"""
Key Press Handler

A standardized utility for handling key presses between Python and C# applications.
This module provides consistent key mapping and event handling to ensure compatibility
between the Python GUI application and the C# macro automator.
"""

import os
import sys
import json
import time
import keyboard
import pyautogui
import win32gui
import win32con
import win32api
import win32process
import ctypes
from ctypes import wintypes

# Windows API constants
WM_KEYDOWN = 0x0100
WM_KEYUP = 0x0101
WM_SYSKEYDOWN = 0x0104
WM_SYSKEYUP = 0x0105

# Virtual key code mapping (Python key name to Windows VK code)
KEY_MAPPING = {
    # Special keys
    'enter': 0x0D,
    'tab': 0x09,
    'space': 0x20,
    'backspace': 0x08,
    'escape': 0x1B,
    'up': 0x26,
    'down': 0x28,
    'left': 0x25,
    'right': 0x27,
    'ctrl': 0x11,
    'alt': 0x12,
    'shift': 0x10,
    'win': 0x5B,
    # Function keys
    'f1': 0x70,
    'f2': 0x71,
    'f3': 0x72,
    'f4': 0x73,
    'f5': 0x74,
    'f6': 0x75,
    'f7': 0x76,
    'f8': 0x77,
    'f9': 0x78,
    'f10': 0x79,
    'f11': 0x7A,
    'f12': 0x7B,
    # Numeric keys
    '0': 0x30,
    '1': 0x31,
    '2': 0x32,
    '3': 0x33,
    '4': 0x34,
    '5': 0x35,
    '6': 0x36,
    '7': 0x37,
    '8': 0x38,
    '9': 0x39,
    # Letter keys (a-z)
    'a': 0x41,
    'b': 0x42,
    'c': 0x43,
    'd': 0x44,
    'e': 0x45,
    'f': 0x46,
    'g': 0x47,
    'h': 0x48,
    'i': 0x49,
    'j': 0x4A,
    'k': 0x4B,
    'l': 0x4C,
    'm': 0x4D,
    'n': 0x4E,
    'o': 0x4F,
    'p': 0x50,
    'q': 0x51,
    'r': 0x52,
    's': 0x53,
    't': 0x54,
    'u': 0x55,
    'v': 0x56,
    'w': 0x57,
    'x': 0x58,
    'y': 0x59,
    'z': 0x5A,
    # Punctuation and other keys
    '.': 0xBE,  # Period
    ',': 0xBC,  # Comma
    ';': 0xBA,  # Semicolon
    "'": 0xDE,  # Apostrophe
    '[': 0xDB,  # Left bracket
    ']': 0xDD,  # Right bracket
    '\\': 0xDC,  # Backslash
    '/': 0xBF,  # Forward slash
    '-': 0xBD,  # Minus
    '=': 0xBB,  # Equals
    '`': 0xC0,  # Backtick
}

# Reverse mapping (Windows VK code to Python key name)
VK_TO_KEY = {v: k for k, v in KEY_MAPPING.items()}

class KeyPressHandler:
    """Handler for standardized key press events between Python and C#"""
    
    def __init__(self, debug=False):
        """Initialize the key press handler
        
        Args:
            debug: Enable debug logging
        """
        self.debug = debug
        self.active_modifiers = set()
        self.last_key_time = 0
        self.key_log = []
        
    def log(self, message):
        """Log a debug message if debug mode is enabled"""
        if self.debug:
            print(f"[KeyPressHandler] {message}")
    
    def send_key_to_window(self, hwnd, key, modifiers=None):
        """Send a key press to a specific window
        
        Args:
            hwnd: Window handle to send the key to
            key: Key name (string) to send
            modifiers: List of modifier keys (ctrl, alt, shift, win)
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Ensure the window is in the foreground
            if not self.activate_window(hwnd):
                self.log(f"Failed to activate window {hwnd}")
                return False
            
            # Small delay to ensure window is active
            time.sleep(0.05)
            
            # Get virtual key code
            vk_code = self.get_vk_code(key)
            if vk_code is None:
                self.log(f"Unknown key: {key}")
                return False
            
            # Handle modifiers
            modifier_vks = []
            if modifiers:
                for mod in modifiers:
                    mod_vk = self.get_vk_code(mod.lower())
                    if mod_vk:
                        modifier_vks.append(mod_vk)
            
            # Press modifier keys
            for mod_vk in modifier_vks:
                win32api.PostMessage(hwnd, WM_KEYDOWN, mod_vk, 0)
                time.sleep(0.01)
            
            # Press and release the main key
            win32api.PostMessage(hwnd, WM_KEYDOWN, vk_code, 0)
            time.sleep(0.01)
            win32api.PostMessage(hwnd, WM_KEYUP, vk_code, 0)
            time.sleep(0.01)
            
            # Release modifier keys in reverse order
            for mod_vk in reversed(modifier_vks):
                win32api.PostMessage(hwnd, WM_KEYUP, mod_vk, 0)
                time.sleep(0.01)
            
            self.log(f"Sent key {key} to window {hwnd}")
            return True
            
        except Exception as e:
            self.log(f"Error sending key to window: {str(e)}")
            return False
    
    def send_key_combo(self, key_combo, target_window=None):
        """Send a key combination (e.g., 'ctrl+shift+p')
        
        Args:
            key_combo: String representing the key combination
            target_window: Optional window title or handle to target
            
        Returns:
            True if successful, False otherwise
        """
        try:
            self.log(f"Sending key combo: {key_combo} to window: {target_window}")
            
            # Parse key combination
            keys = key_combo.lower().split('+')
            
            # Separate modifiers from the main key
            modifiers = []
            main_key = None
            
            for key in keys:
                key = key.strip()
                if key in ('ctrl', 'alt', 'shift', 'win'):
                    modifiers.append(key)
                else:
                    main_key = key
            
            # If no main key was found, use the last key as the main key
            if main_key is None and keys:
                main_key = keys[-1]
                if main_key in modifiers:
                    modifiers.remove(main_key)
            
            self.log(f"Parsed combo - Modifiers: {modifiers}, Main key: {main_key}")
            
            # Check if we have a Ctrl+Alt combination (which can be problematic)
            has_ctrl_alt = 'ctrl' in modifiers and 'alt' in modifiers
            
            # Find target window if specified
            hwnd = None
            if target_window:
                if isinstance(target_window, int):
                    hwnd = target_window
                else:
                    hwnd = self.find_window_by_title(target_window)
                
                if hwnd:
                    # For Ctrl+Alt combinations, use a different approach
                    if has_ctrl_alt:
                        self.log(f"Using special handling for Ctrl+Alt+{main_key} combination")
                        return self._send_ctrl_alt_combo_to_window(hwnd, main_key)
                    else:
                        return self.send_key_to_window(hwnd, main_key, modifiers)
            
            # If no target window or window not found, use pyautogui
            if has_ctrl_alt:
                self.log(f"Using special handling for Ctrl+Alt+{main_key} combination")
                # Use a more direct approach for Ctrl+Alt combinations
                try:
                    # Press modifier keys
                    for mod in modifiers:
                        pyautogui.keyDown(mod)
                        self.log(f"Pressed modifier: {mod}")
                        time.sleep(0.05)  # Slightly longer delay for modifiers
                    
                    # Press and release the main key
                    pyautogui.press(main_key)
                    self.log(f"Pressed main key: {main_key}")
                    time.sleep(0.05)
                    
                    # Release modifier keys in reverse order
                    for mod in reversed(modifiers):
                        pyautogui.keyUp(mod)
                        self.log(f"Released modifier: {mod}")
                        time.sleep(0.05)
                    
                    return True
                except Exception as e:
                    self.log(f"Error with special Ctrl+Alt handling: {str(e)}")
                    # Fall back to standard method
            
            # Standard method
            if modifiers:
                keys_to_press = modifiers + [main_key]
                pyautogui.hotkey(*keys_to_press)
            else:
                pyautogui.press(main_key)
            
            self.log(f"Sent key combo: {key_combo}")
            return True
            
        except Exception as e:
            self.log(f"Error sending key combo: {str(e)}")
            return False
    
    def _send_ctrl_alt_combo_to_window(self, hwnd, key):
        """Special handling for Ctrl+Alt combinations to a window
        
        Args:
            hwnd: Window handle
            key: The main key to press with Ctrl+Alt
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Ensure the window is in the foreground
            if not self.activate_window(hwnd):
                self.log(f"Failed to activate window {hwnd}")
                return False
            
            # Small delay to ensure window is active
            time.sleep(0.05)
            
            # Get virtual key codes
            ctrl_vk = self.get_vk_code('ctrl')
            alt_vk = self.get_vk_code('alt')
            key_vk = self.get_vk_code(key)
            
            self.log(f"Key codes - Ctrl: {ctrl_vk}, Alt: {alt_vk}, Key '{key}': {key_vk}")
            
            if key_vk is None:
                self.log(f"Unknown key: {key}")
                return False
            
            # Use extended key flag for Alt key and main key
            extended_flag = 0x0001  # KEYEVENTF_EXTENDEDKEY
            
            # Press Ctrl first
            win32api.PostMessage(hwnd, WM_KEYDOWN, ctrl_vk, 0)
            time.sleep(0.05)  # Increased delay
            
            # Press Alt with extended flag
            win32api.PostMessage(hwnd, WM_SYSKEYDOWN, alt_vk, (1 << 24) | (1 << 29))
            time.sleep(0.05)  # Increased delay
            
            # Press and release the main key
            scan_code = win32api.MapVirtualKey(key_vk, 0)
            lparam = (scan_code << 16) | 1 | (1 << 24)  # Repeat=1, extended=1
            win32api.PostMessage(hwnd, WM_KEYDOWN, key_vk, lparam)
            time.sleep(0.05)  # Increased delay
            
            # Release the main key
            lparam = (scan_code << 16) | (1 << 30) | (1 << 31) | 1 | (1 << 24)  # Up=1, previous=1
            win32api.PostMessage(hwnd, WM_KEYUP, key_vk, lparam)
            time.sleep(0.05)  # Increased delay
            
            # Release Alt
            win32api.PostMessage(hwnd, WM_SYSKEYUP, alt_vk, (1 << 30) | (1 << 31) | (1 << 24) | (1 << 29))
            time.sleep(0.05)  # Increased delay
            
            # Release Ctrl
            win32api.PostMessage(hwnd, WM_KEYUP, ctrl_vk, (1 << 30) | (1 << 31))
            time.sleep(0.05)  # Increased delay
            
            self.log(f"Sent Ctrl+Alt+{key} to window {hwnd}")
            return True
            
        except Exception as e:
            self.log(f"Error sending Ctrl+Alt combo to window: {str(e)}")
            return False
    
    def get_vk_code(self, key_name):
        """Get the virtual key code for a key name
        
        Args:
            key_name: Name of the key
            
        Returns:
            Virtual key code or None if not found
        """
        key_name = key_name.lower()
        
        # Check direct mapping
        if key_name in KEY_MAPPING:
            return KEY_MAPPING[key_name]
        
        # Handle numeric keys (0-9)
        if key_name.isdigit() and len(key_name) == 1:
            # For digits, use the ASCII value (0x30 for '0', 0x31 for '1', etc.)
            return ord(key_name)
        
        # Handle single character keys
        if len(key_name) == 1:
            # Use the Windows API to get the virtual key code
            vk_code = win32api.VkKeyScan(key_name)
            if vk_code != -1:
                return vk_code & 0xFF
        
        self.log(f"No virtual key code found for: {key_name}")
        return None
    
    def activate_window(self, hwnd):
        """Activate a window by bringing it to the foreground
        
        Args:
            hwnd: Window handle
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Check if window exists
            if not win32gui.IsWindow(hwnd):
                return False
            
            # Get current foreground window
            current_hwnd = win32gui.GetForegroundWindow()
            
            # If already in foreground, return success
            if current_hwnd == hwnd:
                return True
            
            # Bring window to foreground
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
            
            # Verify window was brought to foreground
            time.sleep(0.05)
            return win32gui.GetForegroundWindow() == hwnd
            
        except Exception as e:
            self.log(f"Error activating window: {str(e)}")
            return False
    
    def find_window_by_title(self, title_pattern):
        """Find a window by its title
        
        Args:
            title_pattern: Full or partial window title
            
        Returns:
            Window handle or None if not found
        """
        result = []
        
        def enum_windows_callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                window_title = win32gui.GetWindowText(hwnd)
                if title_pattern.lower() in window_title.lower():
                    result.append(hwnd)
        
        win32gui.EnumWindows(enum_windows_callback, None)
        
        if result:
            return result[0]
        return None
    
    def register_hotkey(self, key_combo, callback):
        """Register a global hotkey
        
        Args:
            key_combo: Key combination string (e.g., 'ctrl+alt+1')
            callback: Function to call when hotkey is pressed
            
        Returns:
            True if registered successfully, False otherwise
        """
        try:
            keyboard.add_hotkey(key_combo.lower(), callback)
            self.log(f"Registered hotkey: {key_combo}")
            return True
        except Exception as e:
            self.log(f"Error registering hotkey: {str(e)}")
            return False
    
    def convert_yaml_key_format(self, yaml_data):
        """Convert between different YAML key formats
        
        Args:
            yaml_data: YAML data dictionary
            
        Returns:
            Converted YAML data
        """
        if 'macros' not in yaml_data:
            yaml_data = {'macros': {}, 'settings': {}}
        
        # Process each macro
        for macro_name, macro_data in yaml_data['macros'].items():
            if 'actions' not in macro_data:
                continue
            
            # Process each action
            new_actions = []
            for action in macro_data['actions']:
                # Convert old format (Type/Parameters) to new format (type/params)
                if 'Type' in action and 'Parameters' in action:
                    action_type = action['Type'].lower()
                    params = action['Parameters']
                    
                    new_action = {'type': action_type}
                    for param_name, param_value in params.items():
                        new_action[param_name.lower()] = param_value
                    
                    new_actions.append(new_action)
                
                # Handle nested 'value' actions
                elif 'type' in action and 'value' in action and action['type'] == 'Actions':
                    for nested_action in action['value']:
                        if 'Type' in nested_action and 'Parameters' in nested_action:
                            action_type = nested_action['Type'].lower()
                            params = nested_action['Parameters']
                            
                            new_action = {'type': action_type}
                            for param_name, param_value in params.items():
                                new_action[param_name.lower()] = param_value
                            
                            new_actions.append(new_action)
                
                # Already in new format
                elif 'type' in action:
                    new_actions.append(action)
            
            # Replace actions with standardized format
            macro_data['actions'] = new_actions
        
        return yaml_data

# Singleton instance
_handler = None

def get_handler(debug=False):
    """Get the singleton KeyPressHandler instance"""
    global _handler
    if _handler is None:
        _handler = KeyPressHandler(debug=debug)
    return _handler

# Convenience functions
def send_key(key, target_window=None):
    """Send a single key press"""
    return get_handler().send_key_combo(key, target_window)

def send_key_combo(key_combo, target_window=None):
    """Send a key combination"""
    return get_handler().send_key_combo(key_combo, target_window)

def find_window(title_pattern):
    """Find a window by title"""
    return get_handler().find_window_by_title(title_pattern)

def register_hotkey(key_combo, callback):
    """Register a global hotkey"""
    return get_handler().register_hotkey(key_combo, callback)

def convert_yaml(yaml_data):
    """Convert YAML data to standardized format"""
    return get_handler().convert_yaml_key_format(yaml_data)
