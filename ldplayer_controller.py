"""
LD Player Controller Module

This module provides functionality to control LD Player instances by sending hotkeys
and automating operations. It runs as a background process that can be started and stopped.
"""

import time
import keyboard
import threading
import win32gui
import win32con
import pyautogui
import random
from datetime import datetime

class LDPlayerController:
    """Controller for automating LD Player operations"""
    
    def __init__(self, log_callback=None):
        """Initialize the controller
        
        Args:
            log_callback: Function to call for logging messages
        """
        self.running = False
        self.thread = None
        self.log_callback = log_callback
        self.window_handles = []
        self.current_window_index = 0
        
    def log(self, message):
        """Log a message using the callback if available"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        formatted_message = f"[{timestamp}] {message}"
        if self.log_callback:
            self.log_callback(formatted_message)
        else:
            print(formatted_message)
    
    def find_ldplayer_windows(self):
        """Find all LD Player windows and store their handles"""
        self.window_handles = []
        
        def enum_windows_callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                window_title = win32gui.GetWindowText(hwnd)
                if "LDPlayer" in window_title:
                    self.window_handles.append(hwnd)
        
        win32gui.EnumWindows(enum_windows_callback, None)
        self.log(f"Found {len(self.window_handles)} LD Player windows")
        return len(self.window_handles) > 0
    
    def activate_window(self, window_handle):
        """Activate a specific window by its handle"""
        try:
            # Bring window to foreground
            win32gui.ShowWindow(window_handle, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(window_handle)
            
            # Get window position and size for centering mouse
            rect = win32gui.GetWindowRect(window_handle)
            x = (rect[0] + rect[2]) // 2
            y = (rect[1] + rect[3]) // 2
            
            # Move mouse to center of window
            pyautogui.moveTo(x, y)
            
            window_title = win32gui.GetWindowText(window_handle)
            self.log(f"Activated window: {window_title}")
            return True
        except Exception as e:
            self.log(f"Error activating window: {str(e)}")
            return False
    
    def cycle_to_next_window(self):
        """Cycle to the next LD Player window"""
        if not self.window_handles:
            if not self.find_ldplayer_windows():
                self.log("No LD Player windows found")
                return False
        
        if self.window_handles:
            self.current_window_index = (self.current_window_index + 1) % len(self.window_handles)
            return self.activate_window(self.window_handles[self.current_window_index])
        return False
    
    def press_key(self, key):
        """Press a key in the active window"""
        pyautogui.press(key)
        self.log(f"Pressed key: {key}")
    
    def click_at_position(self, x_percent, y_percent):
        """Click at a relative position in the current window"""
        if not self.window_handles:
            self.log("No LD Player windows available")
            return
        
        hwnd = self.window_handles[self.current_window_index]
        rect = win32gui.GetWindowRect(hwnd)
        width = rect[2] - rect[0]
        height = rect[3] - rect[1]
        
        x = rect[0] + int(width * x_percent)
        y = rect[1] + int(height * y_percent)
        
        pyautogui.click(x, y)
        self.log(f"Clicked at position: {x_percent:.2f}, {y_percent:.2f}")
    
    def perform_random_action(self):
        """Perform a random action in the current window"""
        actions = [
            # Navigate home screen
            lambda: self.press_key('esc'),
            lambda: self.click_at_position(0.5, 0.5),  # Center click
            lambda: self.click_at_position(0.5, 0.8),  # Bottom center
            lambda: self.click_at_position(0.5, 0.2),  # Top center
            
            # Game-specific actions (customize as needed)
            lambda: self.press_key('space'),
            lambda: self.press_key('enter'),
            lambda: self.press_key('tab'),
            
            # Cycle windows
            lambda: self.cycle_to_next_window(),
        ]
        
        # Choose and execute a random action
        action = random.choice(actions)
        action()
        
        # Random delay between 1-3 seconds
        return 1 + random.random() * 2
    
    def controller_loop(self):
        """Main loop for the controller"""
        self.log("LD Player controller started")
        
        if not self.find_ldplayer_windows():
            self.log("No LD Player windows found. Stopping controller.")
            self.running = False
            return
        
        try:
            while self.running:
                # Perform a random action and get the delay
                delay = self.perform_random_action()
                
                # Sleep for the calculated delay
                time.sleep(delay)
                
        except Exception as e:
            self.log(f"Error in controller loop: {str(e)}")
        finally:
            self.log("LD Player controller stopped")
    
    def start(self):
        """Start the controller in a separate thread"""
        if self.running:
            self.log("Controller is already running")
            return False
        
        self.running = True
        self.thread = threading.Thread(target=self.controller_loop, daemon=True)
        self.thread.start()
        return True
    
    def stop(self):
        """Stop the controller"""
        if not self.running:
            self.log("Controller is not running")
            return False
        
        self.running = False
        if self.thread:
            self.thread.join(timeout=1.0)
            self.thread = None
        
        return True
