"""
Macro Automator

A flexible tool for automating macros in various applications.
Features:
- Send keystrokes to applications
- Perform mouse actions
- Execute sequences with timing controls
- Loop actions with configurable iterations
- Target specific windows by title or process
"""

import time
import sys
import os
import json
import argparse
import threading
import logging
from datetime import datetime
import pyautogui
import keyboard
import win32gui
import win32con
import win32process
import psutil
import yaml

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("macro_automator.log"),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

class MacroAutomator:
    """Main class for automating macros across applications"""
    
    def __init__(self, config_file=None):
        """Initialize the automator
        
        Args:
            config_file: Path to a YAML config file with macro definitions
        """
        self.running = False
        self.paused = False
        self.current_macro = None
        self.stop_event = threading.Event()
        self.macro_sequences = {}
        self.target_windows = []
        self.current_window_index = 0
        
        # Load configuration if provided
        if config_file and os.path.exists(config_file):
            self.load_config(config_file)
    
    def load_config(self, config_file):
        """Load macro configurations from a YAML file
        
        Args:
            config_file: Path to the configuration file
        """
        try:
            with open(config_file, 'r') as f:
                config = yaml.safe_load(f)
            
            # Load macro sequences
            if 'macros' in config:
                self.macro_sequences = config['macros']
                logger.info(f"Loaded {len(self.macro_sequences)} macro sequences from config")
            
            # Load global settings
            if 'settings' in config:
                settings = config['settings']
                if 'default_delay' in settings:
                    pyautogui.PAUSE = settings['default_delay']
                    logger.info(f"Set default delay to {pyautogui.PAUSE} seconds")
            
            return True
        except Exception as e:
            logger.error(f"Error loading config: {str(e)}")
            return False
    
    def save_config(self, config_file):
        """Save current macro configurations to a YAML file
        
        Args:
            config_file: Path to save the configuration
        """
        try:
            config = {
                'settings': {
                    'default_delay': pyautogui.PAUSE
                },
                'macros': self.macro_sequences
            }
            
            with open(config_file, 'w') as f:
                yaml.dump(config, f, default_flow_style=False)
            
            logger.info(f"Saved configuration to {config_file}")
            return True
        except Exception as e:
            logger.error(f"Error saving config: {str(e)}")
            return False
    
    def find_windows(self, title_pattern=None, process_name=None):
        """Find windows matching the specified criteria
        
        Args:
            title_pattern: String pattern to match in window titles
            process_name: Process name to match
            
        Returns:
            List of window handles
        """
        self.target_windows = []
        
        def enum_windows_callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
                window_title = win32gui.GetWindowText(hwnd)
                
                # Check title pattern if specified
                title_match = True
                if title_pattern and title_pattern.lower() not in window_title.lower():
                    title_match = False
                
                # Check process name if specified
                process_match = True
                if process_name:
                    try:
                        _, pid = win32process.GetWindowThreadProcessId(hwnd)
                        process = psutil.Process(pid)
                        if process.name().lower() != process_name.lower():
                            process_match = False
                    except:
                        process_match = False
                
                if title_match and process_match:
                    self.target_windows.append({
                        'hwnd': hwnd,
                        'title': window_title
                    })
        
        win32gui.EnumWindows(enum_windows_callback, None)
        logger.info(f"Found {len(self.target_windows)} matching windows")
        return len(self.target_windows) > 0
    
    def activate_window(self, window_info):
        """Activate a specific window
        
        Args:
            window_info: Dictionary with window information
        
        Returns:
            True if successful, False otherwise
        """
        try:
            hwnd = window_info['hwnd']
            
            # Bring window to foreground
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
            
            # Get window position and size for centering mouse
            rect = win32gui.GetWindowRect(hwnd)
            x = (rect[0] + rect[2]) // 2
            y = (rect[1] + rect[3]) // 2
            
            # Move mouse to center of window
            pyautogui.moveTo(x, y)
            
            logger.info(f"Activated window: {window_info['title']}")
            return True
        except Exception as e:
            logger.error(f"Error activating window: {str(e)}")
            return False
    
    def cycle_to_next_window(self):
        """Cycle to the next target window
        
        Returns:
            True if successful, False otherwise
        """
        if not self.target_windows:
            logger.warning("No target windows available")
            return False
        
        self.current_window_index = (self.current_window_index + 1) % len(self.target_windows)
        return self.activate_window(self.target_windows[self.current_window_index])
    
    def execute_action(self, action):
        """Execute a single macro action
        
        Args:
            action: Dictionary describing the action to perform
        
        Returns:
            True if successful, False otherwise
        """
        try:
            action_type = action.get('type', '').lower()
            
            # Handle different action types
            if action_type == 'keypress':
                key = action.get('key', '')
                if key:
                    pyautogui.press(key)
                    logger.info(f"Pressed key: {key}")
            
            elif action_type == 'keydown':
                key = action.get('key', '')
                if key:
                    pyautogui.keyDown(key)
                    logger.info(f"Key down: {key}")
            
            elif action_type == 'keyup':
                key = action.get('key', '')
                if key:
                    pyautogui.keyUp(key)
                    logger.info(f"Key up: {key}")
            
            elif action_type == 'hotkey':
                keys = action.get('keys', [])
                if keys:
                    pyautogui.hotkey(*keys)
                    logger.info(f"Pressed hotkey: {'+'.join(keys)}")
            
            elif action_type == 'type':
                text = action.get('text', '')
                if text:
                    pyautogui.write(text)
                    logger.info(f"Typed text: {text}")
            
            elif action_type == 'click':
                x = action.get('x', None)
                y = action.get('y', None)
                button = action.get('button', 'left')
                
                if x is not None and y is not None:
                    pyautogui.click(x, y, button=button)
                    logger.info(f"Clicked at position: ({x}, {y}) with {button} button")
                else:
                    pyautogui.click(button=button)
                    logger.info(f"Clicked with {button} button at current position")
            
            elif action_type == 'move':
                x = action.get('x', 0)
                y = action.get('y', 0)
                duration = action.get('duration', 0.25)
                
                pyautogui.moveTo(x, y, duration=duration)
                logger.info(f"Moved to position: ({x}, {y})")
            
            elif action_type == 'sleep':
                duration = action.get('duration', 1.0)
                logger.info(f"Sleeping for {duration} seconds")
                time.sleep(duration)
            
            elif action_type == 'cycle_window':
                self.cycle_to_next_window()
            
            else:
                logger.warning(f"Unknown action type: {action_type}")
                return False
            
            return True
        
        except Exception as e:
            logger.error(f"Error executing action: {str(e)}")
            return False
    
    def execute_sequence(self, sequence_name, iterations=1):
        """Execute a sequence of macro actions
        
        Args:
            sequence_name: Name of the sequence to execute
            iterations: Number of times to repeat the sequence
        
        Returns:
            True if successful, False otherwise
        """
        if sequence_name not in self.macro_sequences:
            logger.error(f"Sequence '{sequence_name}' not found")
            return False
        
        sequence = self.macro_sequences[sequence_name]
        actions = sequence.get('actions', [])
        
        if not actions:
            logger.warning(f"Sequence '{sequence_name}' has no actions")
            return False
        
        logger.info(f"Executing sequence '{sequence_name}' ({iterations} iterations)")
        self.current_macro = sequence_name
        
        try:
            for iteration in range(iterations):
                if self.stop_event.is_set():
                    logger.info("Execution stopped by user")
                    break
                
                logger.info(f"Starting iteration {iteration + 1}/{iterations}")
                
                for action in actions:
                    # Check if execution was stopped
                    if self.stop_event.is_set():
                        logger.info("Execution stopped by user")
                        break
                    
                    # Check if execution is paused
                    while self.paused and not self.stop_event.is_set():
                        time.sleep(0.1)
                    
                    # Execute the action
                    self.execute_action(action)
                
                # Check for inter-iteration delay
                if iteration < iterations - 1:
                    delay = sequence.get('iteration_delay', 0)
                    if delay > 0:
                        logger.info(f"Waiting {delay} seconds before next iteration")
                        time.sleep(delay)
            
            logger.info(f"Sequence '{sequence_name}' completed")
            self.current_macro = None
            return True
        
        except Exception as e:
            logger.error(f"Error executing sequence: {str(e)}")
            self.current_macro = None
            return False
    
    def start_sequence(self, sequence_name, iterations=1):
        """Start executing a sequence in a separate thread
        
        Args:
            sequence_name: Name of the sequence to execute
            iterations: Number of times to repeat the sequence
        
        Returns:
            True if started successfully, False otherwise
        """
        if self.running:
            logger.warning("A macro sequence is already running")
            return False
        
        self.running = True
        self.stop_event.clear()
        
        thread = threading.Thread(
            target=self._sequence_thread,
            args=(sequence_name, iterations),
            daemon=True
        )
        thread.start()
        
        return True
    
    def _sequence_thread(self, sequence_name, iterations):
        """Thread function for executing a sequence
        
        Args:
            sequence_name: Name of the sequence to execute
            iterations: Number of times to repeat the sequence
        """
        try:
            self.execute_sequence(sequence_name, iterations)
        finally:
            self.running = False
    
    def stop(self):
        """Stop the currently running macro sequence"""
        if not self.running:
            logger.info("No macro sequence is currently running")
            return
        
        logger.info("Stopping macro execution")
        self.stop_event.set()
    
    def pause(self):
        """Pause the currently running macro sequence"""
        if not self.running:
            logger.info("No macro sequence is currently running")
            return
        
        if self.paused:
            logger.info("Macro execution is already paused")
            return
        
        logger.info("Pausing macro execution")
        self.paused = True
    
    def resume(self):
        """Resume the paused macro sequence"""
        if not self.running:
            logger.info("No macro sequence is currently running")
            return
        
        if not self.paused:
            logger.info("Macro execution is not paused")
            return
        
        logger.info("Resuming macro execution")
        self.paused = False
    
    def create_sequence(self, name, actions=None, iteration_delay=0):
        """Create a new macro sequence
        
        Args:
            name: Name of the sequence
            actions: List of action dictionaries
            iteration_delay: Delay between iterations in seconds
        
        Returns:
            True if successful, False otherwise
        """
        if name in self.macro_sequences:
            logger.warning(f"Sequence '{name}' already exists, overwriting")
        
        self.macro_sequences[name] = {
            'actions': actions or [],
            'iteration_delay': iteration_delay
        }
        
        logger.info(f"Created sequence '{name}' with {len(actions or [])} actions")
        return True
    
    def add_action_to_sequence(self, sequence_name, action):
        """Add an action to an existing sequence
        
        Args:
            sequence_name: Name of the sequence
            action: Action dictionary to add
        
        Returns:
            True if successful, False otherwise
        """
        if sequence_name not in self.macro_sequences:
            logger.error(f"Sequence '{sequence_name}' not found")
            return False
        
        self.macro_sequences[sequence_name]['actions'].append(action)
        logger.info(f"Added action to sequence '{sequence_name}'")
        return True
    
    def get_status(self):
        """Get the current status of the automator
        
        Returns:
            Dictionary with status information
        """
        return {
            'running': self.running,
            'paused': self.paused,
            'current_macro': self.current_macro,
            'available_macros': list(self.macro_sequences.keys()),
            'target_windows': len(self.target_windows),
            'current_window': self.current_window_index if self.target_windows else -1
        }

def main():
    """Main entry point for the command-line interface"""
    parser = argparse.ArgumentParser(description='Macro Automator')
    parser.add_argument('--config', help='Path to configuration file')
    parser.add_argument('--sequence', help='Name of the sequence to run')
    parser.add_argument('--iterations', type=int, default=1, help='Number of iterations')
    parser.add_argument('--window', help='Window title pattern to target')
    parser.add_argument('--process', help='Process name to target')
    
    args = parser.parse_args()
    
    # Create the automator
    automator = MacroAutomator(args.config)
    
    # Find target windows if specified
    if args.window or args.process:
        automator.find_windows(args.window, args.process)
    
    # Run the specified sequence if provided
    if args.sequence:
        if args.sequence in automator.macro_sequences:
            automator.execute_sequence(args.sequence, args.iterations)
        else:
            logger.error(f"Sequence '{args.sequence}' not found")
            return 1
    else:
        logger.info("No sequence specified. Use --sequence to run a macro.")
        return 0
    
    return 0

if __name__ == "__main__":
    sys.exit(main())
