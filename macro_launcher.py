import webbrowser
import json
import pyautogui
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import cv2
import numpy as np
from PIL import Image, ImageGrab
import win32gui
import win32con
import win32ui
from ctypes import windll
from datetime import datetime
import re

# Load persistent variable
def load_variable():
    if os.path.exists('persistent_variable.json'):
        with open('persistent_variable.json', 'r') as f:
            return json.load(f)
    return {'count': 0, 'last_verification_code': None, 'salted_email': None}

# Save persistent variable
def save_variable(data):
    with open('persistent_variable.json', 'w') as f:
        json.dump(data, f)

def get_verification_info():
    try:
        # Set up Chrome driver
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
        
        # Go to Gmail
        driver.get('https://mail.google.com')
        
        # Wait for user to log in manually (first time only)
        try:
            # Wait for the Gmail inbox to load (max 60 seconds)
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.xpath, "//input[@aria-label='Search mail']"))
            )
        except TimeoutException:
            print("Timeout waiting for Gmail to load. Please make sure you're logged in.")
            driver.quit()
            return None, None

        # Search for Yostar verification email
        search_box = driver.find_element(By.xpath, "//input[@aria-label='Search mail']")
        search_box.clear()
        search_box.send_keys('subject:[Yostar] Your Verification Code')
        search_box.submit()

        # Wait for search results
        time.sleep(2)  # Give time for search results to load

        try:
            # Find and click the most recent email
            email_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.xpath, "//span[contains(text(), '[Yostar] Your Verification Code')]"))
            )
            email_element.click()

            # Wait for email content to load and find the verification code
            email_body = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.xpath, "//div[@role='main']"))
            )
            
            # Get the email subject for verification code
            subject_element = driver.find_element(By.xpath, "//h2[contains(@class, 'gmail_msg')]")
            subject_text = subject_element.text
            code_start = subject_text.find('is ') + 3
            verification_code = subject_text[code_start:code_start+6]

            # Find the salted email in the span element
            try:
                email_span = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.xpath, "//span[contains(@class, 'g2') and @email]"))
                )
                salted_email = email_span.get_attribute('email')
            except TimeoutException:
                print("Could not find salted email element")
                salted_email = None

            # Save to persistent variable
            persistent_data = load_variable()
            persistent_data['last_verification_code'] = verification_code
            persistent_data['salted_email'] = salted_email
            save_variable(persistent_data)
            
            driver.quit()
            return verification_code, salted_email

        except TimeoutException:
            print("No verification email found")
            driver.quit()
            return None, None

    except Exception as e:
        print(f"Error getting verification info: {str(e)}")
        try:
            driver.quit()
        except:
            pass
        return None, None

# Open Gmail
def open_gmail():
    webbrowser.open('https://mail.google.com')

# Open application by title
def open_application(title):
    os.system(f'start {title}')

def process_memorias_screenshot(image_path, output_path=None):
    """
    Process a Memorias screenshot to remove unnecessary elements and keep only the cards and rating.
    
    Args:
        image_path: Path to the input image
        output_path: Path to save the processed image. If None, will append '_processed' to input filename
    """
    # Read the image
    img = cv2.imread(image_path)
    if img is None:
        raise ValueError(f"Could not read image at {image_path}")

    # Convert to RGB for better color detection
    img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    
    # Find the rating text area (usually in the top portion)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY)
    
    # Find the top boundary where the cards start
    top_boundary = 0
    for y in range(binary.shape[0]):
        if np.mean(binary[y]) < 250:  # Look for first non-white row
            top_boundary = max(0, y - 10)  # Add some padding
            break
    
    # Find the bottom boundary where the cards end
    bottom_boundary = binary.shape[0]
    for y in range(binary.shape[0] - 1, 0, -1):
        if np.mean(binary[y]) < 250:  # Look for last non-white row
            bottom_boundary = min(binary.shape[0], y + 10)  # Add some padding
            break
    
    # Crop the image
    cropped = img[top_boundary:bottom_boundary, :]
    
    # If output path not specified, create one
    if output_path is None:
        base, ext = os.path.splitext(image_path)
        output_path = f"{base}_processed{ext}"
    
    # Save the processed image
    cv2.imwrite(output_path, cropped)
    return output_path

def find_window_by_title(title_pattern):
    """
    Find a window that matches the given pattern.
    Returns a list of (hwnd, title) tuples.
    """
    result = []
    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            window_title = win32gui.GetWindowText(hwnd)
            print(f"Found window: '{window_title}'")
            if title_pattern.lower() in window_title.lower():
                result.append((hwnd, window_title))
    win32gui.EnumWindows(callback, None)
    return result

def set_window_title(hwnd, new_title):
    """
    Set a new title for the specified window handle.
    """
    try:
        win32gui.SetWindowText(hwnd, new_title)
        return True
    except Exception as e:
        print(f"Error setting window title: {str(e)}")
        return False

def rename_window(old_title_pattern, new_title):
    """
    Find a window by its title pattern and rename it.
    Returns True if successful, False otherwise.
    """
    windows = find_window_by_title(old_title_pattern)
    if not windows:
        print(f"No window found matching pattern: {old_title_pattern}")
        return False
    
    # If multiple windows found, use the first one
    hwnd, current_title = windows[0]
    if set_window_title(hwnd, new_title):
        print(f"Successfully renamed window from '{current_title}' to '{new_title}'")
        return True
    return False

def get_next_salt_number():
    """Get and increment the salt number from persistent storage"""
    try:
        with open('persistent_variable.json', 'r') as f:
            data = json.load(f)
            current_salt = data.get('salted_email', 0)
    except (FileNotFoundError, json.JSONDecodeError):
        current_salt = 0
    
    # Increment and save
    next_salt = current_salt + 1
    with open('persistent_variable.json', 'w') as f:
        json.dump({'salted_email': next_salt}, f)
    
    return current_salt

def generate_salted_email(number):
    """Generate a salted email with the given number"""
    return f"pimsh0+{number}@hotmail.com"

def rename_hbr_windows():
    """
    Rename all HBR windows with salted email addresses
    """
    try:
        # Get current salt number
        salt = load_variable().get('salted_email', 1)
        print(f"Current salt number: {salt}")
        
        # Get all HBR windows
        windows = []
        def callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title.startswith('HBR'):
                    windows.append((hwnd, title))
                elif title.startswith('pimsh0'):
                    # If it's already a pimsh0 window, treat it as a new HBR window
                    windows.append((hwnd, 'HBR'))
        win32gui.EnumWindows(callback, None)
        
        if not windows:
            print("No HBR windows found")
            return
            
        print(f"Found {len(windows)} windows to rename")
        
        # Sort windows by title to maintain order
        windows.sort(key=lambda x: x[1])
        
        # Rename each window
        renamed = 0
        for hwnd, _ in windows:
            try:
                # Generate new salted email
                new_title = f"pimsh0+{salt}@hotmail.com"
                print(f"Renaming window to: {new_title}")
                
                # Set the window title
                win32gui.SetWindowText(hwnd, new_title)
                renamed += 1
                
                # Increment salt
                salt += 1
                
            except Exception as e:
                print(f"Error renaming window: {str(e)}")
                continue
        
        # Save new salt number
        if renamed > 0:
            save_variable({'salted_email': salt})
            print(f"Updated salt number to: {salt}")
        
        print(f"Renamed {renamed} windows")
        return renamed
        
    except Exception as e:
        print(f"Error renaming windows: {str(e)}")
        return 0

def type_salted_emails():
    """
    Find all windows with salted email titles and type the email into each window
    """
    typed_windows = []
    
    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if '@hotmail.com' in title:
                try:
                    # Get the current window placement info
                    placement = win32gui.GetWindowPlacement(hwnd)
                    
                    # Activate window without minimizing others
                    win32gui.ShowWindow(hwnd, 1)  # SW_SHOWNORMAL = 1
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(0.5)  # Wait for window to be active
                    
                    # Type out the email
                    pyautogui.write(title)
                    typed_windows.append(title)
                    
                    # Restore original window state
                    win32gui.SetWindowPlacement(hwnd, placement)
                    
                    time.sleep(0.2)  # Small delay between windows
                except Exception as e:
                    print(f"Error typing in window {title}: {str(e)}")
    
    win32gui.EnumWindows(callback, None)
    return typed_windows

def load_gmail_config():
    try:
        with open('config.json', 'r') as f:
            config = json.load(f)
            return config['gmail']
    except FileNotFoundError:
        print("Please create config.json with Gmail credentials")
        return None

def save_verification_code(timestamp, code, recipient_email, salted_email):
    """Save verification code to verification_codes.json"""
    try:
        # Load existing codes
        codes = {}
        if os.path.exists('verification_codes.json'):
            with open('verification_codes.json', 'r') as f:
                codes = json.load(f)
        
        # Use salted email as key if available, otherwise use recipient email
        email_key = salted_email if salted_email else recipient_email
        
        # Add new code with timestamp and recipient info
        codes[email_key] = {
            'code': code,
            'timestamp': timestamp,
            'full_recipient': recipient_email
        }
        
        # Save updated codes
        with open('verification_codes.json', 'w') as f:
            json.dump(codes, f, indent=4)
            
    except Exception as e:
        print(f"Error saving verification code: {str(e)}")

def force_outlook_sync():
    """Force Outlook to sync all accounts"""
    try:
        import win32com.client
        import pythoncom
        
        # Initialize COM
        pythoncom.CoInitialize()
        
        # Get Outlook application
        outlook = win32com.client.Dispatch("Outlook.Application")
        
        # Get Explorer (main Outlook window)
        explorer = outlook.ActiveExplorer()
        if not explorer:
            print("No active Outlook window found")
            pythoncom.CoUninitialize()
            return False
            
        print("Starting Outlook sync...")
        
        # Press F9 to force sync
        explorer.CommandBars.ExecuteMso("SendReceiveAll")
        print("Triggered Send/Receive All")
        
        # Give some time for sync to complete
        print("Waiting for sync to complete...")
        time.sleep(5)  # Increased wait time to allow for sync
        
        # Clean up COM
        pythoncom.CoUninitialize()
        
        return True
    except Exception as e:
        print(f"Error forcing Outlook sync: {str(e)}")
        try:
            pythoncom.CoUninitialize()
        except:
            pass
        return False

def scan_outlook_for_codes():
    """
    Scan Outlook for verification codes, focusing on unread emails first
    """
    import win32com.client
    import pythoncom
    import re
    from datetime import datetime
    
    try:
        print("Forcing Outlook sync before scanning...")
        force_outlook_sync()
        
        # Initialize COM
        pythoncom.CoInitialize()
        
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # Find the Outlook.com account's inbox
        found_inbox = None
        for account in namespace.Accounts:
            if 'outlook' in account.DisplayName.lower():
                print(f"Found Outlook account: {account.DisplayName}")
                try:
                    # Get the root folder for this account
                    root_folder = account.DeliveryStore.GetRootFolder()
                    # Search through subfolders to find Inbox
                    for folder in root_folder.Folders:
                        if folder.Name.lower() == 'inbox':
                            found_inbox = folder
                            break
                except Exception as e:
                    print(f"Error accessing folders for {account.DisplayName}: {str(e)}")
                    continue
                
                if found_inbox:
                    break
        
        if not found_inbox:
            print("Could not find Outlook.com inbox, falling back to default inbox")
            found_inbox = namespace.GetDefaultFolder(6)  # 6 is the index for inbox
            
        print(f"Scanning folder: {found_inbox.Name}")
        
        # Look for verification code emails
        found_codes = {}
        
        # First check unread emails
        print("Checking unread emails...")
        unread_emails = found_inbox.Items.Restrict("[Unread]=True")
        count = 0
        # Sort by received time in descending order (newest first)
        unread_emails.Sort("[ReceivedTime]", True)
        
        for email in unread_emails:
            if count >= 100:  # Limit to last 100 unread emails
                break
            if process_email(email, found_codes):
                print("Successfully processed unread email")
            count += 1
        
        # If no codes found, check recent read emails
        if not found_codes:
            print("No codes found in unread emails, checking recent read emails...")
            # Get today's emails
            filter_date = datetime.now().strftime("%m/%d/%Y")
            read_emails = found_inbox.Items.Restrict(f"[ReceivedTime] >= '{filter_date}'")
            
            count = 0
            for email in read_emails:
                if count >= 10:  # Limit to 10 recent emails
                    break
                if not email.UnRead:  # Only process read emails
                    if process_email(email, found_codes):
                        print("Successfully processed read email")
                    count += 1
        
        # Clean up COM
        pythoncom.CoUninitialize()
        
        return found_codes
        
    except Exception as e:
        print(f"Error connecting to Outlook: {str(e)}")
        try:
            pythoncom.CoUninitialize()
        except:
            pass
        return None
    finally:
        pythoncom.CoUninitialize()

def process_email(email, codes_dict):
    """
    Process a single email to look for verification code
    """
    try:
        subject = email.Subject
        # Look for verification code in subject with format "[Yostar] Your Verification Code is XXXXXX"
        match = re.search(r'\[Yostar\] Your Verification Code is (\d{6})', subject)
        if match:
            code = match.group(1)
            # Get recipient
            recipient = email.To
            received_time = email.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
            
            # Save the code
            save_verification_code(received_time, code, recipient, recipient)
            print(f"Found code {code} for {recipient} in subject: {subject}")
            
            # Delete the email after processing
            email.Delete()
            print(f"Deleted email with verification code {code}")
            
            return True
            
        return False
        
    except Exception as e:
        print(f"Error processing email: {str(e)}")
        return False

def enter_verification_codes():
    """
    Find windows with titles matching salted emails in verification_codes.json
    and enter their corresponding verification codes
    """
    try:
        # Load verification codes
        if not os.path.exists('verification_codes.json'):
            print("No verification codes found")
            return
            
        with open('verification_codes.json', 'r') as f:
            codes = json.load(f)
            
        if not codes:
            print("No verification codes found")
            return
            
        print(f"Loaded {len(codes)} verification codes:")
        for email, data in codes.items():
            print(f"  {email}: {data['code']}")
        
        # Get all windows
        print("\nSearching for email windows...")
        # Look for windows with titles matching any of our email addresses
        email_windows = []
        def callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                window_title = win32gui.GetWindowText(hwnd)
                if window_title in codes:
                    email_windows.append((hwnd, window_title))
                print(f"Found window: '{window_title}'")
        win32gui.EnumWindows(callback, None)
        
        if not email_windows:
            print("No email windows found")
            return
            
        print(f"\nFound {len(email_windows)} email windows")
        
        codes_entered = 0
        for hwnd, email_key in email_windows:
            print(f"\nProcessing window for {email_key}")
            try:
                # Activate window
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(1.0)  # Wait longer for window to activate
                
                # Get window rect to ensure it's visible
                rect = win32gui.GetWindowRect(hwnd)
                if rect[2] - rect[0] <= 0 or rect[3] - rect[1] <= 0:
                    print(f"Window appears to be minimized or hidden, skipping")
                    continue
                
                # Type the verification code
                code = codes[email_key]['code']
                print(f"Typing code {code} for {email_key}")
                pyautogui.typewrite(code, interval=0.1)
                print(f"Successfully entered code")
                codes_entered += 1
                time.sleep(0.5)  # Wait between windows
            except Exception as e:
                print(f"Error entering code for window {email_key}: {str(e)}")
                continue
                    
        print(f"\nEntered {codes_entered} verification codes")
        return codes_entered
    
    except Exception as e:
        print(f"Error entering verification codes: {str(e)}")
        return 0

# Main function
if __name__ == '__main__':
    persistent_data = load_variable()
    persistent_data['count'] += 1
    save_variable(persistent_data)

    # Rename HBR windows
    renamed = rename_hbr_windows()
    for old_title, new_title in renamed:
        print(f"Renamed '{old_title}' to '{new_title}'")

    verification_code, salted_email = get_verification_info()
    if verification_code:
        print(f"Latest verification code: {verification_code}")
        print(f"Salted email: {salted_email}")
        pyautogui.write(verification_code)
    else:
        print("No verification info found")
    open_application('Notepad')  # Change this to the desired application title
    typed = type_salted_emails()
    for window in typed:
        print(f"Typed email in window: {window}")

    # Outlook automation
    codes = scan_outlook_for_codes()
    if codes:
        print("Found verification codes:")
        for email, code in codes.items():
            print(f"{email}: {code}")
    else:
        print("No verification codes found")

    # Enter verification codes
    entered = enter_verification_codes()
    if entered:
        print(f"Entered {entered} verification codes")
    else:
        print("No verification codes entered")
