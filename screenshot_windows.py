import os
import time
from datetime import datetime
import win32gui
import win32ui
from ctypes import windll
from PIL import Image

def ensure_screenshots_dir():
    """Create screenshots directory if it doesn't exist"""
    screenshots_dir = os.path.join(os.getcwd(), 'screenshots')
    if not os.path.exists(screenshots_dir):
        os.makedirs(screenshots_dir)
    return screenshots_dir

def capture_window_screenshot(hwnd, save_path):
    """
    Capture a screenshot of a specific window and save it
    """
    # Get the window size
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width = right - left
    height = bottom - top

    # Get the window DC
    hwndDC = win32gui.GetWindowDC(hwnd)
    mfcDC = win32ui.CreateDCFromHandle(hwndDC)
    saveDC = mfcDC.CreateCompatibleDC()

    # Create a bitmap object
    saveBitMap = win32ui.CreateBitmap()
    saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
    saveDC.SelectObject(saveBitMap)

    # Copy the screen into our memory device context
    result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 3)

    # Convert the raw data into a format PIL understands
    bmpinfo = saveBitMap.GetInfo()
    bmpstr = saveBitMap.GetBitmapBits(True)
    im = Image.frombuffer(
        'RGB',
        (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
        bmpstr, 'raw', 'BGRX', 0, 1)

    # Clean up
    win32gui.DeleteObject(saveBitMap.GetHandle())
    saveDC.DeleteDC()
    mfcDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwndDC)

    # Save the image
    im.save(save_path)
    return save_path

def screenshot_windows():
    """
    Take screenshots of all windows with email addresses as titles
    """
    screenshots_dir = ensure_screenshots_dir()
    
    # Find all windows with email addresses
    email_windows = []
    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if '@' in title:  # Simple check for email address
                email_windows.append((hwnd, title))
    win32gui.EnumWindows(callback, None)
    
    # Take screenshots
    screenshots_taken = []
    for hwnd, title in email_windows:
        try:
            # Bring window to front
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.5)  # Wait for window to be in foreground
            
            # Generate filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{title.replace('@', '_at_')}_{timestamp}.png"
            screenshot_path = os.path.join(screenshots_dir, filename)
            
            # Take screenshot
            capture_window_screenshot(hwnd, screenshot_path)
            screenshots_taken.append({
                'title': title,
                'path': screenshot_path
            })
        except Exception as e:
            print(f"Error capturing screenshot for {title}: {str(e)}")
    
    return screenshots_taken

if __name__ == "__main__":
    print("Taking screenshots of windows with email addresses as titles...")
    screenshots = screenshot_windows()
    
    if screenshots:
        print("\nScreenshots taken:")
        for screenshot in screenshots:
            print(f"Window: {screenshot['title']}")
            print(f"Saved to: {screenshot['path']}")
            print("-" * 50)
    else:
        print("No windows with email addresses found")
