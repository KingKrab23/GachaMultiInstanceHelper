from macro_launcher import rename_and_screenshot_hbr_windows

def load_salted_email_number():
    try:
        with open('persistent_variable.json', 'r') as f:
            data = json.load(f)
            return data.get('salted_email', 0)
    except FileNotFoundError:
        return 0

def save_salted_email_number(number):
    with open('persistent_variable.json', 'w') as f:
        json.dump({'salted_email': number}, f)

def rename_hbr_windows():
    renamed_windows = []
    current_number = load_salted_email_number()

    def callback(hwnd, _):
        nonlocal current_number
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            # Check for both HBR and LDPlayer windows
            if any(pattern in title for pattern in ['HBR-', 'LDPlayer-']):
                try:
                    new_title = f"pimsh0+{current_number}@hotmail.com"
                    win32gui.SetWindowText(hwnd, new_title)
                    renamed_windows.append((title, new_title))
                    current_number += 1
                except Exception as e:
                    print(f"Error renaming window {title}: {str(e)}")

    win32gui.EnumWindows(callback, None)
    
    if renamed_windows:
        save_salted_email_number(current_number)
        
    return renamed_windows

def rename_and_screenshot_hbr_windows():
    import pyautogui
    renamed = rename_hbr_windows()
    if renamed:
        for old_title, new_title in renamed:
            # Take a screenshot of the renamed window
            screenshot = pyautogui.screenshot()
            screenshot_path = f"{new_title}.png"
            screenshot.save(screenshot_path)
            print(f"Screenshot saved to: {screenshot_path}")
    return renamed

if __name__ == "__main__":
    # Rename all HBR windows and take screenshots
    processed = rename_and_screenshot_hbr_windows()
    
    if processed:
        print("\nProcessed windows:")
        for old_title, new_title in processed:
            print(f"Renamed: '{old_title}' -> '{new_title}'")
            print("-" * 50)
    else:
        print("No HBR windows found to process")
