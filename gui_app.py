import customtkinter as ctk
from macro_launcher import rename_hbr_windows, type_salted_emails, scan_outlook_for_codes, enter_verification_codes, force_outlook_sync
from screenshot_windows import screenshot_windows
import threading
import queue
import time

class MacroGUI(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configure window
        self.title("Macro Control Panel")
        self.geometry("600x600")
        
        # Configure grid layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(6, weight=1)  # Make the log box expand

        # Create main frame
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=6)  # Larger column for scan outlook
        self.main_frame.grid_columnconfigure(1, weight=1)  # Smaller column for force sync

        # Create buttons
        self.rename_button = ctk.CTkButton(
            self.main_frame, 
            text="Rename HBR Windows to Salted Emails",
            command=self.rename_windows
        )
        self.rename_button.grid(row=0, column=0, padx=20, pady=10, sticky="ew", columnspan=2)

        self.type_emails_button = ctk.CTkButton(
            self.main_frame,
            text="Type Salted Emails in LD Player Windows",
            command=self.type_emails
        )
        self.type_emails_button.grid(row=1, column=0, padx=20, pady=10, sticky="ew", columnspan=2)

        self.enter_codes_button = ctk.CTkButton(
            self.main_frame,
            text="Enter Verification Codes in LD PlayerWindows",
            command=self.enter_codes
        )
        self.enter_codes_button.grid(row=2, column=0, padx=20, pady=10, sticky="ew", columnspan=2)

        self.scan_outlook_button = ctk.CTkButton(
            self.main_frame,
            text="Scan Outlook for Verification Codes",
            command=self.scan_outlook
        )
        self.scan_outlook_button.grid(row=3, column=0, padx=20, pady=10, sticky="ew")

        self.force_sync_button = ctk.CTkButton(
            self.main_frame,
            text="Force Outlook Sync",
            command=self.force_sync
        )
        self.force_sync_button.grid(row=3, column=1, padx=20, pady=10, sticky="ew")

        self.screenshot_button = ctk.CTkButton(
            self.main_frame,
            text="Take Screenshots of LD Player Windows",
            command=self.take_screenshots
        )
        self.screenshot_button.grid(row=4, column=0, padx=20, pady=10, sticky="ew", columnspan=2)

        # Create status display
        self.status_label = ctk.CTkLabel(self, text="Status: Ready", anchor="w")
        self.status_label.grid(row=1, column=0, padx=20, pady=(0, 10), sticky="ew")

        # Create log display
        self.log_frame = ctk.CTkFrame(self)
        self.log_frame.grid(row=2, column=0, padx=20, pady=(0, 20), sticky="nsew")
        self.log_frame.grid_columnconfigure(0, weight=1)
        self.log_frame.grid_rowconfigure(0, weight=1)

        self.log_text = ctk.CTkTextbox(self.log_frame, height=200)
        self.log_text.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # Message queue for thread-safe logging
        self.msg_queue = queue.Queue()
        self.after(100, self.check_queue)

    def check_queue(self):
        """Check for new messages in the queue and display them"""
        try:
            while True:
                msg = self.msg_queue.get_nowait()
                self.log_text.insert("end", msg + "\n")
                self.log_text.see("end")
                self.msg_queue.task_done()
        except queue.Empty:
            pass
        self.after(100, self.check_queue)

    def log(self, message):
        """Thread-safe logging"""
        self.msg_queue.put(message)

    def set_status(self, status):
        """Thread-safe status update"""
        self.after(0, lambda: self.status_label.configure(text=f"Status: {status}"))

    def rename_windows(self):
        """Handle window renaming"""
        def run():
            self.set_status("Renaming windows...")
            self.rename_button.configure(state="disabled")
            try:
                renamed = rename_hbr_windows()
                if renamed:
                    self.log("\nRenamed windows:")
                    for old_title, new_title in renamed:
                        self.log(f"'{old_title}' -> '{new_title}'")
                else:
                    self.log("No HBR windows found to rename")
            except Exception as e:
                self.log(f"Error: {str(e)}")
            finally:
                self.rename_button.configure(state="normal")
                self.set_status("Ready")

        threading.Thread(target=run, daemon=True).start()

    def take_screenshots(self):
        """Handle screenshot capture"""
        def run():
            self.set_status("Taking screenshots...")
            self.screenshot_button.configure(state="disabled")
            try:
                screenshots = screenshot_windows()
                if screenshots:
                    self.log("\nScreenshots taken:")
                    for screenshot in screenshots:
                        self.log(f"Window: {screenshot['title']}")
                        self.log(f"Saved to: {screenshot['path']}")
                        self.log("-" * 50)
                else:
                    self.log("No windows with email addresses found")
            except Exception as e:
                self.log(f"Error: {str(e)}")
            finally:
                self.screenshot_button.configure(state="normal")
                self.set_status("Ready")

        threading.Thread(target=run, daemon=True).start()

    def type_emails(self):
        """Handle typing salted emails into windows"""
        def run():
            self.set_status("Typing emails...")
            self.type_emails_button.configure(state="disabled")
            try:
                typed = type_salted_emails()
                if typed:
                    self.log("\nTyped emails in windows:")
                    for window in typed:
                        self.log(f"Window: {window}")
                else:
                    self.log("No windows with email addresses found")
            except Exception as e:
                self.log(f"Error: {str(e)}")
            finally:
                self.type_emails_button.configure(state="normal")
                self.set_status("Ready")

        threading.Thread(target=run, daemon=True).start()

    def scan_outlook(self):
        """Scan Outlook for verification codes"""
        self.log("Scanning Outlook for verification codes...")
        threading.Thread(target=self._scan_outlook_thread).start()

    def enter_codes(self):
        """Enter verification codes in matching windows"""
        self.log("Entering verification codes in windows...")
        threading.Thread(target=self._enter_codes_thread).start()

    def _enter_codes_thread(self):
        """Thread function for entering verification codes"""
        try:
            codes_entered = enter_verification_codes()
            if codes_entered:
                self.log(f"Successfully entered {codes_entered} verification codes")
            else:
                self.log("No verification codes were entered")
        except Exception as e:
            self.log(f"Error entering verification codes: {str(e)}")

    def _scan_outlook_thread(self):
        """Thread function for scanning Outlook"""
        try:
            codes = scan_outlook_for_codes()
            if codes:
                self.log("\nFound verification codes:")
                for email, code in codes.items():
                    self.log(f"Email: {email}")
                    self.log(f"Code: {code}")
                    self.log("-" * 50)
            else:
                self.log("No verification codes found")
        except Exception as e:
            self.log(f"Error scanning Outlook: {str(e)}")

    def force_sync(self):
        """Force Outlook to sync"""
        self.log("Forcing Outlook to sync...")
        threading.Thread(target=self._force_sync_thread).start()

    def _force_sync_thread(self):
        """Thread function for forcing Outlook sync"""
        try:
            if force_outlook_sync():
                self.log("Sync command sent to Outlook successfully")
            else:
                self.log("Failed to sync Outlook")
        except Exception as e:
            self.log(f"Error forcing sync: {str(e)}")

if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")
    
    app = MacroGUI()
    app.mainloop()
