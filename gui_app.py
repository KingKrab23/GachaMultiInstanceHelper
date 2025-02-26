import customtkinter as ctk
from macro_launcher import rename_hbr_windows, type_salted_emails, scan_outlook_for_codes, enter_verification_codes, force_outlook_sync
from screenshot_windows import screenshot_windows
from image_matcher import match_memorias
import threading
import queue
import time
import json
import os
from datetime import datetime
import csv
from PIL import Image
from pathlib import Path
import functools

# Global cache for memoria images to improve performance
MEMORIA_IMAGE_CACHE = {}

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

        self.match_memorias_button = ctk.CTkButton(
            self.main_frame,
            text="Match Memorias in Screenshots",
            command=self.match_memorias
        )
        self.match_memorias_button.grid(row=5, column=0, padx=20, pady=10, sticky="ew", columnspan=2)

        self.view_results_button = ctk.CTkButton(
            self.main_frame,
            text="View Memoria Match Results",
            command=self.view_memoria_results
        )
        self.view_results_button.grid(row=6, column=0, padx=20, pady=10, sticky="ew", columnspan=2)

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
        """Take screenshots of all windows with email addresses as titles"""
        self.status_label.configure(text="Status: Taking screenshots...")
        
        def run_task():
            try:
                screenshots = screenshot_windows()
                self.log(f"Screenshots taken: {len(screenshots)}")
                for screenshot in screenshots:
                    self.log(f"- {screenshot['title']}: {screenshot['path']}")
                self.status_label.configure(text="Status: Screenshots complete")
            except Exception as e:
                self.log(f"Error taking screenshots: {str(e)}")
                self.status_label.configure(text="Status: Error taking screenshots")
        
        threading.Thread(target=run_task).start()

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

    def match_memorias(self):
        """Match memorias in screenshots and score them"""
        self.status_label.configure(text="Status: Matching memorias...")
        
        def run_task():
            try:
                # Define custom scores for memorias
                custom_scores = {
                    'yingying-ss1': 20,
                    'aoi-ss2': 5,
                    'yuina-ss3': 1,
                    'seira-ss1': 1,
                    'chie-ss2': 1,
                    'seika-ss2': 20,
                    'seika-ss1': 3,
                    'tama-ss1': 1,
                    'tama-ss4': 2
                }
                
                # Define scoring criteria
                scoring_criteria = {
                    'match_confidence_weight': 0.6,
                    'position_weight': 0.2,
                    'size_weight': 0.1,
                    'color_similarity_weight': 0.1
                }
                
                # Ask if user wants to skip already processed screenshots
                skip_processed = True
                
                # Run the matching process with custom scores
                results = match_memorias(custom_scores, scoring_criteria, skip_processed=skip_processed)
                
                # Log summary of results
                total_emails = len(results)
                total_matches = sum(len(matches) for matches in results.values())
                self.log(f"Memoria matching complete: {total_matches} matches found across {total_emails} emails")
                
                # Log detailed results for top matches
                for email, match_results in results.items():
                    # Get the best match for each email
                    best_matches = []
                    for match_result in match_results:
                        if match_result['matches']:
                            best_match = match_result['matches'][0]  # Top match
                            best_matches.append({
                                'memoria': best_match['memoria_name'],
                                'custom_score': best_match['custom_score'],
                                'match_quality_score': best_match['match_quality_score'],
                                'screenshot': os.path.basename(match_result['screenshot_path'])
                            })
                    
                    # Sort by custom score and then by match quality score
                    best_matches.sort(key=lambda x: (x['custom_score'], x['match_quality_score']), reverse=True)
                    top_matches = best_matches[:3]
                    
                    if top_matches:
                        self.log(f"\nEmail: {email}")
                        for i, match in enumerate(top_matches):
                            self.log(f"  Match {i+1}: {match['memoria']} (Score: {match['custom_score']}, Quality: {match['match_quality_score']:.2f}) in {match['screenshot']}")
                
                # Save results to a more readable format
                self.save_readable_results(results)
                
                self.status_label.configure(text="Status: Memoria matching complete")
            except Exception as e:
                self.log(f"Error matching memorias: {str(e)}")
                self.status_label.configure(text="Status: Error matching memorias")
        
        threading.Thread(target=run_task).start()
    
    def save_readable_results(self, results):
        """Save match results in a more readable format"""
        readable_results = {}
        
        # First, get all unique emails from screenshots directory
        all_emails = set()
        screenshots_dir = os.path.join(os.getcwd(), 'screenshots')
        if os.path.exists(screenshots_dir):
            for filename in os.listdir(screenshots_dir):
                if filename.endswith('.png') and '@' in filename:
                    email = self._extract_email_from_filename(filename)
                    if email:
                        all_emails.add(email)
        
        # Initialize all emails with zero scores
        for email in all_emails:
            readable_results[email] = {
                'matching_memorias': [],
                'score': 0
            }
        
        # Update with actual match results
        for email, match_results in results.items():
            if email not in readable_results:
                readable_results[email] = {
                    'matching_memorias': [],
                    'score': 0
                }
            
            # Collect all memoria matches for this email
            all_matches = []
            for match_result in match_results:
                for match in match_result.get('matches', []):
                    all_matches.append({
                        'memoria': match['memoria_name'],
                        'custom_score': match['custom_score'],
                        'match_quality_score': match['match_quality_score']
                    })
            
            # Get unique memorias with their best scores
            memoria_scores = {}
            for match in all_matches:
                memoria = match['memoria']
                custom_score = match['custom_score']
                match_quality = match['match_quality_score']
                
                if memoria not in memoria_scores or match_quality > memoria_scores[memoria]['match_quality_score']:
                    memoria_scores[memoria] = {
                        'custom_score': custom_score,
                        'match_quality_score': match_quality
                    }
            
            # Calculate overall score (sum of custom scores of top 3 memorias)
            top_memorias = sorted(memoria_scores.items(), key=lambda x: (x[1]['custom_score'], x[1]['match_quality_score']), reverse=True)[:3]
            if top_memorias:
                total_score = sum(info['custom_score'] for _, info in top_memorias)
                readable_results[email]['score'] = total_score
                readable_results[email]['matching_memorias'] = [
                    {'name': name, 'score': info['custom_score']} for name, info in top_memorias
                ]
        
        # Save to file
        with open('memoria_match_results.json', 'w') as f:
            json.dump(readable_results, f, indent=2)
        
        self.log("Saved readable results to memoria_match_results.json")
    
    def _extract_email_from_filename(self, filename):
        """Extract email address from screenshot filename."""
        # Assuming format like "email_at_domain.com_timestamp.png"
        parts = filename.split('_')
        if len(parts) >= 3:
            email_parts = parts[:-2]  # Skip timestamp and extension
            email = '_'.join(email_parts).replace('_at_', '@')
            return email
        return None

    def view_memoria_results(self):
        """Open a new window to display memoria match results in a pretty format"""
        try:
            # Load results from file
            if os.path.exists('memoria_match_results.json'):
                with open('memoria_match_results.json', 'r') as f:
                    results = json.load(f)
            else:
                self.log("No results file found")
                return
            
            # Create a new window
            results_window = ctk.CTkToplevel(self)
            results_window.title("Memoria Match Results")
            results_window.geometry("800x600")
            results_window.grid_columnconfigure(0, weight=1)
            results_window.grid_rowconfigure(1, weight=1)

            # Create header frame
            header_frame = ctk.CTkFrame(results_window)
            header_frame.grid(row=0, column=0, padx=20, pady=(20, 0), sticky="ew")
            header_frame.grid_columnconfigure(0, weight=1)

            # Add title
            title_label = ctk.CTkLabel(
                header_frame, 
                text="Memoria Match Results", 
                font=ctk.CTkFont(size=20, weight="bold")
            )
            title_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")

            # Add export buttons
            export_frame = ctk.CTkFrame(header_frame)
            export_frame.grid(row=0, column=1, padx=10, pady=10, sticky="e")

            export_csv_button = ctk.CTkButton(
                export_frame,
                text="Export to CSV",
                command=lambda: self.export_results_to_csv(results)
            )
            export_csv_button.grid(row=0, column=0, padx=5, pady=5)

            export_txt_button = ctk.CTkButton(
                export_frame,
                text="Export to Text",
                command=lambda: self.export_results_to_text(results)
            )
            export_txt_button.grid(row=0, column=1, padx=5, pady=5)

            # Create main content frame with scrolling
            content_frame = ctk.CTkScrollableFrame(results_window)
            content_frame.grid(row=1, column=0, padx=20, pady=20, sticky="nsew")
            content_frame.grid_columnconfigure(0, weight=1)

            # Sort results by score (highest first)
            sorted_results = sorted(
                [(email, data) for email, data in results.items()],
                key=lambda x: x[1]['score'],
                reverse=True
            )

            # Preload all memoria images in a separate thread to improve performance
            threading.Thread(target=self._preload_memoria_images, args=(sorted_results,), daemon=True).start()

            # Add results to the content frame
            for i, (email, data) in enumerate(sorted_results):
                # Create a card for each email
                card = ctk.CTkFrame(content_frame)
                card.grid(row=i, column=0, padx=10, pady=10, sticky="ew")
                card.grid_columnconfigure(1, weight=1)

                # Rank number
                rank_label = ctk.CTkLabel(
                    card,
                    text=f"#{i+1}",
                    font=ctk.CTkFont(size=16, weight="bold"),
                    width=40,
                    height=40
                )
                rank_label.grid(row=0, column=0, rowspan=2, padx=(10, 5), pady=10)

                # Email and score
                email_label = ctk.CTkLabel(
                    card,
                    text=email,
                    font=ctk.CTkFont(size=14, weight="bold"),
                    anchor="w"
                )
                email_label.grid(row=0, column=1, padx=5, pady=(10, 5), sticky="w")

                score_label = ctk.CTkLabel(
                    card,
                    text=f"Score: {data['score']}",
                    font=ctk.CTkFont(size=14),
                    text_color=("#00AA00" if data['score'] >= 20 else 
                                "#AAAA00" if data['score'] >= 5 else "#666666"),
                )
                score_label.grid(row=0, column=2, padx=5, pady=(10, 5), sticky="e")

                # Memoria list - Only create if there are matching memorias
                if data['matching_memorias']:
                    memoria_frame = ctk.CTkFrame(card)
                    memoria_frame.grid(row=1, column=1, columnspan=2, padx=5, pady=(0, 10), sticky="ew")
                    memoria_frame.grid_columnconfigure(1, weight=1)

                    # Calculate dynamic row configuration based on number of memorias
                    num_memorias = len(data['matching_memorias'])
                    
                    # Use a grid layout for memorias that adapts to the count
                    columns = min(3, num_memorias)  # Maximum 3 columns
                    if columns > 0:
                        for col in range(columns):
                            memoria_frame.grid_columnconfigure(col, weight=1)
                    
                    # Place memorias in a grid layout
                    for j, memoria in enumerate(data['matching_memorias']):
                        # Calculate row and column in the grid
                        row = j // columns
                        col = j % columns
                        
                        # Create a frame for each memoria
                        memoria_item_frame = ctk.CTkFrame(memoria_frame)
                        memoria_item_frame.grid(row=row, column=col, padx=5, pady=5, sticky="ew")
                        memoria_item_frame.grid_columnconfigure(1, weight=1)
                        
                        # Load the memoria image
                        img_label = None
                        try:
                            memoria_name = memoria['name']
                            
                            # Create a larger fixed-size container for the image
                            img_container = ctk.CTkFrame(memoria_item_frame, width=60, height=60, fg_color="transparent")
                            img_container.grid(row=0, column=0, padx=5, pady=2)
                            img_container.grid_propagate(False)  # Prevent the frame from resizing to fit contents
                            
                            # Add a placeholder while the image loads
                            placeholder = ctk.CTkLabel(img_container, text="...", fg_color="transparent")
                            placeholder.place(relx=0.5, rely=0.5, anchor="center")
                            
                            # Schedule the image to be loaded and displayed
                            # Stagger loading to improve performance - load first visible ones immediately
                            delay = 10 if i < 5 else (50 * j + 500 * (i - 5))  # Prioritize first 5 rows, delay others
                            results_window.after(delay, lambda m=memoria_name, c=img_container: 
                                               self._display_memoria_image(m, c))
                            
                            img_label = True  # Mark that we have an image placeholder
                            
                        except Exception as e:
                            print(f"Error setting up image: {str(e)}")

                        # Memoria name and score
                        name_label = ctk.CTkLabel(
                            memoria_item_frame,
                            text=memoria['name'],
                            anchor="w"
                        )
                        name_label.grid(row=0, column=1 if img_label else 0, padx=5, pady=2, sticky="w")

                        score_label = ctk.CTkLabel(
                            memoria_item_frame,
                            text=f"Value: {memoria['score']}",
                            text_color=("#00AA00" if memoria['score'] >= 20 else 
                                        "#AAAA00" if memoria['score'] >= 5 else "#666666"),
                        )
                        score_label.grid(row=0, column=2, padx=5, pady=2, sticky="e")
                else:
                    # If no memorias matched, show a compact message
                    no_memoria_label = ctk.CTkLabel(
                        card,
                        text="No memorias matched",
                        font=ctk.CTkFont(size=12),
                        text_color="#666666",
                        anchor="w"
                    )
                    no_memoria_label.grid(row=1, column=1, columnspan=2, padx=5, pady=(0, 10), sticky="w")

            # If no results, show a message
            if not sorted_results:
                no_results_label = ctk.CTkLabel(
                    content_frame,
                    text="No memoria match results found. Run 'Match Memorias in Screenshots' first.",
                    font=ctk.CTkFont(size=14),
                    anchor="center"
                )
                no_results_label.grid(row=0, column=0, padx=20, pady=20)

        except Exception as e:
            self.log(f"Error displaying results: {str(e)}")
            
    def _preload_memoria_images(self, sorted_results):
        """Preload all memoria images in the background to improve performance"""
        global MEMORIA_IMAGE_CACHE
        
        # Collect all unique memoria names
        memoria_names = set()
        for _, data in sorted_results:
            for memoria in data['matching_memorias']:
                memoria_names.add(memoria['name'])
        
        # Load all images into cache
        for memoria_name in memoria_names:
            try:
                memoria_path = os.path.join('memorias', f"{memoria_name}.png")
                if os.path.exists(memoria_path) and memoria_name not in MEMORIA_IMAGE_CACHE:
                    img = Image.open(memoria_path)
                    MEMORIA_IMAGE_CACHE[memoria_name] = img
            except Exception as e:
                print(f"Error preloading image {memoria_name}: {str(e)}")
                
    def _display_memoria_image(self, memoria_name, container):
        """Display a memoria image in the given container"""
        global MEMORIA_IMAGE_CACHE
        
        try:
            # Clear existing children
            for widget in container.winfo_children():
                widget.destroy()
                
            # Get the image from cache or load it
            if memoria_name in MEMORIA_IMAGE_CACHE:
                img = MEMORIA_IMAGE_CACHE[memoria_name]
            else:
                memoria_path = os.path.join('memorias', f"{memoria_name}.png")
                if os.path.exists(memoria_path):
                    img = Image.open(memoria_path)
                    MEMORIA_IMAGE_CACHE[memoria_name] = img
                else:
                    # If image doesn't exist, show a placeholder
                    placeholder = ctk.CTkLabel(container, text="?", fg_color="transparent")
                    placeholder.place(relx=0.5, rely=0.5, anchor="center")
                    return
            
            # Get the original dimensions
            orig_width, orig_height = img.size
            
            # Calculate the aspect ratio
            aspect_ratio = orig_width / orig_height
            
            # Use fixed container dimensions since winfo_width/height might not be ready
            container_width = 60
            container_height = 60
            
            # Calculate the best fit dimensions that preserve aspect ratio
            if aspect_ratio > 1:  # Wider than tall
                new_width = min(container_width, int(container_height * aspect_ratio))
                new_height = int(new_width / aspect_ratio)
            else:  # Taller than wide or square
                new_height = min(container_height, int(container_width / aspect_ratio))
                new_width = int(new_height * aspect_ratio)
            
            # Ensure dimensions are at least 1 pixel
            new_width = max(1, new_width)
            new_height = max(1, new_height)
            
            # Create the CTkImage with the calculated dimensions
            ctk_img = ctk.CTkImage(
                light_image=img,
                dark_image=img,
                size=(new_width, new_height)
            )
            
            # Create and place the image label
            img_label = ctk.CTkLabel(container, text="", image=ctk_img, fg_color="transparent")
            img_label.place(relx=0.5, rely=0.5, anchor="center")
            
        except Exception as e:
            print(f"Error displaying image {memoria_name}: {str(e)}")
            # Show error placeholder
            error_label = ctk.CTkLabel(container, text="!", fg_color="transparent", text_color="red")
            error_label.place(relx=0.5, rely=0.5, anchor="center")

    def export_results_to_csv(self, results):
        """Export results to a CSV file"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"memoria_results_{timestamp}.csv"
            
            with open(filename, 'w', newline='') as csvfile:
                writer = csv.writer(csvfile)
                
                # Write header
                writer.writerow(['Rank', 'Email', 'Total Score', 'Memorias'])
                
                # Sort results by score
                sorted_results = sorted(
                    [(email, data) for email, data in results.items()],
                    key=lambda x: x[1]['score'],
                    reverse=True
                )
                
                # Write data
                for i, (email, data) in enumerate(sorted_results):
                    memorias = ", ".join([f"{m['name']} ({m['score']})" for m in data['matching_memorias']])
                    writer.writerow([i+1, email, data['score'], memorias])
            
            self.log(f"Results exported to {filename}")
        except Exception as e:
            self.log(f"Error exporting to CSV: {str(e)}")

    def export_results_to_text(self, results):
        """Export results to a text file"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"memoria_results_{timestamp}.txt"
            
            with open(filename, 'w') as txtfile:
                txtfile.write("MEMORIA MATCH RESULTS\n")
                txtfile.write("====================\n\n")
                
                # Sort results by score
                sorted_results = sorted(
                    [(email, data) for email, data in results.items()],
                    key=lambda x: x[1]['score'],
                    reverse=True
                )
                
                # Write data
                for i, (email, data) in enumerate(sorted_results):
                    txtfile.write(f"#{i+1} - {email} (Score: {data['score']})\n")
                    txtfile.write("Memorias:\n")
                    for memoria in data['matching_memorias']:
                        txtfile.write(f"  - {memoria['name']} (Value: {memoria['score']})\n")
                    txtfile.write("\n")
            
            self.log(f"Results exported to {filename}")
        except Exception as e:
            self.log(f"Error exporting to text: {str(e)}")

if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")
    
    app = MacroGUI()
    app.mainloop()
