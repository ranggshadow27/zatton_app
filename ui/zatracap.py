import customtkinter as ctk
import tkinter as tk
from tkcalendar import DateEntry
from tkinter import filedialog, messagebox
from datetime import datetime
import threading

# Import dari core
from core.utils import check_license
from core.automation_core import initialize_driver, is_driver_alive, main

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class ZabbixAutomationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Zatton App")
        self.root.geometry("380x650")
        self.root.minsize(380, 650)
        self.root.maxsize(380, 650)

        try:
            self.root.iconbitmap("mhg-zabbix-capture.ico")
        except tk.TclError as e:
            print(f"Error setting icon: {e}. Using default icon.")

        self.excel_file = tk.StringVar(value="")
        self.save_folder = tk.StringVar(value="")
        self.is_running = False
        self.time_running = False
        self.start_time = None
        self.driver = None                     # ← tambahan driver management

        self.create_widgets()

    def create_widgets(self):
        # SEMUA WIDGET SAMA PERSIS KODE ASLI LU (aku copy paste full)
        title_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        title_frame.pack(pady=(20, 10), padx=20, fill="x")
        
        title_label = ctk.CTkLabel(title_frame, text="•   Zatton App   •", font=("Inter", 14, "bold"))
        title_label.pack(pady=2)
        
        title_label = ctk.CTkLabel(title_frame, text="⚡ Zabbix Automation | Router Traffic Capture ⚡", font=("Inter", 12))
        title_label.pack(pady=0)

        # Excel File Selection (sama persis)
        excel_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        excel_frame.pack(pady=1, padx=20, fill="x")
        excel_label = ctk.CTkLabel(excel_frame, text="Pick an Excel File:", font=("Inter", 12))
        excel_label.pack(anchor="w")
        self.excel_entry = ctk.CTkEntry(excel_frame, textvariable=self.excel_file, width=240, font=("Inter", 12))
        self.excel_entry.pack(side="left", padx=(0, 10), pady=5)
        excel_browse = ctk.CTkButton(excel_frame, text="Browse", command=self.browse_excel, font=("Inter", 12), fg_color="#535353", hover_color="#676767")
        excel_browse.pack(side="right")

        # Save Folder Selection (sama persis)
        folder_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        folder_frame.pack(pady=1, padx=20, fill="x")
        folder_label = ctk.CTkLabel(folder_frame, text="Select Save Folder:", font=("Inter", 12))
        folder_label.pack(anchor="w")
        self.folder_entry = ctk.CTkEntry(folder_frame, textvariable=self.save_folder, width=240, font=("Inter", 12))
        self.folder_entry.pack(side="left", padx=(0, 10), pady=5)
        folder_browse = ctk.CTkButton(folder_frame, text="Browse", command=self.browse_folder, font=("Inter", 12), fg_color="#535353", hover_color="#676767")
        folder_browse.pack(side="right")

        # Date Selection (sama persis)
        date_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        date_frame.pack(pady=10, padx=20, fill="x")
        
        start_date_subframe = ctk.CTkFrame(date_frame, fg_color="transparent")
        start_date_subframe.grid(row=0, column=0, padx=(0, 5), sticky="w")
        start_date_label = ctk.CTkLabel(start_date_subframe, text="Start Date:", font=("Inter", 12))
        start_date_label.pack(anchor="w")
        current_month_start = datetime.now().replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        self.start_date_entry = DateEntry(start_date_subframe, width=50, font=("Inter", 12), date_pattern="dd/mm/yyyy")
        self.start_date_entry.set_date(current_month_start)
        self.start_date_entry.pack(pady=5)

        end_date_subframe = ctk.CTkFrame(date_frame, fg_color="transparent")
        end_date_subframe.grid(row=0, column=1, padx=(5, 0), sticky="e")
        end_date_label = ctk.CTkLabel(end_date_subframe, text="End Date:", font=("Inter", 12))
        end_date_label.pack(anchor="w")
        today = datetime.now()
        self.end_date_entry = DateEntry(end_date_subframe, width=50, font=("Inter", 12), date_pattern="dd/mm/yyyy")
        self.end_date_entry.set_date(today)
        self.end_date_entry.pack(pady=5)

        date_frame.grid_columnconfigure(0, weight=1)
        date_frame.grid_columnconfigure(1, weight=1)

        # Progress & Time
        progress_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        progress_frame.pack(pady=5, padx=20, fill="x")
        self.progress_label = ctk.CTkLabel(progress_frame, text="Ready to Start", font=("Inter", 12))
        self.progress_label.pack(anchor="w")
        self.time_elapsed_label = ctk.CTkLabel(progress_frame, text="Time Elapsed: 0:00:00:00", font=("Inter", 12))
        self.time_elapsed_label.pack(anchor="w")
        self.time_elapsed_label.pack_forget()
        self.progress_bar = ctk.CTkProgressBar(progress_frame, width=350)
        self.progress_bar.set(0)
        self.progress_bar.pack(pady=5)

        # Log
        log_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        log_frame.pack(pady=5, padx=20, fill="both", expand=True)
        self.log_textbox = ctk.CTkTextbox(log_frame, height=150, font=("Inter", 12))
        self.log_textbox.pack(fill="both", expand=True)
        self.log_textbox.configure(state="disabled")

        # Start Button
        start_button_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        start_button_frame.pack(pady=5, padx=20, fill="x")
        self.start_button = ctk.CTkButton(start_button_frame, text="Start Automation", command=self.start_automation_thread, font=("Inter", 12), height=32)
        self.start_button.pack(pady=(10, 2))

        watermark_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        watermark_frame.pack(pady=0, padx=10, fill="x")
        watermark_label = ctk.CTkLabel(watermark_frame, text="©2025 ratipray • MHG NOC Software", font=("Inter", 10, "bold"), text_color="#505050")
        watermark_label.pack(pady=(0,10))

    def browse_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.excel_file.set(file_path)

    def browse_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.save_folder.set(folder_path)

    def format_date(self, date_obj):
        return date_obj.strftime("%d %B %Y")

    def validate_dates(self, start_date, end_date):
        try:
            start = datetime.strptime(start_date, "%d %B %Y")
            end = datetime.strptime(end_date, "%d %B %Y")
            if start > end:
                messagebox.showerror("Error", "Start date cannot be later than end date")
                return False
            return True
        except ValueError as e:
            messagebox.showerror("Error", f"Invalid date format: {e}")
            return False

    def add_log(self, message):
        self.root.after(0, self._add_log, message)

    def _add_log(self, message):
        self.log_textbox.configure(state="normal")
        lines = self.log_textbox.get("1.0", "end").strip().split("\n")
        if len(lines) >= 20:
            self.log_textbox.delete("1.0", "2.0")
        self.log_textbox.insert("end", message + "\n")
        self.log_textbox.see("end")
        self.log_textbox.configure(state="disabled")

    def update_progress(self, current, total, site_id):
        self.root.after(0, self._update_progress, current, total, site_id)

    def _update_progress(self, current, total, site_id):
        progress = current / total if total > 0 else 0
        self.progress_bar.set(progress)
        self.progress_label.configure(text=f"Progress: {current}/{total} Capturing site {site_id}")

    def update_time_elapsed(self):
        if self.time_running and self.start_time:
            elapsed = datetime.now() - self.start_time
            total_seconds = int(elapsed.total_seconds())
            days = total_seconds // (24 * 3600)
            hours = (total_seconds % (24 * 3600)) // 3600
            minutes = (total_seconds % 3600) // 60
            seconds = total_seconds % 60
            time_str = f"{days}:{hours:02d}:{minutes:02d}:{seconds:02d}"
            self.time_elapsed_label.configure(text=f"Time Elapsed: {time_str}")
            self.root.after(1000, self.update_time_elapsed)
            return time_str
        return "0:00:00:00"

    def start_automation_thread(self):
        if self.is_running:
            messagebox.showwarning("Warning", "Automation is already running!")
            return

        # === LICENSE CHECK (baru) ===
        success, msg = check_license()
        if not success:
            messagebox.showerror("License Error", msg)
            return

        excel_file = self.excel_file.get()
        save_folder = self.save_folder.get()
        start_date = self.format_date(self.start_date_entry.get_date())
        end_date = self.format_date(self.end_date_entry.get_date())

        if not excel_file or not save_folder:
            messagebox.showerror("Error", "Please select Excel file and save folder")
            return

        if not self.validate_dates(start_date, end_date):
            return

        # === DRIVER REUSE LOGIC (baru) ===
        if self.driver is None or not is_driver_alive(self.driver):
            self.driver = initialize_driver()
            self.add_log("> WebDriver baru diinisialisasi")
        else:
            self.add_log("> Menggunakan WebDriver yang sudah terbuka (reuse)")

        self.is_running = True
        self.time_running = True
        self.start_time = datetime.now()
        self.time_elapsed_label.pack(anchor="w")
        self.update_time_elapsed()
        self.start_button.configure(state="disabled")

        threading.Thread(target=self.run_automation, 
                         args=(excel_file, save_folder, start_date, end_date), 
                         daemon=True).start()

    def run_automation(self, excel_file, save_folder, start_date, end_date):
        try:
            self.add_log("> Starting Automation..")
            main(excel_file, save_folder, start_date, end_date, self.add_log, self.update_progress, self.driver)
            self.root.after(0, lambda: messagebox.showinfo("Success", "Automation completed!"))
            time.sleep(0.25)
            time_str = self.update_time_elapsed()
            self.add_log(f"> Time Elapsed: {time_str}")
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", f"Automation failed: {e}"))
            time.sleep(0.25)
            time_str = self.update_time_elapsed()
            self.add_log(f"> Time Elapsed: {time_str}")
        finally:
            self.is_running = False
            self.time_running = False
            self.root.after(0, lambda: self.start_button.configure(state="normal"))
            self.root.after(0, lambda: self.time_elapsed_label.pack_forget())
            # TIDAK ADA driver.quit() lagi → sesuai request lu

if __name__ == "__main__":
    root = ctk.CTk()
    app = ZabbixAutomationApp(root)
    root.mainloop()