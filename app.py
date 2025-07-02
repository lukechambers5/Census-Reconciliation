import ttkbootstrap as tb
import customtkinter as ctk
from customtkinter import CTkImage
import tableauserverclient as TSC
from tkinter import messagebox, filedialog
import threading
import os
from collections import defaultdict
from tableau_fetch import TableauFetcher
from process_elite_and_larkin import process_excel_file
from oldest_dos import get_oldest_dos
from process_concord import process_concord
from PIL import Image, ImageTk, ImageSequence

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

class TableauApp(tb.Window):
    def __init__(self):
        super().__init__(themename="cyborg")
        self.title("Census Reconciliation Tool")
        self.geometry("1050x600")
        self.resizable(False, False)

        self.credentials = {}

        # Initialize frames
        self.login_frame = ctk.CTkFrame(self)
        self.main_frame = ctk.CTkFrame(self)

        self.build_login_frame()
        self.build_main_frame()

        self.login_frame.pack(fill="both", expand=True)

    def build_login_frame(self):
        container = self.login_frame

        label_font = ("Segoe UI", 14)
        button_font = ("Segoe UI", 14, "bold")

        ctk.CTkLabel(container, text="").pack(pady=35)

        ctk.CTkLabel(container, text="Username:", font=label_font).pack(pady=(20,5))
        self.username_entry = ctk.CTkEntry(container, width=300, height=35)
        self.username_entry.pack(pady=(0,10))

        ctk.CTkLabel(container, text="Password:", font=label_font).pack(pady=(0,5))
        self.password_entry = ctk.CTkEntry(container, width=300, height=35, show="*")
        self.password_entry.pack(pady=(0,10))

        self.error_label = ctk.CTkLabel(container, text="", text_color="#FF5555", font=("Segoe UI", 10, "bold"))
        self.error_label.pack(pady=(0,10))

        ctk.CTkButton(container, text="Login", width=200, height=40, font=button_font,
                    command=self.submit_login).pack(pady=(10,0))
        
        self.username_entry.bind('<Return>', lambda e: self.submit_login())
        self.password_entry.bind('<Return>', lambda e: self.submit_login())

        self.username_entry.bind('<Key>', self.clear_error)
        self.password_entry.bind('<Key>', self.clear_error)
       
    def submit_login(self):
        user = self.username_entry.get().strip()
        pw = self.password_entry.get().strip()
        if not user or not pw:
            self.error_label.configure(text="Username and password are required.")
            return
        try:
            auth = TSC.TableauAuth(user, pw, '')
            server = TSC.Server("https://tableau.blitzmedical.com", use_server_version=True)
            with server.auth.sign_in(auth):
                self.credentials['username'] = user
                self.credentials['password'] = pw
                self.fetcher = TableauFetcher(
                    username=user,
                    password=pw,
                    output_callback=self.append_output,
                    progress_callback=self.update_progress
                )
                self.login_frame.pack_forget()
                self.main_frame.pack(fill="both", expand=True)
        except Exception:
            self.error_label.configure(text="Invalid login credentials. Please try again.")

    def build_main_frame(self):
        self.encounter_lookup = defaultdict(lambda: defaultdict(list))
        self.df_tableau = None
        self.uploaded_file_path = None

        label_font = ("Segoe UI", 14)

        # Select Client Label
        ctk.CTkLabel(self.main_frame, text="Select Client:", font=label_font).pack(pady=(20, 5))

        # Rounded dropdown (OptionMenu)
        self.site_choice = ctk.StringVar(value="Select Client")
        self.client_menu = ctk.CTkOptionMenu(self.main_frame, variable=self.site_choice,
                                            values=["Larkin", "Elite", "Concord"], width=300)
        self.client_menu.pack(pady=(0, 10))

        # Modern progress bar
        self.progress = ctk.CTkProgressBar(self.main_frame, width=675)
        self.progress.set(0)  # Set initial value
        self.progress.pack(pady=10)

        # Modern text box
        self.output_text = ctk.CTkTextbox(self.main_frame, height=250, width=675, font=("Consolas", 13))
        self.output_text.pack(pady=10)

        # Button container with rounded buttons
        btn_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")  # keep transparent for layout only
        btn_frame.pack(pady=(0, 10))

        self.upload_btn = ctk.CTkButton(btn_frame, text="Upload & Process File", width=200,
                                        command=self.upload_file)
        self.upload_btn.pack(side="left", padx=10)

        self.process_btn = ctk.CTkButton(btn_frame, text="Process Tableau Data", width=200,
                                        command=self.start_processing)
        self.process_btn.pack(side="left", padx=10)

        
        self.spinner_label = ctk.CTkLabel(self.main_frame, text="")
        self.spinner_label.place(relx=0.75, rely=0.125, anchor="center")

    def start_spinner(self):
        gif_path = os.path.join(os.path.dirname(__file__), "public", "spinner.gif")

        try:
            pil_image = Image.open(gif_path)
            self.spinner_frames = [
                CTkImage(light_image=frame.convert("RGBA").copy(), size=(50, 50))
                for frame in ImageSequence.Iterator(pil_image)
            ]

            self.spinner_index = 0
            self.spinner_running = True
            self.spinner_label.configure(image=self.spinner_frames[0])

            def animate():
                if self.spinner_running:
                    self.spinner_index = (self.spinner_index + 1) % len(self.spinner_frames)
                    self.spinner_label.configure(image=self.spinner_frames[self.spinner_index])
                    self.after(50, animate)

            animate()

        except FileNotFoundError:
            print(f"[ERROR] Spinner GIF not found at: {gif_path}")

    def stop_spinner(self):
        self.spinner_running = False
        self.spinner_label.place_forget()

    def clear_error(self, event=None):
        self.error_label.configure(text="")

    def append_output(self, text):
        self.output_text.after(0, lambda: self.output_text.insert("end", text))

    def update_progress(self, val):
        self.progress.after(0, lambda: self.progress.set(val))

    def fetch_tableau_data(self, date):
        self.progress.set(0)
        self.upload_btn.configure(state="disabled")

        site = self.site_choice.get()
        if site == "Larkin":
            license_key = "137797"
        elif site == "Elite":
            license_key = "160214"
        elif site == "Concord":
            license_key = ""
        else:
            messagebox.showwarning("Site Required", "Please select a site before fetching.")
            self.upload_btn.configure(state="normal")
            return

        def worker():
            try:
                self.append_output("Connecting to Tableau...\n")
                df = self.fetcher.fetch_data(license_key, filter_values=date)
                if df is not None:
                    self.df_tableau = df
                    self.encounter_lookup = self.fetcher.encounter_lookup
                    self.append_output("Fetch complete.\n")
                else:
                    self.append_output("Fetch failed.\n")
            finally:
                self.upload_btn.configure(state="normal")
                self.after(0, self.stop_spinner)

        threading.Thread(target=worker, daemon=True).start()

    def start_processing(self):
        site = self.site_choice.get()
        if site == "Larkin":
            license_key = "137797"
        elif site == "Elite":
            license_key = "160214"
        elif site == "Concord":
            license_key = ""
        else:
            messagebox.showwarning("Site Required", "Please select a site before processing.")
            return

        self.process_file(license_key)
    
    def upload_file(self):
        self.uploaded_file_path = filedialog.askopenfilename(
            title="Select Patient List",
            filetypes=[
                ("Excel files", ("*.xlsx", "*.xls")),
                ("CSV files",   ("*.csv",)),
                ("All files",   ("*.*",)),
            ]
        )
        if not self.uploaded_file_path:
            return
        
        self.output_text.delete("1.0", "end")

        def worker():
            self.after(0, self.start_spinner)
            self.append_output("Finding oldest date to filter...\n")
            date = get_oldest_dos(self.uploaded_file_path)
            self.append_output(f"Oldest date found: {date}\n") 
            self.after(0, lambda: self.fetch_tableau_data(date))


        threading.Thread(target=worker, daemon=True).start()

    def process_file(self, license_key):
        if(license_key != ""):
            if self.encounter_lookup is None:
                messagebox.showwarning(
                    "Data Missing",
                    "Please fetch Tableau data before uploading Excel file."
                )
                return

        file_path = self.uploaded_file_path
        if not file_path:
            messagebox.showwarning(
                "Missing File",
                "No file has been uploaded yet. Please upload a file first."
            )
            return

        def worker():
            try:
                self.after(0, self.start_spinner)
                if(license_key != ""):
                    processed_path = process_excel_file(
                        file_path,
                        license_key,
                        encounter_lookup=self.encounter_lookup,
                        df_tableau=self.df_tableau,
                        output_callback=self.append_output,
                        tableau_fetcher=self.fetcher,
                    )
                    if processed_path:
                        if messagebox.askyesno(
                            "Open File",
                            f"Processed file saved:\n{processed_path}\n\nDo you want to open it?" 
                        ):
                            os.startfile(processed_path)
                else:
                    self.append_output("\nProcessing data...\n")
                    processed_path = process_concord(self.df_tableau, file_path)

                    if processed_path:
                        if messagebox.askyesno(
                            "Open File",
                            f"Processed file saved:\n{processed_path}\n\nDo you want to open it?"
                        ):
                            os.startfile(processed_path)
            finally:
                self.after(0, self.stop_spinner)

        threading.Thread(target=worker, daemon=True).start()

if __name__ == "__main__":
    app = TableauApp()
    app.mainloop()
