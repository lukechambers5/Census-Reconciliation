import ttkbootstrap as tb
import customtkinter as ctk
from customtkinter import CTkImage
import tableauserverclient as TSC
from tkinter import messagebox, filedialog
import threading
import os
import sys
from collections import defaultdict
from tableau_fetch import TableauFetcher
from process_elite_and_larkin import process_excel_file
from oldest_dos import get_oldest_dos
from process_concord import process_concord
from PIL import Image, ImageTk, ImageSequence

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

def get_resource_path(relative_path):
    """ Get absolute path to resource, works for dev and PyInstaller """
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

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
        self.username_entry = ctk.CTkEntry(container, width=400, height=45)
        self.username_entry.pack(pady=(0,10))

        ctk.CTkLabel(container, text="Password:", font=label_font).pack(pady=(0,5))
        self.password_entry = ctk.CTkEntry(container, width=400, height=45, show="*")
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
        self.output_text.configure(state="disabled")

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

        # Help label at bottom-left
        self.help_label = ctk.CTkLabel(self.main_frame, text="Need Help?", text_color="#1e90ff", cursor="hand2")
        self.help_label.place(relx=0.9, rely=0.96, anchor="sw")
        self.help_label.bind("<Button-1>", lambda e: self.after(0, self.open_help_window))

    def open_help_window(self):
        help_win = ctk.CTkToplevel(self)
        help_win.title("Help Center")
        help_win.geometry("800x700")  
        help_win.resizable(False, False)

        tabview = ctk.CTkTabview(help_win, width=780, height=660)
        tabview.pack(padx=10, pady=10)

        tabview.add("Elite / Larkin")
        tabview.add("Concord")

        # Elite/Larkin HELP
        elite_larkin_help = ctk.CTkTextbox(tabview.tab("Elite / Larkin"), wrap="word")
        elite_larkin_help.insert("0.0",
            "Requirements:\n"
            "• File must be Excel format (.xlsx or .xls)\n"
            "• Must contain column: Patient Name (Last, First)\n"
            "• Must contain column: Date of Service\n\n"

            "Error Status Meanings:\n"
            "• MISMATCH DOS:\n"
            "   └─ A match was found by name, but the Date of Service does not match Tableau.\n\n"
            "• INVALID CODE IN TABLEAU:\n"
            "   └─ A Tableau match was found, but no valid CPT code was present in Tableau.\n"
            "   └─ Valid CPT codes include: 99XXX, LWBS, AMA\n\n"
            "• NAME NOT FOUND IN TABLEAU:\n"
            "   └─ No matching patient name and DOS combination was found in Tableau.\n\n"

            "Elite:\n"
            "• Handles billing codes like 99XXX, LWBS, AMA, 0, and NULL\n"
            "• Deterines status based on billing code and presence in Tableau\n"
            "• Adds column E&M (Pro) based on billing code logic\n\n"
            "Larkin:\n"
            "• Used more direct lookup and sets reconciliation based on matches and preexisting status values\n"
            "• Also attempts to fill Patient MRn and patient DOB from Tableau if abailable\n\n"
            "If you still aren't getting expected output, look at screenshot of raw file that will process correctly, and make sure your file has same layout before uploading.\n\n"
        )
        elite_larkin_help.insert("end", "📁 View Unprocessed Larkin File Example\n", "larkin_link")
        elite_larkin_help.insert("end", "\n") 
        elite_larkin_help.insert("end", "📁 View Unprocessed Elite File Example\n", "elite_link")

        elite_larkin_help.tag_config("larkin_link", foreground="cyan", underline=True)
        elite_larkin_help.tag_config("elite_link", foreground="cyan", underline=True)

        elite_larkin_help.tag_bind("larkin_link", "<Enter>", lambda e: elite_larkin_help.configure(cursor="hand2"))
        elite_larkin_help.tag_bind("larkin_link", "<Leave>", lambda e: elite_larkin_help.configure(cursor=""))

        elite_larkin_help.tag_bind("elite_link", "<Enter>", lambda e: elite_larkin_help.configure(cursor="hand2"))
        elite_larkin_help.tag_bind("elite_link", "<Leave>", lambda e: elite_larkin_help.configure(cursor=""))

        def open_larkin_file(event=None):
            file_path = get_resource_path(os.path.join("public", "larkin.png"))
            os.startfile(file_path)

        def open_elite_file(event=None):
            file_path = get_resource_path(os.path.join("public", "elite.png"))
            os.startfile(file_path)

        elite_larkin_help.tag_bind("larkin_link", "<Button-1>", open_larkin_file)
        elite_larkin_help.tag_bind("elite_link", "<Button-1>", open_elite_file)

        elite_larkin_help.pack(expand=True, fill="both", padx=10, pady=10)
        elite_larkin_help.configure(state="disabled")

        # Concord Help 
        concord_help = ctk.CTkTextbox(tabview.tab("Concord"), wrap="word")
        concord_help.insert("0.0",
            "Requirements:\n"
            "• File must be Excel (.xlsx, .xls) or CSV (.csv)\n"
            "• Required columns:\n"
            "   └─ Patient Name (Last, First)\n"
            "   └─ Date of Service\n"
            "   └─ Account Number\n"
            "   └─ Medical Record Number\n\n"

            "Auto-Filtered Out:\n"
            "• Non-Blitz departments are removed automatically.\n\n"

            "Processing Logic:\n"
            "• Three ID columns are created using Date of Service combined with:\n"
            "   └─ Account Number → ID (DOS_ACCT)\n"
            "   └─ Medical Record Number → ID2 (DOS_MRN)\n"
            "   └─ Full Patient Name → ID3 (DOS_Patient Name)\n\n"

            "Matching Logic:\n"
            "• The app attempts to match each row with Tableau data using:\n"
            "   └─ (Last Name, First Name, Date of Service) or\n"
            "   └─ (Last Name, First Name, MRN)\n"
            "• If matched, Tableau fields like Provider, Carrier, and Facility are filled in.\n"
            "• If no match is found, those columns are filled with '#N/A'\n\n"
            "If you still aren't getting expected output, look at screenshot of raw file that will process correctly, and make sure your file has same layout before uploading.\n\n"
        )

        concord_help.insert("end", "📁 View Unprocessed Concord File Example\n", "concord_link")

        concord_help.tag_config("concord_link", foreground="cyan", underline=True)

        concord_help.tag_bind("concord_link", "<Enter>", lambda e: concord_help.configure(cursor="hand2"))
        concord_help.tag_bind("concord_link", "<Leave>", lambda e: concord_help.configure(cursor=""))

        def open_concord_file(event=None):
            file_path = get_resource_path(os.path.join("public", "concord.png"))
            os.startfile(file_path)

        concord_help.tag_bind("concord_link", "<Button-1>", open_concord_file)

        concord_help.pack(expand=True, fill="both", padx=10, pady=10)
        concord_help.configure(state="disabled")

        help_win.mainloop()


    def start_spinner(self):
        self.spinner_label.place(relx=0.71, rely=0.14, anchor="center")
        gif_path = os.path.join(os.path.dirname(__file__), "public", "spinner.gif")

        try:
            pil_image = Image.open(gif_path)
            self.spinner_frames = [
                CTkImage(light_image=frame.convert("RGBA").copy(), size=(25, 25))
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
        self.output_text.after(0, lambda: self._safe_insert_output(text))

    def _safe_insert_output(self, text):
        self.output_text.configure(state="normal")
        self.output_text.insert("end", text)
        self.output_text.configure(state="disabled")

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
        

        def worker():
            self.output_text.delete("1.0", "end")
            self.append_output("Connecting to Tableau...\n")
            self.after(0, self.start_spinner)
            date = get_oldest_dos(self.uploaded_file_path)
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
        
        self.after(0, self.start_spinner)

        def worker():
            try:
                if(license_key != ""):
                    processed_path = process_excel_file(
                        file_path,
                        license_key,
                        encounter_lookup=self.encounter_lookup,
                        df_tableau=self.df_tableau,
                        output_callback=self.append_output,
                        tableau_fetcher=self.fetcher,
                    )
                else:
                    self.append_output("\nProcessing data...\n")
                    processed_path = process_concord(self.df_tableau, file_path)
            finally:
                self.after(0, self.stop_spinner)
                if processed_path:
                        if messagebox.askyesno(
                            "Open File",
                            f"Processed file saved:\n{processed_path}\n\nDo you want to open it?"
                        ):
                            os.startfile(processed_path)

        threading.Thread(target=worker, daemon=True).start()

if __name__ == "__main__":
    app = TableauApp()
    app.mainloop()
