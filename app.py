import ttkbootstrap as tb
import customtkinter as ctk
import tableauserverclient as TSC
from tkinter import messagebox, filedialog
import threading
import os
from collections import defaultdict
from tableau_fetch import TableauFetcher
from process_elite_and_larkin import process_excel_file
from oldest_dos import get_oldest_dos
from process_concord import process_concord

# Configure CustomTkinter appearance
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

class TableauApp(tb.Window):
    def __init__(self):
        super().__init__(themename="cyborg")
        # Hide main window until login succeeds
        self.withdraw()

        credentials = {}
        # -- Login Dialog --
        login = ctk.CTkToplevel(self)
        login.title("Login to Tableau")
        login.geometry("800x350")
        login.resizable(False, False)
        login.grab_set()

        container = ctk.CTkFrame(login)
        container.pack(expand=True, fill="both", padx=20, pady=20)

        label_font  = ("Segoe UI", 14)
        button_font = ("Segoe UI", 14, "bold")

        ctk.CTkLabel(container, text="Username:", font=label_font).pack(pady=(0,5))
        username_entry = ctk.CTkEntry(container, width=300, height=35, corner_radius=8)
        username_entry.pack(pady=(0,10))

        ctk.CTkLabel(container, text="Password:", font=label_font).pack(pady=(0,5))
        password_entry = ctk.CTkEntry(container, width=300, height=35, show="*", corner_radius=8)
        password_entry.pack(pady=(0,10))

        error_label = ctk.CTkLabel(container, text="", text_color="#FF5555", font=("Segoe UI", 10, "bold"))
        error_label.pack(pady=(0,10))

        def clear_error(event=None):
            error_label.configure(text="")
        username_entry.bind("<Key>", clear_error)
        password_entry.bind("<Key>", clear_error)
        login.bind('<Return>', lambda e: submit())

        def submit():
            user = username_entry.get().strip()
            pw   = password_entry.get().strip()
            if not user or not pw:
                error_label.configure(text="Username and password are required.")
                return
            try:
                auth   = TSC.TableauAuth(user, pw, '')
                server = TSC.Server("https://tableau.blitzmedical.com", use_server_version=True)
                with server.auth.sign_in(auth):
                    credentials['username'] = user
                    credentials['password'] = pw
                    login.destroy()
            except Exception:
                error_label.configure(text="Invalid login credentials. Please try again.")

        ctk.CTkButton(container, text="Login", width=200, height=40, corner_radius=20,
                       font=button_font, command=submit).pack(pady=(10,0))

        self.wait_window(login)
        if not credentials:
            self.destroy()
            return

        # -- Main Window --
        self.deiconify()
        self.title("Census Reconciliation Tool")
        self.geometry("1050x600")
        self.resizable(False, False)

        self.encounter_lookup = defaultdict(lambda: defaultdict(list))
        self.df_tableau = None
        self.uploaded_file_path = None
        self.fetcher = TableauFetcher(
            username=credentials['username'],
            password=credentials['password'],
            output_callback=self.append_output,
            progress_callback=self.update_progress
        )

        self.site_choice = tb.StringVar(value="Select Client")
        tb.Combobox(self, textvariable=self.site_choice,
                    values=["Larkin","Elite","Concord"], state="readonly", width=20).pack(pady=10)

        self.progress = tb.Progressbar(self, mode="determinate", bootstyle="success", length=800)
        self.progress.pack(pady=10)

        self.output_text = tb.ScrolledText(self, height=13, width=80, font=("Consolas", 9))
        self.output_text.pack(pady=10)

        btn_frame = tb.Frame(self)
        btn_frame.pack(pady=(0,10))

        # Single button to select file and kick off all processing
        self.upload_btn = tb.Button(btn_frame, text="Upload & Process File",
                                    bootstyle="primary", command=self.upload_file) 
        self.upload_btn.pack(side="left", padx=5)

        self.process_btn = tb.Button(
            btn_frame,
            text="Process Tableau Data",
            bootstyle="success",
            command=self.start_processing
        )
        self.process_btn.pack(side="left", padx=5)

    def append_output(self, text):
        self.output_text.after(0, lambda: self.output_text.insert("end", text))

    def update_progress(self, val):
        self.progress.after(0, lambda: self.progress.config(value=val))

    def fetch_tableau_data(self, date):
        self.output_text.delete("1.0", "end")
        self.append_output("Connecting to Tableau...\n")
        self.progress.configure(mode="determinate", maximum=100)
        self.progress['value'] = 0
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
                # Pass patient-name list to fetcher
                df = self.fetcher.fetch_data(license_key, filter_values=date)
                if df is not None:
                    self.df_tableau = df
                    self.encounter_lookup = self.fetcher.encounter_lookup
                    self.append_output("\nFetch complete.\n")
                else:
                    self.append_output("\nFetch failed.\n")
            finally:
                self.upload_btn.configure(state="normal")

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
            date = get_oldest_dos(self.uploaded_file_path)
            # Once names are ready, trigger fetch
            self.after(0, lambda: self.fetch_tableau_data(date))

        threading.Thread(target=worker, daemon=True).start()



    def process_file(self, license_key):
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
                process_concord(self.df_tableau, file_path)

        threading.Thread(target=worker, daemon=True).start()

if __name__ == "__main__":
    app = TableauApp()
    app.mainloop()
