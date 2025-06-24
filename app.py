import ttkbootstrap as tb
import customtkinter as ctk
import tableauserverclient as TSC
from tkinter import messagebox, filedialog
import threading
import os
from collections import defaultdict
from tableau_fetch import TableauFetcher
from excel_processing import process_excel_file

# Configure CustomTkinter appearance
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

class TableauApp(tb.Window):
    def __init__(self):
        super().__init__(themename="cyborg")
        # Hide main window until login succeeds
        self.withdraw()

        credentials = {}
        # Create a CustomTkinter login dialog
        login = ctk.CTkToplevel(self)
        login.title("Login to Tableau")
        login.geometry("800x350")
        login.resizable(False, False)
        login.grab_set()

        # Container frame for inputs
        container = ctk.CTkFrame(login)
        container.pack(expand=True, fill="both", padx=20, pady=20)

        label_font  = ("Segoe UI", 14)
        button_font = ("Segoe UI", 14, "bold")

        # Username
        ctk.CTkLabel(container, text="Username:", font=label_font).pack(pady=(0,5))
        username_entry = ctk.CTkEntry(
            container,
            width=300, height=35,
            corner_radius=8
        )
        username_entry.pack(pady=(0,10))

        # Password
        ctk.CTkLabel(container, text="Password:", font=label_font).pack(pady=(0,5))
        password_entry = ctk.CTkEntry(
            container,
            width=300, height=35,
            corner_radius=8,
            show="*"
        )
        password_entry.pack(pady=(0,10))

        # Error label
        error_label = ctk.CTkLabel(
            container,
            text="",
            text_color="#FF5555",
            font=("Segoe UI", 10, "bold")
        )
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
                # Attempt Tableau sign-in
                auth   = TSC.TableauAuth(user, pw, '')
                server = TSC.Server(
                    "https://tableau.blitzmedical.com",
                    use_server_version=True
                )
                with server.auth.sign_in(auth):
                    credentials['username'] = user
                    credentials['password'] = pw
                    login.destroy()
            except Exception:
                error_label.configure(text="Invalid login credentials. Please try again.")

        # Login button
        ctk.CTkButton(
            container,
            text="Login",
            width=200, height=40,
            corner_radius=20,
            font=button_font,
            command=submit
        ).pack(pady=(10,0))

        # Wait until login dialog closes
        self.wait_window(login)
        # If no credentials, exit
        if not credentials:
            self.destroy()
            return

        # Show main window
        self.deiconify()
        self.title("Census Reconciliation Tool")
        self.geometry("1050x600")
        self.resizable(False, False)

        # Initialize data structures
        self.encounter_lookup = defaultdict(lambda: defaultdict(list))
        self.df_tableau = None
        self.fetcher = TableauFetcher(
            username=credentials['username'],
            password=credentials['password'],
            output_callback=self.append_output,
            progress_callback=self.update_progress
        )

        # Site dropdown
        self.site_choice = tb.StringVar(value="Select Client")
        tb.Combobox(
            self,
            textvariable=self.site_choice,
            values=["Larkin","Elite","Concord"],
            state="readonly",
            width=20
        ).pack(pady=10)

        # Progress bar
        self.progress = tb.Progressbar(
            self,
            mode="determinate",
            bootstyle="success",
            length=800
        )
        self.progress.pack(pady=10)

        # Output text area
        self.output_text = tb.ScrolledText(
            self,
            height=13, width=80,
            font=("Consolas", 9)
        )
        self.output_text.pack(pady=10)

        # Fetch & Upload buttons
        btn_frame = tb.Frame(self)
        btn_frame.pack(pady=(0,10))
        self.fetch_btn = tb.Button(
            btn_frame,
            text="Fetch Tableau View",
            bootstyle="primary",
            command=self.fetch_tableau_data
        )
        self.fetch_btn.pack(side="left", padx=5)
        self.upload_btn = tb.Button(
            btn_frame,
            text="Upload Excel File",
            bootstyle="primary",
            command=self.upload_file_with_license
        )
        self.upload_btn.pack(side="left", padx=5)

    def append_output(self, text):
        self.output_text.after(0, lambda: self.output_text.insert("end", text))

    def update_progress(self, val):
        self.progress.after(0, lambda: self.progress.config(value=val))

    def fetch_tableau_data(self):
        # Clear output and reset progress
        self.output_text.delete("1.0", "end")
        self.append_output("Connecting to Tableau...\n")
        self.progress.configure(mode="determinate", maximum=100)
        self.progress['value'] = 0
        self.fetch_btn.configure(state="disabled")

        site = self.site_choice.get()
        if site == "Larkin":
            license_key = "137797"
        elif site == "Elite":
            license_key = "160214"
        elif site == "Concord":
            license_key = ""
        else:
            messagebox.showwarning(
                "Site Required",
                "Please select a site before fetching."
            )
            self.fetch_btn.configure(state="normal")
            return

        def worker():
            try:
                df = self.fetcher.fetch_data(license_key)
                if df is not None:
                    self.df_tableau = df
                    self.encounter_lookup = self.fetcher.encounter_lookup
                    self.append_output("\nFetch complete.\n")
                else:
                    self.append_output("\nFetch failed.\n")
            finally:
                self.fetch_btn.configure(state="normal")

        threading.Thread(target=worker, daemon=True).start()

    def upload_file_with_license(self):
        site = self.site_choice.get()
        if site == "Larkin":
            license_key = "137797"
        elif site == "Elite":
            license_key = "160214"
        elif site == "Concord":
            license_key = "159127"
        else:
            messagebox.showwarning(
                "Site Required",
                "Please select a site before uploading."
            )
            return
        self.upload_file(license_key)

    def upload_file(self, license_key):
        if self.encounter_lookup is None:
            messagebox.showwarning(
                "Data Missing",
                "Please fetch Tableau data before uploading Excel file."
            )
            return

        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not file_path:
            return

        def worker():
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

        threading.Thread(target=worker, daemon=True).start()

if __name__ == "__main__":
    app = TableauApp()
    app.mainloop()
