import ttkbootstrap as tb
from ttkbootstrap.constants import *
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox, filedialog, simpledialog
import os
import threading
from collections import defaultdict
from tableau_fetch import TableauFetcher
import tableauserverclient as TSC
import tableau_fetch
from excel_processing import process_excel_file
from config import TABLEAU_SERVER, SITE_ID

class TableauApp(tb.Window):
    def __init__(self):
        super().__init__(themename="cyborg")

        self.withdraw()

        login = tk.Toplevel()
        login.title("Login to Tableau")
        login.geometry("1050x600")
        login.resizable(False, False)
        login.grab_set()

        # Centering container frame
        container = tk.Frame(login)
        container.place(relx=0.5, rely=0.5, anchor='center')

        label_font = ("Segoe UI", 14)
        entry_font = ("Segoe UI", 14)
        button_font = ("Segoe UI", 14, "bold")

        tk.Label(container, text="Username:", font=label_font).pack(pady=(15, 5))
        username_entry = tk.Entry(container, width=40, font=entry_font)
        username_entry.pack()

        tk.Label(container, text="Password:", font=label_font).pack(pady=(15, 5))
        password_entry = tk.Entry(container, show="*", width=40, font=entry_font)
        password_entry.pack()

        credentials = {}

        style = ttk.Style()
        style.configure("Error.TLabel", foreground="red", font=("Segoe UI", 8, "bold"))

        # Create label using style
        error_label = ttk.Label(container, text="", style="Error.TLabel")
        error_label.pack(pady=(10, 10))
        
        def clear_error(event=None):
            error_label.config(text="")

        username_entry.bind("<Key>", clear_error)
        password_entry.bind("<Key>", clear_error)
        login.bind('<Return>', lambda event: submit())

        def submit():
            username = username_entry.get().strip()
            password = password_entry.get().strip()

            if not username or not password:
                error_label.config(text="Username and password are required.")
                return

            try:
                tableau_auth = TSC.TableauAuth(username, password, SITE_ID) 
                server = TSC.Server(TABLEAU_SERVER, use_server_version=True)

                with server.auth.sign_in(tableau_auth):
                    credentials["username"] = username
                    credentials["password"] = password
                    login.destroy()

            except Exception as e:
                error_label.config(text="Invalid login credentials. Please try again.")

        tk.Button(container, text="Login", width=20, height=2, font=button_font, command=submit).pack(pady=(20, 5))


        self.wait_window(login)

        if not credentials:
            self.destroy()
            return
        
        self.deiconify() 
        
        self.title("Census Reconciliation Tool")
        self.geometry("1050x600")
        self.resizable(False, False)

        self.encounter_lookup = defaultdict(lambda: defaultdict(list))
        self.df_tableau = None

        # Create TableauFetcher instance with UI callbacks
        self.fetcher = TableauFetcher(
            username=credentials["username"],
            password=credentials["password"],
            output_callback=self.append_output,
            progress_callback=self.update_progress
        )

        # Dropdown
        self.site_choice = tb.StringVar(value="Select Client")
        self.site_dropdown = tb.Combobox(
            self,
            textvariable=self.site_choice,
            values=["Larkin", "Elite"],
            state="readonly",
            width=20
        )
        self.site_dropdown.pack(pady=10)

        # Progress bar
        self.progress = tb.Progressbar(
            self,
            mode="determinate",
            bootstyle="success",
            length=800
        )
        self.progress.pack(pady=10)
    
        # Output text area
        self.output_text = tb.ScrolledText(self, height=13, width=80, font=("Consolas", 9))
        self.output_text.pack(pady=10)

        self.fetch_btn = tb.Button(self, text="Fetch Tableau View")
        self.fetch_btn.pack(pady=(0, 10))

        self.upload_btn = tb.Button(self, text="Upload Excel File")
        self.upload_btn.pack(pady=(0, 10))

        self.fetch_btn.configure(command=self.fetch_tableau_data)
        self.upload_btn.configure(command=self.upload_file_with_license)

    def append_output(self, text):
        self.output_text.after(0, lambda: self.output_text.insert("end", text))

    def update_progress(self, val):
        self.progress.after(0, lambda: self.progress.config(value=val))

    def fetch_tableau_data(self):
        self.output_text.delete("1.0", "end")
        self.output_text.insert("end", "Connecting to Tableau...\n")
        self.update()

        site = self.site_choice.get()
        if site == "Larkin":
            license_key = "137797"
        elif site == "Elite":
            license_key = "160214"
        else:
            self.output_text.insert("end", "Please select a site (Larkin or Elite).\n")
            return

        self.progress.configure(mode="determinate", maximum=100)
        self.progress['value'] = 0
        self.progress.update()

        def run():
            df = self.fetcher.fetch_data(license_key)
            if df is not None:
                self.df_tableau = df
                self.encounter_lookup = self.fetcher.encounter_lookup

        threading.Thread(target=run, daemon=True).start()

    def upload_file_with_license(self):
        site = self.site_choice.get()
        if site == "Larkin":
            license_key = "137797"
        elif site == "Elite":
            license_key = "160214"
        else:
            self.output_text.insert("end", "Please select a site (Larkin or Elite).\n")
            return

        self.upload_file(license_key)

    def upload_file(self, license_key):
        if self.encounter_lookup is None:
            messagebox.showwarning("Data Missing", "Please fetch Tableau data before uploading Excel file.")
            return

        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return
        tableau_fetcher = self.fetcher
        def run_process():
            processed_path = process_excel_file(
                file_path,
                license_key,
                encounter_lookup=self.encounter_lookup,
                df_tableau=self.df_tableau,
                output_callback=self.append_output,
                tableau_fetcher=tableau_fetcher,
            )
            if processed_path:
                if messagebox.askyesno("Open File", f"Processed file saved:\n{processed_path}\n\nDo you want to open it?"):
                    os.startfile(processed_path)

        threading.Thread(target=run_process, daemon=True).start()

if __name__ == "__main__":
    app = TableauApp()
    app.mainloop()
