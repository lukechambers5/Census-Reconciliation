import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import messagebox, filedialog
import os
import threading
from collections import defaultdict
from tableau_fetch import TableauFetcher
import tableau_fetch
from excel_processing import process_excel_file

class TableauApp(tb.Window):
    def __init__(self):
        super().__init__(themename="cyborg")
        self.title("Census Reconciliation Tool")
        self.geometry("500x550")
        self.resizable(False, False)

        self.encounter_lookup = defaultdict(lambda: defaultdict(list))
        self.df_tableau = None

        # Create TableauFetcher instance with UI callbacks
        self.fetcher = TableauFetcher(
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
            length=300
        )
        self.progress.pack(pady=10)
    
        # Output text area
        self.output_text = tb.ScrolledText(self, height=13, width=55, font=("Consolas", 9))
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
                tableau_fetcher=tableau_fetcher
            )
            if processed_path:
                if messagebox.askyesno("Open File", f"Processed file saved:\n{processed_path}\n\nDo you want to open it?"):
                    os.startfile(processed_path)

        threading.Thread(target=run_process, daemon=True).start()

if __name__ == "__main__":
    app = TableauApp()
    app.mainloop()
