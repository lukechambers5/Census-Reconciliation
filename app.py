import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
import pandas as pd
from io import BytesIO
import tableauserverclient as TSC
import os
from dotenv import load_dotenv
import traceback

load_dotenv()
ACCESS_TOKEN = os.getenv("ACCESS_TOKEN")
VIEW_ID = os.getenv("VIEW_ID")

TABLEAU_SERVER = 'https://tableau.blitzmedical.com'
TOKEN_NAME = "Census"
SITE_ID = ''  # adjust if needed


class TableauApp(tb.Window):
    def __init__(self):
        super().__init__(themename="darkly")
        self.title("Census Reconciliation Tool")
        self.geometry("1000x700")

        self.charge_code_lookup = None  # To store Tableau charge codes

        # Fetch Tableau Data Section (first)
        self.fetch_btn = tb.Button(self, text="Fetch Tableau View", bootstyle=SUCCESS, command=self.fetch_tableau_data)
        self.fetch_btn.pack(pady=10, fill=X, padx=10)

        self.output_text = tb.ScrolledText(self, height=7)
        self.output_text.pack(fill='x', padx=10, pady=5)

        # Upload Excel File Section (second)
        self.upload_btn = tb.Button(self, text="Upload Excel File", bootstyle=PRIMARY, command=self.upload_file)
        self.upload_btn.pack(pady=10, fill=X, padx=10)

        self.excel_preview = tb.ScrolledText(self, height=10)
        self.excel_preview.pack(fill='both', expand=True, padx=10, pady=5)

    def fetch_tableau_data(self):
        self.output_text.delete("1.0", "end")
        self.output_text.insert("end", "Connecting to Tableau...\n")
        self.update()

        tableau_auth = TSC.PersonalAccessTokenAuth(TOKEN_NAME, ACCESS_TOKEN, SITE_ID)
        server = TSC.Server(TABLEAU_SERVER, use_server_version=True)

        try:
            with server.auth.sign_in(tableau_auth):
                target_view = server.views.get_by_id(VIEW_ID)
                if not target_view:
                    self.output_text.insert("end", "Target view not found.\n")
                    return
                self.output_text.insert("end", "Found the target view!\n")

                server.views.populate_csv(target_view)
                csv_bytes = b"".join(target_view.csv)
                df_tableau = pd.read_csv(BytesIO(csv_bytes))

                required_cols = ['Last Name', 'FirstName', 'Charge Code']
                missing_cols = [col for col in required_cols if col not in df_tableau.columns]
                if missing_cols:
                    self.output_text.insert("end", f"Tableau data missing columns: {missing_cols}\n")
                    return
                
                self.charge_code_lookup = df_tableau.set_index(['Last Name', 'FirstName'])['Charge Code'].to_dict()
                self.output_text.insert("end", f"Tableau data fetched successfully with {len(self.charge_code_lookup)} records.\n")

        except Exception as e:
            self.output_text.insert("end", f"Tableau fetch failed: {e}\n")
            self.output_text.insert("end", traceback.format_exc())

    def upload_file(self):
        if self.charge_code_lookup is None:
            messagebox.showwarning("Data Missing", "Please fetch Tableau data before uploading Excel file.")
            return

        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return

        try:
            self.df_excel = pd.read_excel(file_path)

            # Convert 'Date of Service' to datetime
            if 'Date of Service' in self.df_excel.columns:
                self.df_excel['Date of Service'] = pd.to_datetime(self.df_excel['Date of Service'], errors='coerce')

            # Split 'Patient Name' into Last Name and First Name
            if 'Patient Name' in self.df_excel.columns:
                names_split = self.df_excel['Patient Name'].astype(str).str.split(',', n=1, expand=True)
                last_name_col = names_split[0].str.strip()
                first_name_col = names_split[1].str.strip() if names_split.shape[1] > 1 else ''

                patient_name_index = self.df_excel.columns.get_loc('Patient Name')
                self.df_excel.insert(patient_name_index + 1, 'Last Name', last_name_col)
                self.df_excel.insert(patient_name_index + 2, 'First Name', first_name_col)

            # Create ID1, ID2, ID3 columns
            if 'Date of Service' in self.df_excel.columns and 'Patient DOB' in self.df_excel.columns and 'Last Name' in self.df_excel.columns and 'Patient MRN' in self.df_excel.columns:
                excel_serial_DOS = (self.df_excel["Date of Service"] - pd.Timestamp("1899-12-30")).dt.days
                excel_serial_DOB = (self.df_excel["Patient DOB"] - pd.Timestamp("1899-12-30")).dt.days
                self.df_excel.insert(0, 'ID1', self.df_excel["Patient MRN"].astype(str) + excel_serial_DOS.astype(str))
                self.df_excel.insert(self.df_excel.columns.get_loc('ID1') + 1, 'ID2', excel_serial_DOS.astype(str) + excel_serial_DOB.astype(str) + self.df_excel["Last Name"])
                self.df_excel.insert(self.df_excel.columns.get_loc('ID2') + 1, 'ID3', '')  # Placeholder for ID3
            else:
                self.output_text.insert("end", "Missing columns required for ID generation.\n")

            # Add empty columns for reconciliation status
            self.df_excel["Census Reconciliation"] = ""
            self.df_excel["UNBILLED"] = ""

            # Ensure 'E&M (Pro)' column exists
            if 'E&M (Pro)' not in self.df_excel.columns:
                self.df_excel['E&M (Pro)'] = ""
            census_rec = ""
            # Update 'E&M (Pro)' based on Tableau charge codes
            def get_charge_code_and_census(row):
                last = str(row.get('Last Name', '')).strip().upper()
                first = str(row.get('First Name', '')).strip().upper()
                key = (last, first)
                code = self.charge_code_lookup.get(key, '')
                if code == "LWBS":
                    census_rec = "LWBS"
                elif code == "AMA":
                    census_rec = "AMA"
                elif code == "0":
                    census_rec = "NO CHARGE"
                elif code == "NULL":
                    census_rec = ""
                elif code.startswith("99"):
                    census_rec = "BILLED"
                else:
                    census_rec = "TEMPORARY VALUE"
                return pd.Series([code, census_rec])

            self.df_excel[['E&M (Pro)', 'Census Reconciliation']] = self.df_excel.apply(get_charge_code_and_census, axis=1)


            # Save processed file
            from pathlib import Path
            new_file_path = Path(file_path).with_name(f"PROCESSED__________{Path(file_path).stem}.xlsx")
            self.df_excel.to_excel(new_file_path, index=False)

            # Update preview
            self.excel_preview.delete("1.0", "end")
            self.excel_preview.insert("end", self.df_excel.head(20).to_string())

            if messagebox.askyesno("Open File", f"Processed file saved:\n{new_file_path}\n\nDo you want to open it?"):
                os.startfile(new_file_path)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load/process Excel file:\n{e}")
            self.output_text.insert("end", traceback.format_exc())


if __name__ == "__main__":
    app = TableauApp()
    app.mainloop()
