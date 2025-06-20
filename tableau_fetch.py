from collections import defaultdict
import tableauserverclient as TSC
import pandas as pd
from io import BytesIO
from config import VIEW_ID, TABLEAU_SERVER, SITE_ID

class TableauFetcher:
    def __init__(self, username, password, output_callback=None, progress_callback=None):
        self.username = username
        self.password = password
        self.encounter_lookup = defaultdict(lambda: defaultdict(list))
        self.patient_info_lookup = {}
        self.df_tableau = None
        self.output_callback = output_callback
        self.progress_callback = progress_callback

    def _safe_insert(self, text):
        if self.output_callback:
            self.output_callback(text)

    def _update_progress(self, value):
        if self.progress_callback:
            self.progress_callback(value)

    def fetch_data(self, license_key):
        try:
            tableau_auth = TSC.TableauAuth(self.username, self.password, SITE_ID)
            server = TSC.Server(TABLEAU_SERVER, use_server_version=True)

            with server.auth.sign_in(tableau_auth):
                self._update_progress(5)
                target_view = server.views.get_by_id(VIEW_ID)
                if not target_view:
                    self._safe_insert("Target view not found.\n")
                    return None

                self._update_progress(10)

                req_option = TSC.CSVRequestOptions()
                req_option.max_rows = -1
                req_option.include_all_columns = True
                req_option.vf("Charge Code", "")
                req_option.vf("Last Name", "")
                req_option.vf("License Key", license_key)

                server.views.populate_csv(target_view, req_options=req_option)
                self._update_progress(15)
                self._safe_insert("Downloading data, this may take some time...\n")

                csv_bytes = b"".join(target_view.csv)
                df = pd.read_csv(BytesIO(csv_bytes), on_bad_lines='warn', engine="python")
                self._update_progress(35)

                self._safe_insert("Processing data...\n")

                total_rows = len(df)
                for i, (_, row) in enumerate(df.iterrows()):
                    last = str(row['Last Name']).strip().upper()
                    first = str(row['FirstName']).strip().upper()
                    code = str(row['Charge Code']).strip().upper()
                    dos = str(row['DOS']).strip()
                    appointment_num = str(row['Appointment FID']).strip()
                    dob = str(row['DOB']).strip()
                    mrn = str(row['Chart Number']).strip();
                    name_key = (last, first)
                    if name_key not in self.patient_info_lookup:
                        self.patient_info_lookup[name_key] = {
                            "dob": dob,
                            "mrn": mrn
                        }
                    if (code, dos) not in self.encounter_lookup[(last, first)][appointment_num]:
                        self.encounter_lookup[(last, first)][appointment_num].append((code, dos))

                    if i % max(1, total_rows // 20) == 0:
                        progress_val = 40 + int((i / total_rows) * 60)
                        self._update_progress(progress_val)

                self._safe_insert(f"Retrieved {len(df)} rows from Tableau.\n")

                self.df_tableau = df
                self._update_progress(100)
                return df

        except Exception as e:
            self._safe_insert(f"Error fetching Tableau data: {e}\n")
            self._update_progress(100)
            return None
