from collections import defaultdict
import tableauserverclient as TSC
import pandas as pd
from pandas.errors import EmptyDataError
from io import BytesIO
from datetime import datetime, timedelta

class TableauFetcher:
    #
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
        else:
            print(text, end="")

    def _update_progress(self, value):
        if self.progress_callback:
            self.progress_callback(value)

    def fetch_data(self, license_key, filter_values):
        try:
            tableau_auth = TSC.TableauAuth(self.username, self.password, '')
            server = TSC.Server(
                "https://tableau.blitzmedical.com",
                use_server_version=True
            )
            server.add_http_options({'timeout': 3600})

            with server.auth.sign_in(tableau_auth):
                self._update_progress(0.05)
                target = "Concord Census Reconciliation View"
                if(license_key != ""):
                    target = "EHP Census Reconciliation Details"
                all_views = list(TSC.Pager(server.views))
                oldest_dos = filter_values
                yesterday = (datetime.today() - timedelta(days=1)).strftime("%#m/%#d/%Y")
                
                matched_views = [view for view in all_views if view.name == target]
                target_view = matched_views[0]
                opts = TSC.CSVRequestOptions()
                opts.max_rows = -1
                opts.include_all_columns = True

                def generate_dates(from_date, to_date):
                    date_format = "%Y-%m-%d"
                    from_date = datetime.strptime(from_date, date_format)
                    to_date = datetime.strptime(to_date, date_format)

                    dates = []
                    current_date = from_date
                    while current_date <= to_date:
                        dates.append(current_date.strftime(date_format))
                        current_date += timedelta(days=1)

                    return ",".join(dates)

                def normalize_date(d):
                    return datetime.strptime(d, "%m/%d/%Y").strftime("%Y-%m-%d")
                
                start = normalize_date(oldest_dos)
                end = normalize_date(yesterday)
                date_range = generate_dates(start, end)
                opts.vf("DOS", date_range)
                if license_key in ('160214', '137797'):
                    opts.vf("Charge Code", "")
                    opts.vf("Last Name", "")
                    opts.vf("License Key", license_key)
                    opts.vf("CPT Code")
                server.views.populate_csv(target_view, req_options=opts)
                slice_bytes = b"".join (target_view.csv)
                if not slice_bytes.strip():
                    self._safe_insert("No rows returned from Tableau.\n")
                    self._update_progress(1)
                    return None

                try:
                    df = pd.read_csv(BytesIO(slice_bytes), on_bad_lines='warn', engine="python")
                except EmptyDataError:
                    self._safe_insert("Returned CSV had no columns.\n")
                    self._update_progress(1)
                    return None
                
                # If needed, build encounter lookups for license-key mode
                if license_key in ('160214', '137797'):
                    total_rows = len(df)
                    for i, (_, row) in enumerate(df.iterrows()):
                        last = str(row['Last Name']).strip().upper()
                        first = str(row['FirstName']).strip().upper()
                        code = str(row['Charge Code']).strip().upper()
                        dos = str(row['DOS']).strip()
                        appointment_num = str(row['Appointment FID']).strip()
                        dob = str(row['DOB']).strip()
                        mrn = str(row['Chart Number']).strip()
                        name_key = (last, first)
                        provider = str(row.get('Provider', '')).strip()

                        if name_key not in self.patient_info_lookup:
                            self.patient_info_lookup[name_key] = {"dob": dob, "mrn": mrn}

                        if (code, dos) not in [(c, d) for c, d, _ in self.encounter_lookup[(last, first)][appointment_num]]:
                            self.encounter_lookup[(last, first)][appointment_num].append((code, dos, provider))

                        if i % max(1, total_rows // 20) == 0:
                            self._update_progress(40 + int(((i / total_rows) * 60)) / 100 )
                self._safe_insert(f"Retrieved {len(df)} rows from Tableau. ")
                self.df_tableau = df
                self._update_progress(1)
                
                return df

        except Exception as e:
            self._safe_insert(f"Error fetching Tableau data: {e}\n")
            self._update_progress(1)
            return None
          
