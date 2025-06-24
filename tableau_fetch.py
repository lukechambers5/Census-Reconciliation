from collections import defaultdict
import tableauserverclient as TSC
import pandas as pd
from pandas.errors import EmptyDataError
from io import BytesIO

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

    def fetch_data(self, license_key, filter_values):
        try:
            tableau_auth = TSC.TableauAuth(self.username, self.password, '')
            server = TSC.Server(
                "https://tableau.blitzmedical.com",
                use_server_version=True
            )

            # Determine view based on license_key
            view_id = (
                "28773dd1-0c1e-445b-96f9-264fcdd0f485"
                if license_key in ('160214', '137797')
                else "19921e6e-ff1c-44fa-810d-b129b2806078"
            )
            with server.auth.sign_in(tableau_auth):
                self._update_progress(5)
                target_view = server.views.get_by_id(view_id)
                if not target_view:
                    self._safe_insert("Target view not found.\n")
                    return None
                self._update_progress(10)
                self._safe_insert(f"Filtering for {len(filter_values)} patient names and license/group...\n")

                # Batch patient names into chunks to avoid URL limits
                chunk_size = 150
                batches = [
                    filter_values[i : i + chunk_size]
                    for i in range(0, len(filter_values), chunk_size)
                ]
                dfs = []
                total_batches = len(batches)
       
                for idx, batch in enumerate(batches):
                    opts = TSC.CSVRequestOptions()
                    opts.max_rows = -1
                    opts.include_all_columns = True
                    # apply patient name filters
                    for name in batch:
                        opts.vf("Patient Name", name)
                    # apply license key or group filter
                    if license_key in ('160214', '137797'):
                        opts.vf("Charge Code", "")
                        opts.vf("Last Name", "")
                        opts.vf("License Key", license_key)
                    else:
                        opts.vf("License Key (group)", "Blitz Concord")

                    server.views.populate_csv(target_view, req_options=opts)
                    slice_bytes = b"".join(target_view.csv)
                    # slice_bytes is showing up as b'\r\n' - search up what this means and maybe filter issue idk?
                    print(slice_bytes)
                    if not slice_bytes.strip():
                        self._safe_insert(f"Batch {idx+1}/{total_batches} returned no rows—skipping.\n")
                        # still bump progress so the bar moves
                        prog = 15 + int((idx + 1) / total_batches * 20)
                        self._update_progress(prog)
                        continue
                    try:
                        df_slice = pd.read_csv(
                            BytesIO(slice_bytes),
                            on_bad_lines='warn',
                            engine="python"
                        )
                    except EmptyDataError:
                        self._safe_insert(f"Batch {idx+1} had no columns—skipping.\n")
                        prog = 15 + int((idx + 1) / total_batches * 20)
                        self._update_progress(prog)
                        continue
                    dfs.append(df_slice)

                    # update progress
                    prog = 15 + int((idx + 1) / total_batches * 20)
                    self._update_progress(prog)

                # Combine all slices
                df = pd.concat(dfs, ignore_index=True)
                self._safe_insert("Downloaded and concatenated all filtered slices.\n")

                # If needed, build encounter lookups for license-key mode
                if license_key in ('160214', '137797'):
                    self._safe_insert("Building encounter lookup...\n")
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

                        if (code, dos) not in [c_d for c_d, _p in self.encounter_lookup[(last, first)][appointment_num]]:
                            self.encounter_lookup[(last, first)][appointment_num].append((code, dos, provider))

                        if i % max(1, total_rows // 20) == 0:
                            self._update_progress(40 + int((i / total_rows) * 60))

                self._safe_insert(f"Retrieved {len(df)} rows from Tableau.\n")
                self.df_tableau = df
                self._update_progress(100)
                return df

        except Exception as e:
            self._safe_insert(f"Error fetching Tableau data: {e}\n")
            self._update_progress(100)
            return None
