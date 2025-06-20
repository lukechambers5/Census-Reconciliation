import pandas as pd
from pathlib import Path
import os
import traceback

def process_excel_file(file_path, license_key, encounter_lookup=None, df_tableau=None, output_callback=None, tableau_fetcher=None):
    try:
        if output_callback:
            output_callback("Processing Excel file... May take some time for larger files\n")
        df_excel = pd.read_excel(file_path)
        
        # Convert 'Date of Service' to datetime
        if 'Date of Service' in df_excel.columns:
            df_excel['Date of Service'] = pd.to_datetime(df_excel['Date of Service'], errors='coerce')

        # Split 'Patient Name' into Last Name and First Name
        if 'Patient Name' in df_excel.columns:
            names_split = df_excel['Patient Name'].astype(str).str.split(',', n=1, expand=True)
            last_name_col = names_split[0].str.strip()
            first_name_col = names_split[1].str.strip() if names_split.shape[1] > 1 else ''
            patient_name_index = df_excel.columns.get_loc('Patient Name')
            df_excel.insert(patient_name_index + 1, 'Last Name', last_name_col)
            df_excel.insert(patient_name_index + 2, 'First Name', first_name_col)

        if(license_key == "137797"):
            if 'Last Name' in df_excel.columns and 'First Name' in df_excel.columns:
                if 'Patient Account #' in df_excel.columns:
                    acct_index = df_excel.columns.get_loc('Patient Account #')
                    df_excel.insert(acct_index + 1, 'Patient MRN', "")
                    df_excel.insert(acct_index + 2, 'Patient DOB', "")
                else:
                    # Fallback
                    df_excel['Patient MRN'] = ""
                    df_excel['Patient DOB'] = ""
                for idx, row in df_excel.iterrows():
                    last = str(row["Last Name"]).strip().upper()
                    first = str(row["First Name"]).strip().upper()
                    key = (last, first)

                    patient_info = tableau_fetcher.patient_info_lookup.get(key) 
                    if patient_info:
                        raw_dob = patient_info.get("dob", "")
                        try:
                            parsed_dob = pd.to_datetime(raw_dob, errors="coerce")
                            if pd.notnull(parsed_dob):
                                formatted_dob = parsed_dob.strftime("%m/%d/%Y")
                            else:
                                formatted_dob = ""
                        except Exception:
                            formatted_dob = ""

                        df_excel.at[idx, "Patient DOB"] = formatted_dob
                        df_excel.at[idx, "Patient MRN"] = patient_info.get("mrn", "")
        if(license_key == "160214" or license_key == "137797"):
            # Create ID1, ID2, ID3 columns
            if all(col in df_excel.columns for col in ['Date of Service', 'Patient DOB', 'Last Name', 'Patient MRN']):
                df_excel["Date of Service"] = pd.to_datetime(df_excel["Date of Service"], errors="coerce").dt.normalize()
                excel_serial_DOS = (df_excel["Date of Service"] - pd.Timestamp("1899-12-30")).dt.days
                df_excel["Patient DOB"] = pd.to_datetime(df_excel["Patient DOB"], errors="coerce").dt.normalize()
                excel_serial_DOB = (df_excel["Patient DOB"] - pd.Timestamp("1899-12-30")).dt.days
                df_excel.insert(0, 'ID1', df_excel["Patient MRN"].astype(str) + excel_serial_DOS.astype(str))
                df_excel.insert(df_excel.columns.get_loc('ID1') + 1, 'ID2', excel_serial_DOS.astype(str) + excel_serial_DOB.astype(str) + df_excel["Last Name"])
                df_excel.insert(df_excel.columns.get_loc('ID2') + 1, 'ID3', '')  # Placeholder for ID3
            else:
                if output_callback:
                    output_callback("Missing columns required for ID generation.\n")

            # Add empty columns for reconciliation status
            df_excel["Census Reconciliation"] = ""
            df_excel["UNBILLED"] = ""

            # Ensure 'E&M (Pro)' column exists
            if 'E&M (Pro)' not in df_excel.columns:
                df_excel['E&M (Pro)'] = ""
            if(license_key == "160214"):
                # Update 'E&M (Pro)' based on encounter_lookup
                def get_encounter(row):
                    last = str(row.get('Last Name', '')).strip().upper()
                    first = str(row.get('First Name', '')).strip().upper()
                    key = (last, first)

                    encounters = encounter_lookup.get(key, {}) if encounter_lookup else {}
                    use_code = ""
                    census_rec = ""
                    status = "OPEN"

                    if encounters:
                        for appt_id, code_dos_set in encounters.items():
                            if use_code:
                                break
                            for code, dos in code_dos_set:
                                tableau_dos = dos
                                dos_excel = row['Date of Service']
                                if pd.notnull(dos_excel):
                                    excel_dos = f"{dos_excel.month}/{dos_excel.day}/{dos_excel.year}"
                                    if tableau_dos == excel_dos:
                                        if code.startswith("99") or code in ["LWBS", "AMA", "0", "NULL"]:
                                            status = "OPEN"
                                            use_code = code
                                            break
                                        else:
                                            status = 'INVALID CHARGE CODE'
                                    else:
                                        status = "MISMATCH DOS"
                    else:
                        status = "NAME NOT FOUND"

                    if use_code == "LWBS":
                        census_rec = "LWBS"
                    elif use_code == "AMA":
                        census_rec = "AMA"
                    elif use_code == "0":
                        census_rec = "NON ED ENCOUNTERS"
                    elif use_code == "NULL":
                        census_rec = ""
                    elif use_code.startswith("99"):
                        census_rec = "BILLED"
                    else:
                        census_rec = "#N/A"

                    return pd.Series([use_code, census_rec, status])

                df_excel[['E&M (Pro)', 'Census Reconciliation', 'Status']] = df_excel.apply(get_encounter, axis=1)

                condition = (
                    df_excel['Status'].str.upper() == 'OPEN'
                ) & (
                    df_excel['Census Reconciliation'].isin(['BILLED', 'LWBS', 'AMA'])
                )
                df_excel.loc[condition, 'Status'] = 'DE_COMPLETE'

            if(license_key == "137797"):
                df_excel.loc[df_excel['Status'] == "DE_COMPLETE", 'Census Reconciliation'] = "BILLED"

        # Save processed file
        new_file_path = Path(file_path).with_name(f"PROCESSED______{Path(file_path).stem}.xlsx")
        with pd.ExcelWriter(new_file_path, engine='xlsxwriter', datetime_format='mm/dd/yyyy') as writer:
            df_excel.to_excel(writer, index=False)

        if output_callback:
            output_callback(f"Processed file saved: {new_file_path}\n")

        return new_file_path

    except Exception as e:
        if output_callback:
            output_callback(f"Error processing Excel file: {e}\n{traceback.format_exc()}")
        return None
