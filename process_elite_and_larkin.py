import pandas as pd
import numpy as np
from pathlib import Path
import traceback

def process_excel_file(file_path, license_key, encounter_lookup=None, df_tableau=None, tableau_fetcher=None, output_callback=None):
    try:
        if output_callback:
            output_callback("Processing Excel file... May take some time for larger files\n")

        xl = pd.ExcelFile(file_path)
        for sheet in xl.sheet_names:
            try:
                df = xl.parse(sheet)
                if "Date of Service" in df.columns:
                    if output_callback:
                        output_callback(f"[INFO] Using sheet: {sheet}\n")
                    break
            except Exception:
                continue
        else:
            raise ValueError("No sheet contains 'Date of Service' column")

        # Convert DOS column to datetime after finding correct sheet
        df["Date of Service"] = pd.to_datetime(df["Date of Service"], errors="coerce")

        # GET FIRST KEY
        if 'Patient Name' in df.columns:
            names = df['Patient Name'].astype(str).str.split(',', n=1, expand=True)
            df['Last Name']  = names[0].str.strip().str.upper()
            df['First Name'] = names[1].str.strip().str.upper().fillna("")
            # key is first token before any space
            df['FirstKey']  = df['First Name'].str.split().str[0]

        if 'PatientName' in df.columns:
            names = df['PatientName'].astype(str).str.split(',', n=1, expand=True)
            df['Last Name']  = names[0].str.strip().str.upper()
            df['First Name'] = names[1].str.strip().str.upper().fillna("")
            # key is first token before any space
            df['FirstKey']  = df['First Name'].str.split().str[0]

        # COLUMNS TO ADD
        cols_to_init = [
            'Provider','Patient MRN','Patient DOB',
            'ID1','ID2','ID3','Census Reconciliation','UNBILLED','E&M (Pro)','Status'
        ]
        for col in cols_to_init:
            if col not in df.columns:
                df[col] = ""

        df['DosNormalize'] = df['Date of Service'].dt.normalize()

        # CREATE ENCOUNTER LOOKUP
        if encounter_lookup and license_key in ('160214', '137797'):
            enc_rows = []
            for (last, first), appts in encounter_lookup.items():
                key_first = first.upper().split()[0]
                for appt_id, entries in appts.items():
                    for code, dos_str, provider in entries:
                        enc_rows.append({
                            'Last Name': last,
                            'FirstKey': key_first,
                            'DosLookup': pd.to_datetime(
                                dos_str, format='%m/%d/%Y', errors='coerce'
                            ).normalize(),
                            'Code': code,
                            'ProviderLookup': provider
                        })
            enc_df = pd.DataFrame(enc_rows)

            enc_df['is_99'] = enc_df['Code'].str.startswith('99', na=False)

            enc_df = enc_df.sort_values(
                ['Last Name','FirstKey','DosLookup','is_99'],
                ascending=[True, True, True, False]
            )

            enc_df = enc_df.drop_duplicates(
                subset=['Last Name','FirstKey','DosLookup'],
                keep='first'
            ).drop(columns='is_99')

            df = df.merge(
                enc_df,
                left_on=['Last Name','FirstKey','DosNormalize'],
                right_on=['Last Name','FirstKey','DosLookup'],
                how='left'
            )

            df['Provider'] = df['ProviderLookup'].fillna("")

            # LICENSE-SPECIFIC LOGIC
            if license_key == '160214':
                cond_billed = df['Code'].str.startswith('99', na=False)
                cond_lwbs   = df['Code']=='LWBS'
                cond_ama    = df['Code']=='AMA'
                cond_zero   = df['Code']=='0'
                cond_null   = df['Code']=='NULL'

                df['use_code'] = np.select(
                    [cond_lwbs, cond_ama, cond_zero, cond_null, cond_billed],
                    ['LWBS','AMA','0','NULL', df['Code']],
                    default=''
                )
                df['Census Reconciliation'] = np.select(
                    [cond_lwbs, cond_ama, cond_zero, cond_null, cond_billed],
                    ['LWBS','AMA','NON ED ENCOUNTERS','', 'BILLED'],
                    default='#N/A'
                )
                
                df['Status'] = np.where(df['use_code']=='', 'MISMATCH DOS', 'OPEN')

                df.loc[
                    cond_billed & df['Census Reconciliation'].isin(['BILLED','LWBS','AMA']),
                    'Status'
                ] = 'DE_COMPLETE'

                # INVALID CODES
                cond_invalid = df['Code'].notna() & ~(
                    cond_lwbs | cond_ama | cond_zero | cond_null | cond_billed
                )
                df.loc[cond_invalid, 'Status'] = 'INVALID CODE IN TABLEAU'

                df['E&M (Pro)'] = df['use_code']

                tableau_keys = set(encounter_lookup.keys())
                key_series = pd.Series(list(zip(df['Last Name'], df['FirstKey'])))
                mask_name_exists = key_series.isin(tableau_keys)
                df.loc[~mask_name_exists, 'Status'] = 'NAME NOT FOUND IN TABLEAU'
                
            elif license_key == '137797':
                tableau_keys = set(encounter_lookup.keys())
                patient_keys    = list(zip(df['Last Name'], df['FirstKey']))
                mask_exists     = [k in tableau_keys for k in patient_keys]
                mask_dos_matched = df['Code'].notna()
                statuses        = df['Status'].fillna('').astype(str)

                # ABANDONED = blank, else BILELD or MISMATCHED depending on if DOS matches, and lastly NAME NOT IN TABLEAU if none
                df['Census Reconciliation'] = [
                    "" if stat == 'ABANDONED' else
                    'BILLED'         if ex and matched else
                    'MISMATCHED DOS' if ex and not matched else
                    'NAME NOT IN TABLEAU'
                    for stat, ex, matched in zip(statuses, mask_exists, mask_dos_matched)
                ]

        if tableau_fetcher and getattr(tableau_fetcher, 'patient_info_lookup', None) and license_key == '137797':
            mrn_map = {}
            dob_map = {}
            for (l, f), info in tableau_fetcher.patient_info_lookup.items():
                key = (l.upper(), f.upper())
                mrn_map[key] = info.get('mrn', '') 
                raw_dob = info.get('dob', '')
                dob_ts  = pd.to_datetime(raw_dob, errors='coerce')
                if pd.notnull(dob_ts):
                    dob_map[key] = dob_ts.strftime('%m/%d/%Y')
                else:
                    dob_map[key] = ""
            keys = list(zip(df['Last Name'], df['FirstKey']))
            df['Patient MRN'] = [mrn_map.get(k, "") for k in keys]
            df['Patient DOB'] = [dob_map.get(k, "") for k in keys]

        # GENERATE IDs
        if license_key in ('160214','137797'):
            df['DosNorm'] = df['Date of Service'].dt.normalize()
            df['DobNorm'] = pd.to_datetime(
                df['Patient DOB'], format='%m/%d/%Y', errors='coerce'
            ).dt.normalize()
            serial_DOS = (df['DosNorm'] - pd.Timestamp('1899-12-30')).dt.days.astype(str)
            serial_DOB = (df['DobNorm'] - pd.Timestamp('1899-12-30')).dt.days.astype(str)

            df['ID1'] = df['Patient MRN'].astype(str) + serial_DOS
            df['ID2'] = serial_DOS + serial_DOB + df['Last Name']
            df['ID3'] = ''

        # DROP TEMP COLUMNS
        df.drop(columns=[
            'DosNormalize','DosLookup','ProviderLookup','Code',
            'use_code','DosNorm','DobNorm','MRN','DobLookup','FirstKey'
        ], errors='ignore', inplace=True)

        # REORDER COLUMNS
        desired = [
            'ID1','ID2','ID3',
            'Date of Service','Date Billed','Facility','Patient Account #',
            'Patient MRN','Patient DOB','Patient Name','Last Name','First Name',
            'E&M (Fac)','E&M (Pro)','Status','Census Reconciliation','UNBILLED','Provider'
        ]
        cols = [c for c in desired if c in df.columns] + [c for c in df.columns if c not in desired]
        df = df[cols]

        out = Path(file_path).with_name(f"PROCESSED______{Path(file_path).stem}.xlsx")
        df.to_excel(out, index=False)
        if output_callback:
            output_callback(f"Processed file saved: {out}\n")
        return out

    except Exception:
        if output_callback:
            output_callback(traceback.format_exc())
        return None
