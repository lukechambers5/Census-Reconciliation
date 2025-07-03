import pandas as pd
import os

def process_concord(df_tableau, file_path):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".csv":
        df = pd.read_csv(file_path)
    else:
        df = pd.read_excel(file_path)

    # FILTER OUT NON-BLITZ
    df = df[~((df['Location Code'] == 'CMG_ADVHMA') & (df['Department Code'] == 'ED'))]
    df = df[~((df['Location Code'] == 'CMG_BREMH') & (df['Department Code'] == 'URGENTCARE'))]
    df = df[~((df['Location Code'] == 'CMG_BREMH') & (df['Department Code'] == 'ED'))]
    df = df[~((df['Location Code'] == 'CMG_CHSAL') & (df['Department Code'] == 'TELEPULM'))]
    df = df[~((df['Location Code'] == 'CMG_CHSBV') & (df['Department Code'] == 'TELEPULM'))]
    df = df[~((df['Location Code'] == 'CMG_CHSKV') & (df['Department Code'] == 'TELEPULM'))]
    df = df[~((df['Location Code'] == 'CMG_DEMFD') & (df['Department Code'] == 'ED'))]
    df = df[~((df['Location Code'] == 'CMG_MCGHTN') & (df['Department Code'] == 'ED'))]
    df = df[~((df['Location Code'] == 'CMG_RMCTN') & (df['Department Code'] == 'ED'))]
    df = df[~((df['Location Code'] == 'CMG_SUCCH') & (df['Department Code'] == 'ED'))]
    df = df[~((df['Location Code'] == 'CMG_TAYRH') & (df['Department Code'] == 'ED'))]
    df = df[~((df['Location Code'] == 'CMG_WDLN') & (df['Department Code'] == 'HOSPITALIST'))]

    # NAME DICTIONARY
    df_tableau['FirstName'] = df_tableau['FirstName'].astype(str).str.strip().str.upper().str.split().str[0]
    df_tableau['Last Name'] = df_tableau['Last Name'].astype(str).str.strip()
    df_tableau['DOS'] = pd.to_datetime(df_tableau['DOS'], errors='coerce')
    df_tableau['Chart Number'] = df_tableau['Chart Number'].astype(str).str.strip()
    name_lookup = {}
    for _, row in df_tableau.iterrows():
        key_dos = (
            row['Last Name'],
            row['FirstName'],
            row['DOS'],
        )
        key_mrn = (
            row['Last Name'],
            row['FirstName'],
            row['Chart Number'],
        )
        name_lookup[key_dos] = row.to_dict()
        name_lookup[key_mrn] = row.to_dict()
        
    # ID INSERTION LOGIC and Tableau Fetch LOGIC
    df.insert(0, 'ID (DOS_ACCT)', '')
    df.insert(1, 'ID2 (DOS_MRN)', '') 
    df.insert(2, 'ID3 (DOS_Patient Name)', '')
    df.insert(3, 'Patient Name ', '')
    df.insert(4, 'Facility', '')
    df.insert(5, 'Carrier', '')
    df.insert(6, 'Provider', '')

    df['Patient Name'] = df['Patient Name'].astype(str).str.strip()
    df['Date of Service'] = df['Date of Service'].astype(str).str.strip()

    # Go row by row in df
    for idx, row in df.iterrows():
        try:
            date_obj = pd.to_datetime(row['Date of Service'])
            serial_date = str((date_obj - pd.Timestamp("1899-12-30")).days)

            acct = str(row.get('Account Number', '')).strip()
            acct = ''.join(filter(str.isdigit, acct))

            mrn = str(row.get('Medical Record Number', '')).strip()
            mrn = ''.join(filter(str.isdigit, mrn)) 

            patient_name = str(row.get('Patient Name', '')).strip()

            last_first = patient_name.split(',')
            if len(last_first) != 2:
                continue

            last = last_first[0].strip().upper()
            first = last_first[1].strip().upper().split()[0]

            # IDs
            if acct:
                combined_id = serial_date + acct
                df.at[idx, 'ID (DOS_ACCT)'] = combined_id
            if mrn:
                combined_id_2 = serial_date + mrn
                df.at[idx, 'ID2 (DOS_MRN)'] = combined_id_2
            if patient_name:
                combined_id_3 = serial_date + patient_name
                df.at[idx, 'ID3 (DOS_Patient Name)'] = combined_id_3
            
            # GET TABLEAU DATA
            match = None
            key = (last, first, date_obj)
            key2 = (last, first, mrn)
            if key in name_lookup:
                match = name_lookup[key]
            elif key2 in name_lookup:
                match = name_lookup[key2]
                
            if match:
                df.at[idx, 'Patient Name '] = match.get('Patient Name', '')
                df.at[idx, 'Provider'] = match.get('Provider', '')
                df.at[idx, 'Carrier'] = match.get('Carrier', '')
                df.at[idx, 'Facility'] = match.get('Facility Name', '')
            else:
                df.at[idx, 'Patient Name '] = '#N/A'
                df.at[idx, 'Provider'] = '#N/A'
                df.at[idx, 'Carrier'] = '#N/A'
                df.at[idx, 'Facility'] = '#N/A'

        except Exception as e:
            print(f"Row {idx} error: {e}")

    new_file_path = os.path.join(os.path.dirname(file_path), "PROCESSED_____" + os.path.basename(file_path))

    if ext == ".csv":
        df.to_csv(new_file_path, index=False)
    else:
        df.to_excel(new_file_path, index=False)
    
    return new_file_path 