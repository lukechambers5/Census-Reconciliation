import pandas as pd
import os

def get_oldest_dos(file_path):
    ext = os.path.splitext(file_path)[1].lower()

    if ext == ".csv":
        df = pd.read_csv(file_path, usecols=["Date of Service"])
    else:
        xl = pd.ExcelFile(file_path)
        for sheet in xl.sheet_names:
            try:
                df = xl.parse(sheet, usecols=["Date of Service"])
                break
            except ValueError:
                continue
        else:
            raise ValueError("No sheet contains 'Date of Service' column")
    print(df.columns.to_list())
    df["Date of Service"] = pd.to_datetime(df["Date of Service"], errors="coerce")
    oldest_dos = df["Date of Service"].dropna().min()

    if pd.notnull(oldest_dos) and int(oldest_dos.strftime("%Y")) < 2023:
        return "01/01/2024"
    
    return oldest_dos.strftime("%#m/%#d/%Y") if pd.notnull(oldest_dos) else None
