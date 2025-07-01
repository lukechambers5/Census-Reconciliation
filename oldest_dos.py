import pandas as pd
import os

def get_oldest_dos(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".csv":
        df = pd.read_csv(file_path, usecols=["Date of Service"])
    else:
        df = pd.read_excel(file_path, usecols=["Date of Service"])

    if "Date of Service" not in df.columns:
        raise ValueError("Missing 'Date of Service' column")

    df["Date of Service"] = pd.to_datetime(df["Date of Service"], errors="coerce")
    oldest_dos = df["Date of Service"].dropna().min()
    print("OLDEST DOS:", oldest_dos.strftime("%#m/%#d/%Y"))
    return oldest_dos.strftime("%#m/%#d/%Y") if pd.notnull(oldest_dos) else None
