import pandas as pd
import os

def get_names(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".csv":
        df = pd.read_csv(file_path)
    else:
        df = pd.read_excel(file_path)

    patient_names = (
        df["Patient Name"]
        .dropna()
        .astype(str)
        .str.upper()
        .str.replace(", ", ",", regex=False)
        .unique()
        .tolist()
    )

    return patient_names