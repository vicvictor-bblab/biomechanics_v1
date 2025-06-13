import os
import pandas as pd


def load_excel_or_csv(filepath):
    """Load a CSV or Excel file into a DataFrame"""
    if filepath.lower().endswith('.csv'):
        df = pd.read_csv(filepath)
    else:
        xls = pd.ExcelFile(filepath)
        sheet_name = xls.sheet_names[0]
        df = xls.parse(sheet_name)
    return df


def load_multiple_files(filepaths):
    """Return a dictionary mapping file base names to DataFrames."""
    df_dict = {}
    for fp in filepaths:
        base = os.path.basename(fp)
        try:
            df = load_excel_or_csv(fp)
            df_dict[base] = df
        except Exception as e:
            print(f"Failed to load {fp}: {e}")
    return df_dict
