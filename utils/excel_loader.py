# excel_chart_ui/utils/excel_loader.py

import pandas as pd

def load_excel_sheets_and_columns(filepath):
    excel_file = pd.ExcelFile(filepath)
    sheets = {}
    for sheet_name in excel_file.sheet_names:
        df = excel_file.parse(sheet_name)
        sheets[sheet_name] = (df, df.columns.astype(str).tolist())
    return sheets
