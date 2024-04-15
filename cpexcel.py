import tkinter as tk
from tkinter import filedialog
import pandas as pd
from openpyxl import Workbook
import os

def upload_file():
    file_path = filedialog.askopenfilename()
    if file_path.endswith('.xlsx'):
        df = pd.read_excel(file_path)
        return df
    else:
        print("Please select a .xlsx file.")

def copy_to_new_excel(df):
    if df is not None:
        base_filename = "copied_data.xlsx"
        i = 1
        new_file_path = base_filename
        while os.path.exists(new_file_path):
            new_file_path = f"copied_data_{i}.xlsx"
            i += 1
        with pd.ExcelWriter(new_file_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        print("Data successfully copied to:", new_file_path)

# Call the function to upload the Excel file
df = upload_file()

# Call the function to copy the data to a new Excel file
copy_to_new_excel(df)