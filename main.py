from fpdf import FPDF
import pandas as pd
import glob
import openpyxl

# Input DATA
filepaths = glob.glob('invoices/*.xlsx')

for file in filepaths:
    df = pd.read_excel(file, sheet_name='Sheet 1')

    for i, row in df.iterrows():
      print(f'File{i}\n'
            f'{row}\n')






# Output
