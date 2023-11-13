from fpdf import FPDF
import pandas as pd
import glob
import openpyxl
from pathlib import Path


filepaths = glob.glob('invoices/*.xlsx')

for file in filepaths:
    df = pd.read_excel(file, sheet_name='Sheet 1')
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=False, margin=0)
    pdf.add_page()

    filename = Path(file).stem
    invoice_nr = filename.split('-')[0]
    invoice_date = filename.split('-')[1]
    pdf.set_font(family='Times', style='B', size=12)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=0, h=12, txt=f"Invoice nr. {invoice_nr}", align='L', ln=1)
    pdf.cell(w=0, h=5, txt=f"Date {invoice_date}", align='L', ln=1)

    pdf.output(f'PDF/{filename}.pdf')







