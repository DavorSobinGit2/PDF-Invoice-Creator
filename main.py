import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path
import time

filepaths = glob.glob("invoices/*.xlsx")


for index, filepath in enumerate(filepaths):
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    filename = Path(filepath).stem
    invoice_number = filename.split("-")[0]
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice Nr.{invoice_number}")
    pdf.ln(0)
    curr_date = time.strftime("%Y.%m.%d")
    pdf.cell(w=50, h=30, txt=f"Date: {curr_date}")
    pdf.output(f"PDFs/Invoice-{index+1}.pdf")
