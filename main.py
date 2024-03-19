import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path
import time

# Getting the filepaths from our folder invoices (all that contain .xlsx
filepaths = glob.glob("invoices/*.xlsx")


for index, filepath in enumerate(filepaths):
    filename = Path(filepath).stem # Returns just the filename
    invoice_number = filename.split("-")[0] # Splitting the number away
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    pdf.set_font(family="Times", size=16, style="B")

    # Printing out the invoice number of each filename
    pdf.cell(w=50, h=8, txt=f"Invoice Nr.{invoice_number}")
    pdf.ln(0)
    curr_date = time.strftime("%Y.%m.%d")
    # Printing out the current date of the invoice
    pdf.cell(w=50, h=30, txt=f"Date: {curr_date}", ln=1)

    # Reading the excel file to get the values
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    # Setting up the header
    pdf.set_font(family="Times", size=12, style="B")
    pdf.cell(w=30, h=8, txt="Product Id", border=1)
    pdf.cell(w=35, h=8, txt="Product Name", border=1)
    pdf.cell(w=40, h=8, txt="Amount Purchased", border=1)
    pdf.cell(w=35, h=8, txt="Price Per Unit", border=1)
    pdf.cell(w=40, h=8, txt="Total Price", border=1)
    pdf.ln(0)

    # Printing out the values of each header/row
    for id, row in df.iterrows():
        pdf.ln(8)
        pdf.set_font(family="Times", size=12)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), ln=0, border=1)
        pdf.cell(w=60, h=8, txt=str(row["product_name"]), ln=0, border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), ln=0, border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), ln=0, border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), ln=0, border=1)
    pdf.output(f"PDFs/Invoice-{index+1}.pdf")
