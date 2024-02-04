import pandas as zx
import glob
from fpdf import FPDF
from pathlib import Path

filepath = glob.glob("invoices/*.xlsx")

for i in filepath:
    data = zx.read_excel(i, sheet_name="Sheet 1")

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_font(family="Times", size=16, style="B")
    pdf.add_page()

    filename = Path(i).stem
    invoice_num = filename.split("-")[0]
    date = filename.split("-")[1]

    pdf.cell(w=50, h=8, txt=f"Invoice no. {invoice_num}", ln=1)
    pdf.cell(w=50, h=8, txt=f"Date: {date}")
    pdf.output(f"PDF/{filename}.pdf")


