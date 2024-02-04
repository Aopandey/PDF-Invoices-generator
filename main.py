import pandas as zx
import glob
from fpdf import FPDF
from pathlib import Path

filepath = glob.glob("invoices/*.xlsx")

for i in filepath:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_font(family="Times", size=16, style="B")
    pdf.add_page()

    filename = Path(i).stem
    invoice_num = filename.split("-")[0]
    date = filename.split("-")[1]

    pdf.cell(w=50, h=8, txt=f"Invoice no. {invoice_num}", ln=1)
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    data = zx.read_excel(i, sheet_name="Sheet 1")
    columns_excel = data.columns
    columns_excel = [item.replace("_", " ").title() for item in columns_excel]
    pdf.set_font(family='Times', size=10, style="B")
    pdf.set_text_color(85, 85, 85)
    pdf.cell(w=25, h=8, txt=columns_excel[0], border=1)
    pdf.cell(w=60, h=8, txt=columns_excel[1], border=1)
    pdf.cell(w=35, h=8, txt=columns_excel[2], border=1)
    pdf.cell(w=25, h=8, txt=columns_excel[3], border=1)
    pdf.cell(w=25, h=8, txt=columns_excel[4], border=1, ln=1)

    for index, row in data.iterrows():
        pdf.set_font(family='Times', size=10)
        pdf.set_text_color(85, 85, 85)
        pdf.cell(w=25, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=60, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=35, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=25, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=25, h=8, txt=str(row["total_price"]), border=1, ln=1)

    sum_total = data["total_price"].sum()
    pdf.set_font(family='Times', size=10)
    pdf.set_text_color(85, 85, 85)
    pdf.cell(w=25, h=8, txt="", border=1)
    pdf.cell(w=60, h=8, txt="", border=1)
    pdf.cell(w=35, h=8, txt="", border=1)
    pdf.cell(w=25, h=8, txt="", border=1)
    pdf.cell(w=25, h=8, txt=str(sum_total), border=1, ln=1)

    pdf.set_font(family='Times', size=14, style="B")
    pdf.cell(w=25, h=8, txt=f"The total price is: {sum_total}", ln=3)

    pdf.set_font(family='Times', size=14, style="B")
    pdf.cell(w=28, h=8, txt=f"PythonHow")
    pdf.image("pythonhow.png", w=10)

    pdf.output(f"PDF/{filename}.pdf")


