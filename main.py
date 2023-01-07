import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_nr = filename.split("-")[0]
    invoice_date = filename.split("-")[1]
    header = str(df.columns).title().replace("_", " ")

    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=0, h=10, txt=f"Invoices nr.{invoice_nr}", ln=1)
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=0, h=10, txt=f"Date: {invoice_date}", ln=1)

    columns = df.columns
    header = [item.title().replace("_", " ") for item in columns]
    pdf.set_font(family="Times", size=12)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=30, h=10, txt=header[0], border=1)
    pdf.cell(w=60, h=10, txt=header[1], border=1)
    pdf.cell(w=40, h=10, txt=header[2], border=1)
    pdf.cell(w=30, h=10, txt=header[3], border=1)
    pdf.cell(w=30, h=10, txt=header[4], border=1, ln=1)

    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=12)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(w=30, h=10, txt=str(row["product_id"]), border=1)
        pdf.cell(w=60, h=10, txt=str(row["product_name"]), border=1)
        pdf.cell(w=40, h=10, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=10, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=10, txt=str(row["total_price"]), border=1, ln=1)

    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=12)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=30, h=10, txt="", border=1)
    pdf.cell(w=60, h=10, txt="", border=1)
    pdf.cell(w=40, h=10, txt="", border=1)
    pdf.cell(w=30, h=10, txt="", border=1)
    pdf.cell(w=30, h=10, txt=str(total_sum), border=1, ln=1)

    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=0, h=10, txt="", ln=1)
    pdf.cell(w=0, h=10, txt=f"The total price is {total_sum}", ln=1)
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=32, h=10, txt=f"PythonHow")
    pdf.image("pythonhow.png", w=10)

    pdf.output(f"PDFs/{filename}.pdf")