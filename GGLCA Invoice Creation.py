import pandas as pd
import os
import random
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from datetime import datetime

def create_invoice_pdf(row, filename):
    c = canvas.Canvas(filename, pagesize=A4)
    width, height = A4

    # Header
    c.setFont("Helvetica-Bold", 22)
    c.drawString(width/2 - 40, height - 60, "Invoice")

    # Organization details
    c.setFont("Helvetica", 12)
    c.drawString(50, height - 100, "Prachesta Socio Cultural Trust")
    c.drawString(50, height - 120, "Sobha Lakeview Clubhouse")
    c.drawString(50, height - 140, "Bellandur, 560103")
    c.drawString(50, height - 160, "GST Registration : GHFRT3456S")

    # Invoice To section (right aligned)
    c.setFont("Helvetica-Bold", 12)
    c.drawRightString(width - 50, height - 200, "Invoice To")

    # Two line gap before buyer details
    c.setFont("Helvetica", 12)
    c.drawRightString(width - 50, height - 240, row["Buyer Name"])
    c.drawRightString(width - 50, height - 260, row["Buyer Address"])
    c.drawRightString(width - 50, height - 280, row["Buyer GSTIN"])

    # Generate random invoice number: YYYYMMDD + random 4 digits
    try:
        date_obj = datetime.strptime(str(row["Date"]), "%m/%d/%Y")
    except:
        try:
            date_obj = datetime.strptime(str(row["Date"]), "%Y-%m-%d")
        except:
            date_obj = datetime.today()
    invoice_number = date_obj.strftime("%Y%m%d") + str(random.randint(1000, 9999))

    # 3 line gap before invoice number
    c.setFont("Helvetica", 12)
    c.drawString(50, height - 320, f"Invoice Number: {invoice_number}")

    # 3 line gap before Purpose + Amount
    c.drawString(50, height - 360, str(row["Purpose"]))
    c.drawRightString(width - 50, height - 360, str(row["Amount"]))

    # Two line gap before Date field
    formatted_date = date_obj.strftime("%d-%m-%Y")
    c.drawString(50, height - 400, f"Date: {formatted_date}")

    # Footer
    c.setFont("Helvetica-Oblique", 10)
    c.drawString(50, 80, "Authorized Signatory")

    c.save()

def generate_pdfs_from_excel(excel_file):
    df = pd.read_excel(excel_file)

    base_dir = os.path.dirname(excel_file)
    base_name = os.path.splitext(os.path.basename(excel_file))[0]

    for idx, row in df.iterrows():
        buyer_name = str(row["Buyer Name"]).replace(" ", "_")
        filename = os.path.join(base_dir, f"{base_name}_{buyer_name}_{idx+1}.pdf")
        create_invoice_pdf(row, filename)
        print(f"Created: {filename}")

if __name__ == "__main__":
    excel_file = input("Please enter the Excel file path: ").strip()
    generate_pdfs_from_excel(excel_file)