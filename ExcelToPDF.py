import pandas as pd
import os
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

def create_invoice_pdf(row, filename):
    c = canvas.Canvas(filename, pagesize=A4)
    width, height = A4

    # Title
    c.setFont("Helvetica-Bold", 16)
    c.drawString(200, height - 50, "Tax Invoice")

    # Buyer details
    c.setFont("Helvetica", 12)
    c.drawString(50, height - 100, f"Buyer Name: {row['Buyer Name']}")
    c.drawString(50, height - 120, f"Buyer Address: {row['Buyer Address']}")
    c.drawString(50, height - 140, f"Buyer GSTIN: {row['Buyer GSTIN']}")
    c.drawString(50, height - 160, f"Place of Supply: {row['Place of Supply']}")

    # Optional fields
    if 'Billing Address' in row and pd.notna(row['Billing Address']):
        c.drawString(50, height - 180, f"Billing Address: {row['Billing Address']}")
    if 'Shipping Address' in row and pd.notna(row['Shipping Address']):
        c.drawString(50, height - 200, f"Shipping Address: {row['Shipping Address']}")
    if 'Contact Details' in row and pd.notna(row['Contact Details']):
        c.drawString(50, height - 220, f"Contact Details: {row['Contact Details']}")
    if 'PAN' in row and pd.notna(row['PAN']):
        c.drawString(50, height - 240, f"PAN: {row['PAN']}")

    # Footer
    c.setFont("Helvetica-Oblique", 10)
    c.drawString(50, 50, "Authorized Signatory")

    c.save()

def generate_pdfs_from_excel(excel_file):
    df = pd.read_excel(excel_file)

    # Get directory and base name
    base_dir = os.path.dirname(excel_file)
    base_name = os.path.splitext(os.path.basename(excel_file))[0]

    for idx, row in df.iterrows():
        buyer_name = str(row['Buyer Name']).replace(" ", "_")
        filename = os.path.join(base_dir, f"{base_name}_{buyer_name}_{idx+1}.pdf")
        create_invoice_pdf(row, filename)
        print(f"Created: {filename}")

if __name__ == "__main__":
    excel_file = input("Please enter the Excel file path: ").strip()
    generate_pdfs_from_excel(excel_file)