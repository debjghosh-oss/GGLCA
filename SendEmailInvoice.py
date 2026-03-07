import os
import smtplib
import pandas as pd
from email.message import EmailMessage

def send_invoice(to_email, subject, body, attachment_path, smtp_server, smtp_port, sender_email, sender_password):
    msg = EmailMessage()
    msg["From"] = sender_email
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.set_content(body)

    # Attach PDF
    with open(attachment_path, "rb") as f:
        file_data = f.read()
        file_name = os.path.basename(attachment_path)
    msg.add_attachment(file_data, maintype="application", subtype="pdf", filename=file_name)

    # Send email
    with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
        server.login(sender_email, sender_password)
        server.send_message(msg)
        print(f"Sent invoice to {to_email}")

def mail_invoices(excel_file, output_dir, smtp_server, smtp_port, sender_email, sender_password):
    df = pd.read_excel(excel_file)

    for idx, row in df.iterrows():
        buyer_name = str(row["Buyer Name"]).replace(" ", "_")
        pdf_file = os.path.join(output_dir, f"Invoice_{buyer_name}_{idx+1}.pdf")

        if os.path.exists(pdf_file):
            subject = f"GGLCA Invoice - {row['Buyer Name']}"
            body = f"Dear {row['Buyer Name']},\n\nPlease find attached your invoice.\n\nRegards,\nGGLCA"
            send_invoice(row["Email"], subject, body, pdf_file,
                         smtp_server, smtp_port, sender_email, sender_password)
        else:
            print(f"Invoice PDF not found for {row['Buyer Name']}")

if __name__ == "__main__":
    excel_file = r"C:\Users\debjy\PycharmProjects\HelloWorld\data\invoices.xlsx"
    output_dir = r"C:\Users\debjy\PycharmProjects\HelloWorld\output"

    # Configure your SMTP settings
    smtp_server = "smtp.gmail.com"   # or your corporate SMTP
    smtp_port = 465                  # SSL port
    sender_email = "deb.j.ghosh@gmail.com"
    sender_password = "eqfz jkxh bheq syjw"

    mail_invoices(excel_file, output_dir, smtp_server, smtp_port, sender_email, sender_password)