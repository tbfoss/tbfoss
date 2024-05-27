# automation MM loans
import pandas as pd
import aspose.words as aw
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import smtplib
import tkinter as tk
from tkinter import messagebox

# Constants
EXCEL_FILE = r"C:\Users\tryggvi.bergstad\OneDrive - Vátryggingafélag Íslands hf\orders.xlsx"
SHEET_NAME = 'MM'
WORD_TEMPLATE = r"C:\Users\tryggvi.bergstad\OneDrive - Vátryggingafélag Íslands hf\template.docx"
EMAIL_ADDRESS = 'tryggvi.bergstad@fossar.is'
EMAIL_PASSWORD = 'password'  #  <================================================================ þarf að plugga inn
SMTP_SERVER = 'srelay.hysing.is'
SMTP_PORT = 587
EMAIL_BODY = 'Hæ,\n\nGetið þið afgreitt meðfylgjandi peningamarkaðslán?\n\nMbk,'

def find_order_details(order_number):
    # Read the Excel file from the specified sheet
    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
    
    # Ensure the order number is searched in column A
    order_details = df[df[df.columns[0]] == int(order_number)]
    
    if order_details.empty:
        print(f"Order number {order_number} not found.")
        return None
    else:
        return order_details.iloc[0]

def create_pdf_filename(order_details):
    # Extract required details for the filename
    viðskiptamaður = order_details['C']
    viðskiptadagur = pd.to_datetime(order_details['F']).strftime('%d-%m-%Y')
    tímalengd = order_details['B']

    # Construct the filename
    filename = f"Staðfesting%20á%20peningamarkaðsláni_{viðskiptamaður}_{viðskiptadagur}_{tímalengd}.pdf"
    return filename

def create_word_doc(order_details, pdf_filename):
    # Load the template Word document
    doc = aw.Document(WORD_TEMPLATE)
    
    # Define the placeholders and their corresponding data fields
    placeholders = {
        'viðskiptamaður': 'C',
        'kennitala': 'D',
        'heimilisfang': 'E',
        'dagsetning': 'F',  
        'tilvísun': 'A',    
        'viðskiptadagur': 'F',
        'fyrsti_vaxtadagur': 'G',
        'gjalddagi': 'H',
        'vextir': 'I',
        'dagaregla': 'J',  
        'höfuðstóll': 'K',
        'áfallnir_vextir_á_höfuðstóli': 'L',
        'greiðsla_á_gjalddaga': 'F',
        'greiðslufyrirkomulag': 'O'
    }
    
    builder = aw.DocumentBuilder(doc)
    builder.move_to_document_start()
    
    for placeholder, column in placeholders.items():
        value = order_details[column] if column in order_details else ""
        builder.replace(f'{{{{ {placeholder} }}}}', str(value))
    
    # Save the document as a PDF
    doc.save(pdf_filename)

def send_email_with_attachment(to_email, subject, body, attachment):
    # Create the email message
    msg = MIMEMultipart()
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = "itstryggvi@gmail.com"
    msg['Subject'] = subject

    # Attach the email body
    msg.attach(MIMEText(body, 'plain'))

    # Attach the PDF file
    with open(attachment, 'rb') as attachment_file:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment_file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename= {attachment}')
        msg.attach(part)

    # Send the email
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        server.send_message(msg)


def process_order(order_number):
    # Find the corresponding order details in the Excel sheet
    order_details = find_order_details(order_number)
    
    if order_details is not None:
        # Create the PDF filename
        pdf_filename = create_pdf_filename(order_details)
        
        # Create a Word doc with the order details and save it as a PDF
        create_word_doc(order_details, pdf_filename)
        
        # Create the email subject
        email_subject = f"Peningamarkaðslán #{order_number}"
        
        # Send the PDF as an email attachment
        send_email_with_attachment('itstryggvi@gmail.com', email_subject, EMAIL_BODY, pdf_filename)
        messagebox.showinfo("Success", "Email sent successfully.")
    else:
        messagebox.showerror("Error", f"Order number {order_number} not found.")

def on_submit():
    order_number = order_number_entry.get()
    process_order(order_number)

# Create the main window
root = tk.Tk()
root.title("Order Processing")

# Create and place the widgets
tk.Label(root, text="Enter the order number:").grid(row=0, column=0, padx=10, pady=10)
order_number_entry = tk.Entry(root)
order_number_entry.grid(row=0, column=1, padx=10, pady=10)
submit_button = tk.Button(root, text="Submit", command=on_submit)
submit_button.grid(row=1, columnspan=2, pady=10)

# Start the GUI event loop
root.mainloop()
