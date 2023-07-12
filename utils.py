
import PyPDF2
import os
import shutil
import smtplib, ssl
from email.message import EmailMessage
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText

class PDf_email_generator:
    def __init__(self, File_listener_folder, Output_File_Location):
        self.File_listener_folder = File_listener_folder
        self.Output_File_Location = Output_File_Location


    def pdf_reader(self):
        pdfFileObj = open(f"{self.file_path}", 'rb')
        # Create a PDF reader object
        pdfReader = PyPDF2.PdfReader(pdfFileObj)
        # Get the total number of pages
        numPages = len(pdfReader.pages) 
        # Loop through each page and extract the text
        for pageNum in range(numPages):
            pageObj = pdfReader.pages[pageNum]
            text = pageObj.extract_text()
            print(text)

        for line in text.splitlines():
            if "Bill To" in line:
                self.Bill_to = line.split("Bill To: ")[1]
                self.Bill_to = self.Bill_to.split("Invoice No: ")[0]
                print(self.Bill_to)
            if "Invoice No" in line:
                self.invoice_no = line.split("Invoice No: ")[1]
                print(self.invoice_no)

            if "Due Date" in line:
                self.Due_date = line.split("Due Date: ")[1]
                print(self.Due_date)

            if "Total " in line:
                self.Total_amount = line.split("Total ")[1]
                print(self.Total_amount)
        # Close the PDF file object
        pdfFileObj.close()
        return self


        
    def rename_pdf(self):
    # get the source file name
        source_filename = self.file_path
        new_pdf_name = f"{self.invoice_no}.pdf"
        
        # get the destination folder
        destination_folder = self.Output_File_Location
        
        # concatenate the file name to the destination folder
        destination_file_path = os.path.join(destination_folder, new_pdf_name)
        
        try:
            # rename the file
            os.rename(source_filename, destination_file_path)
            print("PDF file has successfully renamed as:", new_pdf_name)
            
            # move file to the destination folder
            shutil.move(destination_file_path, destination_folder)
            print("PDF file has moved to:", destination_folder)
        except Exception as e:
            print("error message:", str(e))
        return self

    def email_sender(self):
        msg = MIMEMultipart()
        port = 465  # For SSL
        smtp_server = "smtp.gmail.com"
        sender_email = "@gmail.com" # Enter your address
        receiver_email = ""  # Enter receiver address
        password = ""
        
        body = f"""
        Hi,

        I am writing to inform you that the invoice {self.invoice_no} for the services we rendered to {self.Bill_to} Store is now ready and attached in this email. The invoice amount is {self.Total_amount} in total.
        We would appreciate prompt payment of this invoice within {self.Due_date}. If you have any questions or concerns regarding the invoice, Please feel free to reach out to us.

        Thank you for your business.

        Best regards, 
        Josh R"""
        
        msg.attach(MIMEText(body, 'plain'))
        msg['Subject'] = f"Invoice {self.invoice_no} for - from SMTP server"
        msg['From'] = sender_email
        msg['To'] = receiver_email
        
        ### attach the pdf file

        with open(f"{self.Output_File_Location}/{self.invoice_no}.pdf", "rb") as f:
            part = MIMEApplication(f.read(),Name = f"{self.invoice_no}.pdf")
        part['Content-Disposition'] = f'attachment; filename = {self.invoice_no}.pdf'
        msg.attach(part)

        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
            server.login(sender_email, password)
            server.send_message(msg, from_addr=sender_email, to_addrs=receiver_email)



    def file_runner(self):
        if len(os.listdir(self.File_listener_folder)) == 0:
            print("Folder is empty.")
        else:
            # Read the file if it exists
            print(os.listdir(self.File_listener_folder))
            file_name = os.listdir(self.File_listener_folder)[0]
            self.file_path = os.path.join(self.File_listener_folder, file_name)
            print(file_name)
            print(self.file_path)
            if os.path.isfile(self.file_path):
                self.pdf_reader().rename_pdf().email_sender()
            else:
                assert ("File does not exist in the folder.")
            return self





