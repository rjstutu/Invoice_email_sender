## external import
import PyPDF2
import os
import shutil
import smtplib, ssl
from email.message import EmailMessage
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email import policy
from email.parser import BytesParser
from email.utils import parseaddr
import imaplib

## internal import
from config import *



class PDf_email_generator:
    def __init__(self, File_listener_folder, Output_File_Location):
        self.File_listener_folder = File_listener_folder
        self.Output_File_Location = Output_File_Location
        ## default self values
        self.Send_to = None
        self.File_No = None
        self.Due_Date = None
        self.Total_amount = None


    def pdf_reader(self):
        pdfFileObj = open(f"{self.file_path}", 'rb')
        # Create a PDF reader object
        pdfReader = PyPDF2.PdfReader(pdfFileObj)
        # Get the total number of pages
        numPages = len(pdfReader.pages) 
        # Loop through each page and extract the text
        all_text = ""
        for pageNum in range(numPages):
            pageObj = pdfReader.pages[pageNum]
            all_text += pageObj.extract_text()
            print(all_text)

        for key in KEY_INFO_DICT_NEW.keys():
            if key in all_text:
                self.file_purpose = key
                print(f"self.file_purpose: {self.file_purpose}")
                break
        
        # Extract the information

        for key, info in KEY_INFO_DICT_NEW[self.file_purpose][EXTRACT_KEY_DICT].items():
            if info in all_text:
                setattr(self, key, all_text.split(info)[1].split('\n')[0].strip())

                if key == SEND_TO:  ## usually followed by "Invoice No" or "Estimate No"
                    setattr(self, key, getattr(self, key).split(KEY_INFO_DICT_NEW[self.file_purpose][EXTRACT_KEY_DICT][FILE_NO])[0])
                print(f"self.{key}:  {getattr(self, key)}")
                print("------------------")
                continue
        pdfFileObj.close()
        return self


        
    def rename_pdf(self):
    # get the source file name
        source_filename = self.file_path
        new_pdf_name = f"{self.File_No}-{service_party}.pdf"
        print(f"self.File_No: {self.File_No}")
        print("source_filename:", source_filename)
        print("new_pdf_name:", new_pdf_name)
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
        smtp_server = "smtp.gmail.com"

        email_obj = email_body(self.File_No, self.Send_to, self.Total_amount, self.Due_Date)

        if self.file_purpose == INVOICE:
            body = email_body(self.File_No, self.Send_to, self.Total_amount, self.Due_Date).invoice_email_body()
        elif self.file_purpose == ESTIMATE:
            body = email_body(self.File_No, self.Send_to, self.Total_amount, self.Due_Date).estimate_email_body()
        else:
            body = "Error: File purpose is not identified."

        purpose_method_mapping = {
            INVOICE: email_obj.invoice_email_body(),
            ESTIMATE: email_obj.estimate_email_body()
        }

        try:
            body = purpose_method_mapping.get(self.file_purpose)
        except:
            body = "Error: File purpose is not identified."

        msg.attach(MIMEText(body, 'plain'))
        msg['Subject'] = f"{self.file_purpose} {self.File_No} for {service_party} - from SMTP server"
        msg['From'] = sender_email
        msg['To'] = receiver_email
        
        ### attach the pdf file

        with open(f"{self.Output_File_Location}/{self.File_No}-{service_party}.pdf", "rb") as f:
            part = MIMEApplication(f.read(),Name = f"{self.File_No}-{service_party}.pdf")
        part['Content-Disposition'] = f'attachment; filename = {self.File_No}-{service_party}.pdf'
        msg.attach(part)

        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
            server.login(sender_email, password)
            server.send_message(msg, from_addr=sender_email, to_addrs=receiver_email)

    def email_listener(self, user_email, user_password, download_folder='.'):
        
        # Connect to the Gmail IMAP server
        mail = imaplib.IMAP4_SSL('imap.gmail.com')
        mail.login(user_email, user_password)
        mail.select('inbox')

        # Search for all unread emails
        result, data = mail.search(None, '(UNSEEN SUBJECT "execute")')
        if result == 'OK':
            for num in data[0].split():
                # Fetch the email by ID
                result, email_data = mail.fetch(num, '(RFC822)')
                if result == 'OK':
                    raw_email = email_data[0][1]
                    # Parse the email content
                    msg = BytesParser(policy=policy.default).parsebytes(raw_email)
                    
                    # Extract sender's email address
                    from_email = parseaddr(msg["from"])[1]

                    # Check if the sender is in the whitelist
                    if from_email not in WHITELIST:
                        print(f"Ignored email from non-whitelisted sender: {from_email}")
                        continue

                    print("Subject:", msg["subject"])
                    print("From:", msg["from"])
                    
                    # Checking for attachments
                    for part in msg.walk():
                        content_disposition = str(part.get("Content-Disposition"))
                        if "attachment" in content_disposition:
                            filename = part.get_filename()
                            if filename:
                                filepath = os.path.join(download_folder, filename)
                                with open(filepath, 'wb') as f:
                                    f.write(part.get_payload(decode=True))
                                print(f"Downloaded attachment to {filepath}")
                    
                    # Mark the email as read
                    mail.store(num, '+FLAGS', '\\Seen')
        mail.logout()
        return self

    def file_runner(self):
        self.email_listener(user_email= sender_email,
                                    user_password = password, 
                                    download_folder= self.File_listener_folder)
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





