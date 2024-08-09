import win32com.client
from datetime import datetime, timedelta
import os

def extract_email_body_and_save():
    # Create an instance of the Outlook application
    outlook = win32com.client.Dispatch("Outlook.Application")
    
    # Get the namespace (MAPI)
    namespace = outlook.GetNamespace("MAPI")
    
    # Get the default Inbox folder
    inbox = namespace.GetDefaultFolder(6)  # 6 refers to the Inbox
    
    # Define today's date
    today = datetime.today().date()
    filter_str = f"[ReceivedTime] >= '{today.strftime('%m/%d/%Y 00:00 AM')}' AND [ReceivedTime] < '{(today + timedelta(days=1)).strftime('%m/%d/%Y 00:00 AM')}'"
    
    # Get the messages in the Inbox received today
    messages = inbox.Items.Restrict(filter_str)
    
    # Directory to save the XML files
    save_dir = r'\email_extracter\MQIN'
    #save_dir = r'G:\MQIN_backup'
    
    # Create the directory if it doesn't exist
    if not os.path.exists(save_dir):
        os.makedirs(save_dir)
    
    # Iterate over the messages
    for message in messages:
        # Check if the email is unread and has the desired subject
        if message.UnRead and message.Subject == "FW:Sent from AirCentre Movement Manager":
            # Get the email body
            email_body = message.Body.strip()
            
            # Find the start of the XML declaration
            xml_start = email_body.find('<?xml version="1.0" encoding="UTF-8"?>')
            if xml_start != -1:
                # Extract the XML content starting from the declaration
                email_body = email_body[xml_start:]
                
                # Create a filename using a unique identifier
                # FS_Export_YYYYMMDDHHMMSSFFF_Seq.xml
                now = datetime.now()
                seq_number = 1
                while True:
                    filename = f"FS_Export_{now.strftime('%Y%m%d%H%M%S%f')}.xml"
                    file_path = os.path.join(save_dir, filename)
                    if not os.path.exists(file_path):
                        break
                    seq_number += 1
                
                # Save the email body to an XML file
                with open(file_path, "w", encoding="utf-8") as file:
                    file.write(email_body)
                
                print(f"Email body saved to {file_path}")
                
                # Mark the email as read
                message.UnRead = False
                message.Save()

# Call the function to extract the email body and save it
extract_email_body_and_save()
