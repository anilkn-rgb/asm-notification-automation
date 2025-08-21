import imaplib
import email
import os
import re
import json
import urllib.request
import ssl
from datetime import datetime, timedelta

def download_asm_directly():
    """Download the ASM PDF from daily Extranet Files email"""
    
    # Get credentials from environment variables (GitHub Secrets)
    GMAIL_USER = os.environ.get('GMAIL_USER')
    GMAIL_APP_PASSWORD = os.environ.get('GMAIL_APP_PASSWORD')
    WEBHOOK_URL = os.environ.get('WEBHOOK_URL')
    
    # Fallback to hardcoded values if running locally
    if not GMAIL_USER:
        GMAIL_USER = "anil.kn@etssecurities.com"
        GMAIL_APP_PASSWORD = "lymu sjca qnzf vose"
        WEBHOOK_URL = "https://chat.googleapis.com/v1/spaces/AAAAhPquChA/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=WxHW1xhb0p9ou1a8FVNoHP0UaNVxj7RhthcuqHPhwZw"
    
    # Set download folder (GitHub Actions vs local)
    if os.environ.get('GITHUB_ACTIONS'):
        DOWNLOAD_FOLDER = "asm_downloads"
        os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)
    else:
        DOWNLOAD_FOLDER = r"D:\new python folder\ASM alerts"
    
    print("DOWNLOADING ASM PDF FROM DAILY EXTRANET FILES")
    print("=" * 50)
    
    try:
        # Connect to Gmail
        print("Connecting to Gmail...")
        mail = imaplib.IMAP4_SSL('imap.gmail.com')
        mail.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        mail.select('inbox')
        print("Connected to Gmail")
        
        # Get yesterday's date (since emails are about previous day's data)
        yesterday = datetime.now() - timedelta(days=1)
        date_format_1 = yesterday.strftime("%d.%m.%Y")  # 20.08.2025
        date_format_2 = yesterday.strftime("%Y-%m-%d")  # 2025-08-20
        
        print(f"Searching for Extranet Files email for {date_format_1}...")
        
        # Search for yesterday's Extranet Files email
        search_queries = [
            f'SUBJECT "Extranet Files - {date_format_1}"',
            f'FROM "krishnamurthy.ks@etssecurities.com" SUBJECT "Extranet Files"',
            f'FROM "adarsh.mishra@etssecurities.com" SUBJECT "Extranet Files"',
            'SUBJECT "Extranet Files"'
        ]
        
        target_uid = None
        found_subject = None
        
        for query in search_queries:
            print(f"Trying search: {query}")
            result, email_ids = mail.search(None, query)
            
            if result == 'OK' and email_ids[0]:
                email_list = email_ids[0].split()
                
                # Check the most recent emails first
                for uid in reversed(email_list[-5:]):
                    try:
                        result, msg_data = mail.fetch(uid, '(RFC822)')
                        if result == 'OK':
                            email_message = email.message_from_bytes(msg_data[0][1])
                            subject = email_message.get('Subject', '')
                            sender = email_message.get('From', '')
                            date_received = email_message.get('Date', '')
                            
                            print(f"Checking email {uid.decode()}: {subject}")
                            print(f"From: {sender}")
                            
                            # Check if this is a recent Extranet Files email
                            if 'Extranet Files' in subject:
                                # Check for ASM PDF attachment
                                has_asm_pdf = False
                                for part in email_message.walk():
                                    if part.get_content_disposition() == 'attachment':
                                        filename = part.get_filename() or ''
                                        if 'ASM' in filename and filename.endswith('.pdf'):
                                            has_asm_pdf = True
                                            print(f"Found ASM PDF attachment: {filename}")
                                            break
                                
                                if has_asm_pdf:
                                    target_uid = uid.decode()
                                    found_subject = subject
                                    print(f"Found target email: {target_uid}")
                                    break
                    except Exception as e:
                        print(f"Error checking email {uid}: {e}")
                        continue
                
                if target_uid:
                    break
        
        if not target_uid:
            print("No recent Extranet Files email with ASM PDF found")
            # Send notification about missing email
            try:
                message = f"No Extranet Files email found for {date_format_1}\nPlease check if the email was sent."
                payload = {"text": message}
                data = json.dumps(payload).encode('utf-8')
                
                req = urllib.request.Request(
                    WEBHOOK_URL,
                    data=data,
                    headers={'Content-Type': 'application/json'}
                )
                
                context = ssl.create_default_context()
                with urllib.request.urlopen(req, context=context, timeout=30) as response:
                    if response.getcode() == 200:
                        print("'No email found' notification sent successfully!")
            except Exception as e:
                print(f"Error sending notification: {e}")
            return
        
        print(f"Processing email: {found_subject}")
        result, msg_data = mail.fetch(target_uid, '(RFC822)')
        
        if result == 'OK':
            email_message = email.message_from_bytes(msg_data[0][1])
            
            # Download ASM PDF attachment
            pdf_downloaded = False
            for part in email_message.walk():
                if part.get_content_disposition() == 'attachment':
                    filename = part.get_filename()
                    
                    if filename and 'ASM' in filename and filename.lower().endswith('.pdf'):
                        print(f"Found ASM PDF: {filename}")
                        
                        # Get attachment data
                        attachment_data = part.get_payload(decode=True)
                        
                        # Save to folder
                        pdf_path = os.path.join(DOWNLOAD_FOLDER, f"ASM_{date_format_2}.pdf")
                        
                        with open(pdf_path, 'wb') as f:
                            f.write(attachment_data)
                        
                        print(f"Downloaded: {pdf_path}")
                        pdf_downloaded = True
                        
                        # Extract data from PDF
                        print("\nEXTRACTING DATA FROM PDF")
                        print("=" * 30)
                        
                        # Read PDF content
                        with open(pdf_path, 'rb') as file:
                            content = file.read()
                        
                        # Extract text
                        text_content = content.decode('latin-1', errors='ignore')
                        print(f"Extracted {len(text_content)} characters of text")
                        
                        # DEBUG - Show sample content
                        print("\nDEBUG - Sample text content:")
                        print("=" * 40)
                        print(text_content[:800])
                        print("=" * 40)
                        print(f"Z00018 found in text: {'Z00018' in text_content}")
                        print(f"Z00008 found in text: {'Z00008' in text_content}")
                        print(f"ETS SECURITIES found in text: {'ETS SECURITIES' in text_content}")
                        print("=" * 40)
                        
                        # Extract client data for the two specific client codes
                        client_data = []
                        target_codes = ['Z00018', 'Z00008']
                        
                        for client_code in target_codes:
                            if client_code in text_content:
                                print(f"Processing {client_code}...")
                                
                                # Find all lines containing this client code
                                lines = text_content.split('\n')
                                for line in lines:
                                    if client_code in line and 'ETS SECURITIES' in line:
                                        print(f"Found line for {client_code}: {line}")
                                        
                                        # Extract all decimal numbers from the line
                                        numbers = re.findall(r'\d+\.\d{2}', line)
                                        print(f"{client_code} numbers found: {numbers}")
                                        
                                        if len(numbers) >= 3:
                                            # Table structure: [Potential Losses] [Applicable Margins] [Additional Surveillance Margin]
                                            # We want the Additional Surveillance Margin (3rd number)
                                            margin = numbers[2]
                                            
                                            client_data.append({
                                                'client_code': client_code,
                                                'additional_margin': margin
                                            })
                                            print(f"Extracted: {client_code} -> {margin} Cr")
                                        elif len(numbers) >= 1:
                                            # Fallback: use last number if fewer than 3 columns
                                            margin = numbers[-1]
                                            client_data.append({
                                                'client_code': client_code,
                                                'additional_margin': margin
                                            })
                                            print(f"Fallback extraction: {client_code} -> {margin} Cr")
                                        break
                        
                        # Extract reference number
                        ref_match = re.search(r'NCL/RMG/\d{4}/\d{5}', text_content)
                        ref_number = ref_match.group(0) if ref_match else None
                        
                        # Extract date from PDF content
                        date_matches = re.findall(r'(August \d{1,2}, \d{4})', text_content)
                        pdf_date = date_matches[0] if date_matches else date_format_1
                        
                        print(f"\nEXTRACTED DATA:")
                        print(f"Reference: {ref_number}")
                        print(f"Date: {pdf_date}")
                        print(f"Client data count: {len(client_data)}")
                        for client in client_data:
                            print(f"{client['client_code']}: {client['additional_margin']} Cr")
                        
                        # Create message
                        if client_data:
                            message = f"ASM as on {pdf_date}\n\n"
                            message += "Client Code / Additional Surveillance Margin:\n"
                            for client in client_data:
                                message += f"{client['client_code']}: Rs{client['additional_margin']} Cr\n"
                            if ref_number:
                                message += f"\nReference: {ref_number}"
                        else:
                            message = f"ASM is NIL for {pdf_date}"
                        
                        print(f"\nMESSAGE TO SEND:")
                        print("=" * 30)
                        print(message)
                        print("=" * 30)
                        
                        # Send to Google Chat
                        print("\nSending to Google Chat...")
                        try:
                            payload = {"text": message}
                            data = json.dumps(payload).encode('utf-8')
                            
                            req = urllib.request.Request(
                                WEBHOOK_URL,
                                data=data,
                                headers={'Content-Type': 'application/json'}
                            )
                            
                            context = ssl.create_default_context()
                            with urllib.request.urlopen(req, context=context, timeout=30) as response:
                                if response.getcode() == 200:
                                    print("Message sent successfully!")
                                else:
                                    print(f"Failed to send. Status: {response.getcode()}")
                        except Exception as e:
                            print(f"Error sending message: {e}")
                        
                        break
            
            if not pdf_downloaded:
                print("No ASM PDF found in the Extranet Files email")
        else:
            print("Failed to fetch email")
        
        # Close connection
        mail.close()
        mail.logout()
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    download_asm_directly()
