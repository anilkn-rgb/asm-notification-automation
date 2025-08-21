import imaplib
import email
import os
import re
import json
import urllib.request
import ssl
from datetime import datetime

def download_asm_directly():
    """Download the ASM PDF directly from the known email"""
    
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
    
    print("🚀 DOWNLOADING ASM PDF DIRECTLY")
    print("=" * 50)
    
    try:
        # Connect to Gmail
        print("📧 Connecting to Gmail...")
        mail = imaplib.IMAP4_SSL('imap.gmail.com')
        mail.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        mail.select('inbox')
        print("✅ Connected to Gmail")
        
        # Find the most recent ASM email (instead of hardcoded UID)
        print("🔍 Searching for ASM emails...")
        
        # Search for emails with ASM in subject, from last 30 days
        result, email_ids = mail.search(None, 'SUBJECT "ASM"')
        
        target_uid = None
        if result == 'OK' and email_ids[0]:
            # Get the most recent email
            email_list = email_ids[0].split()
            
            # Check last 10 emails to find one with PDF attachment
            for uid in reversed(email_list[-10:]):
                try:
                    result, msg_data = mail.fetch(uid, '(RFC822)')
                    if result == 'OK':
                        email_message = email.message_from_bytes(msg_data[0][1])
                        subject = email_message.get('Subject', '')
                        
                        # Skip reply emails
                        if 'Re:' in subject or 'RE:' in subject:
                            continue
                            
                        # Check if email has PDF attachment
                        has_pdf = False
                        for part in email_message.walk():
                            if part.get_content_disposition() == 'attachment':
                                filename = part.get_filename()
                                if filename and 'ASM' in filename and filename.lower().endswith('.pdf'):
                                    has_pdf = True
                                    break
                        
                        if has_pdf:
                            target_uid = uid.decode()
                            print(f"✅ Found recent ASM email with PDF: UID {target_uid}")
                            print(f"📧 Subject: {subject}")
                            break
                except:
                    continue
        
        # Fallback to UID 908 if no recent email found
        if not target_uid:
            target_uid = '908'
            print(f"⚠️ Using fallback UID {target_uid}")
        
        print(f"📧 Fetching email UID {target_uid}...")
        result, msg_data = mail.fetch(target_uid, '(RFC822)')
        
        if result == 'OK':
            email_message = email.message_from_bytes(msg_data[0][1])
            subject = email_message.get('Subject', 'No Subject')
            date_received = email_message.get('Date', 'No Date')
            print(f"✅ Email subject: {subject}")
            print(f"📅 Email date: {date_received}")
            
            # Extract date from email content or use current date
            current_date = datetime.now().strftime("%Y-%m-%d")
            
            # Download PDF attachment
            pdf_downloaded = False
            for part in email_message.walk():
                if part.get_content_disposition() == 'attachment':
                    filename = part.get_filename()
                    
                    if filename and 'ASM' in filename and filename.lower().endswith('.pdf'):
                        print(f"🔍 Found ASM PDF: {filename}")
                        
                        # Get attachment data
                        attachment_data = part.get_payload(decode=True)
                        
                        # Save to folder
                        pdf_path = os.path.join(DOWNLOAD_FOLDER, f"ASM_{current_date}.pdf")
                        
                        with open(pdf_path, 'wb') as f:
                            f.write(attachment_data)
                        
                        print(f"✅ Downloaded: {pdf_path}")
                        pdf_downloaded = True
                        
                        # Now extract data from PDF
                        print("\n📄 EXTRACTING DATA FROM PDF")
                        print("=" * 30)
                        
                        # Read PDF content
                        with open(pdf_path, 'rb') as file:
                            content = file.read()
                        
                        # Extract text
                        text_content = content.decode('latin-1', errors='ignore')
                        print(f"✅ Extracted {len(text_content)} characters of text")
                        
                        # Extract client data manually
                        client_data = []
                        
                        # Look for Z00018
                        if 'Z00018' in text_content:
                            pos = text_content.find('Z00018')
                            snippet = text_content[pos:pos+300]
                            numbers = re.findall(r'\d{1,4}\.\d{2}', snippet)
                            print(f"Z00018 numbers: {numbers}")
                            if len(numbers) >= 3:
                                client_data.append({
                                    'client_code': 'Z00018',
                                    'additional_margin': numbers[2]  # Usually the 3rd number
                                })
                        
                        # Look for Z00008
                        if 'Z00008' in text_content:
                            pos = text_content.find('Z00008')
                            snippet = text_content[pos:pos+300]
                            numbers = re.findall(r'\d{1,4}\.\d{2}', snippet)
                            print(f"Z00008 numbers: {numbers}")
                            if len(numbers) >= 3:
                                client_data.append({
                                    'client_code': 'Z00008',
                                    'additional_margin': numbers[2]  # Usually the 3rd number
                                })
                        
                        # Extract reference
                        ref_match = re.search(r'NCL/RMG/\d{4}/\d{5}', text_content)
                        ref_number = ref_match.group(0) if ref_match else None
                        
                        # Extract date from PDF content
                        date_matches = re.findall(r'(August \d{1,2}, \d{4})', text_content)
                        pdf_date = date_matches[0] if date_matches else current_date
                        
                        print(f"\n📊 EXTRACTED DATA:")
                        print(f"Reference: {ref_number}")
                        print(f"Date: {pdf_date}")
                        for client in client_data:
                            print(f"{client['client_code']}: ₹{client['additional_margin']} Cr")
                        
                        # Create message
                        if client_data:
                            message = f"ASM as on {pdf_date}\n\n"
                            message += "Client Code / Additional Surveillance Margin:\n"
                            for client in client_data:
                                message += f"{client['client_code']}: ₹{client['additional_margin']} Cr\n"
                            if ref_number:
                                message += f"\nReference: {ref_number}"
                        else:
                            message = f"ASM is NIL for {pdf_date}"
                        
                        print(f"\n📤 MESSAGE TO SEND:")
                        print("=" * 30)
                        print(message)
                        print("=" * 30)
                        
                        # Send to Google Chat
                        print("\n📤 Sending to Google Chat...")
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
                                    print("✅ Message sent successfully!")
                                else:
                                    print(f"❌ Failed to send. Status: {response.getcode()}")
                        except Exception as e:
                            print(f"❌ Error sending message: {e}")
                        
                        break
            
            if not pdf_downloaded:
                print("❌ No ASM PDF found in email")
        else:
            print("❌ Failed to fetch email")
        
        # Close connection
        mail.close()
        mail.logout()
        
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    download_asm_directly()
