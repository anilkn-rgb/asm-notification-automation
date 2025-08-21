import imaplib
import email
import os
import re
import json
import urllib.request
import ssl
from datetime import datetime, timedelta

def download_asm_directly():
    """Download the ASM PDF directly from the most recent email"""
    
    # Get credentials from environment variables (GitHub Secrets)
    GMAIL_USER = os.environ.get('GMAIL_USER')
    GMAIL_APP_PASSWORD = os.environ.get('GMAIL_APP_PASSWORD')
    WEBHOOK_URL = os.environ.get('WEBHOOK_URL')
    
    # Validate that all secrets are available
    if not all([GMAIL_USER, GMAIL_APP_PASSWORD, WEBHOOK_URL]):
        print("❌ Missing required environment variables!")
        return
    
    # Create download folder in current directory
    DOWNLOAD_FOLDER = "asm_downloads"
    os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)
    
    print("🚀 DOWNLOADING ASM PDF DIRECTLY")
    print("=" * 50)
    
    try:
        # Connect to Gmail
        print("📧 Connecting to Gmail...")
        mail = imaplib.IMAP4_SSL('imap.gmail.com')
        mail.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        mail.select('inbox')
        print("✅ Connected to Gmail")
        
        # Search for recent ASM emails - try multiple approaches
        print("🔍 Searching for recent ASM emails...")
        
        # Try different search strategies
        search_strategies = [
            'UNSEEN SUBJECT "ASM"',  # Unread ASM emails first
            'SUBJECT "Additional Margin"',  # Full subject
            'SUBJECT "ASM"',  # Any ASM emails
            'FROM "nse" SUBJECT "ASM"',  # From NSE with ASM
            'ALL'  # Last resort - get all emails and filter
        ]
        
        latest_email_id = None
        
        for strategy in search_strategies:
            print(f"🔍 Trying: {strategy}")
            result, email_ids = mail.search(None, strategy)
            
            if result == 'OK' and email_ids[0]:
                email_list = email_ids[0].split()
                if email_list:
                    # Check each email to find the most recent ASM one
                    for email_id in reversed(email_list[-10:]):  # Check last 10 emails
                        try:
                            result, msg_data = mail.fetch(email_id, '(RFC822)')
                            if result == 'OK':
                                email_message = email.message_from_bytes(msg_data[0][1])
                                subject = email_message.get('Subject', '')
                                date_str = email_message.get('Date', '')
                                
                                print(f"📧 Checking email {email_id.decode()}: {subject}")
                                
                                if 'ASM' in subject or 'Additional' in subject:
                                    latest_email_id = email_id
                                    print(f"✅ Found ASM email: {email_id.decode()}")
                                    break
                        except:
                            continue
                    
                    if latest_email_id:
                        break
        
        if not latest_email_id:
            print("⚠️ No recent ASM emails found, using fallback")
            # Try to find ANY email with ASM in subject from last 50 emails
            result, all_emails = mail.search(None, 'ALL')
            if result == 'OK' and all_emails[0]:
                email_list = all_emails[0].split()
                # Check last 50 emails
                for email_id in reversed(email_list[-50:]):
                    try:
                        result, msg_data = mail.fetch(email_id, '(ENVELOPE)')
                        if result == 'OK':
                            # Simple check for ASM in the envelope
                            envelope_str = str(msg_data[0][1])
                            if 'ASM' in envelope_str or 'Additional' in envelope_str:
                                latest_email_id = email_id
                                break
                    except:
                        continue
        
        if not latest_email_id:
            print("❌ No ASM emails found at all")
            return
        
        print(f"📧 Fetching email {latest_email_id.decode()}...")
        result, msg_data = mail.fetch(latest_email_id, '(RFC822)')
        
        if result == 'OK':
            email_message = email.message_from_bytes(msg_data[0][1])
            subject = email_message.get('Subject', 'No Subject')
            date_received = email_message.get('Date', 'No Date')
            print(f"✅ Email subject: {subject}")
            print(f"📅 Email date: {date_received}")
            
            # Extract date from email subject or use current date
            email_date_match = re.search(r'(\d{1,2}[-/.]\d{1,2}[-/.]\d{4})', subject)
            if email_date_match:
                email_date = email_date_match.group(1)
            else:
                email_date = datetime.now().strftime("%Y-%m-%d")
            
            # Download PDF attachment
            pdf_downloaded = False
            for part in email_message.walk():
                if part.get_content_disposition() == 'attachment':
                    filename = part.get_filename()
                    
                    if filename and ('ASM' in filename or 'Additional' in filename) and filename.lower().endswith('.pdf'):
                        print(f"🔍 Found ASM PDF: {filename}")
                        
                        # Get attachment data
                        attachment_data = part.get_payload(decode=True)
                        
                        # Save to folder
                        pdf_path = os.path.join(DOWNLOAD_FOLDER, f"ASM_latest.pdf")
                        
                        with open(pdf_path, 'wb') as f:
                            f.write(attachment_data)
                        
                        print(f"✅ Downloaded: {pdf_path}")
                        pdf_downloaded = True
                        
                        # Extract data from PDF
                        print("\n📄 EXTRACTING DATA FROM PDF")
                        print("=" * 30)
                        
                        # Read PDF content
                        with open(pdf_path, 'rb') as file:
                            content = file.read()
                        
                        # Extract text
                        text_content = content.decode('latin-1', errors='ignore')
                        print(f"✅ Extracted {len(text_content)} characters of text")
                        
                        # Extract client data - look for the table format
                        client_data = []
                        
                        # Look for the specific pattern from the PDF
                        # Find lines with client codes and margin amounts
                        lines = text_content.split('\n')
                        for i, line in enumerate(lines):
                            if 'Z00018' in line or 'Z00008' in line:
                                print(f"📋 Found line: {line}")
                                
                                # Extract client code
                                client_match = re.search(r'(Z\d{5})', line)
                                if client_match:
                                    client_code = client_match.group(1)
                                    
                                    # Look for the margin amount in the same line or nearby lines
                                    # Pattern: look for number followed by decimal (like 280.99)
                                    margin_matches = re.findall(r'(\d+\.\d{2})', line)
                                    
                                    if margin_matches:
                                        # Usually the Additional Surveillance Margin is the last number in the line
                                        margin = margin_matches[-1]
                                        
                                        client_data.append({
                                            'client_code': client_code,
                                            'additional_margin': margin
                                        })
                                        print(f"✅ Found: {client_code} -> ₹{margin} Cr")
                        
                        # Extract reference number
                        ref_match = re.search(r'NCL/RMG/\d{4}/\d{5}', text_content)
                        ref_number = ref_match.group(0) if ref_match else None
                        
                        # Extract date from PDF content
                        date_match = re.search(r'(August \d{1,2}, \d{4})', text_content)
                        pdf_date = date_match.group(1) if date_match else email_date
                        
                        print(f"\n📊 EXTRACTED DATA:")
                        print(f"Reference: {ref_number}")
                        print(f"PDF Date: {pdf_date}")
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
