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
    
    # Validate that all secrets are available
    if not all([GMAIL_USER, GMAIL_APP_PASSWORD, WEBHOOK_URL]):
        print("❌ Missing required environment variables!")
        print(f"GMAIL_USER: {'✅' if GMAIL_USER else '❌'}")
        print(f"GMAIL_APP_PASSWORD: {'✅' if GMAIL_APP_PASSWORD else '❌'}")
        print(f"WEBHOOK_URL: {'✅' if WEBHOOK_URL else '❌'}")
        return
    
    # Create download folder in current directory (GitHub Actions environment)
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
        
        # Search for ASM emails instead of using hardcoded UID
        print("🔍 Searching for recent ASM emails...")
        search_criteria = '(SUBJECT "ASM" FROM "nse")'
        result, email_ids = mail.search(None, search_criteria)
        
        if result == 'OK' and email_ids[0]:
            # Get the most recent email (last in the list)
            latest_email_id = email_ids[0].split()[-1]
            print(f"📧 Found email ID: {latest_email_id.decode()}")
        else:
            # Fallback to the hardcoded UID if search fails
            print("⚠️ Search failed, using fallback UID 908")
            latest_email_id = b'908'
        
        print(f"📧 Fetching email...")
        result, msg_data = mail.fetch(latest_email_id, '(RFC822)')
        
        if result == 'OK':
            email_message = email.message_from_bytes(msg_data[0][1])
            subject = email_message.get('Subject', 'No Subject')
            print(f"✅ Email subject: {subject}")
            
            # Download PDF attachment
            pdf_downloaded = False
            for part in email_message.walk():
                if part.get_content_disposition() == 'attachment':
                    filename = part.get_filename()
                    
                    if filename and 'ASM' in filename and filename.lower().endswith('.pdf'):
                        print(f"🔍 Found ASM PDF: {filename}")
                        
                        # Get attachment data
                        attachment_data = part.get_payload(decode=True)
                        
                        # Save to folder (with current date)
                        current_date = datetime.now().strftime("%Y-%m-%d")
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
                        
                        print(f"\n📊 EXTRACTED DATA:")
                        print(f"Reference: {ref_number}")
                        for client in client_data:
                            print(f"{client['client_code']}: ₹{client['additional_margin']} Cr")
                        
                        # Create message
                        if client_data:
                            message = f"ASM as on {current_date}\n\n"
                            message += "Client Code / Additional Surveillance Margin:\n"
                            for client in client_data:
                                message += f"{client['client_code']}: ₹{client['additional_margin']} Cr\n"
                            if ref_number:
                                message += f"\nReference: {ref_number}"
                        else:
                            message = f"ASM is NIL for {current_date}"
                        
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
                # Send notification that no ASM was found
                try:
                    current_date = datetime.now().strftime("%Y-%m-%d")
                    message = f"No ASM PDF found for {current_date}"
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
                            print("✅ 'No ASM' notification sent successfully!")
                except Exception as e:
                    print(f"❌ Error sending 'No ASM' notification: {e}")
        else:
            print("❌ Failed to fetch email")
        
        # Close connection
        mail.close()
        mail.logout()
        
    except Exception as e:
        print(f"❌ Error: {e}")
        # Send error notification
        try:
            error_message = f"ASM Script Error: {str(e)}"
            payload = {"text": error_message}
            data = json.dumps(payload).encode('utf-8')
            
            req = urllib.request.Request(
                WEBHOOK_URL,
                data=data,
                headers={'Content-Type': 'application/json'}
            )
            
            context = ssl.create_default_context()
            with urllib.request.urlopen(req, context=context, timeout=30) as response:
                if response.getcode() == 200:
                    print("✅ Error notification sent successfully!")
        except Exception as notify_error:
            print(f"❌ Failed to send error notification: {notify_error}")

if __name__ == "__main__":
    download_asm_directly()
