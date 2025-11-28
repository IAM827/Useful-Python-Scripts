import requests
import time
import re
from datetime import datetime, timedelta

# ============================================================
# CONFIGURATION
# ============================================================

CLIENT_ID = "YOUR_CLIENT_ID"
YOUR_EMAIL = "YOUR_EMAIL_ADDRESS"

# Keywords that trigger reminder creation
REMINDER_KEYWORDS = ["deadline", "due", "reminder", "meeting set-up"]

# How many days before the deadline to send reminder
REMINDER_DAYS_BEFORE = 1

# How often to check for new emails (in seconds)
CHECK_INTERVAL = 300  # Check every 5 minutes

# Track processed emails to avoid duplicates
processed_emails = set()

# ============================================================
# AUTHENTICATION
# ============================================================

def get_access_token_device_code():
    """Get access token using device code flow"""
    
    device_code_url = "https://login.microsoftonline.com/common/oauth2/v2.0/devicecode"
    
    data = {
        "client_id": CLIENT_ID,
        "scope": "Mail.Read Tasks.ReadWrite Calendars.ReadWrite offline_access"
    }
    
    print("Requesting authentication...")
    
    try:
        response = requests.post(device_code_url, data=data)
        response.raise_for_status()
        device_code_response = response.json()
        
        if 'error' in device_code_response or 'message' not in device_code_response:
            print(f"\n❌ Authentication error")
            return None
        
    except requests.exceptions.RequestException as e:
        print(f"\n❌ Network error: {e}")
        return None
    
    print("\n" + "="*60)
    print("AUTHENTICATION REQUIRED")
    print("="*60)
    print(f"\n{device_code_response['message']}\n")
    print(f"User Code: {device_code_response['user_code']}")
    print(f"Visit: {device_code_response['verification_uri']}")
    print("\nWaiting for authentication...")
    print("="*60 + "\n")
    
    token_url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
    
    token_data = {
        "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
        "client_id": CLIENT_ID,
        "device_code": device_code_response['device_code']
    }
    
    while True:
        try:
            token_response = requests.post(token_url, data=token_data)
            token_result = token_response.json()
            
            if "access_token" in token_result:
                print("✓ Authentication successful!\n")
                return token_result["access_token"]
            elif token_result.get("error") == "authorization_pending":
                print(".", end="", flush=True)
                time.sleep(5)
            else:
                print(f"\n❌ Error: {token_result.get('error_description', 'Unknown error')}")
                return None
        except Exception as e:
            print(f"\n❌ Error: {e}")
            return None

# ============================================================
# EMAIL FUNCTIONS
# ============================================================

def get_recent_emails(access_token, max_emails=20):
    """Get recent emails from inbox"""
    
    url = "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages"
    
    params = {
        "$top": max_emails,
        "$select": "id,subject,bodyPreview,body,receivedDateTime,from",
        "$orderby": "receivedDateTime DESC"
    }
    
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    
    try:
        response = requests.get(url, headers=headers, params=params)
        if response.status_code == 200:
            return response.json().get('value', [])
        else:
            print(f"❌ Error fetching emails: {response.status_code}")
            return []
    except Exception as e:
        print(f"❌ Error: {e}")
        return []

def extract_dates_from_text(text):
    """Extract potential dates from email text"""
    
    dates = []
    
    # Common date patterns
    patterns = [
        r'\b(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})\b',  # MM/DD/YYYY or DD-MM-YYYY
        r'\b(\d{4})[/-](\d{1,2})[/-](\d{1,2})\b',    # YYYY-MM-DD
        r'\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]* (\d{1,2}),? (\d{4})\b',  # Month DD, YYYY
        r'\b(\d{1,2}) (Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]* (\d{4})\b',  # DD Month YYYY
    ]
    
    month_map = {
        'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
        'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
    }
    
    for pattern in patterns:
        matches = re.finditer(pattern, text, re.IGNORECASE)
        for match in matches:
            try:
                groups = match.groups()
                
                if len(groups) == 3:
                    if groups[0].isdigit() and groups[1].isdigit():
                        # Numeric date format
                        if len(groups[0]) == 4:  # YYYY-MM-DD
                            year, month, day = int(groups[0]), int(groups[1]), int(groups[2])
                        else:  # MM/DD/YYYY or DD/MM/YYYY - assume MM/DD/YYYY
                            month, day, year = int(groups[0]), int(groups[1]), int(groups[2])
                            if year < 100:
                                year += 2000
                    else:
                        # Month name format
                        if groups[0].isdigit():  # DD Month YYYY
                            day = int(groups[0])
                            month = month_map.get(groups[1][:3].lower())
                            year = int(groups[2])
                        else:  # Month DD, YYYY
                            month = month_map.get(groups[0][:3].lower())
                            day = int(groups[1])
                            year = int(groups[2])
                    
                    date = datetime(year, month, day)
                    if date >= datetime.now():  # Only future dates
                        dates.append(date)
            except (ValueError, TypeError):
                continue
    
    return dates

def check_for_keywords(email):
    """Check if email contains reminder keywords"""
    
    subject = email.get('subject', '').lower()
    body = email.get('bodyPreview', '').lower()
    
    for keyword in REMINDER_KEYWORDS:
        if keyword.lower() in subject or keyword.lower() in body:
            return True
    
    return False

# ============================================================
# REMINDER FUNCTIONS
# ============================================================

def create_outlook_reminder(access_token, subject, date, email_subject):
    """Create a reminder/task in Outlook"""
    
    url = "https://graph.microsoft.com/v1.0/me/outlook/tasks"
    
    reminder_date = date - timedelta(days=REMINDER_DAYS_BEFORE)
    
    task_data = {
        "subject": subject,
        "body": {
            "contentType": "text",
            "content": f"Reminder for: {email_subject}\nDeadline: {date.strftime('%B %d, %Y')}"
        },
        "dueDateTime": {
            "dateTime": date.isoformat(),
            "timeZone": "UTC"
        },
        "reminderDateTime": {
            "dateTime": reminder_date.isoformat(),
            "timeZone": "UTC"
        },
        "importance": "high"
    }
    
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    
    try:
        response = requests.post(url, headers=headers, json=task_data)
        return response.status_code == 201
    except Exception as e:
        print(f"❌ Error creating reminder: {e}")
        return False

def send_reminder_email(access_token, subject, date, email_subject):
    """Send reminder email to yourself"""
    
    url = "https://graph.microsoft.com/v1.0/me/sendMail"
    
    reminder_date = date - timedelta(days=REMINDER_DAYS_BEFORE)
    
    message_body = f"""This is an automated reminder.

Subject: {email_subject}
Deadline: {date.strftime('%A, %B %d, %Y')}

You will be reminded on: {reminder_date.strftime('%A, %B %d, %Y')}

This reminder was automatically created by the Reminder Generator script.
"""
    
    email_data = {
        "message": {
            "subject": f"Reminder: {subject}",
            "body": {
                "contentType": "Text",
                "content": message_body
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": YOUR_EMAIL
                    }
                }
            ]
        }
    }
    
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    
    try:
        response = requests.post(url, headers=headers, json=email_data)
        return response.status_code == 202
    except Exception as e:
        print(f"❌ Error sending email: {e}")
        return False

# ============================================================
# MAIN SCRIPT
# ============================================================

def main():
    print("\n" + "="*60)
    print("REMINDER GENERATOR - AUTO-CREATE REMINDERS FROM EMAILS")
    print("="*60 + "\n")
    
    print(f"✓ Client ID configured")
    print(f"✓ Your email: {YOUR_EMAIL}")
    print(f"✓ Keywords: {', '.join(REMINDER_KEYWORDS)}")
    print(f"✓ Reminder: {REMINDER_DAYS_BEFORE} day(s) before deadline")
    print(f"✓ Check interval: Every {CHECK_INTERVAL} seconds")
    print()
    
    # Authenticate
    access_token = get_access_token_device_code()
    
    if not access_token:
        print("❌ Authentication failed!")
        return
    
    print("="*60)
    print("MONITORING EMAILS FOR REMINDERS")
    print("="*60)
    print("Press Ctrl+C to stop\n")
    
    reminders_created = 0
    check_count = 0
    
    try:
        while True:
            check_count += 1
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            # Get recent emails
            emails = get_recent_emails(access_token)
            
            if emails:
                for email in emails:
                    email_id = email.get('id')
                    
                    # Skip if already processed
                    if email_id in processed_emails:
                        continue
                    
                    subject = email.get('subject', 'No Subject')
                    body_preview = email.get('bodyPreview', '')
                    full_body = email.get('body', {}).get('content', '')
                    
                    # Check if email contains reminder keywords
                    if check_for_keywords(email):
                        print(f"\n[{current_time}] 📧 Found potential reminder:")
                        print(f"  Subject: {subject}")
                        
                        # Extract dates from email
                        dates = extract_dates_from_text(subject + " " + body_preview + " " + full_body)
                        
                        if dates:
                            for date in dates[:1]:  # Use first date found
                                print(f"  Deadline found: {date.strftime('%B %d, %Y')}")
                                
                                # Create reminder
                                reminder_subject = f"Deadline: {subject[:50]}"
                                
                                if create_outlook_reminder(access_token, reminder_subject, date, subject):
                                    reminders_created += 1
                                    print(f"  ✓ Reminder created in Outlook Tasks")
                                
                                if send_reminder_email(access_token, reminder_subject, date, subject):
                                    print(f"  ✓ Reminder email sent")
                                
                                print(f"  Total reminders: {reminders_created}")
                        else:
                            print(f"  ⚠️  No date found in email - skipping")
                        
                        processed_emails.add(email_id)
            
            if check_count % 10 == 0:
                print(f"[{current_time}] Checked emails. Total reminders created: {reminders_created}")
            
            # Wait before next check
            time.sleep(CHECK_INTERVAL)
            
    except KeyboardInterrupt:
        print("\n\n" + "="*60)
        print("STOPPED BY USER")
        print("="*60)
        print(f"Total reminders created: {reminders_created}")
        print("="*60 + "\n")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("\n" + "="*60)
        print("ERROR OCCURRED")
        print("="*60)
        print(f"\n{type(e).__name__}: {str(e)}\n")
        import traceback
        traceback.print_exc()
    finally:
        input("\nPress Enter to close...")
