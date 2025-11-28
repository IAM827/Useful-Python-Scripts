import requests
import time
from datetime import datetime

# ============================================================
# CONFIGURATION
# ============================================================

CLIENT_ID = "YOUR_CLIENT_ID"

# Auto-reply settings
AUTO_REPLY_ENABLED = False  # Set to True to enable auto-replies
AUTO_REPLY_MESSAGE = """Greetings {sender_name},

Please be advised that your email has been received, however, I'm not available to respond at the moment and will do so upon my return. 

If it's urgent, you're welcome to give me a call.

Kindest Regards."""

# How often to check for new emails (in seconds)
CHECK_INTERVAL = 300  # Check every 5 minutes

# Track replied emails to avoid duplicate responses
replied_emails = set()

# ============================================================
# AUTHENTICATION
# ============================================================

def get_access_token_device_code():
    """Get access token using device code flow"""
    
    device_code_url = "https://login.microsoftonline.com/common/oauth2/v2.0/devicecode"
    
    data = {
        "client_id": CLIENT_ID,
        "scope": "Mail.ReadWrite Mail.Send offline_access"
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

def get_unread_emails(access_token, max_emails=20):
    """Get unread emails from inbox"""
    
    url = "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages"
    
    params = {
        "$top": max_emails,
        "$filter": "isRead eq false",
        "$select": "id,subject,from,receivedDateTime,conversationId",
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

def send_auto_reply(access_token, to_email, to_name, subject):
    """Send automatic reply to email"""
    
    url = "https://graph.microsoft.com/v1.0/me/sendMail"
    
    # Format the message with sender's name
    message_body = AUTO_REPLY_MESSAGE.format(sender_name=to_name)
    
    email_data = {
        "message": {
            "subject": f"Re: {subject}",
            "body": {
                "contentType": "Text",
                "content": message_body
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": to_email,
                        "name": to_name
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
        print(f"❌ Error sending reply: {e}")
        return False

def mark_as_read(access_token, email_id):
    """Mark email as read"""
    
    url = f"https://graph.microsoft.com/v1.0/me/messages/{email_id}"
    
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    
    data = {"isRead": True}
    
    try:
        response = requests.patch(url, headers=headers, json=data)
        return response.status_code == 200
    except Exception as e:
        print(f"❌ Error marking as read: {e}")
        return False

# ============================================================
# MAIN SCRIPT
# ============================================================

def main():
    print("\n" + "="*60)
    print("EMAIL RESPONSE BOT - VACATION AUTO-REPLY")
    print("="*60 + "\n")
    
    print(f"✓ Client ID configured")
    print(f"✓ Check interval: Every {CHECK_INTERVAL} seconds")
    print(f"✓ Auto-reply status: {'ENABLED ✓' if AUTO_REPLY_ENABLED else 'DISABLED ✗'}")
    
    if not AUTO_REPLY_ENABLED:
        print("\n⚠️  WARNING: Auto-reply is currently DISABLED!")
        print("To enable, set AUTO_REPLY_ENABLED = True in the script")
        print("="*60 + "\n")
        response = input("Continue anyway? (y/n): ")
        if response.lower() != 'y':
            return
    
    print()
    
    # Authenticate
    access_token = get_access_token_device_code()
    
    if not access_token:
        print("❌ Authentication failed!")
        return
    
    print("="*60)
    print("MONITORING INBOX FOR NEW EMAILS")
    print("="*60)
    print("Press Ctrl+C to stop\n")
    
    replied_count = 0
    check_count = 0
    
    try:
        while True:
            check_count += 1
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            # Get unread emails
            unread_emails = get_unread_emails(access_token)
            
            if unread_emails:
                print(f"\n[{current_time}] Found {len(unread_emails)} unread email(s)")
                
                for email in unread_emails:
                    email_id = email.get('id')
                    conversation_id = email.get('conversationId')
                    subject = email.get('subject', 'No Subject')
                    sender = email.get('from', {}).get('emailAddress', {})
                    sender_email = sender.get('address', 'Unknown')
                    sender_name = sender.get('name', 'Unknown')
                    
                    # Skip if we've already replied to this conversation
                    if conversation_id in replied_emails:
                        print(f"  ⏭️  Skipped (already replied): {subject[:40]}...")
                        continue
                    
                    print(f"  📧 New email: {subject[:40]}...")
                    print(f"     From: {sender_name} ({sender_email})")
                    
                    if AUTO_REPLY_ENABLED:
                        # Send auto-reply
                        if send_auto_reply(access_token, sender_email, sender_name, subject):
                            replied_count += 1
                            replied_emails.add(conversation_id)
                            print(f"     ✓ Auto-reply sent (Total: {replied_count})")
                            
                            # Mark as read
                            mark_as_read(access_token, email_id)
                        else:
                            print(f"     ❌ Failed to send auto-reply")
                    else:
                        print(f"     ⏸️  Auto-reply disabled - no action taken")
            else:
                if check_count % 10 == 0:
                    print(f"[{current_time}] No new emails. Total replies sent: {replied_count}")
            
            # Wait before next check
            time.sleep(CHECK_INTERVAL)
            
    except KeyboardInterrupt:
        print("\n\n" + "="*60)
        print("STOPPED BY USER")
        print("="*60)
        print(f"Total auto-replies sent: {replied_count}")
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
