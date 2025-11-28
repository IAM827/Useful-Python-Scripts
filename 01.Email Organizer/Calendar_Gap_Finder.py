import requests
import time
from datetime import datetime, timedelta

# ============================================================
# CONFIGURATION
# ============================================================

CLIENT_ID = "YOUR_CLIENT_ID"

# Working hours (24-hour format)
WORK_START_HOUR = 8  # 8 AM
WORK_END_HOUR = 21   # 9 PM

# Minimum gap duration to report (in hours)
MIN_GAP_DURATION = 2

# How many days ahead to check
DAYS_AHEAD = 7  # Check next 7 days

# ============================================================
# AUTHENTICATION
# ============================================================

def get_access_token_device_code():
    """Get access token using device code flow"""
    
    device_code_url = "https://login.microsoftonline.com/common/oauth2/v2.0/devicecode"
    
    data = {
        "client_id": CLIENT_ID,
        "scope": "Calendars.Read offline_access"
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
# CALENDAR FUNCTIONS
# ============================================================

def get_calendar_events(access_token, start_date, end_date):
    """Get calendar events within date range"""
    
    url = "https://graph.microsoft.com/v1.0/me/calendar/calendarView"
    
    params = {
        "startDateTime": start_date.isoformat(),
        "endDateTime": end_date.isoformat(),
        "$select": "subject,start,end",
        "$orderby": "start/dateTime"
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
            print(f"❌ Error fetching events: {response.status_code}")
            return []
    except Exception as e:
        print(f"❌ Error: {e}")
        return []

def find_gaps_for_day(events, date):
    """Find free time gaps in a specific day"""
    
    # Define working hours for the day
    work_start = date.replace(hour=WORK_START_HOUR, minute=0, second=0, microsecond=0)
    work_end = date.replace(hour=WORK_END_HOUR, minute=0, second=0, microsecond=0)
    
    # Filter events for this specific day within working hours
    day_events = []
    for event in events:
        start_dt = datetime.fromisoformat(event['start']['dateTime'].replace('Z', '+00:00'))
        end_dt = datetime.fromisoformat(event['end']['dateTime'].replace('Z', '+00:00'))
        
        # Check if event overlaps with this day's working hours
        if start_dt.date() == date.date() or end_dt.date() == date.date():
            # Clip event times to working hours
            clipped_start = max(start_dt.replace(tzinfo=None), work_start)
            clipped_end = min(end_dt.replace(tzinfo=None), work_end)
            
            if clipped_start < clipped_end:
                day_events.append({
                    'start': clipped_start,
                    'end': clipped_end,
                    'subject': event.get('subject', 'No Subject')
                })
    
    # Sort events by start time
    day_events.sort(key=lambda x: x['start'])
    
    # Find gaps
    gaps = []
    current_time = work_start
    
    for event in day_events:
        # Check if there's a gap before this event
        if current_time < event['start']:
            gap_duration = (event['start'] - current_time).total_seconds() / 3600  # Convert to hours
            
            if gap_duration >= MIN_GAP_DURATION:
                gaps.append({
                    'start': current_time,
                    'end': event['start'],
                    'duration': gap_duration
                })
        
        # Move current_time to end of this event
        current_time = max(current_time, event['end'])
    
    # Check if there's a gap after the last event until end of work day
    if current_time < work_end:
        gap_duration = (work_end - current_time).total_seconds() / 3600
        
        if gap_duration >= MIN_GAP_DURATION:
            gaps.append({
                'start': current_time,
                'end': work_end,
                'duration': gap_duration
            })
    
    return gaps, day_events

# ============================================================
# MAIN SCRIPT
# ============================================================

def main():
    print("\n" + "="*60)
    print("CALENDAR GAP FINDER - FIND FREE TIME SLOTS")
    print("="*60 + "\n")
    
    print(f"✓ Client ID configured")
    print(f"✓ Working hours: {WORK_START_HOUR}:00 - {WORK_END_HOUR}:00")
    print(f"✓ Minimum gap: {MIN_GAP_DURATION} hours")
    print(f"✓ Checking: Next {DAYS_AHEAD} days")
    print()
    
    # Authenticate
    access_token = get_access_token_device_code()
    
    if not access_token:
        print("❌ Authentication failed!")
        return
    
    # Calculate date range
    start_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    end_date = start_date + timedelta(days=DAYS_AHEAD)
    
    print(f"Fetching calendar events from {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}...")
    
    # Get all events in the range
    events = get_calendar_events(access_token, start_date, end_date)
    
    print(f"✓ Found {len(events)} events\n")
    
    print("="*60)
    print("FREE TIME SLOTS")
    print("="*60 + "\n")
    
    total_gaps = 0
    total_free_hours = 0
    
    # Check each day
    for day_offset in range(DAYS_AHEAD):
        current_date = start_date + timedelta(days=day_offset)
        
        # Skip weekends if desired (optional)
        # if current_date.weekday() >= 5:  # Saturday = 5, Sunday = 6
        #     continue
        
        gaps, day_events = find_gaps_for_day(events, current_date)
        
        if gaps or day_events:
            print(f"📅 {current_date.strftime('%A, %B %d, %Y')}")
            print("-" * 60)
            
            if day_events:
                print(f"   Scheduled meetings: {len(day_events)}")
                for event in day_events:
                    print(f"   • {event['start'].strftime('%I:%M %p')} - {event['end'].strftime('%I:%M %p')}: {event['subject']}")
            else:
                print(f"   No meetings scheduled")
            
            if gaps:
                print(f"\n   ✓ Free time slots ({len(gaps)} found):")
                for gap in gaps:
                    print(f"   → {gap['start'].strftime('%I:%M %p')} - {gap['end'].strftime('%I:%M %p')} ({gap['duration']:.1f} hours)")
                    total_gaps += 1
                    total_free_hours += gap['duration']
            else:
                print(f"\n   ❌ No free slots available (minimum {MIN_GAP_DURATION} hours)")
            
            print()
    
    # Summary
    print("="*60)
    print("SUMMARY")
    print("="*60)
    print(f"Total free slots found: {total_gaps}")
    print(f"Total free hours: {total_free_hours:.1f} hours")
    print(f"Average per day: {total_free_hours / DAYS_AHEAD:.1f} hours")
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
