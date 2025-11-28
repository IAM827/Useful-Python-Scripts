import requests
from datetime import datetime, timedelta
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.units import inch
import time

# ============================================================
# CONFIGURATION
# ============================================================

CLIENT_ID = "YOUR_CLIENT_ID"
YOUR_EMAIL = "YOUR_EMAIL_ADDRESS"

# Summary settings
SUMMARY_TYPE = "weekly"  # Options: "daily", "weekly", "monthly"

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

def get_date_range(summary_type):
    """Calculate date range based on summary type"""
    
    today = datetime.now()
    
    if summary_type == "daily":
        start_date = today.replace(hour=0, minute=0, second=0, microsecond=0)
        end_date = today.replace(hour=23, minute=59, second=59, microsecond=999999)
    elif summary_type == "weekly":
        # Get start of current week (Monday)
        start_date = today - timedelta(days=today.weekday())
        start_date = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
        # Get end of current week (Sunday)
        end_date = start_date + timedelta(days=6, hours=23, minutes=59, seconds=59)
    elif summary_type == "monthly":
        # Get start of current month
        start_date = today.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        # Get end of current month
        if today.month == 12:
            end_date = today.replace(year=today.year + 1, month=1, day=1) - timedelta(seconds=1)
        else:
            end_date = today.replace(month=today.month + 1, day=1) - timedelta(seconds=1)
    else:
        start_date = today
        end_date = today
    
    return start_date, end_date

def get_calendar_events(access_token, start_date, end_date):
    """Get calendar events within date range"""
    
    url = "https://graph.microsoft.com/v1.0/me/calendar/calendarView"
    
    params = {
        "startDateTime": start_date.isoformat() + "Z",
        "endDateTime": end_date.isoformat() + "Z",
        "$select": "subject,start,end,location,attendees,organizer,bodyPreview",
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

# ============================================================
# PDF GENERATION
# ============================================================

def create_pdf_summary(events, start_date, end_date, filename):
    """Generate PDF summary of meetings"""
    
    doc = SimpleDocTemplate(filename, pagesize=letter)
    story = []
    styles = getSampleStyleSheet()
    
    # Custom styles
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=24,
        textColor=colors.HexColor('#1a73e8'),
        spaceAfter=30,
        alignment=1  # Center
    )
    
    heading_style = ParagraphStyle(
        'CustomHeading',
        parent=styles['Heading2'],
        fontSize=16,
        textColor=colors.HexColor('#1a73e8'),
        spaceAfter=12
    )
    
    # Title
    title = Paragraph(f"Meeting Summary Report", title_style)
    story.append(title)
    
    # Date range
    date_range = Paragraph(
        f"<b>Period:</b> {start_date.strftime('%B %d, %Y')} - {end_date.strftime('%B %d, %Y')}",
        styles['Normal']
    )
    story.append(date_range)
    story.append(Spacer(1, 0.3*inch))
    
    # Summary statistics
    total_meetings = len(events)
    total_hours = sum([
        (datetime.fromisoformat(event['end']['dateTime']) - 
         datetime.fromisoformat(event['start']['dateTime'])).total_seconds() / 3600
        for event in events
    ])
    
    stats_data = [
        ["Total Meetings:", str(total_meetings)],
        ["Total Hours:", f"{total_hours:.1f} hours"],
        ["Average per Day:", f"{total_meetings / 7:.1f} meetings"]
    ]
    
    stats_table = Table(stats_data, colWidths=[2*inch, 2*inch])
    stats_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#f0f0f0')),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 12),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
        ('GRID', (0, 0), (-1, -1), 1, colors.white)
    ]))
    
    story.append(stats_table)
    story.append(Spacer(1, 0.5*inch))
    
    # Group events by day
    events_by_day = {}
    for event in events:
        start_dt = datetime.fromisoformat(event['start']['dateTime'])
        day_key = start_dt.strftime('%Y-%m-%d')
        
        if day_key not in events_by_day:
            events_by_day[day_key] = []
        events_by_day[day_key].append(event)
    
    # Meetings details
    story.append(Paragraph("Meeting Details", heading_style))
    story.append(Spacer(1, 0.2*inch))
    
    for day_key in sorted(events_by_day.keys()):
        day_events = events_by_day[day_key]
        day_date = datetime.strptime(day_key, '%Y-%m-%d')
        
        # Day header
        day_header = Paragraph(
            f"<b>{day_date.strftime('%A, %B %d, %Y')}</b> ({len(day_events)} meetings)",
            styles['Heading3']
        )
        story.append(day_header)
        story.append(Spacer(1, 0.1*inch))
        
        # Meeting table for this day
        meeting_data = [["Time", "Subject", "Location", "Attendees"]]
        
        for event in day_events:
            start_dt = datetime.fromisoformat(event['start']['dateTime'])
            end_dt = datetime.fromisoformat(event['end']['dateTime'])
            
            time_str = f"{start_dt.strftime('%I:%M %p')} - {end_dt.strftime('%I:%M %p')}"
            subject = event.get('subject', 'No Subject')
            location = event.get('location', {}).get('displayName', 'No Location')
            
            attendees = event.get('attendees', [])
            attendee_count = len(attendees)
            attendee_str = f"{attendee_count} attendees" if attendee_count > 0 else "No attendees"
            
            meeting_data.append([time_str, subject, location, attendee_str])
        
        meeting_table = Table(meeting_data, colWidths=[1.5*inch, 2.5*inch, 1.5*inch, 1.2*inch])
        meeting_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a73e8')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('VALIGN', (0, 0), (-1, -1), 'TOP')
        ]))
        
        story.append(meeting_table)
        story.append(Spacer(1, 0.3*inch))
    
    # Build PDF
    doc.build(story)
    print(f"✓ PDF generated: {filename}")

# ============================================================
# MAIN SCRIPT
# ============================================================

def main():
    print("\n" + "="*60)
    print("MEETING SUMMARY GENERATOR - WEEKLY PDF REPORT")
    print("="*60 + "\n")
    
    # Authenticate
    access_token = get_access_token_device_code()
    
    if not access_token:
        print("❌ Authentication failed!")
        return
    
    # Get date range
    start_date, end_date = get_date_range(SUMMARY_TYPE)
    
    print(f"Fetching meetings from {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}...")
    
    # Get calendar events
    events = get_calendar_events(access_token, start_date, end_date)
    
    if not events:
        print("❌ No meetings found in the specified date range.")
        return
    
    print(f"✓ Found {len(events)} meetings")
    
    # Generate PDF
    filename = f"meeting_summary_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.pdf"
    
    print(f"Generating PDF report...")
    create_pdf_summary(events, start_date, end_date, filename)
    
    print("\n" + "="*60)
    print("SUMMARY COMPLETE")
    print("="*60)
    print(f"Report saved as: {filename}")
    print(f"Total meetings: {len(events)}")
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
