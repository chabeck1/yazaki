import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from icalendar import Calendar, Event, vCalAddress, vText
from datetime import datetime, timedelta
import pytz
import uuid

# --- 1. CONFIGURE YOUR DETAILS HERE ---
SENDER_GMAIL = "thepattyc123@gmail.com"      # Your full Gmail address
GMAIL_APP_PASSWORD = "avds ognu jdsv meeu"    # The 16-character App Password (no spaces)
RECIPIENT_EMAIL = "charles.beck@us.yazaki.com"
SENDER_NAME = "Test Script"                  # The name that will appear as the organizer
RECIPIENT_NAME = "Charles Beck"              # The name of the person you're inviting

def send_outlook_compatible_invite(
    sender_email,
    app_password,
    recipient_email,
    recipient_name,
    organizer_name,
    subject,
    description,
    start_time,
    duration_hours=1
):
    """
    Creates and sends a calendar invitation that Outlook will parse
    (explicit method=REQUEST and UTC timestamps).
    """
    # --- Create the Main Message ---
    msg = MIMEMultipart('alternative')
    msg['From'] = f'"{organizer_name}" <{sender_email}>'
    msg['To']   = f'"{recipient_name}" <{recipient_email}>'
    msg['Subject'] = subject

    # --- Plain-text fallback ---
    msg.attach(MIMEText(description, 'plain'))

    # --- Build the iCalendar object ---
    cal = Calendar()
    cal.add('prodid', '-//My Test Script//EN')
    cal.add('version', '2.0')
    cal.add('method', 'REQUEST')

    event = Event()
    # Organizer
    organizer = vCalAddress(f'MAILTO:{sender_email}')
    organizer.params['cn'] = vText(organizer_name)
    event.add('organizer', organizer)
    # Attendee
    attendee = vCalAddress(f'MAILTO:{recipient_email}')
    attendee.params['cn']       = vText(recipient_name)
    attendee.params['ROLE']     = vText('REQ-PARTICIPANT')
    attendee.params['PARTSTAT'] = vText('NEEDS-ACTION')
    attendee.params['RSVP']     = vText('TRUE')
    event.add('attendee', attendee)

    event.add('summary', subject)
    # -- FORCE UTC timestamps to avoid missing VTIMEZONE blocks --
    start_utc = start_time.astimezone(pytz.utc)
    end_utc   = start_utc + timedelta(hours=duration_hours)
    event.add('dtstart', start_utc)
    event.add('dtend',   end_utc)
    event.add('dtstamp', datetime.now(pytz.utc))
    event['uid'] = str(uuid.uuid4())
    event.add('description', description)

    cal.add_component(event)

    # --- Create the calendar MIME part with explicit method=REQUEST ---
    ical_str = cal.to_ical().decode('utf-8')
    ical_part = MIMEText(ical_str, _subtype='calendar')
    ical_part.replace_header(
        'Content-Type',
        'text/calendar; charset="UTF-8"; method=REQUEST; name="invite.ics"'
    )
    msg.attach(ical_part)

    # --- Send via Gmail SMTP over SSL ---
    try:
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(sender_email, app_password)
        server.sendmail(sender_email, [recipient_email], msg.as_string())
        server.quit()
        print(f"[+] Success! Invitation sent to {recipient_email}")
    except Exception as e:
        print(f"[!!!] FAILED to send invitation.")
        print(f"    Error: {e}")

if __name__ == "__main__":
    # --- 2. RUN THE SCRIPT ---
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    test_subject     = f"Windchill Test (Corrected Script) @ {now_str}"
    test_description = "This is a test meeting invite using the corrected script with proper attendee parameters."

    # schedule 2 minutes from now
    local_tz   = datetime.now().astimezone().tzinfo
    start_time = (datetime.now() + timedelta(minutes=2)).astimezone(local_tz)

    print(f"[*] Generating test meeting: '{test_subject}'")
    send_outlook_compatible_invite(
        sender_email=SENDER_GMAIL,
        app_password=GMAIL_APP_PASSWORD,
        recipient_email=RECIPIENT_EMAIL,
        recipient_name=RECIPIENT_NAME,
        organizer_name=SENDER_NAME,
        subject=test_subject,
        description=test_description,
        start_time=start_time
    )
