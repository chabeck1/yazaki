import os
import time
import shutil
import subprocess
import re
from icalendar import Calendar, vCalAddress
from datetime import datetime, timedelta

# --- CONFIGURATION ---
INCOMING_FOLDER = r"C:\Users\10817991\OneDrive - Yazaki\Pictures\Saved Pictures\incoming"
PROCESSED_FOLDER = r"C:\Users\10817991\OneDrive - Yazaki\Pictures\Saved Pictures\processed"
ERROR_FOLDER = r"C:\Users\10817991\OneDrive - Yazaki\Pictures\Saved Pictures\error"
POLL_INTERVAL = 300  # Time in seconds to wait between checking the incoming folder.

def format_duration(timedelta_obj):
    """Formats a timedelta object into a string like '90 minutes'."""
    total_minutes = int(timedelta_obj.total_seconds() // 60)
    return f"{total_minutes} minutes" if total_minutes >= 0 else "0 minutes"

def get_windchill_user_summary_from_email(email):
    """
    Queries Windchill for a user's full summary string (e.g., 'Charles Beck (Software - Intern)')
    based on their email address.
    """
    if not email:
        return None
        
    command = f'im issues --query="User Profiles: Active" --fields="ID,Summary,Email" | findstr /i "{re.escape(email)}"'
    print(f"    - Searching for Windchill user summary for: {email}")
    
    try:
        result = subprocess.run(command, capture_output=True, text=True, shell=True)
        stdout = result.stdout.strip()

        if result.returncode == 0 and stdout:
            first_line = stdout.splitlines()[0]
            prefix = "User Profile (Active):"
            
            if prefix in first_line:
                summary_start_index = first_line.find(prefix) + len(prefix)
                summary_and_email = first_line[summary_start_index:].strip()
                
                email_pos = summary_and_email.lower().rfind(email.lower())
                
                if email_pos != -1:
                    summary = summary_and_email[:email_pos].strip()
                    print(f"      - Found user summary: {summary}")
                    return summary

        print(f"      - User summary not found in Windchill.")
        return None
            
    except Exception as e:
        print(f"[!] An error occurred while searching for user '{email}': {e}")
        return None

def process_ics_file(filepath):
    """
    Reads an .ics file, cleans the description, builds and executes a command
    to create a Windchill item, and returns True on success or False on failure.
    """
    print(f"--- Processing file: {os.path.basename(filepath)} ---")
    
    try:
        with open(filepath, 'rb') as f:
            cal = Calendar.from_ical(f.read())
        event = next((c for c in cal.walk() if c.name == "VEVENT"), None)
    except Exception as e:
        print(f"[!] Error parsing .ics file: {e}. Moving to error folder.")
        return False

    if not event:
        print("[!] No meeting event (VEVENT) found. Moving to error folder.")
        return False

    # --- Step 1: Extract data from .ics file ---
    organizer_email = None
    organizer = event.get('organizer')
    if organizer:
        organizer_email = re.sub(r'mailto:', '', str(organizer), flags=re.IGNORECASE)
        
    attendee_emails = set()
    attendees = event.get('attendee', [])
    if not isinstance(attendees, list):
        attendees = [attendees]
        
    for attendee in attendees:
        if isinstance(attendee, vCalAddress):
            # Skip attendees that are resources (like meeting rooms)
            if 'RESOURCE' in str(attendee.params.get('CUTYPE', '')).upper():
                print(f"      - Skipping resource: {attendee.params.get('CN', 'Unknown Resource')}")
                continue

            attendee_email = re.sub(r'mailto:', '', str(attendee), flags=re.IGNORECASE)
            if attendee_email.lower() != (organizer_email or '').lower():
                attendee_emails.add(attendee_email)

    # --- Step 2: Look up Windchill user summaries ---
    initiator_summary = None
    if organizer_email:
        print("--- Finding meeting initiator in Windchill ---")
        initiator_summary = get_windchill_user_summary_from_email(organizer_email)

    windchill_attendee_summaries = []
    if attendee_emails:
        print("--- Finding meeting attendees in Windchill ---")
        for email in sorted(list(attendee_emails)):
            summary = get_windchill_user_summary_from_email(email)
            if summary:
                windchill_attendee_summaries.append(summary)
    
    # --- Step 3: Assemble the 'im createissue' command ---
    create_parts = ['im', 'createissue', '--type=Meeting']
    
    title = str(event.get('summary', 'No Title')).replace('"', '\\"')
    create_parts.append(f'--field="Title={title}"')
    
    full_description = str(event.get('description', 'No Agenda Provided'))
    footer_delimiter = '________________________________________________________________________________'
    cleaned_description = full_description.split(footer_delimiter, 1)[0].strip()
    if cleaned_description: # Only add description if it's not empty after cleaning
        sanitized_description = cleaned_description.replace('"', '\\"')
        create_parts.append(f'--field="Description={sanitized_description}"')

    # Only add fields if we have data for them
    if initiator_summary:
        create_parts.append(f'--field="Initiator={initiator_summary}"')
    
    create_parts.append('--field="Topic=Meeting"')

    dtstart = event.get('dtstart').dt
    dtend = event.get('dtend').dt
    if isinstance(dtstart, datetime):
        create_parts.append(f'--field="Scheduled Date={dtstart.strftime("%b %d, %Y")}"')
        create_parts.append(f'--field="Scheduled Time={dtstart.strftime("%I:%M %p")}"') # Use AM/PM for clarity
        if isinstance(dtend, datetime):
            duration = format_duration(dtend - dtstart)
            create_parts.append(f'--field="Scheduled Duration={duration}"')

    if windchill_attendee_summaries:
        attendees_str = ",".join(windchill_attendee_summaries)
        create_parts.append(f'--field="Scheduled Attendees={attendees_str}"')

    # --- Step 4: Execute the command ---
    command_string = " ".join(create_parts)
    print("--- Assembled Windchill Command ---")
    print(f"    - Executing: {command_string}")
    
    try:
        result = subprocess.run(command_string, capture_output=True, text=True, check=True, shell=True)
        
        print("--- Command Output ---")
        if result.stdout:
            print(f"    - STDOUT: {result.stdout.strip()}")

        match = re.search(r"Created issue (\d+)", result.stdout)
        if match:
            item_number = match.group(1)
            print(f"\n[+] Success! Created Meeting item: {item_number}")
        else:
            print("\n[+] Success! Command executed, but the item ID could not be parsed from output.")
            
        return True
    except subprocess.CalledProcessError as e:
        print("[!] FAILED to create item. The command returned an error.")
        if e.stdout:
            print(f"    - STDOUT: {e.stdout.strip()}")
        if e.stderr:
            print(f"    - STDERR: {e.stderr.strip()}")
        return False

def main():
    """Main loop to continuously monitor the incoming folder for new files."""
    print("--- Windchill Meeting Creation Listener starting ---")
    print(f"--- Monitoring folder: {INCOMING_FOLDER} ---")
    
    while True:
        try:
            print(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Checking for new .ics files...")
            files_to_process = [f for f in os.listdir(INCOMING_FOLDER) if f.lower().endswith('.ics')]
            
            if not files_to_process:
                print("...No new files found.")
            else:
                print(f"...Found {len(files_to_process)} new file(s) to process.")
                for filename in files_to_process:
                    source_path = os.path.join(INCOMING_FOLDER, filename)
                    success = process_ics_file(source_path)

                    if success:
                        destination_path = os.path.join(PROCESSED_FOLDER, filename)
                    else:
                        destination_path = os.path.join(ERROR_FOLDER, filename)
                    
                    print(f"--- Moving to: {destination_path} ---")
                    shutil.move(source_path, destination_path)
                    print("-" * 50)

            print(f"--- Sleeping for {POLL_INTERVAL} seconds... ---")
            time.sleep(POLL_INTERVAL)

        except FileNotFoundError:
            print(f"[!!!] CRITICAL ERROR: A folder was not found. Please check paths:")
            print(f"    - Incoming: {INCOMING_FOLDER}, Processed: {PROCESSED_FOLDER}, Error: {ERROR_FOLDER}")
            print("--- Exiting script. ---")
            break
        except Exception as e:
            print(f"[!!!] A critical error occurred in the main loop: {e}")
            print(f"--- Waiting for {POLL_INTERVAL} seconds before retrying... ---")
            time.sleep(POLL_INTERVAL)

if __name__ == "__main__":
    main()