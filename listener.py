import os
import time
import shutil
import subprocess
from icalendar import Calendar
from datetime import datetime, timedelta
import re  # Added to parse the item number from CLI output

# --- CONFIGURATION ---
# Define the core directories for the file processing workflow.
INCOMING_FOLDER = r"C:\Users\10817991\OneDrive - Yazaki\Pictures\Saved Pictures\incoming"    # Monitored for new .ics files.
PROCESSED_FOLDER = r"C:\Users\10817991\OneDrive - Yazaki\Pictures\Saved Pictures\processed"   # For successfully processed files.
ERROR_FOLDER = r"C:\Users\10817991\OneDrive - Yazaki\Pictures\Saved Pictures\error"           # For files that failed processing.
POLL_INTERVAL = 300  # Time in seconds to wait between checking the incoming folder.

def format_duration(timedelta_obj):
    """Formats a timedelta object into a string like '90 minutes'."""
    total_minutes = int(timedelta_obj.total_seconds() // 60)
    if total_minutes < 0:
        return "0 minutes"
    return f"{total_minutes} minutes"

def process_ics_file(filepath):
    """
    Reads an .ics file, builds a command to create a Windchill item, executes it,
    and returns True on success or False on failure.
    """
    print(f"--- Processing file: {os.path.basename(filepath)} ---")
    
    try:
        # Read and parse the file to extract the main VEVENT (the meeting data).
        with open(filepath, 'rb') as f:
            cal = Calendar.from_ical(f.read())
        event = next((c for c in cal.walk() if c.name == "VEVENT"), None)
    except Exception as e:
        print(f"[!] Error parsing .ics file: {e}. Moving to error folder.")
        return False

    if not event:
        print("[!] No meeting event (VEVENT) found. Moving to error folder.")
        return False

    # --- Assemble the 'im createissue' command ---
    create_parts = ['im', 'createissue', '--type=Meeting']
    
    # Sanitize text from the .ics file to prevent breaking the command string.
    title = str(event.get('summary', 'No Title'))
    sanitized_title = title.replace('"', '\\"')
    create_parts.append(f'--field="Title={sanitized_title}"')
    
    description = str(event.get('description', 'No Agenda Provided'))
    sanitized_description = description.replace('"', '\\"')
    create_parts.append(f'--field="Description={sanitized_description}"')

    # Extract and format other meeting details for the command's fields.
    organizer_email = str(event.get('organizer', '')).replace('MAILTO:', '')
    create_parts.append(f'--field="Initiator={organizer_email}"')
    
    create_parts.append('--field="Topic=Meeting"')

    dtstart = event.get('dtstart').dt
    dtend = event.get('dtend').dt
    if isinstance(dtstart, datetime):
        create_parts.append(f'--field="Scheduled Date={dtstart.strftime("%b %d, %Y")}"')
        create_parts.append(f'--field="Scheduled Time={dtstart.strftime("%H:%M:%S")}"')
        if isinstance(dtend, datetime):
            duration = format_duration(dtend - dtstart)
            create_parts.append(f'--field="Scheduled Duration={duration}"')

    # --- Execute the command and handle the outcome ---
    command_string = " ".join(create_parts)
    print(f"    - Executing: {command_string}")
    try:
        # Run the command, capture output, and auto-check for execution errors.
        result = subprocess.run(command_string, capture_output=True, text=True, check=True, shell=True)
        stdout = result.stdout.strip()
        print(f"[+] Success! Item created.")
        print(stdout)
        # Parse the item number from the CLI output
        match = re.search(r"\b(\d+)\b", stdout)
        if match:
            item_number = match.group(1)
            print(f"Meeting item number: {item_number}")
        return True
    except subprocess.CalledProcessError as e:
        # If the command returns a non-zero exit code, it failed.
        print(f"[!] FAILED to create item.")
        print(f"    - STDERR: {e.stderr.strip()}") # Log the error message from the command.
        return False

def main():
    """Main loop to continuously monitor the incoming folder for new files."""
    print("--- Windchill Meeting Creation Listener starting ---")
    print(f"--- Monitoring folder: {INCOMING_FOLDER} ---")
    
    while True:
        try:
            print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Checking for new .ics files...")
            # Find all files in the directory with a .ics extension.
            files_to_process = [f for f in os.listdir(INCOMING_FOLDER) if f.lower().endswith('.ics')]
            
            if not files_to_process:
                print("...No new files found.")
            else:
                print(f"...Found {len(files_to_process)} new file(s) to process.")
                for filename in files_to_process:
                    source_path = os.path.join(INCOMING_FOLDER, filename)
                    
                    # Process the file and get a success/failure result.
                    success = process_ics_file(source_path)

                    # Move the original file to the appropriate folder based on the outcome.
                    if success:
                        destination_path = os.path.join(PROCESSED_FOLDER, filename)
                    else:
                        destination_path = os.path.join(ERROR_FOLDER, filename)
                    
                    print(f"--- Moving to: {destination_path} ---")
                    shutil.move(source_path, destination_path)
                    print("-" * 50)

            # Wait before the next check.
            print(f"--- Sleeping for {POLL_INTERVAL} seconds... ---")
            time.sleep(POLL_INTERVAL)

        except FileNotFoundError:
            # Handle the critical error where a configured folder is missing.
            print(f"[!!!] CRITICAL ERROR: A folder was not found. Please check paths:")
            print(f"    - Incoming: {INCOMING_FOLDER}, Processed: {PROCESSED_FOLDER}, Error: {ERROR_FOLDER}")
            print("--- Exiting script. ---")
            break # Exit the loop to stop the script.
        except Exception as e:
            # Catch any other unexpected errors to prevent the script from crashing.
            print(f"[!!!] A critical error occurred in the main loop: {e}")
            print(f"--- Waiting for {POLL_INTERVAL} seconds before retrying... ---")
            time.sleep(POLL_INTERVAL)

# Standard entry point to start the script's main loop.
if __name__ == "__main__":
    main()
