"""
Run Calendar Sync
Automatically adds your monthly shift schedule to Google Calendar.
"""

import json
import os
import sys
from datetime import datetime, timedelta
import openpyxl
import xlrd
import pandas as pd
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Google Calendar API scope
SCOPES = ['https://www.googleapis.com/auth/calendar']

DEFAULT_DOCTOR_NAME = ""
DEFAULT_CALENDAR_ID = 'primary'
SETTINGS_FILENAME = 'user_settings.json'


def load_user_settings(script_dir):
    """Load persisted user settings (doctor name, calendar ID)."""
    settings_path = os.path.join(script_dir, SETTINGS_FILENAME)
    if os.path.exists(settings_path):
        try:
            with open(settings_path, 'r', encoding='utf-8') as handle:
                data = json.load(handle)
                if isinstance(data, dict):
                    return data
        except (json.JSONDecodeError, OSError):
            pass
    return {}


def save_user_settings(script_dir, settings):
    """Persist user settings to disk."""
    settings_path = os.path.join(script_dir, SETTINGS_FILENAME)
    try:
        with open(settings_path, 'w', encoding='utf-8') as handle:
            json.dump(settings, handle, indent=2)
    except OSError:
        pass


def interactive_select(options, prompt, default_index=0):
    """
    Present a list of options and allow selection via arrow keys (Windows) or numeric entry.
    Returns the selected index.
    """
    if not options:
        return None

    index = default_index if 0 <= default_index < len(options) else 0
    supports_arrows = os.name == 'nt'
    msvcrt = None

    if supports_arrows:
        try:
            import msvcrt as _msvcrt
            msvcrt = _msvcrt
        except ImportError:
            supports_arrows = False

    if supports_arrows and msvcrt:
        print(f"\n{prompt}")
        for idx, option in enumerate(options, 1):
            default_note = " (last used)" if idx - 1 == index else ""
            print(f"  {idx}. {option}{default_note}")
        print("Use Up/Down arrows or type a number. Press Enter to confirm.")

        current_message = f"Current selection: {options[index]}"
        print(current_message, end='', flush=True)
        digits = ""

        while True:
            key = msvcrt.getch()

            if key in (b'\r', b'\n'):
                if digits:
                    try:
                        number = int(digits)
                        if 1 <= number <= len(options):
                            index = number - 1
                    except ValueError:
                        pass
                print()
                return index

            if key in (b'\x00', b'\xe0'):
                arrow = msvcrt.getch()
                if arrow == b'H':  # Up
                    index = (index - 1) % len(options)
                elif arrow == b'P':  # Down
                    index = (index + 1) % len(options)
                digits = ""
                new_message = f"Current selection: {options[index]}"
                blank = ' ' * max(len(current_message), len(new_message) + 4)
                print('\r' + blank, end='', flush=True)
                current_message = new_message
                print('\r' + current_message, end='', flush=True)
                continue

            if key == b'\x08':  # Backspace
                digits = digits[:-1]
                display = f"Number entry: {digits}" if digits else f"Current selection: {options[index]}"
                blank = ' ' * max(len(current_message), len(display) + 4)
                print('\r' + blank, end='', flush=True)
                current_message = display
                print('\r' + current_message, end='', flush=True)
                continue

            if key.isdigit():
                digits += key.decode()
                display = f"Number entry: {digits}"
                blank = ' ' * max(len(current_message), len(display) + 4)
                print('\r' + blank, end='', flush=True)
                current_message = display
                print('\r' + current_message, end='', flush=True)
                continue

    # Fallback to simple numeric selection
    print(f"\n{prompt}")
    for idx, option in enumerate(options, 1):
        default_note = " (last used)" if idx - 1 == index else ""
        print(f"  {idx}. {option}{default_note}")

    while True:
        choice = input(f"Enter number (1-{len(options)}) or press Enter for {options[index]}: ").strip()
        if not choice:
            return index
        if choice.isdigit():
            number = int(choice)
            if 1 <= number <= len(options):
                return number - 1
        print("Invalid selection. Try again.")

def get_catchup_times(group):
    """Return catchup time range based on group number"""
    if 2 <= group <= 5:
        return ('00:00', '02:00')
    elif 6 <= group <= 9:
        return ('02:00', '04:00')
    elif 10 <= group <= 12:
        return ('04:00', '06:00')
    elif 13 <= group <= 18:
        return None  # Down after midnight
    return None

def authenticate_google():
    """Authenticate with Google Calendar API"""
    # Get script directory for credentials
    script_dir = os.path.dirname(os.path.abspath(__file__))
    token_path = os.path.join(script_dir, 'token.json')
    credentials_path = os.path.join(script_dir, 'credentials.json')

    # Remove any stored token to force a fresh login each time
    if os.path.exists(token_path):
        try:
            os.remove(token_path)
        except OSError:
            pass

    if not os.path.exists(credentials_path):
        print("\nERROR: credentials.json file not found!")
        print(f"Please place credentials.json in: {script_dir}")
        input("\nPress Enter to exit...")
        sys.exit(1)

    flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
    return flow.run_local_server(port=0)

def find_schedule_file():
    """Find the XLS/XLSX schedule file in the root directory"""
    # Get root directory (2 levels up from script)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    root_dir = os.path.dirname(os.path.dirname(script_dir))

    possible_files = []
    for file in os.listdir(root_dir):
        if file.lower().endswith(('.xls', '.xlsx')) and not file.startswith('~'):
            possible_files.append(file)

    if len(possible_files) == 0:
        print(f"\nERROR: No Excel file found in {root_dir}!")
        print("Please place your schedule Excel file in the calendar folder.")
        return None
    elif len(possible_files) == 1:
        return os.path.join(root_dir, possible_files[0])
    else:
        print("\nMultiple Excel files found. Please select one:")
        for i, file in enumerate(possible_files, 1):
            print(f"{i}. {file}")

        while True:
            try:
                choice = int(input("\nEnter number: "))
                if 1 <= choice <= len(possible_files):
                    return os.path.join(root_dir, possible_files[choice - 1])
            except ValueError:
                pass
            print("Invalid choice. Try again.")

def read_schedule_data(filename):
    """Read the entire schedule file and return rows as dictionaries."""
    print(f"\nReading schedule from: {filename}")

    try:
        schedule_rows = []

        with open(filename, 'rb') as handle:
            first_bytes = handle.read(100)
            is_html = b'<html' in first_bytes.lower() or b'<hr' in first_bytes.lower()

        if is_html:
            print("Detected HTML format, parsing...")
            df = pd.read_html(filename)[0]

            for _, row in df.iterrows():
                shift_data = row.to_dict()
                session_date = shift_data.get('Session Start Date')
                if isinstance(session_date, str):
                    parsed = False
                    for fmt in ('%m/%d/%Y', '%Y-%m-%d'):
                        try:
                            shift_data['Session Start Date'] = datetime.strptime(session_date, fmt)
                            parsed = True
                            break
                        except ValueError:
                            continue
                    if not parsed:
                        shift_data['Session Start Date'] = session_date
                schedule_rows.append(shift_data)

        elif filename.lower().endswith('.xls'):
            workbook = xlrd.open_workbook(filename)
            sheet = workbook.sheet_by_index(0)
            headers = [sheet.cell_value(0, col) for col in range(sheet.ncols)]

            for row_idx in range(1, sheet.nrows):
                row = [sheet.cell_value(row_idx, col) for col in range(sheet.ncols)]
                if not row or not row[0]:
                    continue

                shift_data = dict(zip(headers, row))
                session_date = shift_data.get('Session Start Date')
                if isinstance(session_date, (int, float)):
                    shift_data['Session Start Date'] = xlrd.xldate_as_datetime(session_date, workbook.datemode)
                schedule_rows.append(shift_data)

        else:
            workbook = openpyxl.load_workbook(filename)
            sheet = workbook.active
            headers = [cell.value for cell in sheet[1]]

            for row in sheet.iter_rows(min_row=2, values_only=True):
                if not row or not row[0]:
                    continue

                shift_data = dict(zip(headers, row))
                schedule_rows.append(shift_data)

        return schedule_rows

    except Exception as error:
        print(f"\nERROR reading Excel file: {error}")
        return []


def extract_unique_names(schedule_rows):
    """Return a sorted list of provider names discovered in the schedule data."""
    seen = {}
    for row in schedule_rows:
        full_name = row.get('Full name')
        if not full_name:
            continue
        name_str = str(full_name).strip()
        if not name_str:
            continue
        lower = name_str.lower()
        if lower not in seen:
            seen[lower] = name_str
    return sorted(seen.values(), key=lambda name: name.lower())


def filter_shifts_by_doctor(schedule_rows, doctor_name):
    """Filter schedule rows for the specified doctor name."""
    doctor_name_lower = doctor_name.lower()
    shifts = []

    for shift_data in schedule_rows:
        full_name = shift_data.get('Full name', '')
        if full_name and doctor_name_lower in str(full_name).lower():
            shifts.append(shift_data)

    print(f"Found {len(shifts)} shifts for {doctor_name}")
    return shifts


def select_calendar(service, previous_calendar_id=None):
    """Prompt the user to pick a calendar they can modify."""
    print("\nFetching calendars you can edit...")
    calendars = []
    page_token = None

    while True:
        result = service.calendarList().list(pageToken=page_token).execute()
        for item in result.get('items', []):
            access = item.get('accessRole')
            if access not in ('owner', 'writer'):
                continue
            calendars.append({
                'id': item['id'],
                'summary': item.get('summaryOverride') or item.get('summary') or item['id'],
                'primary': item.get('primary', False)
            })
        page_token = result.get('nextPageToken')
        if not page_token:
            break

    if not calendars:
        print("\nERROR: No editable calendars available.")
        return None, None

    options = []
    default_index = 0
    for idx, cal in enumerate(calendars):
        label = cal['summary']
        if cal['primary']:
            label += " (primary)"
        options.append(label)
        if previous_calendar_id and cal['id'] == previous_calendar_id:
            default_index = idx

    selected_index = interactive_select(options, "Select the calendar to update:", default_index)
    selected_calendar = calendars[selected_index]
    print(f"Using calendar: {selected_calendar['summary']} ({selected_calendar['id']})")
    return selected_calendar['id'], selected_calendar['summary']


def prompt_manual_name(previous_name=None):
    """Prompt the user to type their name manually."""
    fallback = previous_name or DEFAULT_DOCTOR_NAME
    while True:
        if previous_name:
            prompt = f"Type your name as it appears in the schedule (Enter to keep '{previous_name}'): "
        else:
            prompt = "Type your name exactly as it appears in the schedule: "

        entry = input(prompt).strip()
        if entry:
            return entry
        if previous_name:
            return previous_name
        if fallback:
            print(f"Using default name: {fallback}")
            return fallback
        print("Name is required. Please try again.")


def choose_doctor_name(schedule_rows, previous_name=None):
    """Determine which doctor name to use for this run."""
    names = extract_unique_names(schedule_rows)

    if names:
        options = names.copy()
        default_index = 0

        if previous_name:
            match = next((idx for idx, name in enumerate(options) if name.lower() == previous_name.lower()), None)
            if match is not None:
                default_index = match
            else:
                options.insert(0, previous_name)
                default_index = 0

        options.append("Enter a different name...")
        selection = interactive_select(options, "Select your name:", default_index)

        if selection == len(options) - 1:
            return prompt_manual_name(previous_name)
        return options[selection]

    print("\nCould not automatically detect any provider names in this file.")
    return prompt_manual_name(previous_name)

def create_calendar_event(service, calendar_id, event_data):
    """Create a single event in Google Calendar"""
    try:
        event = service.events().insert(calendarId=calendar_id, body=event_data).execute()
        return event
    except HttpError as error:
        print(f"Error creating event: {error}")
        return None

def format_datetime(date, time_str):
    """Convert date and time string to RFC3339 format for Google Calendar"""
    # Date comes as datetime object from Excel
    if isinstance(date, datetime):
        date_str = date.strftime('%Y-%m-%d')
    else:
        date_str = str(date)

    # Combine date and time
    dt_str = f"{date_str}T{time_str}:00"
    return dt_str

def clean_value(value, default=''):
    """Convert NaN or None values to default"""
    if pd.isna(value):
        return default
    return str(value)

def delete_existing_sync_events(service, calendar_id, start_date=None, end_date=None, fallback_days=60):
    """Delete existing Run Calendar events from the calendar."""
    print("\nSearching for existing Run Calendar events to delete...")

    from datetime import timezone

    now_utc = datetime.now(timezone.utc)

    if start_date:
        start_utc = datetime.combine(start_date.date(), datetime.min.time()).replace(tzinfo=timezone.utc)
    else:
        start_utc = now_utc

    if end_date:
        end_utc = datetime.combine(end_date.date(), datetime.max.time()).replace(tzinfo=timezone.utc)
    else:
        end_utc = now_utc + timedelta(days=fallback_days)

    time_min_dt = max(now_utc, start_utc)

    if end_utc <= time_min_dt:
        end_utc = time_min_dt + timedelta(days=1)

    time_min = time_min_dt.isoformat().replace('+00:00', 'Z')
    time_max = end_utc.isoformat().replace('+00:00', 'Z')

    try:
        # Search for events with "Run Calendar" or "Catchup" in the title
        events_result = service.events().list(
            calendarId=calendar_id,
            timeMin=time_min,
            timeMax=time_max,
            singleEvents=True,
            orderBy='startTime'
        ).execute()

        events = events_result.get('items', [])
        deleted_count = 0

        for event in events:
            summary = event.get('summary', '')
            if 'Run Calendar' in summary or 'Catchup' in summary:
                try:
                    service.events().delete(
                        calendarId=calendar_id,
                        eventId=event['id']
                    ).execute()
                    deleted_count += 1
                except HttpError as error:
                    print(f"Error deleting event: {error}")

        print(f"Deleted {deleted_count} existing Run Calendar events")
        return deleted_count

    except HttpError as error:
        print(f"Error searching for events: {error}")
        return 0

def create_events_from_shifts(service, calendar_id, shifts):
    """Create Google Calendar events from shift data"""
    events_created = 0
    
    for shift in shifts:
        session_date = shift['Session Start Date']
        shift_type = shift['Half or Full']
        detail = clean_value(shift.get('Detail', ''))
        group = shift['Group']
        location = clean_value(shift.get('Med Center', ''), default='Home')
        
        events_to_create = []
        location_suffix = f' - {location}' if location else ''

        if shift_type == 'Full':
            # Full shift consists of front and back portions with catchup blocks in between
            next_day = session_date + timedelta(days=1)

            # Front portion: 6:30 PM to midnight
            events_to_create.append({
                'summary': f'Run Calendar Shift (Front){location_suffix}',
                'location': location,
                'description': f'Full shift - Front portion\nGroup: {group}',
                'start': {'dateTime': format_datetime(session_date, '18:30'), 'timeZone': 'America/Los_Angeles'},
                'end': {'dateTime': format_datetime(next_day, '00:00'), 'timeZone': 'America/Los_Angeles'},
                'colorId': '9',  # Blue
                'reminders': {
                    'useDefault': False,
                    'overrides': [
                        {'method': 'popup', 'minutes': 1440},  # 1 day before
                        {'method': 'popup', 'minutes': 10}
                    ]
                }
            })
            
            # Back half catchup and shift based on group
            catchup = get_catchup_times(group)
            if catchup:
                # Back portion: midnight to catchup start - only if catchup doesn't start at midnight
                if catchup[0] != '00:00':
                    events_to_create.append({
                        'summary': f'Run Calendar Shift (Back){location_suffix}',
                        'location': location,
                        'description': f'Full shift - Back portion\nGroup: {group}',
                        'start': {'dateTime': format_datetime(next_day, '00:00'), 'timeZone': 'America/Los_Angeles'},
                        'end': {'dateTime': format_datetime(next_day, catchup[0]), 'timeZone': 'America/Los_Angeles'},
                        'colorId': '9',  # Blue
                        'reminders': {'useDefault': False, 'overrides': []}
                    })

                # Catchup period
                events_to_create.append({
                    'summary': f'Catchup Time{location_suffix}',
                    'location': location,
                    'description': f'Back half catchup (Group {group})',
                    'start': {'dateTime': format_datetime(next_day, catchup[0]), 'timeZone': 'America/Los_Angeles'},
                    'end': {'dateTime': format_datetime(next_day, catchup[1]), 'timeZone': 'America/Los_Angeles'},
                    'colorId': '10',  # Green
                    'reminders': {'useDefault': False, 'overrides': []}
                })

                # Back portion: after catchup ends until 7am
                events_to_create.append({
                    'summary': f'Run Calendar Shift (Back){location_suffix}',
                    'location': location,
                    'description': f'Full shift - Back portion\nGroup: {group}',
                    'start': {'dateTime': format_datetime(next_day, catchup[1]), 'timeZone': 'America/Los_Angeles'},
                    'end': {'dateTime': format_datetime(next_day, '07:00'), 'timeZone': 'America/Los_Angeles'},
                    'colorId': '9',  # Blue
                    'reminders': {'useDefault': False, 'overrides': []}
                })
            else:
                # Groups 12-17: Down after midnight, no back catchup, no back shift
                pass
        
        elif detail == 'First':
            # Front half: 6:30 PM to midnight, then catchup midnight to 1am
            next_day = session_date + timedelta(days=1)


            events_to_create.append({
                'summary': f'Run Calendar Front Half{location_suffix}',
                'location': location,
                'description': f'Front half shift\nGroup: {group}',
                'start': {'dateTime': format_datetime(session_date, '18:30'), 'timeZone': 'America/Los_Angeles'},
                'end': {'dateTime': format_datetime(next_day, '00:00'), 'timeZone': 'America/Los_Angeles'},
                'colorId': '9',  # Blue
                'reminders': {
                    'useDefault': False,
                    'overrides': [
                        {'method': 'popup', 'minutes': 1440},  # 1 day before
                        {'method': 'popup', 'minutes': 10}
                    ]
                }
            })
            
            # Catchup: midnight to 1am
            events_to_create.append({
                'summary': f'Catchup Time{location_suffix}',
                'location': location,
                'description': f'Front half catchup\nGroup: {group}',
                'start': {'dateTime': format_datetime(next_day, '00:00'), 'timeZone': 'America/Los_Angeles'},
                'end': {'dateTime': format_datetime(next_day, '01:00'), 'timeZone': 'America/Los_Angeles'},
                'colorId': '10',  # Green
                'reminders': {'useDefault': False, 'overrides': []}
            })
        
        elif detail == 'Second':
            # Back half: shift is actually next day (overnight/graveyard)
            next_day = session_date + timedelta(days=1)
            catchup = get_catchup_times(group)

            if catchup:
                # Shift before catchup (midnight to catchup start) - only if catchup doesn't start at midnight
                if catchup[0] != '00:00':
                    events_to_create.append({
                        'summary': f'Run Calendar Back Half{location_suffix}',
                        'location': location,
                        'description': f'Back half shift\nGroup: {group}',
                        'start': {'dateTime': format_datetime(next_day, '00:00'), 'timeZone': 'America/Los_Angeles'},
                        'end': {'dateTime': format_datetime(next_day, catchup[0]), 'timeZone': 'America/Los_Angeles'},
                        'colorId': '9',  # Blue
                        'reminders': {
                            'useDefault': False,
                            'overrides': [
                                {'method': 'popup', 'minutes': 1440},  # 1 day before
                                {'method': 'popup', 'minutes': 10}
                            ]
                        }
                    })
                else:
                    # If catchup starts at midnight, add reminders to catchup event (it's the first event)
                    events_to_create.append({
                        'summary': f'Catchup Time{location_suffix}',
                        'location': location,
                        'description': f'Back half catchup (Group {group})',
                        'start': {'dateTime': format_datetime(next_day, catchup[0]), 'timeZone': 'America/Los_Angeles'},
                        'end': {'dateTime': format_datetime(next_day, catchup[1]), 'timeZone': 'America/Los_Angeles'},
                        'colorId': '10',  # Green
                        'reminders': {
                            'useDefault': False,
                            'overrides': [
                                {'method': 'popup', 'minutes': 1440},  # 1 day before
                                {'method': 'popup', 'minutes': 10}
                            ]
                        }
                    })
                    # Skip adding the second catchup event below since we already added it
                    catchup_added = True

                # Catchup period (only if not already added above)
                if catchup[0] != '00:00':
                    events_to_create.append({
                        'summary': f'Catchup Time{location_suffix}',
                        'location': location,
                        'description': f'Back half catchup (Group {group})',
                        'start': {'dateTime': format_datetime(next_day, catchup[0]), 'timeZone': 'America/Los_Angeles'},
                        'end': {'dateTime': format_datetime(next_day, catchup[1]), 'timeZone': 'America/Los_Angeles'},
                        'colorId': '10',  # Green
                        'reminders': {'useDefault': False, 'overrides': []}
                    })

                # Shift after catchup
                events_to_create.append({
                    'summary': f'Run Calendar Back Half{location_suffix}',
                    'location': location,
                    'description': f'Back half shift\nGroup: {group}',
                    'start': {'dateTime': format_datetime(next_day, catchup[1]), 'timeZone': 'America/Los_Angeles'},
                    'end': {'dateTime': format_datetime(next_day, '07:00'), 'timeZone': 'America/Los_Angeles'},
                    'colorId': '9',  # Blue
                    'reminders': {'useDefault': False, 'overrides': []}
                })
            else:
                # Groups 12-17: Down after midnight, no shift
                pass
        
        # Create all events for this shift
        for event_data in events_to_create:
            if create_calendar_event(service, calendar_id, event_data):
                events_created += 1
    
    return events_created

def main():
    print("=" * 60)
    print("RUN CALENDAR SCHEDULE TO GOOGLE CALENDAR")
    print("=" * 60)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    settings = load_user_settings(script_dir)
    last_doctor_name = settings.get('doctor_name', DEFAULT_DOCTOR_NAME)
    last_calendar_id = settings.get('calendar_id', DEFAULT_CALENDAR_ID)

    # Find schedule file
    schedule_file = find_schedule_file()
    if not schedule_file:
        input("\nPress Enter to exit...")
        return
    
    schedule_rows = read_schedule_data(schedule_file)
    if not schedule_rows:
        print("\nNo recognizable data found in the selected schedule file.")
        input("\nPress Enter to exit...")
        return

    doctor_name = choose_doctor_name(schedule_rows, previous_name=last_doctor_name)
    shifts = filter_shifts_by_doctor(schedule_rows, doctor_name)

    while not shifts:
        print(f"\nNo shifts found for '{doctor_name}'. Please choose another name.")
        doctor_name = choose_doctor_name(schedule_rows, previous_name=doctor_name)
        shifts = filter_shifts_by_doctor(schedule_rows, doctor_name)

    settings['doctor_name'] = doctor_name
    save_user_settings(script_dir, settings)
    
    # Show summary
    print(f"\nLooking for shifts assigned to: {doctor_name}")
    print(f"\nFound shifts:")
    print(f"  - Total shifts: {len(shifts)}")

    shift_dates = [shift['Session Start Date'] for shift in shifts if shift.get('Session Start Date')]
    if not shift_dates:
        print("\nERROR: Unable to determine shift dates from the schedule.")
        input("\nPress Enter to exit...")
        return

    earliest_shift = min(shift_dates)
    latest_shift = max(shift_dates) + timedelta(days=1)

    shift_types = {}
    for shift in shifts:
        key = f"{shift['Half or Full']} - {shift.get('Detail', 'N/A')}"
        shift_types[key] = shift_types.get(key, 0) + 1
    
    for shift_type, count in shift_types.items():
        print(f"  - {shift_type}: {count}")
    
    # Confirm before proceeding
    print("\nThis will DELETE existing Run Calendar events and add new ones to your Google Calendar.")
    response = input("Continue? (yes/no): ").strip().lower()
    if response not in ['yes', 'y']:
        print("Cancelled.")
        input("\nPress Enter to exit...")
        return
    
    # Authenticate with Google
    print("\nAuthenticating with Google Calendar...")
    try:
        creds = authenticate_google()
        service = build('calendar', 'v3', credentials=creds)
        print("Authentication successful.")
    except Exception as e:
        print(f"\nERROR during authentication: {e}")
        input("\nPress Enter to exit...")
        return

    # Select calendar
    calendar_id, calendar_label = select_calendar(service, last_calendar_id)
    if not calendar_id:
        input("\nPress Enter to exit...")
        return

    settings['calendar_id'] = calendar_id
    settings['calendar_label'] = calendar_label
    save_user_settings(script_dir, settings)

    # Delete existing Run Calendar events
    delete_existing_sync_events(
        service,
        calendar_id=calendar_id,
        start_date=earliest_shift,
        end_date=latest_shift
    )

    # Create events
    print("\nCreating calendar events...")
    events_created = create_events_from_shifts(service, calendar_id, shifts)
    
    print("\n" + "=" * 60)
    calendar_display = calendar_label or calendar_id
    print(f"SUCCESS! Created {events_created} calendar events in '{calendar_display}'")
    print("=" * 60)
    print("\nCheck your Google Calendar to see your shifts!")
    
    input("\nPress Enter to exit...")

if __name__ == '__main__':
    main()
