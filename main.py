import os
import pytz
import pandas as pd
import warnings
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from datetime import datetime

SCOPES = ['https://www.googleapis.com/auth/calendar']

def get_calendar_service(verbosity=False):
    """
    Authenticate and return a Google Calendar API service.

    Example:
        service = get_calendar_service()
    """
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('calendar', 'v3', credentials=creds)

    if verbosity:
        print("Authenticated Google Calendar API service.")

    return service


def fetch_events(service, date, verbosity=False):
    """
    Fetch events for a specific date from all calendars.

    Example:
        df = fetch_events(service, '2023-09-13')
    """
    assert service is not None, "Service must be initialized."

    # Convert to the user's local time zone (Jerusalem, GMT+2)
    tz = pytz.timezone('Asia/Jerusalem')
    dt = datetime.fromisoformat(date)
    dt = tz.localize(dt)

    time_min = dt.strftime('%Y-%m-%dT00:00:00+02:00')
    time_max = dt.strftime('%Y-%m-%dT23:59:59+02:00')

    if verbosity:
        print(f"Fetching events for {date}.")

    calendar_list = service.calendarList().list().execute().get('items', [])
    df = pd.DataFrame()

    for calendar in calendar_list:
        calendar_id = calendar['id']
        calendar_name = calendar['summary']

        events_result = service.events().list(
            calendarId=calendar_id, timeMin=time_min,
            timeMax=time_max, singleEvents=True,
            orderBy='startTime').execute()

        events = events_result.get('items', [])

        if not events and verbosity:
            print(f"No events found for calendar: {calendar_name}")

        for event in events:
            event_data = {
                'Calendar Name': calendar_name,
                'Event Summary': event.get('summary', 'Unknown'),
                'Event Start': event['start'].get('dateTime', event['start'].get('date')),
                'Event End': event['end'].get('dateTime', event['end'].get('date')),
            }
            df = df.append(event_data, ignore_index=True)

    if df.empty:
        warnings.warn("No events found for the specified date.")

    return df


def add_event(service, calendar_id, summary, start_time, end_time, description='', location='', verbosity=False):
    """
    Add an event to a specific calendar.

    Example:
        add_event(service, 'primary', 'Meeting', '2023-09-13T10:00:00', '2023-09-13T11:00:00')
    """
    assert service is not None, "Service must be initialized."
    assert calendar_id and summary and start_time and end_time, "Essential event details must be provided."

    event = {
        'summary': summary,
        'location': location,
        'description': description,
        'start': {
            'dateTime': start_time,
            'timeZone': 'Asia/Jerusalem',
        },
        'end': {
            'dateTime': end_time,
            'timeZone': 'Asia/Jerusalem',
        },
    }

    event = service.events().insert(calendarId=calendar_id, body=event).execute()

    if verbosity:
        print(f"Event created: {event.get('htmlLink')}")


def create_calendar(service, summary, verbosity=False):
    """
    Create a new calendar.

    Example:
        create_calendar(service, 'My New Calendar')
    """
    assert service is not None, "Service must be initialized."
    assert summary, "Calendar summary must be provided."

    calendar = {
        'summary': summary,
        'timeZone': 'Asia/Jerusalem'
    }

    created_calendar = service.calendars().insert(body=calendar).execute()

    if verbosity:
        print(f"Created calendar: {created_calendar['summary']}, ID: {created_calendar['id']}")


if __name__ == '__main__':
    service = get_calendar_service(verbosity=True)

    # Example usage
    # create_calendar(service, 'My New Calendar', verbosity=True)
    # add_event(service, 'primary', 'Meeting with John', '2023-09-13T10:00:00', '2023-09-13T11:00:00', verbosity=True)
    df = fetch_events(service, '2023-09-13', verbosity=True)
    print(df)
