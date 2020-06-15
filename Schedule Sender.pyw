from __future__ import print_function
import datetime
import calendar
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import smtplib

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']


def main():
    """Google boilerplate interaction for the calendar api"""
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)

    # Call the Calendar API
    # Gets current time in specific iso format
    now = datetime.datetime.utcnow().isoformat() + 'Z'
    # Just adds 7 days to the current time
    end_time = datetime.datetime.now() + datetime.timedelta(days=7)
    # Formats the end time properly
    end_time = end_time.isoformat() + 'Z'
    # Gets all events in next 7 days
    events_result = service.events().list(calendarId='experimacmichigan@gmail.com', timeMin=now, singleEvents=True,
                                          orderBy='startTime', timeMax=end_time).execute()
    events = events_result.get('items', [])

    # Checks for any events in list
    if not events:
        print('No upcoming events found.')
    my_work_events = []
    # Loops through checking if 'brett' is in any of the events, can be changed based on person
    for event in events:
        if 'brett' in event['summary'].lower():
            my_work_events.append(event)
    # Sends events matching persons name to the text function.
    send_work_dates(my_work_events)


def send_work_dates(work_events):
    reformatted_datetimes = []
    # Goes through a loop, reformatting all the information. Changes it to make it more readable for sending.
    for event in work_events:
        (start, end, date) = reformat_datetime(event)
        date = datetime.datetime.strptime(date, '%Y-%m-%d').strftime('%A')
        start = datetime.datetime.strptime(start, '%H:%M').strftime('%I:%M%p')
        end = datetime.datetime.strptime(end, '%H:%M').strftime('%I:%M%p')
        reformatted_datetimes.append((start, end, date))

    # Opens the mail server connection
    smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
    # Says hello to the mail server to check connection
    smtpObj.ehlo()
    # Enables tls
    smtpObj.starttls()
    # Logs into google account. First parameter is eamil, second is password. Omitted for upload
    smtpObj.login(' brett.csotty@gmail.com ', '  ')
    # Reformats the list to put all the dates into the single array index
    for x in range(1, len(reformatted_datetimes)):
        reformatted_datetimes[0] += reformatted_datetimes[x]

    reformatted_datetimes = reformatted_datetimes[0]
    # Sets up the start of the body for the message. the double \n is to indicate message start
    body = "Subject: Work Schedule.\n\n \n"

    detail_index = 1
    # Loops through and formats based upon detail. 1 is start time, 2 is end time, and 3 is day.
    for detail in reformatted_datetimes:
        if detail_index == 3:
            body += detail + "\n"
            detail_index = 1
        elif detail_index == 2:
            body += detail + " "
            detail_index += 1
        else:
            body += detail + "-"
            detail_index += 1
    # Sends the email from the listed email to all the phone numbers. Uses email to text function of carriers
    # For multiple numbers, they must be in list. Filler number put in for example
    send_schedule_status = smtpObj.sendmail(' brett.csotty@gmail.com ',
                                            '1234567890@txt.att.net', body)
    # Checks for message failure, then ends smtp connection
    if send_schedule_status != {}:
        print('There was an error sending the schedule')
    smtpObj.quit()


def reformat_datetime(event):
    # Using string splicing to select certain features of original string.
    start_time = event['start'].get('dateTime')[11:16]
    end_time = event['end'].get('dateTime')[11:16]
    date = event['start'].get('dateTime')[0:10]
    return start_time, end_time, date


if __name__ == '__main__':
    main()
