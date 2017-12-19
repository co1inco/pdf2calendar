

from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

import datetime

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/calendar-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/calendar'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Calendar API Python Quickstart'


class gCalendar():
    def __init__(self):
        credentials = self.get_credentials()
        
        http = credentials.authorize(httplib2.Http())
        self.service = discovery.build('calendar', 'v3', http=http)

        
    def get_credentials(self):
        """Gets valid user credentials from storage.

        If nothing has been stored, or if the stored credentials are invalid,
        the OAuth2 flow is completed to obtain the new credentials.

        Returns:
        Credentials, the obtained credential.
        """
    
        home_dir = os.path.expanduser('~')
        credential_dir = os.path.join(home_dir, '.credentials')
        if not os.path.exists(credential_dir):
            os.makedirs(credential_dir)
        credential_path = os.path.join(credential_dir,
                                       'calendar-python-quickstart.json')

        store = Storage(credential_path)
        credentials = store.get()
        if not credentials or credentials.invalid:
            flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
            flow.user_agent = APPLICATION_NAME
            if flags:
                credentials = tools.run_flow(flow, store, flags)
            else: # Needed only for compatibility with Python 2.6
                credentials = tools.run(flow, store)
            print('Storing credentials to ' + credential_path)
        return credentials


    def createEvent(self, calendarID, dateTimeStart, dateTimeEnd, location="Unknown", eventName="Unknown"):
        # 2017-05-28T17:11:00+01:00   2017-05-28T17:11:00.000000Z
        startTime   = dateTimeStart + ":00+01:00"
        endTime     = dateTimeEnd + ":00+01:00"
    
        event = {
            "start": { "dateTime": startTime},
            "end" : { "dateTime": endTime},
            "summary" : eventName,
            "location" : location
            }
    
        out = self.service.events().insert(calendarId=calendarID, body=event).execute()


    def getCalendarList(self):
        calendars = self.service.calendarList().list().execute()
        calendarList = calendars['items']
        return calendarList
        

if __name__ == '__main__':


    now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
    now = now[:3] + '6' + now[4:]


#    for i in calendarList:
#        print(i['summary'])

    calendarId = 'i7qulknrv641gpshdv5p4344r8@group.calendar.google.com'

    i = gCalendar()



