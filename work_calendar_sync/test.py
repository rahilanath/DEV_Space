import datetime as dt
import pandas as pd
import win32com.client
from cal_setup import get_calendar_service

def get_google_events(begin, end):
    service = get_calendar_service()
    
    minTime = begin.isoformat() + 'Z'
    maxTime = end.isoformat() + 'Z'

    cal_subject = []
    cal_start = []
    cal_end = []
    cal_category = []

    events_result = service.events().list(
        calendarId='primary', timeMin=minTime,
        timeMax=maxTime, singleEvents=True,
        orderBy='startTime').execute()
    events = events_result.get('items', [])

    for event in events:
        if event['start'].get('dateTime') is None:
            startStr = event['start'].get('dateTime', event['start'].get('date'))
            endStr = event['end'].get('dateTime', event['start'].get('date'))

        else:
            startStr = event['start'].get('dateTime', event['start'].get('date'))
            endStr = event['end'].get('dateTime', event['end'].get('date'))
        
        try:
            category = event['colorId']

        except:
            category = None

        cal_subject.append(event['summary'])
        cal_start.append(startStr.replace('T',' ')[:19])
        cal_end.append(endStr.replace('T',' ')[:19])
        cal_category.append(category)

    df = pd.DataFrame({'subject': cal_subject,
                       'start': cal_start,
                       'end': cal_end,
                       'category': cal_category})

    df['start'] = pd.to_datetime(df['start'])
    df['end'] = pd.to_datetime(df['end'])
    df['category'] = df['category'].replace('6','Non-Mandatory')
    df['category'] = df['category'].replace('11','Mandatory')
    df['category'] = df['category'].replace('5','Sticky')
    df['category'] = df['category'].replace('3','Reminder')
    df['category'] = df['category'].replace('2','Time-Off')
    df['category'] = df['category'].replace('None','')

    print(df)

end = dt.datetime.combine(dt.date.today(), dt.datetime.min.time()) + dt.timedelta(days=14)
begin = dt.datetime.combine(dt.date.today(), dt.datetime.min.time()) - dt.timedelta(days=7)

get_google_events(begin, end)