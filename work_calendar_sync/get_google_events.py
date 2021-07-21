import datetime as dt
import pandas as pd
from cal_setup import get_calendar_service

def main():
    service = get_calendar_service()
    
    now = dt.datetime.now()
    priorNow = now - dt.timedelta(days=7)
    futureNow = now + dt.timedelta(days=7)
    minTime = priorNow.isoformat() + 'Z'
    maxTime = futureNow.isoformat() + 'Z'

    cal_subject = []
    cal_start = []
    cal_end = []

    events_result = service.events().list(
        calendarId='primary', timeMin=minTime,
        timeMax=maxTime, singleEvents=True,
        orderBy='startTime').execute()
    events = events_result.get('items', [])

    if not events:
        print('No events found.')
    for event in events:
        startStr = event['start'].get('dateTime', event['start'].get('date'))
        endStr = event['end'].get('dateTime', event['end'].get('date'))

        cal_subject.append(event['summary'])
        cal_start.append(startStr.replace('T',' ')[:19])
        cal_end.append(endStr.replace('T',' ')[:19])

    df = pd.DataFrame({'subject': cal_subject,
                       'start': cal_start,
                       'end': cal_end})
    
    # df['start'] = df['start'].dt.tz_convert(None)
    # df['end'] = df['end'].dt.tz_convert(None)
    return(df)

if __name__ == '__main__':
    main()