import datetime as dt
import pandas as pd
import win32com.client
from cal_setup import get_calendar_service

def get_outlook_calendar(begin, end):
    outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
    calendar = outlook.getDefaultFolder(9).Items
    calendar.IncludeRecurrences = True
    calendar.Sort('[Start]')
    restriction = "[Start] >= '" + begin.strftime('%m/%d/%Y') + "' AND [END] <= '" + end.strftime('%m/%d/%Y') + "'"
    calendar = calendar.Restrict(restriction)
    return calendar

def get_outlook_events(calendar):
    events = [evt for evt in calendar if evt.end.year == dt.date.today().year]
    #  and 'OOO' not in evt.subject.upper() and 'out of office' not in evt.subject.lower()]
    cal_subject = [evt.subject for evt in events]
    cal_start = [evt.start for evt in events]
    cal_end = [evt.end for evt in events]
    cal_category = [evt.categories for evt in events]

    df = pd.DataFrame({'subject': cal_subject,
                       'start': cal_start,
                       'end': cal_end,
                       'category': cal_category})
    
    df['subject'] = df['subject'].astype(object)
    df['start'] = df['start'].dt.tz_convert(None)
    df['end'] = df['end'].dt.tz_convert(None)
    df['category'] = df['category'].astype(object)
    return df

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

    return(df)

def create_google_events(missing_event_list):
    service = get_calendar_service()

    for event in missing_event_list:
        subject = event[0]
        start = event[1].isoformat()
        end = event[2].isoformat()

        if event[3] == 'Mandatory':
            color = '11'
        elif event[3] == 'Non-Mandatory':
            color = '6'
        elif event[3] == 'Sticky':
            color = '5'
        elif event[3] == 'Reminder':
            color = '3'
        elif event[3] == 'Time-Off':
            color = '2'
        elif event[3] == ' ':
            color = 'None'

        print(f'Creating event: {subject} from {start} to {end} with {color} priority...')

        if color == 'None':
            event_result = service.events().insert(calendarId='primary',
                body={
                    "summary": subject,
                    "colorId": color,
                    "start": {"dateTime": start, "timeZone": 'America/Los_Angeles'},
                    "end": {"dateTime": end, "timeZone": 'America/Los_Angeles'},
                }
            ).execute()
        
        else:
            event_result = service.events().insert(calendarId='primary',
                body={
                    "summary": subject,
                    "colorId": color,
                    "start": {"dateTime": start, "timeZone": 'America/Los_Angeles'},
                    "end": {"dateTime": end, "timeZone": 'America/Los_Angeles'},
                }
            ).execute()

def main():
    end = dt.datetime.combine(dt.date.today(), dt.datetime.min.time()) + dt.timedelta(days=7)
    begin = dt.datetime.combine(dt.date.today(), dt.datetime.min.time()) - dt.timedelta(days=7)

    outlook_cal = get_outlook_calendar(begin, end)
    outlook_events = get_outlook_events(outlook_cal)
    google_events = get_google_events(begin, end)

    merged = outlook_events.merge(google_events, on=['subject','start','end','category'], how='outer', indicator=True).loc[lambda x : x['_merge']=='left_only'].drop('_merge', axis='columns')
    missing_events = merged.values

    merged.to_excel('merged.xlsx')
    create_google_events(missing_events)

    print('Calendar sync complete...')
    
if __name__ == '__main__':
   main()