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
    cal_duration = [evt.duration for evt in events]
    cal_category = [evt.categories for evt in events]

    df = pd.DataFrame({'subject': cal_subject,
                       'start': cal_start,
                       'end': cal_end,
                       'duration': cal_duration,
                       'category': cal_category})
    
    df['subject'] = df['subject'].astype(object)
    df['start'] = df['start'].dt.tz_convert(None)
    df['end'] = df['end'].dt.tz_convert(None)
    df['category'] = df['category'].astype(object)

    return(df)

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

def main():
    end = dt.datetime.combine(dt.date.today(), dt.datetime.min.time()) + dt.timedelta(days=14)
    begin = dt.datetime.combine(dt.date.today(), dt.datetime.min.time()) - dt.timedelta(days=7)

    outlook_cal = get_outlook_calendar(begin, end)
    outlook_events = get_outlook_events(outlook_cal)

    google_events = get_google_events(begin, end)

    full_merge = outlook_events.merge(google_events, on=['subject','start','end'], how='outer', indicator=True)
    full_merge.to_excel('last_full_merge.xlsx')

    left_only_merge = outlook_events.merge(google_events, on=['subject','start','end','category'], how='outer', indicator=True).loc[lambda x : x['_merge']=='left_only'].drop('_merge', axis='columns')
    missing_events = left_only_merge.values

    right_only_merge = Outlook_events.merge(google_events, on=['subject','start','end','category'], how='outer', indicator=True).loc[lambda x : x['_merge']=='right_only'].drop('_merge', axis='columns')
    cancelled_events = right_only_merge.values
    
if __name__ == '__main__':
   main()