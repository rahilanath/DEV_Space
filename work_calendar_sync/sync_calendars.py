import datetime as dt
import pandas as pd
import win32com.client
import os, sys
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

    cal_event_id = []
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
        endStr = event['end'].get('dateTime', event['start'].get('date'))

        try:
            category = event['colorId']

        except:
            category = None

        cal_event_id.append(event['id'])
        cal_subject.append(event['summary'])
        cal_start.append(startStr.replace('T',' ')[:19])
        cal_end.append(endStr.replace('T',' ')[:19])
        cal_category.append(category)

    df = pd.DataFrame({'event_id': cal_event_id,
                       'subject': cal_subject,
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
        elif event[3] == '':
            color = 'None'
        else:
            color = 'None'

        print(f'Creating event: {subject} from {start} to {end} with {color} priority...')

        if color == 'None':
            event_result = service.events().insert(calendarId='primary',
                body={
                    "summary": subject,
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

def delete_google_events(cancelled_event_list):
    service = get_calendar_service()

    for event in cancelled_event_list:
        event_id = event[4]
        subject = event[0]
        start = event[1]
        end = event[2]

        print(f'Deleting event: {subject} from {start} to {end}...')

        cancelled_event = service.events().delete(calendarId='primary', eventId=event_id).execute()

def main():
    end = dt.datetime.combine(dt.date.today(), dt.datetime.min.time()) + dt.timedelta(days=14)
    begin = dt.datetime.combine(dt.date.today(), dt.datetime.min.time()) - dt.timedelta(days=7)

    outlook_cal = get_outlook_calendar(begin, end)
    outlook_events = get_outlook_events(outlook_cal)

    try:
        google_events = get_google_events(begin, end)

    except:
        os.remove('./token.pickle')
        google_events = get_google_events(begin, end)

    merge_dictionary={"left_only":"missing", "right_only":"cancelled","both":"synced"}

    all_events_merged_df = outlook_events.merge(google_events, on=['subject','start','end','category'], how='outer', indicator=True)
    all_events_merged_df['_merge'] = all_events_merged_df['_merge'].map(merge_dictionary)
    all_events_merged_df.to_excel('last_all_events_merged_list.xlsx')

    missing_events_df = all_events_merged_df.loc[lambda x : x['_merge']=='missing'].drop(columns=['_merge','event_id'])
    missing_events_df.to_excel('last_missing_event_list.xlsx')
    missing_events = missing_events_df.values

    cancelled_events_df = all_events_merged_df.loc[lambda x : x['_merge']=='cancelled'].drop(columns=['_merge'])
    cancelled_events_df.to_excel('last_cancelled_event_list.xlsx')
    cancelled_events = cancelled_events_df.values

    create_google_events(missing_events)
    delete_google_events(cancelled_events)

    print('Calendar sync complete...')

if __name__ == '__main__':
   main()