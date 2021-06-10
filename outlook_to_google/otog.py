import datetime as dt
import win32com.client

outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')

for test in outlook:
    print()

# def get_calendar(begin,end):
#     outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
#     calendar = outlook.getDefaultFolder(9).Items
#     calendar.IncludeRecurrences = True
#     calendar.Sort('[Start]')

#     restriction = "[Start] >= '" + begin.strftime('%m/%d/%Y') + "' AND [END] <= '" + end.strftime('%m/%d/%Y') + "'"
#     calendar = calendar.Restrict(restriction)
#     return calendar

# cal = get_calendar(dt.datetime(2021,6,6), dt.datetime(2021,6,12))

# for meeting in cal:
#     print(meeting.subject)
#     print(meeting.username)