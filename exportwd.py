from icalendar import Calendar, Event, Alarm
import pytz
import datetime
from datetime import datetime as dutil
import pandas as pd
import sys
import uuid

warning = -15 # How many minutes before class to show reminder
timezone = 'US/Eastern'

def parse_event (index):
    event = Event()
    event['uid'] = str(uuid.uuid4())
    event.add('dtstamp', dutil.now())

    event.add('summary', dp['Course Listing'][index] + ' | ' + dp['Instructional Format'][index])
    event.add('description', dp['Delivery Mode'][index])

    meeting_pat = dp['Meeting Patterns'][index].strip().split('|')
    event.add('location', meeting_pat[2])

    times = meeting_pat[1].split(' - ')
    start = times[0].split(':')
    end = times[1].split(':')
    startTime = int(start[0])
    startMinutes = int(start[1][0:2])
    endTime = int(end[0])
    endMinutes = int(end[1][0:2])

    if 'PM' in start[1] and int(start[0]) != 12:
        startTime += 12
    if 'PM' in end[1] and int(end[0]) != 12:
        endTime += 12

    workdayToIcal = {'M': 'mo', 'T': 'tu', 'W': 'we', 'R': 'th', 'F': 'fr'}
    weekdays = list(workdayToIcal.values())
    days = []
    for letter in workdayToIcal:
        if letter in meeting_pat[0]:
            days.append(workdayToIcal[letter])

    startDate = dp['Start Date'][index].date()
    while startDate.weekday() >= len(weekdays) or weekdays[startDate.weekday()] not in days:
        startDate = datetime.date(startDate.year, startDate.month, startDate.day + 1)

    print(f"{dp['Course Listing'][index]} | {dp['Instructional Format'][index]}")
    print(meeting_pat[1].strip())
    print(f"{startTime:02d}:{startMinutes:02d} - {endTime:02d}:{endMinutes:02d}")
    print()

    event.add('dtstart', dutil.combine(startDate, datetime.time(startTime, startMinutes, tzinfo=pytz.timezone(timezone))))
    event.add('dtend', dutil.combine(startDate, datetime.time(endTime, endMinutes, tzinfo=pytz.timezone(timezone))))
    event.add('rrule', {'freq': 'weekly', 'until': dp['End Date'][index].date() + datetime.timedelta(days=1), 'byday': days})
    alarm = Alarm()
    alarm.add('action', 'DISPLAY')
    alarm.add('trigger', datetime.timedelta(minutes = -warning))
    event.add_component(alarm)
    return event

path = 'View_My_Courses.xlsx'
if len(sys.argv) > 1:
    path = sys.argv[1]

dp = pd.read_excel(path, sheet_name=0, skiprows=[0, 1])

# RRULE:FREQ=WEEKLY;COUNT=30;BYDAY=MO,TU,TH,FR

# 'Course Listing' + ' | ' + 'Instructional Format' goes into title
# 'Meeting Patterns'.split('|') [0] goes to BYDAY, [1] goes to start and end times, [2] goes to location
# 'Start Date' goes to series start date
# 'End Date' goes to series end date

cal = Calendar()
cal.add('prodid', '-//My calendar product//mxm.dk//')
cal.add('version', '2.0')

for index in range(len(dp['Course Listing'])):
    cal.add_component(parse_event(index))

f = open('Workday Schedule.ics', 'wb')
f.write(cal.to_ical())
f.close()