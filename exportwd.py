from icalendar import Calendar, Event, Alarm
import pytz
import datetime
from datetime import datetime as dt
import pandas as pd
import sys
import uuid

warning = 15 # How many minutes before class to show reminder
timezone = 'US/Eastern' # The time zone to use for the calendar events

def parse_event (row):
    event = Event()
    event['uid'] = str(uuid.uuid4())
    event.add('dtstamp', dt.now())

    event.add('summary', row.course_listing + ' | ' + row.instructional_format)
    event.add('description', row.delivery_mode)

    meeting_pat = row.meeting_patterns.strip().split('|')
    event.add('location', meeting_pat[2])

    times = meeting_pat[1].split(' - ')
    start = dt.strptime(times[0].strip(), "%I:%M %p")
    start.replace(tzinfo=pytz.timezone(timezone))

    end = dt.strptime(times[1].strip(), "%I:%M %p")
    end.replace(tzinfo=pytz.timezone(timezone))

    workdayToIcal = {'M': 'mo', 'T': 'tu', 'W': 'we', 'R': 'th', 'F': 'fr'}
    weekdays = list(workdayToIcal.values())
    days = []
    for letter in workdayToIcal:
        if letter in meeting_pat[0]:
            days.append(workdayToIcal[letter])

    startDate = row.start_date.date()
    while startDate.weekday() >= len(weekdays) or weekdays[startDate.weekday()] not in days:
        startDate = datetime.date(startDate.year, startDate.month, startDate.day + 1)

    print(event['summary'])
    print(f"Given time: {meeting_pat[1].strip()}")
    print(f"Parsed time: {start:%I:%M %p} - {end:%I:%M %p}")

    event.add('dtstart', dt.combine(startDate, start.timetz()))
    event.add('dtend', dt.combine(startDate, end.timetz()))
    event.add('rrule', {'freq': 'weekly', 'until': row.end_date.date() + datetime.timedelta(days=1), 'byday': days})

    alarm = Alarm()
    alarm.add('action', 'DISPLAY')
    alarm.add('trigger', datetime.timedelta(minutes = warning))
    alarm.add('description', "display a notification for the event")
    event.add_component(alarm)

    eventDays = str.join(", ", event['rrule']['BYDAY'])
    print(f"Calendar event:\nSummary: {event['summary']}\nLocation: {event['location']}\nRepeats on {eventDays} until {event['rrule']['UNTIL'] - datetime.timedelta(days=1):%A, %B %d %y}\nReminder {warning} minutes before the event")
    print()
    return event

path = 'View_My_Courses.xlsx'
if len(sys.argv) > 1:
    path = sys.argv[1]

excelData = pd.read_excel(path, sheet_name=0, skiprows=[0, 1])

# Replace names with spaces to names with underscores and convert to lowercase
excelData.columns = [c.lower().replace(' ', '_') for c in excelData.columns]

# RRULE:FREQ=WEEKLY;COUNT=30;BYDAY=MO,TU,TH,FR

# 'Course Listing' + ' | ' + 'Instructional Format' goes into title
# 'Meeting Patterns'.split('|') [0] goes to BYDAY, [1] goes to start and end times, [2] goes to location
# 'Start Date' goes to series start date
# 'End Date' goes to series end date

cal = Calendar()
cal.add('prodid', '-//Workday Class Schedule//https://github.com/FlynnD273/ImportWorkdayClasses//')
cal.add('version', '2.0')

for row in excelData.itertuples():
    cal.add_component(parse_event(row))

f = open('Workday Schedule.ics', 'wb')
f.write(cal.to_ical())
f.close()
