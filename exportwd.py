from icalendar import Calendar, Event, Alarm
import datetime as dt
import arrow
import pandas as pd
import sys
import uuid

warning = 15 # How many minutes before class to show reminder
timezone = 'US/Eastern' # The time zone to use for the calendar events

def to_dt (arr: arrow.Arrow) -> dt.datetime:
    return dt.datetime(arr.year, arr.month, arr.day, arr.hour, arr.minute, arr.second, tzinfo= arr.tzinfo)

def parse_event (row):
    now = arrow.now()
    event = Event()
    event['uid'] = str(uuid.uuid4())
    event.add('dtstamp', to_dt(arrow.now()))

    event.add('summary', row.course_listing + ' | ' + row.instructional_format)
    event.add('description', row.delivery_mode)

    meeting_pat = row.meeting_patterns.strip().split('|')
    event.add('location', meeting_pat[2])

    times = meeting_pat[1].split(' - ')
    start = dt.datetime.strptime(times[0].strip(), "%I:%M %p")
    start = arrow.get(start, timezone)
    start = start.replace(year= now.year, month= now.month, day= now.day)

    # Convert to UTC time 
    # Some programs don't take into account timezones properly
    start = start.to("UTC")

    end = dt.datetime.strptime(times[1].strip(), "%I:%M %p")
    end = arrow.get(end, timezone)
    end = end.replace(year= now.year, month= now.month, day= now.day)
    end = end.to("UTC")

    workdayToIcal = {'M': 'mo', 'T': 'tu', 'W': 'we', 'R': 'th', 'F': 'fr'}
    weekdays = list(workdayToIcal.values())
    days = []
    for letter in workdayToIcal:
        if letter in meeting_pat[0]:
            days.append(workdayToIcal[letter])

    startDate = arrow.get(row.start_date)
    while startDate.weekday() >= len(weekdays) or weekdays[startDate.weekday()] not in days:
        startDate = startDate.shift(days=1)

    event.add('dtstart', to_dt(startDate.replace(hour=start.hour, minute=start.minute)))
    event.add('dtend', to_dt(startDate.replace(hour=end.hour, minute=end.minute)))
    event.add('rrule', {'freq': 'weekly', 'until': to_dt(arrow.get(row.end_date).shift(days=1)), 'byday': days})

    alarm = Alarm()
    alarm.add('action', 'DISPLAY')
    alarm.add('trigger', dt.timedelta(minutes= warning))
    event.add_component(alarm)

    eventDays = str.join(", ", event['rrule']['BYDAY'])

    print(event['summary'])
    print(f"Given time: {meeting_pat[1].strip()}")
    # print(f"Parsed time: {start} - {end}")
    print(f"Parsed time: {start.to(timezone):hh:mm A} - {end.to(timezone):hh:mm A}")

    print(f"Calendar event:\nSummary: {event['summary']}\nLocation: {event['location']}\nRepeats on {eventDays} until {event['rrule']['UNTIL'] - dt.timedelta(days=1):%A, %B %d %y}\nReminder {warning} minutes before the event")
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
