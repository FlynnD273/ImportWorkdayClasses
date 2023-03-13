# ImportWorkdayClasses
Converts an Excel spreadsheet containing a class schedule downloaded from Workday to an iCalendar file that you can import into any calendar app (for example, Outlook) 

I haven't added much in the way of customization. So far, you can change the number of minutes before the event you want a reminder and the time zone. 

The Python script will read a file called "View_My_Courses.xlsx" in the same directory as the script, and output a file called "Workday Schedule.ics" as the output. It will also print out the classes in the console so you can double-check that it's correct before importing the ics file into your calendar.

