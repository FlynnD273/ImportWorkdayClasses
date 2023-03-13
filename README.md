# ImportWorkdayClasses
Converts an Excel spreadsheet containing a class schedule downloaded from Workday to an iCalendar file that you can import into any calendar app (for example, Outlook) 

I haven't added much in the way of customization. So far, you can change the number of minutes before the event you want a reminder and the time zone. 

The Python script will read a file called `View_My_Courses.xlsx` in the same directory as the script, and output a file called `Workday Schedule.ics` as the output. It will also print out the classes in the console so you can double-check that it's correct before importing the ics file into your calendar.

## Requirements
You must have Python installed. The required modules can be found in the `requirements.txt` file.
You can install the modules by either running the `install_requirements.bat` script, or by running `
pip install -r requirements.txt` from project directory.

## Running The Script
To run the script, simply run `python importwd.py` from the project directory.

