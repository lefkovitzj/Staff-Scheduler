# Staff-Scheduler
Create tentative aquatics facility staff schedules from a given spreadsheet containing availability for the upcoming scheduling period.


This project contains two application methods, but both require the following inputs:
1. Number of staff per shift of each type (Lifeguards, Managers, Pool Attendants).
2. The maximum ideal number of consecutive shifts worked by one staff member (will be scheduled only if necessary).
3. Input file path (launches file explorer).
4. Output file path.

<hr>

The input file for this application must be in the format of "Sample Availability Sheet.xlsx" in order to run correctly. All three notebook tabs must be included, and the format must remain largely the same. You can make the following changes to the :
1. Change the text representing the staff members who are available (currently U1,U2, U3, ...).
2. Change the text representing the day.
3. Change the text representing the shift.
4. Add more staff members below the current ones.
5. Add more days (with shifts) in the same format as the currently included shifts.
6. Change the document filename.

Make sure both the input file and output file (overwritten if it exists) are not open in any other program before running this application.

<hr>

This project requires the source files "utils.py", "wrapper.py", "cmd_app.py" and "gui_app.py" to have full feature access. You could ignore either one of the "..._app.py" files, but then you will not have access to that file's application type (Command Line Interface or Graphical User Interface). 
The Python Library "openpyxl" must be installed. You can do this by running the command "pip install openpyxl". To learn more about openpyxl, <a href="https://openpyxl.readthedocs.io/en/stable/">check out the documentation</a>.

I have only been able to test this application on the Windows operating system, but it should work on any Operating System.

Feel free to reach out to me with ideas, questions, or other comments regarding this project by opening an Issue!
