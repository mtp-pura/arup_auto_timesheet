# arup_auto_timesheet

Converts your outlook calendar into a timesheet using CODES that you define (jobs..xlsx) and use in the appointment title. For example, if you include 'PERSLEAVE' in the title, it will suck that in and classify it as 'Personal Leave' in the final timesheet. 'SKIP' on the otherhand will ignore that appointment and not include it in your timesheet. The codes MUST be defined by each user.

Import_TS_Calendar_Hourly.xlsx - this is the default timesheet for importing. this was pulled from the timesheet systems website


cal_py.py - the is the main application. sorry for not having much documentation... 

  The user must acquire their own SECRET_ID (see Client Credentials) from below [Azure Portal Link](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Overview/appId/73fee778-cbb5-4c82-81cd-13503338d848/isMSAApp~/false) and add it into the cal_py script.  Also read here for [O365 Integration](https://github.com/O365/python-o365?tab=readme-ov-file#authentication). Do not share this ID with others.
  ![image](https://github.com/user-attachments/assets/6067f63d-7142-4622-9680-75ce05667483).

  
jobs_py.xlsx - this is where the CODES (case-sensitive) are defined. I left mine in as an example. but you must replace them with your own. dont use words, use unique letter combinations for each project.


o365_token.txt - need to acquire your o365 token for your personal calendar, but I can't recall exactly how. replace 'INSERT_KEY_HERE' with your token. do not share with others.


powershell_command.txt - once you have it running, use this command to convert the files into an executable. it will only work with your credentials, so don't bother sharing.


# Usage:
- use you outlook calendar to build out your week. add appointments as you do them and/or plan your week ahead.
- if using cal_py.py from the command line - adjust the "answers" (line 50-52) to select the week you want to ingest.
- run the script and categorise any uncategorised appointments.
  - it may require you to AUTH with your MS Authenticator, if you are not logged into your browser
  - it will load a blank page, copy the URL of the blank page back into the command line and Enter.
  - each new run will replace the existing sheet from that period, so you can run it as many times as you like to check your hours for that period.
- it will save the exported timesheet in the folder at ./exported_formatted
- open the timesheet file to quickly check the output and save. (you must do this otherwise the timesheet system will reject it)
- import timesheet into the Infor Expense Manager.
- Validate one last time and submit

# Notes
Week is Monday to Sunday. If you run it on Monday and want to do last week, you will need to change the python script. (Lines 50-52)
ADHOC allows you to pick any week given a Monday start date (Line 63).
