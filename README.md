# arup_auto_timesheet

Converts your outlook calendar into a timesheet using CODES that you define (jobs..xlsx) and use in the appointment title. For example, if you include 'PERSLEAVE' in the title, it will suck that in and classify it as 'Personal Leave' in the final timesheet. 'SKIP' on the otherhand will ignore that appointment and not include it in your timesheet. The codes MUST be defined by each user.

```
.
├── Import_TS_Calendar_Hourly.xlsx  (Default timesheet for importing, pulled from the timesheet system website)
├── cal_py.py                      (Main application script)
│   └── Documentation:
│       ├── User must acquire their own SECRET_ID (see Client Credentials in Azure) and add it into this script.
│       ├── Refer to [O365 Integration](https://github.com/O365/python-o365?tab=readme-ov-file#authentication) for authentication details and API.
│       └── [Image of Client Credentials] ([https://github.com/user-attachments/assets/6067f63d-7142-4622-9680-75ce05667483](https://github.com/user-attachments/assets/6067f63d-7142-4622-9680-75ce05667483))
├── jobs_py.xlsx                   (Defines CODES (case-sensitive) for projects. Replace example codes with your own unique letter combinations.)
└── powershell_command.txt         (Command to convert the application into an executable, specific to your credentials)
```

# Steps

1.  **Install python.** Up to you how you want to do this- I was running Windows Subsystem for Linux (available in Software Store) on my Arup machine.
2.  **Clone the repo.** Clone or download
    ```python
    git clone https://github.com/mtp-pura/arup_auto_timesheet.git
    ```
3.  **Install dependencies:** Run the command to install required Python libraries (pandas, datetime, Inquirer, etc.). Note that there might be an error with `zoneinfo` but it might not matter.
    ```python
    pip install pandas datetime inquirer openpyxl zoneinfo re os sys O365 python-dateutil
    ```
4.  **Open the main script.** `cal_py.py`
5.  **Get the secret ID:** Follow the portal link to the Azure app. [Azure Portal Link](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Overview/appId/73fee778-cbb5-4c82-81cd-13503338d848/isMSAApp~/false)
    * **Add a client credential:** Navigate to "certificates and secrets" and create a "new client secret."
    * **Name the secret** and set an expiration (e.g., 24 months).
    * **Copy the "value"** of the newly created secret. This is the actual secret ID needed in the code.
    * **Add the copied "value" into the secret ID spot** in the main script. Save the script.
10. **Double-check the dates** in the script to ensure it's pointed at the desired week (current week, last week, or ad hoc dates). Modify the date settings in the script if needed, depending on which week you need to access.
    * Weeks run Monday to Sunday.
11. **Edit the `jobs.py` file:**
    * Replace the default, custom project codes for your appointments (Column A) and charge codes (Column B) with your own.
12. **Add the relevant code to your calendar appointments' subject lines.** The code should match one of the keys defined in `jobs.py`. You can add any other descriptive text, before or after the code.
    * **Note** that 'SKIP' is in-built and used to ignore any appointments
    * **Note** that you can edit other people's appointments to add your project code, but if they update anything, you will need to edit again.
13. **Run the script for the first time:** (e.g., `python3 cal.py`). First run will require you to grant permission to use the API.
    * **Enter a justification** for accessing the API to **Request approval** from the digital rights team. This approval might take some time.
    * **Once approved, run the script again.** You will be prompted for your password or authenticator ID.
    * **Authenticate:** A blank screen will appear in your browser.
    * **Copy the authenticated URL** from the browser's address bar.
    * **Paste the authenticated URL** into the prompt in your terminal.
14. **Authenticating after first run.** if it's been more than 60 minutes since the last authentication, you will have to reauthenticate.
15. **Fix unmatched appointments.** Any appointments that do not match your Project codes will be flagged, you can assign one using arrow keys, but also can go back into the calendar and rerun.
16. **The script will then print out your timesheet data** in the terminal, which you can quickly review. Make adjustments and rerun, or ENTER any key to close.
17. **The script will also generate a timesheet file** (in the `exported_formatted` folder) based on your calendar entries.
18. **Open the generated timesheet file.**
19. **Save the timesheet file** (even if you don't make any changes) to ensure proper formatting for the expense system.
20. **Import the generated timesheet file** into your expense management system.
21. **Submit the timesheet.**
22. **Convert to EXE.** If everything is running as you'd like, you can convert the script to you own personalised EXE, but running the script in `powershell_command.txt` from within the root folder of the repo. Not necessary but might simplify things further and save you from having to open a command line etc.
