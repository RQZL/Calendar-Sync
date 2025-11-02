# Run Calendar Sync

Turn the monthly schedule into Google Calendar events with one double-click. The batch file asks for your name, calendar, and Google sign-in each time, so no one edits the script.

---

## 1. What You Need
1. Windows 10 or 11 PC with internet access.
2. A Google account that owns or can edit the target calendar.
3. This folder exactly as provided; keep every file together.
4. The latest schedule export (.xls or .xlsx) placed in the project root.

**Included tools**
- RUN_CALENDAR.bat - launches the workflow.
- system\scripts\run-calendar-script.py - automation logic.
- system\scripts\credentials.sample.json - template for your Google OAuth client (copy to credentials.json before first run; the real file stays out of Git).
- system\python\... - bundled portable Python runtime.

After the first run, the script writes system\scripts\user_settings.json to remember the latest doctor name and calendar.

---

## 2. First Run
1. If system\scripts\credentials.json is missing, copy credentials.sample.json, rename it to credentials.json, and fill in your Google OAuth client ID and secret. Keep the real file private - .gitignore already keeps it out of Git.
2. Drop the newest schedule file into the project root (same folder as RUN_CALENDAR.bat). If multiple schedules exist, the script will let you choose one.
3. Double-click RUN_CALENDAR.bat.
4. When prompted, pick your name with the arrow keys or by typing the number shown. If it is not listed, choose **Enter a different name...** and type it manually. The selection is saved for next time.
5. Review the shift summary and type yes to continue.
6. Sign in through the browser window that appears and allow access. Tokens are cleared on each launch, so this happens every run.
7. Choose which Google Calendar to update. The last used calendar is preselected, but you can move to another editable calendar before pressing Enter.
8. Watch for the success message. Only events inside the schedule's date range are refreshed; older events stay put.
9. Check Google Calendar to confirm.

---

## 3. Monthly Routine
1. Replace the old schedule file with the new month's export (or add the new file alongside the old one).
2. Run RUN_CALENDAR.bat.
3. Accept the stored name and calendar or pick new ones.
4. Approve the Google login and let the script finish.

That is all; stored preferences make each run a quick confirm-and-go.

---

## 4. Troubleshooting and Tips
- **Login window closes or shows invalid_grant**  
  Run the batch file again and complete the sign-in. The script always asks for a fresh login.

- **Name not detected correctly**  
  Choose **Enter a different name...** and type it exactly as shown in the spreadsheet. The script remembers the new value.

- **Switching calendars**  
  Pick a different calendar from the list after signing in. It becomes the new default.

- **Moving to another computer**  
  Copy the entire project folder to the new PC. Run RUN_CALENDAR.bat; the script will prompt for a name and calendar and rebuild user_settings.json.

- **Sharing with teammates**  
  Keep the real credentials.json outside of Git. When you hand someone a copy of the folder, drop your credential file into system\scripts before zipping or sending it so they can run the workflow immediately.

With these steps, even a non-technical teammate can load the latest schedule, double-click RUN_CALENDAR.bat, and have their shifts on Google Calendar in minutes.

