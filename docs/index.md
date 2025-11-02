# Run Calendar Sync

Run Calendar Sync converts monthly schedule spreadsheets into Google Calendar events. The Windows batch launcher guides teammates through selecting the latest Excel file, choosing their name and destination calendar, and completing Google OAuth sign-in.

## What it does
- Imports `.xls`/`.xlsx` schedule exports.
- Discovers provider names and remembers the last selection.
- Requests a fresh Google login on every run (no shared tokens).
- Updates only events in the current schedule window, leaving prior entries untouched.

## Requirements
- Windows 10/11 PC with internet access.
- Google account that can edit the destination calendar.
- Private OAuth credentials (`credentials.json`) copied into `system/scripts` before launch.

## Support
For questions or issues, open a GitHub issue on the project repository.
