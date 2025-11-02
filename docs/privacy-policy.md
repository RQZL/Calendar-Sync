# Run Calendar Sync Privacy Policy

Effective: {{DATE}}

Run Calendar Sync is a desktop utility that runs entirely on your Windows PC. It reads an Excel schedule that you provide, signs you in to Google Calendar, and creates events on the calendar you choose. The project maintainers do not collect, transmit, or store any personal information.

## Information the tool processes locally
- **Schedule data**: Shift details read from the Excel file you select. The data is used to generate calendar events and is not uploaded anywhere except to your Google Calendar via the Google Calendar API.
- **Google account details**: OAuth tokens and credentials are stored only on your device inside `system/scripts`. Tokens are cleared each time you run the batch file, prompting a fresh login.
- **Preferences**: The script saves your last-used name and calendar ID in `system/scripts/user_settings.json` so future runs can default to those choices.

## Information sharing
Run Calendar Sync does not share data with any external services other than Google Calendar, which is required to create events on your behalf. The application never sends usage analytics, and it does not include any third-party advertising or tracking.

## Your responsibilities
- Keep `credentials.json` private. This file contains your Google OAuth client ID and secret. Do not publish or share it publicly.
- Review the events created in Google Calendar after each run to ensure they match the intended schedule.

## Contact
If you have questions about this policy or encounter issues, please open an issue in the GitHub repository that hosts Run Calendar Sync.
