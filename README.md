# Google Calendar to Spreadsheet

This Google Apps Script automatically updates a Google Spreadsheet with events from a Google Calendar.

## Setup

1. Open the script file and replace `YOUR_EMAIL_ADDRESS` with your Google Calendar email address.
2. Replace `YOUR_SPREADSHEET_ID` with the ID of your Google Spreadsheet.
3. Replace `KEYWORD1`, `KEYWORD2`, `KEYWORD3` with the keywords you want to exclude from the calendar events.

## Usage

The script will run daily and update the spreadsheet with new events from the calendar. Events that contain any of the excluded keywords will not be added to the spreadsheet.
