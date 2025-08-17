# Host Sheet Script

This is the wrapper script for the Google Sheet that uses the CoachToolsCore library.

## Setup Instructions

1. Get your Google Sheet's script ID:
   - Open your Google Sheet
   - Go to Extensions → Apps Script
   - Copy the script ID from the URL

2. Get your library's script ID:
   - From your main library project (../src/), run `clasp open`
   - Copy the script ID from the URL

3. Update the configuration:
   - Replace `YOUR_GOOGLE_SHEET_SCRIPT_ID_HERE` in `.clasp.json` with your sheet's script ID
   - Replace `YOUR_LIBRARY_SCRIPT_ID_HERE` in `src/appsscript.json` with your library's script ID

4. Deploy:
   ```
   clasp push
   ```

## Usage

- `clasp push` - Push changes to your Google Sheet's script
- `clasp pull` - Pull changes from your Google Sheet's script
- `clasp open` - Open the Apps Script editor for your sheet

## File Structure

- `setupSheet.gs` - The main wrapper functions that create menus and call library functions
- `appsscript.json` - Configuration including library dependencies
