# Web App Deployment Guide

This document explains how to deploy the attendance tracker as a web app for true mobile access.

## Why Web App?

The Google Sheets mobile app **does not support custom menus or sidebars**. While the sidebar works great on desktop, mobile users need the web app version for proper functionality.

## Deployment Steps

1. **Open the Apps Script Editor**
   - Go to your Google Sheet
   - Click `Extensions > Apps Script`

2. **Deploy as Web App**
   - Click `Deploy > New deployment`
   - Choose type: **Web app**
   - Execute as: **Me (your-email@domain.com)**
   - Who has access: **Anyone with the link** (recommended) or your domain
   - Click **Deploy**

3. **Copy the Web App URL**
   - Copy the web app URL provided after deployment
   - It will look like: `https://script.google.com/macros/s/.../exec`

4. **Configure URL in Script (Optional)**
   - Option A: Paste the URL into `getWebAppUrl_()` function in `setupSheet.gs`
   - Option B: Store via Script Properties:
     ```javascript
     PropertiesService.getScriptProperties().setProperty('ATTENDANCE_WEB_APP_URL', 'your-url-here');
     ```

## Using the Web App

### For Coaches (Desktop)
- Use the **Attendance** menu in Google Sheets
- Click **"📋 Open Attendance Tracker (Sidebar)"** for the familiar sidebar
- Click **"🔗 Show Web App Link"** to get the mobile URL

### For Mobile Users
- Open the web app URL in your mobile browser
- **Add to Home Screen** for one-tap access:
  - **iOS Safari**: Tap share button → "Add to Home Screen"
  - **Android Chrome**: Tap menu → "Add to Home screen"

## Features

Both sidebar and web app versions include:
- ✅ Date selector (defaults to today)
- ✅ Roster sorted by: Varsity M/F → JV M/F → Name
- ✅ Bulk select All/None buttons
- ✅ Auto-save to "Master Attendance" sheet
- ✅ Prefill existing attendance data
- ✅ Mobile-optimized touch targets

## Technical Details

### Data Flow
- **Reads from**: "Swimmers" sheet (Name, Level, Gender)
- **Writes to**: "Master Attendance" sheet (auto-created if missing)
- **Key**: `(Date, Name)` pairs for upsert logic

### Mobile Optimizations
- Touch-friendly 52px row height
- Sticky header with date/save controls
- Responsive design for phone screens
- No external dependencies

### Security
- Web app executes as the sheet owner
- Data stays within your Google account
- No external API calls or data sharing

## Troubleshooting

### "Script function not found" Error
- Ensure you've pushed the latest code with `clasp push`
- Check that functions are properly defined in `setupSheet.gs`

### Web App Not Loading
- Verify deployment is active and published
- Check execution permissions and sharing settings
- Try opening in incognito/private mode

### Mobile Menu Not Showing
- This is expected - use the web app URL instead
- Google Sheets mobile app doesn't support custom menus

## Support

If you encounter issues:
1. Check the Apps Script logs for errors
2. Verify the "Swimmers" sheet has the expected columns
3. Ensure deployment permissions are correctly set
4. Try redeploying if the web app stops working

The web app provides the same functionality as the sidebar but works reliably on all mobile devices.
