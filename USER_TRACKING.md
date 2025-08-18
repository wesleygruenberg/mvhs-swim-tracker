# User Tracking in Attendance System

## Overview

The attendance system now tracks **who made changes** and **how they accessed the system** for audit and accountability purposes.

## New Tracking Fields

### **UpdatedBy Column**
- **Purpose**: Records the email address of the user who made the attendance change
- **How it works**: Uses `Session.getActiveUser().getEmail()` to capture the authenticated user
- **Privacy**: Only shows email addresses of users with access to the Google Sheet

### **Source Column**  
- **Purpose**: Identifies whether the change was made via desktop sidebar or mobile web app
- **Values**:
  - `"Desktop Sidebar"` - Change made through Google Sheets sidebar interface
  - `"Mobile Web App"` - Change made through the standalone web app
- **Detection**: Analyzes the browser's User Agent string to identify mobile devices

## What Gets Tracked

### **Every attendance record includes:**
1. **Date** - The attendance date (yyyy-mm-dd)
2. **Name** - The swimmer's name  
3. **Present** - Whether they attended (TRUE/FALSE)
4. **Level** - Varsity/JV designation
5. **Gender** - M/F designation
6. **Timestamp** - When the record was created/updated (ISO format)
7. **UpdatedBy** - Email of the person who made the change ⭐ *NEW*
8. **Source** - Desktop vs Mobile interface used ⭐ *NEW*

## Sample Data

```
Date       | Name      | Present | Level   | Gender | Timestamp            | UpdatedBy           | Source
2025-08-18 | John Doe  | TRUE    | Varsity | M      | 2025-08-18T10:30:00Z | coach@school.edu    | Desktop Sidebar
2025-08-18 | Jane Doe  | FALSE   | JV      | F      | 2025-08-18T15:45:00Z | assistant@school.edu| Mobile Web App
```

## Benefits

### **Accountability**
- See exactly who recorded each swimmer's attendance
- Useful for multi-coach teams or when assistants help with attendance

### **Usage Analytics**
- Track adoption of mobile vs desktop interfaces
- Identify which coaches prefer which access method
- Monitor mobile web app usage patterns

### **Audit Trail**
- Complete history of who changed what and when
- Helpful for resolving attendance disputes
- Required for some school districts' record-keeping policies

## Privacy & Security

### **What's Tracked:**
- ✅ User email addresses (already have sheet access)
- ✅ Device type (mobile vs desktop)
- ✅ Timestamp of changes

### **What's NOT Tracked:**
- ❌ Personal device information
- ❌ IP addresses
- ❌ Location data
- ❌ Full browser details

### **Data Storage:**
- All data stays within your Google Workspace
- No external servers or third-party services
- Same security as your Google Sheets

## Technical Implementation

### **Client Side (JavaScript)**
```javascript
// Captures user agent and sends to server
const USER_AGENT = navigator.userAgent;
google.script.run.api_saveAttendanceWithUserInfo(date, payload, USER_AGENT);
```

### **Server Side (Google Apps Script)**
```javascript
// Detects device type and captures user
const userEmail = Session.getActiveUser().getEmail();
const source = userAgent.includes('Mobile') ? 'Mobile Web App' : 'Desktop Sidebar';
```

## Disabling Tracking (Optional)

If you prefer not to track users, you can:

1. **Remove columns**: Delete the "UpdatedBy" and "Source" columns from existing attendance sheets
2. **Use legacy API**: Call `api_saveAttendance()` instead of `api_saveAttendanceWithUserInfo()`
3. **Data still works**: All core attendance functionality remains unchanged

## Migration

### **Existing Data**
- Old attendance records without tracking info remain intact
- New tracking fields will be blank for historical data
- System automatically adds tracking columns when first used

### **Compatibility**
- Fully backward compatible with existing attendance data
- No data loss or corruption during upgrade
- Works with any existing "Master Attendance" sheet structure

This tracking enhancement provides valuable insights while maintaining the same user-friendly experience for taking attendance.
