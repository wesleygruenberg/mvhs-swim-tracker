---
applyTo: '**'
---
You’ve got a solid v1, but two big truths to weld into the next pass:

1. The **Sheets mobile app does not render custom menus or sidebars**. Your summary’s “works on mobile via Google Sheets app” is a mirage.
2. You already have a **two-project clasp layout** (core lib + host-sheet). Keep attendance in the **host-sheet** project, but add a **standalone Web App** so Brigitta gets a true phone UI.

Below is a **delta brief for Copilot** that builds on what you shipped. It preserves everything you’ve done (sorting, bulk All/None, normalized storage), and adds a web app that hits the *same* APIs.

---

# Copilot Delta: Convert Attendance to Dual-Mode (Sidebar + Web App)

## What stays the same

* Keep your existing **Attendance sidebar** (desktop convenience).
* Keep your **server API**:

  * `api_getRosterAndAttendance(dateStr)`
  * `api_saveAttendance(dateStr, list)`
* Keep the **“Swimmers”** sheet schema and use **Name** as the key (until SwimmerID exists).
* Keep **Varsity M/F → JV M/F → Name** sorting and **All/None** bulk actions.

## What changes (additions)

* Add a **web app endpoint** that serves the *same* HTML UI, optimized for mobile.
* Provide a **menu item** that shows the web app link (so you can grab/share it).
* Since includes were flaky, keep **JS embedded in the HTML** (single file).

---

## Files (host-sheet project)

**New/Updated**

* `setupSheet.gs` → add `doGet()` and a menu item to surface the web app URL.
* `AttendanceUI.html` → keep your embedded JS/CSS; add a tiny banner for web mode.
* *(Optional)* `appsscript.json` → scopes (usually inferred).

### `setupSheet.gs` (delta)

```javascript
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Attendance')
    .addItem('📋 Open Attendance Tracker (Sidebar)', 'openAttendanceSidebar')
    .addItem('🔗 Show Web App Link', 'showAttendanceWebLink')
    .addToUi();
}

function openAttendanceSidebar() {
  const t = HtmlService.createTemplateFromFile('AttendanceUI');
  t.defaultDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  t.isWebApp = false;
  SpreadsheetApp.getUi()
    .showSidebar(t.evaluate().setTitle('Attendance').setWidth(360));
}

// NEW: Web app entrypoint for phones
function doGet(e) {
  const t = HtmlService.createTemplateFromFile('AttendanceUI');
  t.defaultDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  t.isWebApp = true;
  return t.evaluate()
    .setTitle('Attendance')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

// Helper: show the current deployment URL in a dialog
function showAttendanceWebLink() {
  const url = getWebAppUrl_();
  SpreadsheetApp.getUi().alert(
    url
      ? 'Attendance Web App URL:\n\n' + url + '\n\nOpen on your phone and “Add to Home Screen.”'
      : 'No web app deployment found.\nUse: Deploy → New deployment → Web app.'
  );
}

// Try to read the web app URL from the last deployment (best-effort)
function getWebAppUrl_() {
  try {
    // Apps Script doesn’t provide a direct getter; hardcode after deploy if needed.
    // Option A (manual): paste your deployed URL below and return it.
    // return 'https://script.google.com/macros/s/.../exec';

    // Option B (Property): store once after deploy
    return PropertiesService.getScriptProperties().getProperty('ATTENDANCE_WEB_APP_URL') || '';
  } catch (e) {
    return '';
  }
}
```

> After you deploy the web app once, **paste** the URL into `getWebAppUrl_()` or **store** it via Script Properties (e.g., run `PropertiesService.getScriptProperties().setProperty('ATTENDANCE_WEB_APP_URL', '<url>')`).

### `AttendanceUI.html` (delta)

Keep your current HTML/CSS/JS. Add two tiny tweaks:

1. Accept the templated flags/values:

```html
<input id="date" type="date" value="<?= defaultDate ?>">
<script>const IS_WEB_APP = <?= isWebApp ? 'true' : 'false' ?>;</script>
```

2. Show a small hint in web mode:

```html
<? if (isWebApp) { ?>
  <div style="padding:8px 12px; font-size:12px; color:#555; border-bottom:1px dashed #ddd;">
    Tip: Add this page to your home screen for one-tap attendance.
  </div>
<? } ?>
```

No other structural change required. Your client JS continues to use:

```js
google.script.run.withSuccessHandler(...).api_getRosterAndAttendance(date)
google.script.run.withSuccessHandler(...).api_saveAttendance(date, payload)
```

These calls work in both sidebar **and** web app contexts.

---

## Web App Deployment

Once, in the host-sheet script:

1. **Deploy → New deployment**
2. **Type**: Web app
3. **Execute as**: **Me**
4. **Who has access**: “Anyone with the link” (simple) or domain-only
5. **Deploy** → copy URL
6. Either paste into `getWebAppUrl_()` or set `ATTENDANCE_WEB_APP_URL` via Script Properties.

On Brigitta’s phone: open the URL in Safari/Chrome → **Add to Home Screen**.

---

## Data contract (confirm against your current sheets)

* **Swimmers** (used as roster)

  * Name (primary key for now)
  * Level: `Varsity`/`JV` (case-normalize)
  * Gender: `M`/`F` (case-normalize)
  * *(Others allowed: Grad Year, Notes — ignored by attendance)*

* **Master Attendance** (created if missing)

  * `Date` (yyyy-mm-dd)
  * `Name` (acts as key)
  * `Present` (TRUE/FALSE)
  * `Level` (denorm)
  * `Gender` (denorm)
  * `Timestamp` (ISO)

**Key**: `(Date, Name)` until you add `SwimmerID`.
If you later introduce IDs, keep Name as a denorm column and swap the key to `(Date, SwimmerID)`.

---

## Edge-case welds (add if not present)

* **Case/whitespace normalization** for Level/Gender/Name (trim, uppercase first letter, etc.) before sorting and writing.
* **Dedup by Name** if the Swimmers sheet ever gets accidental duplicates (pick first; log a warning).
* **Inactive handling** (optional now). If you later add an `Active` column, filter to `Active == TRUE` in UI but preserve historical rows.

---

## Scopes (usually inferred)

* `spreadsheets.currentonly`
* `script.container.ui` (for sidebar)
* `userinfo.email` (web app “execute as me”)
* No external fetch needed.

---

## Quick regression checklist

* Desktop: Sidebar opens, defaults to today, loads/prefills, saves, All/None work.
* Mobile browser: Web app URL loads, same behavior, finger-friendly UI.
* Sorting: Varsity M → Varsity F → JV M → JV F → name A→Z.
* Existing attendance for a date is pre-checked and **editable**.
* Writes normalize to `Master Attendance` without duplicating keys.

---

## Post-merge polish (optional, easy)

* Add **group All/None** buttons per header (Varsity M, Varsity F, …) in the web app too.
* Store **last-used date** in `UserProperties` and preselect it on load.
* Add a **Notes** field per swimmer (extend schema when ready).

Ship this and you’ll have both routes open: comfy sidebar on desktop, true one-tap app on phone—same brain behind both doors.
