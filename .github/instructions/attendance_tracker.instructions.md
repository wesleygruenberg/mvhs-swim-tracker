---
applyTo: '**'
---

Got it. Here’s a drop-in Copilot instruction set to build **Option 2: a mobile-friendly Attendance sidebar** (Apps Script HTML Service) that:

- Defaults the date to **today** but lets Brigitta change it.
- **Prefills** from the Master Attendance sheet if data already exists for that date, and stays editable.
- Shows a **single scrolling phone-friendly list**, grouped/sorted **Varsity M/F → JV M/F**, with stable ordering inside each group.
- Writes back to the Master Attendance tab.

Use this as your Copilot task brief. It’s explicit about files, functions, and conventions so it can implement without asking you a bunch of questions.

---

# Copilot Task: Mobile Attendance Sidebar (Apps Script)

## Project context

- Google Sheets bound script (Coach Tools project).
- Add a custom menu item **Coach Tools → Attendance** to launch a **sidebar**.
- Sidebar shows date selector + roster list (checkboxes/toggles), optimized for mobile use in the Google Sheets app.
- Reads/writes to a **Master Attendance** sheet (normalized rows).
- Roster source is **Roster** sheet.

## Data contracts

### 1) `Roster` sheet (assume these headers; adjust mapping if different)

- `SwimmerID` (stable ID or fallback to Name if you don’t have IDs)
- `Name` (e.g., “Adelaide Gruenberg”)
- `Level` (exact `Varsity` or `JV`)
- `Gender` (exact `M` or `F`)
- `Active` (`TRUE`/`FALSE`) — only show `TRUE` in list

> If your roster uses different header names, define a single mapping object at the top of `Attendance.gs` and keep the rest of the code using those logical keys.

### 2) `Master Attendance` sheet (normalized table)

Headers (exact, in row 1):

- `Date` (yyyy-mm-dd)
- `SwimmerID` (or Name if you don’t have IDs—pick one and stick to it)
- `Present` (`TRUE`/`FALSE`)
- `Name` (denormalized for readability—optional but helpful)
- `Level` (`Varsity`/`JV`) (optional but helpful)
- `Gender` (`M`/`F`) (optional but helpful)
- `Timestamp` (ISO string when last updated)

Primary key: `(Date, SwimmerID)`.

## Files to add/modify

- `Code.gs` (menu hook only; keep minimal)
- `Attendance.gs` (all server-side logic for roster + attendance CRUD + sorting)
- `AttendanceUI.html` (sidebar UI; load template with `HtmlService`)
- `AttendanceUI.css` (inline into HTML via templating or `<style>` tag—simpler to keep as an include file via templating)
- `AttendanceUI.js.html` (client JS; use templated `.html` include so Apps Script can inline easily)

> Use `HtmlService.createTemplateFromFile('AttendanceUI')`, and `<?= include('AttendanceUI.js') ?>` pattern for includes.
> Keep UI single-file where possible; fewer round-trips is better on mobile.

## Permissions / manifest

- Ensure scope to show sidebar: `https://www.googleapis.com/auth/script.container.ui`
- If you hit a permission error, run any function or re-authorize from the bound sheet.

---

## Menu & Sidebar bootstrap (`Code.gs`)

```javascript
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Coach Tools')
    .addItem('Attendance', 'openAttendanceSidebar')
    .addToUi();
}

function openAttendanceSidebar() {
  const t = HtmlService.createTemplateFromFile('AttendanceUI');
  // Default date = today in local time (yyyy-mm-dd)
  t.defaultDate = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
  const html = t
    .evaluate()
    .setTitle('Attendance')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME) // default anyway
    .setWidth(360); // narrow for phone
  SpreadsheetApp.getUi().showSidebar(html);
}

// Helper to include partials (css/js)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
```

---

## Server logic (`Attendance.gs`)

**Goals**

- Fetch roster (active swimmers) and **sort** by: Level (Varsity first, then JV), Gender (M then F), then Name ascending.
- For a given date, fetch existing attendance → return presence map so UI can prefill.
- Save posted attendance (upsert rows by `(Date, SwimmerID)`).

```javascript
const SHEET_NAMES = {
  roster: 'Roster',
  attendance: 'Master Attendance',
};

const ROSTER_HEADERS = {
  id: 'SwimmerID',
  name: 'Name',
  level: 'Level',
  gender: 'Gender',
  active: 'Active',
};

const ATTEND_HEADERS = {
  date: 'Date',
  id: 'SwimmerID',
  present: 'Present',
  name: 'Name',
  level: 'Level',
  gender: 'Gender',
  ts: 'Timestamp',
};

function getRosterSorted() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAMES.roster);
  if (!sh) throw new Error('Missing Roster sheet');

  const values = sh.getDataRange().getValues();
  const header = values.shift();
  const idx = indexMap(header);

  const rows = values
    .map(r => ({
      id: val(r, idx, ROSTER_HEADERS.id),
      name: val(r, idx, ROSTER_HEADERS.name),
      level: val(r, idx, ROSTER_HEADERS.level),
      gender: val(r, idx, ROSTER_HEADERS.gender),
      active: toBool(val(r, idx, ROSTER_HEADERS.active)),
    }))
    .filter(x => x.active);

  // Sorting: Varsity → JV; M → F; Name asc
  const levelOrder = { Varsity: 0, JV: 1 };
  const genderOrder = { M: 0, F: 1 };

  rows.sort((a, b) => {
    const lv = (levelOrder[a.level] ?? 99) - (levelOrder[b.level] ?? 99);
    if (lv !== 0) return lv;
    const gv = (genderOrder[a.gender] ?? 99) - (genderOrder[b.gender] ?? 99);
    if (gv !== 0) return gv;
    return a.name.localeCompare(b.name, 'en', { sensitivity: 'base' });
  });

  return rows;
}

function getAttendanceForDate(yyyy_mm_dd) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAMES.attendance);
  ensureAttendanceHeader(sh);

  const data = sh.getDataRange().getValues();
  const header = data.shift();
  const idx = indexMap(header);

  const presentById = new Map();
  data.forEach(r => {
    const d = toDateString(val(r, idx, ATTEND_HEADERS.date));
    if (d !== yyyy_mm_dd) return;
    const id = String(val(r, idx, ATTEND_HEADERS.id) ?? '').trim();
    if (!id) return;
    const present = toBool(val(r, idx, ATTEND_HEADERS.present));
    presentById.set(id, present);
  });

  return Object.fromEntries(presentById); // { [id]: true/false }
}

function upsertAttendance(
  yyyy_mm_dd,
  attendanceArray /* [{id, present, name, level, gender}] */
) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAMES.attendance);
  ensureAttendanceHeader(sh);

  const range = sh.getDataRange();
  const values = range.getValues();
  const header = values.shift();
  const idx = indexMap(header);

  // Build index of row by (date,id)
  const keyOf = (d, id) => `${d}::${id}`;
  const rowIndexByKey = new Map();
  for (let i = 0; i < values.length; i++) {
    const r = values[i];
    const d = toDateString(val(r, idx, ATTEND_HEADERS.date));
    const id = String(val(r, idx, ATTEND_HEADERS.id) ?? '').trim();
    if (d && id) rowIndexByKey.set(keyOf(d, id), i + 2); // +2 offset for header+1-based
  }

  const nowIso = new Date().toISOString();
  const writes = [];

  attendanceArray.forEach(it => {
    const k = keyOf(yyyy_mm_dd, String(it.id));
    const row = rowIndexByKey.get(k);
    if (row) {
      // Update in place
      writes.push({
        range: sh.getRange(row, idx[ATTEND_HEADERS.present] + 1, 1, 1),
        values: [[!!it.present]],
      });
      writes.push({
        range: sh.getRange(row, idx[ATTEND_HEADERS.ts] + 1, 1, 1),
        values: [[nowIso]],
      });
      // Opportunistic denorm refresh if columns exist
      if (idx[ATTEND_HEADERS.name] !== undefined) {
        writes.push({
          range: sh.getRange(row, idx[ATTEND_HEADERS.name] + 1, 1, 1),
          values: [[it.name ?? '']],
        });
      }
      if (idx[ATTEND_HEADERS.level] !== undefined) {
        writes.push({
          range: sh.getRange(row, idx[ATTEND_HEADERS.level] + 1, 1, 1),
          values: [[it.level ?? '']],
        });
      }
      if (idx[ATTEND_HEADERS.gender] !== undefined) {
        writes.push({
          range: sh.getRange(row, idx[ATTEND_HEADERS.gender] + 1, 1, 1),
          values: [[it.gender ?? '']],
        });
      }
    } else {
      // Append
      const rowVals = [];
      rowVals[idx[ATTEND_HEADERS.date]] = yyyy_mm_dd;
      rowVals[idx[ATTEND_HEADERS.id]] = String(it.id);
      rowVals[idx[ATTEND_HEADERS.present]] = !!it.present;
      if (idx[ATTEND_HEADERS.name] !== undefined)
        rowVals[idx[ATTEND_HEADERS.name]] = it.name ?? '';
      if (idx[ATTEND_HEADERS.level] !== undefined)
        rowVals[idx[ATTEND_HEADERS.level]] = it.level ?? '';
      if (idx[ATTEND_HEADERS.gender] !== undefined)
        rowVals[idx[ATTEND_HEADERS.gender]] = it.gender ?? '';
      if (idx[ATTEND_HEADERS.ts] !== undefined)
        rowVals[idx[ATTEND_HEADERS.ts]] = nowIso;

      // Fill sparse to full row length
      for (let i = 0; i < header.length; i++)
        if (rowVals[i] === undefined) rowVals[i] = '';
      sh.appendRow(rowVals);
    }
  });

  // Batch setRange calls
  const batchSets = writes.reduce((acc, w) => {
    w.range.setValues(w.values);
    return acc;
  }, 0);

  return { ok: true, updated: attendanceArray.length };
}

// ----- helpers -----
function ensureAttendanceHeader(sh) {
  if (!sh) throw new Error('Missing Master Attendance sheet');
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const want = [
    ATTEND_HEADERS.date,
    ATTEND_HEADERS.id,
    ATTEND_HEADERS.present,
    ATTEND_HEADERS.name,
    ATTEND_HEADERS.level,
    ATTEND_HEADERS.gender,
    ATTEND_HEADERS.ts,
  ];
  // If header is missing or partial, set it (idempotent, preserves existing trailing columns)
  if (!header || header[0] !== ATTEND_HEADERS.date) {
    sh.clearContents();
    sh.getRange(1, 1, 1, want.length).setValues([want]);
  }
}

function indexMap(header) {
  const map = {};
  header.forEach((h, i) => {
    map[String(h).trim()] = i;
  });
  return map;
}

function val(row, idx, key) {
  const i = idx[key];
  return i === undefined ? '' : row[i];
}

function toBool(v) {
  if (typeof v === 'boolean') return v;
  const s = String(v).toLowerCase();
  return s === 'true' || s === 'y' || s === 'yes' || s === '1';
}

function toDateString(v) {
  if (!v) return '';
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  // strings like "2025-08-18"
  const s = String(v).trim();
  // Normalize if needed
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? s : '';
}

// Exposed to client
function api_getRosterAndAttendance(yyyy_mm_dd) {
  const roster = getRosterSorted();
  const presentById = getAttendanceForDate(yyyy_mm_dd);
  // Merge present flags (default false if not set)
  const merged = roster.map(r => ({
    id: r.id || r.name, // fallback if IDs not used
    name: r.name,
    level: r.level,
    gender: r.gender,
    present: !!presentById[r.id || r.name],
  }));
  return { date: yyyy_mm_dd, roster: merged };
}

function api_saveAttendance(
  yyyy_mm_dd,
  attendanceList /* [{id, present, name, level, gender}] */
) {
  return upsertAttendance(yyyy_mm_dd, attendanceList);
}
```

---

## Sidebar UI (`AttendanceUI.html`)

- One column, finger-sized rows, sticky date header + Save button.
- Keep CSS simple and inline for performance; avoid external fonts.
- Use `google.script.run.withSuccessHandler(...).api_*` for RPC.

```html
<!DOCTYPE html>
<html>
  <head>
    <base target="_top" />
    <meta
      name="viewport"
      content="width=device-width, initial-scale=1, maximum-scale=1"
    />
    <style>
      :root {
        --pad: 12px;
        --row: 48px;
      }
      body {
        margin: 0;
        font-family:
          system-ui,
          -apple-system,
          Segoe UI,
          Roboto,
          sans-serif;
      }
      .bar {
        position: sticky;
        top: 0;
        background: #fff;
        z-index: 10;
        border-bottom: 1px solid #ddd;
        padding: 8px var(--pad);
        display: flex;
        gap: 8px;
        align-items: center;
      }
      input[type='date'] {
        flex: 1;
        height: 36px;
      }
      button {
        height: 36px;
        border: 1px solid #333;
        background: #111;
        color: #fff;
        border-radius: 8px;
        padding: 0 12px;
      }
      .groupTitle {
        padding: 8px var(--pad);
        font-weight: 600;
        color: #333;
        background: #f6f6f6;
        border-top: 1px solid #eee;
      }
      .row {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 0 var(--pad);
        height: var(--row);
        border-bottom: 1px solid #eee;
      }
      .name {
        flex: 1;
        padding-right: 10px;
      }
      .tag {
        font-size: 12px;
        color: #666;
        min-width: 56px;
        text-align: right;
      }
      .chk {
        width: 24px;
        height: 24px;
      }
      .toast {
        position: fixed;
        left: 50%;
        transform: translateX(-50%);
        bottom: 12px;
        background: #111;
        color: #fff;
        padding: 10px 14px;
        border-radius: 8px;
        display: none;
      }
    </style>
  </head>
  <body>
    <div class="bar">
      <input id="date" type="date" value="<?= defaultDate ?>" />
      <button id="saveBtn">Save</button>
    </div>

    <div id="list"></div>
    <div id="toast" class="toast"></div>

    <?= include('AttendanceUI.js') ?>
  </body>
</html>
```

---

## Client logic (`AttendanceUI.js.html`)

```html
<script>
  (function () {
    const dateEl = document.getElementById('date');
    const listEl = document.getElementById('list');
    const saveBtn = document.getElementById('saveBtn');
    const toastEl = document.getElementById('toast');

    // Stateful copy so we can save without re-reading DOM
    let state = { date: dateEl.value, roster: [] };

    function showToast(msg) {
      toastEl.textContent = msg;
      toastEl.style.display = 'block';
      setTimeout(() => (toastEl.style.display = 'none'), 1400);
    }

    function load(dateStr) {
      google.script.run
        .withSuccessHandler(render)
        .withFailureHandler(err => showToast('Load failed'))
        .api_getRosterAndAttendance(dateStr);
    }

    function render(payload) {
      state.date = payload.date;
      state.roster = payload.roster.slice(); // copy
      listEl.innerHTML = '';

      // We already get sorted order from server.
      // Create group headings when level/gender changes: Varsity M/F → JV M/F
      let lastGroup = '';
      state.roster.forEach((r, i) => {
        const group = `${r.level} ${r.gender}`;
        if (group !== lastGroup) {
          const gt = document.createElement('div');
          gt.className = 'groupTitle';
          gt.textContent = group;
          listEl.appendChild(gt);
          lastGroup = group;
        }

        const row = document.createElement('div');
        row.className = 'row';

        const name = document.createElement('div');
        name.className = 'name';
        name.textContent = r.name;

        const tag = document.createElement('div');
        tag.className = 'tag';
        tag.textContent = r.level;

        const chk = document.createElement('input');
        chk.type = 'checkbox';
        chk.className = 'chk';
        chk.checked = !!r.present;
        chk.addEventListener('change', () => {
          r.present = chk.checked;
        });

        row.appendChild(name);
        row.appendChild(tag);
        row.appendChild(chk);
        listEl.appendChild(row);
      });

      showToast('Loaded');
    }

    function save() {
      // Minimal payload
      const payload = state.roster.map(r => ({
        id: r.id,
        name: r.name,
        level: r.level,
        gender: r.gender,
        present: !!r.present,
      }));

      google.script.run
        .withSuccessHandler(() => showToast('Saved'))
        .withFailureHandler(err => showToast('Save failed'))
        .api_saveAttendance(state.date, payload);
    }

    // Events
    dateEl.addEventListener('change', () => {
      const d = dateEl.value;
      if (!d) return;
      load(d);
    });
    saveBtn.addEventListener('click', save);

    // Initial load
    load(dateEl.value);
  })();
</script>
```

---

## UX details & behaviors

- **Default date** is injected from server (`defaultDate`) as `yyyy-mm-dd`.
- Changing the date triggers a fresh read. If rows exist for that date, the checkboxes are **prefilled** accordingly.
- **Save** does an upsert per swimmer `(Date, SwimmerID)`, preserving any previously saved swimmers not visible (e.g., inactive ones) and updating denormalized fields.
- **Sorting** is done server-side to guarantee a single source of truth: `Varsity M/F → JV M/F → Name`.
- **Groups** get lightweight headers (`Varsity M`, `Varsity F`, `JV M`, `JV F`) as the list scrolls.

---

## Edge cases & safeguards

- If `SwimmerID` is empty, fallback to `Name` for the key—but prefer adding a real ID to avoid future name collisions.
- If `Master Attendance` is missing, it is **created with headers** (idempotent).
- If roster columns differ, adjust the `ROSTER_HEADERS` mapping in one place.
- Only `Active == TRUE` swimmers render; deactivated swimmers’ historical attendance stays in the table but won’t appear in today’s UI.

---

## Testing checklist

1. Create minimal `Roster` with a mix of `Varsity/JV` and `M/F`, mark a few inactive.
2. Open sidebar → confirm default date = today and list groups in the right order.
3. Toggle a few checkboxes → Save → Verify rows appear/updated in **Master Attendance**.
4. Change date → mark different pattern → Save → flip back to today → confirm **prefill**.
5. Change a swimmer’s Level/Gender → open sidebar → verify new grouping order (sorting server-side).
6. Mark a swimmer inactive → open sidebar → verify they’re hidden; historical rows remain intact.

---

## Commit & style rules (match owner preferences)

- **Only** remove trailing spaces/tabs on edited or adjacent lines. Do **not** mass-reformat.
- Make **logical commits**:
  - `feat(attendance): add sidebar UI and server API`
  - `chore(attendance): create Master Attendance schema & helpers`
  - `fix(attendance): roster mapping + sort stability`

- Don’t stage random helper/test files you created locally unless they’re part of the solution.
- Include clear summaries of what changed and why.

---

## Future enhancements (optional hooks)

- “All / None” toggles per group header.
- Filter to show only **Absent** after saving (quick double-check).
- Store last used date in `PropertiesService.getUserProperties()` and pre-select it when sidebar opens.
- Add a text field for quick note per swimmer (extend Master Attendance schema with `Note`).

---

This gives Copilot everything it needs: file scaffolding, server APIs, UI, sorting rules, and I/O contracts. When you paste this into your repo as an issue/task or prompt, it should be able to implement in one pass. When it lands, we can wire a one-tap icon in the menu or an onOpen toast that hints “Attendance is in Coach Tools → Attendance” for extra discoverability.
