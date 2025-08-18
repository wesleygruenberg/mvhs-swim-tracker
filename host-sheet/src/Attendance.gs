const SHEET_NAMES_ATTENDANCE = {
  roster: 'Swimmers',
  attendance: 'Master Attendance',
};

const ROSTER_HEADERS = {
  name: 'Name',
  gradYear: 'Grad Year',
  gender: 'Gender',
  level: 'Level',
  notes: 'Notes',
};

const ATTEND_HEADERS = {
  date: 'Date',
  name: 'Name',
  present: 'Present',
  excused: 'Excused',
  level: 'Level',
  gender: 'Gender',
  ts: 'Timestamp',
};

function getRosterSorted() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAMES_ATTENDANCE.roster);
  if (!sh) throw new Error('Missing Swimmers sheet');

  const values = sh.getDataRange().getValues();
  const header = values.shift();
  const idx = indexMap(header);

  const rows = values
    .map(r => ({
      name: val(r, idx, ROSTER_HEADERS.name),
      gradYear: val(r, idx, ROSTER_HEADERS.gradYear),
      gender: val(r, idx, ROSTER_HEADERS.gender),
      level: val(r, idx, ROSTER_HEADERS.level),
      notes: val(r, idx, ROSTER_HEADERS.notes),
    }))
    .filter(x => x.name && x.name.trim() !== ''); // Only include rows with names

  // Normalize gender and level values
  rows.forEach(r => {
    // Normalize gender (M/Male -> M, F/Female -> F)
    const genderStr = String(r.gender || '').toUpperCase().trim();
    if (genderStr === 'MALE' || genderStr === 'M') {
      r.gender = 'M';
    } else if (genderStr === 'FEMALE' || genderStr === 'F') {
      r.gender = 'F';
    }

    // Normalize level (V -> Varsity, JV -> JV)
    const levelStr = String(r.level || '').toUpperCase().trim();
    if (levelStr === 'V' || levelStr === 'VARSITY') {
      r.level = 'Varsity';
    } else if (levelStr === 'JV') {
      r.level = 'JV';
    }
  });

  // Filter out any with invalid gender/level after normalization
  const validRows = rows.filter(x => 
    (x.gender === 'M' || x.gender === 'F') && 
    (x.level === 'Varsity' || x.level === 'JV')
  );

  // Sorting: Varsity → JV; M → F; Name asc
  const levelOrder = { Varsity: 0, JV: 1 };
  const genderOrder = { M: 0, F: 1 };

  validRows.sort((a, b) => {
    const lv = (levelOrder[a.level] ?? 99) - (levelOrder[b.level] ?? 99);
    if (lv !== 0) return lv;
    const gv = (genderOrder[a.gender] ?? 99) - (genderOrder[b.gender] ?? 99);
    if (gv !== 0) return gv;
    return a.name.localeCompare(b.name, 'en', { sensitivity: 'base' });
  });

  return validRows;
}

function getAttendanceForDate(yyyy_mm_dd) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAMES_ATTENDANCE.attendance);
  ensureAttendanceHeader(sh);

  const data = sh.getDataRange().getValues();
  const header = data.shift();
  const idx = indexMap(header);

  const attendanceByName = new Map();
  data.forEach(r => {
    const d = toDateString(val(r, idx, ATTEND_HEADERS.date));
    if (d !== yyyy_mm_dd) return;
    const name = String(val(r, idx, ATTEND_HEADERS.name) ?? '').trim();
    if (!name) return;
    const present = toBool(val(r, idx, ATTEND_HEADERS.present));
    const excused = toBool(val(r, idx, ATTEND_HEADERS.excused));
    attendanceByName.set(name, { present, excused });
  });

  return Object.fromEntries(attendanceByName); // { [name]: {present: true/false, excused: true/false} }
}

function upsertAttendance(
  yyyy_mm_dd,
  attendanceArray /* [{name, present, excused, level, gender}] */
) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAMES_ATTENDANCE.attendance);
  ensureAttendanceHeader(sh);

  const range = sh.getDataRange();
  const values = range.getValues();
  const header = values.shift();
  const idx = indexMap(header);

  // Build index of row by (date,name)
  const keyOf = (d, name) => `${d}::${name}`;
  const rowIndexByKey = new Map();
  for (let i = 0; i < values.length; i++) {
    const r = values[i];
    const d = toDateString(val(r, idx, ATTEND_HEADERS.date));
    const name = String(val(r, idx, ATTEND_HEADERS.name) ?? '').trim();
    if (d && name) rowIndexByKey.set(keyOf(d, name), i + 2); // +2 offset for header+1-based
  }

  const nowIso = new Date().toISOString();
  const writes = [];

  attendanceArray.forEach(it => {
    const k = keyOf(yyyy_mm_dd, String(it.name));
    const row = rowIndexByKey.get(k);
    if (row) {
      // Update in place
      writes.push({
        range: sh.getRange(row, idx[ATTEND_HEADERS.present] + 1, 1, 1),
        values: [[!!it.present]],
      });
      writes.push({
        range: sh.getRange(row, idx[ATTEND_HEADERS.excused] + 1, 1, 1),
        values: [[!!it.excused]],
      });
      writes.push({
        range: sh.getRange(row, idx[ATTEND_HEADERS.ts] + 1, 1, 1),
        values: [[nowIso]],
      });
      // Opportunistic denorm refresh if columns exist
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
      rowVals[idx[ATTEND_HEADERS.name]] = String(it.name);
      rowVals[idx[ATTEND_HEADERS.present]] = !!it.present;
      rowVals[idx[ATTEND_HEADERS.excused]] = !!it.excused;
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
  if (!sh) {
    // Create the sheet if it doesn't exist
    const ss = SpreadsheetApp.getActive();
    sh = ss.insertSheet(SHEET_NAMES_ATTENDANCE.attendance);
  }
  
  const lastCol = sh.getLastColumn();
  if (lastCol === 0) {
    // Sheet is empty, add headers
    const want = [
      ATTEND_HEADERS.date,
      ATTEND_HEADERS.name,
      ATTEND_HEADERS.present,
      ATTEND_HEADERS.excused,
      ATTEND_HEADERS.level,
      ATTEND_HEADERS.gender,
      ATTEND_HEADERS.ts,
    ];
    sh.getRange(1, 1, 1, want.length).setValues([want]);
    return;
  }
  
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const want = [
    ATTEND_HEADERS.date,
    ATTEND_HEADERS.name,
    ATTEND_HEADERS.present,
    ATTEND_HEADERS.excused,
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
  const attendanceByName = getAttendanceForDate(yyyy_mm_dd);
  // Merge attendance flags (default false if not set)
  const merged = roster.map(r => ({
    name: r.name,
    level: r.level,
    gender: r.gender,
    present: !!(attendanceByName[r.name]?.present),
    excused: !!(attendanceByName[r.name]?.excused),
  }));
  return { date: yyyy_mm_dd, roster: merged };
}

function api_saveAttendance(
  yyyy_mm_dd,
  attendanceList /* [{name, present, excused, level, gender}] */
) {
  return upsertAttendance(yyyy_mm_dd, attendanceList);
}
