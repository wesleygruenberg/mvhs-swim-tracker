/** =========================
 *  Coach Tools for MVHS Swim — v2.2
 *  Adds: Add Result sidebar, Add Meet sidebar, Add Event sidebar,
 *        Clone Clean Baseline (baseline events, no swimmers/meets)
 *  Keeps: Settings, JV toggle, PR Summary/Dashboard, presets, snapshots, usage checks, JV support, sample team
 *  =========================
 */
const LIB_VER = 'v2.2.0'; // bump each push

// Sheet name constants - centralized to avoid typos and enable easy renaming
const SHEET_NAMES = {
  MEET_ENTRY: 'Meet Entry',
  SWIMMERS: 'Swimmers',
  MEETS: 'Meets',
  EVENTS: 'Events',
  RESULTS: 'Results',
  MEET_EVENTS: 'Meet Events',
  LINEUP_CHECK: 'Lineup Check',
  PR_SUMMARY: 'PR Summary',
  SWIMMER_DASHBOARD: 'Swimmer Dashboard',
  COACH_PACKET: 'Coach Packet',
  SETTINGS: 'Settings',
};

// Configuration constants
const CONFIG = {
  MAX_ENTRY_ROWS: 206,
  MIN_BUFFER_ROWS: 1000,
  BUFFER_EXTRA_ROWS: 200,
};

const EVENT_TYPES = {
  INDIVIDUAL: 'Individual',
  RELAY: 'Relay',
};

const STROKES = {
  FREESTYLE: 'Freestyle',
  BACKSTROKE: 'Backstroke',
  BREASTSTROKE: 'Breaststroke',
  BUTTERFLY: 'Butterfly',
  IM: 'IM',
  MEDLEY: 'Medley',
};

function libInfo() {
  const id = ScriptApp.getScriptId();
  SpreadsheetApp.getActive().toast(
    `CoachToolsCore ${LIB_VER}\nScript ID: ${id}`,
    'Coach Tools',
    6
  );
  return { version: LIB_VER, id };
}

/**
 * About dialog for Coach Tools
 */
function aboutCoachTools() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'About Coach Tools',
    `MVHS Swim Coach Tools ${LIB_VER}

Features:
• Meet management and lineup tracking
• Swimmer and event management 
• Results tracking with PR analysis
• Roster ranking analysis from CSV data
• JV/Varsity support
• Bulk import capabilities
• Snapshot and reporting tools

New in this version:
• CSV Roster Rankings: Generate male/female team rankings with individual event ranks, best ranks, and average rankings

Use "Coach Tools > Roster > Generate Roster Rankings from CSV" to analyze your team's performance data.`,
    ui.ButtonSet.OK
  );
}

// Function to be called from host sheet's onOpen()
function setupCoachToolsMenu() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Coach Tools')
    .addItem('About Coach Tools', 'aboutCoachTools')
    .addItem('Refresh All (safe)', 'refreshAll')
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu('Results')
        .addItem('Add Result (sidebar)', 'openAddResultSidebar')
        .addItem('Refresh PR Summary & Dashboard', 'refreshPRs')
    )
    .addSubMenu(
      ui
        .createMenu('Admin')
        .addItem('Ensure Settings Sheet', 'ensureSettingsSheet')
        .addItem('Apply Limits from Settings', 'applyLimitsFromSettings')
        .addItem('Ensure JV Toggle on Meets', 'ensureMeetsHasJVColumn')
        .addItem(
          'Enable JV/Varsity Support (add JV events + reseed)',
          'enableJVSupport'
        )
        .addItem(
          'Clear Sample Data (Results & assignments)',
          'adminClearSampleData'
        )
        .addItem('Add Meet (sidebar)', 'openAddMeetSidebar')
        .addItem('Add Event (sidebar)', 'openAddEventSidebar')
        .addSeparator()
        .addItem('🧪 Test Roster Rankings', 'testRosterRankingsWithSampleData')
    )
    .addSubMenu(
      ui
        .createMenu('Clone')
        .addItem('Make Clean Copy (reset data)', 'cloneMakeCleanCopy')
        .addItem(
          'New Season Copy (carry forward, drop seniors)',
          'cloneNewSeasonCarryForward'
        )
        .addItem(
          'Clone Clean Baseline (baseline events, no meets/swimmers)',
          'cloneCleanBaseline'
        )
    )
    .addSubMenu(
      ui
        .createMenu('Roster')
        .addItem('Add Swimmer + PRs (sidebar)', 'openAddSwimmerSidebar')
        .addItem(
          'Generate Sample Team (50: 25F/25M, 10V/15JV each)',
          'generateSampleTeam50'
        )
        .addSeparator()
        .addItem(
          'Generate Roster Rankings from CSV',
          'generateRosterRankingsFromCSV'
        )
    )
    .addSubMenu(
      ui
        .createMenu('Import')
        .addItem('Bulk Import (CSV paste)', 'openBulkImportSidebar')
    )
    .addItem('Ensure Meet Presets Table', 'ensureMeetPresetsTemplate')
    .addItem('Apply Meet Presets to Lineup', 'applyMeetPresets')
    .addItem('Check Lineup (Usage & Violations)', 'checkLineup')
    .addItem('Create Snapshot of Current Lineup', 'createSnapshot')
    .addItem('Build Coach Packet (print view)', 'buildCoachPacket')
    .addItem('Debug: Dump Filter State', 'debugDumpFilters')
    .addToUi();
}

/** ---------- Refresh All ---------- */
function refreshAll() {
  try {
    const s = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.PR_SUMMARY);
    if (s) clearAllFilters_(s);
  } catch (e) {
    console.log('Failed to clear PR Summary filters:', e.message);
  }
  try {
    const s = SpreadsheetApp.getActive().getSheetByName(
      SHEET_NAMES.LINEUP_CHECK
    );
    if (s) clearAllFilters_(s);
  } catch (e) {
    console.log('Failed to clear Lineup Check filters:', e.message);
  }

  try {
    ensureSettingsSheet();
    applyLimitsFromSettings();
    ensureMeetsHasJVColumn();
    setupValidations();
    ensureMeetEventsTemplate();
    applyMeetPresets();
    refreshPRs();
    checkLineup();
    buildCoachPacket();
    toast('Refresh All complete.');
  } catch (e) {
    toast('Refresh All error: ' + e.message);
    console.error(e);
  }
}

/** =========================
 * SETTINGS
 * ========================= */
function ensureSettingsSheet() {
  const ss = SpreadsheetApp.getActive();
  let set = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!set) set = ss.insertSheet(SHEET_NAMES.SETTINGS);
  if (set.getLastRow() < 2) {
    set.clear();
    const rows = [
      ['Settings', ''],
      ['', ''],
      ['Season Name', '2025 HS'],
      ['Season Start Year', new Date().getFullYear()],
      ['Drop Grad Year on New Season Copy', new Date().getFullYear() + 1],
      ['', ''],
      ['Limits', ''],
      ['Max Individual Events', 2],
      ['Max Relay Events', 2],
      ['', ''],
      [
        'Notes',
        'Change values in column B; Admin → Apply Limits pushes B8/B9 into Meet Entry.',
      ],
    ];
    set.getRange(1, 1, rows.length, 2).setValues(rows);
    set.getRange('A1').setFontWeight('bold').setFontSize(14);
    set.getRange('A7').setFontWeight('bold');
    set.setColumnWidths(1, 2, 240);
  }
  toast('Settings sheet verified.');
}
function readSettings_(ss) {
  const set = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!set)
    return {
      seasonName: 'Season',
      seasonYear: new Date().getFullYear(),
      dropGradYear: new Date().getFullYear() + 1,
      maxInd: 2,
      maxRel: 2,
    };
  const getVal = label => {
    const f = set.createTextFinder(label).matchEntireCell(true).findNext();
    return f ? set.getRange(f.getRow(), 2).getValue() : null;
  };
  return {
    seasonName: String(getVal('Season Name') || 'Season'),
    seasonYear: Number(getVal('Season Start Year') || new Date().getFullYear()),
    dropGradYear: Number(
      getVal('Drop Grad Year on New Season Copy') ||
        new Date().getFullYear() + 1
    ),
    maxInd: Number(getVal('Max Individual Events') || 2),
    maxRel: Number(getVal('Max Relay Events') || 2),
  };
}
function applyLimitsFromSettings() {
  const ss = SpreadsheetApp.getActive();
  const { maxInd, maxRel } = readSettings_(ss);
  const entry = mustSheet('Meet Entry');
  entry.getRange('B2').setValue(maxInd);
  entry.getRange('B3').setValue(maxRel);
  toast(`Limits set: Individual=${maxInd}, Relay=${maxRel}.`);
}

/** =========================
 * VALIDATIONS / PRESETS / REPORTS
 * ========================= */
function setupValidations() {
  const ss = SpreadsheetApp.getActive();
  const entry = mustSheet(SHEET_NAMES.MEET_ENTRY);
  const sw = mustSheet(SHEET_NAMES.SWIMMERS);
  const me = mustSheet(SHEET_NAMES.MEETS);
  const ev = mustSheet(SHEET_NAMES.EVENTS);
  const results = mustSheet(SHEET_NAMES.RESULTS);

  ensureSwimmersLevelColumn_();

  ss.setNamedRange('SwimmerNames', sw.getRange('A2:A'));
  ss.setNamedRange('MeetNames', me.getRange('A2:A'));
  ss.setNamedRange('EventNames', ev.getRange('A2:A'));

  const startRow = 6,
    last = CONFIG.MAX_ENTRY_ROWS;
  entry.getRange(`A${startRow}:A${last}`).insertCheckboxes();

  const dvMeet = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getRangeByName('MeetNames'), true)
    .build();
  const dvSwimmer = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getRangeByName('SwimmerNames'), true)
    .build();
  const dvEvent = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getRangeByName('EventNames'), true)
    .build();

  entry.getRange('B1').setDataValidation(dvMeet);
  entry.getRange(`H${startRow}:H${last}`).setDataValidation(dvSwimmer);
  entry.getRange(`I${startRow}:L${last}`).setDataValidation(dvSwimmer);

  const resLast = Math.max(
    CONFIG.MIN_BUFFER_ROWS,
    results.getLastRow() + CONFIG.BUFFER_EXTRA_ROWS
  );
  results.getRange('A2:A' + resLast).setDataValidation(dvMeet);
  results.getRange('B2:B' + resLast).setDataValidation(dvEvent);
  results.getRange('C2:C' + resLast).setDataValidation(dvSwimmer);
  results.getRange('D2:E' + resLast).setNumberFormat('mm:ss.00');

  const rules = [];
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(
        `=AND($H6<>"",COUNTIFS($H$6:$H$${CONFIG.MAX_ENTRY_ROWS},$H6,$C$6:$C$${CONFIG.MAX_ENTRY_ROWS},"Individual",$A$6:$A$${CONFIG.MAX_ENTRY_ROWS},TRUE)>$B$2)`
      )
      .setRanges([entry.getRange(`H6:H${CONFIG.MAX_ENTRY_ROWS}`)])
      .setBackground('#F4CCCC')
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(
        `=AND($I6<>"",SUMPRODUCT(($I$6:$L$${CONFIG.MAX_ENTRY_ROWS}=$I6)*($A$6:$A$${CONFIG.MAX_ENTRY_ROWS}=TRUE))>$B$3)`
      )
      .setRanges([entry.getRange(`I6:L${CONFIG.MAX_ENTRY_ROWS}`)])
      .setBackground('#F4CCCC')
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(I6<>"",COUNTIF($I6:$L6,I6)>1)`)
      .setRanges([entry.getRange('I6:L206')])
      .setBackground('#FFE699')
      .build()
  );
  entry.setConditionalFormatRules(rules);

  toast('Validations & formatting refreshed.');
}

function ensureMeetEventsTemplate() {
  const ss = SpreadsheetApp.getActive();
  const me = mustSheet(SHEET_NAMES.MEETS);
  const ev = mustSheet(SHEET_NAMES.EVENTS);
  const out =
    ss.getSheetByName(SHEET_NAMES.MEET_EVENTS) ||
    ss.insertSheet(SHEET_NAMES.MEET_EVENTS);

  if (out.getLastRow() < 1) {
    out
      .getRange(1, 1, 1, 4)
      .setValues([['Meet', 'Event', 'Active?', 'Notes']])
      .setFontWeight('bold');
  }

  const last = out.getLastRow();
  const existing = new Set();
  const data = last >= 2 ? out.getRange(2, 1, last - 1, 2).getValues() : [];
  for (const [m, e] of data) if (m && e) existing.add(m + '|' + e);

  const meets = getColValues(me, 1, 2);
  const evLast = ev.getLastRow();
  const evRows =
    evLast >= 2 ? ev.getRange(2, 1, evLast - 1, 5).getValues() : []; // Event,Type,Dist,Stroke,DefaultActive

  const rowsToAppend = [];
  for (const m of meets) {
    for (const r of evRows) {
      const [ename, , , , defActive] = r;
      if (!ename) continue;
      const key = m + '|' + ename;
      if (!existing.has(key)) {
        rowsToAppend.push([m, ename, !!defActive, '']);
        existing.add(key);
      }
    }
  }
  if (rowsToAppend.length > 0) {
    out
      .getRange(out.getLastRow() + 1, 1, rowsToAppend.length, 4)
      .setValues(rowsToAppend);
  }

  out.autoResizeColumns(1, 4);
  toast('Meet Events table is ready.');
}

function ensureMeetsHasJVColumn() {
  const me = mustSheet('Meets');
  const headers = me
    .getRange(1, 1, 1, me.getLastColumn() || 1)
    .getValues()[0]
    .map(h => String(h || '').trim());
  let col = headers.findIndex(h => h.toLowerCase() === 'has jv?') + 1;
  if (!col) {
    col = me.getLastColumn() + 1;
    me.getRange(1, col).setValue('Has JV?').setFontWeight('bold');
  }
  const startRow = 2,
    endRow = Math.max(me.getLastRow(), 100);
  me.getRange(startRow, col, endRow - startRow + 1, 1).insertCheckboxes();
  const last = me.getLastRow();
  if (last >= startRow) {
    const rng = me.getRange(startRow, col, last - startRow + 1, 1);
    const vals = rng.getValues().map(r => [r[0] === '' ? true : r[0]]);
    rng.setValues(vals);
  }
}
function getMeetHasJV_(meetName) {
  if (!meetName) return true;
  const me = mustSheet('Meets');
  const last = me.getLastRow();
  if (last < 2) return true;
  const headers = me
    .getRange(1, 1, 1, me.getLastColumn())
    .getValues()[0]
    .map(h => String(h || '').trim());
  let jvCol = headers.findIndex(h => h.toLowerCase() === 'has jv?') + 1;
  if (!jvCol) return true;
  const meets = me
    .getRange(2, 1, last - 1, 1)
    .getValues()
    .map(r => String(r[0] || '').trim());
  const idx = meets.findIndex(m => m === meetName);
  if (idx < 0) return true;
  const val = me.getRange(2 + idx, jvCol).getValue();
  return val === '' ? true : !!val;
}
function setPresetsJVForMeet_(meetName, hasJV) {
  const presets = mustSheet('Meet Events');
  const last = presets.getLastRow();
  if (last < 2) return;
  const rows = presets.getRange(2, 1, last - 1, 3).getValues();
  let touched = 0;
  for (let i = 0; i < rows.length; i++) {
    const [m, ev] = rows[i];
    if (m === meetName && /\(JV\)\s*$/.test(String(ev || ''))) {
      if (!hasJV && rows[i][2] !== false) {
        rows[i][2] = false;
        touched++;
      }
    }
  }
  if (touched) presets.getRange(2, 1, last - 1, 3).setValues(rows);
}

function applyMeetPresets() {
  const ss = SpreadsheetApp.getActive();
  const entry = mustSheet(SHEET_NAMES.MEET_ENTRY);
  const presets = mustSheet(SHEET_NAMES.MEET_EVENTS);

  ensureMeetsHasJVColumn();

  const selected = (entry.getRange('B1').getDisplayValue() || '').trim();
  if (!selected) return toast('Pick a meet in B1 first.');

  const hasJV = getMeetHasJV_(selected);
  setPresetsJVForMeet_(selected, hasJV);

  const pLast = presets.getLastRow();
  const pVals =
    pLast >= 2 ? presets.getRange(2, 1, pLast - 1, 3).getValues() : [];
  const map = new Map();
  for (const [meet, ev, active] of pVals) {
    if (meet === selected && ev) map.set(ev, !!active);
  }

  const startRow = 6;
  const lastRow = findLastDataRow(entry, 2, startRow);
  for (let r = startRow; r <= lastRow; r++) {
    const evName = entry.getRange(r, 2).getDisplayValue();
    if (!evName) continue;
    let active = map.has(evName) ? map.get(evName) : true;
    if (!hasJV && /\(JV\)\s*$/.test(evName)) active = false;
    entry.getRange(r, 1).setValue(active);
  }

  toast(
    `Applied presets for "${selected}" (${hasJV ? 'JV enabled' : 'JV disabled'}).`
  );
}

function checkLineup() {
  const ss = SpreadsheetApp.getActive();
  const entry = mustSheet(SHEET_NAMES.MEET_ENTRY);
  const sw = mustSheet(SHEET_NAMES.SWIMMERS);
  const out =
    ss.getSheetByName(SHEET_NAMES.LINEUP_CHECK) ||
    ss.insertSheet(SHEET_NAMES.LINEUP_CHECK);
  out.clear();

  const maxInd = Number(entry.getRange('B2').getValue() || 2);
  const maxRel = Number(entry.getRange('B3').getValue() || 2);
  const swimmers = getColValues(sw, 1, 2);

  const levelCol = findHeaderColumn_(sw, 'Level');
  const nameLevel = {};
  if (levelCol) {
    const last = sw.getLastRow();
    const names = last >= 2 ? sw.getRange(2, 1, last - 1, 1).getValues() : [];
    const levels =
      last >= 2 ? sw.getRange(2, levelCol, last - 1, 1).getValues() : [];
    for (let i = 0; i < names.length; i++)
      if (names[i][0])
        nameLevel[names[i][0]] = String(levels[i][0] || '').trim();
  }

  const startRow = 6;
  const lastRow = findLastDataRow(entry, 2, startRow);

  const indiv = Object.fromEntries(swimmers.map(s => [s, 0]));
  const relay = Object.fromEntries(swimmers.map(s => [s, 0]));
  const dupViolations = [];
  const assignViolations = [];
  const jvMismatch = [];

  for (let r = startRow; r <= lastRow; r++) {
    const active = entry.getRange(r, 1).getValue() === true;
    const evName = entry.getRange(r, 2).getDisplayValue();
    const type = entry.getRange(r, 3).getDisplayValue();
    if (!active || !evName) continue;

    const isJVEvent = /\(JV\)\s*$/.test(evName);

    if (type === EVENT_TYPES.INDIVIDUAL) {
      const name = entry.getRange(r, 8).getDisplayValue().trim();
      if (name) {
        indiv[name] = (indiv[name] || 0) + 1;
        if (isJVEvent && (nameLevel[name] || '').toLowerCase() === 'varsity') {
          jvMismatch.push([r, evName, name]);
        }
      }
    } else if (type === EVENT_TYPES.RELAY) {
      const names = entry
        .getRange(r, 9, 1, 4)
        .getDisplayValues()[0]
        .map(x => x.trim())
        .filter(Boolean);
      const dups = findDuplicates(names);
      if (dups.length) dupViolations.push([r, evName, dups.join(', ')]);
      for (const n of names) {
        relay[n] = (relay[n] || 0) + 1;
        if (isJVEvent && (nameLevel[n] || '').toLowerCase() === 'varsity') {
          jvMismatch.push([r, evName, n]);
        }
      }
    }
  }

  const header = [
    [
      'Swimmer',
      EVENT_TYPES.INDIVIDUAL,
      EVENT_TYPES.RELAY,
      'Limit (Ind)',
      'Limit (Rel)',
      'Status',
    ],
  ];
  const rows = [];
  for (const s of swimmers) {
    if (!s) continue;
    const i = indiv[s] || 0;
    const r = relay[s] || 0;
    const status = i > maxInd || r > maxRel ? 'OVER' : 'OK';
    rows.push([s, i, r, maxInd, maxRel, status]);
    if (i > maxInd)
      assignViolations.push(['', '', s, EVENT_TYPES.INDIVIDUAL, i]);
    if (r > maxRel) assignViolations.push(['', '', s, EVENT_TYPES.RELAY, r]);
  }
  rows.sort((a, b) => a[0].localeCompare(b[0]));

  out.getRange(1, 1, 1, 6).setValues(header).setFontWeight('bold');
  if (rows.length) out.getRange(2, 1, rows.length, 6).setValues(rows);
  safeCreateFilter_(out, out.getRange(1, 1, Math.max(2, rows.length + 1), 6));

  out.autoResizeColumns(1, 6);

  const rng = out.getRange(2, 6, Math.max(rows.length, 1), 1);
  const rules = [];
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('OVER')
      .setBackground('#F4CCCC')
      .setRanges([rng])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('OK')
      .setBackground('#D9EAD3')
      .setRanges([rng])
      .build()
  );
  out.setConditionalFormatRules(rules);

  let row = rows.length + 3;
  out
    .getRange(row, 1, 1, 3)
    .setValues([['Duplicate swimmers in a relay row', '(Row)', '(Event)']])
    .setFontWeight('bold');
  row++;
  if (dupViolations.length) {
    out.getRange(row, 1, dupViolations.length, 3).setValues(dupViolations);
    row += dupViolations.length + 1;
  } else {
    out.getRange(row, 1).setValue('None');
    row += 2;
  }

  out
    .getRange(row, 1, 1, 5)
    .setValues([
      ['Assignments over limits', '(Row)', '(Event)', '(Type)', '(Count)'],
    ])
    .setFontWeight('bold');
  row++;
  if (assignViolations.length) {
    out
      .getRange(row, 1, assignViolations.length, 5)
      .setValues(assignViolations);
    row += assignViolations.length + 1;
  } else {
    out.getRange(row, 1).setValue('None');
    row += 2;
  }

  out
    .getRange(row, 1, 1, 3)
    .setValues([
      [
        'JV/VARSITY mismatches (Varsity swimmers in JV events)',
        '(Row)',
        '(Event)',
      ],
    ])
    .setFontWeight('bold');
  row++;
  if (jvMismatch.length) {
    out
      .getRange(row, 1, jvMismatch.length, 2)
      .setValues(jvMismatch.map(x => [x[0], x[1]]));
  } else {
    out.getRange(row, 1).setValue('None');
  }

  toast('Lineup Check generated.');
}

function createSnapshot() {
  const ss = SpreadsheetApp.getActive();
  const src = mustSheet(SHEET_NAMES.MEET_ENTRY);
  const meet = src.getRange('B1').getDisplayValue() || 'Unspecified Meet';
  const stamp = Utilities.formatDate(
    new Date(),
    ss.getSpreadsheetTimeZone(),
    'yyyy-MM-dd HHmm'
  );
  const name = `Lineup — ${meet} — ${stamp}`;
  const snap = src.copyTo(ss).setName(name);
  snap.getDataRange().copyTo(snap.getDataRange(), { contentsOnly: true });
  toast(`Snapshot saved: ${name}`);
}
function setupPRSuite() {
  createPRSummary();
  createSwimmerDashboard();
}

function createPRSummary() {
  const ss = SpreadsheetApp.getActive();
  const results = ss.getSheetByName(SHEET_NAMES.RESULTS);
  if (!results) throw new Error(`Missing '${SHEET_NAMES.RESULTS}' sheet.`);

  const outName = 'PR Summary';
  let out = ss.getSheetByName(outName) || ss.insertSheet(outName);
  out.clear();

  // Results columns: A Meet, B Event, C Swimmer, D Seed, E Final, F Place, G Notes, H Date, I Is PR?, J Current PR
  const lastRow = results.getLastRow();
  if (lastRow < 2) {
    out.getRange(1, 1).setValue('No results yet.');
    return;
  }
  const data = results.getRange(2, 1, lastRow - 1, 10).getValues();

  // Build maps keyed by "swimmer|event"
  const best = new Map(); // key -> {time, meet, date, count}
  const latest = new Map(); // key -> {time, date, meet}
  for (const r of data) {
    const [meet, event, swimmer, seed, finalTime, , , date] = r;
    if (!swimmer || !event || finalTime === '' || finalTime == null) continue;
    const key = swimmer + '|' + event;
    const t = finalTime; // stored as a serial number (fraction of a day)
    // Count + best
    const b = best.get(key);
    if (!b) best.set(key, { time: t, meet, date, count: 1 });
    else {
      b.count++;
      if (t < b.time) {
        b.time = t;
        b.meet = meet;
        b.date = date;
      }
    }
    // Latest (by date)
    const L = latest.get(key);
    if (!L || (date && date > L.date)) latest.set(key, { time: t, date, meet });
  }

  // Emit rows
  const rows = [];
  for (const [key, v] of best.entries()) {
    const [swimmer, event] = key.split('|');
    const L = latest.get(key);
    rows.push([
      swimmer,
      event,
      v.time, // PR Time
      v.meet || '', // PR Meet
      v.date || '', // PR Date
      v.count, // Races
      L ? L.time : '', // Last Swim
    ]);
  }
  rows.sort((a, b) => a[0].localeCompare(b[0]) || a[1].localeCompare(b[1]));

  // Header
  const header = [
    'Swimmer',
    'Event',
    'PR Time',
    'PR Meet',
    'PR Date',
    'Races',
    'Last Swim',
    'Δ vs PR',
  ];
  out
    .getRange(1, 1, 1, header.length)
    .setValues([header])
    .setFontWeight('bold');

  if (rows.length) {
    out.getRange(2, 1, rows.length, 7).setValues(rows);
    // Add Δ vs PR
    out
      .getRange(2, 8, rows.length, 1)
      .setFormulaR1C1('=IF(AND(RC[-1]<>"",RC[-5]<>""),RC[-1]-RC[-5],"")');
  }

  // Formats + niceties
  out.setFrozenRows(1);
  safeCreateFilter_(out, out.getRange(1, 1, Math.max(2, rows.length + 1), 8));
  out.getRange('C2:C').setNumberFormat('mm:ss.00'); // PR Time
  out.getRange('G2:G').setNumberFormat('mm:ss.00'); // Last Swim
  out.getRange('H2:H').setNumberFormat('[m]:ss.00'); // Δ vs PR (can be 0:xx.xx)
  out.autoResizeColumns(1, 8);
}

function createSwimmerDashboard() {
  const ss = SpreadsheetApp.getActive();
  const prs = ss.getSheetByName('PR Summary') || createPRSummary();
  const swSheet = ss.getSheetByName('Swimmers');
  if (!swSheet) throw new Error("Missing 'Swimmers' sheet.");

  // Ensure a named range for swimmers exists (dynamic full column)
  ss.setNamedRange('SwimmerNames', swSheet.getRange('A2:A'));

  const name = 'Swimmer Dashboard';
  let dash = ss.getSheetByName(name) || ss.insertSheet(name);
  dash.clear();

  // Title + selector
  dash
    .getRange('A1')
    .setValue('Swimmer Dashboard')
    .setFontWeight('bold')
    .setFontSize(14);
  dash.getRange('A3').setValue('Swimmer:').setFontWeight('bold');
  dash
    .getRange('B3')
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInRange(ss.getRangeByName('SwimmerNames'), true)
        .build()
    );
  dash.getRange('B3').setNote('Select a swimmer to filter PRs');

  // Headers
  const headers = [
    'Event',
    'PR Time',
    'PR Meet',
    'PR Date',
    'Races',
    'Last Swim',
    'Δ vs PR',
  ];
  dash.getRange('A5:G5').setValues([headers]).setFontWeight('bold');

  // Filtered view (array formula pulls from PR Summary)
  // PR Summary layout: A Swimmer, B Event, C PR Time, D PR Meet, E PR Date, F Races, G Last Swim, H Δ
  dash.getRange('A6').setFormula(
    `
=IF(B3="","",
  QUERY('PR Summary'!A2:H,
    "select B,C,D,E,F,G,H where A = '" & B3 & "' order by B",
    0
  )
)
  `.trim()
  );

  // Formats
  dash.getRange('B6:B').setNumberFormat('mm:ss.00'); // PR Time
  dash.getRange('G6:G').setNumberFormat('[m]:ss.00'); // Δ vs PR
  dash.getRange('F6:F').setNumberFormat('mm:ss.00'); // Last Swim
  dash.setFrozenRows(5);
  dash.autoResizeColumns(1, 7);
}

function refreshPRs() {
  createPRSummary();
  createSwimmerDashboard();
  toast('PR Summary & Dashboard refreshed.');
}
function createPRSummary() {
  const ss = SpreadsheetApp.getActive();
  const results = mustSheet(SHEET_NAMES.RESULTS);
  const out =
    ss.getSheetByName(SHEET_NAMES.PR_SUMMARY) ||
    ss.insertSheet(SHEET_NAMES.PR_SUMMARY);
  out.clear();
  const last = results.getLastRow();
  if (last < 2) {
    out.getRange(1, 1).setValue('No results yet.');
    return;
  }
  const vals = results.getRange(2, 1, last - 1, 10).getValues();
  const best = new Map(),
    latest = new Map();
  for (const [meet, ev, sw, , fin, , , date] of vals) {
    if (!sw || !ev || fin === '' || fin == null) continue;
    const key = sw + '|' + ev,
      t = fin;
    const b = best.get(key);
    if (!b) best.set(key, { time: t, meet, date, count: 1 });
    else {
      b.count++;
      if (t < b.time) {
        b.time = t;
        b.meet = meet;
        b.date = date;
      }
    }
    const L = latest.get(key);
    if (!L || (date && date > L.date)) latest.set(key, { time: t, date, meet });
  }
  const rows = [];
  for (const [k, v] of best.entries()) {
    const [sw, ev] = k.split('|');
    const L = latest.get(k);
    rows.push([
      sw,
      ev,
      v.time,
      v.meet || '',
      v.date || '',
      v.count,
      L ? L.time : '',
    ]);
  }
  rows.sort((a, b) => a[0].localeCompare(b[0]) || a[1].localeCompare(b[1]));
  const header = [
    'Swimmer',
    'Event',
    'PR Time',
    'PR Meet',
    'PR Date',
    'Races',
    'Last Swim',
    'Δ vs PR',
  ];
  out
    .getRange(1, 1, 1, header.length)
    .setValues([header])
    .setFontWeight('bold');
  if (rows.length) {
    out.getRange(2, 1, rows.length, 7).setValues(rows);
    out
      .getRange(2, 8, rows.length, 1)
      .setFormulaR1C1('=IF(AND(RC[-1]<>"",RC[-5]<>""),RC[-1]-RC[-5],"")');
  }
  out.setFrozenRows(1);
  safeCreateFilter_(out, out.getRange(1, 1, Math.max(2, rows.length + 1), 8));

  out.getRange('C2:C').setNumberFormat('mm:ss.00');
  out.getRange('G2:G').setNumberFormat('mm:ss.00');
  out.getRange('H2:H').setNumberFormat('[m]:ss.00');
  out.autoResizeColumns(1, 8);
}
function createSwimmerDashboard() {
  const ss = SpreadsheetApp.getActive();
  const prs = ss.getSheetByName('PR Summary') || createPRSummary();
  const sw = mustSheet('Swimmers');
  ss.setNamedRange('SwimmerNames', sw.getRange('A2:A'));
  let dash =
    ss.getSheetByName('Swimmer Dashboard') ||
    ss.insertSheet('Swimmer Dashboard');
  dash.clear();
  dash
    .getRange('A1')
    .setValue('Swimmer Dashboard')
    .setFontWeight('bold')
    .setFontSize(14);
  dash.getRange('A3').setValue('Swimmer:').setFontWeight('bold');
  dash
    .getRange('B3')
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInRange(ss.getRangeByName('SwimmerNames'), true)
        .build()
    );
  const headers = [
    'Event',
    'PR Time',
    'PR Meet',
    'PR Date',
    'Races',
    'Last Swim',
    'Δ vs PR',
  ];
  dash.getRange('A5:G5').setValues([headers]).setFontWeight('bold');
  dash.getRange('A6').setFormula(
    `
=IF(B3="","",
  QUERY('PR Summary'!A2:H,
    "select B,C,D,E,F,G,H where A = '" & B3 & "' order by B",
    0
  )
)`.trim()
  );
  dash.getRange('B6:B').setNumberFormat('mm:ss.00');
  dash.getRange('G6:G').setNumberFormat('[m]:ss.00');
  dash.getRange('F6:F').setNumberFormat('mm:ss.00');
  dash.setFrozenRows(5);
  dash.autoResizeColumns(1, 7);
}

function buildCoachPacket() {
  const ss = SpreadsheetApp.getActive();
  const entry = mustSheet('Meet Entry');
  const cp =
    ss.getSheetByName(SHEET_NAMES.COACH_PACKET) ||
    ss.insertSheet(SHEET_NAMES.COACH_PACKET);
  cp.clear();

  const meet = entry.getRange('B1').getDisplayValue() || 'Unspecified Meet';
  cp.getRange('A1')
    .setValue(`Coach Packet — ${meet}`)
    .setFontWeight('bold')
    .setFontSize(14);

  const startRow = 6;
  const lastRow = findLastDataRow(entry, 2, startRow);
  const rows = [['Event', 'Type', 'Heat', 'Lane', 'Individual / Relay Legs']];
  for (let r = startRow; r <= lastRow; r++) {
    const active = entry.getRange(r, 1).getValue() === true;
    if (!active) continue;
    const ev = entry.getRange(r, 2).getDisplayValue();
    const type = entry.getRange(r, 3).getDisplayValue();
    const heat = entry.getRange(r, 6).getDisplayValue();
    const lane = entry.getRange(r, 7).getDisplayValue();
    if (type === EVENT_TYPES.INDIVIDUAL) {
      const n = entry.getRange(r, 8).getDisplayValue();
      rows.push([ev, type, heat, lane, n || '—']);
    } else {
      const legs = entry
        .getRange(r, 9, 1, 4)
        .getDisplayValues()[0]
        .filter(Boolean)
        .join(' • ');
      rows.push([ev, type, heat, lane, legs || '—']);
    }
  }
  if (rows.length === 1) rows.push(['(no active events)', '', '', '', '']);

  cp.getRange(3, 1, rows.length, 5).setValues(rows);
  cp.getRange(3, 1, 1, 5).setFontWeight('bold');
  cp.getRange('A3:E').setWrap(true).setVerticalAlignment('middle');
  cp.setFrozenRows(3);
  cp.setColumnWidth(1, 220);
  cp.setColumnWidth(2, 90);
  cp.setColumnWidth(3, 60);
  cp.setColumnWidth(4, 60);
  cp.setColumnWidth(5, 380);

  toast('Coach Packet built.');
}

/** =========================
 * ROSTER RANKING FUNCTIONALITY
 * ========================= */

/**
 * Generate roster rankings from CSV data
 * Creates separate male/female rankings and calculates aggregate stats
 */
function generateRosterRankingsFromCSV() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'CSV Data Input',
    'Please paste your CSV data (including headers):',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const csvData = response.getResponseText().trim();
  if (!csvData) {
    toast('No CSV data provided.');
    return;
  }

  try {
    processRosterRankings_(csvData);
  } catch (e) {
    toast('Error processing CSV: ' + e.message);
    console.error('CSV processing error:', e);
  }
}

function processRosterRankings_(csvData) {
  // Parse CSV into rows
  const rows = Utilities.parseCsv(csvData);
  if (rows.length < 2) {
    throw new Error('CSV must have at least header and one data row');
  }

  const header = rows[0];
  const swimmers = rows.slice(1);

  // Identify event columns (skip Name, Gender)
  const eventCols = [];
  for (let i = 2; i < header.length; i++) {
    const eventName = header[i] ? header[i].trim() : '';
    if (eventName && eventName !== '') {
      eventCols.push({ name: eventName, idx: i });
    }
  }

  if (eventCols.length === 0) {
    throw new Error('No event columns found in CSV');
  }

  // Build event rankings: {event: {M: [swimmerObj], F: [swimmerObj]}}
  const eventRankings = {};

  for (const { name, idx } of eventCols) {
    const male = [],
      female = [];

    for (const row of swimmers) {
      const swimmerName = row[0] ? row[0].trim() : '';
      const gender = row[1] ? row[1].trim().toUpperCase() : '';
      const timeStr = row[idx] ? row[idx].trim() : '';

      if (swimmerName && gender && timeStr && timeStr !== '') {
        const timeSeconds = parseTimeToSeconds_(timeStr);
        if (timeSeconds > 0) {
          const swimmer = {
            name: swimmerName,
            gender: gender,
            time: timeSeconds,
            timeDisplay: timeStr,
          };

          if (gender === 'M') {
            male.push(swimmer);
          } else if (gender === 'F') {
            female.push(swimmer);
          }
        }
      }
    }

    // Sort by time ascending (fastest first)
    male.sort((a, b) => a.time - b.time);
    female.sort((a, b) => a.time - b.time);

    eventRankings[name] = { M: male, F: female };
  }

  // Generate summary for each swimmer
  const maleRoster = [];
  const femaleRoster = [];

  for (const row of swimmers) {
    const name = row[0] ? row[0].trim() : '';
    const gender = row[1] ? row[1].trim().toUpperCase() : '';

    if (!name || !gender) continue;

    const ranks = [];
    let bestRank = null;
    let bestEvent = null;

    for (const { name: eventName, idx } of eventCols) {
      const timeStr = row[idx] ? row[idx].trim() : '';
      if (timeStr && timeStr !== '') {
        const timeSeconds = parseTimeToSeconds_(timeStr);
        if (timeSeconds > 0) {
          const ranking = eventRankings[eventName][gender];
          const rank = ranking.findIndex(s => s.name === name) + 1;

          if (rank > 0) {
            ranks.push({ event: eventName, rank: rank, time: timeStr });

            if (!bestRank || rank < bestRank) {
              bestRank = rank;
              bestEvent = eventName;
            }
          }
        }
      }
    }

    if (ranks.length > 0) {
      const avgRank = ranks.reduce((sum, r) => sum + r.rank, 0) / ranks.length;
      const eventRanksList = ranks
        .map(r => `${r.event}: #${r.rank} (${r.time})`)
        .join('\n');

      const rosterEntry = [
        name,
        ranks.length, // Number of events
        bestRank,
        bestEvent,
        Math.round(avgRank * 100) / 100, // Round to 2 decimals
        eventRanksList,
      ];

      if (gender === 'M') {
        maleRoster.push(rosterEntry);
      } else if (gender === 'F') {
        femaleRoster.push(rosterEntry);
      }
    }
  }

  // Sort rosters by average rank (best average first)
  maleRoster.sort((a, b) => a[4] - b[4]); // Sort by average rank
  femaleRoster.sort((a, b) => a[4] - b[4]);

  // Create output sheets
  createRosterSheet_('Male Roster Rankings', maleRoster);
  createRosterSheet_('Female Roster Rankings', femaleRoster);

  toast(
    `Roster rankings generated! Created Male (${maleRoster.length}) and Female (${femaleRoster.length}) roster sheets.`
  );
}

function createRosterSheet_(sheetName, rosterData) {
  const ss = SpreadsheetApp.getActive();

  // Delete existing sheet if it exists
  const existingSheet = ss.getSheetByName(sheetName);
  if (existingSheet) {
    ss.deleteSheet(existingSheet);
  }

  // Create new sheet
  const sheet = ss.insertSheet(sheetName);

  // Set headers
  const headers = [
    'Swimmer',
    'Events Count',
    'Best Rank',
    'Best Event',
    'Average Rank',
    'All Event Rankings',
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  // Add data
  if (rosterData.length > 0) {
    sheet
      .getRange(2, 1, rosterData.length, headers.length)
      .setValues(rosterData);
  }

  // Format the sheet
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);

  // Add conditional formatting for best ranks
  const bestRankRange = sheet.getRange(2, 3, Math.max(rosterData.length, 1), 1);
  const rule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(1)
    .setBackground('#34A853') // Green for 1st place
    .setFontColor('#FFFFFF')
    .setRanges([bestRankRange])
    .build();

  const rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(2)
    .setBackground('#FBBC04') // Yellow for 2nd place
    .setRanges([bestRankRange])
    .build();

  const rule3 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(3)
    .setBackground('#FF9900') // Orange for 3rd place
    .setRanges([bestRankRange])
    .build();

  sheet.setConditionalFormatRules([rule1, rule2, rule3]);

  // Set text wrapping for the rankings column
  if (rosterData.length > 0) {
    sheet.getRange(2, 6, rosterData.length, 1).setWrap(true);
  }
}

// Helper: parse time string to seconds (supports mm:ss.xx, ss.xx, and various formats)
function parseTimeToSeconds_(timeStr) {
  if (!timeStr || timeStr.trim() === '') return 0;

  // Clean the time string - remove parentheses and extra text
  let cleanTime = timeStr.replace(/\([^)]*\)/g, '').trim();

  // Handle various invalid formats
  if (cleanTime.includes('?') || cleanTime === '00:00.0' || cleanTime === '') {
    return 0;
  }

  // Handle obvious errors like 22:00.0 for backstroke (probably 22.00 seconds)
  if (cleanTime.match(/^\d{2}:\d{2}\.\d$/)) {
    const parts = cleanTime.split(':');
    const minutes = parseInt(parts[0], 10);
    const seconds = parseFloat(parts[1]);

    // If minutes seem too high for swimming (>10 minutes), treat as seconds
    if (minutes > 10) {
      return minutes + seconds;
    }
    return minutes * 60 + seconds;
  }

  // Handle mm:ss.xx format
  if (cleanTime.includes(':')) {
    const parts = cleanTime.split(':');
    if (parts.length === 2) {
      const minutes = parseInt(parts[0], 10);
      const seconds = parseFloat(parts[1]);
      return minutes * 60 + seconds;
    }
  }

  // Handle ss.xx format or just seconds
  const numValue = parseFloat(cleanTime);
  return isNaN(numValue) ? 0 : numValue;
}

/**
 * Test the roster ranking function with the provided sample CSV data
 */
function testRosterRankingsWithSampleData() {
  const sampleCSV = `Name,Gender (M/F),50 Free,50 Fly,100 Fly,100 IM,200 IM,100 Back,200 Free,100 Breast,100 Free,200F,500 Free
Abigal Chally,F,00:28.9,,1:12.21,,2:54.12,22:00.0,2:26.80,1:37.52,1:05.9,,,
Ace Garcia,M,00:23.21,,00:28.41,,08:00.0,1:01.16,,,38:24.0,,5:06.90
Connor Roy,M,00:24.48,,1:00.30,,2:14.16,04:00.0,,1:04.13,00:28.48,,34:00.0
Garret Black,M,00:27.57,,1:23.12,,2:46.556,1:14.21,2:24.19,1:13.15,02:24.0,,6:38.75
Jacek Brown,M,00:23.77,,00:58.83,,2:13.96,1:09.2,01:00.0,1:05?,04:48.0,,27:00.0
Quetzal Carrillo,F,00:27.4,,1:13.00,,2:30.6,1:13.12,2:12.95,1:19.24,1:00.97,,6:19.11
Olivia Hussman,F,00:27.4,,1:18.00,,2:47.88,1:15.68,2:23.63,1:25.37,1:01.92,,6:31.97`;

  try {
    processRosterRankings_(sampleCSV);
    toast('Test completed! Check the Male and Female Roster Rankings sheets.');
  } catch (e) {
    toast('Test failed: ' + e.message);
    console.error('Test error:', e);
  }
}

/** =========================
 * ADMIN & ROSTER + JV SUPPORT
 * ========================= */
function ensureSwimmersLevelColumn_() {
  const sw = mustSheet('Swimmers');
  const headers = sw
    .getRange(1, 1, 1, Math.max(sw.getLastColumn(), 5))
    .getValues()[0];
  const norm = headers.map(h =>
    String(h || '')
      .trim()
      .toLowerCase()
  );
  if (!norm.includes('level')) {
    sw.insertColumnAfter(3); // D
    sw.getRange(1, 4).setValue('Level').setFontWeight('bold');
    if (!sw.getRange(1, 5).getValue())
      sw.getRange(1, 5).setValue('Notes').setFontWeight('bold');
  }
  // Ensure base headers exist
  sw.getRange(1, 1, 1, 5)
    .setValues([['Name', 'Grad Year', 'Gender', 'Level', 'Notes']])
    .setFontWeight('bold');
}

function adminClearSampleData() {
  const ss = SpreadsheetApp.getActive();
  const results = mustSheet('Results');
  const entry = mustSheet('Meet Entry');
  const rLast = results.getLastRow();
  if (rLast >= 2) results.getRange(2, 1, rLast - 1, 10).clearContent();
  reseedMeetEntryFromEvents_();
  entry.getRange('B1').setValue('');
  toast(
    'Sample data cleared (Results & assignments). Meet Entry reseeded from Events.'
  );
}

function generateSampleTeam50() {
  const ss = SpreadsheetApp.getActive();
  const sw = mustSheet('Swimmers');
  ensureSwimmersLevelColumn_();
  const last = sw.getLastRow();
  if (last >= 2) sw.getRange(2, 1, last - 1, sw.getLastColumn()).clearContent();

  const firstF = [
    'Avery',
    'Riley',
    'Jordan',
    'Taylor',
    'Casey',
    'Parker',
    'Quinn',
    'Rowan',
    'Emerson',
    'Hayden',
    'Morgan',
    'Reese',
    'Skyler',
    'Alex',
    'Drew',
    'Logan',
    'Cameron',
    'Charlie',
    'Harper',
    'Kendall',
    'Sage',
    'Blake',
    'Finley',
    'Sydney',
    'Payton',
  ];
  const firstM = [
    'Liam',
    'Noah',
    'Oliver',
    'Elijah',
    'James',
    'Benjamin',
    'Lucas',
    'Henry',
    'Alexander',
    'Mason',
    'Michael',
    'Ethan',
    'Daniel',
    'Jacob',
    'Logan',
    'Jackson',
    'Levi',
    'Sebastian',
    'Mateo',
    'Jack',
    'Owen',
    'Theodore',
    'Aiden',
    'Samuel',
    'Joseph',
  ];
  const lastNames = [
    'Brooks',
    'Carter',
    'Diaz',
    'Ellis',
    'Foster',
    'Garcia',
    'Hayes',
    'Ingram',
    'Jensen',
    'Kim',
    'Lopez',
    'Miller',
    'Nguyen',
    'Ortiz',
    'Patel',
    'Quintero',
    'Rivera',
    'Shaw',
    'Turner',
    'Underwood',
    'Vargas',
    'Walker',
    'Xu',
    'Young',
    'Zimmerman',
  ];

  const year = readSettings_(ss).seasonYear || new Date().getFullYear();
  const grads = [year + 1, year + 2, year + 3, year + 4];
  const rows = [];
  function pick(pool, n) {
    const a = pool.slice();
    for (let i = a.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [a[i], a[j]] = [a[j], a[i]];
    }
    return a.slice(0, n);
  }
  const fNames = pick(firstF, 25),
    mNames = pick(firstM, 25),
    lNames = pick(lastNames, 50);

  for (let i = 0; i < 25; i++)
    rows.push([
      `${fNames[i]} ${lNames[i]}`,
      grads[i % 4],
      'F',
      i < 10 ? 'Varsity' : 'JV',
      '',
    ]);
  for (let i = 0; i < 25; i++)
    rows.push([
      `${mNames[i]} ${lNames[25 + i]}`,
      grads[(i + 1) % 4],
      'M',
      i < 10 ? 'Varsity' : 'JV',
      '',
    ]);

  sw.getRange(2, 1, rows.length, 5).setValues(rows);
  toast(
    'Sample team generated: 50 swimmers (25F/25M; 10 Varsity + 15 JV per gender).'
  );
}

function enableJVSupport() {
  const ss = SpreadsheetApp.getActive();
  const ev = mustSheet('Events');
  const last = ev.getLastRow();
  if (last < 2) throw new Error('Events sheet is empty.');
  const rows = ev.getRange(2, 1, last - 1, 5).getValues();
  const existing = new Set(rows.map(r => r[0]));
  const toAppend = [];
  for (const r of rows) {
    const name = String(r[0] || '');
    if (!name || /\(JV\)\s*$/.test(name)) continue;
    const jvName = `${name} (JV)`;
    if (!existing.has(jvName)) {
      const copy = r.slice();
      copy[0] = jvName;
      toAppend.push(copy);
    }
  }
  if (toAppend.length)
    ev.getRange(ev.getLastRow() + 1, 1, toAppend.length, 5).setValues(toAppend);
  reseedMeetEntryFromEvents_();
  ensureMeetEventsTemplate();
  applyMeetPresets();
  toast(
    'JV support enabled: JV event variants added, Meet Entry reseeded, presets refreshed.'
  );
}

function reseedMeetEntryFromEvents_() {
  const entry = mustSheet('Meet Entry');
  const ev = mustSheet('Events');
  const lastEntry = entry.getLastRow();
  if (lastEntry > 5)
    entry.getRange(6, 1, lastEntry - 5, 12).clear({ contentsOnly: true });
  const rows = ev
    .getRange(2, 1, Math.max(ev.getLastRow() - 1, 0), 5)
    .getValues();
  let r = 6;
  for (const e of rows) {
    const [name, type, dist, stroke, defActive] = e;
    if (!name) continue;
    entry.getRange(r, 1).setValue(!!defActive);
    entry.getRange(r, 2).setValue(name);
    entry.getRange(r, 3).setValue(type);
    entry.getRange(r, 4).setValue(dist);
    entry.getRange(r, 5).setValue(stroke);
    r++;
  }
}

/** =========================
 * CLONE: CLEAN, NEW SEASON, CLEAN BASELINE
 * ========================= */
function cloneMakeCleanCopy() {
  const src = SpreadsheetApp.getActive();
  ensureSettingsSheet();
  const settings = readSettings_(src);
  const newName = `${settings.seasonName} — CLEAN COPY — ${timestamp_()}`;
  const file = DriveApp.getFileById(src.getId());
  const copy = file.makeCopy(newName);
  const tgt = SpreadsheetApp.openById(copy.getId());
  resetDataInCopy_(tgt, {
    carryForward: false,
    dropGradYear: settings.dropGradYear,
  });
  toast(`Clean copy created.\nURL: ${copy.getUrl()}`);
}

function cloneNewSeasonCarryForward() {
  const src = SpreadsheetApp.getActive();
  ensureSettingsSheet();
  const settings = readSettings_(src);
  const nextSeasonName = `${settings.seasonName || 'Season'} NEXT`;
  const newName = `${nextSeasonName} — NEW SEASON — ${timestamp_()}`;
  const file = DriveApp.getFileById(src.getId());
  const copy = file.makeCopy(newName);
  const tgt = SpreadsheetApp.openById(copy.getId());
  resetDataInCopy_(tgt, {
    carryForward: true,
    dropGradYear: settings.dropGradYear,
  });
  const set = tgt.getSheetByName('Settings');
  if (set) {
    const finder = set
      .createTextFinder('Season Start Year')
      .matchEntireCell(true)
      .findNext();
    if (finder) {
      const row = finder.getRow();
      const cur = Number(
        set.getRange(row, 2).getValue() || settings.seasonYear
      );
      set.getRange(row, 2).setValue(cur + 1);
    }
  }
  toast(`New season copy created.\nURL: ${copy.getUrl()}`);
}

// NEW: make a copy with baseline events, no swimmers, no meets
function cloneCleanBaseline() {
  const src = SpreadsheetApp.getActive();
  const newName = `Swim Tracker — CLEAN BASELINE — ${timestamp_()}`;
  const file = DriveApp.getFileById(src.getId());
  const copy = file.makeCopy(newName);
  const tgt = SpreadsheetApp.openById(copy.getId());

  // Reset sheets
  // Swimmers -> headers only
  const sw = tgt.getSheetByName('Swimmers') || tgt.insertSheet('Swimmers');
  sw.clear();
  sw.getRange(1, 1, 1, 5)
    .setValues([['Name', 'Grad Year', 'Gender', 'Level', 'Notes']])
    .setFontWeight('bold');

  // Meets -> headers + Has JV? column, no rows
  const me = tgt.getSheetByName('Meets') || tgt.insertSheet('Meets');
  me.clear();
  me.getRange(1, 1, 1, 5)
    .setValues([['Meet', 'Date', 'Location', 'Course', 'Season/Notes']])
    .setFontWeight('bold');
  tgt.setActiveSheet(me);
  ensureMeetsHasJVColumn(); // adds the Has JV? column

  // Events -> baseline set (no JV)
  const ev = tgt.getSheetByName('Events') || tgt.insertSheet('Events');
  ev.clear();
  ev.getRange(1, 1, 1, 5)
    .setValues([['Event', 'Type', 'Distance', 'Stroke', 'Default Active?']])
    .setFontWeight('bold');
  const baseline = [
    ['200 Medley Relay', 'Relay', 200, 'Medley', true],
    ['200 Freestyle', 'Individual', 200, 'Freestyle', true],
    ['200 Individual Medley', 'Individual', 200, 'IM', true],
    ['50 Freestyle', 'Individual', 50, 'Freestyle', true],
    ['100 Butterfly', 'Individual', 100, 'Butterfly', true],
    ['100 Freestyle', 'Individual', 100, 'Freestyle', true],
    ['500 Freestyle', 'Individual', 500, 'Freestyle', true],
    ['200 Freestyle Relay', 'Relay', 200, 'Freestyle', true],
    ['100 Backstroke', 'Individual', 100, 'Backstroke', true],
    ['100 Breaststroke', 'Individual', 100, 'Breaststroke', true],
    ['400 Freestyle Relay', 'Relay', 400, 'Freestyle', true],
    // extras default OFF
    ['200 Backstroke', 'Individual', 200, 'Backstroke', false],
    ['200 Breaststroke', 'Individual', 200, 'Breaststroke', false],
    ['200 Butterfly', 'Individual', 200, 'Butterfly', false],
    ['400 Individual Medley', 'Individual', 400, 'IM', false],
    ['50 Butterfly', 'Individual', 50, 'Butterfly', false],
    ['50 Backstroke', 'Individual', 50, 'Backstroke', false],
    ['50 Breaststroke', 'Individual', 50, 'Breaststroke', false],
  ];
  if (baseline.length)
    ev.getRange(2, 1, baseline.length, 5).setValues(baseline);

  // Results -> header only
  const res = tgt.getSheetByName('Results') || tgt.insertSheet('Results');
  res.clear();
  res
    .getRange(1, 1, 1, 10)
    .setValues([
      [
        'Meet',
        'Event',
        'Swimmer',
        'Seed Time (mm:ss.00)',
        'Final Time (mm:ss.00)',
        'Place',
        'Notes',
        'Date',
        'Is PR?',
        'Current PR',
      ],
    ])
    .setFontWeight('bold');

  // Meet Entry -> reseed
  const entry =
    tgt.getSheetByName('Meet Entry') || tgt.insertSheet('Meet Entry');
  // If sheet exists, keep top rows (labels) and reseed; else you may want to copy from source—here we do a minimal rebuild:
  entry.clear();
  entry.getRange(1, 1).setValue('Selected Meet').setFontWeight('bold');
  entry.getRange(1, 2).setValue('');
  entry
    .getRange(2, 1)
    .setValue('Max Individual Events per Swimmer')
    .setFontWeight('bold');
  entry.getRange(2, 2).setValue(2);
  entry
    .getRange(3, 1)
    .setValue('Max Relay Events per Swimmer')
    .setFontWeight('bold');
  entry.getRange(3, 2).setValue(2);
  entry
    .getRange(4, 1, 1, 12)
    .setValues([
      [
        'Active?',
        'Event',
        'Type',
        'Distance',
        'Stroke',
        'Heat',
        'Lane',
        'Swimmer (Individual)',
        'Relay Leg 1',
        'Relay Leg 2',
        'Relay Leg 3',
        'Relay Leg 4',
      ],
    ])
    .setFontWeight('bold');
  tgt.setActiveSheet(ev); // reseed uses active file's Events
  SpreadsheetApp.setActiveSpreadsheet(tgt);
  reseedMeetEntryFromEvents_();

  // Meet Events -> just header
  const presets =
    tgt.getSheetByName('Meet Events') || tgt.insertSheet('Meet Events');
  presets.clear();
  presets
    .getRange(1, 1, 1, 4)
    .setValues([['Meet', 'Event', 'Active?', 'Notes']])
    .setFontWeight('bold');

  // Derived views -> remove; will rebuild on demand
  ['PR Summary', 'Swimmer Dashboard', 'Lineup Check', 'Coach Packet'].forEach(
    n => {
      const sh = tgt.getSheetByName(n);
      if (sh) tgt.deleteSheet(sh);
    }
  );
  // Snapshots
  tgt.getSheets().forEach(sh => {
    if (sh.getName().startsWith('Lineup — ')) tgt.deleteSheet(sh);
  });

  // Final: set validations in the copy
  setupValidationsFor_(tgt);
  ensureMeetEventsTemplateFor_(tgt);

  toast(`Clean baseline clone created:\n${copy.getUrl()}`);
}

function resetDataInCopy_(ss, opts) {
  const { carryForward, dropGradYear } = opts;

  if (!ss.getSheetByName('Settings')) {
    const setSrc = SpreadsheetApp.getActive().getSheetByName('Settings');
    if (setSrc) {
      const setCopy = ss.insertSheet('Settings');
      const rng = setSrc.getDataRange();
      setCopy
        .getRange(1, 1, rng.getNumRows(), rng.getNumColumns())
        .setValues(rng.getValues());
    } else {
      ss.insertSheet('Settings');
    }
  }
  const results = ss.getSheetByName('Results');
  if (results) {
    const last = results.getLastRow();
    if (last >= 2) results.getRange(2, 1, last - 1, 10).clearContent();
  }
  const entry = ss.getSheetByName('Meet Entry');
  const events = ss.getSheetByName('Events');
  if (entry && events) {
    entry.getRange('B1').setValue('');
    const startRow = 6;
    const lastRow = findLastDataRow(entry, 2, startRow);
    const evMap = new Map();
    const er = events.getLastRow();
    const eRows = er >= 2 ? events.getRange(2, 1, er - 1, 5).getValues() : [];
    for (const r of eRows) evMap.set(r[0], !!r[4]);
    for (let r = startRow; r <= lastRow; r++) {
      const ev = entry.getRange(r, 2).getDisplayValue();
      entry.getRange(r, 1).setValue(evMap.has(ev) ? evMap.get(ev) : true);
      entry.getRange(r, 6, 1, 7).clearContent();
    }
    const { maxInd, maxRel } = readSettings_(ss);
    entry.getRange('B2').setValue(maxInd);
    entry.getRange('B3').setValue(maxRel);
  }
  const sw = ss.getSheetByName('Swimmers');
  if (sw) {
    if (carryForward) {
      const last = sw.getLastRow();
      if (last >= 2) {
        const vals = sw.getRange(2, 1, last - 1, 4).getValues();
        const kept = vals.filter(r => Number(r[1]) !== Number(dropGradYear));
        sw.getRange(2, 1, last - 1, 4).clearContent();
        if (kept.length) sw.getRange(2, 1, kept.length, 4).setValues(kept);
      }
    }
  }
  ['PR Summary', 'Swimmer Dashboard', 'Lineup Check', 'Coach Packet'].forEach(
    n => {
      const sh = ss.getSheetByName(n);
      if (sh) ss.deleteSheet(sh);
    }
  );
  ss.getSheets().forEach(sh => {
    const name = sh.getName();
    if (name.startsWith('Lineup — ')) ss.deleteSheet(sh);
  });
  setupValidationsFor_(ss);
  ensureMeetEventsTemplateFor_(ss);
}

/** Parametric helpers for copies */
function setupValidationsFor_(ss) {
  const entry = _mustSheet(ss, SHEET_NAMES.MEET_ENTRY);
  const sw = _mustSheet(ss, SHEET_NAMES.SWIMMERS);
  const me = _mustSheet(ss, SHEET_NAMES.MEETS);
  const ev = _mustSheet(ss, SHEET_NAMES.EVENTS);
  const results = _mustSheet(ss, SHEET_NAMES.RESULTS);

  ss.setNamedRange('SwimmerNames', sw.getRange('A2:A'));
  ss.setNamedRange('MeetNames', me.getRange('A2:A'));
  ss.setNamedRange('EventNames', ev.getRange('A2:A'));

  const startRow = 6,
    last = CONFIG.MAX_ENTRY_ROWS;
  entry.getRange(`A${startRow}:A${last}`).insertCheckboxes();

  const dvMeet = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getRangeByName('MeetNames'), true)
    .build();
  const dvSwimmer = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getRangeByName('SwimmerNames'), true)
    .build();
  entry.getRange('B1').setDataValidation(dvMeet);
  entry.getRange(`H${startRow}:H${last}`).setDataValidation(dvSwimmer);
  entry.getRange(`I${startRow}:L${last}`).setDataValidation(dvSwimmer);

  const resLast = Math.max(
    CONFIG.MIN_BUFFER_ROWS,
    results.getLastRow() + CONFIG.BUFFER_EXTRA_ROWS
  );
  const dvEvent = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getRangeByName('EventNames'), true)
    .build();
  results.getRange('A2:A' + resLast).setDataValidation(dvMeet);
  results.getRange('B2:B' + resLast).setDataValidation(dvEvent);
  results.getRange('C2:C' + resLast).setDataValidation(dvSwimmer);
  results.getRange('D2:E' + resLast).setNumberFormat('mm:ss.00');
}
function ensureMeetEventsTemplateFor_(ss) {
  const me = _mustSheet(ss, 'Meets');
  const ev = _mustSheet(ss, 'Events');
  const out = ss.getSheetByName('Meet Events') || ss.insertSheet('Meet Events');
  if (out.getLastRow() < 1)
    out
      .getRange(1, 1, 1, 4)
      .setValues([['Meet', 'Event', 'Active?', 'Notes']])
      .setFontWeight('bold');
  const last = out.getLastRow();
  const existing = new Set();
  const data = last >= 2 ? out.getRange(2, 1, last - 1, 2).getValues() : [];
  for (const [m, e] of data) if (m && e) existing.add(m + '|' + e);
  const meets = _getColValues(me, 1, 2);
  const evLast = ev.getLastRow();
  const evRows =
    evLast >= 2 ? ev.getRange(2, 1, evLast - 1, 5).getValues() : [];
  const rowsToAppend = [];
  for (const m of meets)
    for (const r of evRows) {
      const [ename, , , , defActive] = r;
      if (!ename) continue;
      const key = m + '|' + ename;
      if (!existing.has(key)) {
        rowsToAppend.push([m, ename, !!defActive, '']);
        existing.add(key);
      }
    }
  if (rowsToAppend.length)
    out
      .getRange(out.getLastRow() + 1, 1, rowsToAppend.length, 4)
      .setValues(rowsToAppend);
}

/** =========================
 * ROSTER: Add Swimmer + PRs (existing)
 * ========================= */
function openAddSwimmerSidebar() {
  const html = HtmlService.createHtmlOutput(addSwimmerSidebarHtml_()).setTitle(
    'Add Swimmer + PRs'
  );
  SpreadsheetApp.getUi().showSidebar(html);
}
function getIndividualEventsForPR_() {
  const ev = mustSheet('Events');
  const last = ev.getLastRow();
  if (last < 2) return [];
  const vals = ev.getRange(2, 1, last - 1, 2).getValues();
  return vals
    .map(r => ({ name: String(r[0] || ''), type: String(r[1] || '') }))
    .filter(
      x => x.type === 'Individual' && x.name && !/\(JV\)\s*$/.test(x.name)
    )
    .map(x => x.name);
}
function addSwimmerWithPRs(payload) {
  const ss = SpreadsheetApp.getActive();
  const sw = mustSheet('Swimmers');
  const results = mustSheet('Results');
  ensureSwimmersLevelColumn_();

  const name = String(payload.name || '').trim();
  if (!name) throw new Error('Name is required.');
  const grad = Number(payload.gradYear || '');
  const gender = String(payload.gender || '').trim() || '';
  const level = String(payload.level || '').trim() || '';
  const date = payload.date ? new Date(payload.date) : new Date();
  const prs = payload.prs || {};

  const last = sw.getLastRow();
  let rowIdx = -1;
  if (last >= 2) {
    const names = sw
      .getRange(2, 1, last - 1, 1)
      .getValues()
      .map(r => String(r[0] || ''));
    rowIdx = names.findIndex(n => n === name);
  }
  const levelCol = findHeaderColumn_(sw, 'Level') || 4;
  if (rowIdx >= 0) {
    const r = 2 + rowIdx;
    if (grad) sw.getRange(r, 2).setValue(grad);
    if (gender) sw.getRange(r, 3).setValue(gender);
    if (level) sw.getRange(r, levelCol).setValue(level);
  } else {
    sw.getRange(sw.getLastRow() + 1, 1, 1, 5).setValues([
      [name, grad || '', gender || '', level || '', ''],
    ]);
  }

  const rows = [];
  const meetLabel = 'PR Baseline';
  for (const [evName, tStr] of Object.entries(prs)) {
    const serial = parseTimeSerial_(tStr);
    if (serial == null) continue;
    rows.push([
      meetLabel,
      evName,
      name,
      '',
      serial,
      '',
      'Added via sidebar',
      date,
    ]);
  }
  if (rows.length) {
    const startRow = results.getLastRow() + 1;
    results.getRange(startRow, 1, rows.length, 8).setValues(rows);
  }
  try {
    setupValidations();
  } catch (e) {
    console.log('Failed to setup validations:', e.message);
  }
  try {
    refreshPRs();
  } catch (e) {
    console.log('Failed to refresh PRs:', e.message);
  }

  return { added: rowIdx < 0, prCount: rows.length };
}
function parseTimeSerial_(s) {
  if (s == null) return null;
  s = String(s).trim();
  if (!s) return null;
  let m = s.match(/^(\d+):(\d{1,2})(?:\.(\d+))?$/);
  if (m) {
    const minutes = parseInt(m[1], 10);
    const seconds = parseInt(m[2], 10) + (m[3] ? parseFloat('0.' + m[3]) : 0);
    const total = minutes * 60 + seconds;
    return total / 86400;
  }
  m = s.match(/^(\d+(?:\.\d+)?)$/);
  if (m) return parseFloat(m[1]) / 86400;
  return null;
}
function findHeaderColumn_(sheet, headerText) {
  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map(h =>
      String(h || '')
        .trim()
        .toLowerCase()
    );
  const idx = headers.indexOf(String(headerText).trim().toLowerCase());
  return idx >= 0 ? idx + 1 : 0;
}
function addSwimmerSidebarHtml_() {
  const events = JSON.stringify(getIndividualEventsForPR_());
  return `
<!doctype html><html><head><meta charset="utf-8">
<style>
body{font:13px/1.4 Arial,sans-serif;padding:12px}h2{margin:0 0 8px}
label{display:block;margin-top:8px;font-weight:bold}
input,select{width:100%;box-sizing:border-box;padding:6px}
.grid{display:grid;grid-template-columns:1fr 120px;gap:6px 8px}.row{display:contents}
.fine{color:#666;font-size:11px}.btn{margin-top:12px;width:100%;padding:8px;font-weight:bold}
.ok{background:#1e8e3e;color:#fff;border:0}.warn{background:#e37400;color:#fff;border:0}
.section{margin-top:12px;border-top:1px solid #ddd;padding-top:8px}.pill{display:inline-block;padding:2px 6px;background:#eee;border-radius:999px;font-size:11px;margin-left:6px}
</style></head><body>
<h2>Add Swimmer <span class="pill">+ optional PRs</span></h2>
<label>Name</label><input id="name" type="text" placeholder="First Last">
<div class="grid">
  <div class="row"><label>Grad Year</label><input id="grad" type="number" min="2024" max="2035" step="1"></div>
  <div class="row"><label>Gender</label><select id="gender"><option value="">—</option><option>F</option><option>M</option><option>N/A</option></select></div>
  <div class="row"><label>Level</label><select id="level"><option value="">—</option><option>Varsity</option><option>JV</option></select></div>
  <div class="row"><label>PR Date</label><input id="date" type="date"></div>
</div>
<div class="section">
  <label>Personal Records (optional)</label>
  <div class="fine">Enter times as <b>mm:ss.xx</b> or <b>ss.xx</b>. Leave blank to skip.</div>
  <div id="events" class="grid"></div>
</div>
<button class="btn ok" onclick="submitForm()">Add Swimmer & PRs</button>
<button class="btn warn" onclick="google.script.host.close()">Close</button>
<script>
  const EVENTS = ${events}; const evDiv = document.getElementById('events');
  function addEventRow(name){const wrap=document.createElement('div');wrap.className='row';
    const lab=document.createElement('div');lab.textContent=name;
    const inpWrap=document.createElement('div');const inp=document.createElement('input');
    inp.type='text';inp.placeholder='e.g., 1:05.32 or 28.75';inp.dataset.event=name;inpWrap.appendChild(inp);
    wrap.appendChild(lab);wrap.appendChild(inpWrap);evDiv.appendChild(wrap);}
  EVENTS.forEach(addEventRow);
  function submitForm(){
    const name=document.getElementById('name').value.trim(); if(!name){alert('Name is required.');return;}
    const grad=document.getElementById('grad').value, gender=document.getElementById('gender').value, level=document.getElementById('level').value, date=document.getElementById('date').value;
    const prs={}; document.querySelectorAll('#events input[type=text]').forEach(i=>{const v=i.value.trim(); if(v) prs[i.dataset.event]=v;});
    google.script.run.withSuccessHandler(res=>{alert('Saved ✓ ' + (res.prCount||0) + ' PRs recorded');google.script.host.close();})
      .withFailureHandler(err=>alert('Error: '+err.message)).addSwimmerWithPRs({name:name,gradYear:grad,gender:gender,level:level,date:date,prs:prs});
  }
</script></body></html>`;
}

/** =========================
 * RESULTS: Add Result (NEW)
 * ========================= */
function openAddResultSidebar() {
  const html = HtmlService.createHtmlOutput(addResultSidebarHtml_()).setTitle(
    'Add Result'
  );
  SpreadsheetApp.getUi().showSidebar(html);
}
function listMeetNames_() {
  return getColValues(mustSheet(SHEET_NAMES.MEETS), 1, 2);
}
function listSwimmerNames_() {
  return getColValues(mustSheet(SHEET_NAMES.SWIMMERS), 1, 2);
}
function listEventNames_() {
  return getColValues(mustSheet(SHEET_NAMES.EVENTS), 1, 2);
}
function listActiveEventsForMeet(meet) {
  if (!meet) return listEventNames_();
  const presets = mustSheet('Meet Events');
  const last = presets.getLastRow();
  if (last < 2) return listEventNames_();
  const vals = presets.getRange(2, 1, last - 1, 3).getValues();
  const set = [];
  for (const [m, e, active] of vals) {
    if (m === meet && !!active && e) set.push(e);
  }
  return set.length ? set : listEventNames_();
}
function getCurrentPR(swimmer, eventName) {
  if (!swimmer || !eventName) return null;
  const res = mustSheet('Results');
  const last = res.getLastRow();
  if (last < 2) return null;
  const vals = res.getRange(2, 1, last - 1, 10).getValues(); // meet,event,swimmer,seed,final,place,notes,date,isPR,curPR
  let best = null;
  for (const r of vals) {
    if (
      String(r[1] || '') === eventName &&
      String(r[2] || '') === swimmer &&
      r[4] !== '' &&
      r[4] != null
    ) {
      const t = Number(r[4]);
      if (best == null || t < best) best = t;
    }
  }
  return best; // serial or null
}
function addResultRow(payload) {
  const res = mustSheet('Results');
  const meet = String(payload.meet || '').trim();
  const eventName = String(payload.event || '').trim();
  const swimmer = String(payload.swimmer || '').trim();
  if (!meet || !eventName || !swimmer)
    throw new Error('Meet, Event, and Swimmer are required.');
  const seedSerial = parseTimeSerial_(payload.seed || '');
  const finalSerial = parseTimeSerial_(payload.final || '');
  if (finalSerial == null)
    throw new Error('Final time is required (mm:ss.xx or ss.xx).');
  const place = payload.place || '';
  const notes = payload.notes || '';
  const date = payload.date ? new Date(payload.date) : new Date();
  res
    .getRange(res.getLastRow() + 1, 1, 1, 8)
    .setValues([
      [
        meet,
        eventName,
        swimmer,
        seedSerial || '',
        finalSerial,
        place,
        notes,
        date,
      ],
    ]);
  try {
    refreshPRs();
  } catch (e) {
    console.log('Failed to refresh PRs after adding result:', e.message);
  }
  return { ok: true };
}
function addResultSidebarHtml_() {
  const meets = JSON.stringify(listMeetNames_());
  const swimmers = JSON.stringify(listSwimmerNames_());
  const allEvents = JSON.stringify(listEventNames_());
  return `
<!doctype html><html><head><meta charset="utf-8">
<style>
body{font:13px/1.4 Arial,sans-serif;padding:12px}h2{margin:0 0 8px}
label{display:block;margin-top:8px;font-weight:bold}
input,select,textarea{width:100%;box-sizing:border-box;padding:6px}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:6px 8px}.row{display:contents}
.fine{color:#666;font-size:11px}.btn{margin-top:12px;width:100%;padding:8px;font-weight:bold}
.ok{background:#1e8e3e;color:#fff;border:0}.warn{background:#e37400;color:#fff;border:0}
.small{font-size:12px}
</style></head><body>
<h2>Add Result</h2>
<label>Meet</label>
<select id="meet"></select>
<label>Event</label>
<select id="event"></select>
<label>Swimmer</label>
<select id="swimmer"></select>
<div class="grid">
  <div class="row"><label>Seed Time</label><input id="seed" type="text" placeholder="mm:ss.xx or ss.xx"></div>
  <div class="row"><label>Final Time*</label><input id="final" type="text" placeholder="mm:ss.xx or ss.xx"></div>
  <div class="row"><label>Place</label><input id="place" type="text" placeholder="e.g., 2"></div>
  <div class="row"><label>Date</label><input id="date" type="date"></div>
</div>
<label>Notes</label><textarea id="notes" rows="2" placeholder="Optional"></textarea>
<div class="fine" id="prhint"></div>
<button class="btn ok" onclick="submitForm()">Save Result</button>
<button class="btn warn" onclick="google.script.host.close()">Close</button>
<script>
const MEETS=${meets}, SWIMMERS=${swimmers}, ALL_EVENTS=${allEvents};
const meetSel=document.getElementById('meet'), eventSel=document.getElementById('event'), swimSel=document.getElementById('swimmer'), hint=document.getElementById('prhint');

function fill(sel, arr){ sel.innerHTML=''; arr.forEach(v=>{const o=document.createElement('option');o.textContent=v;o.value=v; sel.appendChild(o);}); }

fill(meetSel, MEETS); fill(swimSel, SWIMMERS); fill(eventSel, ALL_EVENTS);

meetSel.addEventListener('change', ()=>{ google.script.run.withSuccessHandler(list=>{ fill(eventSel, list); prCheck(); }).listActiveEventsForMeet(meetSel.value); });
eventSel.addEventListener('change', prCheck); swimSel.addEventListener('change', prCheck);

function prCheck(){
  const sw=swimSel.value, ev=eventSel.value; if(!sw||!ev){ hint.textContent=''; return; }
  google.script.run.withSuccessHandler(serial=>{
    if(serial==null){ hint.textContent='No PR recorded yet for this swimmer/event.'; return; }
    // Convert serial days -> mm:ss.xx
    const sec = serial*86400; const m=Math.floor(sec/60); const s=(sec%60).toFixed(2).padStart(5,'0'); 
    hint.innerHTML = 'Current PR: <b>'+m+':'+s+'</b>';
  }).getCurrentPR(sw, ev);
}

function submitForm(){
  const payload = {
    meet: meetSel.value, event: eventSel.value, swimmer: swimSel.value,
    seed: document.getElementById('seed').value, final: document.getElementById('final').value,
    place: document.getElementById('place').value, notes: document.getElementById('notes').value,
    date: document.getElementById('date').value
  };
  google.script.run.withSuccessHandler(()=>{ alert('Saved ✓'); google.script.host.close(); })
    .withFailureHandler(err=>alert('Error: '+err.message))
    .addResultRow(payload);
}
</script></body></html>`;
}

/** =========================
 * ADMIN: Add Meet / Add Event (NEW)
 * ========================= */
function openAddMeetSidebar() {
  const html = HtmlService.createHtmlOutput(addMeetSidebarHtml_()).setTitle(
    'Add Meet'
  );
  SpreadsheetApp.getUi().showSidebar(html);
}
function addMeet(payload) {
  const me = mustSheet('Meets');
  const name = String(payload.name || '').trim();
  if (!name) throw new Error('Meet name is required.');
  const date = payload.date ? new Date(payload.date) : '';
  const loc = String(payload.location || '').trim();
  const course = String(payload.course || '').trim(); // SCY/LCM/SCM
  const notes = String(payload.notes || '').trim();
  // Append
  me.appendRow([name, date, loc, course, notes]);
  ensureMeetsHasJVColumn();
  ensureMeetEventsTemplate();
  setupValidations();
  return { ok: true };
}
function addMeetSidebarHtml_() {
  return `
<!doctype html><html><head><meta charset="utf-8"><style>
body{font:13px/1.4 Arial,sans-serif;padding:12px}h2{margin:0 0 8px}
label{display:block;margin-top:8px;font-weight:bold}input,select,textarea{width:100%;padding:6px;box-sizing:border-box}
.btn{margin-top:12px;width:100%;padding:8px;font-weight:bold;background:#1e8e3e;color:#fff;border:0}
</style></head><body>
<h2>Add Meet</h2>
<label>Name*</label><input id="name" type="text" placeholder="e.g., Boise City Classic">
<label>Date</label><input id="date" type="date">
<label>Location</label><input id="location" type="text" placeholder="Pool name">
<label>Course</label><select id="course"><option value="">—</option><option>SCY</option><option>LCM</option><option>SCM</option></select>
<label>Notes</label><textarea id="notes" rows="2"></textarea>
<button class="btn" onclick="go()">Add Meet</button>
<script>
function go(){
  const p={name:document.getElementById('name').value,date:document.getElementById('date').value,location:document.getElementById('location').value,course:document.getElementById('course').value,notes:document.getElementById('notes').value};
  if(!p.name.trim()){alert('Name is required.');return;}
  google.script.run.withSuccessHandler(()=>{alert('Meet added ✓');google.script.host.close();})
    .withFailureHandler(err=>alert('Error: '+err.message)).addMeet(p);
}
</script></body></html>`;
}

function openAddEventSidebar() {
  const html = HtmlService.createHtmlOutput(addEventSidebarHtml_()).setTitle(
    'Add Event'
  );
  SpreadsheetApp.getUi().showSidebar(html);
}
function addEvent(payload) {
  debugLog_('addEvent', 'called', payload);
  const ev = mustSheet('Events');
  const name = String(payload.name || '').trim();
  if (!name) throw new Error('Event name is required.');
  const type = String(payload.type || '').trim() || 'Individual';
  const dist = Number(payload.distance || '');
  const stroke = String(payload.stroke || '').trim();
  const defActive = !!payload.defaultActive;
  const addJV = !!payload.addJV;
  const reseed = !!payload.reseed;

  debugLog_('addEvent', 'parsed values', {
    name,
    type,
    dist,
    stroke,
    defActive,
    addJV,
    reseed,
  });

  ev.appendRow([name, type, dist || '', stroke, defActive]);
  debugLog_('addEvent', 'added main event row', {
    name,
    type,
    dist,
    stroke,
    defActive,
  });
  if (addJV) {
    ev.appendRow([`${name} (JV)`, type, dist || '', stroke, defActive]);
    debugLog_('addEvent', 'added JV event row', {
      name: `${name} (JV)`,
      type,
      dist,
      stroke,
      defActive,
    });
  }

  if (reseed) {
    reseedMeetEntryFromEvents_();
    debugLog_('addEvent', 'reseeded meet entry');
  }
  ensureMeetEventsTemplate();
  setupValidations();
  debugLog_('addEvent', 'completed successfully');

  return { ok: true };
}
function addEventSidebarHtml_() {
  return `
<!doctype html><html><head><meta charset="utf-8"><style>
body{font:13px/1.4 Arial,sans-serif;padding:12px}h2{margin:0 0 8px}
label{display:block;margin-top:8px;font-weight:bold}input,select{width:100%;padding:6px;box-sizing:border-box}
.row{display:grid;grid-template-columns:1fr 1fr;gap:6px 8px}.btn{margin-top:12px;width:100%;padding:8px;font-weight:bold;background:#1e8e3e;color:#fff;border:0}
</style></head><body>
<h2>Add Event</h2>
<label>Event Name*</label><input id="name" type="text" placeholder="e.g., 100 Backstroke">
<div class="row">
  <div><label>Type</label><select id="type"><option>Individual</option><option>Relay</option></select></div>
  <div><label>Distance</label><input id="dist" type="number" step="1" placeholder="e.g., 100"></div>
</div>
<label>Stroke</label><input id="stroke" type="text" placeholder="e.g., Backstroke / Freestyle / IM">
<label><input id="def" type="checkbox"> Default Active?</label>
<label><input id="jv" type="checkbox"> Also create JV variant</label>
<label><input id="reseed" type="checkbox"> Reseed Meet Entry now (rebuild rows)</label>
<button class="btn" onclick="go()">Add Event</button>
<script>
function go(){
  const p={name:document.getElementById('name').value,type:document.getElementById('type').value,distance:document.getElementById('dist').value,stroke:document.getElementById('stroke').value,defaultActive:document.getElementById('def').checked,addJV:document.getElementById('jv').checked,reseed:document.getElementById('reseed').checked};
  console.log('Form data:', p);
  if(!p.name.trim()){alert('Event name is required.');return;}
  console.log('Calling addEvent with payload:', p);
  google.script.run.withSuccessHandler((result)=>{
    console.log('Success:', result);
    alert('Event added ✓');
    google.script.host.close();
  }).withFailureHandler(err=>{
    console.error('Error:', err);
    alert('Error: '+err.message);
  }).addEvent(p);
}
</script></body></html>`;
}

/** =========================
 * HELPERS
 * ========================= */
function mustSheet(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error(`Missing required sheet: "${name}"`);
  return sh;
}
function _mustSheet(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Missing required sheet in copy: "${name}"`);
  return sh;
}
function getColValues(sheet, col, startRow = 2) {
  const last = sheet.getLastRow();
  if (last < startRow) return [];
  return sheet
    .getRange(startRow, col, last - startRow + 1, 1)
    .getValues()
    .map(r => r[0])
    .filter(Boolean);
}
function _getColValues(sheet, col, startRow = 2) {
  const last = sheet.getLastRow();
  if (last < startRow) return [];
  return sheet
    .getRange(startRow, col, last - startRow + 1, 1)
    .getValues()
    .map(r => r[0])
    .filter(Boolean);
}
function findLastDataRow(sheet, keyCol, startRow) {
  const last = sheet.getLastRow();
  if (last < startRow) return startRow - 1;
  const vals = sheet
    .getRange(startRow, keyCol, last - startRow + 1, 1)
    .getValues()
    .map(r => r[0]);
  let end = startRow - 1;
  for (let i = 0; i < vals.length; i++) if (vals[i]) end = startRow + i;
  return end;
}
function findDuplicates(arr) {
  const seen = new Set(),
    dup = new Set();
  for (const x of arr) {
    if (seen.has(x)) dup.add(x);
    else seen.add(x);
  }
  return [...dup];
}
function toast(msg) {
  SpreadsheetApp.getActive().toast(msg, 'Coach Tools', 5);
}
function timestamp_() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd HHmm'
  );
}

/** =========================
 * IMPORT: Bulk Import (CSV Paste)
 * ========================= */
function openBulkImportSidebar() {
  const html = HtmlService.createHtmlOutput(bulkImportSidebarHtml_()).setTitle(
    'Bulk Import'
  );
  SpreadsheetApp.getUi().showSidebar(html);
}

// Server: do the import
function bulkImport(payload) {
  const type = String(payload.type || '').toLowerCase(); // 'swimmers' | 'meets' | 'pr'
  const csv = String(payload.csv || '').trim();
  const hasHeader = !!payload.hasHeader;
  const defaultDate = payload.defaultDate
    ? new Date(payload.defaultDate)
    : new Date();

  if (!csv) throw new Error('Paste CSV data first.');

  // Parse CSV (handles quotes/commas properly)
  const rows = Utilities.parseCsv(csv).filter(r =>
    r.some(c => String(c).trim() !== '')
  );
  if (!rows.length) throw new Error('No rows detected.');
  const data = hasHeader ? rows.slice(1) : rows;

  if (type === 'swimmers') return importSwimmers_(data);
  if (type === 'meets') return importMeets_(data);
  if (type === 'pr') return importPRs_(data, defaultDate);

  throw new Error('Unknown import type: ' + type);
}

function importSwimmers_(data) {
  const sw = mustSheet('Swimmers');
  ensureSwimmersLevelColumn_(); // ensures Name, Grad Year, Gender, Level, Notes headers exist
  const out = [];
  for (const r of data) {
    const name = (r[0] || '').toString().trim();
    if (!name) continue;
    const grad = r[1] ? Number(r[1]) : '';
    const gender = (r[2] || '').toString().trim();
    const level = (r[3] || '').toString().trim();
    const notes = (r[4] || '').toString();
    out.push([name, grad, gender, level, notes]);
  }
  if (out.length)
    sw.getRange(sw.getLastRow() + 1, 1, out.length, 5).setValues(out);
  setupValidations();
  return { inserted: out.length, kind: 'swimmers' };
}

function importMeets_(data) {
  const me = mustSheet('Meets');
  const out = [];
  const jvMarks = [];
  for (const r of data) {
    const name = (r[0] || '').toString().trim();
    if (!name) continue;
    const date = r[1] ? new Date(r[1]) : '';
    const loc = (r[2] || '').toString().trim();
    const course = (r[3] || '').toString().trim(); // SCY/LCM/SCM
    const notes = (r[4] || '').toString().trim();
    const hasJV = (r[5] || '').toString().trim().toLowerCase();
    out.push([name, date, loc, course, notes]);
    jvMarks.push(hasJV); // remember per-row intent
  }
  if (out.length) {
    const start = me.getLastRow() + 1;
    me.getRange(start, 1, out.length, 5).setValues(out);
    ensureMeetsHasJVColumn();
    // set Has JV? checkboxes using text flags like 'true','yes','y','1'
    const headers = me
      .getRange(1, 1, 1, me.getLastColumn())
      .getValues()[0]
      .map(h =>
        String(h || '')
          .trim()
          .toLowerCase()
      );
    const jvCol = headers.indexOf('has jv?') + 1;
    for (let i = 0; i < jvMarks.length; i++) {
      const val = jvMarks[i];
      const isTrue = ['true', 'yes', 'y', '1'].includes(val);
      me.getRange(start + i, jvCol).setValue(isTrue);
    }
  }
  ensureMeetEventsTemplate(); // cross-join presets for new meets
  setupValidations();
  return { inserted: out.length, kind: 'meets' };
}

function importPRs_(data, fallbackDate) {
  const res = mustSheet('Results');
  const rows = [];
  for (const r of data) {
    const swimmer = (r[0] || '').toString().trim();
    const event = (r[1] || '').toString().trim();
    const timeStr = (r[2] || '').toString().trim();
    if (!swimmer || !event || !timeStr) continue;
    const serial = parseTimeSerial_(timeStr);
    if (serial == null) continue;
    const date = r[3] ? new Date(r[3]) : fallbackDate;
    rows.push([
      'PR Baseline',
      event,
      swimmer,
      '',
      serial,
      '',
      'Imported baseline',
      date,
    ]);
  }
  if (rows.length)
    res.getRange(res.getLastRow() + 1, 1, rows.length, 8).setValues(rows);
  refreshPRs();
  return { inserted: rows.length, kind: 'pr' };
}

function bulkImportSidebarHtml_() {
  const tmpl = `
<!doctype html><html><head><meta charset="utf-8">
<style>
body{font:13px/1.4 Arial,sans-serif;padding:12px}
h2{margin:0 0 8px} label{display:block;margin-top:8px;font-weight:bold}
select,textarea,input{width:100%;box-sizing:border-box;padding:6px}
textarea{height:180px;font-family:ui-monospace,Consolas,Monaco,monospace}
.small{font-size:12px;color:#666} .btn{margin-top:12px;width:100%;padding:8px;font-weight:bold;background:#1e8e3e;color:#fff;border:0}
.row{display:grid;grid-template-columns:1fr 1fr;gap:8px}
pre{background:#f6f6f6;padding:8px;overflow:auto}
</style></head><body>
<h2>Bulk Import (CSV paste)</h2>
<label>Import Type</label>
<select id="type">
  <option value="swimmers">Swimmers</option>
  <option value="meets">Meets</option>
  <option value="pr">PR Baselines</option>
</select>

<div class="row">
  <div>
    <label>Options</label>
    <label class="small"><input id="hasHeader" type="checkbox" checked> First row is header</label>
    <div id="prOpts" style="display:none">
      <label>Default date for PRs (if missing per row)</label>
      <input id="defaultDate" type="date">
    </div>
  </div>
  <div>
    <label>Templates</label>
    <div id="tmpl" class="small"></div>
  </div>
</div>

<label>CSV Data</label>
<textarea id="csv" placeholder="Paste CSV here..."></textarea>

<button class="btn" onclick="go()">Import</button>
<div id="msg" class="small"></div>

<script>
const TMPL = {
  swimmers: "Name,Grad Year,Gender,Level,Notes\\nAlex Rivera,2027,M,Varsity,\\nTaylor Brooks,2028,F,JV,",
  meets: "Meet,Date,Location,Course,Notes,Has JV?\\nBoise City Classic,2025-11-10,City Pool,SCY,Non-conference,Yes",
  pr: "Swimmer,Event,Time,Date(optional)\\nAlex Rivera,100 Freestyle,55.12,2025-10-01\\nTaylor Brooks,200 IM,2:18.90,"
};
const typeSel = document.getElementById('type');
const tmplDiv = document.getElementById('tmpl');
const prOpts = document.getElementById('prOpts');

function renderTmpl(){
  const t = typeSel.value;
  prOpts.style.display = (t==='pr') ? 'block' : 'none';
  tmplDiv.innerHTML = '<pre>'+TMPL[t]+'</pre>';
}
typeSel.addEventListener('change', renderTmpl);
renderTmpl();

function go(){
  const payload = {
    type: typeSel.value,
    csv: document.getElementById('csv').value,
    hasHeader: document.getElementById('hasHeader').checked,
    defaultDate: document.getElementById('defaultDate').value
  };
  document.getElementById('msg').textContent = 'Importing...';
  google.script.run.withSuccessHandler(res=>{
    document.getElementById('msg').textContent = 'Done: '+res.inserted+' '+res.kind+' imported.';
  }).withFailureHandler(err=>{
    document.getElementById('msg').textContent = 'Error: '+err.message;
  }).bulkImport(payload);
}
</script>
</body></html>`;
  return tmpl;
}

// ===== Debug utilities =====
const DEBUG = true; // flip to false to silence sheet logging (console stays)
const AUTO_FILTERS = true; // global off-switch for programmatic filters

function debugLog_(step, msg, data) {
  const ts = new Date();
  const stamp = Utilities.formatDate(
    ts,
    Session.getScriptTimeZone(),
    'HH:mm:ss'
  );
  console.log(
    `[DEBUG ${stamp}] [${step}] ${msg} ${data ? JSON.stringify(data) : ''}`
  );
  if (!DEBUG) return;
  try {
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName('_Debug');
    if (!sh) {
      sh = ss.insertSheet('_Debug');
      sh.getRange(1, 1, 1, 4)
        .setValues([['Time', 'Step', 'Message', 'Data']])
        .setFontWeight('bold');
    }
    sh.appendRow([ts, step, msg, data ? JSON.stringify(data) : '']);
  } catch (_) {}
}

function withStep_(name, fn) {
  const t0 = Date.now();
  debugLog_(name, 'start');
  try {
    const res = fn();
    debugLog_(name, 'ok', { ms: Date.now() - t0 });
    return res;
  } catch (e) {
    debugLog_(name, 'ERROR', {
      ms: Date.now() - t0,
      err: String(e),
      stack: e && e.stack,
    });
    throw new Error(`${name}: ${e.message}`);
  }
}

// Filter state & cleanup (uses Advanced Sheets API if available)
function getFilterState_(sheet) {
  const hasBasic = !!(sheet.getFilter && sheet.getFilter());
  let viewCount = null;
  try {
    const ssId = SpreadsheetApp.getActive().getId();
    const meta = Sheets.Spreadsheets.get(ssId, {
      fields: 'sheets(properties.sheetId,filterViews(filterViewId))',
    });
    const me = (meta.sheets || []).find(
      s => s.properties && s.properties.sheetId === sheet.getSheetId()
    );
    viewCount = me && me.filterViews ? me.filterViews.length : 0;
  } catch (e) {
    viewCount = -1;
  } // -1 means Advanced Service unavailable
  return { hasBasic, viewCount };
}

function clearAllFilters_(sheet) {
  const ssId = SpreadsheetApp.getActive().getId();
  const sheetId = sheet.getSheetId();
  try {
    const meta = Sheets.Spreadsheets.get(ssId, {
      fields: 'sheets(properties.sheetId,filterViews(filterViewId))',
    });
    const me = (meta.sheets || []).find(
      s => s.properties && s.properties.sheetId === sheetId
    );
    const views = me && me.filterViews ? me.filterViews : [];
    const requests = [{ clearBasicFilter: { sheetId } }];
    for (const v of views)
      requests.push({ deleteFilterView: { filterId: v.filterViewId } });
    if (requests.length) Sheets.Spreadsheets.batchUpdate({ requests }, ssId);
    debugLog_('clearAllFilters_', 'cleared', {
      sheet: sheet.getName(),
      views: views.length,
    });
  } catch (e) {
    debugLog_('clearAllFilters_', 'skipped (no Advanced Service?)', {
      sheet: sheet.getName(),
      err: String(e),
    });
  }
}

// Always use this instead of range.createFilter()
function safeCreateFilter_(sheet, range, tag) {
  if (!autoFiltersEnabled_()) {
    debugLog_(
      'safeCreateFilter_',
      'AUTO_FILTERS=false (meet day mode); cleared filters',
      { sheet: sheet.getName(), tag }
    );
    try {
      clearAllFilters_(sheet);
    } catch (e) {
      console.log('Failed to clear filters in meet day mode:', e.message);
    }
    return;
  }
  try {
    clearAllFilters_(sheet);
  } catch (e) {
    console.log('Failed to clear all filters:', e.message);
  }
  try {
    const f = sheet.getFilter && sheet.getFilter();
    if (f) f.remove();
  } catch (e) {
    console.log('Failed to remove existing filter:', e.message);
  }
  try {
    range.createFilter();
    debugLog_('safeCreateFilter_', 'created', { sheet: sheet.getName(), tag });
  } catch (e) {
    const msg = String((e && e.message) || e);
    if (msg.indexOf('already has a filter') === -1) {
      debugLog_('safeCreateFilter_', 'ERROR', {
        sheet: sheet.getName(),
        tag,
        err: msg,
      });
      throw e;
    }
    debugLog_('safeCreateFilter_', 'skipped (already exists)', {
      sheet: sheet.getName(),
      tag,
    });
    SpreadsheetApp.getActive().toast(
      `Skipped filter on ${tag || sheet.getName()} (already exists)`,
      'Coach Tools',
      3
    );
  }
}

// Quick menu hook to dump filter state
function debugDumpFilters() {
  const ss = SpreadsheetApp.getActive();
  ['PR Summary', 'Lineup Check'].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const st = getFilterState_(sh);
    debugLog_('debugDumpFilters', name, st);
    SpreadsheetApp.getActive().toast(
      `${name}: basic=${st.hasBasic} views=${st.viewCount}`,
      'Coach Tools',
      5
    );
  });
}

/***** MEET DAY MODE *****/

// Persist per-spreadsheet
function isMeetDayModeOn_() {
  return (
    PropertiesService.getDocumentProperties().getProperty('MEET_DAY') === '1'
  );
}
function setMeetDayMode_(on) {
  PropertiesService.getDocumentProperties().setProperty(
    'MEET_DAY',
    on ? '1' : '0'
  );
  applyMeetDayModeEffects_(on);
  SpreadsheetApp.getActive().toast(
    `Meet Day Mode: ${on ? 'ON' : 'OFF'}`,
    'Coach Tools',
    5
  );
}
function toggleMeetDayMode() {
  setMeetDayMode_(!isMeetDayModeOn_());
}
function meetDayStatus() {
  const on = isMeetDayModeOn_();
  const id = ScriptApp.getScriptId ? ScriptApp.getScriptId() : '(lib)';
  SpreadsheetApp.getActive().toast(
    `CoachToolsCore • ${on ? 'MEET DAY ON' : 'Meet Day off'} • ${typeof LIB_VER !== 'undefined' ? LIB_VER : ''}`,
    'Coach Tools',
    6
  );
  return { on, version: typeof LIB_VER !== 'undefined' ? LIB_VER : '', id };
}

// Auto-filters should be OFF during meet day
function autoFiltersEnabled_() {
  return !isMeetDayModeOn_();
}

// Apply visual/UX changes for meet mode
function applyMeetDayModeEffects_(on) {
  const ss = SpreadsheetApp.getActive();

  // 1) Clear filters on key sheets and suppress future ones via autoFiltersEnabled_()
  ['PR Summary', 'Lineup Check'].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) {
      try {
        clearAllFilters_(sh);
      } catch (_) {}
    }
  });

  // 2) Lock admin-ish sheets during meet; unlock when off
  const toLock = ['Settings', 'Events', 'Meet Events', 'Results'];
  const tag = 'CoachTools: Meet Day Lock';
  if (on) {
    toLock.forEach(n => {
      const sh = ss.getSheetByName(n);
      if (!sh) return;
      // avoid duplicates
      const existing = sh
        .getProtections(SpreadsheetApp.ProtectionType.SHEET)
        .find(p => p.getDescription() === tag);
      if (!existing) {
        const p = sh.protect().setDescription(tag);
        try {
          p.removeEditors(p.getEditors());
        } catch (_) {}
      }
    });
  } else {
    ss.getSheets().forEach(sh => {
      sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => {
        if (p.getDescription() === tag) p.remove();
      });
    });
  }

  // 3) Hide admin sheets on meet day; unhide when off
  toLock.forEach(n => {
    const sh = ss.getSheetByName(n);
    if (!sh) return;
    try {
      sh.setHidden(on);
    } catch (e) {
      console.log(`Failed to ${on ? 'hide' : 'unhide'} sheet ${n}:`, e.message);
    }
  });

  // 4) Make Coach Packet extra legible
  const cp = ss.getSheetByName('Coach Packet');
  if (cp) {
    try {
      cp.setFrozenRows(3);
      cp.getRange('A1:E1')
        .setFontWeight('bold')
        .setFontSize(on ? 14 : 12);
      cp.getRange('A3:E')
        .setWrap(true)
        .setVerticalAlignment('middle')
        .setFontSize(on ? 12 : 10);
      // subtle borders for readability
      const lr = Math.max(3, cp.getLastRow());
      cp.getRange(3, 1, lr - 2, 5).setBorder(
        false,
        true,
        false,
        true,
        false,
        false,
        '#cccccc',
        SpreadsheetApp.BorderStyle.SOLID
      );
    } catch (_) {}
  }

  // 5) Add a tiny status hint on Meet Entry title
  const entry = ss.getSheetByName('Meet Entry');
  if (entry) {
    const v = entry.getRange('A1').getDisplayValue();
    const base = v.replace(/\s+—\s+MEET DAY.*$/, '');
    entry.getRange('A1').setValue(on ? `${base} — MEET DAY` : base);
  }
}

function buildBulkImportSidebar() {
  return HtmlService.createHtmlOutput(bulkImportSidebarHtml_()).setTitle(
    'Bulk Import'
  );
}

function buildAddResultSidebar() {
  return HtmlService.createHtmlOutput(addResultSidebarHtml_()).setTitle(
    'Add Result'
  );
}

function buildAddMeetSidebar() {
  return HtmlService.createHtmlOutput(addMeetSidebarHtml_()).setTitle(
    'Add Meet'
  );
}

function buildAddEventSidebar() {
  return HtmlService.createHtmlOutput(addEventSidebarHtml_()).setTitle(
    'Add Event'
  );
}
