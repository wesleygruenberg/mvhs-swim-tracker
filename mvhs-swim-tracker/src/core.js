/** =========================
 *  Coach Tools for MVHS Swim — v2.0
 *  Adds: Add Result sidebar, Add Meet sidebar, Add Event sidebar,
 *        Clone Clean Baseline (baseline events, no swimmers/meets)
 *  Keeps: Settings, JV toggle, PR Summary/Dashboard, presets, snapshots, usage checks, JV support, sample team
 *  =========================
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Coach Tools')
    .addItem('Refresh All (safe)', 'refreshAll')
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('Results')
        .addItem('Add Result (sidebar)', 'openAddResultSidebar')
        .addItem('Refresh PR Summary & Dashboard', 'refreshPRs')
    )
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('Admin')
        .addItem('Ensure Settings Sheet', 'ensureSettingsSheet')
        .addItem('Apply Limits from Settings', 'applyLimitsFromSettings')
        .addItem('Ensure JV Toggle on Meets', 'ensureMeetsHasJVColumn')
        .addItem('Enable JV/Varsity Support (add JV events + reseed)', 'enableJVSupport')
        .addItem('Clear Sample Data (Results & assignments)', 'adminClearSampleData')
        .addItem('Add Meet (sidebar)', 'openAddMeetSidebar')
        .addItem('Add Event (sidebar)', 'openAddEventSidebar')
    )
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('Clone')
        .addItem('Make Clean Copy (reset data)', 'cloneMakeCleanCopy')
        .addItem('New Season Copy (carry forward, drop seniors)', 'cloneNewSeasonCarryForward')
        .addItem('Clone Clean Baseline (baseline events, no meets/swimmers)', 'cloneCleanBaseline')
    )
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('Roster')
        .addItem('Add Swimmer + PRs (sidebar)', 'openAddSwimmerSidebar')
        .addItem('Generate Sample Team (50: 25F/25M, 10V/15JV each)', 'generateSampleTeam50')
    )
    .addSubMenu(
    ui.createMenu('Import')
      .addItem('Bulk Import (CSV paste)', 'openBulkImportSidebar')
    )
    .addItem('Ensure Meet Presets Table', 'ensureMeetEventsTemplate')
    .addItem('Apply Meet Presets to Lineup', 'applyMeetPresets')
    .addItem('Check Lineup (Usage & Violations)', 'checkLineup')
    .addItem('Create Snapshot of Current Lineup', 'createSnapshot')
    .addItem('Build Coach Packet (print view)', 'buildCoachPacket')
    .addToUi();
}

/** ---------- Refresh All ---------- */
function refreshAll() {
  try { const s = SpreadsheetApp.getActive().getSheetByName('PR Summary');    if (s && s.getFilter()) s.getFilter().remove(); } catch(e){}
  try { const s = SpreadsheetApp.getActive().getSheetByName('Lineup Check');  if (s && s.getFilter()) s.getFilter().remove(); } catch(e){}

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
  let set = ss.getSheetByName('Settings');
  if (!set) set = ss.insertSheet('Settings');
  if (set.getLastRow() < 2) {
    set.clear();
    const rows = [
      ['Settings', ''],
      ['', ''],
      ['Season Name',        '2025 HS'],
      ['Season Start Year',  new Date().getFullYear()],
      ['Drop Grad Year on New Season Copy', new Date().getFullYear()+1],
      ['', ''],
      ['Limits', ''],
      ['Max Individual Events', 2],
      ['Max Relay Events',      2],
      ['', ''],
      ['Notes', 'Change values in column B; Admin → Apply Limits pushes B8/B9 into Meet Entry.']
    ];
    set.getRange(1,1,rows.length,2).setValues(rows);
    set.getRange('A1').setFontWeight('bold').setFontSize(14);
    set.getRange('A7').setFontWeight('bold');
    set.setColumnWidths(1,2,240);
  }
  toast('Settings sheet verified.');
}
function readSettings_(ss) {
  const set = ss.getSheetByName('Settings');
  if (!set) return { seasonName:'Season', seasonYear:new Date().getFullYear(), dropGradYear:new Date().getFullYear()+1, maxInd:2, maxRel:2 };
  const getVal = (label) => { const f = set.createTextFinder(label).matchEntireCell(true).findNext(); return f ? set.getRange(f.getRow(),2).getValue() : null; };
  return {
    seasonName: String(getVal('Season Name') || 'Season'),
    seasonYear: Number(getVal('Season Start Year') || new Date().getFullYear()),
    dropGradYear: Number(getVal('Drop Grad Year on New Season Copy') || (new Date().getFullYear()+1)),
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
  const entry = mustSheet('Meet Entry');
  const sw = mustSheet('Swimmers');
  const me = mustSheet('Meets');
  const ev = mustSheet('Events');
  const results = mustSheet('Results');

  ensureSwimmersLevelColumn_();

  ss.setNamedRange('SwimmerNames', sw.getRange('A2:A'));
  ss.setNamedRange('MeetNames',    me.getRange('A2:A'));
  ss.setNamedRange('EventNames',   ev.getRange('A2:A'));

  const startRow = 6, last = 206;
  entry.getRange(`A${startRow}:A${last}`).insertCheckboxes();

  const dvMeet    = SpreadsheetApp.newDataValidation().requireValueInRange(ss.getRangeByName('MeetNames'), true).build();
  const dvSwimmer = SpreadsheetApp.newDataValidation().requireValueInRange(ss.getRangeByName('SwimmerNames'), true).build();
  const dvEvent   = SpreadsheetApp.newDataValidation().requireValueInRange(ss.getRangeByName('EventNames'), true).build();

  entry.getRange('B1').setDataValidation(dvMeet);
  entry.getRange(`H${startRow}:H${last}`).setDataValidation(dvSwimmer);
  entry.getRange(`I${startRow}:L${last}`).setDataValidation(dvSwimmer);

  const resLast = Math.max(1000, results.getLastRow()+200);
  results.getRange('A2:A' + resLast).setDataValidation(dvMeet);
  results.getRange('B2:B' + resLast).setDataValidation(dvEvent);
  results.getRange('C2:C' + resLast).setDataValidation(dvSwimmer);
  results.getRange('D2:E' + resLast).setNumberFormat('mm:ss.00');

  const rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND($H6<>"",COUNTIFS($H$6:$H$206,$H6,$C$6:$C$206,"Individual",$A$6:$A$206,TRUE)>$B$2)`)
    .setRanges([entry.getRange('H6:H206')]).setBackground('#F4CCCC').build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND($I6<>"",SUMPRODUCT(($I$6:$L$206=$I6)*($A$6:$A$206=TRUE))>$B$3)`)
    .setRanges([entry.getRange('I6:L206')]).setBackground('#F4CCCC').build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND(I6<>"",COUNTIF($I6:$L6,I6)>1)`)
    .setRanges([entry.getRange('I6:L206')]).setBackground('#FFE699').build());
  entry.setConditionalFormatRules(rules);

  toast('Validations & formatting refreshed.');
}

function ensureMeetEventsTemplate() {
  const ss = SpreadsheetApp.getActive();
  const me = mustSheet('Meets');
  const ev = mustSheet('Events');
  const out = ss.getSheetByName('Meet Events') || ss.insertSheet('Meet Events');

  if (out.getLastRow() < 1) {
    out.getRange(1,1,1,4).setValues([['Meet','Event','Active?','Notes']]).setFontWeight('bold');
  }

  const last = out.getLastRow();
  const existing = new Set();
  const data = (last >= 2) ? out.getRange(2,1,last-1,2).getValues() : [];
  for (const [m,e] of data) if (m && e) existing.add(m + '|' + e);

  const meets = getColValues(me, 1, 2);
  const evLast = ev.getLastRow();
  const evRows = (evLast >= 2) ? ev.getRange(2,1,evLast-1,5).getValues() : []; // Event,Type,Dist,Stroke,DefaultActive

  const rowsToAppend = [];
  for (const m of meets) {
    for (const r of evRows) {
      const [ename,, , , defActive] = r;
      if (!ename) continue;
      const key = m + '|' + ename;
      if (!existing.has(key)) {
        rowsToAppend.push([m, ename, !!defActive, '']);
        existing.add(key);
      }
    }
  }
  if (rowsToAppend.length > 0) {
    out.getRange(out.getLastRow()+1, 1, rowsToAppend.length, 4).setValues(rowsToAppend);
  }

  out.autoResizeColumns(1,4);
  toast('Meet Events table is ready.');
}

function ensureMeetsHasJVColumn() {
  const me = mustSheet('Meets');
  const headers = me.getRange(1,1,1,me.getLastColumn() || 1).getValues()[0].map(h => String(h||'').trim());
  let col = headers.findIndex(h => h.toLowerCase() === 'has jv?') + 1;
  if (!col) {
    col = me.getLastColumn() + 1;
    me.getRange(1, col).setValue('Has JV?').setFontWeight('bold');
  }
  const startRow = 2, endRow = Math.max(me.getLastRow(), 100);
  me.getRange(startRow, col, endRow-startRow+1, 1).insertCheckboxes();
  const last = me.getLastRow();
  if (last >= startRow) {
    const rng = me.getRange(startRow, col, last-startRow+1, 1);
    const vals = rng.getValues().map(r => [r[0] === '' ? true : r[0]]);
    rng.setValues(vals);
  }
}
function getMeetHasJV_(meetName) {
  if (!meetName) return true;
  const me = mustSheet('Meets');
  const last = me.getLastRow();
  if (last < 2) return true;
  const headers = me.getRange(1,1,1,me.getLastColumn()).getValues()[0].map(h => String(h||'').trim());
  let jvCol = headers.findIndex(h => h.toLowerCase() === 'has jv?') + 1;
  if (!jvCol) return true;
  const meets = me.getRange(2,1,last-1,1).getValues().map(r=>String(r[0]||'').trim());
  const idx = meets.findIndex(m => m === meetName);
  if (idx < 0) return true;
  const val = me.getRange(2+idx, jvCol).getValue();
  return val === '' ? true : !!val;
}
function setPresetsJVForMeet_(meetName, hasJV) {
  const presets = mustSheet('Meet Events');
  const last = presets.getLastRow();
  if (last < 2) return;
  const rows = presets.getRange(2,1,last-1,3).getValues();
  let touched = 0;
  for (let i=0;i<rows.length;i++) {
    const [m, ev] = rows[i];
    if (m === meetName && /\(JV\)\s*$/.test(String(ev||''))) {
      if (!hasJV && rows[i][2] !== false) { rows[i][2] = false; touched++; }
    }
  }
  if (touched) presets.getRange(2,1,last-1,3).setValues(rows);
}

function applyMeetPresets() {
  const ss = SpreadsheetApp.getActive();
  const entry = mustSheet('Meet Entry');
  const presets = mustSheet('Meet Events');

  ensureMeetsHasJVColumn();

  const selected = (entry.getRange('B1').getDisplayValue() || '').trim();
  if (!selected) return toast('Pick a meet in B1 first.');

  const hasJV = getMeetHasJV_(selected);
  setPresetsJVForMeet_(selected, hasJV);

  const pLast = presets.getLastRow();
  const pVals = (pLast >= 2) ? presets.getRange(2,1,pLast-1,3).getValues() : [];
  const map = new Map();
  for (const [meet, ev, active] of pVals) {
    if (meet === selected && ev) map.set(ev, !!active);
  }

  const startRow = 6;
  const lastRow = findLastDataRow(entry, 2, startRow);
  for (let r = startRow; r <= lastRow; r++) {
    const evName = entry.getRange(r,2).getDisplayValue();
    if (!evName) continue;
    let active = map.has(evName) ? map.get(evName) : true;
    if (!hasJV && /\(JV\)\s*$/.test(evName)) active = false;
    entry.getRange(r,1).setValue(active);
  }

  toast(`Applied presets for "${selected}" (${hasJV ? 'JV enabled' : 'JV disabled'}).`);
}

function checkLineup() {
  const ss = SpreadsheetApp.getActive();
  const entry = mustSheet('Meet Entry');
  const sw = mustSheet('Swimmers');
  const out = ss.getSheetByName('Lineup Check') || ss.insertSheet('Lineup Check');
  out.clear();

  const maxInd = Number(entry.getRange('B2').getValue() || 2);
  const maxRel = Number(entry.getRange('B3').getValue() || 2);
  const swimmers = getColValues(sw,1,2);

  const levelCol = findHeaderColumn_(sw, 'Level');
  const nameLevel = {};
  if (levelCol) {
    const last = sw.getLastRow();
    const names = (last>=2) ? sw.getRange(2,1,last-1,1).getValues() : [];
    const levels= (last>=2) ? sw.getRange(2,levelCol,last-1,1).getValues() : [];
    for (let i=0;i<names.length;i++) if (names[i][0]) nameLevel[names[i][0]] = String(levels[i][0]||'').trim();
  }

  const startRow = 6;
  const lastRow = findLastDataRow(entry, 2, startRow);

  const indiv = Object.fromEntries(swimmers.map(s=>[s,0]));
  const relay = Object.fromEntries(swimmers.map(s=>[s,0]));
  const dupViolations = [];
  const assignViolations = [];
  const jvMismatch = [];

  for (let r = startRow; r <= lastRow; r++) {
    const active = entry.getRange(r,1).getValue() === true;
    const evName = entry.getRange(r,2).getDisplayValue();
    const type   = entry.getRange(r,3).getDisplayValue();
    if (!active || !evName) continue;

    const isJVEvent = /\(JV\)\s*$/.test(evName);

    if (type === 'Individual') {
      const name = entry.getRange(r,8).getDisplayValue().trim();
      if (name) {
        indiv[name] = (indiv[name]||0) + 1;
        if (isJVEvent && (nameLevel[name]||'').toLowerCase()==='varsity') {
          jvMismatch.push([r, evName, name]);
        }
      }
    } else if (type === 'Relay') {
      const names = entry.getRange(r,9,1,4).getDisplayValues()[0].map(x=>x.trim()).filter(Boolean);
      const dups = findDuplicates(names);
      if (dups.length) dupViolations.push([r, evName, dups.join(', ')]);
      for (const n of names) {
        relay[n] = (relay[n]||0) + 1;
        if (isJVEvent && (nameLevel[n]||'').toLowerCase()==='varsity') {
          jvMismatch.push([r, evName, n]);
        }
      }
    }
  }

  const header = [['Swimmer','Individual','Relay','Limit (Ind)','Limit (Rel)','Status']];
  const rows = [];
  for (const s of swimmers) {
    if (!s) continue;
    const i = indiv[s] || 0;
    const r = relay[s] || 0;
    const status = (i>maxInd || r>maxRel) ? 'OVER' : 'OK';
    rows.push([s,i,r,maxInd,maxRel,status]);
    if (i>maxInd) assignViolations.push(['', '', s, 'Individual', i]);
    if (r>maxRel) assignViolations.push(['', '', s, 'Relay', r]);
  }
  rows.sort((a,b)=> a[0].localeCompare(b[0]));

  out.getRange(1,1,1,6).setValues(header).setFontWeight('bold');
  if (rows.length) out.getRange(2,1,rows.length,6).setValues(rows);
  safeCreateFilter_(out, out.getRange(1,1,Math.max(2,rows.length+1),6));

  out.autoResizeColumns(1,6);

  const rng = out.getRange(2,6,Math.max(rows.length,1),1);
  const rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('OVER').setBackground('#F4CCCC').setRanges([rng]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('OK').setBackground('#D9EAD3').setRanges([rng]).build());
  out.setConditionalFormatRules(rules);

  let row = rows.length + 3;
  out.getRange(row,1,1,3).setValues([['Duplicate swimmers in a relay row','(Row)','(Event)']]).setFontWeight('bold');
  row++;
  if (dupViolations.length) {
    out.getRange(row,1,dupViolations.length,3).setValues(dupViolations);
    row += dupViolations.length + 1;
  } else { out.getRange(row,1).setValue('None'); row += 2; }

  out.getRange(row,1,1,5).setValues([['Assignments over limits','(Row)','(Event)','(Type)','(Count)']]).setFontWeight('bold');
  row++;
  if (assignViolations.length) {
    out.getRange(row,1,assignViolations.length,5).setValues(assignViolations);
    row += assignViolations.length + 1;
  } else { out.getRange(row,1).setValue('None'); row += 2; }

  out.getRange(row,1,1,3).setValues([['JV/VARSITY mismatches (Varsity swimmers in JV events)','(Row)','(Event)']]).setFontWeight('bold');
  row++;
  if (jvMismatch.length) {
    out.getRange(row,1,jvMismatch.length,2).setValues(jvMismatch.map(x=>[x[0],x[1]]));
  } else {
    out.getRange(row,1).setValue('None');
  }

  toast('Lineup Check generated.');
}

function createSnapshot() {
  const ss = SpreadsheetApp.getActive();
  const src = mustSheet('Meet Entry');
  const meet = src.getRange('B1').getDisplayValue() || 'Unspecified Meet';
  const stamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HHmm");
  const name = `Lineup — ${meet} — ${stamp}`;
  const snap = src.copyTo(ss).setName(name);
  snap.getDataRange().copyTo(snap.getDataRange(), {contentsOnly:true});
  toast(`Snapshot saved: ${name}`);
}

function refreshPRs() {
  createPRSummary();
  createSwimmerDashboard();
  toast('PR Summary & Dashboard refreshed.');
}
function createPRSummary() {
  const ss = SpreadsheetApp.getActive();
  const results = mustSheet('Results');
  const out = ss.getSheetByName('PR Summary') || ss.insertSheet('PR Summary');
  out.clear();
  const last = results.getLastRow();
  if (last < 2) { out.getRange(1,1).setValue('No results yet.'); return; }
  const vals = results.getRange(2,1,last-1,10).getValues();
  const best = new Map(), latest = new Map();
  for (const [meet, ev, sw, , fin, , , date] of vals) {
    if (!sw || !ev || fin === "" || fin == null) continue;
    const key = sw + '|' + ev, t = fin;
    const b = best.get(key); if (!b) best.set(key, {time:t, meet, date, count:1});
    else { b.count++; if (t < b.time) { b.time=t; b.meet=meet; b.date=date; } }
    const L = latest.get(key); if (!L || (date && date > L.date)) latest.set(key, {time:t, date, meet});
  }
  const rows = [];
  for (const [k,v] of best.entries()) {
    const [sw,ev] = k.split('|'); const L = latest.get(k);
    rows.push([sw, ev, v.time, v.meet||"", v.date||"", v.count, L?L.time:""]);
  }
  rows.sort((a,b)=> a[0].localeCompare(b[0]) || a[1].localeCompare(b[1]));
  const header = ['Swimmer','Event','PR Time','PR Meet','PR Date','Races','Last Swim','Δ vs PR'];
  out.getRange(1,1,1,header.length).setValues([header]).setFontWeight('bold');
  if (rows.length) {
    out.getRange(2,1,rows.length,7).setValues(rows);
    out.getRange(2,8,rows.length,1).setFormulaR1C1('=IF(AND(RC[-1]<>"",RC[-5]<>""),RC[-1]-RC[-5],"")');
  }
  out.setFrozenRows(1);
  safeCreateFilter_(out, out.getRange(1,1,Math.max(2,rows.length+1),8));

  out.getRange('C2:C').setNumberFormat('mm:ss.00');
  out.getRange('G2:G').setNumberFormat('mm:ss.00');
  out.getRange('H2:H').setNumberFormat('[m]:ss.00');
  out.autoResizeColumns(1,8);
}
function createSwimmerDashboard() {
  const ss = SpreadsheetApp.getActive();
  const prs = ss.getSheetByName('PR Summary') || createPRSummary();
  const sw = mustSheet('Swimmers');
  ss.setNamedRange('SwimmerNames', sw.getRange('A2:A'));
  let dash = ss.getSheetByName('Swimmer Dashboard') || ss.insertSheet('Swimmer Dashboard');
  dash.clear();
  dash.getRange('A1').setValue('Swimmer Dashboard').setFontWeight('bold').setFontSize(14);
  dash.getRange('A3').setValue('Swimmer:').setFontWeight('bold');
  dash.getRange('B3').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(ss.getRangeByName('SwimmerNames'), true).build()
  );
  const headers = ['Event','PR Time','PR Meet','PR Date','Races','Last Swim','Δ vs PR'];
  dash.getRange('A5:G5').setValues([headers]).setFontWeight('bold');
  dash.getRange('A6').setFormula(`
=IF(B3="","",
  QUERY('PR Summary'!A2:H,
    "select B,C,D,E,F,G,H where A = '" & B3 & "' order by B",
    0
  )
)`.trim());
  dash.getRange('B6:B').setNumberFormat('mm:ss.00');
  dash.getRange('G6:G').setNumberFormat('[m]:ss.00');
  dash.getRange('F6:F').setNumberFormat('mm:ss.00');
  dash.setFrozenRows(5);
  dash.autoResizeColumns(1,7);
}

function buildCoachPacket() {
  const ss = SpreadsheetApp.getActive();
  const entry = mustSheet('Meet Entry');
  const cp = ss.getSheetByName('Coach Packet') || ss.insertSheet('Coach Packet');
  cp.clear();

  const meet = entry.getRange('B1').getDisplayValue() || 'Unspecified Meet';
  cp.getRange('A1').setValue(`Coach Packet — ${meet}`).setFontWeight('bold').setFontSize(14);

  const startRow = 6;
  const lastRow = findLastDataRow(entry, 2, startRow);
  const rows = [['Event','Type','Heat','Lane','Individual / Relay Legs']];
  for (let r = startRow; r <= lastRow; r++) {
    const active = entry.getRange(r,1).getValue() === true;
    if (!active) continue;
    const ev = entry.getRange(r,2).getDisplayValue();
    const type = entry.getRange(r,3).getDisplayValue();
    const heat = entry.getRange(r,6).getDisplayValue();
    const lane = entry.getRange(r,7).getDisplayValue();
    if (type === 'Individual') {
      const n = entry.getRange(r,8).getDisplayValue();
      rows.push([ev, type, heat, lane, n || '—']);
    } else {
      const legs = entry.getRange(r,9,1,4).getDisplayValues()[0].filter(Boolean).join(' • ');
      rows.push([ev, type, heat, lane, legs || '—']);
    }
  }
  if (rows.length === 1) rows.push(['(no active events)', '', '', '', '']);

  cp.getRange(3,1,rows.length,5).setValues(rows);
  cp.getRange(3,1,1,5).setFontWeight('bold');
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
 * ADMIN & ROSTER + JV SUPPORT
 * ========================= */
function ensureSwimmersLevelColumn_() {
  const sw = mustSheet('Swimmers');
  const headers = sw.getRange(1,1,1,Math.max(sw.getLastColumn(),5)).getValues()[0];
  const norm = headers.map(h => String(h||'').trim().toLowerCase());
  if (!norm.includes('level')) {
    sw.insertColumnAfter(3); // D
    sw.getRange(1,4).setValue('Level').setFontWeight('bold');
    if (!sw.getRange(1,5).getValue()) sw.getRange(1,5).setValue('Notes').setFontWeight('bold');
  }
  // Ensure base headers exist
  sw.getRange(1,1,1,5).setValues([['Name','Grad Year','Gender','Level','Notes']]).setFontWeight('bold');
}

function adminClearSampleData() {
  const ss = SpreadsheetApp.getActive();
  const results = mustSheet('Results');
  const entry = mustSheet('Meet Entry');
  const rLast = results.getLastRow();
  if (rLast >= 2) results.getRange(2,1,rLast-1,10).clearContent();
  reseedMeetEntryFromEvents_();
  entry.getRange('B1').setValue('');
  toast('Sample data cleared (Results & assignments). Meet Entry reseeded from Events.');
}

function generateSampleTeam50() {
  const ss = SpreadsheetApp.getActive();
  const sw = mustSheet('Swimmers');
  ensureSwimmersLevelColumn_();
  const last = sw.getLastRow();
  if (last >= 2) sw.getRange(2,1,last-1,sw.getLastColumn()).clearContent();

  const firstF = ['Avery','Riley','Jordan','Taylor','Casey','Parker','Quinn','Rowan','Emerson','Hayden','Morgan','Reese','Skyler','Alex','Drew','Logan','Cameron','Charlie','Harper','Kendall','Sage','Blake','Finley','Sydney','Payton'];
  const firstM = ['Liam','Noah','Oliver','Elijah','James','Benjamin','Lucas','Henry','Alexander','Mason','Michael','Ethan','Daniel','Jacob','Logan','Jackson','Levi','Sebastian','Mateo','Jack','Owen','Theodore','Aiden','Samuel','Joseph'];
  const lastNames = ['Brooks','Carter','Diaz','Ellis','Foster','Garcia','Hayes','Ingram','Jensen','Kim','Lopez','Miller','Nguyen','Ortiz','Patel','Quintero','Rivera','Shaw','Turner','Underwood','Vargas','Walker','Xu','Young','Zimmerman'];

  const year = readSettings_(ss).seasonYear || new Date().getFullYear();
  const grads = [year+1, year+2, year+3, year+4];
  const rows = [];
  function pick(pool, n){ const a=pool.slice(); for(let i=a.length-1;i>0;i--){const j=Math.floor(Math.random()*(i+1)); [a[i],a[j]]=[a[j],a[i]];} return a.slice(0,n); }
  const fNames = pick(firstF,25), mNames = pick(firstM,25), lNames = pick(lastNames,50);

  for (let i=0;i<25;i++) rows.push([`${fNames[i]} ${lNames[i]}`, grads[i%4], 'F', (i<10?'Varsity':'JV'), '']);
  for (let i=0;i<25;i++) rows.push([`${mNames[i]} ${lNames[25+i]}`, grads[(i+1)%4], 'M', (i<10?'Varsity':'JV'), '']);

  sw.getRange(2,1,rows.length,5).setValues(rows);
  toast('Sample team generated: 50 swimmers (25F/25M; 10 Varsity + 15 JV per gender).');
}

function enableJVSupport() {
  const ss = SpreadsheetApp.getActive();
  const ev = mustSheet('Events');
  const last = ev.getLastRow();
  if (last < 2) throw new Error('Events sheet is empty.');
  const rows = ev.getRange(2,1,last-1,5).getValues();
  const existing = new Set(rows.map(r => r[0]));
  const toAppend = [];
  for (const r of rows) {
    const name = String(r[0]||'');
    if (!name || /\(JV\)\s*$/.test(name)) continue;
    const jvName = `${name} (JV)`;
    if (!existing.has(jvName)) { const copy = r.slice(); copy[0] = jvName; toAppend.push(copy); }
  }
  if (toAppend.length) ev.getRange(ev.getLastRow()+1,1,toAppend.length,5).setValues(toAppend);
  reseedMeetEntryFromEvents_();
  ensureMeetEventsTemplate();
  applyMeetPresets();
  toast('JV support enabled: JV event variants added, Meet Entry reseeded, presets refreshed.');
}

function reseedMeetEntryFromEvents_() {
  const entry = mustSheet('Meet Entry');
  const ev = mustSheet('Events');
  const lastEntry = entry.getLastRow();
  if (lastEntry > 5) entry.getRange(6,1,lastEntry-5,12).clear({contentsOnly: true});
  const rows = ev.getRange(2,1,Math.max(ev.getLastRow()-1,0),5).getValues();
  let r = 6;
  for (const e of rows) {
    const [name, type, dist, stroke, defActive] = e;
    if (!name) continue;
    entry.getRange(r,1).setValue(!!defActive);
    entry.getRange(r,2).setValue(name);
    entry.getRange(r,3).setValue(type);
    entry.getRange(r,4).setValue(dist);
    entry.getRange(r,5).setValue(stroke);
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
  resetDataInCopy_(tgt, { carryForward: false, dropGradYear: settings.dropGradYear });
  toast(`Clean copy created.\nURL: ${copy.getUrl()}`);
}

function cloneNewSeasonCarryForward() {
  const src = SpreadsheetApp.getActive();
  ensureSettingsSheet();
  const settings = readSettings_(src);
  const nextSeasonName = `${(settings.seasonName || 'Season')} NEXT`;
  const newName = `${nextSeasonName} — NEW SEASON — ${timestamp_()}`;
  const file = DriveApp.getFileById(src.getId());
  const copy = file.makeCopy(newName);
  const tgt = SpreadsheetApp.openById(copy.getId());
  resetDataInCopy_(tgt, { carryForward: true, dropGradYear: settings.dropGradYear });
  const set = tgt.getSheetByName('Settings');
  if (set) {
    const finder = set.createTextFinder('Season Start Year').matchEntireCell(true).findNext();
    if (finder) {
      const row = finder.getRow();
      const cur = Number(set.getRange(row,2).getValue() || settings.seasonYear);
      set.getRange(row,2).setValue(cur + 1);
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
  sw.getRange(1,1,1,5).setValues([['Name','Grad Year','Gender','Level','Notes']]).setFontWeight('bold');

  // Meets -> headers + Has JV? column, no rows
  const me = tgt.getSheetByName('Meets') || tgt.insertSheet('Meets');
  me.clear();
  me.getRange(1,1,1,5).setValues([['Meet','Date','Location','Course','Season/Notes']]).setFontWeight('bold');
  tgt.setActiveSheet(me); ensureMeetsHasJVColumn(); // adds the Has JV? column

  // Events -> baseline set (no JV)
  const ev = tgt.getSheetByName('Events') || tgt.insertSheet('Events');
  ev.clear();
  ev.getRange(1,1,1,5).setValues([['Event','Type','Distance','Stroke','Default Active?']]).setFontWeight('bold');
  const baseline = [
    ['200 Medley Relay','Relay',200,'Medley',true],
    ['200 Freestyle','Individual',200,'Freestyle',true],
    ['200 Individual Medley','Individual',200,'IM',true],
    ['50 Freestyle','Individual',50,'Freestyle',true],
    ['100 Butterfly','Individual',100,'Butterfly',true],
    ['100 Freestyle','Individual',100,'Freestyle',true],
    ['500 Freestyle','Individual',500,'Freestyle',true],
    ['200 Freestyle Relay','Relay',200,'Freestyle',true],
    ['100 Backstroke','Individual',100,'Backstroke',true],
    ['100 Breaststroke','Individual',100,'Breaststroke',true],
    ['400 Freestyle Relay','Relay',400,'Freestyle',true],
    // extras default OFF
    ['200 Backstroke','Individual',200,'Backstroke',false],
    ['200 Breaststroke','Individual',200,'Breaststroke',false],
    ['200 Butterfly','Individual',200,'Butterfly',false],
    ['400 Individual Medley','Individual',400,'IM',false],
    ['50 Butterfly','Individual',50,'Butterfly',false],
    ['50 Backstroke','Individual',50,'Backstroke',false],
    ['50 Breaststroke','Individual',50,'Breaststroke',false],
  ];
  if (baseline.length) ev.getRange(2,1,baseline.length,5).setValues(baseline);

  // Results -> header only
  const res = tgt.getSheetByName('Results') || tgt.insertSheet('Results');
  res.clear();
  res.getRange(1,1,1,10).setValues([['Meet','Event','Swimmer','Seed Time (mm:ss.00)','Final Time (mm:ss.00)','Place','Notes','Date','Is PR?','Current PR']]).setFontWeight('bold');

  // Meet Entry -> reseed
  const entry = tgt.getSheetByName('Meet Entry') || tgt.insertSheet('Meet Entry');
  // If sheet exists, keep top rows (labels) and reseed; else you may want to copy from source—here we do a minimal rebuild:
  entry.clear();
  entry.getRange(1,1).setValue('Selected Meet').setFontWeight('bold');
  entry.getRange(1,2).setValue('');
  entry.getRange(2,1).setValue('Max Individual Events per Swimmer').setFontWeight('bold');
  entry.getRange(2,2).setValue(2);
  entry.getRange(3,1).setValue('Max Relay Events per Swimmer').setFontWeight('bold');
  entry.getRange(3,2).setValue(2);
  entry.getRange(4,1,1,12).setValues([['Active?','Event','Type','Distance','Stroke','Heat','Lane','Swimmer (Individual)','Relay Leg 1','Relay Leg 2','Relay Leg 3','Relay Leg 4']]).setFontWeight('bold');
  tgt.setActiveSheet(ev); // reseed uses active file's Events
  SpreadsheetApp.setActiveSpreadsheet(tgt);
  reseedMeetEntryFromEvents_();

  // Meet Events -> just header
  const presets = tgt.getSheetByName('Meet Events') || tgt.insertSheet('Meet Events');
  presets.clear();
  presets.getRange(1,1,1,4).setValues([['Meet','Event','Active?','Notes']]).setFontWeight('bold');

  // Derived views -> remove; will rebuild on demand
  ['PR Summary','Swimmer Dashboard','Lineup Check','Coach Packet'].forEach(n => { const sh=tgt.getSheetByName(n); if (sh) tgt.deleteSheet(sh); });
  // Snapshots
  tgt.getSheets().forEach(sh => { if (sh.getName().startsWith('Lineup — ')) tgt.deleteSheet(sh); });

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
      setCopy.getRange(1,1,rng.getNumRows(),rng.getNumColumns()).setValues(rng.getValues());
    } else {
      ss.insertSheet('Settings');
    }
  }
  const results = ss.getSheetByName('Results'); if (results) { const last = results.getLastRow(); if (last >= 2) results.getRange(2,1,last-1,10).clearContent(); }
  const entry = ss.getSheetByName('Meet Entry');
  const events = ss.getSheetByName('Events');
  if (entry && events) {
    entry.getRange('B1').setValue('');
    const startRow = 6;
    const lastRow = findLastDataRow(entry, 2, startRow);
    const evMap = new Map();
    const er = events.getLastRow();
    const eRows = (er>=2) ? events.getRange(2,1,er-1,5).getValues() : [];
    for (const r of eRows) evMap.set(r[0], !!r[4]);
    for (let r = startRow; r <= lastRow; r++) {
      const ev = entry.getRange(r,2).getDisplayValue();
      entry.getRange(r,1).setValue(evMap.has(ev) ? evMap.get(ev) : true);
      entry.getRange(r,6,1,7).clearContent();
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
        const vals = sw.getRange(2,1,last-1,4).getValues();
        const kept = vals.filter(r => Number(r[1]) !== Number(dropGradYear));
        sw.getRange(2,1,last-1,4).clearContent();
        if (kept.length) sw.getRange(2,1,kept.length,4).setValues(kept);
      }
    }
  }
  ['PR Summary','Swimmer Dashboard','Lineup Check','Coach Packet'].forEach(n => { const sh = ss.getSheetByName(n); if (sh) ss.deleteSheet(sh); });
  ss.getSheets().forEach(sh => { const name = sh.getName(); if (name.startsWith('Lineup — ')) ss.deleteSheet(sh); });
  setupValidationsFor_(ss);
  ensureMeetEventsTemplateFor_(ss);
}

/** Parametric helpers for copies */
function setupValidationsFor_(ss) {
  const entry = _mustSheet(ss,'Meet Entry');
  const sw = _mustSheet(ss,'Swimmers');
  const me = _mustSheet(ss,'Meets');
  const ev = _mustSheet(ss,'Events');
  const results = _mustSheet(ss,'Results');

  ss.setNamedRange('SwimmerNames', sw.getRange('A2:A'));
  ss.setNamedRange('MeetNames',    me.getRange('A2:A'));
  ss.setNamedRange('EventNames',   ev.getRange('A2:A'));

  const startRow = 6, last = 206;
  entry.getRange(`A${startRow}:A${last}`).insertCheckboxes();

  const dvMeet    = SpreadsheetApp.newDataValidation().requireValueInRange(ss.getRangeByName('MeetNames'), true).build();
  const dvSwimmer = SpreadsheetApp.newDataValidation().requireValueInRange(ss.getRangeByName('SwimmerNames'), true).build();
  entry.getRange('B1').setDataValidation(dvMeet);
  entry.getRange(`H${startRow}:H${last}`).setDataValidation(dvSwimmer);
  entry.getRange(`I${startRow}:L${last}`).setDataValidation(dvSwimmer);

  const resLast = Math.max(1000, results.getLastRow()+200);
  const dvEvent   = SpreadsheetApp.newDataValidation().requireValueInRange(ss.getRangeByName('EventNames'), true).build();
  results.getRange('A2:A' + resLast).setDataValidation(dvMeet);
  results.getRange('B2:B' + resLast).setDataValidation(dvEvent);
  results.getRange('C2:C' + resLast).setDataValidation(dvSwimmer);
  results.getRange('D2:E' + resLast).setNumberFormat('mm:ss.00');
}
function ensureMeetEventsTemplateFor_(ss) {
  const me = _mustSheet(ss,'Meets');
  const ev = _mustSheet(ss,'Events');
  const out = ss.getSheetByName('Meet Events') || ss.insertSheet('Meet Events');
  if (out.getLastRow() < 1) out.getRange(1,1,1,4).setValues([['Meet','Event','Active?','Notes']]).setFontWeight('bold');
  const last = out.getLastRow();
  const existing = new Set();
  const data = (last >= 2) ? out.getRange(2,1,last-1,2).getValues() : [];
  for (const [m,e] of data) if (m && e) existing.add(m + '|' + e);
  const meets = _getColValues(me, 1, 2);
  const evLast = ev.getLastRow();
  const evRows = (evLast >= 2) ? ev.getRange(2,1,evLast-1,5).getValues() : [];
  const rowsToAppend = [];
  for (const m of meets) for (const r of evRows) {
    const [ename,, , , defActive] = r; if (!ename) continue;
    const key = m + '|' + ename;
    if (!existing.has(key)) { rowsToAppend.push([m, ename, !!defActive, '']); existing.add(key); }
  }
  if (rowsToAppend.length) out.getRange(out.getLastRow()+1, 1, rowsToAppend.length, 4).setValues(rowsToAppend);
}

/** =========================
 * ROSTER: Add Swimmer + PRs (existing)
 * ========================= */
function openAddSwimmerSidebar() {
  const html = HtmlService.createHtmlOutput(addSwimmerSidebarHtml_()).setTitle('Add Swimmer + PRs');
  SpreadsheetApp.getUi().showSidebar(html);
}
function getIndividualEventsForPR_() {
  const ev = mustSheet('Events');
  const last = ev.getLastRow();
  if (last < 2) return [];
  const vals = ev.getRange(2,1,last-1,2).getValues();
  return vals.map(r => ({name:String(r[0]||''), type:String(r[1]||'')}))
             .filter(x => x.type === 'Individual' && x.name && !/\(JV\)\s*$/.test(x.name))
             .map(x => x.name);
}
function addSwimmerWithPRs(payload) {
  const ss = SpreadsheetApp.getActive();
  const sw = mustSheet('Swimmers');
  const results = mustSheet('Results');
  ensureSwimmersLevelColumn_();

  const name = String(payload.name||'').trim();
  if (!name) throw new Error('Name is required.');
  const grad = Number(payload.gradYear||'');
  const gender = String(payload.gender||'').trim() || '';
  const level = String(payload.level||'').trim() || '';
  const date  = payload.date ? new Date(payload.date) : new Date();
  const prs   = payload.prs || {};

  const last = sw.getLastRow();
  let rowIdx = -1;
  if (last >= 2) {
    const names = sw.getRange(2,1,last-1,1).getValues().map(r=>String(r[0]||''));
    rowIdx = names.findIndex(n => n === name);
  }
  const levelCol = findHeaderColumn_(sw, 'Level') || 4;
  if (rowIdx >= 0) {
    const r = 2 + rowIdx;
    if (grad)   sw.getRange(r,2).setValue(grad);
    if (gender) sw.getRange(r,3).setValue(gender);
    if (level)  sw.getRange(r,levelCol).setValue(level);
  } else {
    sw.getRange(sw.getLastRow()+1,1,1,5).setValues([[name, grad||'', gender||'', level||'', '']]);
  }

  const rows = [];
  const meetLabel = 'PR Baseline';
  for (const [evName, tStr] of Object.entries(prs)) {
    const serial = parseTimeSerial_(tStr);
    if (serial == null) continue;
    rows.push([meetLabel, evName, name, '', serial, '', 'Added via sidebar', date]);
  }
  if (rows.length) {
    const startRow = results.getLastRow() + 1;
    results.getRange(startRow,1,rows.length,8).setValues(rows);
  }
  try { setupValidations(); } catch(e) {}
  try { refreshPRs(); } catch(e) {}

  return {added: rowIdx < 0, prCount: rows.length};
}
function parseTimeSerial_(s) {
  if (s == null) return null;
  s = String(s).trim();
  if (!s) return null;
  let m = s.match(/^(\d+):(\d{1,2})(?:\.(\d+))?$/);
  if (m) {
    const minutes = parseInt(m[1],10);
    const seconds = parseInt(m[2],10) + (m[3] ? parseFloat('0.'+m[3]) : 0);
    const total = minutes*60 + seconds;
    return total/86400;
  }
  m = s.match(/^(\d+(?:\.\d+)?)$/);
  if (m) return parseFloat(m[1])/86400;
  return null;
}
function findHeaderColumn_(sheet, headerText) {
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(h => String(h||'').trim().toLowerCase());
  const idx = headers.indexOf(String(headerText).trim().toLowerCase());
  return idx >= 0 ? (idx+1) : 0;
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
  const html = HtmlService.createHtmlOutput(addResultSidebarHtml_()).setTitle('Add Result');
  SpreadsheetApp.getUi().showSidebar(html);
}
function listMeetNames_(){ return getColValues(mustSheet('Meets'),1,2); }
function listSwimmerNames_(){ return getColValues(mustSheet('Swimmers'),1,2); }
function listEventNames_(){ return getColValues(mustSheet('Events'),1,2); }
function listActiveEventsForMeet_(meet) {
  if (!meet) return listEventNames_();
  const presets = mustSheet('Meet Events');
  const last = presets.getLastRow(); if (last < 2) return listEventNames_();
  const vals = presets.getRange(2,1,last-1,3).getValues();
  const set = [];
  for (const [m,e,active] of vals) { if (m===meet && !!active && e) set.push(e); }
  return set.length ? set : listEventNames_();
}
function getCurrentPR_(swimmer, eventName) {
  if (!swimmer || !eventName) return null;
  const res = mustSheet('Results');
  const last = res.getLastRow(); if (last < 2) return null;
  const vals = res.getRange(2,1,last-1,10).getValues(); // meet,event,swimmer,seed,final,place,notes,date,isPR,curPR
  let best = null;
  for (const r of vals) {
    if (String(r[1]||'')===eventName && String(r[2]||'')===swimmer && r[4] !== '' && r[4] != null) {
      const t = Number(r[4]); if (best==null || t<best) best = t;
    }
  }
  return best; // serial or null
}
function addResultRow_(payload) {
  const res = mustSheet('Results');
  const meet = String(payload.meet||'').trim();
  const eventName = String(payload.event||'').trim();
  const swimmer = String(payload.swimmer||'').trim();
  if (!meet || !eventName || !swimmer) throw new Error('Meet, Event, and Swimmer are required.');
  const seedSerial = parseTimeSerial_(payload.seed||'');
  const finalSerial = parseTimeSerial_(payload.final||'');
  if (finalSerial == null) throw new Error('Final time is required (mm:ss.xx or ss.xx).');
  const place = payload.place || '';
  const notes = payload.notes || '';
  const date = payload.date ? new Date(payload.date) : new Date();
  res.getRange(res.getLastRow()+1,1,1,8).setValues([[meet,eventName,swimmer,seedSerial||'',finalSerial,place,notes,date]]);
  try { refreshPRs(); } catch(e) {}
  return {ok:true};
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

meetSel.addEventListener('change', ()=>{ google.script.run.withSuccessHandler(list=>{ fill(eventSel, list); prCheck(); }).listActiveEventsForMeet_(meetSel.value); });
eventSel.addEventListener('change', prCheck); swimSel.addEventListener('change', prCheck);

function prCheck(){
  const sw=swimSel.value, ev=eventSel.value; if(!sw||!ev){ hint.textContent=''; return; }
  google.script.run.withSuccessHandler(serial=>{
    if(serial==null){ hint.textContent='No PR recorded yet for this swimmer/event.'; return; }
    // Convert serial days -> mm:ss.xx
    const sec = serial*86400; const m=Math.floor(sec/60); const s=(sec%60).toFixed(2).padStart(5,'0'); 
    hint.innerHTML = 'Current PR: <b>'+m+':'+s+'</b>';
  }).getCurrentPR_(sw, ev);
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
    .addResultRow_(payload);
}
</script></body></html>`;
}

/** =========================
 * ADMIN: Add Meet / Add Event (NEW)
 * ========================= */
function openAddMeetSidebar() {
  const html = HtmlService.createHtmlOutput(addMeetSidebarHtml_()).setTitle('Add Meet');
  SpreadsheetApp.getUi().showSidebar(html);
}
function addMeet_(payload) {
  const me = mustSheet('Meets');
  const name = String(payload.name||'').trim();
  if (!name) throw new Error('Meet name is required.');
  const date = payload.date ? new Date(payload.date) : '';
  const loc = String(payload.location||'').trim();
  const course = String(payload.course||'').trim(); // SCY/LCM/SCM
  const notes = String(payload.notes||'').trim();
  // Append
  me.appendRow([name,date,loc,course,notes]);
  ensureMeetsHasJVColumn();
  ensureMeetEventsTemplate();
  setupValidations();
  return {ok:true};
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
    .withFailureHandler(err=>alert('Error: '+err.message)).addMeet_(p);
}
</script></body></html>`;
}

function openAddEventSidebar() {
  const html = HtmlService.createHtmlOutput(addEventSidebarHtml_()).setTitle('Add Event');
  SpreadsheetApp.getUi().showSidebar(html);
}
function addEvent_(payload) {
  const ev = mustSheet('Events');
  const name = String(payload.name||'').trim();
  if (!name) throw new Error('Event name is required.');
  const type = String(payload.type||'').trim() || 'Individual';
  const dist = Number(payload.distance||'');
  const stroke = String(payload.stroke||'').trim();
  const defActive = !!payload.defaultActive;
  const addJV = !!payload.addJV;
  const reseed = !!payload.reseed;

  ev.appendRow([name, type, dist||'', stroke, defActive]);
  if (addJV) ev.appendRow([`${name} (JV)`, type, dist||'', stroke, defActive]);

  if (reseed) {
    reseedMeetEntryFromEvents_();
  }
  ensureMeetEventsTemplate();
  setupValidations();

  return {ok:true};
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
  if(!p.name.trim()){alert('Event name is required.');return;}
  google.script.run.withSuccessHandler(()=>{alert('Event added ✓');google.script.host.close();})
    .withFailureHandler(err=>alert('Error: '+err.message)).addEvent_(p);
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
function getColValues(sheet, col, startRow=2) {
  const last = sheet.getLastRow();
  if (last < startRow) return [];
  return sheet.getRange(startRow, col, last-startRow+1, 1).getValues().map(r=>r[0]).filter(Boolean);
}
function _getColValues(sheet, col, startRow=2) {
  const last = sheet.getLastRow();
  if (last < startRow) return [];
  return sheet.getRange(startRow, col, last-startRow+1, 1).getValues().map(r=>r[0]).filter(Boolean);
}
function findLastDataRow(sheet, keyCol, startRow) {
  const last = sheet.getLastRow();
  if (last < startRow) return startRow-1;
  const vals = sheet.getRange(startRow, keyCol, last-startRow+1, 1).getValues().map(r=>r[0]);
  let end = startRow-1;
  for (let i=0;i<vals.length;i++) if (vals[i]) end = startRow+i;
  return end;
}
function findDuplicates(arr) {
  const seen = new Set(), dup = new Set();
  for (const x of arr) { if (seen.has(x)) dup.add(x); else seen.add(x); }
  return [...dup];
}
function toast(msg) { SpreadsheetApp.getActive().toast(msg, 'Coach Tools', 5); }
function timestamp_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HHmm'); }

/** =========================
 * IMPORT: Bulk Import (CSV Paste)
 * ========================= */
function openBulkImportSidebar() {
  const html = HtmlService.createHtmlOutput(bulkImportSidebarHtml_()).setTitle('Bulk Import');
  SpreadsheetApp.getUi().showSidebar(html);
}

// Server: do the import
function bulkImport_(payload) {
  const type = String(payload.type||'').toLowerCase(); // 'swimmers' | 'meets' | 'pr'
  const csv  = String(payload.csv||'').trim();
  const hasHeader = !!payload.hasHeader;
  const defaultDate = payload.defaultDate ? new Date(payload.defaultDate) : new Date();

  if (!csv) throw new Error('Paste CSV data first.');

  // Parse CSV (handles quotes/commas properly)
  const rows = Utilities.parseCsv(csv).filter(r => r.some(c => String(c).trim() !== ''));
  if (!rows.length) throw new Error('No rows detected.');
  const data = hasHeader ? rows.slice(1) : rows;

  if (type === 'swimmers') return importSwimmers_(data);
  if (type === 'meets')    return importMeets_(data);
  if (type === 'pr')       return importPRs_(data, defaultDate);

  throw new Error('Unknown import type: ' + type);
}

function importSwimmers_(data) {
  const sw = mustSheet('Swimmers');
  ensureSwimmersLevelColumn_(); // ensures Name, Grad Year, Gender, Level, Notes headers exist
  const out = [];
  for (const r of data) {
    const name = (r[0]||'').toString().trim();
    if (!name) continue;
    const grad = r[1] ? Number(r[1]) : '';
    const gender = (r[2]||'').toString().trim();
    const level  = (r[3]||'').toString().trim();
    const notes  = (r[4]||'').toString();
    out.push([name, grad, gender, level, notes]);
  }
  if (out.length) sw.getRange(sw.getLastRow()+1,1,out.length,5).setValues(out);
  setupValidations();
  return { inserted: out.length, kind:'swimmers' };
}

function importMeets_(data) {
  const me = mustSheet('Meets');
  const out = [];
  const jvMarks = [];
  for (const r of data) {
    const name = (r[0]||'').toString().trim();
    if (!name) continue;
    const date = r[1] ? new Date(r[1]) : '';
    const loc  = (r[2]||'').toString().trim();
    const course = (r[3]||'').toString().trim(); // SCY/LCM/SCM
    const notes  = (r[4]||'').toString().trim();
    const hasJV  = (r[5]||'').toString().trim().toLowerCase();
    out.push([name, date, loc, course, notes]);
    jvMarks.push(hasJV); // remember per-row intent
  }
  if (out.length) {
    const start = me.getLastRow()+1;
    me.getRange(start,1,out.length,5).setValues(out);
    ensureMeetsHasJVColumn();
    // set Has JV? checkboxes using text flags like 'true','yes','y','1'
    const headers = me.getRange(1,1,1,me.getLastColumn()).getValues()[0].map(h=>String(h||'').trim().toLowerCase());
    const jvCol = headers.indexOf('has jv?') + 1;
    for (let i=0;i<jvMarks.length;i++) {
      const val = jvMarks[i];
      const isTrue = ['true','yes','y','1'].includes(val);
      me.getRange(start+i, jvCol).setValue(isTrue);
    }
  }
  ensureMeetEventsTemplate(); // cross-join presets for new meets
  setupValidations();
  return { inserted: out.length, kind:'meets' };
}

function importPRs_(data, fallbackDate) {
  const res = mustSheet('Results');
  const rows = [];
  for (const r of data) {
    const swimmer = (r[0]||'').toString().trim();
    const event   = (r[1]||'').toString().trim();
    const timeStr = (r[2]||'').toString().trim();
    if (!swimmer || !event || !timeStr) continue;
    const serial = parseTimeSerial_(timeStr);
    if (serial == null) continue;
    const date = r[3] ? new Date(r[3]) : fallbackDate;
    rows.push(['PR Baseline', event, swimmer, '', serial, '', 'Imported baseline', date]);
  }
  if (rows.length) res.getRange(res.getLastRow()+1,1,rows.length,8).setValues(rows);
  refreshPRs();
  return { inserted: rows.length, kind:'pr' };
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
  }).bulkImport_(payload);
}
</script>
</body></html>`;
  return tmpl;
}

function safeCreateFilter_(sheet, range) {
  try {
    const f = sheet.getFilter && sheet.getFilter();
    if (f) f.remove();                 // remove existing basic filter
  } catch (e) {
    // ignore; some sheets won't support getFilter in older contexts
  }
  range.createFilter();                // then create a fresh one
}
