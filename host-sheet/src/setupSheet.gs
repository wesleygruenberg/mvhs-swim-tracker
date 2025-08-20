// ===== Host wrapper (bound to the Sheet) =====

// Build the Coach Tools menu on every open
function onOpen() {
  // First, setup the full CoachTools menu from the library
  CoachToolsCore.setupCoachToolsMenu();

  // Then add our attendance feature to it
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Attendance')
    .addItem('📋 Open Attendance Tracker (Sidebar)', 'openAttendanceSidebar')
    .addItem('🔗 Show Web App Link', 'showAttendanceWebLink')
    .addToUi();
}

// 1:1 wrappers to the library (keep these small and boring)
function refreshAll() {
  CoachToolsCore.refreshAll();
}
function aboutCoachTools() {
  CoachToolsCore.aboutCoachTools();
}

function openBulkImportSidebar() {
  const html = CoachToolsCore.buildBulkImportSidebar();
  SpreadsheetApp.getUi().showSidebar(html);
}
function openAddResultSidebar() {
  const html = CoachToolsCore.buildAddResultSidebar();
  SpreadsheetApp.getUi().showSidebar(html);
}
function openAddSwimmerSidebar() {
  CoachToolsCore.openAddSwimmerSidebar();
}
function openAddMeetSidebar() {
  const html = CoachToolsCore.buildAddMeetSidebar();
  SpreadsheetApp.getUi().showSidebar(html);
}
function openAddEventSidebar() {
  const html = CoachToolsCore.buildAddEventSidebar();
  SpreadsheetApp.getUi().showSidebar(html);
}
function openAttendanceSidebar() {
  const t = HtmlService.createTemplateFromFile('AttendanceUI');
  t.defaultDate = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
  t.isWebApp = false;
  SpreadsheetApp.getUi().showSidebar(
    t.evaluate().setTitle('Attendance').setWidth(360)
  );
}

function createAttendanceSummary() {
  CoachToolsCore.createAttendanceSummary();
}

function createTestAttendanceData() {
  CoachToolsCore.createTestAttendanceData();
}

// NEW: Web app entrypoint for phones
function doGet(e) {
  const t = HtmlService.createTemplateFromFile('AttendanceUI');
  t.defaultDate = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
  t.isWebApp = true;
  return t
    .evaluate()
    .setTitle('Attendance')
    .addMetaTag(
      'viewport',
      'width=device-width, initial-scale=1, maximum-scale=1'
    );
}

// Helper: show the current deployment URL in a dialog
function showAttendanceWebLink() {
  const url = getWebAppUrl_();
  SpreadsheetApp.getUi().alert(
    url
      ? 'Attendance Web App URL:\n\n' +
          url +
          '\n\nOpen on your phone and "Add to Home Screen."'
      : 'No web app deployment found.\nUse: Deploy → New deployment → Web app.'
  );
}

// Try to read the web app URL from the last deployment (best-effort)
function getWebAppUrl_() {
  try {
    // Apps Script doesn't provide a direct getter; hardcode after deploy if needed.
    // Option A (manual): paste your deployed URL below and return it.
    // return 'https://script.google.com/macros/s/.../exec';

    // Option B (Property): store once after deploy
    return (
      PropertiesService.getScriptProperties().getProperty(
        'ATTENDANCE_WEB_APP_URL'
      ) || ''
    );
  } catch (e) {
    return '';
  }
}

function refreshPRs() {
  CoachToolsCore.refreshPRs();
}
function checkLineup() {
  CoachToolsCore.checkLineup();
}
function buildCoachPacket() {
  CoachToolsCore.buildCoachPacket();
}
function createSnapshot() {
  CoachToolsCore.createSnapshot();
}
function applyMeetPresets() {
  CoachToolsCore.applyMeetPresets();
}
function ensureMeetPresetsTemplate() {
  CoachToolsCore.ensureMeetEventsTemplate();
}
function setupValidations() {
  CoachToolsCore.setupValidations();
}
function enableJVSupport() {
  CoachToolsCore.enableJVSupport();
}
function adminClearSampleData() {
  CoachToolsCore.adminClearSampleData();
}
function generateSampleTeam50() {
  CoachToolsCore.generateSampleTeam50();
}
function cloneMakeCleanCopy() {
  CoachToolsCore.cloneMakeCleanCopy();
}
function cloneNewSeasonCarryForward() {
  CoachToolsCore.cloneNewSeasonCarryForward();
}
function cloneCleanBaseline() {
  CoachToolsCore.cloneCleanBaseline();
}
function ensureSettingsSheet() {
  CoachToolsCore.ensureSettingsSheet();
}
function applyLimitsFromSettings() {
  CoachToolsCore.applyLimitsFromSettings();
}
function ensureMeetsHasJVColumn() {
  CoachToolsCore.ensureMeetsHasJVColumn();
}
function generateRosterRankingsFromCSV() {
  CoachToolsCore.generateRosterRankingsFromCSV();
}
function createRawTryoutResultsSheet() {
  CoachToolsCore.createRawTryoutResultsSheet();
}
function generateTryoutRankingsFromSheet() {
  CoachToolsCore.generateTryoutRankingsFromSheet();
}
function generateVarsityJVSquads() {
  CoachToolsCore.generateVarsityJVSquads();
}
function processCompleteTryouts() {
  CoachToolsCore.processCompleteTryouts();
}
function generateRosterAnnouncement() {
  CoachToolsCore.generateRosterAnnouncement();
}
function applySwimmersColorCoding() {
  CoachToolsCore.applySwimmersColorCoding();
}
function createPRsFromTryouts() {
  CoachToolsCore.createPRsFromTryouts();
}
function setupRelayEvents() {
  CoachToolsCore.setupRelayEvents();
}
function generateRelayAssignments() {
  CoachToolsCore.generateRelayAssignments();
}
function refreshSwimmerAssignmentSummary() {
  CoachToolsCore.refreshSwimmerAssignmentSummary();
}

function buildAttendanceSidebar() {
  const t = HtmlService.createTemplateFromFile('AttendanceUI');
  // Default date = today in local time (yyyy-mm-dd)
  t.defaultDate = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
  return t
    .evaluate()
    .setTitle('Attendance')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(360);
}

// Helper for HTML template includes
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
