// ===== Host wrapper (bound to the Sheet) =====

// Build the Coach Tools menu on every open
function onOpen() {
  CoachToolsCore.setupCoachToolsMenu();
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
