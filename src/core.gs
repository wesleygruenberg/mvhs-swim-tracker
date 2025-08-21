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
  ATTENDANCE_SUMMARY: 'Attendance Summary',
  MASTER_ATTENDANCE: 'Master Attendance',
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

/**
 * About dialog for Coach Tools
 */
function aboutCoachTools() {
  const ui = SpreadsheetApp.getUi();
  const id = ScriptApp.getScriptId();
  ui.alert(
    'About Coach Tools',
    `MVHS Swim Coach Tools ${LIB_VER}

Features:
• Meet management and lineup tracking
• Swimmer and event management 
• Results tracking with PR analysis
• Roster ranking analysis from CSV data
• Raw tryout data import from CSV
• JV/Varsity support
• Bulk import capabilities
• Snapshot and reporting tools

New in this version:
• CSV Roster Rankings: Generate male/female team rankings with individual event ranks, best ranks, and average rankings
• Raw Tryout Results: Import CSV data directly into a formatted Google Sheet for easy viewing and analysis
• Tryout Rankings: Generate rankings from existing Tryouts sheet using 8 core events (50/100/200/500 Free, 100 Breast/Fly/Back, 200 IM). Missing times are ranked as worst in that event.

Use "Coach Tools > Roster > Generate Roster Rankings from CSV" to analyze your team's performance data.
Use "Coach Tools > Roster > Create Raw Tryout Results Sheet" to import raw CSV data.
Use "Coach Tools > Roster > Generate Tryout Rankings from Sheet" to rank swimmers from existing Tryouts data.

Script ID: ${id}`,
    ui.ButtonSet.OK
  );
  return { version: LIB_VER, id };
}

/**
 * Creates or updates the "Raw Tryout Results" sheet with CSV data
 */
function createRawTryoutResultsSheet() {
  const csvFiles = DriveApp.getFilesByName('MVHS_Times_2025.csv');

  if (!csvFiles.hasNext()) {
    toast(
      'CSV file "MVHS_Times_2025.csv" not found in your Google Drive. Please upload it first.'
    );
    return;
  }

  const csvFile = csvFiles.next();
  const csvContent = csvFile.getBlob().getDataAsString();

  try {
    importRawTryoutData_(csvContent);
    toast('Raw Tryout Results sheet has been created/updated successfully!');
  } catch (e) {
    toast('Error creating Raw Tryout Results sheet: ' + e.message);
    console.error('Import error:', e);
  }
}

/**
 * Generates tryout rankings from the existing Tryouts sheet
 */
function generateTryoutRankingsFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const tryoutsSheet = ss.getSheetByName('Tryouts');
    if (!tryoutsSheet) {
      toast(
        'Error: "Tryouts" sheet not found. Please make sure you have a sheet named "Tryouts" with your tryout data.'
      );
      return;
    }

    const data = tryoutsSheet.getDataRange().getValues();
    if (data.length < 2) {
      toast('Error: No data found in Tryouts sheet.');
      return;
    }

    processTryoutRankings_(data);
    toast(
      'Tryout rankings generated successfully! Check the "Tryout Rankings" sheet.'
    );
  } catch (e) {
    toast('Error generating tryout rankings: ' + e.message);
    console.error('Tryout rankings error:', e);
  }
}

/**
 * Comprehensive function to process complete tryout workflow:
 * 1. Generate tryout rankings from Tryouts sheet
 * 2. Create Varsity/JV squad proposals
 * 3. Add new swimmers to Swimmers tab with PR baselines
 */
function processCompleteTryouts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    console.log('Starting complete tryout processing...');

    // Check for required sheets
    const tryoutsSheet = ss.getSheetByName('Tryouts');
    if (!tryoutsSheet) {
      toast(
        'Error: "Tryouts" sheet not found. Please make sure you have a sheet named "Tryouts" with your tryout data.'
      );
      return;
    }

    const data = tryoutsSheet.getDataRange().getValues();
    if (data.length < 2) {
      toast('Error: No data found in Tryouts sheet.');
      return;
    }

    // Get settings for Varsity/JV generation
    const settingsResult = getVarsitySettings_();
    if (!settingsResult.success) {
      toast(
        `Warning: Could not read Settings sheet (${settingsResult.error}). Using defaults.`
      );
    }
    const settings = settingsResult.success
      ? settingsResult.data
      : {
          varsitySpotsF: 15,
          varsitySpotsM: 15,
          bubbleSize: 3,
        };

    // Step 1: Generate tryout rankings
    console.log('Step 1: Generating tryout rankings...');
    processTryoutRankings_(data);

    // Step 2: Generate Varsity/JV squads
    console.log('Step 2: Generating Varsity/JV squads...');
    generateVarsityJVProposal_(data, settings);

    // Step 3: Add swimmers to Swimmers tab
    console.log('Step 3: Adding swimmers to Swimmers tab...');
    const addedSwimmers = addSwimmersFromTryouts_(data);

    // Step 4: Generate roster announcement
    console.log('Step 4: Generating roster announcement...');
    generateRosterAnnouncement();

    // Show comprehensive success message
    const ui = SpreadsheetApp.getUi();
    const swimmerMessage =
      addedSwimmers.newCount > 0
        ? `${addedSwimmers.newCount} new swimmers added to roster`
        : '';
    const updateMessage =
      addedSwimmers.updatedCount > 0
        ? `${addedSwimmers.updatedCount} existing swimmers updated with levels/notes`
        : '';

    let swimmerSummary = [swimmerMessage, updateMessage]
      .filter(msg => msg)
      .join(', ');
    if (!swimmerSummary) {
      swimmerSummary = 'Swimmer data processed';
    }

    ui.alert(
      'Complete Tryout Processing Successful!',
      `✅ Tryout Rankings: Generated successfully\n` +
        `✅ Varsity/JV Squads: Created with ${settings.varsitySpotsF}F/${settings.varsitySpotsM}M varsity spots\n` +
        `✅ Swimmers Updated: ${swimmerSummary}\n` +
        `✅ PR Baselines: ${addedSwimmers.prCount} tryout times added as PR baselines\n` +
        `✅ Roster Announcement: Clean 4-column format ready for email\n\n` +
        `Check these sheets:\n` +
        `• "Tryout Rankings" - Individual event and overall rankings\n` +
        `• "Varsity/JV - Autogenerated" - Squad proposals with bubble analysis\n` +
        `• "Swimmers" - Updated roster with proper Varsity/JV levels and best events\n` +
        `• "Roster Announcement" - Email-ready 4-column roster format`,
      ui.ButtonSet.OK
    );

    console.log('Complete tryout processing finished successfully');
  } catch (e) {
    toast('Error processing complete tryouts: ' + e.message);
    console.error('Complete tryout processing error:', e);
  }
}

/**
 * Generates proposed Varsity/JV squads based on tryout rankings and settings
 */
function generateVarsityJVSquads() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // First, make sure we have tryout rankings
    let rankingsSheet;
    try {
      rankingsSheet = ss.getSheetByName('Tryout Rankings');
    } catch (e) {
      toast(
        'Error: "Tryout Rankings" sheet not found. Please generate tryout rankings first.'
      );
      return;
    }

    // Get settings
    const settings = getVarsitySettings_();
    if (!settings.success) {
      toast('Error reading settings: ' + settings.error);
      return;
    }

    // Get tryout data to process rankings
    const tryoutsSheet = ss.getSheetByName('Tryouts');
    if (!tryoutsSheet) {
      toast('Error: "Tryouts" sheet not found.');
      return;
    }

    const data = tryoutsSheet.getDataRange().getValues();
    if (data.length < 2) {
      toast('Error: No data found in Tryouts sheet.');
      return;
    }

    generateVarsityJVProposal_(data, settings.data);
    toast(
      'Varsity/JV squad proposals generated successfully! Check the "Varsity/JV - Autogenerated" sheet.'
    );
  } catch (e) {
    toast('Error generating Varsity/JV squads: ' + e.message);
    console.error('Varsity/JV generation error:', e);
  }
}

/**
 * Helper function to import CSV data into the Raw Tryout Results sheet
 * @param {string} csvContent - The CSV content to import
 */
function importRawTryoutData_(csvContent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Parse CSV data
  const csvData = Utilities.parseCsv(csvContent);

  if (csvData.length === 0) {
    throw new Error('CSV file appears to be empty');
  }

  // Check if sheet exists, create if not
  let sheet;
  try {
    sheet = ss.getSheetByName('Raw Tryout Results');
  } catch (e) {
    sheet = ss.insertSheet('Raw Tryout Results');
  }

  // Clear existing content
  sheet.clear();

  // Define the required events for tracking missing times
  const requiredEvents = [
    '50 Free',
    '100 Free',
    '200 Free',
    '500 Free',
    '100 Breast',
    '100 Fly',
    '100 Back',
    '200 IM',
  ];

  // Find event columns in the CSV data
  const headers = csvData[0];
  const eventColumns = [];
  requiredEvents.forEach(eventName => {
    for (let i = 0; i < headers.length; i++) {
      const header = headers[i].toString().trim();
      if (header === eventName) {
        eventColumns.push({ index: i, name: eventName });
        break;
      }
    }
  });

  // Add "Missing Events" column to headers
  const enhancedHeaders = [...headers, 'Missing Events'];
  const numCols = enhancedHeaders.length;

  // Prepare enhanced data with missing events count
  const enhancedData = [enhancedHeaders];

  for (let i = 1; i < csvData.length; i++) {
    const row = csvData[i];
    let missingCount = 0;

    // Count missing events for this swimmer
    eventColumns.forEach(event => {
      const timeValue = row[event.index];
      const hasValidTime =
        timeValue &&
        timeValue.toString().trim() &&
        parseTimeToSeconds_(timeValue.toString().trim()) > 0;
      if (!hasValidTime) {
        missingCount++;
      }
    });

    // Add the missing count to the row
    enhancedData.push([...row, missingCount]);
  }

  const numRows = enhancedData.length;

  // Add all data to sheet
  sheet.getRange(1, 1, numRows, numCols).setValues(enhancedData);

  // Format header row
  if (numRows > 0) {
    sheet
      .getRange(1, 1, 1, numCols)
      .setFontWeight('bold')
      .setBackground('#e6f3ff');
  }

  // Apply color coding to time columns and missing events column
  if (numRows > 1) {
    // Color code event time columns
    eventColumns.forEach(event => {
      const colIndex = event.index + 1; // Convert to 1-based indexing
      const range = sheet.getRange(2, colIndex, numRows - 1, 1);
      const values = range.getValues();

      // Apply conditional formatting based on time presence
      for (let i = 0; i < values.length; i++) {
        const cellRange = sheet.getRange(i + 2, colIndex);
        const timeValue = values[i][0];

        const hasValidTime =
          timeValue &&
          timeValue.toString().trim() &&
          parseTimeToSeconds_(timeValue.toString().trim()) > 0;
        if (!hasValidTime) {
          // Missing time - light red background
          cellRange.setBackground('#ffcccc');
        } else {
          // Has time - light green background
          cellRange.setBackground('#ccffcc');
        }
      }
    });

    // Color code the "Missing Events" column
    const missingEventsCol = numCols;
    const missingRange = sheet.getRange(2, missingEventsCol, numRows - 1, 1);
    const missingValues = missingRange.getValues();

    for (let i = 0; i < missingValues.length; i++) {
      const cellRange = sheet.getRange(i + 2, missingEventsCol);
      const missingCount = missingValues[i][0];

      if (missingCount >= 4) {
        // 4 or more missing - red background with white text
        cellRange
          .setBackground('#ff4444')
          .setFontColor('white')
          .setFontWeight('bold');
      } else if (missingCount >= 2) {
        // 2-3 missing - yellow background
        cellRange.setBackground('#ffff99');
      } else if (missingCount === 1) {
        // 1 missing - light yellow background
        cellRange.setBackground('#ffffcc');
      } else {
        // All events present - light green background
        cellRange.setBackground('#ccffcc');
      }
    }
  }

  // Auto-resize columns
  for (let i = 1; i <= numCols; i++) {
    sheet.autoResizeColumn(i);
  }

  // Freeze header row
  if (numRows > 1) {
    sheet.setFrozenRows(1);
  }

  console.log(
    `Imported ${numRows} rows and ${numCols} columns to Raw Tryout Results sheet with color coding and missing events tracking`
  );
}

/**
 * Helper function to read varsity settings from the Settings sheet
 * @returns {Object} Object with success flag and data/error
 */
function getVarsitySettings_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName('Settings');

    if (!settingsSheet) {
      return { success: false, error: 'Settings sheet not found' };
    }

    const data = settingsSheet.getDataRange().getValues();
    const settings = {
      varsitySpotsF: 15, // default
      varsitySpotsM: 15, // default
      bubbleSize: 3, // default
    };

    // Look for the settings values
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row[0] && row[1]) {
        const key = row[0].toString().trim();
        const value = row[1];

        if (key === 'Varsity Spots - F') {
          settings.varsitySpotsF = parseInt(value) || 15;
        } else if (key === 'Varsity Spots - M') {
          settings.varsitySpotsM = parseInt(value) || 15;
        } else if (key === 'Bubble size') {
          settings.bubbleSize = parseInt(value) || 3;
        }
      }
    }

    console.log('Settings loaded:', settings);
    return { success: true, data: settings };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Helper function to generate the Varsity/JV proposal
 * @param {Array} tryoutData - Raw tryout data from sheet
 * @param {Object} settings - Settings object with varsity spots and bubble size
 */
function generateVarsityJVProposal_(tryoutData, settings) {
  // First, get the processed rankings
  const headers = tryoutData[0];
  const swimmers = tryoutData.slice(1);

  // Define the required events
  const requiredEvents = [
    '50 Free',
    '100 Free',
    '200 Free',
    '500 Free',
    '100 Breast',
    '100 Fly',
    '100 Back',
    '200 IM',
  ];

  // Find event columns
  const eventColumns = [];
  requiredEvents.forEach(eventName => {
    for (let i = 0; i < headers.length; i++) {
      const header = headers[i].toString().trim();
      if (header === eventName) {
        eventColumns.push({ index: i, name: eventName });
        break;
      }
    }
  });

  // Separate swimmers by gender and calculate rankings
  const maleSwimmers = swimmers.filter(
    row => row[1] && row[1].toString().toUpperCase() === 'M'
  );
  const femaleSwimmers = swimmers.filter(
    row => row[1] && row[1].toString().toUpperCase() === 'F'
  );

  // Calculate rankings for each gender
  const maleRankings = calculateTryoutRankings_(
    maleSwimmers,
    eventColumns,
    'Male'
  );
  const femaleRankings = calculateTryoutRankings_(
    femaleSwimmers,
    eventColumns,
    'Female'
  );

  // Create the proposal sheet
  createVarsityJVSheet_(maleRankings, femaleRankings, settings);
}

/**
 * Create the Varsity/JV proposal sheet
 * @param {Array} maleRankings - Male swimmer rankings
 * @param {Array} femaleRankings - Female swimmer rankings
 * @param {Object} settings - Settings with varsity spots and bubble size
 */
function createVarsityJVSheet_(maleRankings, femaleRankings, settings) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create or get the proposal sheet
  let sheet;
  try {
    sheet = ss.getSheetByName('Varsity/JV - Autogenerated');
    sheet.clear();
  } catch (e) {
    sheet = ss.insertSheet('Varsity/JV - Autogenerated');
  }

  let currentRow = 1;

  // Add header with timestamp
  sheet
    .getRange(currentRow, 1)
    .setValue(
      `Varsity/JV Squad Proposal - Generated ${new Date().toLocaleString()}`
    );
  sheet
    .getRange(currentRow, 1, 1, 7)
    .setBackground('#2c5aa0')
    .setFontColor('white')
    .setFontWeight('bold');
  currentRow += 2;

  // Helper function to write a gender's squad breakdown
  function writeGenderSquads(rankings, genderLabel, varsitySpots) {
    if (rankings.length === 0) return;

    // Gender header
    sheet
      .getRange(currentRow, 1)
      .setValue(`${genderLabel} Squad (${rankings.length} swimmers)`);
    sheet
      .getRange(currentRow, 1, 1, 7)
      .setBackground('#4a90e2')
      .setFontColor('white')
      .setFontWeight('bold');
    currentRow++;

    // Determine squad assignments
    const varsityCount = Math.min(varsitySpots, rankings.length);
    const bubbleStart = Math.max(0, varsitySpots - settings.bubbleSize);
    const bubbleEnd = Math.min(
      rankings.length,
      varsitySpots + settings.bubbleSize
    );

    // Write Varsity section
    if (varsityCount > 0) {
      sheet.getRange(currentRow, 1).setValue('VARSITY');
      sheet
        .getRange(currentRow, 1, 1, 7)
        .setBackground('#90EE90')
        .setFontWeight('bold');
      currentRow++;

      // Varsity headers
      const headers = [
        'Rank',
        'Name',
        'Avg Rank',
        'Best Rank',
        'Best Event',
        'Missing Events',
        'Status',
      ];
      sheet.getRange(currentRow, 1, 1, headers.length).setValues([headers]);
      sheet
        .getRange(currentRow, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#e6f3ff');
      currentRow++;

      for (let i = 0; i < varsityCount; i++) {
        const swimmer = rankings[i];
        let status = 'Varsity';
        if (i >= bubbleStart) {
          status = 'Varsity (Bubble - Last In)';
        }

        const row = [
          i + 1,
          swimmer.name,
          swimmer.avgRank || 'N/A',
          swimmer.bestRank || 'N/A',
          swimmer.bestEvent || 'N/A',
          swimmer.missingEvents || 0,
          status,
        ];

        sheet.getRange(currentRow, 1, 1, row.length).setValues([row]);

        // Apply color coding based on missing events
        const rowRange = sheet.getRange(currentRow, 1, 1, row.length);
        if (swimmer.missingEvents >= 4) {
          // 4+ missing events - red background with white text
          rowRange.setBackground('#ff4444').setFontColor('white');
        } else if (i >= bubbleStart) {
          rowRange.setBackground('#FFE4B5'); // Light orange for bubble
        }
        currentRow++;
      }
    }

    currentRow++; // Space

    // Write JV/Bubble section
    if (rankings.length > varsitySpots) {
      sheet.getRange(currentRow, 1).setValue('JV / BUBBLE ANALYSIS');
      sheet
        .getRange(currentRow, 1, 1, 7)
        .setBackground('#FFA500')
        .setFontColor('white')
        .setFontWeight('bold');
      currentRow++;

      // JV headers
      const headers = [
        'Rank',
        'Name',
        'Avg Rank',
        'Best Rank',
        'Best Event',
        'Missing Events',
        'Status',
      ];
      sheet.getRange(currentRow, 1, 1, headers.length).setValues([headers]);
      sheet
        .getRange(currentRow, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#e6f3ff');
      currentRow++;

      for (let i = varsitySpots; i < rankings.length; i++) {
        const swimmer = rankings[i];
        let status = 'JV';
        if (i < bubbleEnd) {
          status = 'JV (Bubble - First Out)';
        }

        const row = [
          i + 1,
          swimmer.name,
          swimmer.avgRank || 'N/A',
          swimmer.bestRank || 'N/A',
          swimmer.bestEvent || 'N/A',
          swimmer.missingEvents || 0,
          status,
        ];

        sheet.getRange(currentRow, 1, 1, row.length).setValues([row]);

        // Apply color coding based on missing events
        const rowRange = sheet.getRange(currentRow, 1, 1, row.length);
        if (swimmer.missingEvents >= 4) {
          // 4+ missing events - red background with white text
          rowRange.setBackground('#ff4444').setFontColor('white');
        } else if (i < bubbleEnd) {
          rowRange.setBackground('#FFE4B5'); // Light orange for bubble
        }
        currentRow++;
      }
    }

    currentRow += 2; // Space between genders
  }

  // Write both gender sections
  writeGenderSquads(femaleRankings, 'Female', settings.varsitySpotsF);
  writeGenderSquads(maleRankings, 'Male', settings.varsitySpotsM);

  // Add summary section
  sheet.getRange(currentRow, 1).setValue('SUMMARY');
  sheet
    .getRange(currentRow, 1, 1, 6)
    .setBackground('#2c5aa0')
    .setFontColor('white')
    .setFontWeight('bold');
  currentRow++;

  const summaryData = [
    ['', 'Female', 'Male'],
    ['Total Swimmers', femaleRankings.length, maleRankings.length],
    ['Varsity Spots', settings.varsitySpotsF, settings.varsitySpotsM],
    ['Bubble Size', settings.bubbleSize, settings.bubbleSize],
    [
      'JV Swimmers',
      Math.max(0, femaleRankings.length - settings.varsitySpotsF),
      Math.max(0, maleRankings.length - settings.varsitySpotsM),
    ],
  ];

  sheet.getRange(currentRow, 1, summaryData.length, 3).setValues(summaryData);
  sheet
    .getRange(currentRow, 1, 1, 3)
    .setFontWeight('bold')
    .setBackground('#e6f3ff');

  // Auto-resize columns
  for (let i = 1; i <= 7; i++) {
    sheet.autoResizeColumn(i);
  }

  // Freeze header row
  sheet.setFrozenRows(1);

  console.log(
    `Created Varsity/JV proposal with ${femaleRankings.length} female and ${maleRankings.length} male swimmers`
  );
}

/**
 * Helper function to add swimmers from tryout data to the Swimmers tab
 * @param {Array} tryoutData - The raw tryout data
 * @returns {Object} Summary of swimmers added and PRs created
 */
function addSwimmersFromTryouts_(tryoutData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const swimmersSheet = mustSheet('Swimmers');
  const resultsSheet = mustSheet('Results');

  ensureSwimmersLevelColumn_();

  const headers = tryoutData[0];
  const swimmers = tryoutData.slice(1);

  // Define the required events for PR baseline
  const requiredEvents = [
    '50 Free',
    '100 Free',
    '200 Free',
    '500 Free',
    '100 Breast',
    '100 Fly',
    '100 Back',
    '200 IM',
  ];

  // Find event columns
  const eventColumns = [];
  requiredEvents.forEach(eventName => {
    for (let i = 0; i < headers.length; i++) {
      const header = headers[i].toString().trim();
      if (header === eventName) {
        eventColumns.push({ index: i, name: eventName });
        break;
      }
    }
  });

  // Get Varsity settings to determine squad assignments
  const settingsResult = getVarsitySettings_();
  const settings = settingsResult.success
    ? settingsResult.data
    : {
        varsitySpotsF: 15,
        varsitySpotsM: 15,
        bubbleSize: 3,
      };

  // Calculate rankings to determine Varsity/JV assignments and best events
  const maleSwimmers = swimmers.filter(
    row => row[1] && row[1].toString().toUpperCase() === 'M'
  );
  const femaleSwimmers = swimmers.filter(
    row => row[1] && row[1].toString().toUpperCase() === 'F'
  );

  const maleRankings = calculateTryoutRankings_(
    maleSwimmers,
    eventColumns,
    'Male'
  );
  const femaleRankings = calculateTryoutRankings_(
    femaleSwimmers,
    eventColumns,
    'Female'
  );

  // Create squad assignment maps
  const squadAssignments = new Map();

  // Assign male squad levels
  maleRankings.forEach((swimmer, index) => {
    const level = index < settings.varsitySpotsM ? 'V' : 'JV';
    squadAssignments.set(swimmer.name, {
      level: level,
      ranking: swimmer,
      overallRank: index + 1,
    });
  });

  // Assign female squad levels
  femaleRankings.forEach((swimmer, index) => {
    const level = index < settings.varsitySpotsF ? 'V' : 'JV';
    squadAssignments.set(swimmer.name, {
      level: level,
      ranking: swimmer,
      overallRank: index + 1,
    });
  });

  // Function to calculate best events for notes
  function calculateBestEvents(ranking, allRankings) {
    const bestEvents = [];

    // Find events where swimmer is top 5 on their gender team
    eventColumns.forEach(event => {
      const swimmerRank = ranking.eventRanks[event.name];
      if (swimmerRank && swimmerRank <= 5) {
        bestEvents.push(`${event.name} (#${swimmerRank})`);
      }
    });

    // If less than 3 best events, add their next best events
    if (bestEvents.length < 3) {
      const eventRanks = [];
      eventColumns.forEach(event => {
        const rank = ranking.eventRanks[event.name];
        if (rank && rank > 5) {
          // Only consider events not already in bestEvents
          eventRanks.push({ event: event.name, rank: rank });
        }
      });

      // Sort by rank and add up to 3 total
      eventRanks.sort((a, b) => a.rank - b.rank);
      const needed = 3 - bestEvents.length;
      for (let i = 0; i < Math.min(needed, eventRanks.length); i++) {
        bestEvents.push(`${eventRanks[i].event} (#${eventRanks[i].rank})`);
      }
    }

    return bestEvents.join(', ');
  }

  // Get existing swimmers and their row positions
  const existingSwimmers = new Map(); // Map swimmer name to row number
  const lastRow = swimmersSheet.getLastRow();
  if (lastRow >= 2) {
    const existingData = swimmersSheet
      .getRange(2, 1, lastRow - 1, 5) // Get all 5 columns including Notes
      .getValues();
    existingData.forEach((row, index) => {
      if (row[0]) {
        const swimmerName = row[0].toString().trim();
        existingSwimmers.set(swimmerName, index + 2); // +2 because index is 0-based and data starts at row 2
      }
    });
  }

  let newSwimmersCount = 0;
  let updatedSwimmersCount = 0;
  let totalPRsAdded = 0;
  const meetLabel = 'Tryout Baseline';
  const tryoutDate = new Date();

  // Process each swimmer
  swimmers.forEach(row => {
    const name = row[0] ? row[0].toString().trim() : '';
    const gender = row[1] ? row[1].toString().toUpperCase().trim() : '';

    if (!name || !gender) return; // Skip invalid rows

    // Get squad assignment and best events
    const assignment = squadAssignments.get(name);
    if (!assignment) {
      console.log(`No squad assignment found for ${name}, skipping...`);
      return;
    }

    const level = assignment.level;
    const bestEventsNotes = calculateBestEvents(
      assignment.ranking,
      gender === 'M' ? maleRankings : femaleRankings
    );

    // Check if swimmer already exists
    if (existingSwimmers.has(name)) {
      // Update existing swimmer's Level and Notes
      const rowNumber = existingSwimmers.get(name);
      swimmersSheet.getRange(rowNumber, 4).setValue(level); // Column 4 = Level
      swimmersSheet.getRange(rowNumber, 5).setValue(bestEventsNotes); // Column 5 = Notes
      updatedSwimmersCount++;
      console.log(
        `Updated swimmer ${name} with level ${level} and notes: ${bestEventsNotes}`
      );
    } else {
      // Add new swimmer to Swimmers sheet
      // Assume graduation year 2028 for tryout swimmers (can be adjusted later)
      const gradYear = 2028;

      swimmersSheet
        .getRange(swimmersSheet.getLastRow() + 1, 1, 1, 5)
        .setValues([[name, gradYear, gender, level, bestEventsNotes]]);

      newSwimmersCount++;
      existingSwimmers.set(name, swimmersSheet.getLastRow()); // Track new swimmer to prevent duplicates
      console.log(
        `Added new swimmer ${name} with level ${level} and notes: ${bestEventsNotes}`
      );
    }

    // Add PR baselines for each event they have a time for
    const prRows = [];
    eventColumns.forEach(event => {
      const timeValue = row[event.index];
      if (timeValue && timeValue !== '') {
        const timeStr = timeValue.toString().trim();
        const serial = parseTimeSerial_(timeStr);

        if (serial !== null) {
          prRows.push([
            meetLabel,
            event.name,
            name,
            '',
            serial,
            '',
            'Added via tryout processing',
            tryoutDate,
          ]);
          totalPRsAdded++;
        }
      }
    });

    // Add PR rows to Results sheet
    if (prRows.length > 0) {
      const startRow = resultsSheet.getLastRow() + 1;
      resultsSheet.getRange(startRow, 1, prRows.length, 8).setValues(prRows);
    }
  });

  // Refresh validations and PRs
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

  console.log(
    `Added ${newSwimmersCount} new swimmers and updated ${updatedSwimmersCount} existing swimmers with ${totalPRsAdded} PR baselines`
  );

  // Apply color coding to the swimmers sheet
  applySwimmersColorCoding_();

  return {
    newCount: newSwimmersCount,
    updatedCount: updatedSwimmersCount,
    prCount: totalPRsAdded,
  };
}

/**
 * Apply color coding to the Swimmers sheet based on gender and level
 */
function applySwimmersColorCoding_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const swimmersSheet = ss.getSheetByName('Swimmers');

  if (!swimmersSheet) {
    console.log('Swimmers sheet not found, skipping color coding');
    return;
  }

  const lastRow = swimmersSheet.getLastRow();
  if (lastRow < 2) {
    console.log('No swimmer data found, skipping color coding');
    return;
  }

  // Get all swimmer data (Name, Grad Year, Gender, Level, Notes)
  const data = swimmersSheet.getRange(2, 1, lastRow - 1, 5).getValues();

  // Define colors and text colors
  const colorScheme = {
    M_V: { background: '#1c4587', text: '#ffffff' }, // Boys Varsity - Dark Blue with White text
    M_JV: { background: '#6fa8dc', text: '#000000' }, // Boys JV - Light Blue with Black text
    F_V: { background: '#274e13', text: '#ffffff' }, // Girls Varsity - Dark Green with White text
    F_JV: { background: '#93c47d', text: '#000000' }, // Girls JV - Light Green with Black text
    default: { background: '#ffffff', text: '#000000' }, // White with Black text for unknown/other
  };

  // Apply color coding row by row
  for (let i = 0; i < data.length; i++) {
    const rowNum = i + 2; // Data starts at row 2
    const [name, gradYear, gender, level, notes] = data[i];

    if (!gender || !level) continue; // Skip rows with missing data

    const genderStr = gender.toString().toUpperCase().trim();
    const levelStr = level.toString().toUpperCase().trim();

    // Determine color scheme
    let scheme = colorScheme['default'];
    if ((genderStr === 'M' || genderStr === 'MALE') && levelStr === 'V') {
      scheme = colorScheme['M_V'];
    } else if (
      (genderStr === 'F' || genderStr === 'FEMALE') &&
      levelStr === 'V'
    ) {
      scheme = colorScheme['F_V'];
    } else if (
      (genderStr === 'M' || genderStr === 'MALE') &&
      levelStr === 'JV'
    ) {
      scheme = colorScheme['M_JV'];
    } else if (
      (genderStr === 'F' || genderStr === 'FEMALE') &&
      levelStr === 'JV'
    ) {
      scheme = colorScheme['F_JV'];
    }

    // Apply background and text color to the entire row
    const range = swimmersSheet.getRange(rowNum, 1, 1, 5);
    range.setBackground(scheme.background);
    range.setFontColor(scheme.text);
  }

  console.log(`Applied color coding to ${data.length} swimmers`);
}

/**
 * Standalone function to apply color coding to Swimmers sheet
 * Can be called manually from menu
 */
function applySwimmersColorCoding() {
  try {
    applySwimmersColorCoding_();

    SpreadsheetApp.getUi().alert(
      'Color Coding Applied!',
      'Swimmers sheet has been color-coded:\n\n' +
        '🔵 Boys Varsity - Dark Blue\n' +
        '💙 Boys JV - Light Blue\n' +
        '🟢 Girls Varsity - Dark Green\n' +
        '💚 Girls JV - Light Green',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      'Error',
      `Failed to apply color coding: ${e.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    console.error('Color coding error:', e);
  }
}

/**
 * Create Personal Records (PR) sheet from tryout data
 * This establishes baseline times that can be updated throughout the season
 */
function createPRsFromTryouts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // Check for Tryouts sheet
    const tryoutsSheet = ss.getSheetByName('Tryouts');
    if (!tryoutsSheet) {
      SpreadsheetApp.getUi().alert(
        'Error',
        'Tryouts sheet not found. Please ensure you have tryout data first.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    const tryoutData = tryoutsSheet.getDataRange().getValues();
    if (tryoutData.length < 2) {
      SpreadsheetApp.getUi().alert(
        'Error',
        'No data found in Tryouts sheet.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    // Get or create Personal Records sheet
    let prSheet = ss.getSheetByName('Personal Records');
    if (!prSheet) {
      prSheet = ss.insertSheet('Personal Records');
    } else {
      prSheet.clear();
    }

    // Copy tryout data structure to PR sheet
    const headers = tryoutData[0];
    prSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Copy all swimmer data
    const swimmerData = tryoutData.slice(1);
    if (swimmerData.length > 0) {
      prSheet
        .getRange(2, 1, swimmerData.length, headers.length)
        .setValues(swimmerData);
    }

    // Format the PR sheet
    const headerRange = prSheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');

    // Apply time formatting to time columns (skip Name and Gender columns)
    for (let col = 3; col <= headers.length; col++) {
      const columnRange = prSheet.getRange(2, col, swimmerData.length, 1);
      columnRange.setNumberFormat('mm:ss.00');
    }

    // Add note about PRs
    prSheet
      .getRange(swimmerData.length + 3, 1)
      .setValue(
        'Note: These are baseline PRs from tryouts. Update times as swimmers improve throughout the season.'
      );
    prSheet.getRange(swimmerData.length + 3, 1).setFontStyle('italic');

    SpreadsheetApp.getUi().alert(
      'Personal Records Created!',
      `Personal Records sheet has been created with ${swimmerData.length} swimmers.\n\n` +
        `This establishes baseline times from tryouts that you can update throughout the season.\n\n` +
        `Next steps:\n` +
        `• Update PR times as swimmers improve\n` +
        `• Use these PRs for relay assignments and meet lineups`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    console.log(
      `Created Personal Records sheet with ${swimmerData.length} swimmers`
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      'Error',
      `Failed to create Personal Records: ${e.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    console.error('PR creation error:', e);
  }
}

/**
 * Set up relay events structure with priority configuration
 */
function setupRelayEvents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // Define all relay events with default priorities
    const relayEvents = [
      {
        name: '200 Medley',
        level: 'Varsity',
        strokes: ['Back', 'Breast', 'Fly', 'Free'],
        distance: 50,
        priority: 1,
      },
      {
        name: '200 Medley JV',
        level: 'JV',
        strokes: ['Back', 'Breast', 'Fly', 'Free'],
        distance: 50,
        priority: 2,
      },
      {
        name: '200 Free',
        level: 'Varsity',
        strokes: ['Free', 'Free', 'Free', 'Free'],
        distance: 50,
        priority: 3,
      },
      {
        name: '200 Free JV',
        level: 'JV',
        strokes: ['Free', 'Free', 'Free', 'Free'],
        distance: 50,
        priority: 4,
      },
      {
        name: '400 IM',
        level: 'Varsity',
        strokes: ['Fly', 'Back', 'Breast', 'Free'],
        distance: 100,
        priority: 5,
      },
      {
        name: '400 IM JV',
        level: 'JV',
        strokes: ['Fly', 'Back', 'Breast', 'Free'],
        distance: 100,
        priority: 6,
      },
      {
        name: '350 Free',
        level: 'Varsity',
        strokes: ['Free', 'Free', 'Free', 'Free'],
        distance: 87.5,
        priority: 7,
      },
      {
        name: '400 Free',
        level: 'Varsity',
        strokes: ['Free', 'Free', 'Free', 'Free'],
        distance: 100,
        priority: 8,
      },
      {
        name: '400 Free JV',
        level: 'JV',
        strokes: ['Free', 'Free', 'Free', 'Free'],
        distance: 100,
        priority: 9,
      },
      {
        name: '200 Fly',
        level: 'Varsity',
        strokes: ['Fly', 'Fly', 'Fly', 'Fly'],
        distance: 50,
        priority: 10,
      },
      {
        name: '200 Fly JV',
        level: 'JV',
        strokes: ['Fly', 'Fly', 'Fly', 'Fly'],
        distance: 50,
        priority: 11,
      },
      {
        name: '200 Breast',
        level: 'Varsity',
        strokes: ['Breast', 'Breast', 'Breast', 'Breast'],
        distance: 50,
        priority: 12,
      },
      {
        name: '200 Breast JV',
        level: 'JV',
        strokes: ['Breast', 'Breast', 'Breast', 'Breast'],
        distance: 50,
        priority: 13,
      },
      {
        name: '200 Back',
        level: 'Varsity',
        strokes: ['Back', 'Back', 'Back', 'Back'],
        distance: 50,
        priority: 14,
      },
      {
        name: '200 Back JV',
        level: 'JV',
        strokes: ['Back', 'Back', 'Back', 'Back'],
        distance: 50,
        priority: 15,
      },
      {
        name: '200 Medley Co-ed',
        level: 'Varsity',
        strokes: ['Back', 'Breast', 'Fly', 'Free'],
        distance: 50,
        priority: 16,
      },
      {
        name: '200 Free Frosh',
        level: 'JV',
        strokes: ['Free', 'Free', 'Free', 'Free'],
        distance: 50,
        priority: 17,
      },
      {
        name: 'I-tube',
        level: 'Varsity',
        strokes: ['Special', 'Special', 'Special', 'Special'],
        distance: 0,
        priority: 18,
      },
    ];

    // Get or create Relay Events Config sheet
    let relaySheet = ss.getSheetByName('Relay Events Config');
    if (!relaySheet) {
      relaySheet = ss.insertSheet('Relay Events Config');
    } else {
      relaySheet.clear();
    }

    // Set up headers
    const headers = [
      'Priority',
      'Event Name',
      'Level',
      'Gender',
      'Stroke 1',
      'Stroke 2',
      'Stroke 3',
      'Stroke 4',
      'Distance per Leg',
      'Active',
      'Notes',
    ];
    relaySheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Add relay events data
    const eventData = relayEvents.map(event => [
      event.priority,
      event.name,
      event.level,
      event.name.includes('Co-ed') ? 'Mixed' : 'Both', // Gender assignment
      event.strokes[0],
      event.strokes[1],
      event.strokes[2],
      event.strokes[3],
      event.distance,
      true, // Active by default
      event.name === 'I-tube' ? 'Unknown event type - needs clarification' : '',
    ]);

    relaySheet
      .getRange(2, 1, eventData.length, headers.length)
      .setValues(eventData);

    // Format headers
    const headerRange = relaySheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1c4587');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');

    // Set up data validation for Priority column (1-20)
    const priorityRange = relaySheet.getRange(2, 1, eventData.length, 1);
    const priorityRule = SpreadsheetApp.newDataValidation()
      .requireNumberBetween(1, 20)
      .setAllowInvalid(false)
      .setHelpText('Priority from 1 (highest) to 20 (lowest)')
      .build();
    priorityRange.setDataValidation(priorityRule);

    // Set up data validation for Level column
    const levelRange = relaySheet.getRange(2, 3, eventData.length, 1);
    const levelRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Varsity', 'JV'])
      .setAllowInvalid(false)
      .build();
    levelRange.setDataValidation(levelRule);

    // Set up data validation for Gender column
    const genderRange = relaySheet.getRange(2, 4, eventData.length, 1);
    const genderRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Men', 'Women', 'Both', 'Mixed'])
      .setAllowInvalid(false)
      .setHelpText(
        'Men/Women = gender-specific, Both = separate M/F relays, Mixed = co-ed'
      )
      .build();
    genderRange.setDataValidation(genderRule);

    // Set up data validation for Active column
    const activeRange = relaySheet.getRange(2, 10, eventData.length, 1);
    const activeRule = SpreadsheetApp.newDataValidation()
      .requireCheckbox()
      .build();
    activeRange.setDataValidation(activeRule);

    // Color code by level
    for (let i = 0; i < eventData.length; i++) {
      const rowNum = i + 2;
      const level = eventData[i][2];
      const rowRange = relaySheet.getRange(rowNum, 1, 1, headers.length);

      if (level === 'JV') {
        rowRange.setBackground('#e1f5fe'); // Light blue for JV
      } else {
        rowRange.setBackground('#f3e5f5'); // Light purple for Varsity
      }
    }

    // Auto-resize columns
    relaySheet.autoResizeColumns(1, headers.length);

    // Add section break and non-conventional relays table
    const lastRow = relaySheet.getLastRow();
    const sectionBreakRow = lastRow + 2;
    
    // Section header for non-conventional relays
    relaySheet.getRange(sectionBreakRow, 1).setValue('NON-CONVENTIONAL RELAYS (Different Leg Counts)');
    relaySheet.getRange(sectionBreakRow, 1, 1, 8).merge();
    relaySheet.getRange(sectionBreakRow, 1)
      .setBackground('#ff6d01')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

    // Non-conventional relay headers
    const nonConvHeaders = [
      'Priority',
      'Event Name', 
      'Level',
      'Gender',
      'Number of Legs',
      'Stroke Pattern',
      'Distance per Leg',
      'Active'
    ];
    
    relaySheet.getRange(sectionBreakRow + 1, 1, 1, nonConvHeaders.length).setValues([nonConvHeaders]);
    
    // Format non-conventional headers
    const nonConvHeaderRange = relaySheet.getRange(sectionBreakRow + 1, 1, 1, nonConvHeaders.length);
    nonConvHeaderRange.setBackground('#e65100');
    nonConvHeaderRange.setFontColor('#ffffff');
    nonConvHeaderRange.setFontWeight('bold');

    // Add sample non-conventional relays
    const nonConvRelays = [
      [19, '350 Free', 'Varsity', 'Both', 3, 'Free-Free-Free', 117, true],
      [20, 'I-tube', 'Varsity', 'Both', 6, 'Special-Special-Special-Special-Special-Special', 0, true]
    ];

    relaySheet.getRange(sectionBreakRow + 2, 1, nonConvRelays.length, nonConvHeaders.length).setValues(nonConvRelays);

    // Set up data validation for non-conventional relays
    const nonConvStartRow = sectionBreakRow + 2;
    
    // Priority validation
    const nonConvPriorityRange = relaySheet.getRange(nonConvStartRow, 1, nonConvRelays.length, 1);
    nonConvPriorityRange.setDataValidation(priorityRule);

    // Level validation
    const nonConvLevelRange = relaySheet.getRange(nonConvStartRow, 3, nonConvRelays.length, 1);
    nonConvLevelRange.setDataValidation(levelRule);

    // Gender validation
    const nonConvGenderRange = relaySheet.getRange(nonConvStartRow, 4, nonConvRelays.length, 1);
    nonConvGenderRange.setDataValidation(genderRule);

    // Number of legs validation (3-8 legs)
    const legsRange = relaySheet.getRange(nonConvStartRow, 5, nonConvRelays.length, 1);
    const legsRule = SpreadsheetApp.newDataValidation()
      .requireNumberBetween(3, 8)
      .setAllowInvalid(false)
      .setHelpText('Number of legs from 3 to 8')
      .build();
    legsRange.setDataValidation(legsRule);

    // Active validation
    const nonConvActiveRange = relaySheet.getRange(nonConvStartRow, 8, nonConvRelays.length, 1);
    nonConvActiveRange.setDataValidation(activeRule);

    // Color non-conventional relays
    for (let i = 0; i < nonConvRelays.length; i++) {
      const rowNum = nonConvStartRow + i;
      const level = nonConvRelays[i][2];
      const rowRange = relaySheet.getRange(rowNum, 1, 1, nonConvHeaders.length);

      if (level === 'JV') {
        rowRange.setBackground('#fff3e0'); // Light orange for JV non-conventional
      } else {
        rowRange.setBackground('#fce4ec'); // Light pink for Varsity non-conventional
      }
    }

    // Auto-resize all columns
    relaySheet.autoResizeColumns(1, Math.max(headers.length, nonConvHeaders.length));

    SpreadsheetApp.getUi().alert(
      'Relay Events Config Created!',
      `Created configurable relay events sheet with ${relayEvents.length} conventional relay events and ${nonConvRelays.length} non-conventional relays.\n\n` +
        `Features:\n` +
        `• Priority column (1 = highest priority)\n` +
        `• Gender configuration (Men/Women/Both/Mixed)\n` +
        `• Active checkbox to enable/disable events\n` +
        `• Non-conventional relays with custom leg counts\n` +
        `• Color coding: Blue/Purple = Conventional, Orange/Pink = Non-conventional\n\n` +
        `Adjust priorities and settings, then run "Generate Smart Relay Assignments"!`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    console.log(`Set up ${relayEvents.length} configurable relay events`);
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      'Error',
      `Failed to setup relay events: ${e.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    console.error('Relay setup error:', e);
  }
}

/**
 * Generate smart relay assignments with prioritization and constraints
 */
function generateRelayAssignments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // Get relay configuration
    const relayConfigSheet = ss.getSheetByName('Relay Events Config');
    if (!relayConfigSheet) {
      SpreadsheetApp.getUi().alert(
        'Error',
        'Please run "Setup Relay Events" first to create the configuration sheet.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    const configData = relayConfigSheet.getDataRange().getValues();
    const configHeaders = configData[0];
    
    // Parse both conventional and non-conventional relays
    const conventionalRelays = [];
    const nonConventionalRelays = [];
    let foundNonConventionalSection = false;
    
    for (let i = 1; i < configData.length; i++) {
      const row = configData[i];
      
      // Check if we've hit the non-conventional section
      if (row[0] && row[0].toString().includes('NON-CONVENTIONAL RELAYS')) {
        foundNonConventionalSection = true;
        i++; // Skip the header row
        continue;
      }
      
      // Skip empty rows
      if (!row[0] && !row[1]) continue;
      
      // Only process active events
      const activeColumnIndex = foundNonConventionalSection ? 7 : 9;
      if (row[activeColumnIndex] !== true) continue;
      
      if (foundNonConventionalSection) {
        nonConventionalRelays.push({
          type: 'non-conventional',
          priority: row[0],
          eventName: row[1],
          level: row[2],
          genderConfig: row[3],
          numLegs: row[4],
          strokePattern: row[5],
          distance: row[6],
          active: row[7]
        });
      } else {
        conventionalRelays.push({
          type: 'conventional',
          priority: row[0],
          eventName: row[1],
          level: row[2],
          genderConfig: row[3],
          stroke1: row[4],
          stroke2: row[5],
          stroke3: row[6],
          stroke4: row[7],
          distance: row[8],
          active: row[9],
          notes: row[10]
        });
      }
    }

    // Combine and sort all relay configs by priority
    const allRelayConfigs = [...conventionalRelays, ...nonConventionalRelays];
    allRelayConfigs.sort((a, b) => a.priority - b.priority);

    // Get swimmers data
    const swimmersSheet = ss.getSheetByName('Swimmers');
    if (!swimmersSheet) {
      SpreadsheetApp.getUi().alert(
        'Error',
        'Swimmers sheet not found. Please process tryouts first.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    const swimmersData = swimmersSheet.getDataRange().getValues();
    const swimmers = parseSwimmersWithPRs_(swimmersData);

    // Debug: Log swimmer data
    console.log(`Found ${swimmers.length} swimmers:`);
    swimmers.forEach(s => console.log(`${s.name}: ${s.gender}, ${s.level}`));

    // Check for existing relay assignments and warn user
    let resultsSheet = ss.getSheetByName('Relay Assignments');
    let hasExistingAssignments = false;

    if (resultsSheet) {
      const data = resultsSheet.getDataRange().getValues();
      hasExistingAssignments = data.length > 1; // More than just headers

      if (hasExistingAssignments) {
        const response = SpreadsheetApp.getUi().alert(
          'Overwrite Existing Assignments?',
          'You already have relay assignments. This will:\n\n' +
            '• Create a backup copy in "Relay Assignments Backup"\n' +
            '• Overwrite your current assignments with new smart assignments\n\n' +
            'Do you want to continue?',
          SpreadsheetApp.getUi().ButtonSet.YES_NO
        );

        if (response === SpreadsheetApp.getUi().Button.NO) {
          return; // User cancelled
        }

        // Create backup before overwriting
        createRelayAssignmentsBackup_(ss, resultsSheet);
      }
    }

    // Preserve locked relays from existing assignments
    let lockedRelays = [];
    if (resultsSheet && hasExistingAssignments) {
      const existingData = resultsSheet.getDataRange().getValues();
      const existingHeaders = existingData[0];
      const lockIndex = existingHeaders.indexOf('Lock');
      
      if (lockIndex !== -1) {
        // Find locked relays to preserve
        for (let i = 1; i < existingData.length; i++) {
          const row = existingData[i];
          if (row[lockIndex] === true) {
            lockedRelays.push(row);
            console.log(`Preserving locked relay: ${row[0]} ${row[1]} ${row[2]}`);
          }
        }
      }
    }

    // Get or create Results sheet
    if (!resultsSheet) {
      resultsSheet = ss.insertSheet('Relay Assignments');
    } else {
      resultsSheet.clear();
    }

    // Set up headers with support for non-conventional relays
    const headers = [
      'Event',
      'Level', 
      'Gender',
      'Lock',
      'Leg 1',
      'Leg 1 Time',
      'Leg 2',
      'Leg 2 Time',
      'Leg 3',
      'Leg 3 Time',
      'Leg 4',
      'Leg 4 Time',
      'Leg 5',
      'Leg 5 Time',
      'Leg 6',
      'Leg 6 Time',
      'Leg 7',
      'Leg 7 Time',
      'Leg 8',
      'Leg 8 Time',
      'Total Time',
      'Notes',
    ];

    console.log('Setting headers:', headers);

    // Add instruction note at the top
    const instructionText =
      '💡 After making manual changes to assignments, use "Coach Tools > Refresh Swimmer Assignment Summary" to update the summary.';
    resultsSheet.getRange(1, 1).setValue(instructionText);
    resultsSheet.getRange(1, 1, 1, headers.length).merge();
    resultsSheet
      .getRange(1, 1)
      .setBackground('#e3f2fd')
      .setFontStyle('italic')
      .setWrap(true);
    resultsSheet.getRange(2, 1, 1, headers.length).setValues([headers]);

    // Verify headers were set correctly
    const actualHeaders = resultsSheet.getRange(2, 1, 1, headers.length).getValues()[0];
    console.log('Headers set to:', actualHeaders);
    if (actualHeaders[0] !== 'Event') {
      console.error('WARNING: Event header not set correctly! Expected "Event", got:', actualHeaders[0]);
    }

    // Format headers
    const headerRange = resultsSheet.getRange(2, 1, 1, headers.length);
    headerRange.setBackground('#0d47a1');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');

    // Track swimmer assignments (max 4 relays per swimmer)
    const swimmerAssignments = new Map(); // swimmer name -> array of relay assignments
    const relayResults = [];
    let currentRow = 3; // Start from row 3 now since we added instruction row

    // Process each relay configuration
    for (const config of allRelayConfigs) {
      if (config.type === 'conventional') {
        const {
          priority,
          eventName,
          level,
          genderConfig,
          stroke1,
          stroke2,
          stroke3,
          stroke4,
          distance,
          active,
          notes,
        } = config;

      // Determine which genders to create relays for
      const gendersToProcess = getGendersForRelay_(genderConfig);

      for (const gender of gendersToProcess) {
        // Get eligible swimmers for this relay
        const eligibleSwimmers = swimmers.filter(swimmer => {
          // Level match - be more flexible with level matching
          const swimmerLevel = swimmer.level.toUpperCase();
          const relayLevel = level.toUpperCase();
          if (
            swimmerLevel !== 'VARSITY' &&
            swimmerLevel !== 'V' &&
            swimmerLevel !== 'JV' &&
            swimmerLevel !== 'JUNIOR VARSITY'
          ) {
            return false; // Skip swimmers without clear level
          }

          // Normalize levels for comparison
          const normalizedSwimmerLevel =
            swimmerLevel === 'VARSITY' || swimmerLevel === 'V'
              ? 'VARSITY'
              : 'JV';
          const normalizedRelayLevel = relayLevel.toUpperCase();

          // For mixed gender relays or relays marked as "Both" level, allow more flexibility
          const isFlexibleRelay = gender === 'Mixed' || eventName.toLowerCase().includes('frosh') || eventName.toLowerCase().includes('mixed');
          
          if (!isFlexibleRelay && normalizedSwimmerLevel !== normalizedRelayLevel) {
            return false; // Strict level matching for standard relays
          }
          // For flexible relays, allow both JV and Varsity swimmers

          // Gender match
          if (gender === 'Mixed') {
            return true; // Mixed relays include all genders
          } else if (gender === 'Men' && swimmer.gender !== 'M') {
            return false;
          } else if (gender === 'Women' && swimmer.gender !== 'F') {
            return false;
          }

          return true;
        });

        console.log(
          `${eventName} ${gender} ${level}: Found ${eligibleSwimmers.length} eligible swimmers`
        );

        // Smart assignment with 4-relay constraint
        const selectedSwimmers = selectSwimmersForRelay_(
          eligibleSwimmers,
          swimmerAssignments,
          eventName,
          gender
        );

        if (selectedSwimmers.length >= 4) {
          const relayRow = createConventionalRelayRow_(
            eventName,
            level,
            gender,
            selectedSwimmers,
            [stroke1, stroke2, stroke3, stroke4],
            eligibleSwimmers.length
          );

          relayResults.push(relayRow);

          // Track assignments
          selectedSwimmers.slice(0, 4).forEach(swimmer => {
            if (!swimmerAssignments.has(swimmer.name)) {
              swimmerAssignments.set(swimmer.name, []);
            }
            swimmerAssignments.get(swimmer.name).push(`${eventName} ${gender}`);
          });
        } else if (selectedSwimmers.length > 0) {
          // Partial relay - fill what we can
          const relayRow = createPartialConventionalRelayRow_(
            eventName,
            level,
            gender,
            selectedSwimmers,
            [stroke1, stroke2, stroke3, stroke4]
          );
          relayResults.push(relayRow);

          // Track partial assignments
          selectedSwimmers.forEach(swimmer => {
            if (!swimmerAssignments.has(swimmer.name)) {
              swimmerAssignments.set(swimmer.name, []);
            }
            swimmerAssignments.get(swimmer.name).push(`${eventName} ${gender}`);
          });
        } else {
          // No eligible swimmers
          const relayRow = createEmptyConventionalRelayRow_(eventName, level, gender);
          relayResults.push(relayRow);
        }
      }
      } else if (config.type === 'non-conventional') {
        // Handle non-conventional relays
        const {
          priority,
          eventName,
          level,
          genderConfig,
          numLegs,
          strokePattern,
          distancePerLeg,
          active,
          notes,
        } = config;

        if (!active) continue;

        // Determine which genders to create relays for
        const gendersToProcess = getGendersForRelay_(genderConfig);

        for (const gender of gendersToProcess) {
          // Get eligible swimmers for this relay
          const eligibleSwimmers = swimmers.filter(swimmer => {
            // Level match - be more flexible with level matching
            const swimmerLevel = swimmer.level.toUpperCase();
            const relayLevel = level.toUpperCase();
            if (
              swimmerLevel !== 'VARSITY' &&
              swimmerLevel !== 'V' &&
              swimmerLevel !== 'JV' &&
              swimmerLevel !== 'JUNIOR VARSITY'
            ) {
              return false;
            }

            // Normalize levels for comparison
            const normalizedSwimmerLevel =
              swimmerLevel === 'VARSITY' || swimmerLevel === 'V'
                ? 'VARSITY'
                : 'JV';
            const normalizedRelayLevel = relayLevel.toUpperCase();

            // For mixed gender relays or relays marked as "Both" level, allow more flexibility
            const isFlexibleRelay = gender === 'Mixed' || eventName.toLowerCase().includes('frosh') || eventName.toLowerCase().includes('mixed');
            
            if (!isFlexibleRelay && normalizedSwimmerLevel !== normalizedRelayLevel) {
              return false; // Strict level matching for standard relays
            }
            // For flexible relays, allow both JV and Varsity swimmers

            // Gender match
            if (gender === 'Mixed') {
              return true; // Mixed relays include all genders
            } else if (gender === 'Men' && swimmer.gender !== 'M') {
              return false;
            } else if (gender === 'Women' && swimmer.gender !== 'F') {
              return false;
            }

            return true;
          });

          console.log(
            `${eventName} ${gender} ${level}: Found ${eligibleSwimmers.length} eligible swimmers for ${numLegs}-leg relay`
          );

          // Smart assignment with 4-relay constraint
          const selectedSwimmers = selectSwimmersForNonConventionalRelay_(
            eligibleSwimmers,
            swimmerAssignments,
            eventName,
            gender,
            numLegs
          );

          if (selectedSwimmers.length >= numLegs) {
            const relayRow = createNonConventionalRelayRow_(
              eventName,
              level,
              gender,
              selectedSwimmers,
              strokePattern,
              numLegs,
              eligibleSwimmers.length
            );

            relayResults.push(relayRow);

            // Track assignments
            selectedSwimmers.slice(0, numLegs).forEach(swimmer => {
              if (!swimmerAssignments.has(swimmer.name)) {
                swimmerAssignments.set(swimmer.name, []);
              }
              swimmerAssignments.get(swimmer.name).push(`${eventName} ${gender}`);
            });
          } else if (selectedSwimmers.length > 0) {
            // Partial relay - fill what we can
            const relayRow = createPartialNonConventionalRelayRow_(
              eventName,
              level,
              gender,
              selectedSwimmers,
              strokePattern,
              numLegs
            );
            relayResults.push(relayRow);

            // Track partial assignments
            selectedSwimmers.forEach(swimmer => {
              if (!swimmerAssignments.has(swimmer.name)) {
                swimmerAssignments.set(swimmer.name, []);
              }
              swimmerAssignments.get(swimmer.name).push(`${eventName} ${gender}`);
            });
          } else {
            // No eligible swimmers
            const relayRow = createEmptyNonConventionalRelayRow_(eventName, level, gender, numLegs);
            relayResults.push(relayRow);
          }
        }
      }
    }

    // Add preserved locked relays to the results
    if (lockedRelays.length > 0) {
      console.log(`Adding ${lockedRelays.length} locked relays back to results`);
      
      // Make sure locked relays have the correct number of columns to match new structure
      lockedRelays.forEach(lockedRow => {
        if (lockedRow.length < headers.length) {
          // Pad with empty strings to match new header structure
          while (lockedRow.length < headers.length) {
            lockedRow.push('');
          }
        }
        // Ensure Lock column is set to true
        lockedRow[3] = true;
        relayResults.push(lockedRow);
        
        // Track locked relay swimmer assignments
        const event = lockedRow[0];
        const level = lockedRow[1];
        const gender = lockedRow[2];
        const relayName = `${event} (${level} ${gender})`;
        
        // Check all possible leg positions for swimmers
        for (let i = 4; i < headers.length - 2; i += 2) { // Skip Lock, Total Time, Notes columns
          const swimmer = lockedRow[i];
          if (swimmer && swimmer.toString().trim() !== '') {
            if (!swimmerAssignments.has(swimmer)) {
              swimmerAssignments.set(swimmer, []);
            }
            swimmerAssignments.get(swimmer).push(relayName);
          }
        }
      });
    }

    // Write all results to sheet
    if (relayResults.length > 0) {
      resultsSheet
        .getRange(3, 1, relayResults.length, headers.length)
        .setValues(relayResults);

      // Apply dropdowns and formatting
      setupRelayDropdownsAndValidation_(
        resultsSheet,
        swimmers,
        relayResults.length + 2 // +2 to account for instruction row and header row
      );
    }

    // Create swimmer assignment summary
    createSwimmerAssignmentSummary_(ss, swimmerAssignments);

    SpreadsheetApp.getUi().alert(
      'Smart Relay Assignments Created!',
      `Generated ${relayResults.length} relay assignments with enhanced features.\n\n` +
        `New Features:\n` +
        `• Support for non-conventional relays (3-8 legs)\n` +
        `• Lock checkbox to preserve specific assignments\n` +
        `• ${lockedRelays.length} locked relays preserved from previous assignments\n\n` +
        `Existing Features:\n` +
        `• 4-relay maximum per swimmer enforced\n` +
        `• Load balancing: swimmers with fewer relays preferred\n` +
        `• Dropdown lists for easy editing\n` +
        `• Red highlighting: same relay conflicts\n` +
        `• Orange highlighting: >4 total relay violations\n\n` +
        `💡 Tip: Check the "Lock" box for any relay you want to preserve during future regenerations.\n` +
        `After manual changes, use "Coach Tools > Refresh Swimmer Assignment Summary" to update the summary.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    console.log(`Generated ${relayResults.length} relay assignments`);
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      'Error',
      `Failed to generate relay assignments: ${e.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    console.error('Relay assignment error:', e);
  }
}

/**
 * Helper function to determine which genders to process for a relay
 */
function getGendersForRelay_(genderConfig) {
  switch (genderConfig) {
    case 'Men':
      return ['Men'];
    case 'Women':
      return ['Women'];
    case 'Both':
      return ['Men', 'Women'];
    case 'Mixed':
      return ['Mixed'];
    default:
      return ['Men', 'Women']; // Default to both
  }
}

/**
 * Generate optimal relay team using smart assignment algorithm
 */
function generateOptimalRelay_(
  swimmers,
  eventName,
  level,
  gender,
  strokes,
  distance,
  existingAssignments
) {
  // Filter eligible swimmers
  const eligibleSwimmers = swimmers.filter(swimmer => {
    // Level match
    if (swimmer.level !== level) return false;

    // Gender match
    if (gender === 'Mixed') {
      // Mixed relays need both genders - for now, include all
      return true;
    } else if (gender === 'Men' && swimmer.gender !== 'M') {
      return false;
    } else if (gender === 'Women' && swimmer.gender !== 'F') {
      return false;
    }

    // Check if swimmer is already in 4 relays
    const currentAssignments = existingAssignments.get(swimmer.name) || [];
    if (currentAssignments.length >= 4) return false;

    return true;
  });

  if (eligibleSwimmers.length < 4) {
    return {
      success: false,
      notes: `Insufficient eligible swimmers (${eligibleSwimmers.length}/4 needed)`,
    };
  }

  // Step 1: Find best anchor (leg 4) swimmer
  const anchor = findBestSwimmerForStroke_(
    eligibleSwimmers,
    strokes[3],
    distance,
    []
  );
  if (!anchor) {
    return {
      success: false,
      notes: `No suitable anchor swimmer found for ${strokes[3]}`,
    };
  }

  const usedSwimmers = [anchor.name];

  // Step 2: Find leg 1 swimmer
  const leg1 = findBestSwimmerForStroke_(
    eligibleSwimmers,
    strokes[0],
    distance,
    usedSwimmers
  );
  if (!leg1) {
    return {
      success: false,
      notes: `No suitable swimmer found for leg 1 (${strokes[0]})`,
    };
  }
  usedSwimmers.push(leg1.name);

  // Step 3: Find leg 3 swimmer
  const leg3 = findBestSwimmerForStroke_(
    eligibleSwimmers,
    strokes[2],
    distance,
    usedSwimmers
  );
  if (!leg3) {
    return {
      success: false,
      notes: `No suitable swimmer found for leg 3 (${strokes[2]})`,
    };
  }
  usedSwimmers.push(leg3.name);

  // Step 4: Find leg 2 swimmer
  const leg2 = findBestSwimmerForStroke_(
    eligibleSwimmers,
    strokes[1],
    distance,
    usedSwimmers
  );
  if (!leg2) {
    return {
      success: false,
      notes: `No suitable swimmer found for leg 2 (${strokes[1]})`,
    };
  }

  // Calculate total time
  const totalTime = calculateRelayTime_([leg1, leg2, leg3, anchor]);

  return {
    success: true,
    leg1: leg1,
    leg2: leg2,
    leg3: leg3,
    leg4: anchor,
    totalTime: totalTime,
    notes: null,
  };
}

/**
 * Find best available swimmer for specific stroke and distance
 */
function findBestSwimmerForStroke_(swimmers, stroke, distance, usedSwimmers) {
  const availableSwimmers = swimmers.filter(
    swimmer => !usedSwimmers.includes(swimmer.name)
  );

  if (availableSwimmers.length === 0) return null;

  // Find best time for this stroke/distance combination
  let bestSwimmer = null;
  let bestTime = Infinity;

  for (const swimmer of availableSwimmers) {
    const time = getBestTimeForStrokeDistance_(swimmer, stroke, distance);
    if (time !== null && time < bestTime) {
      bestTime = time;
      bestSwimmer = {
        name: swimmer.name,
        time: formatTime_(time),
        rawTime: time,
      };
    }
  }

  // If no exact match, find closest stroke match
  if (!bestSwimmer) {
    for (const swimmer of availableSwimmers) {
      const time = getClosestStrokeTime_(swimmer, stroke);
      if (time !== null && time < bestTime) {
        bestTime = time;
        bestSwimmer = {
          name: swimmer.name,
          time: formatTime_(time) + '*',
          rawTime: time,
        };
      }
    }
  }

  return bestSwimmer;
}

/**
 * Get best time for specific stroke and distance
 */
function getBestTimeForStrokeDistance_(swimmer, stroke, distance) {
  const strokeMap = {
    Free: ['50 Free', '100 Free', '200 Free', '500 Free'],
    Back: ['100 Back'],
    Breast: ['100 Breast'],
    Fly: ['100 Fly'],
    IM: ['200 IM'],
  };

  const events = strokeMap[stroke] || [];
  let bestTime = null;

  for (const event of events) {
    if (swimmer.prs && swimmer.prs[event]) {
      const time = swimmer.prs[event];
      if (bestTime === null || time < bestTime) {
        bestTime = time;
      }
    }
  }

  return bestTime;
}

/**
 * Get closest stroke time when exact match not available
 */
function getClosestStrokeTime_(swimmer, stroke) {
  // Fallback to freestyle if stroke not available
  if (stroke !== 'Free' && swimmer.prs) {
    const freeEvents = ['50 Free', '100 Free', '200 Free'];
    for (const event of freeEvents) {
      if (swimmer.prs[event]) {
        return swimmer.prs[event];
      }
    }
  }
  return null;
}

/**
 * Calculate total relay time
 */
function calculateRelayTime_(legs) {
  const totalSeconds = legs.reduce((sum, leg) => sum + (leg.rawTime || 0), 0);
  return formatTime_(totalSeconds);
}

/**
 * Parse swimmers with their personal records
 */
function parseSwimmersWithPRs_(swimmersData) {
  const headers = swimmersData[0];
  const swimmers = [];

  for (let i = 1; i < swimmersData.length; i++) {
    const row = swimmersData[i];
    if (!row[0]) continue; // Skip empty rows

    const swimmer = {
      name: row[0],
      gradYear: row[1],
      gender: row[2],
      level: row[3],
      notes: row[4],
      personalRecords: {},
    };

    // Get PRs from Personal Records sheet if available
    swimmer.personalRecords = getSwimmerPRs_(swimmer.name);

    swimmers.push(swimmer);
  }

  return swimmers;
}

/**
 * Get swimmer's personal records
 */
function getSwimmerPRs_(swimmerName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prSheet = ss.getSheetByName('Personal Records');
  if (!prSheet) return {};

  const prData = prSheet.getDataRange().getValues();
  const headers = prData[0];

  for (let i = 1; i < prData.length; i++) {
    if (prData[i][0] === swimmerName) {
      const prs = {};
      for (let j = 1; j < headers.length; j++) {
        if (prData[i][j] && prData[i][j] !== '') {
          prs[headers[j]] = parseTimeToSeconds_(prData[i][j]);
        }
      }
      return prs;
    }
  }

  return {};
}

/**
 * Format relay results sheet
 */
function formatRelayResults_(sheet, numRows) {
  // Auto-resize columns for all columns
  sheet.autoResizeColumns(1, Math.max(22, sheet.getLastColumn()));

  // Apply alternating row colors
  for (let i = 2; i <= numRows; i++) {
    const rowRange = sheet.getRange(i, 1, 1, Math.max(22, sheet.getLastColumn()));
    if (i % 2 === 0) {
      rowRange.setBackground('#f8f9fa');
    }

    // Color code by level and gender
    const level = sheet.getRange(i, 2).getValue();
    const gender = sheet.getRange(i, 3).getValue();

    if (level === 'Varsity') {
      if (gender === 'Men') {
        rowRange.setBackground('#e3f2fd'); // Light blue
      } else if (gender === 'Women') {
        rowRange.setBackground('#fce4ec'); // Light pink
      } else {
        rowRange.setBackground('#f3e5f5'); // Light purple for mixed
      }
    } else {
      // JV
      if (gender === 'Men') {
        rowRange.setBackground('#e8f5e8'); // Light green
      } else if (gender === 'Women') {
        rowRange.setBackground('#fff3e0'); // Light orange
      } else {
        rowRange.setBackground('#f5f5f5'); // Light gray for mixed
      }
    }
  }
}

/**
 * Create swimmer assignment summary
 */
function createSwimmerAssignmentSummary_(ss, swimmerAssignments) {
  let summarySheet = ss.getSheetByName('Swimmer Assignment Summary');
  if (!summarySheet) {
    summarySheet = ss.insertSheet('Swimmer Assignment Summary');
  } else {
    summarySheet.clear();
  }

  const headers = ['Swimmer', 'Number of Relays', 'Relay Assignments'];
  summarySheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format headers
  const headerRange = summarySheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#1565c0');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');

  // Create summary data
  const summaryData = [];
  for (const [swimmer, assignments] of swimmerAssignments) {
    summaryData.push([swimmer, assignments.length, assignments.join(', ')]);
  }

  // Sort by number of relays (descending) then by name
  summaryData.sort((a, b) => {
    if (b[1] !== a[1]) return b[1] - a[1];
    return a[0].localeCompare(b[0]);
  });

  if (summaryData.length > 0) {
    summarySheet
      .getRange(2, 1, summaryData.length, headers.length)
      .setValues(summaryData);

    // Color code by number of relays
    for (let i = 0; i < summaryData.length; i++) {
      const rowNum = i + 2;
      const numRelays = summaryData[i][1];
      const rowRange = summarySheet.getRange(rowNum, 1, 1, headers.length);

      if (numRelays >= 4) {
        rowRange.setBackground('#ffcdd2'); // Light red - maxed out
      } else if (numRelays === 3) {
        rowRange.setBackground('#fff3e0'); // Light orange - nearly full
      } else if (numRelays === 2) {
        rowRange.setBackground('#e8f5e8'); // Light green - moderate
      } else {
        rowRange.setBackground('#e3f2fd'); // Light blue - light load
      }
    }
  }

  summarySheet.autoResizeColumns(1, headers.length);
}

/**
 * Refresh swimmer assignment summary from existing Relay Assignments sheet
 */
function refreshSwimmerAssignmentSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // Check if Relay Assignments sheet exists
    const relaySheet = ss.getSheetByName('Relay Assignments');
    if (!relaySheet) {
      SpreadsheetApp.getUi().alert(
        'No Relay Assignments Found',
        'Please generate relay assignments first using "Generate Smart Relay Assignments".',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    // Get relay assignments data
    const data = relaySheet.getDataRange().getValues();
    if (data.length <= 1) {
      SpreadsheetApp.getUi().alert(
        'Empty Relay Assignments',
        'The Relay Assignments sheet appears to be empty. Please generate assignments first.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    // Parse relay assignments to rebuild swimmer assignments map
    const swimmerAssignments = new Map();
    
    // Find the actual header row - skip any instructional text at the top
    let headerRowIndex = 0;
    let headers = data[0];
    
    // Check if first row looks like an instruction (contains emoji or very long text)
    const firstCellText = (data[0][0] || '').toString().trim();
    if (firstCellText.includes('💡') || firstCellText.length > 50 || !firstCellText.includes('Event')) {
      headerRowIndex = 1;
      headers = data[1];
    }

    // Find column indices for swimmer positions - updated for new structure
    // Trim headers to handle any extra whitespace
    const trimmedHeaders = headers.map(h => (h || '').toString().trim());
    const leg1Index = trimmedHeaders.indexOf('Leg 1');
    const leg2Index = trimmedHeaders.indexOf('Leg 2');
    const leg3Index = trimmedHeaders.indexOf('Leg 3');
    const leg4Index = trimmedHeaders.indexOf('Leg 4');
    const leg5Index = trimmedHeaders.indexOf('Leg 5');
    const leg6Index = trimmedHeaders.indexOf('Leg 6');
    const leg7Index = trimmedHeaders.indexOf('Leg 7');
    const leg8Index = trimmedHeaders.indexOf('Leg 8');
    const eventIndex = trimmedHeaders.indexOf('Event');
    const levelIndex = trimmedHeaders.indexOf('Level');
    const genderIndex = trimmedHeaders.indexOf('Gender');

    console.log('Column indices:', {
      leg1Index,
      leg2Index,
      leg3Index,
      leg4Index,
      leg5Index,
      leg6Index,
      leg7Index,
      leg8Index,
      eventIndex,
      levelIndex,
      genderIndex,
    });

    if (leg1Index === -1 || eventIndex === -1) {
      SpreadsheetApp.getUi().alert(
        'Invalid Sheet Format',
        'The Relay Assignments sheet format is not recognized. Please regenerate assignments.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    // Process each relay assignment row (starting from row after headers)
    for (let i = headerRowIndex + 1; i < data.length; i++) {
      const row = data[i];
      const event = row[eventIndex];
      const level = row[levelIndex];
      const gender = row[genderIndex];

      if (!event) continue; // Skip empty rows

      const relayName = `${event} (${level} ${gender})`;

      // Add each swimmer to their assignments - handle up to 8 legs
      [leg1Index, leg2Index, leg3Index, leg4Index, leg5Index, leg6Index, leg7Index, leg8Index].forEach(legIndex => {
        if (legIndex !== -1 && row[legIndex]) {
          const swimmer = row[legIndex].toString().trim();
          if (swimmer && swimmer !== '') {
            if (!swimmerAssignments.has(swimmer)) {
              swimmerAssignments.set(swimmer, []);
            }
            swimmerAssignments.get(swimmer).push(relayName);
          }
        }
      });
    }

    console.log(
      'Found assignments for swimmers:',
      Array.from(swimmerAssignments.keys())
    );

    // Create/update the swimmer assignment summary
    createSwimmerAssignmentSummary_(ss, swimmerAssignments);

    SpreadsheetApp.getUi().alert(
      'Summary Refreshed!',
      `Swimmer Assignment Summary has been refreshed with ${swimmerAssignments.size} swimmers.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    console.log(
      `Refreshed swimmer assignment summary with ${swimmerAssignments.size} swimmers`
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      'Error',
      `Failed to refresh swimmer assignment summary: ${e.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    console.error('Refresh summary error:', e);
  }
}

/**
 * Create backup of existing relay assignments
 */
/**
 * Validates and fixes the Relay Assignments sheet headers if they're missing or corrupted
 */
function validateRelayAssignmentsHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Relay Assignments');
  
  if (!sheet) {
    console.log('No Relay Assignments sheet found');
    return false;
  }

  const expectedHeaders = [
    'Event',
    'Level', 
    'Gender',
    'Lock',
    'Leg 1',
    'Leg 1 Time',
    'Leg 2',
    'Leg 2 Time',
    'Leg 3',
    'Leg 3 Time',
    'Leg 4',
    'Leg 4 Time',
    'Leg 5',
    'Leg 5 Time',
    'Leg 6',
    'Leg 6 Time',
    'Leg 7',
    'Leg 7 Time',
    'Leg 8',
    'Leg 8 Time',
    'Total Time',
    'Notes',
  ];

  // Check if sheet has data
  if (sheet.getLastRow() < 2) {
    console.log('Relay Assignments sheet is empty');
    return false;
  }

  // Find the header row (could be row 1 or 2 if there's an instruction row)
  let headerRow = 1;
  const row1 = sheet.getRange(1, 1, 1, expectedHeaders.length).getValues()[0];
  const row2 = sheet.getRange(2, 1, 1, expectedHeaders.length).getValues()[0];

  if (row1[0] === 'Event') {
    headerRow = 1;
  } else if (row2[0] === 'Event') {
    headerRow = 2;
  } else {
    console.log('Headers need fixing - Event header not found in expected location');
    
    // Try to find headers by looking for "Leg 1" pattern
    if (row1.includes('Leg 1')) {
      headerRow = 1;
    } else if (row2.includes('Leg 1')) {
      headerRow = 2;
    } else {
      console.error('Cannot locate header row');
      return false;
    }

    // Fix the headers
    console.log('Fixing headers at row', headerRow);
    sheet.getRange(headerRow, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      'Headers Fixed',
      'The "Event" header was missing from your Relay Assignments sheet and has been restored. ' +
      'This should fix any issues with creating relay entry sheets.',
      ui.ButtonSet.OK
    );
    
    return true;
  }

  console.log('Headers are correct');
  return true;
}

function createRelayAssignmentsBackup_(ss, sourceSheet) {
  try {
    // Remove existing backup if it exists
    const existingBackup = ss.getSheetByName('Relay Assignments Backup');
    if (existingBackup) {
      ss.deleteSheet(existingBackup);
    }

    // Create new backup by copying the source sheet
    const backupSheet = sourceSheet.copyTo(ss);
    backupSheet.setName('Relay Assignments Backup');

    // Add timestamp to indicate when backup was created
    const timestamp = new Date().toLocaleString();
    backupSheet.insertRowBefore(1);
    backupSheet.getRange(1, 1).setValue(`Backup created: ${timestamp}`);
    backupSheet.getRange(1, 1).setBackground('#fff3e0').setFontStyle('italic');

    console.log('Created relay assignments backup');
  } catch (e) {
    console.error('Failed to create backup:', e);
    // Don't fail the main operation if backup fails
  }
}

/**
 * Select swimmers for a relay while respecting 4-relay maximum constraint
 */
function selectSwimmersForRelay_(
  eligibleSwimmers,
  swimmerAssignments,
  eventName,
  gender
) {
  // First try swimmers who are under the 4-relay limit
  const preferredSwimmers = eligibleSwimmers.filter(swimmer => {
    const currentAssignments = swimmerAssignments.get(swimmer.name) || [];
    return currentAssignments.length < 4;
  });

  // Sort all eligible swimmers by current relay count (ascending) to balance load
  const sortedSwimmers = [...eligibleSwimmers].sort((a, b) => {
    const aCount = (swimmerAssignments.get(a.name) || []).length;
    const bCount = (swimmerAssignments.get(b.name) || []).length;
    if (aCount !== bCount) return aCount - bCount; // Prefer swimmers with fewer relays
    return a.name.localeCompare(b.name); // Stable sort by name
  });

  // Use preferred swimmers first, then fill remaining spots with any available swimmers
  const selectedSwimmers = [];
  
  // Add swimmers under 4-relay limit first
  selectedSwimmers.push(...preferredSwimmers.slice(0, 4));
  
  // If we need more swimmers and don't have enough under the limit, add more
  if (selectedSwimmers.length < 4) {
    const remainingNeeded = 4 - selectedSwimmers.length;
    const additionalSwimmers = sortedSwimmers
      .filter(swimmer => !selectedSwimmers.includes(swimmer))
      .slice(0, remainingNeeded);
    selectedSwimmers.push(...additionalSwimmers);
  }

  console.log(
    `  ${eventName} ${gender}: ${selectedSwimmers.length} swimmers selected (${preferredSwimmers.length} under 4-relay limit, ${selectedSwimmers.length - preferredSwimmers.slice(0, 4).length} over limit)`
  );

  // Return up to 4 swimmers
  return selectedSwimmers;
}

/**
 * Get tentative time for a swimmer in a stroke (simplified)
 */
function getTentativeTime_(swimmer, stroke) {
  if (!swimmer.personalRecords || !stroke) return '';

  // Simple stroke matching - look for closest match
  const strokeLower = stroke.toLowerCase();
  for (const [event, time] of Object.entries(swimmer.personalRecords)) {
    if (
      event.toLowerCase().includes(strokeLower) ||
      (strokeLower.includes('free') && event.toLowerCase().includes('free')) ||
      (strokeLower.includes('back') && event.toLowerCase().includes('back')) ||
      (strokeLower.includes('breast') &&
        event.toLowerCase().includes('breast')) ||
      (strokeLower.includes('fly') && event.toLowerCase().includes('fly'))
    ) {
      return formatTime_(time);
    }
  }

  return 'TBD';
}

/**
 * Setup dropdowns and validation for relay assignments
 */
function setupRelayDropdownsAndValidation_(sheet, swimmers, numRows) {
  // Create dropdown lists for each level and gender combination
  const swimmersByGroup = {};

  swimmers.forEach(swimmer => {
    // Normalize level
    const level =
      swimmer.level.toUpperCase() === 'VARSITY' ||
      swimmer.level.toUpperCase() === 'V'
        ? 'VARSITY'
        : 'JV';
    const gender = swimmer.gender;
    const key = `${level}_${gender}`;

    if (!swimmersByGroup[key]) {
      swimmersByGroup[key] = [];
    }
    swimmersByGroup[key].push(swimmer.name);
  });

  // Add empty option for each group
  Object.keys(swimmersByGroup).forEach(key => {
    swimmersByGroup[key].unshift(''); // Add empty option at beginning
  });

  // Set up data validation for swimmer columns - updated for new column structure
  // With Lock column added, swimmer columns are now: 5, 7, 9, 11, 13, 15, 17, 19 (Legs 1-8)
  const swimmerColumns = [5, 7, 9, 11, 13, 15, 17, 19]; // Leg 1-8 swimmer columns

  for (let row = 3; row <= numRows; row++) {
    // Start from row 3 due to instruction row
    const level = sheet.getRange(row, 2).getValue();
    const gender = sheet.getRange(row, 3).getValue();

    // Set up Lock column validation (column 4)
    const lockCell = sheet.getRange(row, 4);
    const lockRule = SpreadsheetApp.newDataValidation()
      .requireCheckbox()
      .setHelpText('Check to prevent this relay from being regenerated')
      .build();
    lockCell.setDataValidation(lockRule);

    // Determine which swimmers are eligible
    let eligibleSwimmers = [];
    if (gender === 'Mixed') {
      // For mixed relays, include all swimmers of the appropriate level
      const varsityKey = `VARSITY_M`;
      const jvKey = `JV_M`;
      const femaleVarsityKey = `VARSITY_F`;
      const femaleJvKey = `JV_F`;

      if (level.toUpperCase() === 'VARSITY') {
        eligibleSwimmers = [
          ...(swimmersByGroup[varsityKey] || []),
          ...(swimmersByGroup[femaleVarsityKey] || []),
        ];
      } else {
        eligibleSwimmers = [
          ...(swimmersByGroup[jvKey] || []),
          ...(swimmersByGroup[femaleJvKey] || []),
        ];
      }
    } else {
      const genderCode = gender === 'Men' ? 'M' : 'F';
      const key = `${level.toUpperCase()}_${genderCode}`;
      eligibleSwimmers = swimmersByGroup[key] || [''];
    }

    // Remove duplicates and sort
    eligibleSwimmers = [...new Set(eligibleSwimmers)].sort();

    // Apply dropdown validation to each swimmer column
    swimmerColumns.forEach(col => {
      const cell = sheet.getRange(row, col);
      if (eligibleSwimmers.length > 1) {
        // More than just empty option
        const rule = SpreadsheetApp.newDataValidation()
          .requireValueInList(eligibleSwimmers)
          .setAllowInvalid(false)
          .setHelpText(
            `Select from ${eligibleSwimmers.length - 1} eligible swimmers`
          )
          .build();
        cell.setDataValidation(rule);
      }
    });
  }

  // Set up conditional formatting for validation
  setupRelayValidationFormatting_(sheet, numRows);
}

/**
 * Setup conditional formatting for relay validation
 */
function setupRelayValidationFormatting_(sheet, numRows) {
  // Clear existing conditional formatting
  sheet.clearConditionalFormatRules();

  const rules = [];

  // Red highlight for swimmers in multiple legs of same relay
  for (let row = 3; row <= numRows; row++) {
    // Start from row 3 due to instruction row
    // Updated swimmer columns for new structure: 5,7,9,11,13,15,17,19 (Legs 1-8)
    const swimmerColumns = [5, 7, 9, 11, 13, 15, 17, 19]; // Leg columns

    swimmerColumns.forEach((col, index) => {
      const otherCols = swimmerColumns.filter((_, i) => i !== index);

      otherCols.forEach(otherCol => {
        const formula = `=AND(${sheet.getRange(row, col).getA1Notation()}<>"", ${sheet.getRange(row, col).getA1Notation()}=${sheet.getRange(row, otherCol).getA1Notation()})`;

        const rule = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(formula)
          .setBackground('#ffcdd2') // Light red
          .setFontColor('#d32f2f') // Dark red
          .setRanges([sheet.getRange(row, col)])
          .build();

        rules.push(rule);
      });
    });
  }

  // Orange highlight for swimmers in more than 4 relays total
  // Updated to handle all 8 leg columns
  const swimmerColumns = [5, 7, 9, 11, 13, 15, 17, 19]; // Leg columns

  for (let row = 3; row <= numRows; row++) {
    // Start from row 3 due to instruction row
    swimmerColumns.forEach(col => {
      // Create a formula that counts how many times this swimmer appears in the entire sheet
      const cellRef = sheet.getRange(row, col).getA1Notation();
      const countRefs = swimmerColumns.map(c => {
        const columnLetter = String.fromCharCode(64 + c); // Convert column number to letter
        return `COUNTIF(${columnLetter}3:${columnLetter}1000,${cellRef})`;
      }).join('+');
      
      const formula = `=AND(${cellRef}<>"", ${countRefs}>4)`;

      const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(formula)
        .setBackground('#ffe0b2') // Light orange
        .setFontColor('#f57c00') // Dark orange
        .setRanges([sheet.getRange(row, col)])
        .build();

      rules.push(rule);
    });
  }

  sheet.setConditionalFormatRules(rules);

  // Add note about validation in the instruction row
  const noteRange = sheet.getRange(1, 14); // Column N, row 1
  noteRange.setValue('Red = same relay conflict, Orange = >4 relay limit');
  noteRange.setFontSize(10);
  noteRange.setFontColor('#666666');
}

/**
 * Parse swimmer data from Swimmers sheet
 */
function parseSwimmerData_(data) {
  const headers = data[0];
  const swimmers = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0]) {
      // Has name
      swimmers.push({
        name: row[0].toString().trim(),
        gradYear: row[1],
        gender: row[2].toString().toUpperCase().trim(),
        level: row[3].toString().toUpperCase().trim(), // V or JV
        notes: row[4] ? row[4].toString() : '',
      });
    }
  }

  return swimmers;
}

/**
 * Parse PR data from Personal Records sheet
 */
function parsePRData_(data) {
  const headers = data[0];
  const prs = new Map(); // swimmer name -> event times

  // Find event columns (skip Name and Gender)
  const eventColumns = [];
  for (let col = 0; col < headers.length; col++) {
    const header = headers[col].toString().trim();
    if (
      (header !== 'Name' && header !== 'Gender' && header.includes('Free')) ||
      header.includes('Back') ||
      header.includes('Breast') ||
      header.includes('Fly') ||
      header.includes('IM')
    ) {
      eventColumns.push({ col, event: header });
    }
  }

  // Parse swimmer PRs
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = row[0] ? row[0].toString().trim() : '';

    if (name) {
      const swimmerPRs = {};
      eventColumns.forEach(eventCol => {
        const timeValue = row[eventCol.col];
        if (timeValue && timeValue !== '') {
          const timeInSeconds = parseTimeToSeconds_(timeValue);
          if (timeInSeconds > 0) {
            swimmerPRs[eventCol.event] = timeInSeconds;
          }
        }
      });
      prs.set(name, swimmerPRs);
    }
  }

  return prs;
}

/**
 * Parse relay events from Relay Events sheet
 */
function parseRelayEvents_(data) {
  const relayEvents = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0]) {
      relayEvents.push({
        name: row[0].toString().trim(),
        level: row[1].toString().trim(),
        strokes: [row[2], row[3], row[4], row[5]],
        distance: row[6] || 50,
        notes: row[7] || '',
      });
    }
  }

  return relayEvents;
}

/**
 * Generate relay assignments with constraints
 */
function generateRelayAssignments_(swimmers, prs, relayEvents) {
  const assignments = [];
  const swimmerEventCount = new Map(); // Track events per swimmer

  // Initialize event counts
  swimmers.forEach(swimmer => {
    swimmerEventCount.set(swimmer.name, 0);
  });

  // Separate swimmers by gender and level
  const maleVarsity = swimmers.filter(s => s.gender === 'M' && s.level === 'V');
  const maleJV = swimmers.filter(s => s.gender === 'M' && s.level === 'JV');
  const femaleVarsity = swimmers.filter(
    s => s.gender === 'F' && s.level === 'V'
  );
  const femaleJV = swimmers.filter(s => s.gender === 'F' && s.level === 'JV');

  // Process each relay event
  relayEvents.forEach(relay => {
    // Determine eligible swimmers
    let eligibleSwimmers = [];

    if (relay.name.includes('Co-ed')) {
      // Co-ed relays - mix genders, varsity level swimmers
      eligibleSwimmers = [...maleVarsity, ...femaleVarsity];
    } else if (relay.level === 'JV') {
      // JV relays - JV swimmers only, separate by gender
      eligibleSwimmers =
        relay.name.toLowerCase().includes('women') ||
        relay.name.toLowerCase().includes('girls')
          ? femaleJV
          : relay.name.toLowerCase().includes('men') ||
              relay.name.toLowerCase().includes('boys')
            ? maleJV
            : [...maleJV, ...femaleJV];
    } else {
      // Varsity relays - Varsity swimmers can swim, separate by gender
      eligibleSwimmers =
        relay.name.toLowerCase().includes('women') ||
        relay.name.toLowerCase().includes('girls')
          ? femaleVarsity
          : relay.name.toLowerCase().includes('men') ||
              relay.name.toLowerCase().includes('boys')
            ? maleVarsity
            : [...maleVarsity, ...femaleVarsity];
    }

    // For now, assume mixed gender for most relays if not specified
    if (eligibleSwimmers.length === 0) {
      if (relay.level === 'JV') {
        eligibleSwimmers = [...maleJV, ...femaleJV];
      } else {
        eligibleSwimmers = [...maleVarsity, ...femaleVarsity];
      }
    }

    // Assign swimmers to this relay
    const relayTeam = assignSwimmersToRelay_(
      relay,
      eligibleSwimmers,
      prs,
      swimmerEventCount
    );

    if (relayTeam.length === 4) {
      assignments.push({
        eventName: relay.name,
        level: relay.level,
        swimmers: relayTeam,
      });

      // Update event counts
      relayTeam.forEach(swimmer => {
        const currentCount = swimmerEventCount.get(swimmer.name) || 0;
        swimmerEventCount.set(swimmer.name, currentCount + 1);
      });
    }
  });

  return assignments;
}

/**
 * Assign 4 swimmers to a specific relay based on their times and constraints
 */
function assignSwimmersToRelay_(
  relay,
  eligibleSwimmers,
  prs,
  swimmerEventCount
) {
  const relayTeam = [];

  // For each leg of the relay, find the best available swimmer
  relay.strokes.forEach((stroke, legIndex) => {
    // Find best swimmer for this stroke who isn't already on this relay
    let bestSwimmer = null;
    let bestTime = Infinity;

    eligibleSwimmers.forEach(swimmer => {
      // Skip if already on this relay
      if (relayTeam.find(member => member.name === swimmer.name)) return;

      // Prefer swimmers with fewer than 4 events
      const eventCount = swimmerEventCount.get(swimmer.name) || 0;
      if (eventCount >= 4) return; // Skip overloaded swimmers

      // Get best time for this stroke distance
      const swimmerPRs = prs.get(swimmer.name) || {};
      let relevantTime = null;

      // Map stroke to PR event name (this is simplified - you may need to adjust)
      const strokeDistance = relay.distance;
      if (stroke === 'Free') {
        relevantTime =
          strokeDistance <= 50
            ? swimmerPRs['50 Free']
            : strokeDistance <= 100
              ? swimmerPRs['100 Free']
              : strokeDistance <= 200
                ? swimmerPRs['200 Free']
                : swimmerPRs['500 Free'];
      } else if (stroke === 'Back') {
        relevantTime = swimmerPRs['100 Back'];
      } else if (stroke === 'Breast') {
        relevantTime = swimmerPRs['100 Breast'];
      } else if (stroke === 'Fly') {
        relevantTime = swimmerPRs['100 Fly'];
      }

      if (relevantTime && relevantTime < bestTime) {
        bestTime = relevantTime;
        bestSwimmer = swimmer;
      }
    });

    if (bestSwimmer) {
      relayTeam.push({
        name: bestSwimmer.name,
        gender: bestSwimmer.gender,
        level: bestSwimmer.level,
        stroke: stroke,
        leg: legIndex + 1,
        time: bestTime,
      });
    }
  });

  // Reorder swimmers in optimal relay order if we have 4 swimmers
  if (relayTeam.length === 4) {
    // Sort by time to get rankings
    const sortedByTime = [...relayTeam].sort((a, b) => a.time - b.time);

    // Optimal relay order: 2nd best, 4th best, 3rd best, best
    const orderedTeam = [
      { ...sortedByTime[1], leg: 1 }, // 2nd best swims first
      { ...sortedByTime[3], leg: 2 }, // 4th best swims second
      { ...sortedByTime[2], leg: 3 }, // 3rd best swims third
      { ...sortedByTime[0], leg: 4 }, // Best swims anchor
    ];

    return orderedTeam;
  }

  return relayTeam;
}

/**
 * Create relay assignments sheet
 */
function createRelayAssignmentsSheet_(assignments) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let assignmentsSheet = ss.getSheetByName('Relay Assignments');
  if (!assignmentsSheet) {
    assignmentsSheet = ss.insertSheet('Relay Assignments');
  } else {
    assignmentsSheet.clear();
  }

  // Headers
  const headers = [
    'Event',
    'Level',
    'Leg 1',
    'Leg 2',
    'Leg 3',
    'Leg 4',
    'Notes',
  ];
  assignmentsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Data
  const assignmentData = assignments.map(assignment => [
    assignment.eventName,
    assignment.level,
    assignment.swimmers[0]
      ? `${assignment.swimmers[0].name} (${assignment.swimmers[0].stroke})`
      : '',
    assignment.swimmers[1]
      ? `${assignment.swimmers[1].name} (${assignment.swimmers[1].stroke})`
      : '',
    assignment.swimmers[2]
      ? `${assignment.swimmers[2].name} (${assignment.swimmers[2].stroke})`
      : '',
    assignment.swimmers[3]
      ? `${assignment.swimmers[3].name} (${assignment.swimmers[3].stroke})`
      : '',
    assignment.swimmers.length < 4
      ? 'INCOMPLETE - Need more swimmers'
      : 'Complete',
  ]);

  if (assignmentData.length > 0) {
    assignmentsSheet
      .getRange(2, 1, assignmentData.length, headers.length)
      .setValues(assignmentData);
  }

  // Format
  const headerRange = assignmentsSheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');

  console.log(`Created relay assignments for ${assignments.length} events`);
}

/**
 * Generate roster announcement sheet for easy copying to email
 * Creates a clean 4-column format: Varsity Men, Varsity Women, JV Men, JV Women
 */
function generateRosterAnnouncement() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // Check for Swimmers sheet
    const swimmersSheet = ss.getSheetByName('Swimmers');
    if (!swimmersSheet) {
      SpreadsheetApp.getUi().alert(
        'Error',
        'Swimmers sheet not found. Please run "Process Complete Tryouts" first to populate the Swimmers sheet.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    const data = swimmersSheet.getDataRange().getValues();
    if (data.length < 2) {
      SpreadsheetApp.getUi().alert(
        'Error',
        'No swimmer data found in Swimmers sheet.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    // Parse the swimmers data to extract rosters
    const rosters = parseSwimmersForRoster_(data);

    // Create the announcement sheet
    createRosterAnnouncementSheet_(rosters);

    SpreadsheetApp.getUi().alert(
      'Roster Announcement Created!',
      `Created "Roster Announcement" sheet with clean 4-column format:\n\n` +
        `• Varsity Men (${rosters.varsityMen.length} swimmers)\n` +
        `• Varsity Women (${rosters.varsityWomen.length} swimmers)\n` +
        `• JV Men (${rosters.jvMen.length} swimmers)\n` +
        `• JV Women (${rosters.jvWomen.length} swimmers)\n\n` +
        `This sheet is formatted for easy copying to email announcements!`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      'Error',
      `Failed to generate roster announcement: ${e.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    console.error('Roster announcement error:', e);
  }
}

/**
 * Parse swimmers sheet data to extract roster lists
 */
function parseSwimmersForRoster_(data) {
  const rosters = {
    varsityMen: [],
    varsityWomen: [],
    jvMen: [],
    jvWomen: [],
  };

  // Skip header row (row 0)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = row[0] ? row[0].toString().trim() : '';
    const gradYear = row[1];
    const gender = row[2] ? row[2].toString().toUpperCase().trim() : '';
    const level = row[3] ? row[3].toString().toUpperCase().trim() : '';

    // Skip rows with missing critical data
    if (!name || !gender || !level) {
      continue;
    }

    // Normalize gender values
    const isM = gender === 'M' || gender === 'MALE';
    const isF = gender === 'F' || gender === 'FEMALE';

    // Normalize level values
    const isVarsity = level === 'V' || level === 'VARSITY';
    const isJV = level === 'JV' || level === 'JUNIOR VARSITY';

    // Sort into appropriate roster
    if (isM && isVarsity) {
      rosters.varsityMen.push(name);
    } else if (isF && isVarsity) {
      rosters.varsityWomen.push(name);
    } else if (isM && isJV) {
      rosters.jvMen.push(name);
    } else if (isF && isJV) {
      rosters.jvWomen.push(name);
    }
  }

  // Sort each roster alphabetically
  rosters.varsityMen.sort();
  rosters.varsityWomen.sort();
  rosters.jvMen.sort();
  rosters.jvWomen.sort();

  return rosters;
}

/**
 * Create the clean roster announcement sheet
 */
function createRosterAnnouncementSheet_(rosters) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get or create sheet
  let announcementSheet = ss.getSheetByName('Roster Announcement');
  if (!announcementSheet) {
    announcementSheet = ss.insertSheet('Roster Announcement');
  } else {
    announcementSheet.clear();
  }

  // Set up headers
  const headers = ['Varsity Men', 'Varsity Women', 'JV Men', 'JV Women'];
  announcementSheet.getRange(1, 1, 1, 4).setValues([headers]);

  // Format headers
  const headerRange = announcementSheet.getRange(1, 1, 1, 4);
  headerRange.setBackground('#1c4587');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  headerRange.setFontSize(12);

  // Find the maximum roster length to determine how many rows we need
  const maxLength = Math.max(
    rosters.varsityMen.length,
    rosters.varsityWomen.length,
    rosters.jvMen.length,
    rosters.jvWomen.length
  );

  // Fill in the roster data
  if (maxLength > 0) {
    for (let i = 0; i < maxLength; i++) {
      const row = [
        i < rosters.varsityMen.length ? rosters.varsityMen[i] : '',
        i < rosters.varsityWomen.length ? rosters.varsityWomen[i] : '',
        i < rosters.jvMen.length ? rosters.jvMen[i] : '',
        i < rosters.jvWomen.length ? rosters.jvWomen[i] : '',
      ];
      announcementSheet.getRange(i + 2, 1, 1, 4).setValues([row]);
    }

    // Apply alternating row colors for readability
    for (let i = 2; i <= maxLength + 1; i++) {
      const rowRange = announcementSheet.getRange(i, 1, 1, 4);
      if (i % 2 === 0) {
        rowRange.setBackground('#f8f9fa'); // Light gray for even rows
      }
    }
  }

  // Set column widths
  announcementSheet.setColumnWidth(1, 150); // Varsity Men
  announcementSheet.setColumnWidth(2, 150); // Varsity Women
  announcementSheet.setColumnWidth(3, 150); // JV Men
  announcementSheet.setColumnWidth(4, 150); // JV Women

  // Add borders
  const dataRange = announcementSheet.getRange(1, 1, maxLength + 1, 4);
  dataRange.setBorder(true, true, true, true, true, true);

  // Add summary at the bottom
  const summaryRow = maxLength + 3;
  announcementSheet
    .getRange(summaryRow, 1, 1, 4)
    .setValues([
      [
        `Total: ${rosters.varsityMen.length}`,
        `Total: ${rosters.varsityWomen.length}`,
        `Total: ${rosters.jvMen.length}`,
        `Total: ${rosters.jvWomen.length}`,
      ],
    ]);

  const summaryRange = announcementSheet.getRange(summaryRow, 1, 1, 4);
  summaryRange.setFontWeight('bold');
  summaryRange.setHorizontalAlignment('center');
  summaryRange.setBackground('#e8f0fe');

  // Add instructions
  const instructionRow = summaryRow + 2;
  announcementSheet
    .getRange(instructionRow, 1)
    .setValue(
      'Instructions: Select all data above and copy (Ctrl+C) to paste into email announcements.'
    );
  announcementSheet.getRange(instructionRow, 1, 1, 4).merge();
  announcementSheet.getRange(instructionRow, 1).setFontStyle('italic');
  announcementSheet.getRange(instructionRow, 1).setWrap(true);

  console.log(
    `Created roster announcement with ${maxLength} rows and ${rosters.varsityMen.length + rosters.varsityWomen.length + rosters.jvMen.length + rosters.jvWomen.length} total swimmers`
  );
}

/**
 * Helper function to process tryout rankings from sheet data
 * @param {Array} data - The sheet data including headers
 */
function processTryoutRankings_(data) {
  const headers = data[0];
  const swimmers = data.slice(1);

  // Define the specific events to rank on
  const requiredEvents = [
    '50 Free',
    '100 Free',
    '200 Free',
    '500 Free',
    '100 Breast',
    '100 Fly',
    '100 Back',
    '200 IM',
  ];

  // Find the column indices for each required event
  const eventColumns = [];
  requiredEvents.forEach(eventName => {
    for (let i = 0; i < headers.length; i++) {
      const header = headers[i].toString().trim();
      if (header === eventName) {
        eventColumns.push({
          index: i,
          name: eventName,
        });
        break;
      }
    }
  });

  if (eventColumns.length === 0) {
    throw new Error(
      'No required event columns found. Expected: ' + requiredEvents.join(', ')
    );
  }

  console.log(
    `Found ${eventColumns.length} of ${requiredEvents.length} required events:`,
    eventColumns.map(e => e.name)
  );

  // Separate swimmers by gender and calculate rankings
  const maleSwimmers = swimmers.filter(
    row => row[1] && row[1].toString().toUpperCase() === 'M'
  );
  const femaleSwimmers = swimmers.filter(
    row => row[1] && row[1].toString().toUpperCase() === 'F'
  );

  // Process each gender group
  const maleRankings = calculateTryoutRankings_(
    maleSwimmers,
    eventColumns,
    'Male'
  );
  const femaleRankings = calculateTryoutRankings_(
    femaleSwimmers,
    eventColumns,
    'Female'
  );

  // Create the rankings sheet
  createTryoutRankingsSheet_(maleRankings, femaleRankings, eventColumns);
}

/**
 * Calculate rankings for a group of swimmers
 * @param {Array} swimmers - Array of swimmer data rows
 * @param {Array} eventColumns - Array of event column info
 * @param {string} gender - Gender label for logging
 * @returns {Array} Array of swimmer ranking objects
 */
function calculateTryoutRankings_(swimmers, eventColumns, gender) {
  const rankings = [];
  const totalSwimmers = swimmers.filter(
    swimmer => swimmer[0] && swimmer[0].toString().trim()
  ).length;

  // For each event, sort swimmers by time and assign ranks
  const eventRankings = {};

  eventColumns.forEach(event => {
    const swimmerTimes = [];

    // Collect all swimmers with their times (or lack thereof)
    swimmers.forEach((swimmer, swimmerIndex) => {
      const name = swimmer[0] ? swimmer[0].toString().trim() : '';
      if (!name) return; // Skip swimmers without names

      const timeValue = swimmer[event.index];
      let seconds = null;
      let timeString = null;

      if (timeValue && timeValue.toString().trim()) {
        seconds = parseTimeToSeconds_(timeValue.toString().trim());
        if (seconds > 0) {
          timeString = timeValue.toString().trim();
        }
      }

      swimmerTimes.push({
        swimmerIndex: swimmerIndex,
        swimmer: swimmer,
        name: name,
        seconds: seconds,
        timeString: timeString,
        hasTime: seconds !== null && seconds > 0,
      });
    });

    // Sort: swimmers with times first (by time), then swimmers without times
    swimmerTimes.sort((a, b) => {
      if (a.hasTime && b.hasTime) {
        return a.seconds - b.seconds; // Fastest first
      } else if (a.hasTime && !b.hasTime) {
        return -1; // Times come before no-times
      } else if (!a.hasTime && b.hasTime) {
        return 1; // No-times come after times
      } else {
        return 0; // Both have no time, keep original order
      }
    });

    // Assign ranks
    eventRankings[event.name] = {};

    // Count swimmers with valid times
    const swimmersWithTimes = swimmerTimes.filter(entry => entry.hasTime);
    const swimmersWithoutTimes = swimmerTimes.filter(entry => !entry.hasTime);

    // Assign ranks to swimmers with times (1, 2, 3, etc.)
    swimmersWithTimes.forEach((entry, index) => {
      eventRankings[event.name][entry.swimmerIndex] = index + 1;
    });

    // Assign same rank to all swimmers without times (next rank after those with times)
    const noTimeRank = swimmersWithTimes.length + 1;
    swimmersWithoutTimes.forEach(entry => {
      eventRankings[event.name][entry.swimmerIndex] = noTimeRank;
    });
  });

  // Calculate swimmer summaries - now every swimmer gets a rank for every event
  swimmers.forEach((swimmer, swimmerIndex) => {
    const name = swimmer[0] ? swimmer[0].toString().trim() : '';
    if (!name) return;

    const swimmerRanks = [];
    let bestRank = Infinity;
    let bestEvent = '';
    let missingEventsCount = 0; // Count actual missing times

    // Every swimmer gets a rank for every event (including missing times)
    eventColumns.forEach(event => {
      const rank = eventRankings[event.name][swimmerIndex];
      if (rank) {
        swimmerRanks.push(rank);
        if (rank < bestRank) {
          bestRank = rank;
          bestEvent = event.name;
        }
      }

      // Check if this swimmer has a time for this event
      const timeValue = swimmer[event.index];
      const hasValidTime =
        timeValue &&
        timeValue.toString().trim() &&
        parseTimeToSeconds_(timeValue.toString().trim()) > 0;
      if (!hasValidTime) {
        missingEventsCount++;
      }
    });

    // Calculate average rank across ALL required events
    const avgRank =
      swimmerRanks.length > 0
        ? Math.round(
            (swimmerRanks.reduce((sum, rank) => sum + rank, 0) /
              swimmerRanks.length) *
              100
          ) / 100
        : null;

    const rankingData = {
      name: name,
      gender: gender,
      eventRanks: {},
      eventTimes: {},
      bestRank: bestRank === Infinity ? null : bestRank,
      bestEvent: bestEvent || null,
      avgRank: avgRank,
      eventsParticipated: swimmerRanks.length,
      totalEvents: eventColumns.length,
      missingEvents: missingEventsCount, // Use the actual count of missing times
    };

    // Store individual event ranks and times
    eventColumns.forEach(event => {
      rankingData.eventRanks[event.name] =
        eventRankings[event.name][swimmerIndex] || null;
      const timeValue = swimmer[event.index];
      rankingData.eventTimes[event.name] =
        (timeValue && timeValue.toString().trim()) || null;
    });

    rankings.push(rankingData);
  });

  // Sort by average rank (best average first)
  rankings.sort((a, b) => {
    if (a.avgRank === null && b.avgRank === null) return 0;
    if (a.avgRank === null) return 1;
    if (b.avgRank === null) return -1;
    return a.avgRank - b.avgRank;
  });

  return rankings;
}

/**
 * Create the Tryout Rankings sheet with male and female rankings
 * @param {Array} maleRankings - Male swimmer rankings
 * @param {Array} femaleRankings - Female swimmer rankings
 * @param {Array} eventColumns - Event column information
 */
function createTryoutRankingsSheet_(
  maleRankings,
  femaleRankings,
  eventColumns
) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create or get the rankings sheet
  let sheet;
  try {
    sheet = ss.getSheetByName('Tryout Rankings');
    sheet.clear();
  } catch (e) {
    sheet = ss.insertSheet('Tryout Rankings');
  }

  let currentRow = 1;

  // Helper function to write rankings for a gender
  function writeGenderRankings(rankings, genderLabel) {
    if (rankings.length === 0) return;

    // Gender header
    sheet.getRange(currentRow, 1).setValue(`${genderLabel} Tryout Rankings`);
    sheet
      .getRange(currentRow, 1, 1, 5 + eventColumns.length)
      .setBackground('#4a90e2')
      .setFontColor('white')
      .setFontWeight('bold');
    currentRow++;

    // Column headers
    const headers = ['Rank', 'Name', 'Avg Rank', 'Best Rank', 'Best Event'];
    eventColumns.forEach(event => headers.push(event.name + ' (Rank)'));

    sheet.getRange(currentRow, 1, 1, headers.length).setValues([headers]);
    sheet
      .getRange(currentRow, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#e6f3ff');
    currentRow++;

    // Data rows
    rankings.forEach((swimmer, index) => {
      const row = [
        index + 1, // Overall rank
        swimmer.name,
        swimmer.avgRank || 'N/A',
        swimmer.bestRank || 'N/A',
        swimmer.bestEvent || 'N/A',
      ];

      // Add event ranks
      eventColumns.forEach(event => {
        const rank = swimmer.eventRanks[event.name];
        const time = swimmer.eventTimes[event.name];
        if (rank && time) {
          row.push(`${rank} (${time})`);
        } else if (rank) {
          row.push(`${rank} (No Time)`);
        } else {
          row.push('');
        }
      });

      sheet.getRange(currentRow, 1, 1, row.length).setValues([row]);
      currentRow++;
    });

    currentRow++; // Space between sections
  }

  // Write both gender sections
  writeGenderRankings(maleRankings, 'Male');
  writeGenderRankings(femaleRankings, 'Female');

  // Apply conditional formatting to event time columns
  applyTryoutRankingsConditionalFormatting_(
    sheet,
    maleRankings,
    femaleRankings,
    eventColumns
  );

  // Auto-resize columns
  for (let i = 1; i <= 5 + eventColumns.length; i++) {
    sheet.autoResizeColumn(i);
  }

  // Freeze header rows and name column
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);

  console.log(
    `Created Tryout Rankings sheet with ${maleRankings.length} male and ${femaleRankings.length} female swimmers`
  );
}

/**
 * Apply conditional formatting to the Tryout Rankings sheet
 * @param {Sheet} sheet - The Tryout Rankings sheet
 * @param {Array} maleRankings - Male swimmer rankings
 * @param {Array} femaleRankings - Female swimmer rankings
 * @param {Array} eventColumns - Event column information
 */
function applyTryoutRankingsConditionalFormatting_(
  sheet,
  maleRankings,
  femaleRankings,
  eventColumns
) {
  // Helper function to apply formatting for a gender section
  function formatGenderSection(rankings, startRow, genderLabel) {
    if (rankings.length === 0) return startRow;

    const headerRow = startRow;
    const dataStartRow = startRow + 2; // Account for gender header and column headers
    const dataEndRow = dataStartRow + rankings.length - 1;

    // For each event column, apply conditional formatting
    eventColumns.forEach((event, eventIndex) => {
      const colIndex = 6 + eventIndex; // Event columns start at column 6 (after Rank, Name, Avg Rank, Best Rank, Best Event)

      if (dataEndRow >= dataStartRow) {
        const range = sheet.getRange(
          dataStartRow,
          colIndex,
          dataEndRow - dataStartRow + 1,
          1
        );

        // Collect all valid times for this event in this gender to determine quartiles
        const validTimes = [];
        rankings.forEach(swimmer => {
          const timeStr = swimmer.eventTimes[event.name];
          if (
            timeStr &&
            timeStr.trim() &&
            parseTimeToSeconds_(timeStr.trim()) > 0
          ) {
            validTimes.push(parseTimeToSeconds_(timeStr.trim()));
          }
        });

        if (validTimes.length > 0) {
          // Sort times (fastest to slowest)
          validTimes.sort((a, b) => a - b);

          // Calculate quartiles
          const q1Index = Math.floor(validTimes.length * 0.33); // Top 33% = Green
          const q3Index = Math.floor(validTimes.length * 0.67); // Bottom 33% = Red, Middle = Yellow

          const q1Time = validTimes[q1Index];
          const q3Time = validTimes[q3Index];

          // Apply formatting to each cell in this column
          const values = range.getValues();
          for (let i = 0; i < values.length; i++) {
            const cellValue = values[i][0];
            if (cellValue && cellValue.toString().includes('(')) {
              // Extract time from "rank (time)" format
              const timeMatch = cellValue.toString().match(/\(([^)]+)\)/);
              if (timeMatch && timeMatch[1] !== 'No Time') {
                const timeStr = timeMatch[1];
                const seconds = parseTimeToSeconds_(timeStr);

                if (seconds > 0) {
                  const cellRange = sheet.getRange(dataStartRow + i, colIndex);

                  if (seconds <= q1Time) {
                    // Best times - Green
                    cellRange.setBackground('#90EE90'); // Light Green
                  } else if (seconds <= q3Time) {
                    // Medium times - Yellow
                    cellRange.setBackground('#FFFF99'); // Light Yellow
                  } else {
                    // Slowest times - Red
                    cellRange.setBackground('#FFB6C1'); // Light Red
                  }
                }
              } else if (timeMatch && timeMatch[1] === 'No Time') {
                // No time - Gray
                const cellRange = sheet.getRange(dataStartRow + i, colIndex);
                cellRange.setBackground('#D3D3D3'); // Light Gray
              }
            }
          }
        }
      }
    });

    return dataEndRow + 2; // Account for the space after this section
  }

  // Apply formatting to male section (starts at row 1)
  let currentRow = 1;
  if (maleRankings.length > 0) {
    currentRow = formatGenderSection(maleRankings, currentRow, 'Male');
  }

  // Apply formatting to female section
  if (femaleRankings.length > 0) {
    formatGenderSection(femaleRankings, currentRow, 'Female');
  }
}

// Function to be called from host sheet's onOpen()
function setupCoachToolsMenu() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Coach Tools')
    .addItem('About Coach Tools', 'aboutCoachTools')
    .addItem('Refresh All (safe)', 'refreshAll')
    .addSubMenu(
      ui
        .createMenu('Attendance')
        .addItem('📋 Open Attendance Tracker', 'openAttendanceSidebar')
        .addItem('📊 Create Weekly Summary', 'createAttendanceSummary')
        .addSeparator()
        .addItem('🧪 Add Test Data', 'createTestAttendanceData')
    )
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu('Results')
        .addItem('Add Result (sidebar)', 'openAddResultSidebar')
        .addItem('Refresh PR Summary & Dashboard', 'refreshPRs')
    )
    .addSubMenu(
      ui
        .createMenu('Roster')
        .addItem('Add Swimmer + PRs (sidebar)', 'openAddSwimmerSidebar')
        .addSeparator()
        .addItem(
          '📄 Create Raw Tryout Results Sheet',
          'createRawTryoutResultsSheet'
        )
        .addItem(
          '🏊 Generate Tryout Rankings from Sheet',
          'generateTryoutRankingsFromSheet'
        )
        .addItem('🏆 Generate Varsity/JV Squads', 'generateVarsityJVSquads')
        .addItem(
          '📧 Generate Roster Announcement',
          'generateRosterAnnouncement'
        )
        .addSeparator()
        .addItem(
          '⚡ Process Complete Tryouts (Rankings + Squads + Swimmers)',
          'processCompleteTryouts'
        )
        .addSeparator()
        .addItem('🎨 Apply Swimmers Color Coding', 'applySwimmersColorCoding')
        .addSeparator()
        .addItem(
          '📊 Create Personal Records from Tryouts',
          'createPRsFromTryouts'
        )
        .addItem('🏊‍♂️ Setup Relay Events Config', 'setupRelayEvents')
        .addItem(
          '🏊‍♀️ Generate Smart Relay Assignments',
          'generateRelayAssignments'
        )
        .addItem(
          '📊 Refresh Swimmer Assignment Summary',
          'refreshSwimmerAssignmentSummary'
        )
        .addSeparator()
        .addItem(
          '📋 Create My Relay Entry Sheet',
          'createMyRelayEntrySheet'
        )
        .addItem(
          '📋 Create Blank Relay Entry Sheets',
          'createBlankRelayEntrySheets'
        )
        .addSeparator()
        .addItem(
          '⚙️ Setup Team Relay Meet Config',
          'setupTeamRelayMeetConfig'
        )
        .addItem(
          '🔧 Validate Relay Headers',
          'validateRelayAssignmentsHeaders'
        )
        .addSeparator()
        .addItem(
          'Generate Roster Rankings from CSV',
          'generateRosterRankingsFromCSV'
        )
    )
    .addSubMenu(
      ui
        .createMenu('Heat Sheets')
        .addItem('🏊 Setup Lane Assignments', 'setupLaneAssignments')
        .addItem('📄 Generate Relay Heat Sheet', 'generateRelayHeatSheet')
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
        .createMenu('Under Development')
        .addItem(
          'Generate Sample Team (50: 25F/25M, 10V/15JV each)',
          'generateSampleTeam50'
        )
        .addSeparator()
        .addItem('Bulk Import (CSV paste)', 'openBulkImportSidebar')
        .addSeparator()
        .addItem('Ensure Meet Presets Table', 'ensureMeetPresetsTemplate')
        .addItem('Apply Meet Presets to Lineup', 'applyMeetPresets')
        .addItem('Check Lineup (Usage & Violations)', 'checkLineup')
        .addItem('Create Snapshot of Current Lineup', 'createSnapshot')
        .addItem('Build Coach Packet (print view)', 'buildCoachPacket')
    )
    .addToUi();
}

// Sidebar functions for host sheet menu items
function openBulkImportSidebar() {
  const html = buildBulkImportSidebar();
  SpreadsheetApp.getUi().showSidebar(html);
}
function openAddResultSidebar() {
  const html = buildAddResultSidebar();
  SpreadsheetApp.getUi().showSidebar(html);
}

function openAttendanceSidebar() {
  const html = buildAttendanceSidebar();
  SpreadsheetApp.getUi().showSidebar(html);
}

function openAddMeetSidebar() {
  const html = buildAddMeetSidebar();
  SpreadsheetApp.getUi().showSidebar(html);
}
function openAddEventSidebar() {
  const html = buildAddEventSidebar();
  SpreadsheetApp.getUi().showSidebar(html);
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

function createPRSummary() {
  const ss = SpreadsheetApp.getActive();
  const results = mustSheet(SHEET_NAMES.RESULTS);
  const out =
    ss.getSheetByName(SHEET_NAMES.PR_SUMMARY) ||
    ss.insertSheet(SHEET_NAMES.PR_SUMMARY);
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
  const prs = ss.getSheetByName(SHEET_NAMES.PR_SUMMARY) || createPRSummary();
  const swSheet = mustSheet(SHEET_NAMES.SWIMMERS);

  // Ensure a named range for swimmers exists (dynamic full column)
  ss.setNamedRange('SwimmerNames', swSheet.getRange('A2:A'));

  const name = SHEET_NAMES.SWIMMER_DASHBOARD;
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
  QUERY('${SHEET_NAMES.PR_SUMMARY}'!A2:H,
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

function createAttendanceSummary() {
  const ss = SpreadsheetApp.getActive();
  const name = SHEET_NAMES.ATTENDANCE_SUMMARY;
  let summary = ss.getSheetByName(name) || ss.insertSheet(name);
  summary.clear();

  // Check if Master Attendance sheet exists and has data
  const masterSheet = ss.getSheetByName(SHEET_NAMES.MASTER_ATTENDANCE);
  if (!masterSheet) {
    summary
      .getRange('A1')
      .setValue(
        'Error: Master Attendance sheet not found. Please create attendance data first.'
      );
    return;
  }

  const dataRange = masterSheet.getDataRange();
  if (!dataRange || dataRange.getNumRows() < 2) {
    summary
      .getRange('A1')
      .setValue(
        'No attendance data found. Please add some attendance records first.'
      );
    return;
  }

  // Title
  summary
    .getRange('A1')
    .setValue('Weekly Attendance Summary')
    .setFontWeight('bold')
    .setFontSize(14);

  // Debug info
  summary
    .getRange('A2')
    .setValue(
      `Data source: ${SHEET_NAMES.MASTER_ATTENDANCE} (${dataRange.getNumRows() - 1} records)`
    );

  // Weekly data aggregation using QUERY
  const headers = ['Swimmer', 'Team', 'Week', 'Practices', 'Meets Requirement'];
  summary.getRange('A4:E4').setValues([headers]).setFontWeight('bold');

  // Query to aggregate attendance by swimmer and week
  summary.getRange('A5').setFormula(
    `
=QUERY('${SHEET_NAMES.MASTER_ATTENDANCE}'!A:H,
  "select B, D, WEEKNUM(A), count(B) 
   where C = TRUE or E = TRUE
   group by B, D, WEEKNUM(A)
   order by D, B, WEEKNUM(A)",
  0
)
  `.trim()
  );

  // Alternative query to show all data for debugging
  summary
    .getRange('G4')
    .setValue('Debug: All Attendance Data')
    .setFontWeight('bold');
  summary.getRange('G5').setFormula(
    `
=QUERY('${SHEET_NAMES.MASTER_ATTENDANCE}'!A:H,
  "select B, D, A, C where C = TRUE limit 10",
  0
)
  `.trim()
  );

  // Meets requirement column (✅ if >= 3, ❌ if < 3)
  summary.getRange('E5').setFormula('=IF(D5>=3,"✅","❌")');

  // Apply the formula to more rows
  summary.getRange('E5:E50').setFormula('=IF(D5:D50>=3,"✅","❌")');

  // Team-level compliance metrics
  summary
    .getRange('A52')
    .setValue('Team Compliance Metrics')
    .setFontWeight('bold')
    .setFontSize(12);

  // Create simplified metrics for now
  summary.getRange('A54').setValue('Varsity Team').setFontWeight('bold');
  summary.getRange('A55').setValue('% Meeting 3x/week:');
  summary.getRange('B55').setFormula(
    `
=ROUND(COUNTIFS(B:B,"Varsity",E:E,"✅") / COUNTIF(B:B,"Varsity") * 100, 1) & "%"
  `.trim()
  );

  summary.getRange('A56').setValue('Avg practices/week:');
  summary.getRange('B56').setFormula(
    `
=ROUND(AVERAGEIF(B:B,"Varsity",D:D), 1)
  `.trim()
  );

  summary.getRange('A58').setValue('JV Team').setFontWeight('bold');
  summary.getRange('A59').setValue('% Meeting 3x/week:');
  summary.getRange('B59').setFormula(
    `
=ROUND(COUNTIFS(B:B,"JV",E:E,"✅") / COUNTIF(B:B,"JV") * 100, 1) & "%"
  `.trim()
  );

  summary.getRange('A60').setValue('Avg practices/week:');
  summary.getRange('B60').setFormula(
    `
=ROUND(AVERAGEIF(B:B,"JV",D:D), 1)
  `.trim()
  );

  // Format and resize
  summary.setFrozenRows(4);
  summary.autoResizeColumns(1, 8);
}

function createTestAttendanceData() {
  const ss = SpreadsheetApp.getActive();
  let masterSheet = ss.getSheetByName(SHEET_NAMES.MASTER_ATTENDANCE);

  if (!masterSheet) {
    masterSheet = ss.insertSheet(SHEET_NAMES.MASTER_ATTENDANCE);
    // Add headers
    const headers = [
      'Date',
      'Name',
      'Present',
      'Excused',
      'Level',
      'Gender',
      'Timestamp',
      'UpdatedBy',
      'Source',
    ];
    masterSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  // Add some test data
  const testData = [
    [
      '2025-08-12',
      'Alice Johnson',
      true,
      false,
      'Varsity',
      'F',
      new Date().toISOString(),
      'test',
      'test',
    ],
    [
      '2025-08-12',
      'Bob Smith',
      true,
      false,
      'Varsity',
      'M',
      new Date().toISOString(),
      'test',
      'test',
    ],
    [
      '2025-08-12',
      'Carol Davis',
      false,
      true,
      'JV',
      'F',
      new Date().toISOString(),
      'test',
      'test',
    ], // Excused
    [
      '2025-08-14',
      'Alice Johnson',
      true,
      false,
      'Varsity',
      'F',
      new Date().toISOString(),
      'test',
      'test',
    ],
    [
      '2025-08-14',
      'Bob Smith',
      true,
      false,
      'Varsity',
      'M',
      new Date().toISOString(),
      'test',
      'test',
    ],
    [
      '2025-08-14',
      'Carol Davis',
      true,
      false,
      'JV',
      'F',
      new Date().toISOString(),
      'test',
      'test',
    ],
    [
      '2025-08-16',
      'Alice Johnson',
      true,
      false,
      'Varsity',
      'F',
      new Date().toISOString(),
      'test',
      'test',
    ],
    [
      '2025-08-16',
      'Bob Smith',
      false,
      true,
      'Varsity',
      'M',
      new Date().toISOString(),
      'test',
      'test',
    ], // Excused
    [
      '2025-08-16',
      'Carol Davis',
      true,
      false,
      'JV',
      'F',
      new Date().toISOString(),
      'test',
      'test',
    ],
  ];

  const startRow = masterSheet.getLastRow() + 1;
  masterSheet
    .getRange(startRow, 1, testData.length, testData[0].length)
    .setValues(testData);

  SpreadsheetApp.getUi().alert(
    'Test attendance data added! Now try creating the attendance summary again.'
  );
}

function createAttendanceCharts(summary) {
  // For now, let's skip the complex charts and focus on getting the basic data working
  // We can add charts back once we confirm the QUERY is working

  summary
    .getRange('J1')
    .setValue('Charts will be added once data is confirmed working')
    .setFontWeight('bold');
  summary
    .getRange('J2')
    .setValue(
      'Debug: If you see swimmer data in columns A-E, charts can be enabled'
    );
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

// Attendance sidebar functions
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

/**
 * Helper functions for relay row creation
 */

/**
 * Normalize level values to ensure they match data validation rules
 */
function normalizeLevel_(level) {
  if (!level) return 'Varsity'; // Default fallback
  
  const levelStr = level.toString().toUpperCase().trim();
  
  if (levelStr === 'V' || levelStr === 'VARSITY') {
    return 'Varsity';
  } else if (levelStr === 'JV' || levelStr === 'JUNIOR VARSITY') {
    return 'JV';
  }
  
  // If it doesn't match known patterns, try to detect from content
  if (levelStr.includes('JV') || levelStr.includes('JUNIOR')) {
    return 'JV';
  }
  
  // Default to Varsity if unclear
  return 'Varsity';
}

function createConventionalRelayRow_(eventName, level, gender, selectedSwimmers, strokes, eligibleCount) {
  const relayRow = new Array(22).fill(''); // Initialize array with 22 elements
  relayRow[0] = eventName;
  relayRow[1] = normalizeLevel_(level); // Normalize level to ensure validation compliance
  relayRow[2] = gender;
  relayRow[3] = false; // Lock checkbox
  
  // Fill first 4 legs for conventional relays
  for (let i = 0; i < 4 && i < selectedSwimmers.length; i++) {
    relayRow[4 + i * 2] = selectedSwimmers[i].name; // Swimmer name
    relayRow[5 + i * 2] = getTentativeTime_(selectedSwimmers[i], strokes[i]); // Time
  }
  
  relayRow[20] = 'TBD'; // Total Time
  relayRow[21] = `Smart-assigned (${eligibleCount} eligible)`; // Notes
  
  return relayRow;
}

function createPartialConventionalRelayRow_(eventName, level, gender, selectedSwimmers, strokes) {
  const relayRow = new Array(22).fill(''); // Initialize array with 22 elements
  relayRow[0] = eventName;
  relayRow[1] = normalizeLevel_(level); // Normalize level to ensure validation compliance
  relayRow[1] = level;
  relayRow[2] = gender;
  relayRow[3] = false; // Lock checkbox
  
  // Fill available swimmers
  for (let i = 0; i < Math.min(4, selectedSwimmers.length); i++) {
    relayRow[4 + i * 2] = selectedSwimmers[i].name;
    relayRow[5 + i * 2] = getTentativeTime_(selectedSwimmers[i], strokes[i]);
  }
  
  relayRow[20] = 'TBD';
  relayRow[21] = `Partial: ${selectedSwimmers.length}/4 (many swimmers at 4-relay limit)`;
  
  return relayRow;
}

function createEmptyConventionalRelayRow_(eventName, level, gender) {
  const relayRow = new Array(22).fill('');
  relayRow[0] = eventName;
  relayRow[1] = normalizeLevel_(level); // Normalize level to ensure validation compliance
  relayRow[2] = gender;
  relayRow[3] = false; // Lock checkbox
  relayRow[20] = 'TBD';
  relayRow[21] = 'No available swimmers (all at 4-relay limit)';
  
  return relayRow;
}

function createNonConventionalRelayRow_(eventName, level, gender, selectedSwimmers, numLegs, strokePattern, eligibleCount) {
  const relayRow = new Array(22).fill('');
  relayRow[0] = eventName;
  relayRow[1] = normalizeLevel_(level); // Normalize level to ensure validation compliance
  relayRow[2] = gender;
  relayRow[3] = false; // Lock checkbox
  
  // Fill legs based on numLegs
  for (let i = 0; i < numLegs && i < selectedSwimmers.length; i++) {
    relayRow[4 + i * 2] = selectedSwimmers[i].name;
    relayRow[5 + i * 2] = getTentativeTime_(selectedSwimmers[i], 'Free'); // Default to Free for now
  }
  
  relayRow[20] = 'TBD';
  relayRow[21] = `${numLegs}-leg relay: Smart-assigned (${eligibleCount} eligible)`;
  
  return relayRow;
}

function createPartialNonConventionalRelayRow_(eventName, level, gender, selectedSwimmers, numLegs) {
  const relayRow = new Array(22).fill('');
  relayRow[0] = eventName;
  relayRow[1] = normalizeLevel_(level); // Normalize level to ensure validation compliance
  relayRow[2] = gender;
  relayRow[3] = false; // Lock checkbox
  
  // Fill available swimmers
  for (let i = 0; i < selectedSwimmers.length; i++) {
    relayRow[4 + i * 2] = selectedSwimmers[i].name;
    relayRow[5 + i * 2] = getTentativeTime_(selectedSwimmers[i], 'Free');
  }
  
  relayRow[20] = 'TBD';
  relayRow[21] = `Partial ${numLegs}-leg: ${selectedSwimmers.length}/${numLegs} (swimmers at 4-relay limit)`;
  
  return relayRow;
}

function createEmptyNonConventionalRelayRow_(eventName, level, gender, numLegs) {
  const relayRow = new Array(22).fill('');
  relayRow[0] = eventName;
  relayRow[1] = normalizeLevel_(level); // Normalize level to ensure validation compliance
  relayRow[2] = gender;
  relayRow[3] = false; // Lock checkbox
  relayRow[20] = 'TBD';
  relayRow[21] = `${numLegs}-leg relay: No available swimmers (all at 4-relay limit)`;
  
  return relayRow;
}

function selectSwimmersForNonConventionalRelay_(eligibleSwimmers, swimmerAssignments, eventName, gender, numLegs) {
  // First try swimmers who are under the 4-relay limit
  const preferredSwimmers = eligibleSwimmers.filter(swimmer => {
    const currentAssignments = swimmerAssignments.get(swimmer.name) || [];
    return currentAssignments.length < 4;
  });

  // Sort all eligible swimmers by current assignment count then by best time
  const sortedSwimmers = [...eligibleSwimmers].sort((a, b) => {
    const aAssignments = (swimmerAssignments.get(a.name) || []).length;
    const bAssignments = (swimmerAssignments.get(b.name) || []).length;
    
    if (aAssignments !== bAssignments) {
      return aAssignments - bAssignments; // Fewer assignments first
    }
    
    // If same assignment count, sort by best freestyle time
    const aTime = parseTimeToSeconds_(a.personalRecords?.['50 Free'] || '99:99.99');
    const bTime = parseTimeToSeconds_(b.personalRecords?.['50 Free'] || '99:99.99');
    return aTime - bTime;
  });

  // Use preferred swimmers first, then fill remaining spots with any available swimmers
  const selectedSwimmers = [];
  
  // Add swimmers under 4-relay limit first
  selectedSwimmers.push(...preferredSwimmers.slice(0, numLegs));
  
  // If we need more swimmers and don't have enough under the limit, add more
  if (selectedSwimmers.length < numLegs) {
    const remainingNeeded = numLegs - selectedSwimmers.length;
    const additionalSwimmers = sortedSwimmers
      .filter(swimmer => !selectedSwimmers.includes(swimmer))
      .slice(0, remainingNeeded);
    selectedSwimmers.push(...additionalSwimmers);
  }

  console.log(
    `  ${eventName} ${gender}: ${selectedSwimmers.length} swimmers selected for ${numLegs}-leg relay (${preferredSwimmers.length} under 4-relay limit, ${selectedSwimmers.length - preferredSwimmers.slice(0, numLegs).length} over limit)`
  );

  return selectedSwimmers;
}

/**
 * Create relay entry sheet for your team with swimmers and their event assignments
 */
function createMyRelayEntrySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // First, validate and fix relay assignments headers if needed
    validateRelayAssignmentsHeaders();
    
    // Check if Swimmers sheet exists
    const swimmersSheet = ss.getSheetByName('Swimmers');
    if (!swimmersSheet) {
      SpreadsheetApp.getUi().alert(
        'Missing Swimmers Sheet',
        'Please generate the Swimmers sheet first using "Process Complete Tryouts".',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    // Check if Relay Assignments sheet exists
    const relaySheet = ss.getSheetByName('Relay Assignments');
    if (!relaySheet) {
      SpreadsheetApp.getUi().alert(
        'Missing Relay Assignments',
        'Please generate relay assignments first using "Generate Smart Relay Assignments".',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    // Get swimmers data
    const swimmersData = swimmersSheet.getDataRange().getValues();
    const swimmersHeaders = swimmersData[0];
    const nameIndex = swimmersHeaders.indexOf('Name');
    const levelIndex = swimmersHeaders.indexOf('Level');
    const genderIndex = swimmersHeaders.indexOf('Gender');
    const gradeIndex = swimmersHeaders.indexOf('Grad Year');

    if (nameIndex === -1 || levelIndex === -1 || genderIndex === -1) {
      SpreadsheetApp.getUi().alert(
        'Invalid Swimmers Sheet',
        'The Swimmers sheet is missing required columns (Name, Level, Gender).',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    // Get relay assignments data
    const relayData = relaySheet.getDataRange().getValues();
    let relayHeaders;
    let relayDataStart = 1;
    
    console.log('DEBUG: Relay sheet has', relayData.length, 'rows');
    console.log('DEBUG: First row:', relayData[0]);
    
    // Check if sheet is empty
    if (relayData.length === 0 || (relayData.length === 1 && relayData[0].every(cell => !cell))) {
      SpreadsheetApp.getUi().alert(
        'Empty Relay Assignments',
        'The Relay Assignments sheet appears to be empty. Please run "Generate Smart Relay Assignments" first.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Skip instruction row if present
    if (relayData[0][0] && relayData[0][0].toString().includes('💡')) {
      relayHeaders = relayData[1];
      relayDataStart = 2;
    } else {
      relayHeaders = relayData[0];
    }

    // Build swimmer assignments map
    const swimmerEventMap = new Map();
    
    // Find leg columns in relay data
    const legColumns = [];
    for (let i = 0; i < relayHeaders.length; i++) {
      const header = (relayHeaders[i] || '').toString().trim();
      if (header.startsWith('Leg ') && !header.includes('Time')) {
        legColumns.push(i);
      }
    }
    
    const eventIndex = relayHeaders.findIndex(h => (h || '').toString().trim() === 'Event');
    
    // If "Event" column not found, assume first column contains events
    const actualEventIndex = eventIndex !== -1 ? eventIndex : 0;
    
    console.log('DEBUG: Found leg columns:', legColumns);
    console.log('DEBUG: Event column index:', eventIndex);
    console.log('DEBUG: Using event index:', actualEventIndex);
    
    // Check if we found the required columns
    if (legColumns.length === 0) {
      SpreadsheetApp.getUi().alert(
        'Invalid Relay Assignments Sheet', 
        'Could not find any "Leg" columns in Relay Assignments sheet. Please regenerate relay assignments.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Process relay assignments
    for (let i = relayDataStart; i < relayData.length; i++) {
      const row = relayData[i];
      const event = row[eventIndex];
      if (!event) continue;
      
      console.log('DEBUG: Processing event:', event);
      
      // Check each leg for swimmers
      legColumns.forEach(legCol => {
        const swimmer = row[legCol];
        if (swimmer && swimmer.toString().trim()) {
          const swimmerName = swimmer.toString().trim();
          if (!swimmerEventMap.has(swimmerName)) {
            swimmerEventMap.set(swimmerName, new Set());
          }
          swimmerEventMap.get(swimmerName).add(event.toString().trim());
        }
      });
    }

    // Create or update My Relay Entry sheet
    let entrySheet = ss.getSheetByName('My Relay Entry');
    if (entrySheet) {
      ss.deleteSheet(entrySheet);
    }
    entrySheet = ss.insertSheet('My Relay Entry');

    // Set up the relay entry sheet structure based on the CSV
    const relayEvents = [
      '200 Medley', '200 Medley JV', '350 Free', '200 Fly', '200 Fly JV',
      '200 Breast', '200 Breast JV', '200 Free', '200 Free JV', '400 IM',
      '400 IM JV', 'I-tube', '200 Back', '200 Back JV', '400 Free',
      '400 Free JV', '200 Medley Co-ed', '200 Free Frosh'
    ];

    // Team name header - don't merge to avoid freeze conflicts
    entrySheet.getRange(1, 1).setValue('Team Name: Mavs');
    entrySheet.getRange(1, 1).setBackground('#e3f2fd').setFontWeight('bold');

    // Column headers
    const headers = ['Swimmer Name', 'Grade'].concat(relayEvents);
    entrySheet.getRange(2, 1, 1, headers.length).setValues([headers]);
    entrySheet.getRange(2, 1, 1, headers.length)
      .setBackground('#0d47a1')
      .setFontColor('#ffffff')
      .setFontWeight('bold');

    // Separate girls and boys
    const girls = [];
    const boys = [];
    
    for (let i = 1; i < swimmersData.length; i++) {
      const row = swimmersData[i];
      const name = row[nameIndex];
      const gender = row[genderIndex];
      const grade = row[gradeIndex];
      
      if (!name) continue;
      
      const swimmerRow = [name, grade || ''];
      
      // Add X's for events this swimmer is assigned to
      relayEvents.forEach(event => {
        const hasEvent = swimmerEventMap.has(name) && swimmerEventMap.get(name).has(event);
        if (hasEvent) {
          console.log('DEBUG: Found match for', name, 'in event', event);
        }
        swimmerRow.push(hasEvent ? 'X' : '');
      });
      
      if (gender === 'F') {
        girls.push(swimmerRow);
      } else {
        boys.push(swimmerRow);
      }
    }

    // Sort by last name
    const sortByLastName = (a, b) => {
      const aLast = a[0].split(' ').pop() || '';
      const bLast = b[0].split(' ').pop() || '';
      return aLast.localeCompare(bLast);
    };
    
    girls.sort(sortByLastName);
    boys.sort(sortByLastName);

    let currentRow = 3;

    // Add Girls section
    entrySheet.getRange(currentRow, 1).setValue('Girls');
    entrySheet.getRange(currentRow, 1).setFontWeight('bold');
    currentRow++;

    if (girls.length > 0) {
      entrySheet.getRange(currentRow, 1, girls.length, headers.length).setValues(girls);
      currentRow += girls.length;
    }

    // Add some blank rows
    currentRow += 5;

    // Add column headers again for boys section
    entrySheet.getRange(currentRow, 1, 1, headers.length).setValues([headers]);
    entrySheet.getRange(currentRow, 1, 1, headers.length)
      .setBackground('#0d47a1')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    currentRow++;

    // Add Boys section
    entrySheet.getRange(currentRow, 1).setValue('Boys');
    entrySheet.getRange(currentRow, 1).setFontWeight('bold');
    currentRow++;

    if (boys.length > 0) {
      entrySheet.getRange(currentRow, 1, boys.length, headers.length).setValues(boys);
    }

    // Format the sheet
    entrySheet.setColumnWidth(1, 150); // Swimmer Name
    entrySheet.setColumnWidth(2, 60);  // Grade
    for (let i = 3; i <= headers.length; i++) {
      entrySheet.setColumnWidth(i, 80); // Event columns
    }

    // Freeze the header rows
    entrySheet.setFrozenRows(2);
    entrySheet.setFrozenColumns(2);

    SpreadsheetApp.getUi().alert(
      'Success!',
      `My Relay Entry sheet created with ${girls.length} girls and ${boys.length} boys.\n\nSwimmers are marked with X for their assigned relay events.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (e) {
    SpreadsheetApp.getUi().alert(
      'Error',
      `Failed to create relay entry sheet: ${e.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    console.error('Create relay entry error:', e);
  }
}

/**
 * Sets up or updates the Team Relay Meet Config sheet with default team configurations
 */
function setupTeamRelayMeetConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName('Team Relay Meet Config');
  
  if (!configSheet) {
    configSheet = ss.insertSheet('Team Relay Meet Config');
    
    // Set up headers
    const headers = [
      'Team Name',
      'Girls Count',
      'Boys Count'
    ];
    
    configSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Add default teams
    const defaultTeams = [
      ['Team 1', 30, 30],
      ['Team 2', 30, 30],
      ['Team 3', 30, 30],
      ['Team 4', 30, 30],
      ['Team 5', 30, 30]
    ];
    
    configSheet.getRange(2, 1, defaultTeams.length, 3).setValues(defaultTeams);
    
    // Format headers
    const headerRange = configSheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#0d47a1');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    
    // Auto-resize columns
    configSheet.autoResizeColumns(1, headers.length);
    
    // Add instruction
    configSheet.getRange(defaultTeams.length + 3, 1, 1, 3).merge();
    configSheet.getRange(defaultTeams.length + 3, 1).setValue(
      '💡 Modify team names and swimmer counts above. Use "Create Blank Relay Entry Sheets" to generate sheets based on this configuration.'
    );
    configSheet.getRange(defaultTeams.length + 3, 1)
      .setBackground('#e3f2fd')
      .setFontStyle('italic')
      .setWrap(true);
      
    console.log('Created Team Relay Meet Config sheet with default teams');
  }
  
  return configSheet;
}

/**
 * Gets team configuration from the Team Relay Meet Config sheet
 */
function getTeamConfigurations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Team Relay Meet Config');
  
  if (!configSheet) {
    // Create the config sheet if it doesn't exist
    setupTeamRelayMeetConfig();
    return getTeamConfigurations(); // Recursive call to get the data
  }
  
  const data = configSheet.getDataRange().getValues();
  const headers = data[0];
  const teams = [];
  
  // Parse team data (skip header row)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const teamName = row[0];
    const girlsCount = row[1] || 30;
    const boysCount = row[2] || 30;
    
    // Skip empty rows
    if (!teamName || teamName.toString().trim() === '') {
      continue;
    }
    
    // Skip instruction rows
    if (teamName.toString().includes('💡')) {
      continue;
    }
    
    teams.push({
      name: teamName.toString().trim(),
      girlsCount: parseInt(girlsCount) || 30,
      boysCount: parseInt(boysCount) || 30
    });
  }
  
  return teams;
}

/**
 * Create blank relay entry sheets for teams based on Team Relay Meet Config
 */
function createBlankRelayEntrySheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Ensure Team Relay Meet Config exists and get team configurations
    setupTeamRelayMeetConfig();
    const teamConfigs = getTeamConfigurations();
    
    if (teamConfigs.length === 0) {
      SpreadsheetApp.getUi().alert(
        'No Teams Configured',
        'Please add team configurations to the "Team Relay Meet Config" sheet first.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Set up the relay entry sheet structure
    const relayEvents = [
      '200 Medley', '200 Medley JV', '350 Free', '200 Fly', '200 Fly JV',
      '200 Breast', '200 Breast JV', '200 Free', '200 Free JV', '400 IM',
      '400 IM JV', 'I-tube', '200 Back', '200 Back JV', '400 Free',
      '400 Free JV', '200 Medley Co-ed', '200 Free Frosh'
    ];

    const headers = ['Swimmer Name', 'Grade'].concat(relayEvents);
    
    const createdSheets = [];
    const skippedSheets = [];
    
    teamConfigs.forEach(teamConfig => {
      const teamName = teamConfig.name;
      const girlsCount = teamConfig.girlsCount;
      const boysCount = teamConfig.boysCount;
      
      // Check if sheet already exists
      let sheet = ss.getSheetByName(teamName);
      if (sheet) {
        // Update the existing sheet's team name in the header
        sheet.getRange(1, 1).setValue(`Team Name: ${teamName}`);
        skippedSheets.push(teamName);
        console.log(`Sheet "${teamName}" already exists - updated team name only`);
        return; // Skip to next team
      }
      
      // Create new sheet
      sheet = ss.insertSheet(teamName);
      
      // Team name header - don't merge to avoid freeze conflicts
      sheet.getRange(1, 1).setValue(`Team Name: ${teamName}`);
      sheet.getRange(1, 1).setBackground('#e3f2fd').setFontWeight('bold');

      // Column headers
      sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(2, 1, 1, headers.length)
        .setBackground('#0d47a1')
        .setFontColor('#ffffff')
        .setFontWeight('bold');

      // Girls section
      sheet.getRange(3, 1).setValue('Girls');
      sheet.getRange(3, 1).setFontWeight('bold');

      // Add blank rows for girls
      for (let i = 4; i < 4 + girlsCount; i++) {
        sheet.getRange(i, 1, 1, headers.length).setValues([Array(headers.length).fill('')]);
      }

      // Calculate row for boys section header (after girls + buffer)
      const boysHeaderRow = 4 + girlsCount + 5; // 5 row buffer
      
      // Add column headers again for boys section
      sheet.getRange(boysHeaderRow, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(boysHeaderRow, 1, 1, headers.length)
        .setBackground('#0d47a1')
        .setFontColor('#ffffff')
        .setFontWeight('bold');

      // Boys section
      sheet.getRange(boysHeaderRow + 1, 1).setValue('Boys');
      sheet.getRange(boysHeaderRow + 1, 1).setFontWeight('bold');

      // Add blank rows for boys
      for (let i = boysHeaderRow + 2; i < boysHeaderRow + 2 + boysCount; i++) {
        sheet.getRange(i, 1, 1, headers.length).setValues([Array(headers.length).fill('')]);
      }

      // Set up freeze panes (first 2 columns)
      sheet.setFrozenColumns(2);
      
      // Auto-resize columns for readability
      sheet.autoResizeColumns(1, 2);
      
      createdSheets.push(teamName);
      console.log(`Created blank relay entry sheet for ${teamName} (${girlsCount} girls, ${boysCount} boys)`);
    });
    
    // Show summary to user
    let message = '';
    if (createdSheets.length > 0) {
      message += `✅ Created ${createdSheets.length} new relay entry sheets:\n${createdSheets.join(', ')}\n\n`;
    }
    if (skippedSheets.length > 0) {
      message += `⚠️ Updated team names for ${skippedSheets.length} existing sheets:\n${skippedSheets.join(', ')}\n\n`;
    }
    message += `💡 Configure team names and swimmer counts in the "Team Relay Meet Config" sheet.`;
    
    SpreadsheetApp.getUi().alert('Relay Entry Sheets', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
/**
 * Creates or updates the Lane Assignments sheet based on the CSV format
 */
function setupLaneAssignments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let laneSheet = ss.getSheetByName('Lane Assignments');
  
  if (!laneSheet) {
    laneSheet = ss.insertSheet('Lane Assignments');
    
    // Set up headers based on CSV format
    const headers = ['Event #', 'Event Name', 'CENT', 'MVHS', 'MER', 'EAGLE', 'TIMBER'];
    laneSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Add sample data from the CSV
    const sampleData = [
      [1, '200 Medley Relay', 2, 3, 4, 5, 1],
      [2, 'JV 200 Medley Relay', 1, 4, 5, 2, 3],
      [3, '350 Free Relay', 4, 1, 2, 3, 5],
      [4, '200 Fly Relay', 5, 2, 3, 1, 4],
      [5, 'JV 200 Fly Relay', 1, 3, 4, 5, 2],
      [6, '200 Breast Relay', 3, 4, 5, 2, 1],
      [7, 'JV 200 Breast Relay', 4, 5, 2, 1, 3],
      [8, '200 Free Relay', 5, 1, 3, 4, 2],
      [9, 'JV 200 Free Relay', 2, 1, 4, 5, 3],
      [10, '400 IM Relay', 3, 4, 1, 2, 5],
      [11, 'JV 400 IM Relay', 1, 5, 2, 3, 4],
      [13, '200 Back Relay', 5, 2, 3, 4, 1],
      [14, 'JV 200 Back Relay', 1, 3, 4, 5, 2],
      [15, '400 Free Relay', 3, 4, 1, 2, 5],
      [16, 'JV 400 Free Relay', 4, 5, 2, 1, 3],
      [17, '200 Medley Relay Co-ed', 5, 1, 3, 4, 2],
      [18, '200 Frosh Free Relay', 2, 3, 1, 5, 4]
    ];
    
    laneSheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);
    
    // Format headers
    const headerRange = laneSheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#0d47a1');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    
    // Auto-resize columns
    laneSheet.autoResizeColumns(1, headers.length);
    
    console.log('Created Lane Assignments sheet with sample data');
  }
  
  return laneSheet;
}

/**
 * Gets lane assignments from the Lane Assignments sheet
 */
function getLaneAssignments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const laneSheet = ss.getSheetByName('Lane Assignments');
  
  if (!laneSheet) {
    throw new Error('Lane Assignments sheet not found. Please run "Setup Lane Assignments" first.');
  }
  
  const data = laneSheet.getDataRange().getValues();
  const headers = data[0];
  const assignments = [];
  
  // Parse lane assignments (skip header row)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const eventNum = row[0];
    const eventName = row[1];
    
    // Skip empty rows
    if (!eventNum || !eventName) continue;
    
    const laneAssignment = {
      eventNumber: eventNum,
      eventName: eventName.toString().trim(),
      lanes: {
        CENT: row[2],
        MVHS: row[3], 
        MER: row[4],
        EAGLE: row[5],
        TIMBER: row[6]
      }
    };
    
    assignments.push(laneAssignment);
  }
  
  return assignments;
}

/**
 * Gets relay entries from a specific team sheet
 */
function getTeamRelayEntries(teamName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teamSheet = ss.getSheetByName(teamName);
  
  if (!teamSheet) {
    console.log(`Team sheet "${teamName}" not found`);
    return {};
  }
  
  const data = teamSheet.getDataRange().getValues();
  if (data.length < 2) {
    console.log(`Team sheet "${teamName}" has no data`);
    return {};
  }
  
  // Find headers (skip team name header if present)
  let headerRow = 0;
  let headers = data[0];
  
  // Check if first row is team name header
  if (headers[0] && headers[0].toString().includes('Team Name:')) {
    headerRow = 1;
    headers = data[1];
  }
  
  const entries = {};
  
  // Process all swimmers
  for (let i = headerRow + 1; i < data.length; i++) {
    const row = data[i];
    const swimmerName = row[0];
    
    // Skip empty rows or section headers
    if (!swimmerName || swimmerName === 'Girls' || swimmerName === 'Boys') {
      continue;
    }
    
    // Check each event column for 'X' marks
    for (let j = 2; j < headers.length && j < row.length; j++) {
      const eventName = headers[j];
      const isAssigned = row[j] === 'X' || row[j] === 'x';
      
      if (isAssigned && eventName) {
        if (!entries[eventName]) {
          entries[eventName] = [];
        }
        entries[eventName].push(swimmerName.toString().trim());
      }
    }
  }
  
  return entries;
}

/**
 * Generates a relay meet heat sheet and creates a Google Doc
 */
function generateRelayHeatSheet() {
  try {
    // Ensure Lane Assignments sheet exists
    setupLaneAssignments();
    
    // Get lane assignments
    const laneAssignments = getLaneAssignments();
    
    // Get all team entries
    const teamNames = ['CENT', 'MVHS', 'MER', 'EAGLE', 'TIMBER'];
    const allTeamEntries = {};
    
    teamNames.forEach(teamName => {
      allTeamEntries[teamName] = getTeamRelayEntries(teamName);
    });
    
    // Create Google Doc
    const doc = DocumentApp.create('Relay Meet Heat Sheet - ' + new Date().toLocaleDateString());
    const body = doc.getBody();
    
    // Document title
    body.appendParagraph('RELAY MEET HEAT SHEET').setHeading(DocumentApp.ParagraphHeading.TITLE);
    body.appendParagraph('Generated: ' + new Date().toLocaleString()).setAttributes({
      [DocumentApp.Attribute.FONT_SIZE]: 10,
      [DocumentApp.Attribute.ITALIC]: true
    });
    body.appendParagraph(''); // Blank line
    
    // Group events by heat number (based on event order)
    const heats = [];
    let currentHeat = [];
    
    laneAssignments.forEach((assignment, index) => {
      currentHeat.push(assignment);
      
      // Create heat every 3-4 events or at end
      if (currentHeat.length >= 3 || index === laneAssignments.length - 1) {
        heats.push([...currentHeat]);
        currentHeat = [];
      }
    });
    
    // Generate heat sheets
    heats.forEach((heat, heatIndex) => {
      // Heat header
      body.appendParagraph(`HEAT ${heatIndex + 1}`).setHeading(DocumentApp.ParagraphHeading.HEADING1);
      
      heat.forEach(assignment => {
        // Event header
        body.appendParagraph(`Event ${assignment.eventNumber}: ${assignment.eventName}`)
          .setHeading(DocumentApp.ParagraphHeading.HEADING2);
        
        // Create lanes table
        const table = body.appendTable();
        
        // Lane headers
        const headerRow = table.appendTableRow();
        headerRow.appendTableCell('Lane');
        headerRow.appendTableCell('Team');
        headerRow.appendTableCell('Swimmers');
        
        // Style header row
        for (let i = 0; i < headerRow.getNumCells(); i++) {
          const cell = headerRow.getCell(i);
          cell.setBackgroundColor('#e3f2fd');
          cell.getChild(0).asText().setBold(true);
        }
        
        // Sort teams by lane assignment
        const laneOrder = [];
        Object.entries(assignment.lanes).forEach(([team, lane]) => {
          laneOrder.push({ team, lane: parseInt(lane) });
        });
        laneOrder.sort((a, b) => a.lane - b.lane);
        
        // Add team rows
        laneOrder.forEach(({ team, lane }) => {
          const row = table.appendTableRow();
          row.appendTableCell(lane.toString());
          row.appendTableCell(team);
          
          // Get swimmers for this event
          const teamEntries = allTeamEntries[team] || {};
          const swimmers = teamEntries[assignment.eventName] || [];
          
          if (swimmers.length > 0) {
            row.appendTableCell(swimmers.join(', '));
          } else {
            row.appendTableCell('No entries found');
          }
        });
        
        body.appendParagraph(''); // Blank line after each event
      });
      
      body.appendPageBreak(); // New page for each heat
    });
    
    // Open the document
    const url = doc.getUrl();
    
    SpreadsheetApp.getUi().alert(
      'Heat Sheet Generated!',
      `Your relay meet heat sheet has been created.\n\nDocument: ${doc.getName()}\nURL: ${url}\n\nThe document includes ${heats.length} heats with ${laneAssignments.length} events total.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    console.log('Generated heat sheet:', url);
    return doc;
    
  } catch (error) {
    console.error('Error generating heat sheet:', error);
    SpreadsheetApp.getUi().alert(
      'Error',
      'Failed to generate heat sheet: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

  } catch (error) {
    console.error('Error creating blank relay entry sheets:', error);
    SpreadsheetApp.getUi().alert(
      'Error',
      'Failed to create blank relay entry sheets: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}
