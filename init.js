// init.js — Initialisation logic, called from initialisation via google.script.run
//
// Entry point: initialisation(classes, sports, termStartDate, timetableData)
//   - classes       : string[]  — list of class names e.g. ["1A","1B"]
//   - sports        : string[]  — list of sport names e.g. ["Basketball","Volleyball"]
//   - termStartDate : string    — ISO date string of Term 2 Week 1 Monday e.g. "2025-03-10"
//   - timetableData : object[]  — parsed timetable rows, processed client-side before sending
//                                 Each row: { className, day (0–13), period, lesson }

// ---------------------------------------------------------------------------
// Main initialisation function — called from client
// ---------------------------------------------------------------------------

const CLASS_PREFERENCE_SHEET_NAME="Class Preference";
const MATCH_DETAIL_SHEET_NAME="Match Details";

function initialisation(classes, sports, termStartDate, timetableData) {
  try {
    var year = new Date().getFullYear();

    // Step 1: Create master spreadsheet
    var spreadsheetId = createMasterSpreadsheet(classes, timetableData, year);

    // Step 2: Create Class Preference form and link to master
    var prefFormId = createClassPreferenceForm(classes, sports, spreadsheetId, year);

    // Step 3: Create Match Info Update form and link to master
    var matchFormId = createMatchUpdateForm(sports, spreadsheetId, year);

    // Step 4: Save config to script properties
    var config = {
      year:          year,
      termStartDate: termStartDate,
      classes:       classes,
      sports:        sports,
      spreadsheetId: spreadsheetId,
      forms: {
        classPreference: prefFormId,
        matchUpdate:     matchFormId
      }
    };
    saveConfig(config);

    // Step 5: Schedule async rename of form response sheets (avoids sleep delays)
    scheduleFormSheetRename();

    // Return success to client
    return { success: true, spreadsheetId: spreadsheetId, prefFormId: prefFormId, matchFormId: matchFormId };

  } catch (e) {
    // Return failure details to client
    return { success: false, error: e.message };
  }
}


// ---------------------------------------------------------------------------
// Create Master Spreadsheet
// ---------------------------------------------------------------------------
function createMasterSpreadsheet(classes, timetableData, year) {
  var ss = SpreadsheetApp.create(year + ' VJC NSG Allocation System Database');

  // ── Allocation tab ───────────────────────────────────────────────────────
  // Stores final decisions: one row per class-match pairing
  var allocationSheet = ss.getSheetByName('Sheet1'); // rename the default sheet
  allocationSheet.setName('Allocation');
  allocationSheet.appendRow(['Class', 'Match']);

  // ── One timetable tab per class ──────────────────────────────────────────
  // Tab name: "<ClassName> Timetable"
  // Row 1 (header): ["Day", "Period 0", "Period 1", ..., "Period 28"]
  // Rows 2–15: one row per day (Day 0 to Day 13), displayed as Day 1 to Day 14
  // Column 1: day names, Columns 2–30: periods 0–28
  // Cells contain lesson name; multiple lessons in same period are comma-separated
  classes.forEach(function(className) {
    var classSheet = ss.insertSheet(className + ' Timetable');

    // Build heading row: ["Day", "Period 0", "Period 1", ..., "Period 28"]
    var headingRow = ['Day'];
    for (var p = 0; p <= 28; p++) {
      headingRow.push('Period ' + p);
    }
    classSheet.appendRow(headingRow);

    // Build 14 day rows (Day 0 to Day 13)
    // grid[day][period] holds an array of lesson names (for comma-join)
    var grid = [];
    for (var d = 0; d < 14; d++) {
      grid[d] = [];
      for (var p = 0; p <= 28; p++) {
        grid[d][p] = [];
      }
    }

    // Populate grid from timetable data
    // timetableData format: { className: { days: [ [period, lesson], ... ] } }
    if (timetableData[className] && timetableData[className].days) {
      var classDays = timetableData[className].days;
      
      // Iterate through each day (0-13)
      for (var d = 0; d < classDays.length && d < 14; d++) {
        var dayLessons = classDays[d];
        
        // dayLessons is an array of [period, lesson] tuples
        if (Array.isArray(dayLessons)) {
          for (var i = 0; i < dayLessons.length; i++) {
            var tuple = dayLessons[i];
            var period = tuple[0]; // 0-indexed
            var lesson = tuple[1];
            
            if (period >= 0 && period <= 28) {
              grid[d][period].push(lesson);
            }
          }
        }
      }
    }

    // Write each day row to the sheet
    for (var d = 0; d < 14; d++) {
      var dataRow = ['Day ' + (d + 1)]; // Display as Day 1 to Day 14
      for (var p = 0; p <= 28; p++) {
        dataRow.push(grid[d][p].join(', ')); // comma-separated if multiple lessons
      }
      classSheet.appendRow(dataRow);
    }
  });

  return ss.getId();
}


// ---------------------------------------------------------------------------
// Create Class Preference Google Form
// ---------------------------------------------------------------------------
function createClassPreferenceForm(classes, sports, spreadsheetId, year) {
  var form = FormApp.create(year + ' Match Support Class Preference');
  form.setDescription('Fill this in to indicate your class preferences for NSG match support.');

  // Q1 — Class selection (dropdown, required)
  form.addListItem()
      .setTitle('Class')
      .setChoiceValues(classes)
      .setRequired(true);

  // Q2 — Sport ranking grid
  // Rows = sports, Columns = rank numbers (1 to n)
  // "Require one response per row" and "Limit to one response per column"
  // ensures each sport gets a unique rank
  var rankLabels = sports.map(function(_, i) { return String(i + 1); });
  var rankGrid   = form.addGridItem();
  rankGrid.setTitle('Rank each sport (1 = most preferred, no duplicates)')
          .setRows(sports)
          .setColumns(rankLabels)
          .setRequired(true);
  try {
    rankGrid.setValidation(
      FormApp.createGridValidation()
        .requireLimitOneResponsePerColumn()
        .build()
    );
  } catch (e) {
    Logger.log('Grid validation not supported: ' + e.message);
  }

  // Q3 — Classmate involvement checkboxes (not required)
  form.addCheckboxItem()
      .setTitle('Which sports have members of your class participating in?')
      .setChoiceValues(sports)
      .setRequired(false);

  // Link form responses to the master spreadsheet
  form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetId);

  return form.getId();
}


// ---------------------------------------------------------------------------
// Create Match Info Update Google Form
// ---------------------------------------------------------------------------
function createMatchUpdateForm(sports, spreadsheetId, year) {
  var form = FormApp.create(year + ' NSG Match Info Update');
  form.setDescription('Use this form to add or update NSG match details.');

  // Q1 — Match name (short answer, required)
  form.addTextItem()
      .setTitle('Match Name')
      .setRequired(true);

  // Q2 — Sport (multiple choice from sports list, required)
  form.addMultipleChoiceItem()
      .setTitle('Sport')
      .setChoiceValues(sports)
      .setRequired(true);

  // Q3 — Estimated number of classes attending (short answer, required)
  form.addTextItem()
      .setTitle('Estimated Number of Classes')
      .setRequired(true);

  // Q4 — Match date (date picker, required)
  form.addDateItem()
      .setTitle('Date')
      .setRequired(true);

  // Q5 — Estimated leave time (time picker, required)
  form.addTimeItem()
      .setTitle('Estimated Leave Time')
      .setRequired(true);

  // Q6 — Estimated return time (time picker, required)
  form.addTimeItem()
      .setTitle('Estimated Return Time')
      .setRequired(true);

  // Link form responses to the master spreadsheet
  form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetId);

  return form.getId();
}


// ---------------------------------------------------------------------------
// Schedule async renaming of form response sheets
// Runs 2 minutes after initialisation, avoids blocking the main function with sleep()
// ---------------------------------------------------------------------------
function scheduleFormSheetRename() {
  ScriptApp.newTrigger('renameFormResponseSheets')
    .timeBased()
    .after(2 * 60 * 1000) // 2 minutes in milliseconds
    .create();
}

// Triggered ~2 minutes after initialisation
// Finds each form's response sheet by form ID and renames it to the correct tab name
// Then deletes the placeholder tabs created during spreadsheet setup
function renameFormResponseSheets() {
  var config = loadConfig();
  var ss     = SpreadsheetApp.openById(config.spreadsheetId);

  // Rename the form response sheets using form ID (stable, not dependent on tab name)
  getFormResponseSheet(ss, config.forms.classPreference).setName(CLASS_PREFERENCE_SHEET_NAME);
  getFormResponseSheet(ss, config.forms.matchUpdate).setName(MATCH_DETAIL_SHEET_NAME);

  // Clean up this trigger so it does not run again
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'renameFormResponseSheets') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

// Returns the sheet in ss whose linked form URL contains the given formId
function getFormResponseSheet(ss, formId) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var url = sheets[i].getFormUrl();
    if (url && url.indexOf(formId) !== -1) {
      return sheets[i];
    }
  }
  throw new Error('No response sheet found for form ID: ' + formId);
}