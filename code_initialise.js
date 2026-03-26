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


const CLASS_PREFERENCE_SHEET_NAME = "Class Preference";
const MATCH_UPDATE_SHEET_NAME = "Match Update";
const MATCH_INFO_SHEET_NAME = "Match Info";
const ALLOCATION_SHEET_NAME = "Allocation";
const DISPLAY_BY_CLASS_NAME = "Display By Class";
const DISPLAY_BY_EVENT_NAME = "Display By Event";


function initialisation(classes, sports, termStartDate, timetableData) {
  Logger.log("Initialisation Begins!");
  try {
    var year = new Date().getFullYear();

    // Step 1: Create master spreadsheet
    var spreadsheetId = createMasterSpreadsheet(classes, timetableData, year);
    
    Logger.log("SpreadsheetId",spreadsheetId);

    // Step 2: Create Class Preference form and link to master
    var prefFormId = createClassPreferenceForm(classes, sports, spreadsheetId, year);
    Logger.log("pref ID",prefFormId);

    // Step 3: Create Match Info Update form and link to master
    var matchFormId = createMatchUpdateForm(sports, spreadsheetId, year);
    Logger.log("match ID",matchFormId);
    
    const props = PropertiesService.getScriptProperties();
    
    props.setProperty('year', year.toString());
    props.setProperty('termStartDate', termStartDate.toString());
    props.setProperty('classes', classes.toString());
    props.setProperty('sports', sports.toString())
    props.setProperty('spreadsheetId', spreadsheetId);

    // Step 4: Create form submission triggers
    createFormSubmissionTriggers(prefFormId, matchFormId);

    // Return success to client
    return { success: true, spreadsheetId: spreadsheetId, prefFormId: prefFormId, matchFormId: matchFormId };

  } catch (e) {
    // throw new Error("Init Failed!",e.messgae);
    return { success: false, error: e.message };
  }
}


// ---------------------------------------------------------------------------
// Create Master Spreadsheet
// ---------------------------------------------------------------------------

function createMasterSpreadsheet(classes, timetableData, year) {
  // throw new Error(timetableData);
  var ss = SpreadsheetApp.create(year + ' VJC NSG Allocation System Database');

  // ── Allocation tab ───────────────────────────────────────────────────────
  // Stores final decisions: one row per class-match pairing
  var allocationSheet = ss.getSheetByName('Sheet1'); // rename the default sheet
  allocationSheet.setName(ALLOCATION_SHEET_NAME);
  allocationSheet.appendRow(['Class', 'Sport', 'Match Level']);

  // ── Match Info tab ───────────────────────────────────────────────────────
  // Stores consolidated match information (updated from Match Update form responses)
  var matchInfoSheet = ss.insertSheet(MATCH_INFO_SHEET_NAME);
  matchInfoSheet.appendRow(['Sport', 'Match Level', 'Venue', 'Num Classes', 'Date', 'Leave Time', 'Return Time', 'Cancelled']);

  

  // ── Display By Class tab ───────────────────────────────────────────────────────
  // Display Events for each class
  var matchInfoSheet = ss.insertSheet(DISPLAY_BY_CLASS_NAME);
  matchInfoSheet.appendRow(["Class", 'Sport', 'Match Level', 'Venue', 'Date', 'Leave Time', 'Return Time']);

  // ── Display By Event tab ───────────────────────────────────────────────────────
  // Display Class going to each event
  var matchInfoSheet = ss.insertSheet(DISPLAY_BY_EVENT_NAME);
  matchInfoSheet.appendRow(['Sport', 'Match Level', "Class",'Venue', 'Date', 'Leave Time', 'Return Time']);

  // ── One timetable tab per class ──────────────────────────────────────────
  // Tab name: "<ClassName> Timetable"
  // Row 1 (header): ["Day", "Period 0", "Period 1", ..., "Period 28"]
  // Rows 2–15: one row per day (Day 0 to Day 13), displayed as Day 1 to Day 14
  // Column 1: day names, Columns 2–30: periods 0–28
  // Cells contain lesson name; multiple lessons in same period are comma-separated

  /*
  Object.entries(timetableData).forEach(([className,classLessons])=> {

    var classSheet = ss.insertSheet(className + ' Timetable');

    // Build heading row: ["Day", "Period 0", "Period 1", ..., "Period 28"]
    var headingRow = ['Day'];
    for (var p = 0; p <= 28; p++) {
      headingRow.push('Period ' + p);
    }
    classSheet.appendRow(headingRow);

    // Populate grid from timetable data
    // timetableData format: { className: [[lesson 1, lesson 2],[...],[...],...] }

    // Write each day row to the sheet
    for(let d =0; d<14;d++){
      let dataRow = Array(30).fill("");
      
      dataRow[0]='Day ' + (parseInt(d) + 1); // Display as Day 1 to Day 14
      
      if(classLessons[d]) Object.entries(classLessons[d]).forEach(([idx,lesson_name])=>{
        
        if (!dataRow[idx].split(',').includes(lesson_name)){
          dataRow[idx]+= dataRow[idx]?","+lesson_name:lesson_name;
        }
      });
      
      
      classSheet.appendRow(dataRow);
      }
  
  });
  */

  Logger.log("Master sheet Successful");
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
  // Rows = sports, Columns = rank numbers (1 to 10)
  // "Require one response per row" and "Limit to one response per column"
  // ensures each sport gets a unique rank
  var rankLabels = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10'];
  var rankGrid   = form.addGridItem();
  rankGrid.setTitle('Rank each sport (1 = most preferred, up to 10th)')
          .setRows(rankLabels)
          .setColumns(sports)
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
  
  // Brief delay to ensure response sheet is created
  Utilities.sleep(500);
  
  // Rename response sheet
  var ss = SpreadsheetApp.openById(spreadsheetId);
  getFormResponseSheet(ss, form.getId()).setName(CLASS_PREFERENCE_SHEET_NAME);

  return form.getId();
}


// ---------------------------------------------------------------------------
// Create Match Info Update Google Form
// ---------------------------------------------------------------------------
function createMatchUpdateForm(sports, spreadsheetId, year) {
  var form = FormApp.create(year + ' NSG Match Info Update');
  form.setDescription('Use this form to add or update NSG match details.');

  // Q1 — Sport (multiple choice from sports list, required)
  form.addListItem()
      .setTitle('Sport')
      .setChoiceValues(sports)
      .setRequired(true);

  // Q2 — Match Level
  form.addMultipleChoiceItem()
      .setTitle('Match Level')
      .setChoiceValues(["Finals","Semi-Finals","3rd/ 4th"])
      .setRequired(true);

  // Q3 — Venue (short answer, NOT required)
  form.addTextItem()
      .setTitle('Venue')
      .setRequired(false);

  // Q4 — Estimated number of classes attending (short answer, NOT required)
  form.addTextItem()
      .setTitle('Estimated Number of Classes')
      .setRequired(false);

  // Q5 — Match date (date picker, NOT required)
  form.addDateItem()
      .setTitle('Date')
      .setRequired(false);

  // Q6 — Estimated leave time (time picker, NOT required)
  form.addTimeItem()
      .setTitle('Estimated Leave Time')
      .setRequired(false);

  // Q7 — Estimated return time (time picker, NOT required)
  form.addTimeItem()
      .setTitle('Estimated Return Time')
      .setRequired(false);

  // Q8 — Cancelled (checkbox, NOT required)
  form.addCheckboxItem()
      .setTitle('Is this match cancelled?')
      .setChoiceValues(['Yes'])
      .setRequired(false);

  // Link form responses to the master spreadsheet
  form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetId);
  
  // Brief delay to ensure response sheet is created
  Utilities.sleep(500);
  
  // Rename response sheet
  var ss = SpreadsheetApp.openById(spreadsheetId);
  getFormResponseSheet(ss, form.getId()).setName(MATCH_UPDATE_SHEET_NAME);

  return form.getId();
}

// Returns the sheet in ss whose linked form URL contains the given formId
function getFormResponseSheet(ss, formId) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var url = sheets[i].getFormUrl();
    if (url && url.includes(formId)) {
      return sheets[i];
    }
  }
  throw new Error('No response sheet found for form ID: ' + formId);
}


// ---------------------------------------------------------------------------
// Form submission triggers — called automatically on form submit
// ---------------------------------------------------------------------------

function createFormSubmissionTriggers(classPreferenceFormId, matchUpdateFormId) {
  try {
    // Open forms to get Form objects
    const classForm = FormApp.openById(classPreferenceFormId);
    const matchForm = FormApp.openById(matchUpdateFormId);
    
    // Create trigger for Class Preference form
    ScriptApp.newTrigger('onClassPreferenceSubmit')
      .forForm(classForm)
      .onFormSubmit()
      .create();
    Logger.log('Created trigger for Class Preference form');
    
    // Create trigger for Match Update form
    ScriptApp.newTrigger('onMatchUpdateSubmit')
      .forForm(matchForm)
      .onFormSubmit()
      .create();
    Logger.log('Created trigger for Match Update form');
  } catch (e) {
    Logger.log('Failed to create form triggers: ' + e.message);
  }
}

// Triggered when Class Preference form is submitted
function onClassPreferenceSubmit(e) {
  Logger.log('Class Preference form submitted');
  try {
    allocation();
  } catch (error) {
    Logger.log('Error in onClassPreferenceSubmit: ' + error.message);
  }
}

// Triggered when Match Update form is submitted
function onMatchUpdateSubmit(e) {
  Logger.log('Match Update form submitted');
  try {
    updateMatchInfo(e);
    allocation();
  } catch (error) {
    Logger.log('Error in onMatchUpdateSubmit: ' + error.message);
  }
}

// ---------------------------------------------------------------------------
// Update Match Info sheet from latest Match Update submission
// ---------------------------------------------------------------------------
function updateMatchInfo(e) {
  try {
    // Extract form response answers
    const itemResponses = e.response.getItemResponses();
    
    // Map answers: Q1=Sport(0), Q2=MatchLevel(1), Q3=Venue(2), Q4=NumClasses(3), Q5=Date(4), Q6=LeaveTime(5), Q7=ReturnTime(6), Q8=Cancelled(7)
    const sport = itemResponses[0] ? itemResponses[0].getResponse() : null;
    const matchLevel = itemResponses[1] ? itemResponses[1].getResponse() : null;
    const venue = itemResponses[2] ? itemResponses[2].getResponse() : '';
    const numClasses = itemResponses[3] ? itemResponses[3].getResponse() : '';
    const date = itemResponses[4] ? itemResponses[4].getResponse() : '';
    const leaveTime = itemResponses[5] ? itemResponses[5].getResponse() : '';
    const returnTime = itemResponses[6] ? itemResponses[6].getResponse() : '';
    const cancelled = (itemResponses[7] && itemResponses[7].getResponse().length > 0) ? 'Yes' : 'No';
    
    if (!sport || !matchLevel) {
      Logger.log('Sport or Match Level missing in submission');
      return;
    }
    
    const props = PropertiesService.getScriptProperties();
    const spreadsheetId = props.getProperty('spreadsheetId');
    
    if (!spreadsheetId) {
      throw new Error('SpreadsheetId not found in properties');
    }
    
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const infoSheet = ss.getSheetByName(MATCH_INFO_SHEET_NAME);
    
    if (!infoSheet) {
      throw new Error('Match Info sheet not found');
    }
    
    // Get all rows from Match Info sheet
    const infoData = infoSheet.getDataRange().getValues();
    
    // Find matching row by Sport + Match Level (composite key)
    let matchRowIndex = -1;
    for (let i = 1; i < infoData.length; i++) {
      if (infoData[i][0] === sport && infoData[i][1] === matchLevel) {
        matchRowIndex = i + 1; // Convert to 1-based row number
        break;
      }
    }
    
    if (matchRowIndex === -1) {
      // New match, append as new row
      infoSheet.appendRow([sport, matchLevel, venue, numClasses, date, leaveTime, returnTime, cancelled]);
      Logger.log('Added new match: ' + sport + ' - ' + matchLevel);
    } else {
      // Update existing match - use new value if provided, otherwise keep old value
      const oldRow = infoData[matchRowIndex - 1];
      const newVenue = venue || oldRow[2];
      const newNumClasses = numClasses || oldRow[3];
      const newDate = date || oldRow[4];
      const newLeaveTime = leaveTime || oldRow[5];
      const newReturnTime = returnTime || oldRow[6];
      
      infoSheet.getRange(matchRowIndex, 3, 1, 6).setValues([[newVenue, newNumClasses, newDate, newLeaveTime, newReturnTime, cancelled]]);
      Logger.log('Updated match: ' + sport + ' - ' + matchLevel);
    }
    
  } catch (e) {
    Logger.log('updateMatchInfo failed: ' + e.message);
  }
}
