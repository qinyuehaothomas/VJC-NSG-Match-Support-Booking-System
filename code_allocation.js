// alloc.js — Allocation logic
// Entry point: allocation()
//   - Retrieves spreadsheetId from PropertiesService
//   - Reads Class Preference and Match Details form responses
//   - Runs allocation algorithm (blank for now)
//   - Writes results to Allocation sheet

const CLASS_PREFERENCE_SHEET_NAME = "Class Preference";
const MATCH_DETAIL_SHEET_NAME = "Match Details";
const ALLOCATION_SHEET_NAME = "Allocation";

/**
 * Main allocation function
 * Called after form submissions via triggers
 * Retrieves spreadsheetId from PropertiesService
 */
function allocation() {
  Logger.log("Allocation triggered!");
  
  try {
    // Step 1: Get spreadsheetId from properties
    const props = PropertiesService.getScriptProperties();
    const spreadsheetId = props.getProperty('spreadsheetId');
    
    if (!spreadsheetId) {
      throw new Error('SpreadsheetId not found in properties. Has initialization run?');
    }
    
    // Step 2: Open spreadsheet
    const ss = SpreadsheetApp.openById(spreadsheetId);
    Logger.log("Opened spreadsheet: " + ss.getName());
    
    // Step 3: Read form responses
    const classPreferences = readClassPreferences(ss);
    const matchDetails = readMatchDetails(ss);
    
    Logger.log("Class Preferences: " + JSON.stringify(classPreferences));
    Logger.log("Match Details: " + JSON.stringify(matchDetails));
    
    // Step 4: Run allocation algorithm (blank for now)
    const allocations = runAllocationAlgorithm(classPreferences, matchDetails);
    
    Logger.log("Allocations result: " + JSON.stringify(allocations));
    
    // Step 5: Return success
    return { success: true, message: 'Allocation completed' };
    
  } catch (e) {
    Logger.log("Allocation failed: " + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Reads Class Preference form responses
 * Returns array of objects: [{ timestamp, class, sportRankings, classmateInvolvement }, ...]
 */
function readClassPreferences(ss) {
  const sheet = ss.getSheetByName(CLASS_PREFERENCE_SHEET_NAME);
  
  if (!sheet) {
    throw new Error('Class Preference sheet not found');
  }
  
  const data = sheet.getDataRange().getValues();
  
  if (data.length === 0) {
    return [];
  }
  
  const headers = data[0]; // First row is headers
  const preferences = [];
  
  // Parse each response row (starting from row 2)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // Build sport rankings object: { "Sport1": 1, "Sport2": 2, ... }
    const sportRankings = {};
    for (let j = 2; j < headers.length - 1; j++) { // Skip Timestamp (0) and Class (1), and last column (classmate involvement)
      const sport = headers[j];
      const rank = row[j];
      if (sport && rank) {
        sportRankings[sport] = parseInt(rank);
      }
    }
    
    // Get classmate involvement (last column, assuming it's a comma-separated list)
    const classmateInvolvementStr = row[headers.length - 1] || "";
    const classmateInvolvement = classmateInvolvementStr ? classmateInvolvementStr.split(',').map(s => s.trim()) : [];
    
    preferences.push({
      timestamp: row[0],
      class: row[1],
      sportRankings: sportRankings,
      classmateInvolvement: classmateInvolvement
    });
  }
  
  return preferences;
}

/**
 * Reads Match Details form responses
 * Returns array of objects: [{ timestamp, matchName, sport, numClasses, date, leaveTime, returnTime }, ...]
 */
function readMatchDetails(ss) {
  const sheet = ss.getSheetByName(MATCH_DETAIL_SHEET_NAME);
  
  if (!sheet) {
    throw new Error('Match Details sheet not found');
  }
  
  const data = sheet.getDataRange().getValues();
  
  if (data.length === 0) {
    return [];
  }
  
  const headers = data[0]; // First row is headers
  const matches = [];
  
  // Parse each response row (starting from row 2)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    matches.push({
      timestamp: row[0],
      matchName: row[1],
      sport: row[2],
      numClasses: row[3],
      date: row[4],
      leaveTime: row[5],
      returnTime: row[6]
    });
  }
  
  return matches;
}

/**
 * Allocation algorithm (blank implementation)
 * Input: classPreferences (array), matchDetails (array)
 * Output: allocations (array of { class, match } assignments)
 */
function runAllocationAlgorithm(classPreferences, matchDetails) {
  // TODO: Implement allocation algorithm
  // For now, return empty array
  return [];
}
