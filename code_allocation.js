// Constants
const MAX_MATCHES_PER_CLASS = 2;
const MIN_CLASSES_PER_MATCH = 5;
const DEFAULT_MATCH_CAPACITY = 300;

// Main entry point - triggered on form submission
function allocation() {
  Logger.log("Alloc Function triggered");
  try {
    const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('spreadsheetId'));
    const classPrefs = readClassPreferences(ss);
    const matchDetails = readMatchDetails(ss);
    const allocations = runAllocationAlgorithm(classPrefs, matchDetails, ss);
    writeAllocationsToSheet(ss, allocations);
    return { success: true, message: 'Allocation completed' };
  } catch (e) {
    Logger.log("Allocation failed: " + e.message);
    return { success: false, error: e.message };
  }
}

// Read Class Preference responses: [class, sportRankings{}, classmateInvolvement[]]
function readClassPreferences(ss) {
  
  const data = ss.getSheetByName(CLASS_PREFERENCE_SHEET_NAME).getDataRange().getValues();
  const prefs = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[1]) continue;
    
    const sportRankings = {};
    for (let j = 0; j < 10; j++) {
      const sport = (row[2 + j] || '').trim();
      if (sport) sportRankings[sport] = j + 1;
    }
    
    const involvement = (row[12] || '').split(',').map(s => s.trim()).filter(s => s);
    prefs.push({ class: row[1], sportRankings, classmateInvolvement: involvement });
  }
  return prefs;
}

// Read Match Info sheet: [sport, matchLevel, venue, numClasses, date, leaveTime, returnTime, cancelled]
function readMatchDetails(ss) {
  
  const data = ss.getSheetByName(MATCH_INFO_SHEET_NAME).getDataRange().getValues();
  const matches = [];
  
  for (let i = 1; i < data.length; i++) {
    matches.push({
      sport: data[i][0],
      matchLevel: data[i][1],
      venue: data[i][2],
      numClasses: data[i][3],
      date: data[i][4],
      leaveTime: data[i][5],
      returnTime: data[i][6],
      cancelled: data[i][7]
    });
  }
  return matches;
}

// Write allocations to Allocation sheet
function writeAllocationsToSheet(ss, allocations) {
  const sheet = ss.getSheetByName(ALLOCATION_SHEET_NAME);
  if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
  if (allocations.length > 0) sheet.getRange(2, 1, allocations.length, 3).setValues(allocations);
}

// Core allocation algorithm
function runAllocationAlgorithm(classPrefs, matchDetails, ss) {
  
  const today = new Date();
  
  // Deduplicate input by class name - same class only appears once
  const seenClasses = new Set();
  const uniquePrefs = classPrefs.filter(p => {
    if (seenClasses.has(p.class)) return false;
    seenClasses.add(p.class);
    return true;
  });
  
  // Get past allocations and future matches
  const determined = getDeterminedAllocations(ss, matchDetails, today);
  const future = matchDetails.filter(m => m.cancelled !== 'Yes' && new Date(m.date) >= today);
  
  // Initialize tracking: class -> count, match -> classes assigned
  const classCount = {};
  const matchAlloc = {};
  uniquePrefs.forEach(p => classCount[p.class] = 0);
  future.forEach(m => matchAlloc[m.sport + '|' + m.matchLevel] = []);
  
  // Add determined (past) allocations to current tracking
  determined.forEach(a => {
    classCount[a[0]]++;
    const k = a[1] + '|' + a[2];
    if (!matchAlloc[k]) matchAlloc[k] = [];
    matchAlloc[k].push(a[0]);
  });
  
  // Build score matrix: class -> sport -> score
  const scores = buildScoreMatrix(uniquePrefs);
  
  // Sort matches by capacity ascending (fill smallest first)
  const sorted = future.sort((a, b) => (parseInt(a.numClasses) || DEFAULT_MATCH_CAPACITY) - (parseInt(b.numClasses) || DEFAULT_MATCH_CAPACITY));
  
  // Greedy fill: for each match, add best available classes until full
  sorted.forEach(m => {
    const k = m.sport + '|' + m.matchLevel;
    const cap = parseInt(m.numClasses) || DEFAULT_MATCH_CAPACITY;
    
    while (matchAlloc[k].length < cap) {
      const best = uniquePrefs
        .filter(p => classCount[p.class] < MAX_MATCHES_PER_CLASS && !matchAlloc[k].includes(p.class))
        .sort((a, b) => (scores[b.class]?.[m.sport] || 0) - (scores[a.class]?.[m.sport] || 0))[0];
      
      if (!best) break;
      matchAlloc[k].push(best.class);
      classCount[best.class]++;
    }
  });
  
  // Convert to output format: [class, sport, matchLevel]
  const result = [];
  Object.entries(matchAlloc).forEach(([k, classes]) => {
    const [sport, level] = k.split('|');
    classes.forEach(c => result.push([c, sport, level]));
  });
  return result;
}

// Get past allocations (locked-in, before today)
function getDeterminedAllocations(ss, matchDetails, today) {
  const lookup = {};
  matchDetails.forEach(m => lookup[m.sport + '|' + m.matchLevel] = m);
  Logger.log(lookup);
  Logger.log(ss.getSheetByName(ALLOCATION_SHEET_NAME).getDataRange().getValues()
    .slice(1));
  return ss.getSheetByName(ALLOCATION_SHEET_NAME).getDataRange().getValues()
    .slice(1)
    .filter(row => row[0] && lookup[row[1] + '|' + row[2]] && new Date(lookup[row[1] + '|' + row[2]].date) < today);
}

// Build score matrix: score = rank bonus + involvement bonus
function buildScoreMatrix(prefs) {
  const matrix = {};
  const numSports = prefs[0] ? Object.keys(prefs[0].sportRankings || {}).length : 0;
  
  prefs.forEach(p => {
    matrix[p.class] = {};
    Object.entries(p.sportRankings).forEach(([sport, rank]) => {
      const rankScore = (numSports - rank + 1) * 2;
      const involvementBonus = p.classmateInvolvement.includes(sport) ? 1 : 0;
      matrix[p.class][sport] = rankScore + involvementBonus;
    });
  });
  return matrix;
}
