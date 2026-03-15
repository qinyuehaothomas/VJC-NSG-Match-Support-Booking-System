// ============================================================
// allocation.js — Match-to-Class allocation logic (Update function)
// ============================================================

/**
 * runAllocation() is triggered whenever:
 *   - A new Class Preference form response is received
 *   - A new Match form response is received
 *   - The timetable is updated
 *
 * Algorithm (preference-aware, capacity-respecting):
 *   1. Load all matches and classes with their preferences & timetables.
 *   2. For each match (sorted by date/time), determine which classes are
 *      available (no timetable clash) and have indicated involvement.
 *   3. Rank available classes by:
 *        a. Classmate Involvement flag (Q3) — hard preference
 *        b. Sport ranking (Q2) — lower rank number = higher priority
 *        c. Remaining capacity
 *   4. Assign top N classes up to capacity.
 *   5. Write results to the Allocation tab.
 */
function runAllocation() {
  const config = loadConfig();
  if (!config) {
    Logger.log('runAllocation: no config found, aborting.');
    return;
  }

  const ss = SpreadsheetApp.openById(config.spreadsheetId);

  // ── Load data ────────────────────────────────────────────
  const matches     = loadMatches(ss);
  const preferences = loadPreferences(ss, config.classes);
  const timetable   = loadTimetable(ss);

  if (matches.length === 0) {
    Logger.log('runAllocation: no matches found.');
    return;
  }

  // ── Allocation map: matchId -> [className] ───────────────
  const allocation = {};

  // Sort matches by date then start time
  matches.sort((a, b) => {
    const da = new Date(`${a.date} ${a.startTime}`);
    const db = new Date(`${b.date} ${b.startTime}`);
    return da - db;
  });

  // Track how many matches each class is already assigned to (to spread load)
  const classAssignmentCount = {};
  config.classes.forEach(c => classAssignmentCount[c] = 0);

  for (const match of matches) {
    const capacity = parseInt(match.capacity) || config.classes.length;
    const sport    = match.sport;

    // Build scored candidate list
    const candidates = [];

    for (const className of config.classes) {
      const pref = preferences[className] || {};

      // Check timetable clash
      if (hasClash(timetable, className, match.date, match.startTime, match.endTime)) {
        continue; // Skip: class has a lesson at this time
      }

      // Score: involvement flag is a strong positive signal
      const involvement = pref.involvement && pref.involvement.includes(sport) ? 1 : 0;

      // Sport rank: lower number = more preferred. Default to worst rank if not ranked.
      const sportRank = pref.sportRanks && pref.sportRanks[sport] != null
        ? pref.sportRanks[sport]
        : config.sports.length + 1;

      // Tie-break: prefer classes with fewer assignments (spread the load)
      const assignCount = classAssignmentCount[className] || 0;

      candidates.push({ className, involvement, sportRank, assignCount });
    }

    // Sort candidates: involvement desc, then sportRank asc, then assignCount asc
    candidates.sort((a, b) => {
      if (b.involvement !== a.involvement) return b.involvement - a.involvement;
      if (a.sportRank  !== b.sportRank)   return a.sportRank  - b.sportRank;
      return a.assignCount - b.assignCount;
    });

    // Assign top N up to capacity
    const assigned = candidates.slice(0, capacity).map(c => c.className);
    allocation[match.matchId] = assigned;
    assigned.forEach(c => classAssignmentCount[c]++);
  }

  // ── Write to Allocation sheet ────────────────────────────
  writeAllocation(ss, allocation);
  Logger.log('Allocation complete. ' + Object.keys(allocation).length + ' matches processed.');
}

// ── Data loaders ──────────────────────────────────────────────

/**
 * Loads all matches from the Matches sheet.
 * @returns {Object[]} Array of match objects
 */
function loadMatches(ss) {
  const sheet = ss.getSheetByName('Matches');
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const matches = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue; // Skip empty rows (no MatchID)
    const obj = {};
    headers.forEach((h, idx) => {
      // Normalise header to camelCase key
      obj[normHeader(h)] = row[idx];
    });
    matches.push(obj);
  }
  return matches;
}

/**
 * Loads all class preferences from the Preference sheet.
 * Returns a map: className -> { sportRanks: {sport: rank}, involvement: [sport] }
 */
function loadPreferences(ss, classes) {
  const sheet = ss.getSheetByName('Preference');
  if (!sheet || sheet.getLastRow() < 2) return {};

  const data    = sheet.getDataRange().getValues();
  const headers = data[0]; // Dynamic — depends on form column order

  // We need to identify which columns correspond to which question.
  // Google Forms response sheet has: Timestamp, then Q1, Q2 columns, Q3 column
  // Q2 is a grid — each sport gets its own column: "Rank your preferred sports [SportName]"
  // Q3 is a checkbox: "Which sports have classmates participating in?"

  const prefs = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // Find class name column (Q1: "Class")
    const classColIdx = headers.findIndex(h => h.toString().toLowerCase() === 'class');
    const className   = classColIdx >= 0 ? row[classColIdx] : null;
    if (!className || !classes.includes(className)) continue;

    // Parse sport ranks from grid columns
    // Grid columns look like: "Rank your preferred sports [Basketball]"
    const sportRanks = {};
    headers.forEach((h, idx) => {
      const match = h.toString().match(/\[(.+?)\]/);
      if (match && h.toString().toLowerCase().includes('rank')) {
        const sport = match[1];
        const rank  = parseInt(row[idx]);
        if (!isNaN(rank)) sportRanks[sport] = rank;
      }
    });

    // Parse involvement checkboxes
    // Column looks like: "Which sports have classmates participating in?"
    const invColIdx  = headers.findIndex(h =>
      h.toString().toLowerCase().includes('classmate')
    );
    const invRaw     = invColIdx >= 0 ? row[invColIdx] : '';
    const involvement = invRaw ? invRaw.toString().split(', ') : [];

    // Store last preference per class (in case of multiple submissions)
    prefs[className] = { sportRanks, involvement };
  }

  return prefs;
}

/**
 * Loads timetable data.
 * Returns array of {className, date, startTime, endTime, subject}
 */
function loadTimetable(ss) {
  const sheet = ss.getSheetByName('Timetable');
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows    = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const obj = {};
    headers.forEach((h, idx) => obj[h] = row[idx]);
    rows.push({
      className: obj['Class'],
      date:      formatDate(obj['Date']),
      startTime: obj['StartTime'],
      endTime:   obj['EndTime'],
      subject:   obj['Subject']
    });
  }
  return rows;
}

// ── Helpers ───────────────────────────────────────────────────

/**
 * Returns true if the class has a timetable entry clashing with the match window.
 */
function hasClash(timetable, className, matchDate, matchStart, matchEnd) {
  const matchDateStr = formatDate(matchDate);
  return timetable.some(entry =>
    entry.className === className &&
    entry.date      === matchDateStr &&
    timesOverlap(entry.startTime, entry.endTime, matchStart, matchEnd)
  );
}

/**
 * Writes the allocation map to the Allocation sheet.
 * Clears existing data (except header) first.
 */
function writeAllocation(ss, allocation) {
  let sheet = ss.getSheetByName('Allocation');
  if (!sheet) sheet = ss.insertSheet('Allocation');

  // Clear old data below header
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).clearContent();
  }

  // Write header if missing
  if (sheet.getLastRow() === 0 || sheet.getRange(1,1).getValue() !== 'Class') {
    sheet.getRange(1, 1, 1, 2).setValues([['Class', 'MatchID']]);
    formatHeaderRow(sheet, '#db4437');
  }

  const rows = [];
  Object.entries(allocation).forEach(([matchId, classes]) => {
    classes.forEach(cls => rows.push([cls, matchId]));
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }
}

/**
 * Normalise a header string to a simple camelCase key for match objects.
 * e.g. "Estimated Start Time" → "estimatedStartTime"
 * e.g. "MatchID" → "matchId"
 */
function normHeader(h) {
  return h.toString()
    .replace(/[^a-zA-Z0-9 ]/g, '')
    .split(' ')
    .map((w, i) => i === 0 ? w.toLowerCase() : w[0].toUpperCase() + w.slice(1).toLowerCase())
    .join('');
}

/**
 * Formats a Date object or date string to "YYYY-MM-DD".
 */
function formatDate(d) {
  if (!d) return '';
  const dt = d instanceof Date ? d : new Date(d);
  if (isNaN(dt)) return d.toString();
  return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}