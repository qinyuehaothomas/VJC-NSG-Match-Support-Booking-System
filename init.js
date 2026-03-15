// ============================================================
// create_sheet.js — Creates the Master Google Spreadsheet
// ============================================================

/**
 * Creates the Master Spreadsheet with all required tabs.
 * @param {string[]} classes - List of class names
 * @param {string[]} sports  - List of sport names
 * @returns {string} Spreadsheet ID
 */
function createMasterSpreadsheet(classes, sports) {
  const year = new Date().getFullYear();
  const ss   = SpreadsheetApp.create(`NSG Match Support — Master ${year}`);
  const ssId = ss.getId();

  // ── Timetable tab ─────────────────────────────────────────
  const timetableSheet = ss.getSheetByName('Sheet1');
  timetableSheet.setName('Timetable');
  timetableSheet.appendRow(['Class', 'Date', 'StartTime', 'EndTime', 'Subject']);
  formatHeaderRow(timetableSheet, '#1a73e8');

  // ── Preference tab ────────────────────────────────────────
  const prefSheet = ss.insertSheet('Preference');
  // Headers will be set when the form is linked
  formatHeaderRow(prefSheet, '#0f9d58');

  // ── Matches tab ───────────────────────────────────────────
  const matchSheet = ss.insertSheet('Matches');
  matchSheet.appendRow([
    'MatchID', 'Name', 'Sport', 'Date',
    'EstimatedStartTime', 'EstimatedEndTime', 'EstimatedCapacity', 'Venue'
  ]);
  formatHeaderRow(matchSheet, '#f4b400');

  // ── Allocation tab ────────────────────────────────────────
  const allocSheet = ss.insertSheet('Allocation');
  allocSheet.appendRow(['Class', 'MatchID']);
  formatHeaderRow(allocSheet, '#db4437');

  Logger.log('Master spreadsheet created: ' + ssId);
  return ssId;
}

/**
 * Bold + colour the first row of a sheet.
 * @param {Sheet} sheet
 * @param {string} bgColor - Hex colour string
 */
function formatHeaderRow(sheet, bgColor) {
  const range = sheet.getRange(1, 1, 1, sheet.getLastColumn() || 8);
  range.setBackground(bgColor)
       .setFontColor('#ffffff')
       .setFontWeight('bold');
  sheet.setFrozenRows(1);
}

// ============================================================
// create_form.js — Creates Class Preference & Matches Google Forms
// ============================================================

/**
 * Creates both Google Forms and links them to the Master Spreadsheet.
 * @param {string[]} classes        - List of class names
 * @param {string[]} sports         - List of sport names
 * @param {string}   spreadsheetId  - Master spreadsheet ID
 * @returns {{ classPreference: string, matches: string }} Form IDs
 */
function createForms(classes, sports, spreadsheetId) {
  const ss = SpreadsheetApp.openById(spreadsheetId);

  const prefFormId   = createClassPreferenceForm(classes, sports, ss);
  const matchFormId  = createMatchesForm(sports, ss);

  return {
    classPreference: prefFormId,
    matches:         matchFormId
  };
}

// ── Class Preference Form ─────────────────────────────────────

/**
 * Creates the Class Preference Google Form.
 * Q1: Class dropdown
 * Q2: Sport ranking (Multiple Choice Grid)
 * Q3: Classmate involvement (Checkboxes)
 * @returns {string} Form ID
 */
function createClassPreferenceForm(classes, sports, ss) {
  const form = FormApp.create('NSG Class Preference');
  form.setDescription(
    'Please fill in your class preferences for the National School Games.'
  );
  form.setCollectEmail(false);

  // Q1 — Class (dropdown, required)
  const q1 = form.addListItem();
  q1.setTitle('Class')
    .setChoiceValues(classes)
    .setRequired(true);

  // Q2 — Sport Ranking (Multiple Choice Grid)
  // Rows = ranks 1..n, Columns = sports
  const n    = sports.length;
  const rows = Array.from({ length: n }, (_, i) => String(i + 1));

  const q2 = form.addGridItem();
  q2.setTitle('Rank your preferred sports (1 = most preferred)')
    .setRows(rows)
    .setColumns(sports)
    .setRequired(true);

  try {
    // Restrict one response per column (each sport ranked once)
    q2.setValidation(
      FormApp.createGridValidation()
        .requireLimitOneResponsePerColumn()
        .build()
    );
  } catch (e) {
    Logger.log('Grid validation not supported in this context: ' + e);
  }

  // Q3 — Classmate Involvement (Checkboxes, not required)
  const q3 = form.addCheckboxItem();
  q3.setTitle('Which sports have classmates participating in?')
    .setChoiceValues(sports)
    .setRequired(false);

  // Link to Preference sheet
  const prefSheet = ss.getSheetByName('Preference');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  // Move the linked sheet to the 'Preference' position
  Utilities.sleep(2000); // Give Forms time to create the linked sheet
  renameLinkedFormSheet(ss, form.getId(), 'Preference');

  Logger.log('Class Preference Form created: ' + form.getId());
  return form.getId();
}

// ── Matches Form ──────────────────────────────────────────────

/**
 * Creates the Matches Google Form.
 * Questions: Name, Sport, Date, Est Start Time, Est End Time, Est Capacity
 * @returns {string} Form ID
 */
function createMatchesForm(sports, ss) {
  const form = FormApp.create('NSG Match Entry');
  form.setDescription(
    'Enter details for each National School Games match.'
  );
  form.setCollectEmail(false);

  // Match Name (Short answer, required)
  form.addTextItem()
      .setTitle('Match Name / Description')
      .setRequired(true);

  // Sport (Multiple Choice, required)
  form.addMultipleChoiceItem()
      .setTitle('Sport')
      .setChoiceValues(sports)
      .setRequired(true);

  // Date (Date question, required)
  form.addDateItem()
      .setTitle('Date')
      .setRequired(true);

  // Estimated Start Time (Time question, required)
  form.addTimeItem()
      .setTitle('Estimated Start Time')
      .setRequired(true);

  // Estimated End Time (Time question, required)
  form.addTimeItem()
      .setTitle('Estimated End Time')
      .setRequired(true);

  // Estimated Capacity (Short answer with number validation, required)
  const capItem = form.addTextItem();
  capItem.setTitle('Estimated Capacity (number of students)')
         .setRequired(true);
  try {
    capItem.setValidation(
      FormApp.createTextValidation()
        .requireNumberGreaterThan(0)
        .build()
    );
  } catch (e) {
    Logger.log('Capacity validation skipped: ' + e);
  }

  // Venue (Short answer, optional — can be updated later)
  form.addTextItem()
      .setTitle('Venue (leave blank if TBC)')
      .setRequired(false);

  // Link to Matches sheet
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  Utilities.sleep(2000);
  renameLinkedFormSheet(ss, form.getId(), 'Matches');

  // Add MatchID formula to the sheet after linking
  addMatchIdColumn(ss);

  Logger.log('Matches Form created: ' + form.getId());
  return form.getId();
}

// ── Helper ────────────────────────────────────────────────────

/**
 * After a Form is linked to a spreadsheet, Google creates a new sheet
 * named "Form Responses N". This renames it to the target name.
 */
function renameLinkedFormSheet(ss, formId, targetName) {
  const sheets = ss.getSheets();
  for (const sheet of sheets) {
    const name = sheet.getName();
    if (name.startsWith('Form Responses')) {
      // Check if this sheet is linked to our form
      try {
        const form = FormApp.openById(formId);
        if (form.getDestinationId() === ss.getId()) {
          // Remove old sheet with target name if it exists
          const existing = ss.getSheetByName(targetName);
          if (existing && existing.getSheetId() !== sheet.getSheetId()) {
            ss.deleteSheet(existing);
          }
          sheet.setName(targetName);
          return;
        }
      } catch (e) {
        Logger.log('renameLinkedFormSheet error: ' + e);
      }
    }
  }
}

/**
 * Adds a MatchID auto-formula column to the Matches sheet.
 * MatchID = "M" + row number, computed via formula.
 */
function addMatchIdColumn(ss) {
  Utilities.sleep(1000);
  const sheet = ss.getSheetByName('Matches');
  if (!sheet) return;

  // Insert MatchID as the first column
  sheet.insertColumnBefore(1);
  sheet.getRange(1, 1).setValue('MatchID');
  formatHeaderRow(sheet, '#f4b400');

  // Add formula from row 2 onwards (up to 1000 responses)
  const formula = '=IF(B2<>"","M"&ROW()-1,"")';
  sheet.getRange(2, 1).setFormula(formula);
  // Auto-fill pattern hint — users can extend or we can do it in allocation
}

// ============================================================
// create_site.js — Creates the Google Site with all pages
// ============================================================

/**
 * Creates the NSG Match Support Google Site.
 * Pages: Main, View/ByClass, View/ByEvent,
 *        Update/ClassPreference, Update/Timetable, Update/Matches
 *
 * NOTE: Google Sites (new) has limited Apps Script API support.
 * This function creates a Classic Google Sites structure.
 * If your domain uses New Google Sites, see the fallback comment below.
 *
 * @param {{ classPreference: string, matches: string }} formIds
 * @param {string} spreadsheetId
 * @returns {string} Site URL
 */
function createSite(formIds, spreadsheetId) {
  const year  = new Date().getFullYear();
  const name  = `NSG Match Support ${year}`;

  let site;
  try {
    site = SitesApp.createSite('', slugify(name), name,
      'National School Games — Match Support Management Portal');
  } catch (e) {
    // If domain-level creation is restricted, create under user sites
    try {
      site = SitesApp.createSite(name, name,
        'National School Games — Match Support Management Portal');
    } catch (e2) {
      Logger.log('Site creation failed: ' + e2 +
        '\nFalling back to returning a placeholder URL.');
      // Return a note — teacher can manually create site and embed the HTML
      return 'SITE_CREATION_FAILED__SEE_LOG';
    }
  }

  // ── Main (navigation) page ─────────────────────────────────
  const mainPage = site.getRootPage();
  mainPage.setName('Main');
  mainPage.setPageContent(buildMainPageHtml(site, year));

  // ── View / ByClass ─────────────────────────────────────────
  const byClassPage = site.createAnnouncementsPage('ByClass'); // reuse WebPage
  // Actually use createWebPage for content pages:
  site.createWebPage('ByClass', 'ByClass',
    buildEmbedPageHtml('DisplayByClass',
      getScriptUrl(),
      'View Schedule by Class'));

  // ── View / ByEvent ─────────────────────────────────────────
  site.createWebPage('ByEvent', 'ByEvent',
    buildEmbedPageHtml('DisplayByEvent',
      getScriptUrl(),
      'View Schedule by Event'));

  // ── Update / ClassPreference ───────────────────────────────
  const prefFormUrl = FormApp.openById(formIds.classPreference).getPublishedUrl();
  site.createWebPage('ClassPreference', 'ClassPreference',
    buildFormEmbedHtml(prefFormUrl, 'Class Preference'));

  // ── Update / Timetable ─────────────────────────────────────
  site.createWebPage('Timetable', 'Timetable',
    buildTimetableEmbedHtml(getScriptUrl()));

  // ── Update / Matches ──────────────────────────────────────
  const matchFormUrl = FormApp.openById(formIds.matches).getPublishedUrl();
  site.createWebPage('Matches', 'Matches',
    buildFormEmbedHtml(matchFormUrl, 'Match Entry'));

  const siteUrl = site.getUrl();
  Logger.log('Site created: ' + siteUrl);
  return siteUrl;
}

// ── Page HTML builders ────────────────────────────────────────

function buildMainPageHtml(site, year) {
  return `
<h1>NSG Match Support ${year}</h1>
<p>Welcome to the National School Games Match Support Management Portal.</p>
<h2>Navigation</h2>
<ul>
  <li><a href="ByClass">📅 View Schedule by Class</a></li>
  <li><a href="ByEvent">🏅 View Schedule by Event</a></li>
  <li><a href="ClassPreference">📝 Update Class Preference</a></li>
  <li><a href="Timetable">📤 Upload Timetable</a></li>
  <li><a href="Matches">➕ Add / Update Match</a></li>
</ul>`;
}

function buildEmbedPageHtml(functionName, scriptUrl, title) {
  return `
<h2>${title}</h2>
<iframe src="${scriptUrl}?page=${functionName}"
  width="100%" height="800"
  style="border:none;"></iframe>`;
}

function buildFormEmbedHtml(formUrl, title) {
  return `
<h2>${title}</h2>
<iframe src="${formUrl}?embedded=true"
  width="100%" height="900"
  frameborder="0" marginheight="0" marginwidth="0">
  Loading…
</iframe>`;
}

function buildTimetableEmbedHtml(scriptUrl) {
  return `
<h2>Upload Timetable</h2>
<iframe src="${scriptUrl}?page=Timetable"
  width="100%" height="700"
  style="border:none;"></iframe>`;
}

// ── Utilities ─────────────────────────────────────────────────

function slugify(str) {
  return str.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '');
}

/**
 * Returns the deployed web app URL for this script.
 * In production this will be the /exec URL; during development it may be /dev.
 */
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

