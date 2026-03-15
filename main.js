/**
 * Serves the appropriate HTML page based on configuration.
 * If config is missing or the year is not the current year, serve initialize.html,
 * otherwise serve main.html.
 */
function doGet() {
  const props = PropertiesService.getScriptProperties();
  const configJson = props.getProperty('config');
  const currentYear = new Date().getFullYear().toString();

  let needsInit = true;

  // Check Year
  if (configJson) {
    try {
      const config = JSON.parse(configJson);
      if (config.year === currentYear) {
        needsInit = false;
      }
    } catch (e) {
      // If config is corrupted, treat as needing init
      console.error('Config parse error:', e);
    }
  }

  if (needsInit) {
    return HtmlService.createHtmlOutputFromFile('initialise')
      .setTitle('Initialise Match Support System')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else {
    return HtmlService.createTemplateFromFile('main')
      .evaluate()
      .setTitle('Match Support Allocation')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

/**
 * Empty endpoint for initialisation form submission.
 * You will later implement the logic to save classes, sports, term date,
 * and timetable data, and then store the current year in config.
 */
function initialise(classes, sports, termDate, timetable) {
  // TODO: Save the provided data to your spreadsheet(s)
  // and store the current year in config.

  // Example: save config with current year
  const config = {
    year: new Date().getFullYear().toString()
  };
  PropertiesService.getScriptProperties().setProperty('config', JSON.stringify(config));

  // Return something (optional) – the client expects a success response.
  return;
}