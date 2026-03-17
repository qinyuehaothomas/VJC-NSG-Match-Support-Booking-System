/**
 * Serves the appropriate HTML page based on configuration.
 * If config is missing or the year is not the current year, serve initialize.html,
 * otherwise serve main.html.
 */
function doGet() {
  const props = PropertiesService.getScriptProperties();
  const configYear = props.getProperty('year');
  const currentYear = new Date().getFullYear().toString();

  let needsInit = true;

  // Check Year
  if (configYear) {
    try {
      if (configYear === currentYear) {
        needsInit = false;
      }
    } catch (e) {
      // If config is corrupted, treat as needing init
      console.error('Config parse error:', e);
    }
  }

  if (needsInit) {
    return HtmlService.createTemplateFromFile('initialise')
      .evaluate()
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


function include(filename) {
return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}