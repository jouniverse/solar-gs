/**
 * Google Solar GS — Main entry point
 * Menu creation, sidebar launcher, HTML templating helper
 */

/**
 * Runs when the spreadsheet is opened. Adds the Solar API menu.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Solar API")
    .addItem("Open Sidebar", "openSidebar")
    .addSeparator()
    .addItem("Settings", "openSettings")
    .addItem("Help", "openHelp")
    .addSeparator()
    .addItem("Clear Cache", "clearSolarCache")
    .addToUi();
}

/**
 * Runs when the add-on is installed. Delegates to onOpen.
 */
function onInstall() {
  onOpen();
}

/**
 * Opens the Solar Panel sidebar.
 */
function openSidebar() {
  var template = HtmlService.createTemplateFromFile("sidebar");
  var html = template
    .evaluate()
    .setTitle("Solar System Designer")
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Opens the Help dialog.
 */
function openHelp() {
  var html = HtmlService.createHtmlOutputFromFile("help-dialog")
    .setWidth(500)
    .setHeight(480);
  SpreadsheetApp.getUi().showModalDialog(html, "Help — Solar System Designer");
}

/**
 * Opens the Settings dialog.
 */
function openSettings() {
  var html = HtmlService.createHtmlOutputFromFile("settings-dialog")
    .setWidth(420)
    .setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(
    html,
    "Settings — Solar System Designer",
  );
}

/**
 * Saves the API key to script properties.
 * @param {string} key - The Google Maps API key
 * @return {Object} {success, message|error}
 */
function saveApiKey(key) {
  if (!key || !key.trim()) {
    return { success: false, error: "API key cannot be empty." };
  }
  PropertiesService.getScriptProperties().setProperty(
    "GMAPS_API_KEY",
    key.trim(),
  );
  return { success: true, message: "API key saved successfully." };
}

/**
 * Returns the stored API key (masked for display).
 * @return {Object} {hasKey, masked}
 */
function getStoredApiKey() {
  var key =
    PropertiesService.getScriptProperties().getProperty("GMAPS_API_KEY");
  if (!key) return { hasKey: false, masked: "" };
  var masked = key.substring(0, 8) + "..." + key.substring(key.length - 4);
  return { hasKey: true, masked: masked };
}

/**
 * Includes an HTML file's content for templating.
 * Usage in HTML: <?!= include('css') ?>
 * @param {string} filename - The HTML file name (without extension)
 * @return {string} The file content
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
