/**
 * PVGIS Google Sheets Add-on
 * Fetches solar energy data from PVGIS (Photovoltaic Geographical Information System)
 * directly into Google Sheets for analysis.
 */

/**
 * Creates the add-on menu when the spreadsheet opens.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("PVGIS")
    .addItem("Open Sidebar", "showSidebar")
    .addSeparator()
    .addItem("Clear Active Sheet", "clearActiveSheet")
    .addToUi();
}

/**
 * Shows the PVGIS sidebar.
 */
function showSidebar() {
  var html = HtmlService.createTemplateFromFile("Sidebar")
    .evaluate()
    .setTitle("PVGIS Data")
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Includes an HTML file as a template partial.
 * Used in HTML templates via: <?!= include('FileName') ?>
 * @param {string} filename - The name of the HTML file to include (without .html extension).
 * @return {string} The file content as a string.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Clears all content and formatting from the active sheet.
 */
function clearActiveSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();
  sheet.clearFormats();
  SpreadsheetApp.getActiveSpreadsheet().toast("Sheet cleared.", "PVGIS");
}

/**
 * Saves the user's last-used location to PropertiesService.
 * @param {number} lat - Latitude.
 * @param {number} lon - Longitude.
 */
function saveLocation(lat, lon) {
  var props = PropertiesService.getUserProperties();
  props.setProperty("pvgis_lat", String(lat));
  props.setProperty("pvgis_lon", String(lon));
}

/**
 * Loads the user's last-used location from PropertiesService.
 * @return {object} Object with lat and lon, or defaults.
 */
function loadLocation() {
  var props = PropertiesService.getUserProperties();
  var lat = props.getProperty("pvgis_lat");
  var lon = props.getProperty("pvgis_lon");
  return {
    lat: lat ? parseFloat(lat) : 45.0,
    lon: lon ? parseFloat(lon) : 8.0,
  };
}
