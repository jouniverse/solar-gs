var NREL_BASE_URL = "https://developer.nlr.gov";

var ENDPOINTS = {
  PVWATTS_V8: "/api/pvwatts/v8.json",
  SOLAR_RESOURCE: "/api/solar/solar_resource/v1.json",
  DATASET_QUERY_V2: "/api/solar/data_query/v2.json",
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("NREL Solar")
    .addItem("Open Sidebar", "openSidebar")
    .addSeparator()
    .addItem("Set API Key…", "promptApiKey")
    .addToUi();
}

function openSidebar() {
  var html = HtmlService.createTemplateFromFile("Sidebar")
    .evaluate()
    .setTitle("NREL Solar")
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getApiKey() {
  return (
    PropertiesService.getScriptProperties().getProperty("NREL_API_KEY") || ""
  );
}

function promptApiKey() {
  var ui = SpreadsheetApp.getUi();
  var current = getApiKey();
  var msg = current
    ? "Current key: " +
      current.substring(0, 8) +
      "…\nEnter a new key or cancel:"
    : "No API key set. Enter your NREL API key:";
  var result = ui.prompt("NREL API Key", msg, ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() === ui.Button.OK) {
    var key = result.getResponseText().trim();
    if (key) {
      PropertiesService.getScriptProperties().setProperty("NREL_API_KEY", key);
      ui.alert("API key saved.");
    }
  }
}

function saveApiKey(key) {
  if (key && key.trim()) {
    PropertiesService.getScriptProperties().setProperty(
      "NREL_API_KEY",
      key.trim(),
    );
    return true;
  }
  return false;
}
