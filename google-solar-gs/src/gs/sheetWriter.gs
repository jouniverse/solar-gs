/**
 * Google Solar GS — Sheet Writer
 * Creates/overwrites named sheets with solar data
 */

/**
 * Gets or creates a sheet by name, clearing existing content.
 * @param {string} name - Sheet name
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (sheet) {
    sheet.clear();
    // Remove any existing formatting
    var range = sheet.getDataRange();
    if (range.getNumRows() > 0 && range.getNumColumns() > 0) {
      range.clearFormat();
    }
  } else {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

/**
 * Writes a 2D array to a sheet with formatted headers.
 * @param {string} sheetName - The sheet name
 * @param {Array[]} data - 2D array, first row is headers
 */
function writeDataToSheet_(sheetName, data) {
  if (!data || data.length === 0) return;

  var sheet = getOrCreateSheet_(sheetName);
  var numRows = data.length;
  var numCols = data[0].length;

  // Write all data
  sheet.getRange(1, 1, numRows, numCols).setValues(data);

  // Format header row: bold, frozen, light background
  var headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#e8f0fe");
  sheet.setFrozenRows(1);

  // Auto-resize columns
  for (var c = 1; c <= numCols; c++) {
    sheet.autoResizeColumn(c);
  }
}

/**
 * Writes individual panel data to the "Solar Panels" sheet.
 * Called from the sidebar via google.script.run.
 * @param {Object} buildingInsights
 * @param {number} [customCapacityWatts]
 * @return {Object} {success, message}
 */
function writeIndividualPanels(
  buildingInsights,
  customCapacityWatts,
  lat,
  lng,
) {
  try {
    var data = processIndividualPanels(buildingInsights, customCapacityWatts);
    appendLocationInfo_(data, lat, lng);
    writeDataToSheet_("Solar Panels", data);
    return {
      success: true,
      message:
        "Wrote " + (data.length - 1) + ' panels to "Solar Panels" sheet.',
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Writes panel configurations to the "Panel Configs" sheet.
 * @param {Object} buildingInsights
 * @param {number} [customCapacityWatts]
 * @return {Object} {success, message}
 */
function writePanelConfigs(buildingInsights, customCapacityWatts) {
  try {
    var data = processPanelConfigs(buildingInsights, customCapacityWatts);
    writeDataToSheet_("Panel Configs", data);
    return {
      success: true,
      message:
        "Wrote " +
        (data.length - 1) +
        ' configurations to "Panel Configs" sheet.',
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Writes roof segment data to the "Roof Segments" sheet.
 * @param {Object} buildingInsights
 * @return {Object} {success, message}
 */
function writeRoofSegments(buildingInsights, lat, lng) {
  try {
    var data = processRoofSegments(buildingInsights);
    appendLocationInfo_(data, lat, lng);
    writeDataToSheet_("Roof Segments", data);
    return {
      success: true,
      message:
        "Wrote " + (data.length - 1) + ' segments to "Roof Segments" sheet.',
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Writes solar summary to the "Solar Summary" sheet.
 * @param {Object} buildingInsights
 * @return {Object} {success, message}
 */
function writeSolarSummary(buildingInsights) {
  try {
    var data = processSolarSummary(buildingInsights);
    writeDataToSheet_("Solar Summary", data);
    return {
      success: true,
      message: 'Wrote solar summary to "Solar Summary" sheet.',
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Writes energy projection to the "Energy Projection" sheet.
 * @param {Object} buildingInsights
 * @param {number} panelCount
 * @param {number} [customCapacityWatts]
 * @param {number} [degradationRate]
 * @return {Object} {success, message}
 */
function writeEnergyProjection(
  buildingInsights,
  panelCount,
  customCapacityWatts,
  degradationRate,
) {
  try {
    var data = processEnergyProjection(
      buildingInsights,
      panelCount,
      customCapacityWatts,
      degradationRate,
    );
    writeDataToSheet_("Energy Projection", data);
    return {
      success: true,
      message:
        "Wrote " +
        (data.length - 1) +
        '-year projection to "Energy Projection" sheet.',
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Batch export: writes multiple datasets at once.
 * @param {Object} params - {buildingInsights, panelCount, customCapacityWatts, degradationRate, exports: string[]}
 * @return {Object} {success, results: Object[]}
 */
function exportSelectedData(params) {
  var results = [];
  var bi = params.buildingInsights;
  var exports = params.exports || [];

  if (exports.indexOf("panels") !== -1) {
    results.push({
      name: "Solar Panels",
      result: writeIndividualPanels(
        bi,
        params.customCapacityWatts,
        params.lat,
        params.lng,
      ),
    });
  }
  if (exports.indexOf("configs") !== -1) {
    results.push({
      name: "Panel Configs",
      result: writePanelConfigs(bi, params.customCapacityWatts),
    });
  }
  if (exports.indexOf("segments") !== -1) {
    results.push({
      name: "Roof Segments",
      result: writeRoofSegments(bi, params.lat, params.lng),
    });
  }
  if (exports.indexOf("summary") !== -1) {
    results.push({ name: "Solar Summary", result: writeSolarSummary(bi) });
  }
  if (exports.indexOf("projection") !== -1) {
    results.push({
      name: "Energy Projection",
      result: writeEnergyProjection(
        bi,
        params.panelCount,
        params.customCapacityWatts,
        params.degradationRate,
      ),
    });
  }

  var allSuccess = results.every(function (r) {
    return r.result.success;
  });
  return { success: allSuccess, results: results };
}

/**
 * Writes sunshine quantiles to the "Sunshine Quantiles" sheet.
 * @param {Object} buildingInsights
 * @return {Object} {success, message}
 */
function writeSunshineQuantiles(buildingInsights) {
  try {
    var data = processSunshineQuantiles(buildingInsights);
    writeDataToSheet_("Sunshine Quantiles", data);
    return {
      success: true,
      message:
        "Wrote " +
        (data.length - 1) +
        ' quantiles to "Sunshine Quantiles" sheet.',
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Writes cost analysis to the "Cost Analysis" sheet.
 * @param {Object} params - Cost parameters from client
 * @return {Object} {success, message}
 */
function writeCostAnalysis(params) {
  try {
    var data = processCostAnalysis(params);
    writeDataToSheet_("Cost Analysis", data);
    return {
      success: true,
      message: 'Wrote cost analysis to "Cost Analysis" sheet.',
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Appends location info (lat, lng, Google Maps URL) to the right of a 2D data array.
 * Only the header row and the first data row get values; remaining rows get empty cells.
 * @param {Array[]} data - 2D array (mutated in place)
 * @param {number} lat
 * @param {number} lng
 */
function appendLocationInfo_(data, lat, lng) {
  if (!data || data.length === 0 || lat == null || lng == null) return;

  var gmapsUrl = "https://www.google.com/maps?q=" + lat + "," + lng;

  // Add headers
  data[0].push("", "Location Lat", "Location Lng", "Google Maps URL");

  // First data row gets the values
  if (data.length > 1) {
    data[1].push("", lat, lng, gmapsUrl);
  }

  // Remaining rows get empty cells to keep the 2D array rectangular
  for (var i = 2; i < data.length; i++) {
    data[i].push("", "", "", "");
  }
}

/**
 * Writes arbitrary 2D data to a named sheet. Used by client inject buttons.
 * @param {string} sheetName
 * @param {Array[]} data - 2D array, first row is headers
 * @return {Object} {success, message}
 */
function writeCustomData(sheetName, data) {
  try {
    writeDataToSheet_(sheetName, data);
    return {
      success: true,
      message:
        "Wrote " + (data.length - 1) + ' rows to "' + sheetName + '" sheet.',
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}
