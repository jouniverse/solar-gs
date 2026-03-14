/**
 * PVGIS API interaction layer.
 * All API calls go through fetchPvgisData() which uses UrlFetchApp.
 */

/**
 * Fetches data from a PVGIS API endpoint.
 * @param {string} endpoint - The API endpoint name (e.g. 'PVcalc', 'tmy').
 * @param {object} params - Query parameters as key-value pairs.
 * @return {object} Parsed JSON response from PVGIS.
 * @throws {Error} If the API returns an error or the request fails.
 */
function fetchPvgisData(endpoint, params) {
  params.outputformat = "json";

  var queryParts = [];
  for (var key in params) {
    if (
      params[key] !== null &&
      params[key] !== undefined &&
      params[key] !== ""
    ) {
      queryParts.push(
        encodeURIComponent(key) + "=" + encodeURIComponent(params[key]),
      );
    }
  }

  var url = PVGIS_CONFIG.BASE_URL + "/" + endpoint + "?" + queryParts.join("&");

  var options = {
    method: "get",
    muteHttpExceptions: true,
  };

  var response = UrlFetchApp.fetch(url, options);
  var code = response.getResponseCode();
  var body = response.getContentText();

  if (code !== 200) {
    var errorMsg = "PVGIS API error (HTTP " + code + ")";
    try {
      var errData = JSON.parse(body);
      if (errData.message) {
        errorMsg = errData.message;
      }
    } catch (e) {
      if (body) {
        errorMsg += ": " + body.substring(0, 200);
      }
    }
    throw new Error(errorMsg);
  }

  var data = JSON.parse(body);
  return data;
}

// ─── Endpoint-specific wrappers ─────────────────────────────────────────────

/**
 * Fetches grid-connected PV system data and writes to sheet.
 * @param {object} params - Form parameters from the sidebar.
 * @return {object} Summary info for the sidebar.
 */
function fetchGridConnected(params) {
  validateRequired(params, ["lat", "lon", "peakpower", "loss"]);
  var data = fetchPvgisData(PVGIS_CONFIG.ENDPOINTS.GRID_CONNECTED, params);
  writeGridConnectedData(data);
  var totals = data.outputs.totals.fixed;
  return {
    message: "Grid-connected PV data written to sheet.",
    summary: "Annual production: " + totals.E_y + " kWh/y",
  };
}

/**
 * Fetches tracking PV system data and writes to sheet.
 * @param {object} params - Form parameters from the sidebar.
 * @return {object} Summary info for the sidebar.
 */
function fetchTrackingPV(params) {
  validateRequired(params, ["lat", "lon", "peakpower", "loss"]);
  var data = fetchPvgisData(PVGIS_CONFIG.ENDPOINTS.TRACKING_PV, params);
  writeTrackingPVData(data);
  var summary = "";
  var outputs = data.outputs.totals;
  if (outputs.vertical_axis)
    summary += "Vertical axis: " + outputs.vertical_axis.E_y + " kWh/y  ";
  if (outputs.inclined_axis)
    summary += "Inclined axis: " + outputs.inclined_axis.E_y + " kWh/y  ";
  if (outputs.twoaxis)
    summary += "Two-axis: " + outputs.twoaxis.E_y + " kWh/y  ";
  return {
    message: "Tracking PV data written to sheet.",
    summary: summary || "Data written.",
  };
}

/**
 * Fetches off-grid PV system data and writes to sheet.
 * @param {object} params - Form parameters from the sidebar.
 * @return {object} Summary info for the sidebar.
 */
function fetchOffGrid(params) {
  validateRequired(params, [
    "lat",
    "lon",
    "peakpower",
    "batterysize",
    "cutoff",
    "consumptionday",
  ]);
  var data = fetchPvgisData(PVGIS_CONFIG.ENDPOINTS.OFF_GRID, params);
  writeOffGridData(data);
  return {
    message: "Off-grid PV data written to sheet.",
    summary: "12-month performance data written.",
  };
}

/**
 * Fetches monthly radiation data and writes to sheet.
 * @param {object} params - Form parameters from the sidebar.
 * @return {object} Summary info for the sidebar.
 */
function fetchMonthlyRadiation(params) {
  validateRequired(params, ["lat", "lon"]);
  var data = fetchPvgisData(PVGIS_CONFIG.ENDPOINTS.MONTHLY_RADIATION, params);
  writeMonthlyRadiationData(data);
  var count = data.outputs.monthly ? data.outputs.monthly.length : 0;
  return {
    message: "Monthly radiation data written to sheet.",
    summary: count + " records written.",
  };
}

/**
 * Fetches daily radiation data and writes to sheet.
 * @param {object} params - Form parameters from the sidebar.
 * @return {object} Summary info for the sidebar.
 */
function fetchDailyRadiation(params) {
  validateRequired(params, ["lat", "lon", "month"]);
  var data = fetchPvgisData(PVGIS_CONFIG.ENDPOINTS.DAILY_RADIATION, params);
  writeDailyRadiationData(data);
  var count = data.outputs.daily_profile
    ? data.outputs.daily_profile.length
    : 0;
  return {
    message: "Daily radiation data written to sheet.",
    summary: count + " hourly records written.",
  };
}

/**
 * Fetches hourly radiation data and writes to sheet.
 * @param {object} params - Form parameters from the sidebar.
 * @return {object} Summary info for the sidebar.
 */
function fetchHourlyRadiation(params) {
  validateRequired(params, ["lat", "lon"]);
  var data = fetchPvgisData(PVGIS_CONFIG.ENDPOINTS.HOURLY_RADIATION, params);
  writeHourlyRadiationData(data);
  var count = data.outputs.hourly ? data.outputs.hourly.length : 0;
  return {
    message: "Hourly radiation data written to sheet.",
    summary: count + " hourly records written.",
  };
}

/**
 * Fetches TMY data and writes to sheet.
 * @param {object} params - Form parameters from the sidebar.
 * @return {object} Summary info for the sidebar.
 */
function fetchTMY(params) {
  validateRequired(params, ["lat", "lon"]);
  var data = fetchPvgisData(PVGIS_CONFIG.ENDPOINTS.TMY, params);
  writeTMYData(data);
  var count = data.outputs.tmy_hourly ? data.outputs.tmy_hourly.length : 0;
  return {
    message: "TMY data written to sheet.",
    summary: count + " hourly records written.",
  };
}

/**
 * Fetches horizon profile data and writes to sheet.
 * @param {object} params - Form parameters from the sidebar.
 * @return {object} Summary info for the sidebar.
 */
function fetchHorizonProfile(params) {
  validateRequired(params, ["lat", "lon"]);
  var data = fetchPvgisData(PVGIS_CONFIG.ENDPOINTS.HORIZON_PROFILE, params);
  writeHorizonProfileData(data);
  return {
    message: "Horizon profile data written to sheet.",
    summary: "Horizon and sun path data written.",
  };
}

// ─── Validation ─────────────────────────────────────────────────────────────

/**
 * Validates that required parameters are present and non-empty.
 * @param {object} params - Parameters object.
 * @param {string[]} required - Array of required parameter names.
 * @throws {Error} If any required parameter is missing.
 */
function validateRequired(params, required) {
  for (var i = 0; i < required.length; i++) {
    var key = required[i];
    if (
      params[key] === null ||
      params[key] === undefined ||
      params[key] === ""
    ) {
      throw new Error("Missing required parameter: " + key);
    }
  }
}
