/**
 * SheetWriter — writes PVGIS API response data to Google Sheets.
 * Each writer creates/replaces a named sheet and uses batch setValues().
 */

// ─── Helpers ────────────────────────────────────────────────────────────────

/**
 * Gets or creates a sheet with the given name. Clears it if it exists.
 * @param {string} name - Sheet name.
 * @return {Sheet} The sheet object.
 */
function getOrCreateSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (sheet) {
    sheet.clear();
    sheet.clearFormats();
  } else {
    sheet = ss.insertSheet(name);
  }
  ss.setActiveSheet(sheet);
  return sheet;
}

/**
 * Writes a metadata header block to the sheet, including GMAPS URL.
 * @param {Sheet} sheet - The target sheet.
 * @param {object} data - The full API response (contains inputs).
 * @param {string} title - Title for the header.
 * @return {number} The next row after the header.
 */
function writeMetaHeader(sheet, data, title) {
  var row = 1;
  var inputs = data.inputs || {};
  var loc = inputs.location || {};

  sheet.getRange(row, 1).setValue(title).setFontWeight("bold").setFontSize(11);
  row++;
  sheet.getRange(row, 1).setValue("Latitude");
  sheet.getRange(row, 2).setValue(loc.latitude || "");
  row++;
  sheet.getRange(row, 1).setValue("Longitude");
  sheet.getRange(row, 2).setValue(loc.longitude || "");
  row++;

  // Google Maps URL
  var lat = loc.latitude || "";
  var lon = loc.longitude || "";
  if (lat !== "" && lon !== "") {
    sheet.getRange(row, 1).setValue("GMAPS URL");
    sheet
      .getRange(row, 2)
      .setValue("https://www.google.com/maps?q=" + lat + "," + lon);
  }
  row++;

  sheet.getRange(row, 1).setValue("Elevation (m)");
  sheet.getRange(row, 2).setValue(loc.elevation || "");
  row++;

  var meteo = inputs.meteo_data || {};
  if (meteo.radiation_db) {
    sheet.getRange(row, 1).setValue("Radiation DB");
    sheet.getRange(row, 2).setValue(meteo.radiation_db);
    row++;
  }

  sheet.getRange(row, 1).setValue("Fetched");
  sheet.getRange(row, 2).setValue(new Date().toISOString());
  row++;
  row++; // blank row
  return row;
}

/**
 * Writes a 2D array to the sheet starting at the given row, with bold header row.
 * @param {Sheet} sheet - The target sheet.
 * @param {number} startRow - 1-based row to start writing.
 * @param {string[]} headers - Column headers.
 * @param {Array[]} rows - 2D array of data rows.
 */
function writeTable(sheet, startRow, headers, rows) {
  if (headers.length > 0) {
    var headerRange = sheet.getRange(startRow, 1, 1, headers.length);
    headerRange.setValues([headers]).setFontWeight("bold");
  }
  if (rows.length > 0) {
    var dataRange = sheet.getRange(
      startRow + 1,
      1,
      rows.length,
      headers.length,
    );
    dataRange.setValues(rows);
  }
  // Auto-resize columns
  for (var c = 1; c <= headers.length; c++) {
    sheet.autoResizeColumn(c);
  }
}

/**
 * Converts PVGIS time format "YYYYMMDD:HHMM" or "YYYYMMDD:HH:MM" to "MM/DD/YYYY HH:MM:SS".
 * @param {string} pvgisTime - Time string from PVGIS API.
 * @return {string} Formatted time string for Google Sheets.
 */
function formatPvgisTime(pvgisTime) {
  if (!pvgisTime || typeof pvgisTime !== "string") return pvgisTime;
  // Remove any trailing ".000" or similar
  var cleaned = pvgisTime.replace(/\.000$/, "");
  // Match patterns: "YYYYMMDD:HHMM" or "YYYYMMDD:HH:MM"
  var match = cleaned.match(/^(\d{4})(\d{2})(\d{2}):(\d{2}):?(\d{2})$/);
  if (match) {
    var year = match[1];
    var month = match[2];
    var day = match[3];
    var hour = match[4];
    var minute = match[5];
    return month + "/" + day + "/" + year + " " + hour + ":" + minute + ":00";
  }
  return pvgisTime;
}

// ─── Writers ────────────────────────────────────────────────────────────────

/**
 * Writes grid-connected PV data to a sheet.
 */
function writeGridConnectedData(data) {
  var sheet = getOrCreateSheet(PVGIS_CONFIG.SHEET_NAMES.GRID_CONNECTED);
  var row = writeMetaHeader(sheet, data, "PERFORMANCE OF GRID-CONNECTED PV");

  // PV module info
  var pv = data.inputs.pv_module || {};
  sheet.getRange(row, 1).setValue("PV Technology");
  sheet.getRange(row, 2).setValue(pv.technology || "");
  row++;
  sheet.getRange(row, 1).setValue("Peak Power (kWp)");
  sheet.getRange(row, 2).setValue(pv.peak_power || "");
  row++;
  sheet.getRange(row, 1).setValue("System Loss (%)");
  sheet.getRange(row, 2).setValue(pv.system_loss || "");
  row++;
  row++;

  // Monthly data
  var monthly = data.outputs.monthly.fixed || [];
  var headers = [
    "Month",
    "E_d (kWh/d)",
    "E_m (kWh/mo)",
    "H(i)_d (kWh/m²/d)",
    "H(i)_m (kWh/m²/mo)",
    "SD_m (kWh)",
  ];
  var rows = monthly.map(function (m) {
    return [m.month, m.E_d, m.E_m, m["H(i)_d"], m["H(i)_m"], m.SD_m];
  });

  writeTable(sheet, row, headers, rows);
  row += rows.length + 2;

  // Totals
  var totals = data.outputs.totals.fixed || {};
  sheet.getRange(row, 1).setValue("Totals").setFontWeight("bold");
  row++;
  var totalKeys = [
    ["E_d (kWh/d)", "E_d"],
    ["E_m (kWh/mo)", "E_m"],
    ["E_y (kWh/y)", "E_y"],
    ["H(i)_d (kWh/m²/d)", "H(i)_d"],
    ["H(i)_m (kWh/m²/mo)", "H(i)_m"],
    ["H(i)_y (kWh/m²/y)", "H(i)_y"],
    ["SD_m (kWh)", "SD_m"],
    ["SD_y (kWh)", "SD_y"],
    ["Loss AOI (%)", "l_aoi"],
    ["Loss spectral (%)", "l_spec"],
    ["Loss temp+irrad (%)", "l_tg"],
    ["Loss total (%)", "l_total"],
    ["LCOE_pv (currency/kWh)", "LCOE_pv"],
  ];
  totalKeys.forEach(function (pair) {
    var val = totals[pair[1]];
    if (val !== undefined && val !== null) {
      sheet.getRange(row, 1).setValue(pair[0]);
      sheet.getRange(row, 2).setValue(val);
      row++;
    }
  });
}

/**
 * Writes tracking PV data to a sheet.
 * Only writes the tracking types that were requested (no fixed output).
 */
function writeTrackingPVData(data) {
  var sheet = getOrCreateSheet(PVGIS_CONFIG.SHEET_NAMES.TRACKING_PV);
  var row = writeMetaHeader(sheet, data, "PERFORMANCE OF TRACKING PV");

  var pv = data.inputs.pv_module || {};
  sheet.getRange(row, 1).setValue("PV Technology");
  sheet.getRange(row, 2).setValue(pv.technology || "");
  row++;
  sheet.getRange(row, 1).setValue("Peak Power (kWp)");
  sheet.getRange(row, 2).setValue(pv.peak_power || "");
  row++;
  sheet.getRange(row, 1).setValue("System Loss (%)");
  sheet.getRange(row, 2).setValue(pv.system_loss || "");
  row++;
  row++;

  // Only write tracking types (not fixed)
  var trackingTypes = ["vertical_axis", "inclined_axis", "two_axis"];
  trackingTypes.forEach(function (type) {
    var monthly = data.outputs.monthly[type];
    if (!monthly) return;

    var label = type.replace(/_/g, " ").replace(/\b\w/g, function (c) {
      return c.toUpperCase();
    });
    sheet
      .getRange(row, 1)
      .setValue(label + " - Monthly")
      .setFontWeight("bold");
    row++;

    var headers = [
      "Month",
      "E_d (kWh/d)",
      "E_m (kWh/mo)",
      "H(i)_d (kWh/m²/d)",
      "H(i)_m (kWh/m²/mo)",
      "SD_m (kWh)",
    ];
    var rows = monthly.map(function (m) {
      return [m.month, m.E_d, m.E_m, m["H(i)_d"], m["H(i)_m"], m.SD_m];
    });
    writeTable(sheet, row, headers, rows);
    row += rows.length + 2;

    // Full totals for this tracking type
    var totals = data.outputs.totals[type] || {};
    sheet
      .getRange(row, 1)
      .setValue(label + " - Totals")
      .setFontWeight("bold");
    row++;
    var totalKeys = [
      ["E_d (kWh/d)", "E_d"],
      ["E_m (kWh/mo)", "E_m"],
      ["E_y (kWh/y)", "E_y"],
      ["H(i)_d (kWh/m²/d)", "H(i)_d"],
      ["H(i)_m (kWh/m²/mo)", "H(i)_m"],
      ["H(i)_y (kWh/m²/y)", "H(i)_y"],
      ["SD_m (kWh)", "SD_m"],
      ["SD_y (kWh)", "SD_y"],
      ["Loss AOI (%)", "l_aoi"],
      ["Loss spectral (%)", "l_spec"],
      ["Loss temp+irrad (%)", "l_tg"],
      ["Loss total (%)", "l_total"],
    ];
    totalKeys.forEach(function (pair) {
      var val = totals[pair[1]];
      if (val !== undefined && val !== null) {
        sheet.getRange(row, 1).setValue(pair[0]);
        sheet.getRange(row, 2).setValue(val);
        row++;
      }
    });
    row++;
  });
}

/**
 * Writes off-grid PV data to a sheet, including totals and histogram.
 */
function writeOffGridData(data) {
  var sheet = getOrCreateSheet(PVGIS_CONFIG.SHEET_NAMES.OFF_GRID);
  var row = writeMetaHeader(sheet, data, "PERFORMANCE OF OFF-GRID PV SYSTEM");

  var pv = data.inputs.pv_module || {};
  var batt = data.inputs.battery || {};
  var cons = data.inputs.consumption || {};

  sheet.getRange(row, 1).setValue("Peak Power (Wp)");
  sheet.getRange(row, 2).setValue(pv.peak_power || "");
  row++;
  sheet.getRange(row, 1).setValue("Battery Capacity (Wh)");
  sheet.getRange(row, 2).setValue(batt.capacity || "");
  row++;
  sheet.getRange(row, 1).setValue("Discharge Cutoff (%)");
  sheet.getRange(row, 2).setValue(batt.discharge_cutoff_limit || "");
  row++;
  sheet.getRange(row, 1).setValue("Daily Consumption (Wh)");
  sheet.getRange(row, 2).setValue(cons.daily || "");
  row++;
  row++;

  var monthly = data.outputs.monthly || [];
  var headers = [
    "Month",
    "E_d (Wh/d)",
    "E_lost_d (Wh/d)",
    "f_f (%)",
    "f_e (%)",
  ];
  var rows = monthly.map(function (m) {
    return [m.month, m.E_d, m.E_lost_d, m.f_f, m.f_e];
  });

  writeTable(sheet, row, headers, rows);
  row += rows.length + 2;

  // Totals — write all fields present in the response
  if (data.outputs.totals) {
    var t = data.outputs.totals;
    sheet.getRange(row, 1).setValue("Totals").setFontWeight("bold");
    row++;
    var totalKeys = [
      ["d_total (days)", "d_total"],
      ["E_d (Wh/d)", "E_d"],
      ["E_lost_d (Wh/d)", "E_lost_d"],
      ["E_lost (Wh)", "E_lost"],
      ["E_miss (Wh)", "E_miss"],
      ["f_f (%)", "f_f"],
      ["f_e (%)", "f_e"],
    ];
    totalKeys.forEach(function (pair) {
      var val = t[pair[1]];
      if (val !== undefined && val !== null) {
        sheet.getRange(row, 1).setValue(pair[0]);
        sheet.getRange(row, 2).setValue(val);
        row++;
      }
    });
    row++;
  }

  // Histogram — battery charge state distribution
  var histogram = data.outputs.histogram || [];
  if (histogram.length > 0) {
    sheet
      .getRange(row, 1)
      .setValue("Battery Charge State Histogram")
      .setFontWeight("bold");
    row++;
    var histHeaders = ["CS_min (%)", "CS_max (%)", "f_CS (days)"];
    var histRows = histogram.map(function (h) {
      return [h.CS_min, h.CS_max, h.f_CS];
    });
    writeTable(sheet, row, histHeaders, histRows);
    row += histRows.length + 2;
  }
}

/**
 * Writes monthly radiation data to a sheet.
 */
function writeMonthlyRadiationData(data) {
  var sheet = getOrCreateSheet(PVGIS_CONFIG.SHEET_NAMES.MONTHLY_RADIATION);
  var row = writeMetaHeader(sheet, data, "MONTHLY IRRADIATION DATA");

  var monthly = data.outputs.monthly || [];
  if (monthly.length === 0) return;

  // Determine which columns are present from the first record
  var first = monthly[0];
  var colDefs = [];
  colDefs.push({ key: "year", label: "Year" });
  colDefs.push({ key: "month", label: "Month" });
  if (first["H(h)_m"] !== undefined)
    colDefs.push({ key: "H(h)_m", label: "H(h)_m (kWh/m²/mo)" });
  if (first["H(i_opt)_m"] !== undefined)
    colDefs.push({ key: "H(i_opt)_m", label: "H(i_opt)_m (kWh/m²/mo)" });
  if (first["H(i)_m"] !== undefined)
    colDefs.push({ key: "H(i)_m", label: "H(i)_m (kWh/m²/mo)" });
  if (first["Hb(n)_m"] !== undefined)
    colDefs.push({ key: "Hb(n)_m", label: "Hb(n)_m (kWh/m²/mo)" });
  if (first.Kd !== undefined) colDefs.push({ key: "Kd", label: "Kd (ratio)" });
  if (first.T2m !== undefined) colDefs.push({ key: "T2m", label: "T2m (°C)" });

  var headers = colDefs.map(function (c) {
    return c.label;
  });
  var rows = monthly.map(function (m) {
    return colDefs.map(function (c) {
      return m[c.key] !== undefined ? m[c.key] : "";
    });
  });

  writeTable(sheet, row, headers, rows);
}

/**
 * Writes daily radiation data to a sheet.
 * Dynamically detects all columns present in the response.
 */
function writeDailyRadiationData(data) {
  var sheet = getOrCreateSheet(PVGIS_CONFIG.SHEET_NAMES.DAILY_RADIATION);
  var row = writeMetaHeader(sheet, data, "AVERAGE DAILY IRRADIANCE DATA");

  var daily = data.outputs.daily_profile || [];
  if (daily.length === 0) return;

  // All possible daily radiation columns in expected order
  var allPossibleKeys = [
    { key: "month", label: "Month" },
    { key: "time", label: "Time" },
    { key: "G(i)", label: "G(i) (W/m²)" },
    { key: "Gb(i)", label: "Gb(i) (W/m²)" },
    { key: "Gd(i)", label: "Gd(i) (W/m²)" },
    { key: "Gcs(i)", label: "Gcs(i) (W/m²)" },
    { key: "G(n)", label: "G(n) (W/m²)" },
    { key: "Gb(n)", label: "Gb(n) (W/m²)" },
    { key: "Gd(n)", label: "Gd(n) (W/m²)" },
    { key: "Gcs(n)", label: "Gcs(n) (W/m²)" },
    { key: "T2m", label: "T2m (°C)" },
  ];

  var first = daily[0];
  var colDefs = allPossibleKeys.filter(function (c) {
    return first[c.key] !== undefined;
  });

  var headers = colDefs.map(function (c) {
    return c.label;
  });
  var rows = daily.map(function (d) {
    return colDefs.map(function (c) {
      return d[c.key] !== undefined ? d[c.key] : "";
    });
  });

  writeTable(sheet, row, headers, rows);
}

/**
 * Writes hourly radiation data to a sheet. Can be very large (8760+ rows per year).
 * Converts PVGIS time format to Google Sheets-parsable format.
 */
function writeHourlyRadiationData(data) {
  var sheet = getOrCreateSheet(PVGIS_CONFIG.SHEET_NAMES.HOURLY_RADIATION);
  var row = writeMetaHeader(sheet, data, "HOURLY RADIATION DATA");

  var hourly = data.outputs.hourly || [];
  if (hourly.length === 0) return;

  // All possible hourly columns in expected order
  var allPossibleKeys = [
    { key: "time", label: "Time" },
    { key: "P", label: "P (W)" },
    { key: "G(i)", label: "G(i) (W/m²)" },
    { key: "Gb(i)", label: "Gb(i) (W/m²)" },
    { key: "Gd(i)", label: "Gd(i) (W/m²)" },
    { key: "Gr(i)", label: "Gr(i) (W/m²)" },
    { key: "H_sun", label: "Sun Height (°)" },
    { key: "T2m", label: "T2m (°C)" },
    { key: "WS10m", label: "WS10m (m/s)" },
    { key: "Int", label: "Int" },
  ];

  var first = hourly[0];
  var colDefs = allPossibleKeys.filter(function (c) {
    return first[c.key] !== undefined;
  });

  var headers = colDefs.map(function (c) {
    return c.label;
  });

  // Batch write for performance
  var headerRange = sheet.getRange(row, 1, 1, headers.length);
  headerRange.setValues([headers]).setFontWeight("bold");

  var BATCH_SIZE = 10000;
  var dataRow = row + 1;
  for (var i = 0; i < hourly.length; i += BATCH_SIZE) {
    var batch = hourly.slice(i, i + BATCH_SIZE);
    var batchRows = batch.map(function (h) {
      return colDefs.map(function (c) {
        var val = h[c.key] !== undefined ? h[c.key] : "";
        if (c.key === "time") val = formatPvgisTime(val);
        return val;
      });
    });
    sheet
      .getRange(dataRow, 1, batchRows.length, headers.length)
      .setValues(batchRows);
    dataRow += batchRows.length;
  }

  for (var c = 1; c <= headers.length; c++) {
    sheet.autoResizeColumn(c);
  }
}

/**
 * Writes TMY data to a sheet. Includes months_selected summary + ~8760 hourly rows.
 * Converts PVGIS time format to Google Sheets-parsable format.
 */
function writeTMYData(data) {
  var sheet = getOrCreateSheet(PVGIS_CONFIG.SHEET_NAMES.TMY);
  var row = writeMetaHeader(sheet, data, "TYPICAL METEOROLOGICAL YEAR");

  // Months selected
  var months = data.outputs.months_selected || [];
  if (months.length > 0) {
    sheet.getRange(row, 1).setValue("Months Selected").setFontWeight("bold");
    row++;
    var mHeaders = ["Month", "Year"];
    var mRows = months.map(function (m) {
      return [m.month, m.year];
    });
    writeTable(sheet, row, mHeaders, mRows);
    row += mRows.length + 2;
  }

  // TMY hourly data
  var hourly = data.outputs.tmy_hourly || [];
  if (hourly.length === 0) return;

  sheet.getRange(row, 1).setValue("Hourly Data").setFontWeight("bold");
  row++;

  var headers = [
    "Time (UTC)",
    "T2m (°C)",
    "RH (%)",
    "G(h) (W/m²)",
    "Gb(n) (W/m²)",
    "Gd(h) (W/m²)",
    "IR(h) (W/m²)",
    "WS10m (m/s)",
    "WD10m (°)",
    "SP (Pa)",
  ];
  var keys = [
    "time(UTC)",
    "T2m",
    "RH",
    "G(h)",
    "Gb(n)",
    "Gd(h)",
    "IR(h)",
    "WS10m",
    "WD10m",
    "SP",
  ];

  var headerRange = sheet.getRange(row, 1, 1, headers.length);
  headerRange.setValues([headers]).setFontWeight("bold");

  var BATCH_SIZE = 10000;
  var dataRow = row + 1;
  for (var i = 0; i < hourly.length; i += BATCH_SIZE) {
    var batch = hourly.slice(i, i + BATCH_SIZE);
    var batchRows = batch.map(function (h) {
      return keys.map(function (k) {
        var val = h[k] !== undefined ? h[k] : "";
        if (k === "time(UTC)") val = formatPvgisTime(val);
        return val;
      });
    });
    sheet
      .getRange(dataRow, 1, batchRows.length, headers.length)
      .setValues(batchRows);
    dataRow += batchRows.length;
  }

  for (var c = 1; c <= headers.length; c++) {
    sheet.autoResizeColumn(c);
  }
}

/**
 * Writes horizon profile data to a sheet.
 * Includes horizon profile, winter solstice and summer solstice sun paths.
 * Each dataset has its own azimuth column with different key names.
 */
function writeHorizonProfileData(data) {
  var sheet = getOrCreateSheet(PVGIS_CONFIG.SHEET_NAMES.HORIZON_PROFILE);
  var row = writeMetaHeader(sheet, data, "HORIZON PROFILE");

  // Horizon profile heights
  var horizon = data.outputs.horizon_profile || [];
  if (horizon.length > 0) {
    sheet.getRange(row, 1).setValue("Horizon Profile").setFontWeight("bold");
    row++;
    var hHeaders = ["Azimuth (°)", "Horizon Height (°)"];
    var hRows = horizon.map(function (h) {
      return [h.A, h.H_hor];
    });
    writeTable(sheet, row, hHeaders, hRows);
    row += hRows.length + 2;
  }

  // Winter solstice sun path
  var winter = data.outputs.winter_solstice || [];
  if (winter.length > 0) {
    sheet
      .getRange(row, 1)
      .setValue("Winter Solstice Sun Path (Dec 21)")
      .setFontWeight("bold");
    row++;
    var wHeaders = ["A_sun(w) - Azimuth (°)", "H_sun(w) - Sun Elevation (°)"];
    var wRows = winter.map(function (s) {
      return [s["A_sun(w)"], s["H_sun(w)"]];
    });
    writeTable(sheet, row, wHeaders, wRows);
    row += wRows.length + 2;
  }

  // Summer solstice sun path
  var summer = data.outputs.summer_solstice || [];
  if (summer.length > 0) {
    sheet
      .getRange(row, 1)
      .setValue("Summer Solstice Sun Path (Jun 21)")
      .setFontWeight("bold");
    row++;
    var sHeaders = ["A_sun(s) - Azimuth (°)", "H_sun(s) - Sun Elevation (°)"];
    var sRows = summer.map(function (s) {
      return [s["A_sun(s)"], s["H_sun(s)"]];
    });
    writeTable(sheet, row, sHeaders, sRows);
    row += sRows.length + 2;
  }
}
