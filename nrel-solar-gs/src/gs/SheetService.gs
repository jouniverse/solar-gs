var MONTH_NAMES = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December"
];

var MONTH_KEYS = [
  "jan", "feb", "mar", "apr", "may", "jun",
  "jul", "aug", "sep", "oct", "nov", "dec"
];

/**
 * Returns (and optionally creates) a sheet by name, then clears it.
 */
function getOrCreateSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  } else {
    sheet.clear();
  }
  return sheet;
}

// ---------------------------------------------------------------------------
// PVWatts V8
// ---------------------------------------------------------------------------

function writePVWattsToSheet(data) {
  var sheet = getOrCreateSheet_("PVWatts V8");
  var out = data.outputs;
  var info = data.station_info || {};

  var rows = [];

  rows.push(["PVWatts V8 Results"]);
  rows.push(["Station", info.city || "", info.state || "", "Lat: " + (info.lat || ""), "Lon: " + (info.lon || ""), "Elev: " + (info.elev || "") + " m"]);
  rows.push(["Weather Source", info.weather_data_source || ""]);
  rows.push(["Station Map", gmapsUrl_(info.lat, info.lon)]);
  rows.push([]);

  rows.push(["Month", "AC Output (kWh)", "POA Irradiance (kWh/m\u00B2)", "Solar Radiation (kWh/m\u00B2/day)", "DC Output (kWh)"]);

  for (var i = 0; i < 12; i++) {
    rows.push([
      MONTH_NAMES[i],
      round_(out.ac_monthly[i]),
      round_(out.poa_monthly[i]),
      round_(out.solrad_monthly[i]),
      round_(out.dc_monthly[i])
    ]);
  }

  rows.push([
    "Annual",
    round_(out.ac_annual),
    "",
    round_(out.solrad_annual),
    ""
  ]);

  rows.push([]);
  rows.push(["Capacity Factor (%)", round_(out.capacity_factor)]);

  sheet.getRange(1, 1, rows.length, rows[0].length < 6 ? 6 : rows[0].length).setValues(
    rows.map(function(r) {
      while (r.length < 6) r.push("");
      return r;
    })
  );

  formatHeader_(sheet, 1, 6);
  formatHeader_(sheet, 6, 5);
  sheet.autoResizeColumns(1, 6);
}

// ---------------------------------------------------------------------------
// Solar Resource Data
// ---------------------------------------------------------------------------

function writeSolarResourceToSheet(data) {
  var sheet = getOrCreateSheet_("Solar Resource");
  var out = data.outputs;
  var inp = data.inputs || {};

  var rows = [];

  rows.push(["Solar Resource Data"]);
  rows.push(["Location", "Lat: " + (inp.lat || ""), "Lon: " + (inp.lon || ""), gmapsUrl_(inp.lat, inp.lon)]);
  rows.push(["Source", (data.metadata && data.metadata.sources) ? data.metadata.sources.join(", ") : ""]);
  rows.push([]);

  rows.push(["Month", "Avg DNI (kWh/m\u00B2/day)", "Avg GHI (kWh/m\u00B2/day)", "Avg Tilt at Lat (kWh/m\u00B2/day)"]);

  for (var i = 0; i < 12; i++) {
    var key = MONTH_KEYS[i];
    rows.push([
      MONTH_NAMES[i],
      out.avg_dni.monthly[key],
      out.avg_ghi.monthly[key],
      out.avg_lat_tilt.monthly[key]
    ]);
  }

  rows.push(["Annual", out.avg_dni.annual, out.avg_ghi.annual, out.avg_lat_tilt.annual]);

  sheet.getRange(1, 1, rows.length, 4).setValues(
    rows.map(function(r) {
      while (r.length < 4) r.push("");
      return r;
    })
  );

  formatHeader_(sheet, 1, 4);
  formatHeader_(sheet, 5, 4);
  sheet.autoResizeColumns(1, 4);
}

// ---------------------------------------------------------------------------
// Solar Dataset Query V2
// ---------------------------------------------------------------------------

var DATASET_DESCRIPTIONS = {
  nsrdb: "Gridded TMY data from the NREL National Solar Radiation Database (NSRDB)",
  tmy2: "TMY2 station data (NSRDB 1961-1990 Archive)",
  tmy3: "TMY3 station data (NSRDB 1991-2005 Archive)",
  intl: "PVWatts International station data"
};

function writeDatasetQueryToSheet(data) {
  var sheet = getOrCreateSheet_("Dataset Query");
  var out = data.outputs;
  var inp = data.inputs || {};

  var rows = [];

  rows.push(["Solar Dataset Query V2"]);
  rows.push(["Location", "Lat: " + (inp.lat || ""), "Lon: " + (inp.lon || ""), gmapsUrl_(inp.lat, inp.lon)]);
  rows.push([]);

  var headers = [
    "Dataset Key", "Description", "ID", "City", "State",
    "Lat", "Lon", "Timezone", "Elevation (m)", "Distance (m)", "Weather Data Source", "Google Maps"
  ];
  rows.push(headers);

  var keys = ["nsrdb", "tmy2", "tmy3", "intl"];
  for (var k = 0; k < keys.length; k++) {
    var key = keys[k];
    if (!out[key]) continue;
    var d = out[key];
    rows.push([
      key,
      DATASET_DESCRIPTIONS[key] || "",
      d.id || "",
      d.city || "",
      d.state || "",
      d.lat != null ? d.lat : "",
      d.lon != null ? d.lon : "",
      d.timezone != null ? d.timezone : "",
      d.elevation != null ? d.elevation : "",
      d.distance != null ? d.distance : "",
      d.weather_data_source || "",
      gmapsUrl_(d.lat, d.lon)
    ]);
  }

  var colCount = headers.length;
  sheet.getRange(1, 1, rows.length, colCount).setValues(
    rows.map(function(r) {
      while (r.length < colCount) r.push("");
      return r;
    })
  );

  formatHeader_(sheet, 1, colCount);
  formatHeader_(sheet, 4, colCount);
  sheet.autoResizeColumns(1, colCount);
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function round_(val) {
  if (typeof val !== "number") return val;
  return Math.round(val * 100) / 100;
}

function gmapsUrl_(lat, lon) {
  if (lat == null || lon == null || lat === "" || lon === "") return "";
  return "https://www.google.com/maps?q=" + lat + "," + lon;
}

function formatHeader_(sheet, row, cols) {
  sheet.getRange(row, 1, 1, cols)
    .setFontWeight("bold")
    .setBackground("#e8f0fe");
}
