/**
 * Builds a full NREL API URL with query parameters.
 */
function buildUrl_(endpoint, params) {
  var apiKey = getApiKey();
  if (!apiKey) throw new Error("NREL API key not set. Use NREL Solar → Set API Key… first.");
  var url = NREL_BASE_URL + endpoint + "?api_key=" + encodeURIComponent(apiKey);
  for (var key in params) {
    if (params[key] !== "" && params[key] !== null && params[key] !== undefined) {
      url += "&" + encodeURIComponent(key) + "=" + encodeURIComponent(params[key]);
    }
  }
  return url;
}

/**
 * Makes an HTTP GET to the NREL API and returns parsed JSON.
 */
function fetchNrel_(endpoint, params) {
  var url = buildUrl_(endpoint, params);
  var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var code = response.getResponseCode();
  var body = JSON.parse(response.getContentText());

  if (code !== 200) {
    var msg = (body.errors && body.errors.length)
      ? body.errors.join("; ")
      : "HTTP " + code;
    throw new Error("NREL API error: " + msg);
  }
  if (body.errors && body.errors.length) {
    throw new Error("NREL API error: " + body.errors.join("; "));
  }
  return body;
}

// ---------------------------------------------------------------------------
// PVWatts V8
// ---------------------------------------------------------------------------

function queryPVWatts(params) {
  var apiParams = {
    lat: params.lat,
    lon: params.lon,
    system_capacity: params.system_capacity || 4,
    module_type: params.module_type != null ? params.module_type : 0,
    losses: params.losses != null ? params.losses : 14,
    array_type: params.array_type != null ? params.array_type : 1,
    tilt: params.tilt != null ? params.tilt : 20,
    azimuth: params.azimuth != null ? params.azimuth : 180,
    dataset: params.dataset || "nsrdb",
    radius: params.radius != null ? params.radius : 100,
    timeframe: params.timeframe || "monthly",
    dc_ac_ratio: params.dc_ac_ratio || 1.2,
    gcr: params.gcr || 0.4,
    inv_eff: params.inv_eff || 96
  };

  if (params.bifaciality) apiParams.bifaciality = params.bifaciality;
  if (params.albedo) apiParams.albedo = params.albedo;
  if (params.soiling) apiParams.soiling = params.soiling;
  if (params.use_wf_albedo) apiParams.use_wf_albedo = params.use_wf_albedo;

  var data = fetchNrel_(ENDPOINTS.PVWATTS_V8, apiParams);
  writePVWattsToSheet(data);

  return {
    inputs: data.inputs,
    station_info: data.station_info,
    warnings: data.warnings || [],
    ac_annual: data.outputs.ac_annual,
    solrad_annual: data.outputs.solrad_annual,
    capacity_factor: data.outputs.capacity_factor
  };
}

// ---------------------------------------------------------------------------
// Solar Resource Data
// ---------------------------------------------------------------------------

function querySolarResource(params) {
  var apiParams = {
    lat: params.lat,
    lon: params.lon
  };

  var data = fetchNrel_(ENDPOINTS.SOLAR_RESOURCE, apiParams);
  writeSolarResourceToSheet(data);

  return {
    warnings: data.warnings || [],
    metadata: data.metadata,
    annual_dni: data.outputs.avg_dni.annual,
    annual_ghi: data.outputs.avg_ghi.annual,
    annual_tilt: data.outputs.avg_lat_tilt.annual
  };
}

// ---------------------------------------------------------------------------
// Solar Dataset Query V2
// ---------------------------------------------------------------------------

function queryDatasetQuery(params) {
  var apiParams = {
    lat: params.lat,
    lon: params.lon
  };
  if (params.radius != null && params.radius !== "") apiParams.radius = params.radius;
  if (params.all) apiParams.all = 1;

  var data = fetchNrel_(ENDPOINTS.DATASET_QUERY_V2, apiParams);
  writeDatasetQueryToSheet(data);

  var datasetCount = data.outputs ? Object.keys(data.outputs).length : 0;
  return {
    warnings: data.warnings || [],
    datasetCount: datasetCount
  };
}
