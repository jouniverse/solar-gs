/**
 * Google Solar GS — Solar API calls
 * Fetches Building Insights, Data Layers, and GeoTIFF data
 */

/**
 * Fetches Building Insights for a location.
 * @param {number} lat - Latitude
 * @param {number} lng - Longitude
 * @param {string} quality - Required quality (HIGH, MEDIUM, LOW)
 * @return {Object} {success: boolean, data?: Object, error?: string}
 */
function fetchBuildingInsights(lat, lng, quality) {
  try {
    if (!validateCoordinates_(lat, lng)) {
      return {
        success: false,
        error:
          "Invalid coordinates. Latitude must be -90 to 90, longitude -180 to 180.",
      };
    }
    quality = quality || "HIGH";

    // Check cache first
    var cacheKey = buildingInsightsCacheKey_(lat, lng, quality);
    var cached = cacheGet_(cacheKey);
    if (cached) {
      return { success: true, data: JSON.parse(cached), fromCache: true };
    }

    var url = buildBuildingInsightsUrl_(lat, lng, quality);
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var code = response.getResponseCode();
    var body = response.getContentText();

    if (code !== 200) {
      var errorMsg = "API returned status " + code;
      try {
        var errObj = JSON.parse(body);
        if (errObj.error && errObj.error.message) {
          errorMsg = errObj.error.message;
        }
      } catch (e) {
        /* use default message */
      }
      return { success: false, error: errorMsg };
    }

    var data = JSON.parse(body);

    // Cache the response
    cacheSet_(cacheKey, body);

    return { success: true, data: data, fromCache: false };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Fetches Data Layer URLs for a location.
 * @param {number} lat - Latitude
 * @param {number} lng - Longitude
 * @param {number} radiusMeters - Radius in meters (default 50)
 * @param {string} quality - Required quality (HIGH, MEDIUM, LOW)
 * @return {Object} {success: boolean, data?: Object, error?: string}
 */
function fetchDataLayers(lat, lng, radiusMeters, quality) {
  try {
    if (!validateCoordinates_(lat, lng)) {
      return {
        success: false,
        error:
          "Invalid coordinates. Latitude must be -90 to 90, longitude -180 to 180.",
      };
    }
    radiusMeters = radiusMeters || 50;
    quality = quality || "HIGH";

    // Check cache first
    var cacheKey = dataLayersCacheKey_(lat, lng, radiusMeters, quality);
    var cached = cacheGet_(cacheKey);
    if (cached) {
      return { success: true, data: JSON.parse(cached), fromCache: true };
    }

    var url = buildDataLayersUrl_(lat, lng, radiusMeters, quality);
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var code = response.getResponseCode();
    var body = response.getContentText();

    if (code !== 200) {
      var errorMsg = "API returned status " + code;
      try {
        var errObj = JSON.parse(body);
        if (errObj.error && errObj.error.message) {
          errorMsg = errObj.error.message;
        }
      } catch (e) {
        /* use default message */
      }
      return { success: false, error: errorMsg };
    }

    var data = JSON.parse(body);
    cacheSet_(cacheKey, body);

    return { success: true, data: data, fromCache: false };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Downloads a GeoTIFF file and returns it as a base64-encoded string.
 * The URL comes from the Data Layers response and needs the API key appended.
 * @param {string} geoTiffUrl - The GeoTIFF URL from the API response
 * @return {Object} {success: boolean, data?: string, error?: string}
 */
function fetchGeoTiff(geoTiffUrl) {
  try {
    if (!geoTiffUrl) {
      return { success: false, error: "No GeoTIFF URL provided." };
    }

    var apiKey = getApiKey_();
    var fullUrl = geoTiffUrl + "&key=" + apiKey;

    var response = UrlFetchApp.fetch(fullUrl, { muteHttpExceptions: true });
    var code = response.getResponseCode();

    if (code !== 200) {
      return {
        success: false,
        error: "Failed to download GeoTIFF. Status: " + code,
      };
    }

    var blob = response.getBlob();
    var base64 = Utilities.base64Encode(blob.getBytes());

    return { success: true, data: base64 };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Combined fetch: gets both Building Insights and Data Layers for a location.
 * Used as a single call from the sidebar to reduce round-trips.
 * @param {number} lat
 * @param {number} lng
 * @param {number} radiusMeters
 * @param {string} quality
 * @return {Object} {buildingInsights: Object, dataLayers: Object}
 */
function fetchSolarData(lat, lng, radiusMeters, quality) {
  var bi = fetchBuildingInsights(lat, lng, quality);
  var dl = fetchDataLayers(lat, lng, radiusMeters, quality);
  return {
    buildingInsights: bi,
    dataLayers: dl,
  };
}
