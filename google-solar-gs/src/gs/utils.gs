/**
 * Google Solar GS — Shared utilities
 * URL builders, formatters, validation
 */

/**
 * Retrieves the Google Maps API key from script properties.
 * @return {string} The API key
 */
function getApiKey_() {
  var key =
    PropertiesService.getScriptProperties().getProperty("GMAPS_API_KEY");
  if (!key) {
    throw new Error(
      "GMAPS_API_KEY not found in Script Properties. Go to Extensions > Apps Script > Project Settings > Script Properties to add it.",
    );
  }
  return key;
}

/**
 * Builds the Building Insights API URL.
 * @param {number} lat - Latitude
 * @param {number} lng - Longitude
 * @param {string} quality - Required quality (HIGH, MEDIUM, or LOW)
 * @return {string} The API URL
 */
function buildBuildingInsightsUrl_(lat, lng, quality) {
  return (
    "https://solar.googleapis.com/v1/buildingInsights:findClosest" +
    "?location.latitude=" +
    lat.toFixed(6) +
    "&location.longitude=" +
    lng.toFixed(6) +
    "&requiredQuality=" +
    encodeURIComponent(quality) +
    "&key=" +
    getApiKey_()
  );
}

/**
 * Builds the Data Layers API URL.
 * @param {number} lat - Latitude
 * @param {number} lng - Longitude
 * @param {number} radiusMeters - Radius in meters
 * @param {string} quality - Required quality (HIGH, MEDIUM, or LOW)
 * @return {string} The API URL
 */
function buildDataLayersUrl_(lat, lng, radiusMeters, quality) {
  return (
    "https://solar.googleapis.com/v1/dataLayers:get" +
    "?location.latitude=" +
    lat.toFixed(6) +
    "&location.longitude=" +
    lng.toFixed(6) +
    "&radiusMeters=" +
    Math.round(radiusMeters) +
    "&view=FULL_LAYERS" +
    "&requiredQuality=" +
    encodeURIComponent(quality) +
    "&exactQualityRequired=false" +
    "&pixelSizeMeters=0.5" +
    "&key=" +
    getApiKey_()
  );
}

/**
 * Validates latitude and longitude values.
 * @param {number} lat - Latitude (-90 to 90)
 * @param {number} lng - Longitude (-180 to 180)
 * @return {boolean} True if valid
 */
function validateCoordinates_(lat, lng) {
  if (typeof lat !== "number" || typeof lng !== "number") return false;
  if (isNaN(lat) || isNaN(lng)) return false;
  if (lat < -90 || lat > 90) return false;
  if (lng < -180 || lng > 180) return false;
  return true;
}

/**
 * Formats a number to fixed decimal places.
 * @param {number} value
 * @param {number} decimals
 * @return {string}
 */
function formatNumber_(value, decimals) {
  if (value === null || value === undefined) return "-";
  return Number(value).toFixed(decimals || 2);
}

/**
 * Formats a date object from the API response.
 * @param {Object} dateObj - {year, month, day}
 * @return {string} Formatted date string
 */
function formatApiDate_(dateObj) {
  if (!dateObj) return "-";
  var m = dateObj.month < 10 ? "0" + dateObj.month : dateObj.month;
  var d = dateObj.day < 10 ? "0" + dateObj.day : dateObj.day;
  return dateObj.year + "-" + m + "-" + d;
}
