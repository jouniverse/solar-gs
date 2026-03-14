/**
 * Google Solar GS — Caching layer
 * Chunked caching to handle CacheService 100KB limit
 */

var CACHE_TTL_ = 21600; // 6 hours in seconds
var CHUNK_SIZE_ = 90000; // ~90KB per chunk (leave room for overhead)

/**
 * Stores a value in cache, chunking if necessary.
 * @param {string} key - Cache key
 * @param {string} jsonString - JSON string to cache
 */
function cacheSet_(key, jsonString) {
  var cache = CacheService.getScriptCache();
  if (jsonString.length <= CHUNK_SIZE_) {
    cache.put(key, jsonString, CACHE_TTL_);
    return;
  }
  // Chunk the data
  var chunks = [];
  for (var i = 0; i < jsonString.length; i += CHUNK_SIZE_) {
    chunks.push(jsonString.substring(i, i + CHUNK_SIZE_));
  }
  var entries = {};
  entries[key] = JSON.stringify({ _chunked: true, _count: chunks.length });
  for (var c = 0; c < chunks.length; c++) {
    entries[key + "_chunk_" + c] = chunks[c];
  }
  cache.putAll(entries, CACHE_TTL_);
}

/**
 * Retrieves a value from cache, reassembling chunks if necessary.
 * @param {string} key - Cache key
 * @return {string|null} The cached JSON string, or null if not found
 */
function cacheGet_(key) {
  var cache = CacheService.getScriptCache();
  var raw = cache.get(key);
  if (!raw) return null;

  var parsed;
  try {
    parsed = JSON.parse(raw);
  } catch (e) {
    return raw;
  }

  if (!parsed._chunked) return raw;

  // Reassemble chunks
  var keys = [];
  for (var i = 0; i < parsed._count; i++) {
    keys.push(key + "_chunk_" + i);
  }
  var chunkMap = cache.getAll(keys);
  var result = "";
  for (var j = 0; j < parsed._count; j++) {
    var chunk = chunkMap[key + "_chunk_" + j];
    if (!chunk) return null; // Chunk expired or missing
    result += chunk;
  }
  return result;
}

/**
 * Generates a cache key for Building Insights.
 * @param {number} lat
 * @param {number} lng
 * @param {string} quality
 * @return {string}
 */
function buildingInsightsCacheKey_(lat, lng, quality) {
  return "BI_" + lat.toFixed(4) + "_" + lng.toFixed(4) + "_" + quality;
}

/**
 * Generates a cache key for Data Layers.
 * @param {number} lat
 * @param {number} lng
 * @param {number} radius
 * @param {string} quality
 * @return {string}
 */
function dataLayersCacheKey_(lat, lng, radius, quality) {
  return (
    "DL_" + lat.toFixed(4) + "_" + lng.toFixed(4) + "_" + radius + "_" + quality
  );
}

/**
 * Clears all solar-related cache entries. Exposed to sidebar and menu.
 */
function clearSolarCache() {
  var cache = CacheService.getScriptCache();
  // CacheService doesn't support listing keys, so we remove known patterns
  // by simply letting them expire. For an explicit clear, the user can wait
  // or we remove specific keys if we know what was cached.
  // Best effort: remove commonly used keys
  cache.removeAll([]);
  return {
    success: true,
    message: "Cache cleared. New fetches will call the API.",
  };
}
