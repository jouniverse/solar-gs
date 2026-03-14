/**
 * Google Solar GS — Data Processing
 * Transforms API responses into 2D arrays for sheet writing and UI display
 */

/**
 * Processes individual solar panels from building insights.
 * Merges panel data with roof segment stats (pitch, azimuth).
 * @param {Object} buildingInsights - The full building insights response
 * @param {number} [customCapacityWatts] - Custom panel capacity (overrides API default)
 * @return {Array[]} 2D array with headers as first row
 */
function processIndividualPanels(buildingInsights, customCapacityWatts) {
  var sp = buildingInsights.solarPotential;
  if (!sp || !sp.solarPanels || sp.solarPanels.length === 0) {
    return [
      [
        "Panel #",
        "Segment",
        "Latitude",
        "Longitude",
        "Energy (kWh/yr)",
        "Pitch (°)",
        "Azimuth (°)",
        "Orientation",
      ],
    ];
  }

  var defaultCapacity = sp.panelCapacityWatts || 250;
  var capacityRatio = customCapacityWatts
    ? customCapacityWatts / defaultCapacity
    : 1;
  var segments = sp.roofSegmentStats || [];
  var panels = sp.solarPanels;

  var data = [
    [
      "Panel #",
      "Segment",
      "Latitude",
      "Longitude",
      "Energy (kWh/yr)",
      "Pitch (°)",
      "Azimuth (°)",
      "Orientation",
    ],
  ];

  for (var i = 0; i < panels.length; i++) {
    var panel = panels[i];
    var segIdx = panel.segmentIndex || 0;
    var seg = segments[segIdx] || {};
    data.push([
      i + 1,
      segIdx,
      panel.center ? panel.center.latitude : "",
      panel.center ? panel.center.longitude : "",
      Math.round((panel.yearlyEnergyDcKwh || 0) * capacityRatio * 100) / 100,
      seg.pitchDegrees !== undefined
        ? Math.round(seg.pitchDegrees * 100) / 100
        : "",
      seg.azimuthDegrees !== undefined
        ? Math.round(seg.azimuthDegrees * 100) / 100
        : "",
      panel.orientation || "LANDSCAPE",
    ]);
  }

  return data;
}

/**
 * Processes solar panel configurations.
 * @param {Object} buildingInsights
 * @param {number} [customCapacityWatts]
 * @return {Array[]} 2D array with headers
 */
function processPanelConfigs(buildingInsights, customCapacityWatts) {
  var sp = buildingInsights.solarPotential;
  if (!sp || !sp.solarPanelConfigs) {
    return [
      [
        "Config #",
        "Panel Count",
        "Yearly Energy (kWh)",
        "Avg Energy/Panel (kWh)",
      ],
    ];
  }

  var defaultCapacity = sp.panelCapacityWatts || 250;
  var capacityRatio = customCapacityWatts
    ? customCapacityWatts / defaultCapacity
    : 1;
  var configs = sp.solarPanelConfigs;

  var data = [
    [
      "Config #",
      "Panel Count",
      "Yearly Energy (kWh)",
      "Avg Energy/Panel (kWh)",
    ],
  ];

  for (var i = 0; i < configs.length; i++) {
    var cfg = configs[i];
    var energy = (cfg.yearlyEnergyDcKwh || 0) * capacityRatio;
    var count = cfg.panelsCount || 0;
    data.push([
      i + 1,
      count,
      Math.round(energy * 100) / 100,
      count > 0 ? Math.round((energy / count) * 100) / 100 : 0,
    ]);
  }

  return data;
}

/**
 * Processes roof segment statistics.
 * @param {Object} buildingInsights
 * @return {Array[]} 2D array with headers
 */
function processRoofSegments(buildingInsights) {
  var sp = buildingInsights.solarPotential;
  if (!sp || !sp.roofSegmentStats) {
    return [
      [
        "Segment #",
        "Pitch (°)",
        "Azimuth (°)",
        "Area (m²)",
        "Height (m)",
        "Center Lat",
        "Center Lng",
      ],
    ];
  }

  var segments = sp.roofSegmentStats;
  var data = [
    [
      "Segment #",
      "Pitch (°)",
      "Azimuth (°)",
      "Area (m²)",
      "Height (m)",
      "Center Lat",
      "Center Lng",
    ],
  ];

  for (var i = 0; i < segments.length; i++) {
    var seg = segments[i];
    data.push([
      i,
      Math.round((seg.pitchDegrees || 0) * 100) / 100,
      Math.round((seg.azimuthDegrees || 0) * 100) / 100,
      Math.round((seg.stats ? seg.stats.areaMeters2 : 0) * 100) / 100,
      Math.round((seg.planeHeightAtCenterMeters || 0) * 100) / 100,
      seg.center ? seg.center.latitude : "",
      seg.center ? seg.center.longitude : "",
    ]);
  }

  return data;
}

/**
 * Processes a solar summary (key-value pairs).
 * @param {Object} buildingInsights
 * @return {Array[]} 2D array with headers
 */
function processSolarSummary(buildingInsights) {
  var sp = buildingInsights.solarPotential || {};
  var wrs = sp.wholeRoofStats || {};

  var data = [["Parameter", "Value", "Unit"]];
  data.push(["Building Name", buildingInsights.name || "-", ""]);
  data.push([
    "Latitude",
    buildingInsights.center ? buildingInsights.center.latitude : "-",
    "°",
  ]);
  data.push([
    "Longitude",
    buildingInsights.center ? buildingInsights.center.longitude : "-",
    "°",
  ]);
  data.push(["Postal Code", buildingInsights.postalCode || "-", ""]);
  data.push(["Region", buildingInsights.regionCode || "-", ""]);
  data.push(["Imagery Date", formatApiDate_(buildingInsights.imageryDate), ""]);
  data.push(["Imagery Quality", buildingInsights.imageryQuality || "-", ""]);
  data.push(["Max Panel Count", sp.maxArrayPanelsCount || 0, "panels"]);
  data.push([
    "Max Array Area",
    Math.round((sp.maxArrayAreaMeters2 || 0) * 100) / 100,
    "m²",
  ]);
  data.push([
    "Max Sunshine Hours",
    Math.round((sp.maxSunshineHoursPerYear || 0) * 100) / 100,
    "hrs/yr",
  ]);
  data.push([
    "Carbon Offset Factor",
    Math.round((sp.carbonOffsetFactorKgPerMwh || 0) * 100) / 100,
    "kg/MWh",
  ]);
  data.push(["Panel Capacity", sp.panelCapacityWatts || "-", "W"]);
  data.push(["Panel Lifetime", sp.panelLifetimeYears || "-", "years"]);
  data.push([
    "Roof Area (total)",
    Math.round((wrs.areaMeters2 || 0) * 100) / 100,
    "m²",
  ]);
  data.push([
    "Roof Area (ground)",
    Math.round((wrs.groundAreaMeters2 || 0) * 100) / 100,
    "m²",
  ]);
  data.push(["Roof Segments", (sp.roofSegmentStats || []).length, ""]);

  return data;
}

/**
 * Generates a 20-year energy projection with degradation.
 * @param {Object} buildingInsights
 * @param {number} panelCount - Number of panels to model
 * @param {number} [customCapacityWatts]
 * @param {number} [degradationRate] - Annual degradation fraction (default 0.005)
 * @return {Array[]} 2D array with headers
 */
function processEnergyProjection(
  buildingInsights,
  panelCount,
  customCapacityWatts,
  degradationRate,
) {
  degradationRate = degradationRate || 0.005;
  var sp = buildingInsights.solarPotential;
  if (!sp || !sp.solarPanelConfigs) {
    return [["Year", "Energy (kWh)", "Degradation Factor"]];
  }

  // Find config closest to requested panel count
  var configs = sp.solarPanelConfigs;
  var defaultCapacity = sp.panelCapacityWatts || 250;
  var capacityRatio = customCapacityWatts
    ? customCapacityWatts / defaultCapacity
    : 1;

  var baseEnergy = 0;
  for (var i = 0; i < configs.length; i++) {
    if (configs[i].panelsCount >= panelCount) {
      baseEnergy = (configs[i].yearlyEnergyDcKwh || 0) * capacityRatio;
      break;
    }
  }
  // If panelCount exceeds all configs, use the last one
  if (baseEnergy === 0 && configs.length > 0) {
    baseEnergy =
      (configs[configs.length - 1].yearlyEnergyDcKwh || 0) * capacityRatio;
  }

  var lifetime = sp.panelLifetimeYears || 20;
  var data = [
    ["Year", "Energy (kWh)", "Cumulative Energy (kWh)", "Degradation Factor"],
  ];
  var cumulative = 0;

  for (var y = 1; y <= lifetime; y++) {
    var factor = Math.pow(1 - degradationRate, y);
    var energy = Math.round(baseEnergy * factor * 100) / 100;
    cumulative += energy;
    data.push([
      y,
      energy,
      Math.round(cumulative * 100) / 100,
      Math.round(factor * 10000) / 10000,
    ]);
  }

  return data;
}

/**
 * Processes sunshine quantiles from building insights.
 * @param {Object} buildingInsights
 * @return {Array[]} 2D array with headers
 */
function processSunshineQuantiles(buildingInsights) {
  var sp = buildingInsights.solarPotential || {};
  var wrs = sp.wholeRoofStats || {};
  var quantiles = wrs.sunshineQuantiles || [];

  var data = [["Percentile (%)", "Sunshine (kWh/kW/yr)"]];
  for (var i = 0; i < quantiles.length; i++) {
    var pct = Math.round((i / (quantiles.length - 1)) * 100);
    data.push([pct, Math.round(quantiles[i] * 100) / 100]);
  }
  return data;
}

/**
 * Processes cost analysis data for sheet export.
 * @param {Object} params - Cost analysis parameters and results from the client
 * @return {Array[]} 2D array with headers
 */
function processCostAnalysis(params) {
  var data = [["Parameter", "Value", "Unit"]];
  data.push(["Panels Count", params.panelsCount || 0, "panels"]);
  data.push(["Panel Capacity", params.panelCapacityWatts || 0, "W"]);
  data.push(["Installation Size", params.installationSizeKw || 0, "kW"]);
  data.push(["Installation Cost", params.installationCostTotal || 0, "$"]);
  data.push(["Yearly Energy DC", params.yearlyEnergyDcKwh || 0, "kWh"]);
  data.push(["Monthly Energy Bill", params.monthlyAverageEnergyBill || 0, "$"]);
  data.push(["Energy Cost", params.energyCostPerKwh || 0, "$/kWh"]);
  data.push(["Solar Incentives", params.solarIncentives || 0, "$"]);
  data.push([
    "Installation Cost/W",
    params.installationCostPerWatt || 0,
    "$/W",
  ]);
  data.push(["DC to AC Derate", params.dcToAcDerate || 0, ""]);
  data.push([
    "Efficiency Depreciation",
    params.efficiencyDepreciationFactor || 0,
    "",
  ]);
  data.push(["Cost Increase Factor", params.costIncreaseFactor || 0, ""]);
  data.push(["Discount Rate", params.discountRate || 0, ""]);
  data.push(["Lifespan", params.installationLifeSpan || 0, "years"]);
  data.push(["Cost With Solar", params.totalCostWithSolar || 0, "$"]);
  data.push(["Cost Without Solar", params.totalCostWithoutSolar || 0, "$"]);
  data.push(["Savings", params.savings || 0, "$"]);
  return data;
}
