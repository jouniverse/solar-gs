/**
 * PVGIS API Configuration
 */

var PVGIS_CONFIG = {
  BASE_URL: "https://re.jrc.ec.europa.eu/api/v5_3",

  ENDPOINTS: {
    GRID_CONNECTED: "PVcalc",
    TRACKING_PV: "PVcalc",
    OFF_GRID: "SHScalc",
    MONTHLY_RADIATION: "MRcalc",
    DAILY_RADIATION: "DRcalc",
    HOURLY_RADIATION: "seriescalc",
    TMY: "tmy",
    HORIZON_PROFILE: "printhorizon",
  },

  /** Radiation database options */
  RADIATION_DATABASES: [
    { value: "PVGIS-SARAH3", label: "PVGIS-SARAH3 (Europe, Africa, Asia)" },
    { value: "PVGIS-NSRDB", label: "PVGIS-NSRDB (Americas)" },
    { value: "PVGIS-ERA5", label: "PVGIS-ERA5 (Global, inc. high-lat)" },
    { value: "PVGIS-COSMO", label: "PVGIS-COSMO (Europe)" },
  ],

  /** PV technology options */
  PV_TECHNOLOGIES: [
    { value: "crystSi", label: "Crystalline Silicon (original)" },
    { value: "crystSi2025", label: "Crystalline Silicon (2025)" },
    { value: "CIS", label: "CIS" },
    { value: "CdTe", label: "CdTe" },
    { value: "Unknown", label: "Unknown" },
  ],

  /** Mounting position options */
  MOUNTING_TYPES: [
    { value: "free", label: "Free-standing" },
    { value: "building", label: "Building-integrated" },
  ],

  /** Hourly radiation tracking types */
  TRACKING_TYPES: [
    { value: 0, label: "Fixed" },
    { value: 1, label: "Single horizontal axis (N-S)" },
    { value: 2, label: "Two-axis tracking" },
    { value: 3, label: "Vertical axis tracking" },
    { value: 4, label: "Single horizontal axis (E-W)" },
    { value: 5, label: "Single inclined axis (N-S)" },
  ],

  /** Months for daily radiation */
  MONTHS: [
    { value: 0, label: "All months" },
    { value: 1, label: "January" },
    { value: 2, label: "February" },
    { value: 3, label: "March" },
    { value: 4, label: "April" },
    { value: 5, label: "May" },
    { value: 6, label: "June" },
    { value: 7, label: "July" },
    { value: 8, label: "August" },
    { value: 9, label: "September" },
    { value: 10, label: "October" },
    { value: 11, label: "November" },
    { value: 12, label: "December" },
  ],

  /** Year range for databases */
  YEAR_RANGES: {
    "PVGIS-SARAH3": { min: 2005, max: 2023 },
    "PVGIS-NSRDB": { min: 2005, max: 2023 },
    "PVGIS-ERA5": { min: 2005, max: 2023 },
    "PVGIS-COSMO": { min: 2005, max: 2023 },
  },

  /** TMY period options */
  TMY_PERIODS: [
    { value: "PVGIS-SARAH3:2005:2023", label: "PVGIS-SARAH3: 2005 - 2023" },
    { value: "PVGIS-ERA5:2005:2023", label: "PVGIS-ERA5: 2005 - 2023" },
  ],

  /** Default values */
  DEFAULTS: {
    peakpower: 1,
    loss: 14,
    angle: 35,
    aspect: 0,
    pvtechchoice: "crystSi",
    mountingplace: "free",
    lifetime: 25,
    batterysize: 600,
    cutoff: 40,
    consumptionday: 300,
    offgrid_peakpower: 50,
  },

  /** Sheet names for output */
  SHEET_NAMES: {
    GRID_CONNECTED: "PVGIS - Grid Connected",
    TRACKING_PV: "PVGIS - Tracking PV",
    OFF_GRID: "PVGIS - Off-Grid",
    MONTHLY_RADIATION: "PVGIS - Monthly Radiation",
    DAILY_RADIATION: "PVGIS - Daily Radiation",
    HOURLY_RADIATION: "PVGIS - Hourly Radiation",
    TMY: "PVGIS - TMY",
    HORIZON_PROFILE: "PVGIS - Horizon Profile",
  },
};
