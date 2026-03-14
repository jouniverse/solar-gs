# google-solar-gs - Solar System Designer

A Google Sheets add-on that fetches data from the [Google Solar API](https://developers.google.com/maps/documentation/solar) and displays rooftop solar potential, data layer visualizations, energy charts, and cost estimates — all inside a sidebar within Google Sheets. Data can be injected directly into spreadsheet tabs for further analysis.

## Features

- **Location selection** — Search by address or click on an interactive Leaflet map with satellite/street view toggle.
- **Building insights** — View roof area, panel capacity, segment details, and individual panel positions on a satellite map.
- **Data layers** — Visualize GeoTIFF overlays (RGB, annual flux, monthly flux, building mask) on a map with adjustable opacity.
- **Charts** — Interactive Chart.js visualizations of per-panel energy, total energy by configuration, energy depreciation over time, panel distribution by roof segment, and sunshine quantiles.
- **Cost analysis** — Financial calculator for installation cost, utility savings, and lifetime payback with configurable parameters.
- **Sheet injection** — "Inject to Sheet" buttons throughout the app write data directly into named Google Sheets tabs.
- **Caching** — API responses are cached to minimize redundant calls and associated costs.

## UI

![Solar System Designer UI](./imgs/google-solar-gs-ui-screenshots.png)

## Setup

### Prerequisites

- A Google account with access to Google Sheets.
- A **Google Maps Platform API key** with the Solar API enabled. See [Google's documentation](https://developers.google.com/maps/documentation/solar/cloud-setup) for setup instructions.

> **Note:** The Google Solar API is a paid service. Each Building Insights request and each Data Layers request are billed separately. Review the [Solar API pricing](https://developers.google.com/maps/documentation/solar/usage-and-billing) before use.

### Installation

Deploy using [Clasp](https://developers.google.com/apps-script/guides/clasp) or **copy the files manually** into a Google Apps Script project.

#### Manual Copy Instructions

1. Open a Google Sheet.
2. Go to **Extensions > Apps Script**.
3. Create the following `.gs` files and paste the contents from `src/gs/`:
   - `main.gs`
   - `utils.gs`
   - `cache.gs`
   - `solarApi.gs`
   - `dataProcessing.gs`
   - `sheetWriter.gs`
4. Create the following `.html` files and paste the contents from `src/html/`:
   - `sidebar.html`
   - `css.html`
   - `js-main.html`
   - `js-location.html`
   - `js-building.html`
   - `js-datalayers.html`
   - `js-charts.html`
   - `js-cost.html`
   - `help-dialog.html`
   - `settings-dialog.html`
5. Save the project and reload the Google Sheet.
6. A **Solar API** menu will appear in the menu bar.
7. Go to **Solar API > Settings** and enter your Google Maps API key.

## Usage

1. Open the sidebar via **Solar API > Open Sidebar**.
2. In the **Location** tab, search for an address or click the map to place a marker.
3. Click **Fetch Solar Data** to retrieve building insights and data layers from the API.
4. Browse the tabs to explore the results:
   - **Building** — Summary cards, roof segments, individual panels map and table.
   - **Layers** — Toggle data layer overlays (RGB, annual/monthly flux, mask) on a satellite map.
   - **Charts** — Energy charts with per-panel statistics, configurations, and depreciation projections.
   - **Cost** — Configure financial parameters and calculate lifetime savings.
5. Use the **Inject to Sheet** buttons to write data into spreadsheet tabs for further analysis or export.

## Glossary

| Term                               | Description                                                                                                                                                      |
| ---------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Pitch**                          | The angle of a roof segment relative to the horizontal plane, in degrees. A flat roof has a pitch of 0°.                                                         |
| **Azimuth**                        | The compass direction a roof segment faces, in degrees clockwise from north. 0° = north, 90° = east, 180° = south, 270° = west.                                  |
| **Orientation**                    | Whether a solar panel is placed in landscape (L) or portrait (P) orientation on the roof.                                                                        |
| **Annual Flux**                    | The total solar energy received per square meter of roof surface over one year (kWh/m²/yr).                                                                      |
| **Monthly Flux**                   | The solar energy received per square meter for a specific month, provided as 12 separate bands in a single GeoTIFF.                                              |
| **Sunshine Quantiles**             | A distribution of solar irradiance values across the entire roof surface. Shows what percentage of the roof receives at least a given amount of sunshine.        |
| **Yearly Energy DC**               | The total direct-current energy a panel configuration can generate in one year (kWh/yr), before conversion losses.                                               |
| **Panel Capacity**                 | The rated power output of a single solar panel in watts (W).                                                                                                     |
| **DC-to-AC Derate**                | The conversion efficiency factor from DC to AC power. Accounts for inverter losses, wiring losses, and other system inefficiencies. Typically around 0.85 (85%). |
| **Efficiency Depreciation Factor** | The annual degradation rate of panel performance. A factor of 0.995 means panels lose 0.5% efficiency per year.                                                  |
| **Cost Increase Factor**           | The expected annual increase in electricity prices. A factor of 1.022 means a 2.2% yearly increase.                                                              |
| **Discount Rate**                  | The rate used to discount future cash flows to present value. A factor of 1.04 means a 4% annual discount rate.                                                  |
| **DSM (Digital Surface Model)**    | A 3D representation of the Earth's surface including buildings and trees, used to understand roof geometry and shading.                                          |
| **Building Mask**                  | A binary layer that isolates the building footprint from the surrounding area.                                                                                   |
| **CO₂ Offset Factor**              | The amount of carbon dioxide offset per unit of solar energy produced (kg/MWh).                                                                                  |

## Interpreting the Data

### Building Insights

The **Summary Cards** show key metrics at a glance: maximum panel count, usable roof area, peak sunshine hours per year, and CO₂ offset potential. The **Solar Potential** section details roof area (total and ground-projected), segment count, and panel configuration options.

The **Panel Locations** map displays individual panel positions as color-coded dots — blue indicates lower energy yield, red indicates higher yield. Panel positions may appear slightly offset from the satellite roof imagery due to differences between Google's data source and the Esri satellite tiles.

### Data Layers

- **RGB** — A true-color aerial image of the building. Useful as a visual reference.
- **Annual Flux** — A heatmap showing yearly solar irradiance across the roof. Brighter areas receive more sunlight and are better candidates for panel placement.
- **Monthly Flux** — The same as annual flux but broken down by month. Useful for understanding seasonal variation.
- **Mask** — Shows where the building is. Useful for understanding panel placement context.

### Charts

- **Panel # vs. Energy** — Shows the energy yield of each individual panel. Statistics (min, max, mean, median, std dev) are displayed above the chart.
- **Number of panels vs. Total energy** — Shows how total system energy increases as more panels are added.
- **Year vs. Projected energy** — Shows energy output declining over the system lifetime due to panel degradation.
- **Roof segment vs. Panel count** — Shows how panels are distributed across different roof segments.
- **Sunshine quantiles** — Shows the distribution of sunlight across the roof area.

### Cost Analysis

The Cost tab uses configurable parameters (monthly bill, energy cost, incentives, installation cost per watt, and advanced factors) to estimate lifetime savings. The calculation compares projected electricity costs with and without solar over the panel lifetime, accounting for energy price increases, panel degradation, and discount rates.

## Project Structure

```
src/
├── gs/                          # Server-side Google Apps Script
│   ├── main.gs                  # Menu, sidebar, dialogs, API key storage
│   ├── utils.gs                 # Utility functions
│   ├── cache.gs                 # API response caching
│   ├── solarApi.gs              # Solar API request handling
│   ├── dataProcessing.gs        # Transform API responses into 2D arrays
│   └── sheetWriter.gs           # Write data to named sheets
└── html/                        # Client-side HTML/CSS/JS
    ├── sidebar.html             # Main HTML shell with tab navigation
    ├── css.html                 # Stylesheet
    ├── js-main.html             # Global state, tab switching, utilities
    ├── js-location.html         # Location map and geocoding
    ├── js-building.html         # Building insights, panels map, tables
    ├── js-datalayers.html       # GeoTIFF data layer rendering
    ├── js-charts.html           # Chart.js visualizations
    ├── js-cost.html             # Financial cost calculator
    ├── help-dialog.html         # Help dialog content
    └── settings-dialog.html     # API key settings dialog
```

## Client-Side Libraries

| Library                                               | Version | Purpose                                                            |
| ----------------------------------------------------- | ------- | ------------------------------------------------------------------ |
| [Leaflet.js](https://leafletjs.com/)                  | 1.9.4   | Interactive maps (location selection, panels map, data layers map) |
| [geotiff.js](https://geotiffjs.github.io/geotiff.js/) | 2.1.3   | GeoTIFF parsing and raster data extraction                         |
| [proj4js](http://proj4js.org/)                        | 2.12.1  | Coordinate reference system reprojection for GeoTIFF bounds        |
| [Chart.js](https://www.chartjs.org/)                  | 4.4.7   | Interactive chart rendering                                        |

## API Reference

- [Google Solar API Overview](https://developers.google.com/maps/documentation/solar)
- [Building Insights endpoint](https://developers.google.com/maps/documentation/solar/building-insights)
- [Data Layers endpoint](https://developers.google.com/maps/documentation/solar/data-layers)
- [Cost calculation guide](https://developers.google.com/maps/documentation/solar/calculate-costs-typescript)
- [Usage and billing](https://developers.google.com/maps/documentation/solar/usage-and-billing)

## License

**MIT License.**

This project is not currently licensed for redistribution.
