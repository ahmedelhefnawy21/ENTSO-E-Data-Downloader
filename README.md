<h1 align="center">ENTSO-E Data Downloader</h1>

<p align="center">    
  <img src="https://img.shields.io/badge/License-MIT-blue.svg" alt="License: MIT">
</p>

Automated **R** scripts that fetch, clean, and summarise electricity data from the European Network of Transmission System Operators for Electricity ([ENTSO-E](https://www.entsoe.eu/)) Transparency Platform—installed capacity, total load, and generation by technology—for **38 European bidding zones (2020 – 2025)**, saving everything to ready-to-use Excel workbooks.

<p align="center">
  <img src="https://raw.githubusercontent.com/ahmedelhefnawy21/ENTSO-E-Data-Downloader/main/ENTSO-E.svg.png" width="300px" alt="ENTSO-E Logo"><br> 
</p>

**Note: This repository is NOT officially affilited with ENTSO-E.**

---
## Overview

The repository includes four R scripts, each designed for a specific task:

- **`entsoe_capacity.R`**: Downloads yearly installed capacity data per production type for each country, producing individual Excel files and a consolidated workbook.
- **`entsoe_load.R`**: Retrieves hourly total load data per country and year, generating individual Excel files and a summary workbook with trend analysis.
- **`entsoe_generation.R`**: Fetches hourly generation data per production type, organizing it into per-type folders with individual Excel files.
- **`summary_generation.R`**: Aggregates generation data into summary workbooks, one per production type.

### Why it’s useful

- **End-to-end** pipeline: API ▶ tidy tables ▶ Excel workbooks.

- **Self-healing**: resumes after API hiccups & rate limits through smart retries + automatic sleeps.

- **Trend tagging** for load data by having a trend column (“increasing / decreasing / fluctuating”).

- Works **headless**—cron, Task Scheduler, or GitHub Actions.

---
## Coverage

| Category | Technology (bucket name used in this repo) | ENTSO-E PSR code |
|----------|--------------------------------------------|------------------|
| **Renewable** | Wind Onshore | `B19` |
| | Wind Offshore | `B18` |
| | Solar PV | `B16` |
| | Hydro – Run-of-river & poundage | `B12` |
| | Hydro – Water reservoir | `B11` |
| | Hydro – Pumped storage (pump-up mode only) | `B21` |
| | Marine (tidal / wave) | `B23` |
| | Biomass | `B10` |
| | Geothermal | `B17` |
| | Other renewable | `B20` |
| **Low-carbon** | Nuclear | `B09` |
| **Fossil** | Fossil Gas | `B04` |
| | Fossil Hard coal | `B05` |
| | Fossil Brown coal / Lignite | `B02` |
| | Fossil Coal-derived gas | `B03` |
| | Fossil Oil | `B06` |
| | Fossil Oil shale | `B07` |
| | Fossil Peat | `B08` |
| **Other** | Waste (municipal/industrial) | `B15` |

---

### Countries Covered

| Code | Country       | Code | Country       | Code | Country       |
|------|---------------|------|---------------|------|---------------|
| AL   | Albania       | LU   | Luxembourg    | SE   | Sweden        |
| AT   | Austria       | LV   | Latvia        | SI   | Slovenia      |
| BA   | Bosnia & Herzegovina | MD   | Moldova        | SK   | Slovakia      |
| BE   | Belgium       | ME   | Montenegro    | UA   | Ukraine       |
| BG   | Bulgaria      | MK   | North Macedonia | UK   | United Kingdom |
| CH   | Switzerland   | NL   | Netherlands   | XK   | Kosovo        |
| CY   | Cyprus        | NO   | Norway        | CZ   | Czech Republic|
| DE   | Germany       | PL   | Poland        | DK   | Denmark       |
| EE   | Estonia       | PT   | Portugal      | ES   | Spain         |
| FI   | Finland       | RO   | Romania       | FR   | France        |
| GE   | Georgia       | RS   | Serbia        | GR   | Greece        |
| HR   | Croatia       | IT   | Italy         | HU   | Hungary       |
| IE   | Ireland       | LT   | Lithuania     |

---

## Prerequisites

- **R** (version ≥ 4.0 recommended)
- **R Packages**:
  - `entsoeapi`, `openxlsx`, `dplyr`, `tidyr`, `lubridate`, `readxl`, `writexl`, `tibble`, `devtools`
- **ENTSO-E API Token**:
  - Obtain a free token from the [ENTSO-E Transparency Platform](https://transparency.entsoe.eu/).
  - Set it as an environment variable:
    ```bash
    export ENTSOE_PAT=your_36_character_token
    ```
    Or in R:
    ```R
    Sys.setenv(ENTSOE_PAT = "your_36_character_token")
    ```
---

## Installation & usage

1. **Clone the Repository**:
   ```bash
   git clone https://github.com/ahmedelhefnawy21/ENTSO-E-Data-Downloader.git
   cd ENTSO-E-Data-Downloader
   ```

2. **Install R Packages**:
   ```R
   install.packages(c("devtools", "openxlsx", "tibble", "dplyr", "tidyr", "lubridate", "readxl", "writexl"))
   devtools::install_github(repo = "krose/entsoeapi", ref = "master")
   ```

3. **Set Up Your Environment**:
   - Ensure your API token is set (see [Prerequisites](#prerequisites)).


4. **Running the Scripts**:

Run the scripts in R or via the command line. Adjust your working directory accordingly.

#### In R:
```R
source("scripts/entsoe_capacity.R")    # Downloads capacity data
source("scripts/entsoe_load.R")        # Downloads load data
source("scripts/entsoe_generation.R")  # Downloads generation data
source("scripts/summary_generation.R") # Summarizes generation data
```

**Note**: Run `summary_generation.R` only after `entsoe_generation.R` completes.

#### Via Command Line:
```bash
Rscript scripts/entsoe_capacity.R
Rscript scripts/entsoe_load.R
Rscript scripts/entsoe_generation.R
Rscript scripts/summary_generation.R
```

---
## Output Files

- **Capacity**:
  - Individual files: `data/capacity/Inst_Cap_Type_<Country>_<Year>.xlsx`
  - Summary: `data/capacity/capacity_entsoe.xlsx` (columns: DEFINITION, 2020, 2021, ...)

- **Load**:
  - Individual files: `data/load/Total_Load_<Country>_<Year>.xlsx`
  - Summary: `data/load/load_entsoe.xlsx` (columns: Hour, 2020, 2021, ..., Trend)

- **Generation**:
  - Individual files: `data/generation/entsoe_generation_<type>/entsoe_generation_<type>/Actual_Gen_<Type>_<Country>_<Year>.xlsx` (e.g., `entsoe_generation_windon/Actual_Gen_WindOn_AT_2020.xlsx`)
  - Summary file for each type: `data/generation/entsoe_generation_<type>/Actual_Gen_<Type>_summary.xlsx` (sheets per country)
---
## Notes

- **Time Zones**: Load and generation data use local time zones (e.g., CET for Germany, Asia/Tbilisi for Georgia).
- **API Limits**: Scripts include `Sys.sleep(0.2)` or `Sys.sleep(0.5)` to avoid rate limits; adjust if necessary.
- **File Overwrites**: Individual files may overwrite existing ones; summaries update existing data. Back up important files.
- **Portability**: No Excel installation needed (`openxlsx` is used).
- **Customization**: Edit the `years` variable (e.g., `2020:2025`) or country lists in the scripts as needed.

