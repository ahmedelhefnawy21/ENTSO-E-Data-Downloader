## 1) SETUP -------------------------------------------------------------------
setwd("YOUR_WORKING_DIRECTORY")

library(entsoeapi)
library(openxlsx)
library(dplyr)
library(lubridate)
library(tidyr)

security_token <- Sys.getenv("ENTSOE_PAT")

eic_data <- tibble::tribble(
  ~Country, ~EIC,
  "AL","10YAL-KESH-----5",  "AT","10YAT-APG------L",
  "BE","10YBE----------2",  "BA","10YBA-JPCC-----D",
  "BG","10YCA-BULGARIA-R",  "HR","10YHR-HEP------M",
  "CY","10YCY-1001A0003J",  "CZ","10YCZ-CEPS-----N",
  "DK","10Y1001A1001A65H",  "EE","10Y1001A1001A39I",
  "FI","10YFI-1--------U",  "FR","10YFR-RTE------C",
  "GE","10Y1001A1001B012",  "DE","10Y1001A1001A83F",
  "GR","10YGR-HTSO-----Y",  "HU","10YHU-MAVIR----U",
  "IE","10YIE-1001A00010",  "IT","10YIT-GRTN-----B",
  "XK","10Y1001C--00100H",  "LV","10YLV-1001A00074",
  "LT","10YLT-1001A0008Q",  "LU","10YLU-CEGEDEL-NQ",
  "MD","10Y1001A1001A990",  "ME","10YCS-CG-TSO---S",
  "NL","10YNL----------L",  "MK","10YMK-MEPSO----8",
  "NO","10YNO-0--------C",  "PL","10YPL-AREA-----S",
  "PT","10YPT-REN------W",  "RO","10YRO-TEL------P",
  "RS","10YCS-SERBIATSOV",  "SK","10YSK-SEPS-----K",
  "SI","10YSI-ELES-----O",  "ES","10YES-REE------0",
  "SE","10YSE-1--------K",  "UK","10Y1001A1001A92E",
  "UA","10Y1001C--00003F",  "CH","10YCH-SWISSGRIDZ"
)

# PSR reference: keep only production types, with DEFINITION and CODE
psr_list <- asset_types %>%
  slice(-(1:5), -(26:29)) %>%
  select(DEFINITION, CODE)

## 2) Helper: discover all years with valid data -----------------------------
get_available_years <- function(eic,
                                token,
                                first_probe = 2000,
                                empty_streak = 7) {
  this_year     <- year(Sys.Date())
  years_found   <- integer()
  misses_in_row <- 0
  
  for (y in this_year:first_probe) {
    df <- try(
      gen_installed_capacity_per_pt(
        eic            = eic,
        year           = y,
        psr_type       = NULL,
        security_token = token
      ),
      silent = TRUE
    )
    
    # require at least one row AND a PSR-type column
    have_valid <- !inherits(df, "try-error") &&
      nrow(df) > 0 &&
      any(c("ts_mkt_psr_type", "psr_type") %in% names(df))
    
    if (have_valid) {
      years_found   <- c(years_found, y)
      misses_in_row <- 0
    } else {
      misses_in_row <- misses_in_row + 1
      # only stop after we've already found ≥1 valid year
      if (length(years_found) > 0 && misses_in_row >= empty_streak) {
        break
      }
    }
  }
  
  sort(years_found)
}

years_by_cntry <- lapply(
  setNames(eic_data$EIC, eic_data$Country),
  get_available_years,
  token = security_token
)
years_all <- sort(unique(unlist(years_by_cntry)))

cat("Year coverage per country:\n")
print(sapply(years_by_cntry, function(v)
  if (length(v)) paste0(min(v), "–", max(v)) else "none"
))

###############################################################################
# 3) DOWNLOAD / CREATE SINGLE-YEAR FILES --------------------------------------
###############################################################################
for (country in eic_data$Country) {
  yrs <- years_by_cntry[[country]]
  if (length(yrs) == 0) {
    message("—> No capacity data for ", country)
    next
  }
  message("Downloading ", country, ": ", min(yrs), "–", max(yrs))
  
  eic <- eic_data$EIC[match(country, eic_data$Country)]
  
  for (yr in yrs) {
    api_raw <- gen_installed_capacity_per_pt(
      eic            = eic,
      year           = yr,
      psr_type       = NULL,
      security_token = security_token
    )
    
    # Skip if no PSR-type column
    if (!any(c("ts_mkt_psr_type", "psr_type") %in% names(api_raw))) {
      warning("Skipping ", country, "-", yr, ": no PSR-type column")
      next
    }
    # Normalize PSR-type to ts_mkt_psr_type
    if ("psr_type" %in% names(api_raw)) {
      api_raw <- rename(api_raw, ts_mkt_psr_type = psr_type)
    }
    
    # Normalize quantity column
    if ("quantity" %in% names(api_raw) && !"ts_point_quantity" %in% names(api_raw)) {
      api_raw <- rename(api_raw, ts_point_quantity = quantity)
    }
    
    # Build full-template table
    inst_cap_type <- psr_list %>%
      left_join(api_raw, by = c("CODE" = "ts_mkt_psr_type")) %>%
      rename(
        quantity = ts_point_quantity,
        unit     = ts_quantity_measure_unit_name
      ) %>%
      select(DEFINITION, quantity, unit)
    
    write.xlsx(
      inst_cap_type,
      sprintf("Inst_Cap_Type_%s_%d.xlsx", country, yr),
      sheetName = "sheet",
      rowNames  = FALSE
    )
  }
}

###############################################################################
# 4) BUILD / UPDATE CONSOLIDATED WORKBOOK -------------------------------------
###############################################################################
summary_file <- "capacity_entsoe.xlsx"
wb <- if (file.exists(summary_file)) {
  loadWorkbook(summary_file)
} else {
  wb0 <- createWorkbook()
  saveWorkbook(wb0, summary_file, overwrite = TRUE)
  wb0
}

existing_sheets <- sheets(wb)
sheet_list <- vector("list", length = nrow(eic_data))
names(sheet_list) <- eic_data$Country

## 4a) Initialize or import country sheets -----------------------------------
for (country in eic_data$Country) {
  if (country %in% existing_sheets) {
    sheet_data <- read.xlsx(summary_file, sheet = country, check.names = FALSE)
  } else {
    sheet_data <- data.frame(DEFINITION = psr_list$DEFINITION,
                             stringsAsFactors = FALSE)
  }
  
  # ensure the full DEFINITION list is present
  if (!all(psr_list$DEFINITION %in% sheet_data$DEFINITION)) {
    sheet_data <- full_join(psr_list["DEFINITION"], sheet_data,
                            by = "DEFINITION")
  }
  
  # ensure every year column exists
  for (yr in years_all) {
    if (!as.character(yr) %in% names(sheet_data)) {
      sheet_data[[as.character(yr)]] <- NA_real_
    }
  }
  sheet_list[[country]] <- sheet_data
}

## 4b) Merge per-year files into memory --------------------------------------
for (country in eic_data$Country) {
  yrs <- years_by_cntry[[country]]
  for (yr in yrs) {
    f <- sprintf("Inst_Cap_Type_%s_%d.xlsx", country, yr)
    if (!file.exists(f)) next
    
    year_tbl <- read.xlsx(f, sheet = 1, check.names = FALSE)
    
    sheet_list[[country]] <- sheet_list[[country]] %>%
      left_join(year_tbl[c("DEFINITION", "quantity")], by = "DEFINITION") %>%
      mutate(!!as.character(yr) := coalesce(quantity, !!sym(as.character(yr)))) %>%
      select(-quantity)
  }
}

## 4c) Write all sheets back -------------------------------------------------
for (sht in sheets(wb)) removeWorksheet(wb, sht)

for (country in names(sheet_list)) {
  addWorksheet(wb, country)
  writeData(wb, sheet = country, sheet_list[[country]])
}

saveWorkbook(wb, summary_file, overwrite = TRUE)
cat("Consolidated workbook saved to ", summary_file, "\n")
