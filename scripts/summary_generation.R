###############################################################################
# Create summary workbooks from single (per‐country/year) Excel files #
###############################################################################

# ── Libraries ── #
library(readxl)   
library(writexl)
library(dplyr)
library(lubridate)

safe_read <- function(path) {
  tryCatch(
    read_excel(path, sheet = 1),
    error   = function(e) { warning("read: ", basename(path), ": ", e$message); NULL },
    warning = function(w) { warning("read: ", basename(path), ": ", w$message); NULL }
  )
}

# ── Configuration ── #

# PSR-type folders on disk
psr_folders <- c(
  "entsoe_generation_windon",
  "entsoe_generation_windoff",
  "entsoe_generation_solarpv",
  "entsoe_generation_hydro_runofriver_poundage",
  "entsoe_generation_hydro_water_reservoir",
  "entsoe_generation_hydro_pumped_storage",
  "entsoe_generation_marine",
  "entsoe_generation_biomass",
  "entsoe_generation_nuclear",
  "entsoe_generation_geothermal",
  "entsoe_generation_other",           # contains all fossil types
  "entsoe_generation_other_renewable",
  "entsoe_generation_waste"
)

years     <- 2020:2025
max_hours <- 8784

# two-letter country codes in your file names
countries <- c("AL","AT","BE","BA","BG","HR","CY","CZ","DK","EE","FI","FR","GE",
               "DE","GR","HU","IE","IT","XK","LV","LT","LU","MD","ME","NL","MK",
               "NO","PL","PT","RO","RS","SK","SI","ES","SE","UK","UA","CH")

# ── Main loop ── #
for (folder in psr_folders) {
  if (!dir.exists(folder)) {
    warning("Skipping missing folder: ", folder)
    next
  }
  message("Scanning folder: ", folder)
  
  # grab every .xlsx in that folder
  all_files <- list.files(folder, pattern="\\.xlsx$", full.names=TRUE)
  if (length(all_files)==0) next
  
  # extract metadata from filename
  meta <- tibble(path = all_files) %>%
    mutate(fn = basename(path)) %>%
    # prefix = everything before "_<CC>_<YYYY>.xlsx"
    mutate(
      prefix = sub("_[A-Z]{2}_[0-9]{4}\\.xlsx$", "", fn),
      Country = sub("^.*_([A-Z]{2})_[0-9]{4}\\.xlsx$", "\\1", fn),
      Year    = as.integer(sub("^.*_([0-9]{4})\\.xlsx$", "\\1", fn))
    ) %>%
    filter(Country %in% countries, Year %in% years)
  
  # process each PSR prefix separately
  for (psr in unique(meta$prefix)) {
    sub_meta <- filter(meta, prefix==psr)
    message(" ⋯ building summary for PSR type: ", psr)
    
    # prepare list of one data.frame per country
    summary_list <- vector("list", length(countries))
    names(summary_list) <- countries
    
    for (cty in countries) {
      # skeleton: Hour + one column per year
      df_sum <- tibble(Hour = seq_len(max_hours))
      for (yr in years) df_sum[[as.character(yr)]] <- NA_real_
      
      # fill in year by year
      rows <- filter(sub_meta, Country==cty)
      for (yr in years) {
        row <- filter(rows, Year==yr)
        if (nrow(row)==0) next
        dat <- safe_read(row$path[1])
        if (is.null(dat) || nrow(dat)==0) next
        
        # detect timestamp & quantity
        ts_col  <- intersect(names(dat), c("timestamp","datetime","ts_point_dt_start","start"))[1]
        qty_col <- intersect(names(dat), c("ts_point_quantity","quantity","value"))[1]
        if (is.na(ts_col)||is.na(qty_col)) {
          warning(" missing cols in ", basename(row$path[1])); next
        }
        
        d2 <- dat %>%
          rename(time = !!ts_col, data = !!qty_col) %>%
          arrange(time) %>%
          mutate(Hour = row_number()) %>%
          filter(Hour <= max_hours) %>%
          select(Hour, data)
        
        df_sum[ d2$Hour, as.character(yr) ] <- d2$data
      }
      
      summary_list[[cty]] <- df_sum
    }
    
    # write one summary workbook
    out <- file.path(folder, paste0(psr, "_summary.xlsx"))
    writexl::write_xlsx(summary_list, path = out)
    message("   ✔ written: ", out)
  }
}
