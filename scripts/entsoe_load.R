# Load library 
library(entsoeapi)
library(openxlsx)
library(dplyr)
library(tidyr)
library(lubridate)

security_token <- Sys.getenv("ENTSOE_PAT")

# Initiate/create the dataset
eic_data <- data.frame(
  Country = c("AL", "AT", "BE", "BA", "BG", "HR", "CY", "CZ", "DK", "EE", "FI", "FR", "GE", "DE", "GR", "HU", "IE", "IT", "XK", "LV", "LT", "LU", "MD", "ME", "NL", "MK", "NO", "PL", "PT", "RO", "RS", "SK", "SI", "ES", "SE", "UK", "UA", "CH"),
  EIC = c("10YAL-KESH-----5", "10YAT-APG------L", "10YBE----------2", "10YBA-JPCC-----D", "10YCA-BULGARIA-R", "10YHR-HEP------M", "10YCY-1001A0003J", "10YCZ-CEPS-----N", "10Y1001A1001A65H", "10Y1001A1001A39I", "10YFI-1--------U", "10YFR-RTE------C", "10Y1001A1001B012", "10Y1001A1001A83F", "10YGR-HTSO-----Y", "10YHU-MAVIR----U", "10YIE-1001A00010", "10YIT-GRTN-----B", "10Y1001C--00100H", "10YLV-1001A00074", "10YLT-1001A0008Q", "10YLU-CEGEDEL-NQ", "10Y1001A1001A990", "10YCS-CG-TSO---S", "10YNL----------L", "10YMK-MEPSO----8", "10YNO-0--------C", "10YPL-AREA-----S", "10YPT-REN------W", "10YRO-TEL------P", "10YCS-SERBIATSOV", "10YSK-SEPS-----K", "10YSI-ELES-----O", "10YES-REE------0", "10YSE-1--------K", "10Y1001A1001A92E",  "10Y1001C--00003F", "10YCH-SWISSGRIDZ"),
  Timezone = c("CET", "CET", "CET", "CET", "EET", "CET", "EET", "CET", "CET", "EET", "EET", "CET",
               "Asia/Tbilisi", "CET", "EET", "CET", "UTC", "CET", "CET", "EET", "EET", "CET", "EET", "CET",
               "CET", "CET", "CET", "CET", "UTC", "EET", "CET", "CET", "CET", "CET", "CET",
               "UTC", "EET", "CET"))

# Asset types
psr_list<-asset_types
psr_list<-psr_list[-c(1:5, 26:29),]

for (i in 1:nrow(eic_data)) {
  country <- eic_data$Country[i]
  eic <- eic_data$EIC[i]
  tz <- eic_data$Timezone[i]
  
  for (year in 2020:2025) {
    start_date <- as.POSIXct(paste(year, "-01-01 00:00:00", sep = ""), tz)
    end_date <- as.POSIXct(paste(year + 1, "-01-01 00:00:00", sep = ""), tz)
    
    actual_total_load <- load_actual_total(eic, start_date, end_date)

  timestamp_df <- data.frame(timestamp = seq(from = start_date, to = end_date, by = "hour"))

  timestamp_df$timestamp <- force_tz(as.POSIXct(timestamp_df$timestamp), tz)

  if (!is.null(actual_total_load) && nrow(actual_total_load) > 0 &&"ts_point_dt_start" %in% names(actual_total_load)) {
    actual_total_load <- left_join(timestamp_df, actual_total_load, by = c("timestamp" = "ts_point_dt_start")) %>%
      rename("quantity" = ts_point_quantity)%>%
      select(timestamp, quantity)
  } else{
    actual_total_load<-data.frame(
      timestamp=timestamp_df$timestamp,
      quantity=NA_real_
    )
  }

    file_name <- paste("Total_Load_", country, "_", year, ".xlsx", sep = "")

    write.xlsx(actual_total_load, file_name, sheetName="sheet", rowNames = FALSE)

    cat("Data for", country, "and year", year, "written to", file_name, "\n")
  }
}


#-----without timestamp df----

for (i in 1:nrow(eic_data)) {
  country <- eic_data$Country[i]
  eic <- eic_data$EIC[i]
  tz <- eic_data$Timezone[i]
  
  for (year in 2020:2025) {
    start_date <- as.POSIXct(paste(year, "-01-01 00:00:00", sep = ""), tz)
    end_date <- as.POSIXct(paste(year + 1, "-01-01 00:00:00", sep = ""), tz)
    
    actual_total_load <- load_actual_total(eic, start_date, end_date)

    if (!is.null(actual_total_load) && nrow(actual_total_load) > 0 &&"ts_point_dt_start" %in% names(actual_total_load)) {
      actual_total_load <- actual_total_load %>%
        rename(quantity = ts_point_quantity, timestamp = ts_point_dt_start) %>%
        select(timestamp, quantity)
    } else{
      actual_total_load<-data.frame(
        timestamp=timestamp_df$timestamp,
        quantity=NA_real_
      )
    }
    
    file_name <- paste("Total_Load_", country, "_", year, ".xlsx", sep = "")
    
    write.xlsx(actual_total_load, file_name, sheetName="sheet", rowNames = FALSE)
    
    cat("Data for", country, "and year", year, "written to", file_name, "\n")
  }
}

#---#loading summary file----
file_path <- "load_entsoe.xlsx"
years <- 2020:2025

max_hours <- 8784

# --------------------------------------------------------------------------
# 1) CREATE OR OPEN THE FINAL SUMMARY WORKBOOK
# --------------------------------------------------------------------------

if (file.exists(file_path)) {
  wb <- loadWorkbook(file_path)
  cat("Workbook loaded:", file_path, "\n")
} else {
  wb <- createWorkbook()
  saveWorkbook(wb, file_path, overwrite = TRUE)
  wb <- loadWorkbook(file_path)
  cat("Workbook created and saved:", file_path, "\n")
}

# --------------------------------------------------------------------------
# 2) LOOP OVER COUNTRIES AND PREPARE EACH SHEET
# --------------------------------------------------------------------------

for (i in seq_len(nrow(eic_data))) {
  country <- eic_data$Country[i]
  
  # If the sheet doesn't exist, create it from scratch
  if (!(country %in% names(wb))) {
    addWorksheet(wb, country)
    
    # Initialize data frame with Hour = 1..8784 and empty columns for each year
    sheet_data <- data.frame(Hour = 1:max_hours)
    for (yr in years) {
      sheet_data[[as.character(yr)]] <- NA_real_
    }
    
    # Write to the workbook
    writeData(wb, country, sheet_data)
    saveWorkbook(wb, file_path, overwrite = TRUE)
    cat("Created sheet for", country, "\n")
    
  } else {
    # Read existing sheet data
    sheet_data <- read.xlsx(file_path, sheet = country)
    
    # If there's no "Hour" column or the length is wrong, rebuild it robustly
    # We'll ensure we have exactly 1..8784 rows with a "Hour" column
    if (!("Hour" %in% names(sheet_data))) {
      # Create a fresh structure
      new_data <- data.frame(Hour = 1:max_hours)
      # If any year columns existed, they won't match row counts; rebuild them as NA
      for (yr in years) {
        if (!yr %in% names(sheet_data)) {
          new_data[[as.character(yr)]] <- NA_real_
        } else {
          # If the old sheet_data has fewer or more rows, we can't reliably merge;
          # so we keep it simple and set them to NA
          new_data[[as.character(yr)]] <- NA_real_
        }
      }
      sheet_data <- new_data
    } else {
      # Ensure the data frame has exactly max_hours rows (1..8784)
      # If the existing sheet had fewer rows, we expand; if more, we truncate
      old_rows <- nrow(sheet_data)
      if (old_rows < max_hours) {
        # Add more rows
        extra_df <- data.frame(Hour = (old_rows+1):max_hours)
        for (yr in years) {
          if (!(as.character(yr) %in% names(sheet_data))) {
            sheet_data[[as.character(yr)]] <- NA_real_
          }
          extra_df[[as.character(yr)]] <- NA_real_
        }
        sheet_data <- rbind(sheet_data, extra_df)
        
      } else if (old_rows > max_hours) {
        # Truncate
        sheet_data <- sheet_data[1:max_hours, ]
      }
      
      # Now ensure we have columns for each year
      for (yr in years) {
        if (!(as.character(yr) %in% names(sheet_data))) {
          sheet_data[[as.character(yr)]] <- NA_real_
        }
      }
      
      # Also ensure "Hour" goes from 1..max_hours in order
      sheet_data$Hour <- 1:max_hours
    }
  }
  
  # ------------------------------------------------------------------------
  # 3) FOR EACH YEAR, READ THE SINGLE-YEAR FILE (IF IT EXISTS)
  #    AND MERGE INTO 'sheet_data' BY HOUR (1..N).
  # ------------------------------------------------------------------------
  for (yr in years) {
    single_year_file <- paste0("Total_Load_", country, "_", yr, ".xlsx")
    
    if (file.exists(single_year_file)) {
      load_data <- read.xlsx(single_year_file, sheet = 1)
      
      if (!("quantity" %in% names(load_data)) && ("ts_point_quantity" %in% names(load_data))) {
        load_data <- load_data %>%
          rename(quantity = ts_point_quantity)
      }
      if (!("timestamp" %in% names(load_data)) && ("ts_point_dt_start" %in% names(load_data))) {
        load_data <- load_data %>%
          rename(timestamp = ts_point_dt_start)
      }
      
      # Sort by timestamp 
      if ("timestamp" %in% names(load_data)) {
        load_data$timestamp <- as.POSIXct(load_data$timestamp, tz = "UTC")
        load_data <- load_data[order(load_data$timestamp), ]
      }
      
      # Create an hour index 1..nrow(load_data)
      load_data$hour_index <- seq_len(nrow(load_data))
      
      # Now fill the summary’s column for that year
      # If the single-year data has fewer than 'max_hours' rows, the rest remain NA
      # If it has more, we only take the first 'max_hours' anyway
      fill_count <- min(nrow(load_data), max_hours)
      
      sheet_data[1:fill_count, as.character(yr)] <- load_data$quantity[1:fill_count]
      
      cat("Merged", single_year_file, "into", country, "sheet for year", yr, "\n")
    } else {
      cat("File not found:", single_year_file, "– skipping year", yr, "\n")
    }
  }

  # ── Trend column ─────────────────────────────────────────
  trend_fun <- function(x) {
    v <- as.numeric(x)
    v <- v[!is.na(v)]                  # drop NAs
    if (length(v) == 0) return("")     # no data
    
    if (all(v == 0)) return("")        
    
    # strip leading zeros so they don’t spoil monotone tests
    nz <- which(v != 0)
    if (length(nz)) v <- v[min(nz):length(v)]
    
    if (length(v) < 2) return("")      # still not enough points
    
    # “No change”: all equal  OR  last‑3 equal
    if (length(unique(v)) == 1)                     return("No change")
    if (length(v) >= 3 && length(unique(tail(v, 3))) == 1)
      return("No change")
    
    d <- diff(v)                       # year‑to‑year deltas
    if (all(d >= 0) && any(d > 0))     return("increasing")   
    if (all(d <= 0) && any(d < 0))     return("decreasing") 
    "fluctuating"
  }
  
  year_cols <- intersect(as.character(years), names(sheet_data))
  sheet_data$Trend <- apply(
    sheet_data[, year_cols, drop = FALSE],
    1,
    trend_fun
  )
  
  # ------------------------------------------------------------------------
  # 4) SAVE THE UPDATED SHEET BACK INTO THE WORKBOOK
  # ------------------------------------------------------------------------
  # Remove old sheet, then write new one
  if (country %in% names(wb)) {
    removeWorksheet(wb, country)
  }
  addWorksheet(wb, country)
  writeData(wb, sheet = country, sheet_data)
  
  saveWorkbook(wb, file_path, overwrite = TRUE)
  cat("Updated sheet for", country, "\n")
}

# --------------------------------------------------------------------------
# 5) FINAL SAVE
# --------------------------------------------------------------------------
saveWorkbook(wb, file_path, overwrite = TRUE)
cat("Final file saved as", file_path, "\n")
