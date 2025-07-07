###############################################################################
# Download ENTSO-E actual generation per production type and save to Excel  #
###############################################################################

# ── Libraries ── #
library(entsoeapi)
library(openxlsx)
library(dplyr)
library(tidyr)
library(lubridate)

# ── Token ── #
token <- Sys.getenv("ENTSOE_PAT")
if (nchar(token) != 36)
  stop("Set a valid 36-character ENTSOE_PAT environment variable.")

# ── Country ➜ EIC + TZ ── #
eic_data <- data.frame(
  Country = c("AL","AT","BE","BA","BG","HR","CY","CZ","DK","EE","FI","FR","GE",
              "DE","GR","HU","IE","IT","XK","LV","LT","LU","MD","ME","NL","MK",
              "NO","PL","PT","RO","RS","SK","SI","ES","SE","UK","UA","CH"),
  EIC = c("10YAL-KESH-----5","10YAT-APG------L","10YBE----------2","10YBA-JPCC-----D",
          "10YCA-BULGARIA-R","10YHR-HEP------M","10YCY-1001A0003J","10YCZ-CEPS-----N",
          "10Y1001A1001A65H","10Y1001A1001A39I","10YFI-1--------U","10YFR-RTE------C",
          "10Y1001A1001B012","10Y1001A1001A83F","10YGR-HTSO-----Y","10YHU-MAVIR----U",
          "10YIE-1001A00010","10YIT-GRTN-----B","10Y1001C--00100H","10YLV-1001A00074",
          "10YLT-1001A0008Q","10YLU-CEGEDEL-NQ","10Y1001A1001A990","10YCS-CG-TSO---S",
          "10YNL----------L","10YMK-MEPSO----8","10YNO-0--------C","10YPL-AREA-----S",
          "10YPT-REN------W","10YRO-TEL------P","10YCS-SERBIATSOV","10YSK-SEPS-----K",
          "10YSI-ELES-----O","10YES-REE------0","10YSE-1--------K","10Y1001A1001A92E",
          "10Y1001C--00003F","10YCH-SWISSGRIDZ"),
  Timezone = c("CET","CET","CET","CET","EET","CET","EET","CET","CET","EET","EET","CET","UTC+4",
               "CET","EET","CET","UTC","CET","CET","EET","EET","CET","EET","CET","CET","CET",
               "CET","CET","UTC","EET","CET","CET","CET","CET","CET","UTC","EET","CET"),
  stringsAsFactors = FALSE
)

# ── PSR label ➜ code ── #
psr_code_map <- c(
  "Wind Onshore"    = "B19",
  "Wind Offshore"   = "B18",
  "Solar"           = "B16",
  "Hydro Run-of-river and poundage" = "B12",
  "Hydro Water Reservoir"           = "B11",
  "Hydro Pumped Storage"            = "B21",
  "Marine"          = "B23",
  "Biomass"         = "B10",
  "Nuclear"         = "B09",
  "Geothermal"      = "B17",
  "Fossil Peat"           = "B08",
  "Fossil Brown coal/Lignite"       = "B02",
  "Fossil Coal-derived gas"         = "B03",
  "Fossil Gas"                      = "B04",
  "Fossil Hard coal"                = "B05",
  "Fossil Oil"                      = "B06",
  "Fossil Oil shale"                = "B07",
  "Other renewable" = "B02",
  "Waste"           = "B15"
)


# ── API helper ── #
safe_gen_per_prod_type <- function(eic, start_date, end_date, psr_code) {
  rows <- list(); st <- start_date
  while (st < end_date) {
    en <- min(st + months(1), end_date)
    part <- tryCatch(
      gen_per_prod_type(eic, period_start = st, period_end = en,
                        gen_type = psr_code, tidy_output = TRUE),
      error = function(e) {
        message("⚠️ ", eic, " ", format(st, "%Y-%m"), ": ",
                sub("[\r\n].*$", "", e$message))
        NULL
      }
    )
    if (!is.null(part)) rows[[length(rows)+1]] <- collect(part)
    Sys.sleep(0.2); st <- en
  }
  if (!length(rows)) return(NULL)
  df <- bind_rows(rows)
  
  # 1. exact PSR
  psr_cols <- intersect(names(df), c("psrType","ts_mkt_psr_type"))
  if (length(psr_cols)>0) df <- df[df[[psr_cols[1]]] == psr_code, ]
  
  # 2. bidding zone & drop cross-border
  in_cols  <- intersect(names(df), c("inbiddingzone_domain_mrid","inBiddingZone_Domain.mRID"))
  if (length(in_cols)>0) {
    df <- df[df[[in_cols[1]]] == eic, ]
    out_cols <- intersect(names(df), c("outbiddingzone_domain_mrid","outBiddingZone_Domain.mRID"))
    if (length(out_cols)>0) {
      df <- df[is.na(df[[out_cols[1]]]) | df[[out_cols[1]]] == df[[in_cols[1]]], ]
    }
  }
  
  df
}

# ── Main worker ── #
fetch_and_save_entsoe_data <- function(psr_label, folder, prefix, years=2024:2025) {
  psr_code <- psr_code_map[[psr_label]]
  dir.create(folder, recursive=TRUE, showWarnings=FALSE)
  
  for (i in seq_len(nrow(eic_data))) {
    cn  <- eic_data$Country[i]
    eic <- eic_data$EIC[i]
    tz  <- eic_data$Timezone[i]
    
    # compute offset if needed
    is_offset <- grepl("^UTC[+-]\\d+$", tz)
    if (is_offset) {
      off_h <- as.integer(sub("UTC([+-]\\d+)", "\\1", tz))
    }
    
    for (yr in years) {
      # 1) build st/en in UTC, then shift if offset
      st_utc <- as.POSIXct(paste0(yr,"-01-01 00:00:00"), tz="UTC")
      en_utc <- as.POSIXct(paste0(yr+1,"-01-01 00:00:00"), tz="UTC")
      if (is_offset) {
        st <- st_utc + hours(off_h)
        en <- en_utc + hours(off_h)
      } else {
        st <- as.POSIXct(paste0(yr,"-01-01 00:00:00"), tz=tz)
        en <- as.POSIXct(paste0(yr+1,"-01-01 00:00:00"), tz=tz)
      }
      
      raw <- safe_gen_per_prod_type(eic, st, en, psr_code)
      
      if (!is.null(raw) && nrow(raw)>0) {
        dt_col  <- intersect(names(raw), c("datetime","ts_point_dt_start","timestamp","ts_start"))[1]
        qty_col <- intersect(names(raw), c("quantity","ts_point_quantity","value"))[1]
        
        if (!is.na(dt_col) && !is.na(qty_col)) {
          # convert to UTC instants
          df <- raw %>% mutate(datetime_utc = force_tz(.data[[dt_col]], tzone="UTC"))
          
          # shift to local
          if (is_offset) {
            df <- df %>% mutate(datetime_local = datetime_utc + hours(off_h))
          } else {
            df <- df %>% mutate(datetime_local = with_tz(datetime_utc, tzone=tz))
          }
          
          # pick first quantity each hour
          api <- df %>%
            distinct(datetime_local, .keep_all=TRUE) %>%
            mutate(hour = floor_date(datetime_local, "hour")) %>%
            group_by(hour) %>%
            summarise(ts_point_quantity=first(.data[[qty_col]]), .groups="drop") %>%
            transmute(timestamp=hour, ts_point_quantity)
        } else {
          warning("Missing datetime/quantity cols for ",psr_label," in ",cn," ",yr)
          api <- NULL
        }
      } else {
        api <- NULL
      }
      
      # full hourly grid + join
      hours <- tibble(timestamp=seq(floor_date(st,"hour"), en-hours(1), by="hour"))
      tidy  <- if (!is.null(api)) {
        hours %>%
          left_join(api, by="timestamp") %>%
          replace_na(list(ts_point_quantity=0))
      } else {
        hours %>% mutate(ts_point_quantity=0)
      }
      tidy <- tidy %>% mutate(psr_type=psr_label, index=row_number())
      
      # write
      file <- file.path(folder, sprintf("%s%s_%d.xlsx", prefix, cn, yr))
      write.xlsx(tidy, file, sheetName="sheet", rowNames=FALSE)
      cat("✔",psr_label,"|",cn,yr,"→",basename(file),"\n")
      Sys.sleep(0.5)
    }
  }
}

# ── Define & run ── #
gen_types <- list(
  list("Wind Onshore","entsoe_generation_windon","Actual_Gen_WindOn_"),
  list("Wind Offshore","entsoe_generation_windoff","Actual_Gen_WindOff_"),
  list("Solar","entsoe_generation_solarpv","Actual_Gen_Solar_"),
  list("Hydro Run-of-river and poundage","entsoe_generation_hydro_runofriver_poundage","Actual_Gen_Hydro_RP_"),
  list("Hydro Water Reservoir","entsoe_generation_hydro_water_reservoir","Actual_Gen_Hydro_WR_"),
  list("Hydro Pumped Storage","entsoe_generation_hydro_pumped_storage","Actual_Gen_Hydro_PS_"),
  list("Marine","entsoe_generation_marine","Actual_Gen_Marine_"),
  list("Biomass","entsoe_generation_biomass","Actual_Gen_Biomass_"),
  list("Nuclear","entsoe_generation_nuclear","Actual_Gen_Nuclear_"),
  list("Geothermal","entsoe_generation_geothermal","Actual_Gen_Geothermal_"),
  list("Fossil Peat",               "entsoe_generation_Other", "Actual_Gen_Fossil Peat_"),
  list("Fossil Brown coal/Lignite", "entsoe_generation_Other", "Actual_Gen_Fossil_Brown_Coal_Lignite"),
  list("Fossil Coal-derived gas",   "entsoe_generation_Other", "Actual_Gen_Fossil_Coal-derived_gas_"),
  list("Fossil Gas",                "entsoe_generation_Other", "Actual_Gen_Fossil_Gas_"),
  list("Fossil Hard coal",          "entsoe_generation_Other", "Actual_Gen_Fossil_Hard_coal_"),
  list("Fossil Oil",                "entsoe_generation_Other", "Actual_Gen_Fossil_Oil_"),
  list("Fossil Oil shale",          "entsoe_generation_Other", "Actual_Gen_Fossil_Oil_shale_"),
  list("Other renewable","entsoe_generation_other_renewable","Actual_Gen_Other_renewable_"),
  list("Waste","entsoe_generation_waste","Actual_Gen_Waste_")
)

for (gt in gen_types) {
  fetch_and_save_entsoe_data(
    psr_label = gt[[1]],
    folder    = gt[[2]],
    prefix    = gt[[3]],
    years     = 2020:2025
  )
}

