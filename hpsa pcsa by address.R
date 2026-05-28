# =============================================================================
# Primary Care HPSA + PCSA Lookup from Excel Address List
# METHOD: Geocodes address → Census Tract FIPS → joins to:
#           1. HRSA "All HPSAs" XLSX  (HPSA score, type, status)
#           2. HRSA PCSA XLSX         (PCSA ID, name, provider ratio)
# =============================================================================

# --- 0. Install/load packages ------------------------------------------------
pkgs <- c("readxl", "writexl", "httr", "jsonlite", "dplyr", "stringr")
invisible(lapply(pkgs, function(p) {
  if (!requireNamespace(p, quietly = TRUE)) install.packages(p)
  library(p, character.only = TRUE)
}))

# =============================================================================
# 1. CONFIGURATION — edit these lines
# =============================================================================
INPUT_FILE  <- "C:/Users/anguyen/Downloads/rstudio scripts/addresses_cleaned.xlsx"
ADDRESS_COL <- "address"
OUTPUT_FILE <- "C:/Users/anguyen/Downloads/rstudio scripts/hpsa_pcsa_results.xlsx"

# HRSA "All HPSAs" XLSX — Primary Care
# Download from: https://data.hrsa.gov/data/download
# Under: Health Workforce > Shortage Areas > HPSA Primary Care > "All HPSAs" XLSX
HPSA_XLSX <- "C:/Users/anguyen/Downloads/rstudio scripts/hrsa_hpsa_data/BCD_HPSA_FCT_DET_PC.xlsx"

# California PCSA CSV
PCSA_CSV <- "C:/Users/anguyen/Downloads/rstudio scripts/hrsa_hpsa_data/Primary_Care_Shortage_Area_(PCSA).csv"

# California Healthcare Workforce Geography Crosswalk (Census Tract → MSSA)
CROSSWALK_CSV <- "C:/Users/anguyen/Downloads/rstudio scripts/hrsa_hpsa_data/geography-crosswalk.csv"

# Census Geocoding API (geographies endpoint returns tract + county FIPS)
CENSUS_GEO_URL <- "https://geocoding.geo.census.gov/geocoder/geographies/onelineaddress"

# =============================================================================
# 2. LOAD & PREPARE HRSA HPSA DATA
# =============================================================================
message("Loading HRSA HPSA data...")
hpsa_raw <- read_excel(HPSA_XLSX)

hpsa_df <- hpsa_raw %>%
  filter(`HPSA Status` == "Designated") %>%
  mutate(
    county_fips = str_pad(
      as.character(`State and County Federal Information Processing Standard Code`),
      5, pad = "0"
    )
  ) %>%
  select(
    county_fips,
    hpsa_name     = `HPSA Name`,
    hpsa_score    = `HPSA Score`,
    hpsa_type     = `Designation Type`,
    hpsa_status   = `HPSA Status`,
    hpsa_pop_type = `HPSA Population Type`,
    hpsa_degree   = `HPSA Degree of Shortage`,
    pc_mcta_score = `PC MCTA Score`
  )

message(sprintf("Loaded %d designated Primary Care HPSA records.", nrow(hpsa_df)))

# =============================================================================
# 3. LOAD & PREPARE CALIFORNIA PCSA DATA
# =============================================================================
message("Loading California PCSA data...")
pcsa_raw <- read.csv(PCSA_CSV, stringsAsFactors = FALSE)

message("PCSA file columns: ", paste(names(pcsa_raw), collapse = ", "))

# Detect census tract and PCSA score columns automatically
tract_col <- names(pcsa_raw)[str_detect(names(pcsa_raw), regex("census.?tract|tractfips|geoid|tract", ignore_case = TRUE))][1]
score_col <- names(pcsa_raw)[str_detect(names(pcsa_raw), regex("pcsa.?score|score|pcsa", ignore_case = TRUE))][1]
mssa_col  <- names(pcsa_raw)[str_detect(names(pcsa_raw), regex("mssa", ignore_case = TRUE))][1]

message("Tract col: ", tract_col)
message("Score col: ", score_col)
message("MSSA col:  ", mssa_col)

pcsa_df <- pcsa_raw %>%
  select(
    mssa_id         = MSSA_ID,
    mssa_name       = MSSA_NAME,
    pcsa_designated = PCSA,
    pcsa_score      = Score_Tota,
    score_provider  = Score_Prov,
    score_poverty   = Score_Pove,
    provider_ratio  = Provider_R,
    ca_county_fips  = CNTY_FIPS
  ) %>%
  mutate(
    # Convert CA county FIPS (e.g. 59) to match last 3 digits of federal FIPS (e.g. 06059)
    ca_county_fips = as.integer(ca_county_fips)
  )

message(sprintf("Loaded %d PCSA MSSA records.", nrow(pcsa_df)))


#loading in crosswalk data
crosswalk_raw <- read.csv(CROSSWALK_CSV, stringsAsFactors = FALSE)
crosswalk_df <- crosswalk_raw %>%
  select(tract_fips = Census_Tract, mssa_id = MSSA_ID) %>%
  mutate(tract_fips = str_pad(as.character(tract_fips), 11, pad = "0")) %>%
  distinct(tract_fips, .keep_all = TRUE)
nrow(crosswalk_df)

# =============================================================================
# 4. GEOCODING FUNCTION — returns lat/lon + county FIPS + census tract FIPS
# =============================================================================
geocode_with_geographies <- function(address) {
  resp <- tryCatch(
    GET(CENSUS_GEO_URL, query = list(
      address   = address,
      benchmark = "Public_AR_Current",
      vintage   = "Current_Current",
      format    = "json"
    ), timeout(20)),
    error = function(e) NULL
  )
  
  empty <- list(
    lat             = NA_real_,
    lon             = NA_real_,
    matched_address = NA_character_,
    county_fips     = NA_character_,
    county_name     = NA_character_,
    tract_fips      = NA_character_
  )
  
  if (is.null(resp) || http_error(resp)) return(empty)
  
  body    <- fromJSON(content(resp, "text", encoding = "UTF-8"), simplifyVector = FALSE)
  matches <- body$result$addressMatches
  if (length(matches) == 0) return(empty)
  
  m      <- matches[[1]]
  coords <- m$coordinates
  geos   <- m$geographies
  
  # County FIPS (5-digit)
  county_fips <- tryCatch({
    counties <- geos[["Counties"]]
    if (length(counties) > 0) as.character(counties[[1]]$GEOID) else NA_character_
  }, error = function(e) NA_character_)
  
  county_name <- tryCatch({
    counties <- geos[["Counties"]]
    if (length(counties) > 0) as.character(counties[[1]]$NAME) else NA_character_
  }, error = function(e) NA_character_)
  
  # Census Tract FIPS (11-digit: state 2 + county 3 + tract 6)
  tract_fips <- tryCatch({
    tracts <- geos[["Census Tracts"]]
    if (length(tracts) > 0) as.character(tracts[[1]]$GEOID) else NA_character_
  }, error = function(e) NA_character_)
  
  list(
    lat             = as.numeric(coords$y),
    lon             = as.numeric(coords$x),
    matched_address = m$matchedAddress,
    county_fips     = county_fips,
    county_name     = county_name,
    tract_fips      = tract_fips
  )
}

# =============================================================================
# 5. HPSA LOOKUP — by county FIPS
# =============================================================================
lookup_hpsa_by_county <- function(county_fips, hpsa_data) {
  no_match <- list(
    pc_hpsa_designated = FALSE,
    pc_hpsa_score      = NA_real_,
    pc_hpsa_type       = NA_character_,
    pc_hpsa_name       = NA_character_,
    pc_hpsa_pop_type   = NA_character_,
    pc_hpsa_degree     = NA_real_,
    pc_mcta_score      = NA_real_,
    pc_hpsa_status     = "Not Designated"
  )
  
  if (is.na(county_fips)) return(no_match)
  
  matches <- hpsa_data %>% filter(county_fips == !!county_fips)
  if (nrow(matches) == 0) return(no_match)
  
  best <- matches %>% arrange(desc(hpsa_score)) %>% slice(1)
  
  list(
    pc_hpsa_designated = TRUE,
    pc_hpsa_score      = as.numeric(best$hpsa_score),
    pc_hpsa_type       = as.character(best$hpsa_type),
    pc_hpsa_name       = as.character(best$hpsa_name),
    pc_hpsa_pop_type   = as.character(best$hpsa_pop_type),
    pc_hpsa_degree     = as.numeric(best$hpsa_degree),
    pc_mcta_score      = as.numeric(best$pc_mcta_score),
    pc_hpsa_status     = "Designated"
  )
}

# =============================================================================
# 6. PCSA LOOKUP — exact match via census tract → MSSA → PCSA
# =============================================================================
lookup_pcsa_by_tract <- function(tract_fips, crosswalk, pcsa_data) {
  no_match <- list(
    pcsa_designated = NA_character_,
    pcsa_score      = NA_real_,
    mssa_id         = NA_character_,
    mssa_name       = NA_character_,
    provider_ratio  = NA_real_
  )
  
  if (is.na(tract_fips)) return(no_match)
  
  # Step 1: tract FIPS → MSSA ID via crosswalk
  cw <- crosswalk %>% filter(tract_fips == !!tract_fips)
  if (nrow(cw) == 0) return(no_match)
  
  mssa <- cw$mssa_id[1]
  
  # Step 2: MSSA ID → PCSA record
  match <- pcsa_data %>% filter(mssa_id == mssa)
  if (nrow(match) == 0) return(no_match)
  
  list(
    pcsa_designated = as.character(match$pcsa_designated[1]),
    pcsa_score      = as.numeric(match$pcsa_score[1]),
    mssa_id         = as.character(match$mssa_id[1]),
    mssa_name       = as.character(match$mssa_name[1]),
    provider_ratio  = as.numeric(match$provider_ratio[1])
  )
}

# =============================================================================
# 7. MAIN PROCESSING LOOP
# =============================================================================
message("Reading input file: ", INPUT_FILE)
df_in <- read_excel(INPUT_FILE)

if (!ADDRESS_COL %in% names(df_in)) {
  stop("Column '", ADDRESS_COL, "' not found. ",
       "Available: ", paste(names(df_in), collapse = ", "))
}

addresses <- df_in[[ADDRESS_COL]]
n         <- length(addresses)
message(sprintf("Processing %d addresses...\n", n))

results <- vector("list", n)

for (i in seq_len(n)) {
  addr <- addresses[i]
  message(sprintf("[%d/%d] %s", i, n, addr))
  
  geo  <- geocode_with_geographies(addr)
  Sys.sleep(0.5)
  
  hpsa <- lookup_hpsa_by_county(geo$county_fips, hpsa_df)
  pcsa <- lookup_pcsa_by_tract(geo$tract_fips, crosswalk_df, pcsa_df)
  
  results[[i]] <- tibble(
    original_address   = addr,
    matched_address    = geo$matched_address,
    latitude           = geo$lat,
    longitude          = geo$lon,
    county_name        = geo$county_name,
    county_fips        = geo$county_fips,
    census_tract_fips  = geo$tract_fips,
    # HPSA fields
    pc_hpsa_designated = hpsa$pc_hpsa_designated,
    pc_hpsa_score      = hpsa$pc_hpsa_score,
    pc_hpsa_type       = hpsa$pc_hpsa_type,
    pc_hpsa_name       = hpsa$pc_hpsa_name,
    pc_hpsa_pop_type   = hpsa$pc_hpsa_pop_type,
    pc_hpsa_degree     = hpsa$pc_hpsa_degree,
    pc_mcta_score      = hpsa$pc_mcta_score,
    pc_hpsa_status     = hpsa$pc_hpsa_status,
    # PCSA fields
    pcsa_score         = pcsa$pcsa_score,
    mssa_id            = pcsa$mssa_id
  )
  
  message(sprintf(
    "  → County: %s | HPSA: %s (score: %s) | PCSA: %s",
    ifelse(is.na(geo$county_name), "Not found", geo$county_name),
    ifelse(hpsa$pc_hpsa_designated, "Designated", "Not Designated"),
    ifelse(hpsa$pc_hpsa_designated, hpsa$pc_hpsa_score, "—"),
    ifelse(is.na(pcsa$pcsa_score), "Not designated", pcsa$pcsa_score)
  ))
}

# =============================================================================
# 8. COMBINE & EXPORT
# =============================================================================
df_out <- bind_rows(results)

other_cols <- setdiff(names(df_in), ADDRESS_COL)
if (length(other_cols) > 0) {
  df_out <- bind_cols(df_in[, other_cols, drop = FALSE], df_out)
}

write_xlsx(df_out, OUTPUT_FILE)

message("\nDone! Results saved to: ", OUTPUT_FILE)
message(sprintf(
  "Summary: %d HPSA designated | %d not designated | %d geocoding failures | %d PCSA matched",
  sum(df_out$pc_hpsa_designated, na.rm = TRUE),
  sum(!df_out$pc_hpsa_designated & !is.na(df_out$latitude), na.rm = TRUE),
  sum(is.na(df_out$latitude)),
  sum(!is.na(df_out$pcsa_id))
))