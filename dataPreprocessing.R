################################################################################
### SCRIPT PURPOSE
# Prepares SMU- and CU-level salmon outlook data for reporting.
# Combines:
#   - SMU-level outlooks
#   - CU-level outlooks
#   - optional Hatchery/Indicator stock records
#
# Steps:
#   1. Read data from Excel
#   2. Validate SMU/CU combinations
#   3. Flag missing or inconsistent data with CHECK messages
#   4. Assign each row a "Resolution" category
#   5. Produce formatted tables for Word output
#
# Key concepts:
# - parentrowid links CU/Other rows to SMU rows
# - Resolution determines the output structure
# - CHECK rows indicate data issues that need manual review

# Notes:
# This script is







################################################################################
### LOAD LIBRARIES
library(dplyr)
library(readr)
library(tibble)
library(csasdown)
library(officer)
library(flextable)
library(readxl)
library(purrr)
library(stringr)

################################################################################
### LOAD LOOKUP TABLES
# Lookup tables contain official SMU/CU names, CU codes, and hatchery/indicator stock info.
# Used for:
#   - validating SMU/CU entries
#   - mapping CU codes to full names
#   - checking missing/incorrectly-entered CU or stock info

lookup_xl_path = "data/lookupTables.xlsx"

# Phase 1 CU lookup table (official CU codes & names)
crosswalkList = read_excel(
  path  = lookup_xl_path,
  sheet = "phase1culookup"
) %>% as.data.frame(stringsAsFactors = FALSE)

# Hatchery/Indicator stock table (subset may not be required)
hatcheryIndicator = read_excel(
  path  = lookup_xl_path,
  sheet = "hatcheryIndicator"
) %>% as.data.frame(stringsAsFactors = FALSE)

################################################################################
### LOAD DATA SHEETS
# Reads all sheets from input Excel file and loads them as separate data frames.
# Required sheets: Salmon_Outlook_Report (SMU), cu_outlook_records (CU)
# Optional: other_records (Hatchery/Indicator stocks)

dummyPath = "data/09Jan2026Data.xlsx"  # <--- ENTER FILE PATH HERE

# Get sheet names
sheet_names = readxl::excel_sheets(dummyPath)

# Read all sheets into a named list of data frames
df_list = map(sheet_names, ~ read_excel(dummyPath, sheet = .x)) %>%
  set_names(sheet_names)

# Assign each sheet to the global environment as its own data frame
list2env(df_list, envir = .GlobalEnv)

# Stop if required sheets are missing
if (!exists("Salmon_Outlook_Report") | !exists("cu_outlook_records")) {
  stop("Required data frames Salmon_Outlook_Report or cu_outlook_records are missing.")
}

# Flag whether optional Hatchery/Indicator data is present
has_other = exists("other_records")

################################################################################
### COLUMN SELECTION
# Only keep columns needed for report tables.

keep_cols_repeat = c(
  "globalid", "uniquerowid",
  "smu_area", "smu_species", "smu_name",
  "outlook_narrative", "smu_outlook_assignment", "smu_prelim_forecast"
)

keep_cols_cu = c(
  "cu_outlook_selection", "cu_outlook_assignment",
  "cu_prelim_forecast", "cu_count", "parentrowid"
)

keep_cols_other = c(
  "other_outlook_selection", "other_outlook_assignment",
  "other_prelim_forecast", "parentrowid"
)

# Subset SMU-level data
Outlook_Repeat_Test = Salmon_Outlook_Report %>%
  select(all_of(keep_cols_repeat))

################################################################################
### PREPARE CU-LEVEL DATA
# Standardize CU entries and flag missing/incorrect data using CHECK messages
# cu_count counts number of CU codes (comma-separated)
# CHECK logic:
#   1) Both CU selection & outlook missing -> "CHECK: No data entered"
#   2) Outlook assigned but no CU -> "Outlook assigned but no CU specified"

cu_outlook_records = cu_outlook_records %>%
  select(any_of(keep_cols_cu)) %>%
  mutate(
    # CHECK: no CU data entered
    cu_outlook_selection = if_else(
      (is.na(cu_outlook_selection) | cu_outlook_selection == "") &
        (is.na(cu_outlook_assignment) | cu_outlook_assignment == ""),
      "CHECK: No data entered",
      cu_outlook_selection
    ),
    cu_outlook_assignment = if_else(
      (is.na(cu_outlook_selection) | cu_outlook_selection == "") &
        (is.na(cu_outlook_assignment) | cu_outlook_assignment == ""),
      "CHECK: No data entered",
      cu_outlook_assignment
    ),
    # CHECK: outlook assigned, but no CU specified
    cu_outlook_selection = if_else(
      (is.na(cu_outlook_selection) | cu_outlook_selection == "") &
        (!is.na(cu_outlook_assignment) & cu_outlook_assignment != ""),
      "Outlook assigned but no CU specified",
      cu_outlook_selection
    ),
    # Count number of CUs listed (comma-separated)
    cu_count = case_when(
      cu_outlook_selection %in%
        c("CHECK: No data entered", "Outlook assigned but no CU specified") ~ 1L,
      TRUE ~ str_count(coalesce(cu_outlook_selection, ""), ",") + 1L
    )
  )

################################################################################
### PREPARE OTHER RECORDS (HATCHERY / INDICATOR STOCKS)
# Optional hatchery/indicator data handled like CU rows, but flagged as "other"

if (has_other) {
  other_prep = other_records %>%
    select(any_of(keep_cols_other)) %>%
    rename(
      cu_outlook_selection  = other_outlook_selection,
      cu_outlook_assignment = other_outlook_assignment,
      cu_prelim_forecast    = other_prelim_forecast
    ) %>%
    mutate(source = "other")
} else {
  other_prep = tibble()
}

################################################################################
### MERGE SMU AND CU/OTHER DATA
# Link CU/Other rows to SMU rows via parentrowid = uniquerowid
# Full join ensures SMUs without CU data are retained

combined_cu_other = bind_rows(
  cu_outlook_records %>% mutate(source = "cu"),
  other_prep
)

cu_outlook_records_enriched = full_join(
  combined_cu_other,
  Outlook_Repeat_Test,
  by = c("parentrowid" = "uniquerowid")
)

################################################################################
### EXPLICIT SMU-ONLY ROWS
# Some SMUs have valid SMU-level outlooks but no CU entries
# Ensure these SMUs appear in final tables

smu_only_rows = Outlook_Repeat_Test %>%
  filter(!is.na(smu_outlook_assignment), smu_outlook_assignment != "") %>%
  mutate(
    cu_outlook_selection  = NA_character_,
    cu_outlook_assignment = NA_character_,
    cu_prelim_forecast    = NA_character_,
    cu_count              = NA_integer_,
    source                = "smu",
    parentrowid           = uniquerowid
  )

cu_outlook_records_enriched = bind_rows(cu_outlook_records_enriched, smu_only_rows) %>%
  distinct(
    parentrowid, cu_outlook_selection, cu_outlook_assignment, source,
    .keep_all = TRUE
  )

################################################################################
### PREPARE FLAGS AND VALIDATION CHECKS
# Derive logical flags used to assign Resolution
# Flags are internal only, not displayed in output

empty_vals = c(NA_character_, "", "N/A", "NA", "n/a", "na",
               "No data entered", "CHECK: No data entered")

tabPrep = cu_outlook_records_enriched %>%
  mutate(
    # preserve raw values for diagnostics
    smu_raw = smu_outlook_assignment,
    cu_raw  = cu_outlook_assignment,
    # empty checks
    smu_empty = smu_raw %in% empty_vals,
    cu_empty  = cu_raw %in% empty_vals,
    # Data-deficient checks
    smu_dd    = tolower(coalesce(smu_raw, "")) == "data deficient",
    cu_dd     = tolower(coalesce(cu_raw, "")) == "data deficient",
    # Numeric checks
    smu_numeric = suppressWarnings(!is.na(as.numeric(smu_raw))),
    cu_numeric  = suppressWarnings(!is.na(as.numeric(cu_raw))),
    # Certain SMUs exempt from numeric/DD checks
    no_check_smu = smu_area == "FRASER AND INTERIOR" |
      smu_name == "SKEENA COHO SALMON"
  ) %>%
  group_by(smu_name) %>%
  mutate(
    smu_has_any_data     = any(!smu_empty),
    smu_n_distinct_parent = n_distinct(parentrowid),
    smu_duplicate        = smu_n_distinct_parent > 1L,
    smu_name_display     = case_when(
      smu_name == "CENTRAL COAST COHO SALMON" ~ smu_name,
      smu_name == "NO DESIGNATED SMU" ~ smu_name,
      smu_duplicate ~ paste0(smu_name, " — CHECK: Data entered for the same SMU more than once"),
      TRUE ~ smu_name
    ),
    # recalculated CU count (fallback)
    cu_count_calc = case_when(
      cu_empty ~ 0L,
      cu_outlook_selection %in%
        c("CHECK: No data entered", "Outlook assigned but no CU specified") ~ 1L,
      TRUE ~ str_count(coalesce(cu_outlook_selection, ""), ",") + 1L
    )
  ) %>%
  ungroup() %>%
  mutate(
    cu_count = coalesce(cu_count, cu_count_calc),
    # identify genuine CU rows
    is_cu_row =
      !cu_empty &
      !cu_outlook_selection %in%
      c("CHECK: No data entered", "Outlook assigned but no CU specified")
    # Static columns for output
  )

################################################################################
### RESOLUTION LOGIC
# Assign Resolution category for each row based on SMU/CU info and checks

tabPrep = tabPrep %>%
  mutate(
    Resolution = case_when(
      smu_name == "CENTRAL COAST COHO SALMON" ~ "PFMA",
      source == "other" ~ "Hatchery or Indicator Stock",
      smu_empty & cu_empty & !smu_has_any_data ~ "CHECK: Only NA entered",
      !no_check_smu & smu_numeric & cu_numeric ~ "CHECK: SMU and CU both numeric",
      !no_check_smu & smu_numeric & cu_dd      ~ "CHECK: SMU numeric but CU data deficient",
      !no_check_smu & smu_dd      & cu_numeric ~ "CHECK: SMU data deficient but CU numeric",
      !is_cu_row ~ "SMU",
      is_cu_row & cu_count == 1 ~ "CU (singular)",
      is_cu_row & cu_count > 1  ~ "CU (aggregate)",
      TRUE ~ "CHECK: unexpected combination"
    )
  )

################################################################################
### CU NAME MAPPING FUNCTION
# Converts CU codes (comma-separated) to full CU names using lookup table

get_labels = function(cu_string, ref_df) {
  if (is.na(cu_string) | cu_string == "") return(NA_character_)
  if (cu_string %in% c("CHECK: No data entered", "Outlook assigned but no CU specified"))
    return(cu_string)
  cu_codes = str_split(cu_string, ",")[[1]] %>% str_trim() %>% toupper()
  labels = ref_df %>%
    mutate(cu = toupper(cu)) %>%
    filter(cu %in% cu_codes) %>%
    arrange(match(cu, cu_codes)) %>%
    pull(label)
  if (length(labels) == 0) paste(cu_codes, collapse = ", ") else paste(labels, collapse = ", ")
}

tabPrep = tabPrep %>%
  mutate(
    CU_Names = map_chr(cu_outlook_selection, ~ get_labels(.x, crosswalkList)),
    Other_RawSelection =
      if_else(source == "other", coalesce(cu_outlook_selection, ""), NA_character_)
  )

################################################################################
### COMBINE FORECASTS AND NAME FIELDS
# Collapse CU-level forecasts when Resolution == "CU (aggregate)"
# Prepare final display columns: Name, Forecast, Outlook

tabPrep = tabPrep %>%
  group_by(
    smu_name_display, Resolution,
    parentrowid, smu_prelim_forecast, smu_outlook_assignment,
    outlook_narrative, smu_area
  ) %>%
  mutate(
    CU_Forecast = if_else(
      Resolution == "CU (aggregate)",
      paste(unique(na.omit(cu_prelim_forecast)), collapse = ", "),
      cu_prelim_forecast
    ),
    CU_Outlook = if_else(
      Resolution == "CU (aggregate)",
      paste(unique(na.omit(cu_outlook_assignment)), collapse = ", "),
      cu_outlook_assignment
    ),
    CU_CodeList = if_else(
      Resolution == "CU (aggregate)",
      paste(unique(na.omit(cu_outlook_selection)), collapse = ", "),
      cu_outlook_selection
    )
  ) %>%
  ungroup() %>%
  mutate(
    Name = case_when(
      Resolution %in% c("SMU", "PFMA") ~ smu_name_display,
      Resolution %in% c("CU (aggregate)", "CU (singular)") ~ CU_Names,
      Resolution == "Hatchery or Indicator Stock" ~ Other_RawSelection,
      str_detect(Resolution, "^CHECK:") ~ CU_Names,
      TRUE ~ NA_character_
    ),
    Forecast = case_when(
      Resolution %in% c("SMU", "PFMA") ~ smu_prelim_forecast,
      Resolution %in% c("CU (aggregate)", "CU (singular)", "Hatchery or Indicator Stock") ~ CU_Forecast,
      str_detect(Resolution, "^CHECK:") ~ paste0(
        "CHECK VALUE: SMU = ", coalesce(smu_prelim_forecast, "NA"),
        "; CU = ", coalesce(CU_Forecast, "NA"),
        " [", CU_CodeList, "]"
      ),
      TRUE ~ NA_character_
    ),
    Outlook = case_when(
      Resolution %in% c("SMU", "PFMA") ~ smu_outlook_assignment,
      Resolution %in% c("CU (aggregate)", "CU (singular)", "Hatchery or Indicator Stock") ~ CU_Outlook,
      str_detect(Resolution, "^CHECK:") ~ paste0(
        "CHECK VALUE: SMU = ", coalesce(smu_outlook_assignment, "NA"),
        "; CU = ", coalesce(CU_Outlook, "NA"),
        " [", CU_CodeList, "]"
      ),
      TRUE ~ NA_character_
    ),
    Narrative = outlook_narrative,
    row_order = case_when(
      Resolution %in% c("SMU", "PFMA") ~ 1L,
      Resolution %in% c("CU (singular)", "CU (aggregate)") ~ 2L,
      TRUE ~ 3L
    )
  )

################################################################################
### FINAL TABLE PREP
# Arrange, deduplicate, clean species/area names, fix special Outlook cases

tabPrep =
  tabPrep %>%
  arrange(smu_area, smu_species, smu_name_display, row_order) %>%
  distinct(
    smu_area, smu_species, smu_name_display,
    Narrative, Resolution, Name,
    Forecast, Outlook
  ) %>%
  rename(smu_name = smu_name_display) %>%
  select(
    smu_area, smu_species, smu_name, Narrative, Resolution, Name,
    Forecast, Outlook
  ) %>%
  # Remove placeholder rows for fully empty data
  filter(Resolution != "CHECK: Only NA entered") %>%
  # Clean species and area names
  mutate(
    smu_species = case_when(
      smu_species %in% c("Sockeye Lake Type", "Sockeye River Type") ~ "Sockeye",
      smu_species == "Pink Even" ~ "Pink",
      TRUE ~ smu_species
    ),
    smu_area = case_when(
      smu_area == "YUKON TRANSBOUNDARY" ~ "YUKON AND TRANSBOUNDARY",
      TRUE ~ smu_area
    )
  ) %>%
  # Fix specific Outlook anomalies
  mutate(
    Outlook = case_when(
      Name == "MIDDLE FRASER-FRASER CANYON_SP_1.3, LOWER FRASER RIVER_SP_1.3" ~ "1",
      Name == "MIDDLE FRASER RIVER_SP_1.3, NORTH THOMPSON_SP_1.3" ~ "2",
      Name == "MIDDLE FRASER RIVER-PORTAGE_FA_1.3, LOWER FRASER RIVER-UPPER PITT_SU_1.3" ~ "1",
      Name == "MIDDLE FRASER RIVER_SU_1.3, NORTH THOMPSON_SU_1.3" ~ "2 to 3",
      TRUE ~ as.character(Outlook)
    )
  )

