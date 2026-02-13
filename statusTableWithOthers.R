# STEPHEN TO DO:
# Make it so you enter Year --> that determines Pink Even vs Pink Odd
# NORTH COAST: species = Pink not Pink even/odd
# Check lookup table col: which one is actually needed
# Change: include full name of CU in spreadsheet in addition to codes

# Make sure it's clear where to put in file name. Also check name of each sheet

################################################################################
### SCRIPT PURPOSE
# This script prepares SMU- and CU-level salmon outlook data for reporting.
# It combines:
#   - SMU-level outlooks
#   - CU-level outlooks
#   - optional Hatchery/Indicator stock records
#
# The script:
#   1. Reads data from Excel
#   2. Validates SMU/CU combinations
#   3. Flags missing or inconsistent data with  CHECK messages
#   4. Assigns each row a "Resolution" category
#   5. Produces formatted tables for Word output

# IMPORTANT CONCEPTS
# - parentrowid links CU/Other rows to SMU rows
# - Resolution is the core decision variable that drives output structure
# - CHECK rows indicate data problems that should be reviewed manually

################################################################################
### LOAD LIBRARIES

library(dplyr)
library(ggplot2)
library(readr)
library(tibble)
library(csasdown)
library(officer)
library(flextable)
library(readxl)
library(purrr)
library(stringr)

################################################################################
### Load Lookup Tables
# These come from the Salmon Space Crosswalk
# Contain official names for SMUs/CUs, are used to detect missing SMUs/CUs
# And also contain both codes and full names for CUs
# Also includes all names for Hatchery and Indicator stocks. Note that not all
# Hatchery and Indicator stocks are required for the Outlook.

lookup_xl_path = "data/lookupTables.xlsx"

crosswalkList = read_excel(
  path  = lookup_xl_path,
  sheet = "phase1culookup"
) %>% as.data.frame(stringsAsFactors = FALSE)

hatcheryIndicator = read_excel(
  path  = lookup_xl_path,
  sheet = "hatcheryIndicator"
) %>% as.data.frame(stringsAsFactors = FALSE)

################################################################################
### LOAD DATA SHEETS
# This section reads all sheets from the input Excel file and loads them
# into the global environment as separate data frames.
#
# REQUIRED SHEETS:
#   - Salmon_Outlook_Report  (SMU-level data)
#   - cu_outlook_records     (CU-level data)
# OPTIONAL SHEETS:
#   - other_records          (Hatchery / Indicator stocks)
# The script will STOP if required sheets are missing.

# HERE IS WHERE TO ENTER THE FILE PATH FOR THE DATA
dummyPath = "data/09Jan2026Data.xlsx"

# Creates a new data frame from each sheet from the data file
sheet_names = readxl::excel_sheets(dummyPath)

df_list = map(sheet_names, ~ read_excel(dummyPath, sheet = .x)) %>%
  set_names(sheet_names)

list2env(df_list, envir = .GlobalEnv)

if (!exists("Salmon_Outlook_Report") | !exists("cu_outlook_records")) {
  stop("Required data frames Salmon_Outlook_Report or cu_outlook_records are missing.")
}

# Track whether optional Hatchery / Indicator data is present
has_other = exists("other_records")

################################################################################
### Column Selection
# Not all columns in the data spreadsheet are required for the Report/PowerPoint
# Only keep the required ones

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

Outlook_Repeat_Test = Salmon_Outlook_Report %>%
  select(all_of(keep_cols_repeat))

################################################################################
### PREPARE CU-LEVEL DATA
# This block standardizes CU entries and  flags missing or incorrectly-entered CU
# data using CHECK messages.
# Note there are more checks later in the script when these get merged with SMU data
#
# CHECK LOGIC:
#   - If BOTH CU selection and CU outlook are missing:
#       - "CHECK: No data entered"
#   - If an outlook is assigned but no CU is specified:
#       - "Outlook assigned but no CU specified"
#
# cu_count:
#   - Counts number of CU codes supplied (comma-separated)
#   - Used later to distinguish singular vs aggregate CU rows

cu_outlook_records = cu_outlook_records %>%
  select(any_of(keep_cols_cu)) %>%
  mutate(

    # CHECK: no CU information provided was provided
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

    # CHECK: Outlook assigned, but no CU specified
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
### PREPARE OTHER_RECORDS (HATCHERY / INDICATOR STOCKS)

# These records are treated similarly to CU records but are tagged separately

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
### Merge SMU and CU/Other Data

# parentrowid (CU/Other) - uniquerowid (SMU) defines the relationship.
# A full join is used since not all SMUs will have CU-level information reported.

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
### EXPLICIT SMU-ONLY ROWS: COME BACK TO THIS ONE AND CLARIFY

# Some SMUs have valid SMU-level outlooks but NO CU entries.
# This block ensures those SMUs still appear in the final tables.

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
# COME BACK TO THIS ONE TOO

# This section derives logical flags describing WHAT KIND of data was entered.
# These flags are later used to assign a Resolution.
#
# IMPORTANT:
# - Flags do not directly appear in output
# - They exist only to support Resolution logic

empty_vals = c(NA_character_, "", "N/A", "NA", "n/a", "na",
                "No data entered", "CHECK: No data entered")

tabPrep = cu_outlook_records_enriched %>%
  mutate(
    # Preserve raw values for diagnostics
    smu_raw = smu_outlook_assignment,
    cu_raw  = cu_outlook_assignment,

    # Empty checks
    smu_empty = smu_raw %in% empty_vals,
    cu_empty  = cu_raw %in% empty_vals,

    # Data-deficient checks
    smu_dd    = tolower(coalesce(smu_raw, "")) == "data deficient",
    cu_dd     = tolower(coalesce(cu_raw, "")) == "data deficient",

    # Numeric checks
    smu_numeric = suppressWarnings(!is.na(as.numeric(smu_raw))),
    cu_numeric  = suppressWarnings(!is.na(as.numeric(cu_raw))),

    # SMUs exempt from certain numeric/DD checks
    no_check_smu = smu_area == "FRASER AND INTERIOR" |
      smu_name == "SKEENA COHO SALMON"
  ) %>%
  group_by(smu_name) %>%
  mutate(
    smu_has_any_data = any(!smu_empty),
    smu_n_distinct_parent = n_distinct(parentrowid),
    smu_duplicate = smu_n_distinct_parent > 1L,
    smu_name_display = case_when(
      smu_name == "CENTRAL COAST COHO SALMON" ~ smu_name,
      smu_name == "NO DESIGNATED SMU" ~ smu_name,
      smu_duplicate ~ paste0(smu_name, " — CHECK: Data entered for the same SMU more than once"),
      TRUE ~ smu_name
    ),
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

    # Identify rows that genuinely represent CU entries
    is_cu_row =
      !cu_empty &
      !cu_outlook_selection %in%
      c("CHECK: No data entered", "Outlook assigned but no CU specified"),

    ### STATIC COLUMNS REQUIRED FOR LATER TABLE BUILDING
    `Avg Run/Avg Spawners` = "50,000",
    `LRP/LBB` = "n/a",
    `Mgmt Target` = "10,000"
  )

################################################################################
### Resolution Logic

# This section adds the "Resolution" of the data
# i.e., do they report by SMU, CU (singular), CU (aggregate) etc.

# The options include:
# Resolution categories:
#   - "SMU": SMU-only row
#   - "CU (singular)": One CU linked to an SMU
#   - "CU (aggregate)": Multiple CUs aggregated under an SMU
#   - "Hatchery or Indicator Stock"
#   - "PFMA": Special case for Central Coast Coho
#   - "CHECK: <message>": Explicit data issues requiring review
#
# IMPORTANT:
# - Order matters: earlier conditions override later ones
# - Numeric / Data Deficient conflicts are prioritized over SMU/CU classification

tabPrep = tabPrep %>%
  mutate(
    Resolution = case_when(
      smu_name == "CENTRAL COAST COHO SALMON" ~ "PFMA",
      source == "other" ~ "Hatchery or Indicator Stock",

      smu_empty & cu_empty & !smu_has_any_data ~ "CHECK: Only NA entered",

      # CHECKs for numeric/data deficient combinations (priority)
      !no_check_smu & smu_numeric & cu_numeric ~ "CHECK: SMU and CU both numeric",
      !no_check_smu & smu_numeric & cu_dd      ~ "CHECK: SMU numeric but CU data deficient",
      !no_check_smu & smu_dd      & cu_numeric ~ "CHECK: SMU data deficient but CU numeric",

      # SMU-only rows
      !is_cu_row ~ "SMU",

      # CU rows
      is_cu_row & cu_count == 1 ~ "CU (singular)",
      is_cu_row & cu_count > 1  ~ "CU (aggregate)",

      TRUE ~ "CHECK: unexpected combination"
    )
  )

################################################################################
### CU Name Mapping Function
# This could maybe be removed if we just add full CU names to data spreadsheet


# Converts CU codes (e.g., "CU1, CU2") into full CU names using lookup table.
#
# INPUTS:
#   - cu_string : character string of comma-separated CU codes
#   - ref_df    : lookup table with columns `cu` and `label`
#
# OUTPUT:
#   - character string of full CU names in original order


get_labels = function(cu_string, ref_df) {
  # Return NA if no CU information exists
  if (is.na(cu_string) | cu_string == "") return(NA_character_)

  # Preserve CHECK messages
  if (cu_string %in% c("CHECK: No data entered", "Outlook assigned but no CU specified"))
    return(cu_string)

  # Split comma-separated CU codes
  cu_codes = str_split(cu_string, ",")[[1]] %>% str_trim() %>% toupper()

  # Look up full CU names in reference table
  labels = ref_df %>%
    mutate(cu = toupper(cu)) %>%
    filter(cu %in% cu_codes) %>%
    arrange(match(cu, cu_codes)) %>%
    pull(label)

  # If lookup fails, return original codes
  if (length(labels) == 0) paste(cu_codes, collapse = ", ") else paste(labels, collapse = ", ")
}

# Apply CU name mapping and retain raw selections for Hatchery rows
tabPrep = tabPrep %>%
  mutate(
    CU_Names = map_chr(cu_outlook_selection, ~ get_labels(.x, crosswalkList)),
    Other_RawSelection =
      if_else(source == "other", coalesce(cu_outlook_selection, ""), NA_character_)
  )

################################################################################
### COMBINE FORECASTS / OUTLOOKS + NAME FIELDS
#
# This section collapses CU-level values when aggregation is required and
# prepares final display fields used in the output tables.
#
# GROUPING:
#   - Rows are grouped by SMU + Resolution
#   - Aggregated CUs are collapsed into comma-separated strings

tabPrep = tabPrep %>%
  group_by(
    smu_name_display, Resolution,
    parentrowid, smu_prelim_forecast, smu_outlook_assignment,
    outlook_narrative, smu_area
  ) %>%
  mutate(

    # Collapse CU forecasts for aggregated rows
    CU_Forecast = if_else(
      Resolution == "CU (aggregate)",
      paste(unique(na.omit(cu_prelim_forecast)), collapse = ", "),
      cu_prelim_forecast
    ),
    # Collapse CU outlook assignments
    CU_Outlook = if_else(
      Resolution == "CU (aggregate)",
      paste(unique(na.omit(cu_outlook_assignment)), collapse = ", "),
      cu_outlook_assignment
    ),
    # Collapse CU code lists
    CU_CodeList = if_else(
      Resolution == "CU (aggregate)",
      paste(unique(na.omit(cu_outlook_selection)), collapse = ", "),
      cu_outlook_selection
    )
  ) %>%
  ungroup() %>%
  mutate(
    # Display Name column depends on Resolution
    Name = case_when(
      Resolution %in% c("SMU", "PFMA") ~ smu_name_display,
      Resolution %in% c("CU (aggregate)", "CU (singular)") ~ CU_Names,
      Resolution == "Hatchery or Indicator Stock" ~ Other_RawSelection,
      str_detect(Resolution, "^CHECK:") ~ CU_Names,
      TRUE ~ NA_character_
    ),
    # Forecast column depends on Resolution
    Forecast = case_when(
      Resolution %in% c("SMU", "PFMA") ~ smu_prelim_forecast,
      Resolution %in% c("CU (aggregate)", "CU (singular)", "Hatchery or Indicator Stock") ~ CU_Forecast,
      # CHECK rows explicitly show both SMU and CU values
       str_detect(Resolution, "^CHECK:") ~ paste0(
        "CHECK VALUE: SMU = ", coalesce(smu_prelim_forecast, "NA"),
        "; CU = ", coalesce(CU_Forecast, "NA"),
        " [", CU_CodeList, "]"
      ),
      TRUE ~ NA_character_
    ),
    # Outlook column mirrors Forecast logic
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
### Final Table Prep

tabPrep =
  tabPrep %>%
  arrange(smu_area, smu_species, smu_name_display, row_order) %>%
  distinct(
    smu_area, smu_species, smu_name_display,
    Narrative, Resolution, Name,
    `Avg Run/Avg Spawners`, `LRP/LBB`, `Mgmt Target`,
    Forecast, Outlook
  ) %>%
  rename(smu_name = smu_name_display) %>%
  select(
    smu_area, smu_species, smu_name, Narrative, Resolution, Name,
    `Avg Run/Avg Spawners`, `LRP/LBB`, `Mgmt Target`,
    Forecast, Outlook
  ) %>%
  ##### STEPHEN TO LOOK INTO A BETTER WAY OF FIXING
  #### FOR SOME REASON NAS ARE GETTING INCORPORATED IE CHECK: Only NA entered
  ### I believe this is because of my exceptions for FRASER AND INTERIOR data
  # Which allow for SMU AND CU data to be displayed.
  # For now, I am just going to remove any rows with this info
  filter(Resolution != "CHECK: Only NA entered") %>%

  ### Change the
  mutate(

    smu_species = case_when(
      smu_species %in% c("Sockeye Lake Type", "Sockeye River Type") ~ "Sockeye",
      smu_species == "Pink Even" ~ "Pink",
      TRUE ~ smu_species)
  ) %>%

  mutate(
    smu_area = case_when(
      smu_area == "YUKON TRANSBOUNDARY" ~ "YUKON AND TRANSBOUNDARY",
      TRUE ~ smu_area)
  ) %>%

  mutate(
    Outlook = case_when(
      Name == "MIDDLE FRASER-FRASER CANYON_SP_1.3, LOWER FRASER RIVER_SP_1.3" ~ "1",
      Name == "MIDDLE FRASER RIVER_SP_1.3, NORTH THOMPSON_SP_1.3" ~ "2",
      Name == "MIDDLE FRASER RIVER-PORTAGE_FA_1.3, LOWER FRASER RIVER-UPPER PITT_SU_1.3" ~ "1",
      Name == "MIDDLE FRASER RIVER_SU_1.3, NORTH THOMPSON_SU_1.3" ~ "2 to 3",
      TRUE ~ as.character(Outlook)
    )
  )


################################################################################
### Table building and styling functions

# Each SMU gets its own “block” consisting of:
#  1) a grey SMU label row
#  2) the narrative text
#  3) a column header row
#  4) the actual data rows


build_block = function(df_smu) {
  smu  = unique(df_smu$smu_name)
  narr = unique(df_smu$Narrative)

  # Row that just labels the start of a new SMU block
  row1 = data.frame(
    Resolution = "SMU",
    Name = "Narrative",
    Forecast = "",
    Outlook  = "",
    stringsAsFactors = FALSE
  )

  # Row that actually contains the narrative text
  row2 = data.frame(
    Resolution = smu,
    Name = narr,
    Forecast = "",
    Outlook  = "",
    stringsAsFactors = FALSE
  )

  # Row with next set of column names
  row3 = data.frame(
    Resolution = "Resolution",
    Name = "Name",
    Forecast = "Forecast",
    Outlook  = "Outlook",
    stringsAsFactors = FALSE
  )

  # Actual SMU / CU data rows
  data_rows = df_smu %>%
    select(Resolution, Name, Forecast, Outlook)

  bind_rows(row1, row2, row3, data_rows)
}

################################################################################
# This part defines the shading/colouring/line thickness etc of the tables

style_smu_table = function(big_df) {

  # Row where a new SMU block starts
  idx_label  = which(big_df$Resolution == "SMU" & big_df$Name == "Narrative")

  # Row that acts as the column header
  idx_header = which(big_df$Resolution == "Resolution")

  # Draw a thick line *before* each new SMU block except the first one
  idx_separators = idx_label[-1] - 1

  ft = flextable(big_df)

  # Need to delete the real header on purpose to manage flextable styling challenges
  ft = delete_part(ft, part = "header")

  # Reset all borders
  ft = border_remove(ft)

  # Add thick horizontal lines to visually split SMUs
  if (length(idx_separators) > 0) {
    ft = border(
      ft,
      i = idx_separators,
      j = 1:ncol(big_df),
      border.bottom = fp_border(width = 2),
      part = "body"
    )
  }

  # Standard table borders
  ft = border_outer(ft, fp_border(width = 1))
  ft = border_inner_v(ft, fp_border(width = 1))

  all_rows = seq_len(nrow(big_df))
  thin_rows = setdiff(all_rows, idx_separators)

  ft = border(
    ft,
    i = thin_rows,
    j = 1:ncol(big_df),
    border.bottom = fp_border(width= 1),
    part = "body"

  )


  # Merge rows ONLY after borders are done
  for (i in idx_label) {
    ft = merge_at(ft, i = i,     j = 2:4)
    ft = merge_at(ft, i = i + 1, j = 2:4)
  }

  # Styling (safe after merges)
  ft = bg(ft, i = idx_label,  bg = "gray90")
  ft = bold(ft, i = idx_label, bold = TRUE)

  ft = bg(ft, i = idx_header, bg = "gray90")
  ft = bold(ft, i = idx_header, bold = TRUE)

  ft = autofit(ft)
  ft = set_table_properties(ft, layout = "autofit")

  ft
}

################################################################################

make_caption = function(species, area, year = format(Sys.Date(), "%Y")) {
  # Normalize area to Title Case and fix "And" to "and"
  area_titleCase = area %>%
    tolower() %>%
    tools::toTitleCase() %>%
    gsub("\\bAnd\\b", "and", ., perl = TRUE)

  # Exact 5-species lookup
  sci_lookup = c(
    chinook = "Oncorhynchus tshawytscha",
    coho    = "Oncorhynchus kisutch",
    sockeye = "Oncorhynchus nerka",
    chum    = "Oncorhynchus keta",
    pink    = "Oncorhynchus gorbuscha"
  )

  # Key for lookup is lowercased species name
  sp_key = tolower(trimws(species))
  sci = sci_lookup[[sp_key]]  # assumes species is valid

  paste0(
    "Summary of Outlooks, forecasts (where available), and narrative descriptions for ",
    species, " salmon ", sci, " in the ", area_titleCase,
    " area during the ", year, " management cycle. Values are presented for each Stock ",
    "Management Unit (SMU), and where applicable, for associated singular Conservation Units (CUs), ",
    "CU aggregations, and Hatchery or Indicator stocks."
  )
}




################################################################################

make_table = function(area, species, data = tabPrep) {

  df_filtered = data %>%
    filter(smu_area == area, smu_species == species)

  if (nrow(df_filtered) == 0) {
    stop(paste("No data found for:", area, "/", species))
  }

  smu_list   = split(df_filtered, df_filtered$smu_name)
  block_list = lapply(smu_list, build_block)
  big_df     = bind_rows(block_list)

  ft = style_smu_table(big_df)

  caption_text = make_caption(species, area)

  # Split caption around scientific name
  sp_key = tolower(trimws(species))
  sci_lookup = c(
    chinook = "Oncorhynchus tshawytscha",
    coho    = "Oncorhynchus kisutch",
    sockeye = "Oncorhynchus nerka",
    chum    = "Oncorhynchus keta",
    pink    = "Oncorhynchus gorbuscha"
  )
  sci = sci_lookup[[sp_key]]

  parts = strsplit(caption_text, sci)[[1]]

  ft = set_caption(
    ft,
    caption = as_paragraph(
      as_chunk(parts[1], props = fp_text_default(italic = TRUE)),
      as_chunk(sci,        props = fp_text_default(italic = FALSE)),
      as_chunk(parts[2], props = fp_text_default(italic = TRUE))
    )
  )


  ft
}

################################################################################
# Example usage (unchanged)
make_table("FRASER AND INTERIOR", "Chinook")
#make_table("SOUTH COAST", "Chinook")

# End of script