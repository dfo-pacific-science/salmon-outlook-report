###############################################################################
# Script Name: tableTest3.R
#
# Purpose:
#   This script prepares and formats tables summarizing annual status
#   evaluations for salmon stock management units (SMUs) and associated
#   conservation units (CUs). It processes raw input data, performs validation
#   checks, enriches records, and builds styled tables for reporting.
#
# Inputs:
#   - data/phase1culookup.csv: Crosswalk file mapping CU codes to descriptive names.
#   - data/testData3.xlsx: Excel workbook containing multiple sheets with SMU and CU data.
#   - Required data frames in the Excel sheets:
#       * Salmon_Outlook_Report
#       * cu_outlook_records
#
# Outputs:
#   - A processed tibble (`tabPrep`) containing enriched SMU/CU data.
#   - Functions to generate formatted tables (`make_table`) with captions.
#   - Example usage at the bottom demonstrates table creation for specific areas/species.
#
# Dependencies:
#   dplyr, ggplot2, readr, tibble, rosettafish, csasdown, officer, flextable,
#   readxl, purrr, stringr
###############################################################################

# Load required libraries
library(dplyr)
library(ggplot2)
library(readr)
library(tibble)
library(rosettafish)
library(csasdown)
library(officer)
library(flextable)
library(readxl)
library(purrr)
library(stringr)

# TEST

### ---------------------- Data Preparation ---------------------- ###
# Read CU crosswalk file (maps CU codes to descriptive labels)
crosswalkList = read.csv("data/phase1culookup.csv", stringsAsFactors = FALSE)

# Path to Excel workbook containing SMU and CU data
dummyPath = "data/testData3.xlsx"

# Get all sheet names from the workbook
sheet_names = readxl::excel_sheets(dummyPath)

# Read each sheet into a list of data frames, named by sheet
df_list = map(sheet_names, ~ read_excel(dummyPath, sheet = .x)) %>% set_names(sheet_names)

# Load each sheet into the global environment for convenience
list2env(df_list, envir = .GlobalEnv)

### ---------------------- Defensive Checks ---------------------- ###
# Ensure required data frames exist before proceeding
if (!exists("Salmon_Outlook_Report") | !exists("cu_outlook_records")) {
  stop("Required data frames Salmon_Outlook_Report or cu_outlook_records are missing.")
}

### ---------------------- Column Selection ---------------------- ###
# Columns to keep from SMU-level report
keep_cols_repeat = c(
  "globalid", "uniquerowid",
  "smu_area", "smu_species", "smu_name",
  "outlook_narrative", "smu_outlook_assignment", "smu_prelim_forecast"
)

# Columns to keep from CU-level records
keep_cols_cu = c(
  "cu_outlook_selection", "cu_outlook_assignment", "cu_prelim_forecast", "cu_count",
  "parentrowid"
)

# Subset SMU data
Outlook_Repeat_Test = Salmon_Outlook_Report %>% select(all_of(keep_cols_repeat))

# Subset CU data and apply initial checks for missing entries
cu_outlook_records = cu_outlook_records %>%
  select(all_of(keep_cols_cu)) %>%
  mutate(
    # Flag missing CU outlook/selection entries
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
    # Flag cases where outlook assigned but CU not specified
    cu_outlook_selection = if_else(
      (is.na(cu_outlook_selection) | cu_outlook_selection == "") &
        (!is.na(cu_outlook_assignment) & cu_outlook_assignment != ""),
      "Outlook assigned but no CU specified",
      cu_outlook_selection
    ),
    # Calculate CU count based on comma-separated list
    cu_count = case_when(
      cu_outlook_selection %in% c("CHECK: No data entered", "Outlook assigned but no CU specified") ~ 1L,
      TRUE ~ str_count(coalesce(cu_outlook_selection, ""), ",") + 1L
    )
  )

### ---------------------- Merge SMU and CU Data ---------------------- ###
cu_outlook_records_enriched = full_join(
  cu_outlook_records,
  Outlook_Repeat_Test,
  by = c("parentrowid" = "uniquerowid")
)

# Add placeholder columns for metrics
tabPrep = cu_outlook_records_enriched %>%
  mutate(
    `Avg Run/Avg Spawners` = "50,000",
    `LRP/LBB` = "n/a",
    `Mgmt Target` = "10,000"
  )

### ---------------------- Helper Function: CU Labels ---------------------- ###
# Converts CU codes into descriptive names using crosswalk file
get_labels = function(cu_string, ref_df) {
  if (is.na(cu_string) | cu_string == "") return(NA_character_)
  if (cu_string %in% c("CHECK: No data entered", "Outlook assigned but no CU specified")) return(cu_string)

  cu_codes = str_split(cu_string, ",")[[1]] %>% str_trim() %>% toupper()
  labels = ref_df %>%
    mutate(cu = toupper(cu)) %>%
    filter(cu %in% cu_codes) %>%
    arrange(match(cu, cu_codes)) %>%
    pull(label)

  if (length(labels) == 0) {
    paste(cu_codes, collapse = ", ")
  } else {
    paste(labels, collapse = ", ")
  }
}

# Apply CU label mapping
tabPrep = tabPrep %>%
  mutate(CU_Names = map_chr(cu_outlook_selection, ~ get_labels(.x, crosswalkList)))

### ---------------------- Data Cleaning and Checks ---------------------- ###
# Define values considered empty
empty_vals = c(NA_character_, "", "N/A", "NA", "n/a", "na", "No data entered", "CHECK: No data entered")

# Add flags for empty/numeric/data-deficient entries and calculate CU counts
tabPrep = tabPrep %>%
  mutate(
    smu_raw = smu_outlook_assignment,
    cu_raw = cu_outlook_assignment,
    smu_empty = smu_raw %in% empty_vals,
    cu_empty = cu_raw %in% empty_vals,
    smu_dd = tolower(coalesce(smu_raw, "")) == "data deficient",
    cu_dd = tolower(coalesce(cu_raw, "")) == "data deficient",
    smu_numeric = suppressWarnings(!is.na(as.numeric(smu_raw))),
    cu_numeric = suppressWarnings(!is.na(as.numeric(cu_raw)))
  ) %>%
  group_by(smu_name) %>%
  mutate(
    smu_n_distinct_parent = n_distinct(parentrowid),
    smu_duplicate = smu_n_distinct_parent > 1L,
    smu_name_display = if_else(smu_duplicate,
                               paste0(smu_name, " — NOTE: duplicate entries"),
                               smu_name
    ),
    cu_count_calc = case_when(
      cu_empty ~ 0L,
      cu_outlook_selection %in% c("CHECK: No data entered", "Outlook assigned but no CU specified") ~ 1L,
      TRUE ~ str_count(coalesce(cu_outlook_selection, ""), ",") + 1L
    ),
    needs_check = case_when(
      (smu_empty & cu_empty) ~ TRUE,
      (smu_numeric & cu_numeric) ~ TRUE,
      (smu_numeric & cu_dd) ~ TRUE,
      (smu_dd & cu_numeric) ~ TRUE,
      TRUE ~ FALSE
    ),
    check_reason = case_when(
      (smu_empty & cu_empty) ~ "Only NA entered",
      (smu_numeric & cu_numeric) ~ "SMU and CU both numeric",
      (smu_numeric & cu_dd) ~ "SMU numeric but CU data deficient",
      (smu_dd & cu_numeric) ~ "CU numeric but SMU data deficient",
      TRUE ~ NA_character_
    )
  ) %>%
  ungroup() %>%
  mutate(
    cu_count = coalesce(cu_count, cu_count_calc),
    Resolution = case_when(
      smu_empty & cu_empty ~ "CHECK: Only NA entered",
      smu_numeric & cu_numeric ~ "CHECK: SMU and CU both numeric",
      smu_numeric & cu_dd ~ "CHECK: SMU numeric but CU data deficient",
      smu_dd & cu_numeric ~ "CHECK: CU numeric but SMU data deficient",
      cu_empty ~ "SMU",
      !cu_empty & cu_count == 1 ~ "CU (singular)",
      !cu_empty & cu_count > 1 ~ "CU (aggregate)",
      TRUE ~ "CHECK: unexpected combination"
    )
  )

### ---------------------- Combine CU Forecasts/Outlooks ---------------------- ###
tabPrep = tabPrep %>%
  group_by(smu_name_display, Resolution, smu_prelim_forecast, smu_outlook_assignment, outlook_narrative) %>%
  mutate(
    CU_Forecast = paste(na.omit(cu_prelim_forecast), collapse = ", "),
    CU_Outlook = paste(na.omit(cu_outlook_assignment), collapse = ", "),
    CU_CodeList = paste(na.omit(cu_outlook_selection), collapse = ", ")
  ) %>%
  ungroup() %>%
  mutate(
    Name = case_when(
      Resolution == "SMU" ~ smu_name_display,
      Resolution %in% c("CU (aggregate)", "CU (singular)") ~ CU_Names,
      str_detect(Resolution, "^CHECK:") ~ CU_Names,
      TRUE ~ NA_character_
    ),
    Forecast = case_when(
      Resolution == "SMU" ~ smu_prelim_forecast,
      Resolution %in% c("CU (aggregate)", "CU (singular)") ~ CU_Forecast,
      str_detect(Resolution, "^CHECK:") ~ paste0(
        "CHECK VALUE: Both SMU and CU forecasts entered. SMU = ",
        coalesce(smu_prelim_forecast, "NA"),
        "; CU = ", coalesce(CU_Forecast, "NA"),
        " [", CU_CodeList, "]"
      ),
      TRUE ~ NA_character_
    ),
    Outlook = case_when(
      Resolution == "SMU" ~ coalesce(smu_outlook_assignment, NA_character_),
      Resolution %in% c("CU (aggregate)", "CU (singular)") ~ coalesce(cu_outlook_assignment, NA_character_),
      str_detect(Resolution, "^CHECK:") ~ paste0(
        "CHECK VALUE: Both SMU and CU outlooks entered. SMU = ",
        coalesce(smu_outlook_assignment, "NA"),
        "; CU = ", coalesce(CU_Outlook, "NA"),
        " [", CU_CodeList, "]"
      ),
      TRUE ~ NA_character_
    ),
    Narrative = outlook_narrative
  )

### ---------------------- Final Table Prep ---------------------- ###
tabPrep = tabPrep %>%
  distinct(
    smu_area, smu_species, smu_name_display, Narrative, Resolution, Name,
    `Avg Run/Avg Spawners`, `LRP/LBB`, `Mgmt Target`, Forecast, Outlook
  ) %>%
  rename(smu_name = smu_name_display) %>%
  mutate(
    Forecast = coalesce(Forecast, NA_character_) %>%
      str_replace_all("\\s*\\-\\s*", "–")
  ) %>%
  select(
    smu_area, smu_species, smu_name, Narrative, Resolution, Name,
    `Avg Run/Avg Spawners`, `LRP/LBB`, `Mgmt Target`, Forecast, Outlook
  )

### ---------------------- Table-Building Functions ---------------------- ###
# Build a block of rows for a single SMU
build_block = function(df_smu) {
  smu = unique(df_smu$smu_name)
  narr = unique(df_smu$Narrative)

                # Header rows for narrative and column labels
                row1 = data.frame(Resolution = "SMU", Name = "Narrative",
                                  `Avg Run/Avg Spawners` = "", `LRP/LBB` = "", `Mgmt Target` = "",
                                  Forecast = "", Outlook = "", check.names = FALSE, stringsAsFactors = FALSE)
                row2 = data.frame(Resolution = smu, Name = narr,
                                  `Avg Run/Avg Spawners` = "", `LRP/LBB` = "", `Mgmt Target` = "",
                                  Forecast = "", Outlook = "", check.names = FALSE, stringsAsFactors = FALSE)
                row3 = data.frame(Resolution = "Resolution", Name = "Name",
                                  `Avg Run/Avg Spawners` = "Avg Run/Avg Spawners",
                                  `LRP/LBB` = "LRP/LBB", `Mgmt Target` = "Mgmt Target",
                                  Forecast = "Forecast", Outlook = "Outlook", check.names = FALSE, stringsAsFactors = FALSE)

                # Actual data rows
                data_rows = df_smu %>%
                  select(Resolution, Name, `Avg Run/Avg Spawners`, `LRP/LBB`, `Mgmt Target`, Forecast, Outlook)

                bind_rows(row1, row2, row3, data_rows)
}

# Apply styling to the combined table
style_smu_table = function(big_df) {
  ft = flextable(big_df)
  ft = delete_part(ft, part = "header")

  # Identify special rows for styling
  idx_label = which(big_df$Resolution == "SMU" & big_df$Name == "Narrative")
  idx_header = which(big_df$Resolution == "Resolution")

  # Apply background and bold formatting
  ft = bg(ft, i = idx_label, bg = "gray90")
  ft = bold(ft, i = idx_label, bold = TRUE)
  ft = bg(ft, i = idx_header, bg = "gray90")
  ft = bold(ft, i = idx_header, bold = TRUE)

  # Add borders
  ft = border_remove(ft)
  ft = border_outer(ft, border = fp_border(color = "black", width = 1))
  ft = border_inner_h(ft, border = fp_border(color = "black", width = 1))
  ft = border_inner_v(ft, border = fp_border(color = "black", width = 1))

  # Merge cells for narrative rows
  for (i in idx_label) {
    ft = merge_at(ft, i = i, j = 2:7)
    ft = merge_at(ft, i = i + 1, j = 2:7)
  }

  # Add thick separators between SMUs
  if (length(idx_label) > 1) {
    idx_separators = idx_label[-1]
    for (i in idx_separators) {
      ft = border(ft, i = i, j = 1:ncol(big_df),
                  border.top = fp_border(color = "black", width = 2), part = "body")
    }
  }

  ft = autofit(ft)
  ft = set_table_properties(ft, layout = "autofit")
  ft
}

# Generate caption text for the table
make_caption = function(species, area, year = format(Sys.Date(), "%Y")) {
  area_titleCase = area %>%
    tolower() %>% tools::toTitleCase() %>%
    gsub("\\bAnd\\b", "and", .)

  paste0(
    "Summary of metrics informing the annual status evaluation for ",
    species, " in the ", area_titleCase, " area during the ", year, " management cycle. ",
    "Values are presented for each stock management unit (SMU), and where applicable, for associated singular conservation units (CUs), CU aggregations, and Hatchery or Indicator stocks. ",
    "Reported values include recent average run size or spawner abundance, biological reference points (limit reference point [LRP] and lower biological benchmark [LBB]), and management (mgmt.) targets or operational control points used to interpret current conditions. ",
    "The forecast provides a numerical abundance estimate for the return year, while the outlook indicates the categorical status classification (see definitions in Table 1)."
  )
}

# Main function to create a styled table for a given area and species
make_table = function(area, species, data = tabPrep) {
  df_filtered = data %>% filter(smu_area == area, smu_species == species)
  if (nrow(df_filtered) == 0) stop(paste("No data found for:", area, "/", species))

  smu_list = split(df_filtered, df_filtered$smu_name)
  block_list = lapply(smu_list, build_block)
  big_df = bind_rows(block_list)

  ft = style_smu_table(big_df)
  caption_text = make_caption(species = species, area = area)
  ft = set_caption(ft, caption = caption_text)
  ft
}

# Example usage:
make_table("FRASER AND INTERIOR", "Chinook")
make_table("SOUTH COAST", "Chinook")