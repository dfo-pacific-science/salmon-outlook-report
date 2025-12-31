# STEPHEN TO DO:
# Make it so you enter Year --> that determines Pink Even vs Pink Odd
# NORTH COAST: species = Pink not Pink even/odd






################################################################################
# Script Name: statusTables_rewrite.R
#
# Purpose:
#   Rewritten from original tableTest3.R to merge CU-level records and an
#   additional `other_records` data frame (Hatchery or Indicator stocks) into
#   the SMU/CU status table workflow. The `other_records` entries are treated
#   similarly to cu_outlook_records and are assigned Resolution =
#   "Hatchery or Indicator Stock".
#
# Inputs: (same as original script plus `other_records` sheet if present)
#   - data/phase1culookup.csv
#   - data/testData3.xlsx (sheets: Salmon_Outlook_Report, cu_outlook_records,
#       optionally other_records)
#
# Outputs:
#   - tabPrep: combined and cleaned SMU/CU/Other tibble with added flags and
#     labels.
#   - make_table(): function for generating styled output tables.
################################################################################

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

################################################################################
###  Exceptions / Overrides
# 1. PFMA exception:
#    - SMU: "CENTRAL COAST COHO SALMON"
#    - Set Resolution = "PFMA" (gets entered as SMU)
#    - Ignore duplicate SMU checks
# 2. SKEENA COHO exception:
#    - SMU: "SKEENA COHO SALMON"
#    - Allow numeric or data-deficient entries for both SMU and CU
#    - Do not flag as check
# 3. FRASER AND INTERIOR exception:
#    - Area: "FRASER AND INTERIOR"
#    - Allow numeric or data-deficient for both SMU and CU
#    - Do not flag as check

################################################################################
### Prep Data: read in required files

# Set the path to workbook
lookup_xl_path = "data/lookupTables.xlsx"

# Read the crosswalk list from the Excel sheet
crosswalkList = read_excel(
  path  = lookup_xl_path,
  sheet = "phase1culookup"
) %>% as.data.frame(stringsAsFactors = FALSE)

# Read the hatchery/indicator table from the Excel sheet
hatcheryIndicator = read_excel(
  path  = lookup_xl_path,
  sheet = "hatcheryIndicator"
) %>% as.data.frame(stringsAsFactors = FALSE)

# crosswalkList = read.csv("data/phase1culookup.csv", stringsAsFactors = FALSE)

dummyPath = "data/22Dec2025Data.xlsx"
sheet_names = readxl::excel_sheets(dummyPath)
df_list = map(sheet_names, ~ read_excel(dummyPath, sheet = .x)) %>% set_names(sheet_names)
list2env(df_list, envir = .GlobalEnv)

################################################################################
### Initial check: required sheets
if (!exists("Salmon_Outlook_Report") | !exists("cu_outlook_records")) {
  stop("Required data frames Salmon_Outlook_Report or cu_outlook_records are missing.")
}

# other_records is optional; if it doesn't exist we continue without it
has_other = exists("other_records")

################################################################################
### Column Selection: keep only relevant columns
keep_cols_repeat = c(
  "globalid", "uniquerowid",
  "smu_area", "smu_species", "smu_name",
  "outlook_narrative", "smu_outlook_assignment", "smu_prelim_forecast"
)

keep_cols_cu = c(
  "cu_outlook_selection", "cu_outlook_assignment", "cu_prelim_forecast", "cu_count",
  "parentrowid"
)

# Columns for other_records (expected names in the workbook)
keep_cols_other = c(
  "other_outlook_selection", "other_outlook_assignment", "other_prelim_forecast", "parentrowid"
)

Outlook_Repeat_Test = Salmon_Outlook_Report %>% select(all_of(keep_cols_repeat))

################################################################################
### Prepare cu_outlook_records (same as original)
cu_outlook_records = cu_outlook_records %>%
  select(any_of(keep_cols_cu)) %>%
  mutate(
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
    cu_outlook_selection = if_else(
      (is.na(cu_outlook_selection) | cu_outlook_selection == "") &
        (!is.na(cu_outlook_assignment) & cu_outlook_assignment != ""),
      "Outlook assigned but no CU specified",
      cu_outlook_selection
    ),
    cu_count = case_when(
      cu_outlook_selection %in% c("CHECK: No data entered", "Outlook assigned but no CU specified") ~ 1L,
      TRUE ~ str_count(coalesce(cu_outlook_selection, ""), ",") + 1L
    )
  )

################################################################################
### Prepare other_records (if present) and align to CU schema
if (has_other) {
  other_prep = other_records %>% select(any_of(keep_cols_other)) %>%
    rename(
      cu_outlook_selection = other_outlook_selection,
      cu_outlook_assignment = other_outlook_assignment,
      cu_prelim_forecast = other_prelim_forecast
    ) %>%
    mutate(
      # mark source so we can later force Resolution for these rows
      source = "other",
      # ensure NA/blank checks match the CU preprocessing
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
      cu_outlook_selection = if_else(
        (is.na(cu_outlook_selection) | cu_outlook_selection == "") &
          (!is.na(cu_outlook_assignment) & cu_outlook_assignment != ""),
        "Outlook assigned but no CU specified",
        cu_outlook_selection
      ),
      cu_count = case_when(
        cu_outlook_selection %in% c("CHECK: No data entered", "Outlook assigned but no CU specified") ~ 1L,
        TRUE ~ str_count(coalesce(cu_outlook_selection, ""), ",") + 1L
      )
    )
} else {
  other_prep = tibble()
}

################################################################################
### Merge SMU and CU/Other Data
# Combine cu_outlook_records and other_prep so both are treated identically downstream
combined_cu_other = bind_rows(
  cu_outlook_records %>% mutate(source = "cu"),
  other_prep
)

# Join to SMU info by parentrowid -> uniquerowid
cu_outlook_records_enriched = full_join(
  combined_cu_other,
  Outlook_Repeat_Test,
  by = c("parentrowid" = "uniquerowid")
)

# Add placeholder metric columns
tabPrep = cu_outlook_records_enriched %>%
  mutate(
    `Avg Run/Avg Spawners` = "50,000",
    `LRP/LBB` = "n/a",
    `Mgmt Target` = "10,000"
  )

################################################################################
### Add CU/Other names as a column
# Reuse get_labels but if row is from `other` use the raw selection as the Name
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

# Add CU_Names but keep original other selections for later
tabPrep = tabPrep %>%
  mutate(
    CU_Names = map_chr(cu_outlook_selection, ~ get_labels(.x, crosswalkList)),
    Other_RawSelection = if_else(source == "other", coalesce(cu_outlook_selection, ""), NA_character_)
  )

################################################################################
### Data checks and flagging (same logic as original; applied to combined rows)
empty_vals = c(
  NA_character_, "", "N/A", "NA", "n/a", "na",
  "No data entered", "CHECK: No data entered"
)

tabPrep = tabPrep %>%
  mutate(
    smu_raw = smu_outlook_assignment,
    cu_raw  = cu_outlook_assignment,
    smu_empty = smu_raw %in% empty_vals,
    cu_empty  = cu_raw %in% empty_vals,
    smu_dd    = tolower(coalesce(smu_raw, "")) == "data deficient",
    cu_dd     = tolower(coalesce(cu_raw, "")) == "data deficient",
    smu_numeric = suppressWarnings(!is.na(as.numeric(smu_raw))),
    cu_numeric  = suppressWarnings(!is.na(as.numeric(cu_raw)))
  ) %>%
  group_by(smu_name) %>%
  mutate(
    smu_n_distinct_parent = n_distinct(parentrowid),
    smu_duplicate = smu_n_distinct_parent > 1L,
    smu_name_display = if_else(
      smu_duplicate,
      paste0(smu_name, " — CHECK: Data entered for the same SMU more than once"),
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
    # For rows coming from other_records, set Resolution to Hatchery/Indicator override
    Resolution = case_when(
      source == "other" ~ "Hatchery or Indicator Stock",
      smu_empty & cu_empty ~ "CHECK: Only NA entered",
      smu_numeric & cu_numeric ~ "CHECK: SMU and CU both numeric",
      smu_numeric & cu_dd ~ "CHECK: SMU numeric but CU data deficient",
      smu_dd & cu_numeric ~ "CHECK: SMU data deficient but CU numeric",
      cu_empty ~ "SMU",
      !cu_empty & cu_count == 1 ~ "CU (singular)",
      !cu_empty & cu_count > 1 ~ "CU (aggregate)",
      TRUE ~ "CHECK: unexpected combination"
    )
  )

################################################################################
### Combine CU/Other Forecasts/Outlooks — treat `other` like CU-level but prefer other values

# When source == other, we want Name to come from Other_RawSelection (or the selection text),
# Forecast/Outlook to come from cu_prelim_forecast/cu_outlook_assignment (which were renamed from other_*),
# and Resolution already set to "Hatchery or Indicator Stock" above.

tabPrep = tabPrep %>%
  group_by(smu_name_display, Resolution, smu_prelim_forecast, smu_outlook_assignment, outlook_narrative) %>%
  mutate(
    CU_Forecast  = paste(unique(na.omit(cu_prelim_forecast)), collapse = ", "),
    CU_Outlook   = paste(unique(na.omit(cu_outlook_assignment)), collapse = ", "),
    CU_CodeList  = paste(unique(na.omit(cu_outlook_selection)), collapse = ", ")
  ) %>%
  ungroup() %>%
  mutate(
    Name = case_when(
      Resolution == "SMU" ~ smu_name_display,
      Resolution %in% c("CU (aggregate)", "CU (singular)") ~ CU_Names,
      Resolution == "Hatchery or Indicator Stock" & !is.na(Other_RawSelection) & Other_RawSelection != "" ~ Other_RawSelection,
      str_detect(Resolution, "^CHECK:") ~ CU_Names,
      TRUE ~ NA_character_
    ),
    Forecast = case_when(
      Resolution == "SMU" ~ smu_prelim_forecast,
      Resolution %in% c("CU (aggregate)", "CU (singular)") ~ CU_Forecast,
      Resolution == "Hatchery or Indicator Stock" ~ CU_Forecast,
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
      Resolution == "Hatchery or Indicator Stock" ~ coalesce(cu_outlook_assignment, NA_character_),
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

################################################################################
### Final Table Prep

tabPrep = tabPrep %>%
  distinct(
    smu_area, smu_species, smu_name_display, Narrative, Resolution, Name,
    `Avg Run/Avg Spawners`, `LRP/LBB`, `Mgmt Target`, Forecast, Outlook
  ) %>%
  rename(smu_name = smu_name_display) %>%
  # mutate(
  #   Forecast = coalesce(Forecast, NA_character_) %>%
  #     str_replace_all("\\s*\-\s*", "–")
  # ) %>%
  select(
    smu_area, smu_species, smu_name, Narrative, Resolution, Name,
    `Avg Run/Avg Spawners`, `LRP/LBB`, `Mgmt Target`, Forecast, Outlook
  )

################################################################################
### Table building and styling functions (unchanged)

build_block = function(df_smu) {
  smu  = unique(df_smu$smu_name)
  narr = unique(df_smu$Narrative)

  row1 = data.frame(
    Resolution = "SMU",
    Name = "Narrative",
    Forecast = "",
    Outlook  = "",
    stringsAsFactors = FALSE
  )

  row2 = data.frame(
    Resolution = smu,
    Name = narr,
    Forecast = "",
    Outlook  = "",
    stringsAsFactors = FALSE
  )

  row3 = data.frame(
    Resolution = "Resolution",
    Name = "Name",
    Forecast = "Forecast",
    Outlook  = "Outlook",
    stringsAsFactors = FALSE
  )

  data_rows = df_smu %>%
    select(Resolution, Name, Forecast, Outlook)

  bind_rows(row1, row2, row3, data_rows)
}

################################################################################

style_smu_table = function(big_df) {

  # ---- Identify structural rows ONCE ----
  idx_label  = which(big_df$Resolution == "SMU" & big_df$Name == "Narrative")
  idx_header = which(big_df$Resolution == "Resolution")

  # Draw separator BELOW the previous SMU block
  idx_separators = idx_label[-1] - 1

  ft = flextable(big_df)
  ft = delete_part(ft, part = "header")

  # ---- Reset all borders ----
  ft = border_remove(ft)

  # ---- Thick SMU separators FIRST (before merges) ----
  if (length(idx_separators) > 0) {
    ft = border(
      ft,
      i = idx_separators,
      j = 1:ncol(big_df),
      border.bottom = fp_border(width = 2),
      part = "body"
    )
  }

  # ---- Standard table borders ----
  ft = border_outer(ft, fp_border(width = 1))
  ft = border_inner_v(ft, fp_border(width = 1))
  # ft = border_inner_h(ft, fp_border(width = 1))

  all_rows = seq_len(nrow(big_df))
  thin_rows = setdiff(all_rows, idx_separators)

  ft = border(
    ft,
    i = thin_rows,
    j = 1:ncol(big_df),
    border.bottom = fp_border(width= 1),
    part = "body"

  )


  # ---- Merge rows ONLY after borders are done ----
  for (i in idx_label) {
    ft = merge_at(ft, i = i,     j = 2:4)
    ft = merge_at(ft, i = i + 1, j = 2:4)
  }

  # ---- Styling (safe after merges) ----
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

  area_titleCase = area %>%
    tolower() %>%
    tools::toTitleCase() %>%
    gsub("\\bAnd\\b", "and", .)

  paste0(
    "Summary of metrics informing the annual status evaluation for ",
    species, " in the ", area_titleCase, " area during the ", year,
    " management cycle. Values are presented for each stock management unit (SMU), ",
    "and where applicable, for associated singular conservation units (CUs), CU ",
    "aggregations, and Hatchery or Indicator stocks."
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
  ft = set_caption(ft, caption = make_caption(species, area))

  ft
}

################################################################################
# Example usage (unchanged)
#make_table("FRASER AND INTERIOR", "Chinook")
#make_table("SOUTH COAST", "Chinook")

# End of script