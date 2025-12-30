
############################################################
# Script Name: tableTest4.R
# Created by: Stephen Finnis
#
# Purpose:
# Prepare and format summary tables for annual status
# evaluations of salmon stock management units (SMUs),
# their conservation units (CUs), and Hatchery or Indicator
# stocks ("other" records). Cleans and joins raw inputs,
# checks for common data-entry issues, and builds tables.
#
# Inputs:
# - data/phase1culookup.csv: SMU-CU crosswalk (names and codes)
# - data/testData3.xlsx: Contains the SMU, CU, and (optionally) "other" sheets
#   Must include:
#   - Salmon_Outlook_Report
#   - cu_outlook_records
#   Optional:
#   - other_outlook_records
#
# Outputs:
# - tabPrep: combined and cleaned tibble
# - make_table(): function for generating styled output tables
############################################################

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

############################################################
# Prep Data: read in required files
# Read crosswalk file that gives CU names from CU codes
crosswalkList <- read.csv("data/phase1culookup.csv", stringsAsFactors = FALSE)

# Path to Excel workbook containing SMU, CU, and optional "other" data
dummyPath <- "data/testData3.xlsx"

# Read list of sheet names
sheet_names <- readxl::excel_sheets(dummyPath)

# Read each sheet into a list of data frames, named by sheet
df_list <- map(sheet_names, ~ read_excel(dummyPath, sheet = .x)) %>% set_names(sheet_names)

# Load each sheet into the global environment for easier reference
list2env(df_list, envir = .GlobalEnv)

############################################################
# Initial check: make sure required sheets are available
if (!exists("Salmon_Outlook_Report") || !exists("cu_outlook_records")) {
  stop("Required data frames Salmon_Outlook_Report or cu_outlook_records are missing.")
}
has_other <- exists("other_outlook_records")

############################################################
# Column Selection: only keep some column names
# Columns to keep from the SMU sheet
keep_cols_repeat <- c(
  "globalid", "uniquerowid",
  "smu_area", "smu_species", "smu_name",
  "outlook_narrative", "smu_outlook_assignment", "smu_prelim_forecast"
)

# Columns to keep from the CU sheet
keep_cols_cu <- c(
  "cu_outlook_selection", "cu_outlook_assignment", "cu_prelim_forecast", "cu_count",
  "parentrowid"
)

# Columns to keep from the "other" sheet
keep_cols_other <- c(
  "other_outlook_selection", "other_outlook_assignment", "other_prelim_forecast",
  "parentrowid"
)

# Subset SMU data to required columns
Outlook_Repeat_Test <- Salmon_Outlook_Report %>% select(all_of(keep_cols_repeat))

############################################################
# Process CU records (retain your original validations)
cu_outlook_records_proc <- cu_outlook_records %>%
  select(all_of(keep_cols_cu)) %>%
  mutate(
    # Case: absolutely no CU info entered
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
    # Case: CU assignment entered, but CU code never selected
    cu_outlook_selection = if_else(
      (is.na(cu_outlook_selection) | cu_outlook_selection == "") &
        (!is.na(cu_outlook_assignment) & cu_outlook_assignment != ""),
      "Outlook assigned but no CU specified",
      cu_outlook_selection
    ),
    # Add column cu_count from comma-separated lists
    cu_count = case_when(
      cu_outlook_selection %in% c("CHECK: No data entered", "Outlook assigned but no CU specified") ~ 1L,
      TRUE ~ str_count(coalesce(cu_outlook_selection, ""), ",") + 1L
    )
  )

# Bring SMU info into each CU row and tag source
cu_outlook_records_enriched <- full_join(
  cu_outlook_records_proc,
  Outlook_Repeat_Test,
  by = c("parentrowid" = "uniquerowid")
) %>%
  mutate(source_type = "CU")

############################################################
# NEW: Process other_outlook_records if present
other_outlook_records_enriched <- NULL
if (has_other) {
  other_outlook_records_enriched <- other_outlook_records %>%
    select(all_of(keep_cols_other)) %>%
    mutate(
      # Mirror CU validations but tailored for "other" labels
      other_outlook_selection = if_else(
        (is.na(other_outlook_selection) | other_outlook_selection == "") &
          (is.na(other_outlook_assignment) | other_outlook_assignment == ""),
        "CHECK: No data entered",
        other_outlook_selection
      ),
      other_outlook_assignment = if_else(
        (is.na(other_outlook_selection) | other_outlook_selection == "") &
          (is.na(other_outlook_assignment) | other_outlook_assignment == ""),
        "CHECK: No data entered",
        other_outlook_assignment
      ),
      # Assignment entered, but stock name not provided
      other_outlook_selection = if_else(
        (is.na(other_outlook_selection) | other_outlook_selection == "") &
          (!is.na(other_outlook_assignment) & other_outlook_assignment != ""),
        "Outlook assigned but no stock specified",
        other_outlook_selection
      ),
      # Count of "other" entries based on commas
      other_count = case_when(
        other_outlook_selection %in% c("CHECK: No data entered", "Outlook assigned but no stock specified") ~ 1L,
        TRUE ~ str_count(coalesce(other_outlook_selection, ""), ",") + 1L
      )
    ) %>%
    # Join to SMU info
    full_join(Outlook_Repeat_Test, by = c("parentrowid" = "uniquerowid")) %>%
    # Normalize column names into the CU schema so downstream logic can be reused
    rename(
      cu_outlook_selection = other_outlook_selection,
      cu_outlook_assignment = other_outlook_assignment,
      cu_prelim_forecast  = other_prelim_forecast,
      cu_count            = other_count
    ) %>%
    mutate(source_type = "Hatchery or Indicator Stock")
}

############################################################
# Combine CU and other records and add placeholder metrics
tabPrep_input <- bind_rows(
  cu_outlook_records_enriched,
  if (!is.null(other_outlook_records_enriched)) other_outlook_records_enriched
) %>%
  mutate(
    `Avg Run/Avg Spawners` = "50,000",
    `LRP/LBB` = "n/a",
    `Mgmt Target` = "10,000"
  )

############################################################
# Add names column:
# - For CU (codes), use crosswalk to map codes to labels
# - For "Hatchery or Indicator Stock", keep free text (no crosswalk mapping)
get_labels <- function(cu_string, ref_df) {
  if (is.na(cu_string) | cu_string == "") return(NA_character_)
  if (cu_string %in% c("CHECK: No data entered", "Outlook assigned but no CU specified")) return(cu_string)
  cu_codes <- str_split(cu_string, ",")[[1]] %>% str_trim() %>% toupper()
  labels <- ref_df %>%
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

tabPrep <- tabPrep_input %>%
  mutate(
    Unit_Names = case_when(
      source_type == "Hatchery or Indicator Stock" ~ coalesce(cu_outlook_selection, NA_character_),
      TRUE ~ map_chr(cu_outlook_selection, ~ get_labels(.x, crosswalkList))
    )
  )

############################################################
# Do data checks and set Resolution
empty_vals <- c(
  NA_character_, "", "N/A", "NA", "n/a", "na",
  "No data entered", "CHECK: No data entered"
)

tabPrep <- tabPrep %>%
  mutate(
    # Keep original outlook values
    smu_raw = smu_outlook_assignment,
    cu_raw  = cu_outlook_assignment,
    # Flags for later logic
    smu_empty   = smu_raw %in% empty_vals,
    cu_empty    = cu_raw %in% empty_vals,
    smu_dd      = tolower(coalesce(smu_raw, "")) == "data deficient",
    cu_dd       = tolower(coalesce(cu_raw, "")) == "data deficient",
    smu_numeric = suppressWarnings(!is.na(as.numeric(smu_raw))),
    cu_numeric  = suppressWarnings(!is.na(as.numeric(cu_raw)))
  ) %>%
  group_by(smu_name) %>%
  mutate(
    # Detect duplicate SMU rows (multiple parent IDs)
    smu_n_distinct_parent = n_distinct(parentrowid),
    smu_duplicate         = smu_n_distinct_parent > 1L,
    smu_name_display      = if_else(
      smu_duplicate,
      paste0(smu_name, " - CHECK: Data entered for the same SMU more than once"),
      smu_name
    ),
    # Recalculate CU count for consistency
    cu_count_calc = case_when(
      cu_empty ~ 0L,
      cu_outlook_selection %in% c("CHECK: No data entered", "Outlook assigned but no CU specified") ~ 1L,
      TRUE ~ str_count(coalesce(cu_outlook_selection, ""), ",") + 1L
    ),
    # Flag combinations that don't make sense
    needs_check = case_when(
      (smu_empty & cu_empty) ~ TRUE,
      (smu_numeric & cu_numeric) ~ TRUE,
      (smu_numeric & cu_dd) ~ TRUE,
      (smu_dd & cu_numeric) ~ TRUE,
      TRUE ~ FALSE
    ),
    # Human-readable explanations for odd cases
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
    # Final categorization of each row
    Resolution = case_when(
      # force label for other records
      source_type == "Hatchery or Indicator Stock" & !cu_empty ~ "Hatchery or Indicator Stock",
      # checks
      smu_empty & cu_empty ~ "CHECK: Only NA entered",
      smu_numeric & cu_numeric ~ "CHECK: SMU and CU both numeric",
      smu_numeric & cu_dd ~ "CHECK: SMU numeric but CU data deficient",
      smu_dd & cu_numeric ~ "CHECK: SMU data deficient but CU numeric",
      # regular SMU/CU cases
      cu_empty ~ "SMU",
      !cu_empty & cu_count == 1 ~ "CU (singular)",
      !cu_empty & cu_count > 1 ~ "CU (aggregate)",
      TRUE ~ "CHECK: unexpected combination"
    )
  )

############################################################
# Combine subordinate Forecasts/Outlooks and finalize display columns
tabPrep <- tabPrep %>%
  group_by(smu_name_display, Resolution, smu_prelim_forecast, smu_outlook_assignment, outlook_narrative) %>%
  mutate(
    CU_Forecast = paste(na.omit(cu_prelim_forecast), collapse = ", "),
    CU_Outlook  = paste(na.omit(cu_outlook_assignment), collapse = ", "),
    CU_CodeList = paste(na.omit(cu_outlook_selection), collapse = ", ")
  ) %>%
  ungroup() %>%
  mutate(
    # Name displayed depends on Resolution
    Name = case_when(
      Resolution == "SMU" ~ smu_name_display,
      Resolution %in% c("CU (aggregate)", "CU (singular)", "Hatchery or Indicator Stock") ~ Unit_Names,
      str_detect(Resolution, "^CHECK:") ~ Unit_Names,
      TRUE ~ NA_character_
    ),
    # Forecast shown in table
    Forecast = case_when(
      Resolution == "SMU" ~ smu_prelim_forecast,
      Resolution %in% c("CU (aggregate)", "CU (singular)", "Hatchery or Indicator Stock") ~ CU_Forecast,
      str_detect(Resolution, "^CHECK:") ~ paste0(
        "CHECK VALUE: Both SMU and subordinate forecasts entered. SMU = ",
        coalesce(smu_prelim_forecast, "NA"),
        "; CU/Other = ",
        coalesce(CU_Forecast, "NA"),
        " [", CU_CodeList, "]"
      ),
      TRUE ~ NA_character_
    ),
    # Outlook shown in table
    Outlook = case_when(
      Resolution == "SMU" ~ coalesce(smu_outlook_assignment, NA_character_),
      Resolution %in% c("CU (aggregate)", "CU (singular)", "Hatchery or Indicator Stock") ~ coalesce(cu_outlook_assignment, NA_character_),
      str_detect(Resolution, "^CHECK:") ~ paste0(
        "CHECK VALUE: Both SMU and subordinate outlooks entered. SMU = ",
        coalesce(smu_outlook_assignment, "NA"),
        "; CU/Other = ",
        coalesce(CU_Outlook, "NA"),
        " [", CU_CodeList, "]"
      ),
      TRUE ~ NA_character_
    ),
    Narrative = outlook_narrative
  )

############################################################
# Final Table Prep
tabPrep <- tabPrep %>%
  distinct(
    smu_area, smu_species, smu_name_display, Narrative, Resolution, Name,
    `Avg Run/Avg Spawners`, `LRP/LBB`, `Mgmt Target`, Forecast, Outlook
  ) %>%
  rename(smu_name = smu_name_display) %>%
  mutate(
    Forecast = coalesce(Forecast, NA_character_)
    # Respect preference: do not convert hyphen to en dash.
    # If you want to normalize spacing around hyphens, uncomment:
    # %>% str_replace_all("\\s*-\\s*", " - ")
  ) %>%
  select(
    smu_area, smu_species, smu_name, Narrative, Resolution, Name,
    `Avg Run/Avg Spawners`, `LRP/LBB`, `Mgmt Target`, Forecast, Outlook
  )

############################################################
# Table-Building Functions

# Build a block of rows for one SMU (narrative rows + data rows)
build_block <- function(df_smu) {
  smu <- unique(df_smu$smu_name)
  narr <- unique(df_smu$Narrative)

  # Header-style narrative rows
  row1 <- data.frame(
    Resolution = "SMU", Name = "Narrative",
    `Avg Run/Avg Spawners` = "", `LRP/LBB` = "", `Mgmt Target` = "",
    Forecast = "", Outlook = "",
    check.names = FALSE, stringsAsFactors = FALSE
  )
  row2 <- data.frame(
    Resolution = smu, Name = narr,
    `Avg Run/Avg Spawners` = "", `LRP/LBB` = "", `Mgmt Target` = "",
    Forecast = "", Outlook = "",
    check.names = FALSE, stringsAsFactors = FALSE
  )
  # Column label row
  row3 <- data.frame(
    Resolution = "Resolution", Name = "Name",
    `Avg Run/Avg Spawners` = "Avg Run/Avg Spawners",
    `LRP/LBB` = "LRP/LBB", `Mgmt Target` = "Mgmt Target",
    Forecast = "Forecast", Outlook = "Outlook",
    check.names = FALSE, stringsAsFactors = FALSE
  )
  # Actual content rows
  data_rows <- df_smu %>%
    select(Resolution, Name, `Avg Run/Avg Spawners`, `LRP/LBB`, `Mgmt Target`, Forecast, Outlook)

  bind_rows(row1, row2, row3, data_rows)
}

# Apply flextable styling rules to full combined table
style_smu_table <- function(big_df) {
  ft <- flextable(big_df)
  ft <- delete_part(ft, part = "header")

  # Identify narrative rows and header rows
  # FIX: use vectorized '&' not scalar '&&'
  idx_label  <- which(big_df$Resolution == "SMU" & big_df$Name == "Narrative")
  idx_header <- which(big_df$Resolution == "Resolution")

  # Apply simple background and bold styles (only if indices are present)
  if (length(idx_label) > 0) {
    ft <- bg(ft, i = idx_label,  bg = "gray90")
    ft <- bold(ft, i = idx_label, bold = TRUE)
  }
  if (length(idx_header) > 0) {
    ft <- bg(ft, i = idx_header, bg = "gray90")
    ft <- bold(ft, i = idx_header, bold = TRUE)
  }

  # Add borders
  ft <- border_remove(ft)
  ft <- border_outer(ft, border = fp_border(color = "black", width = 1))
  ft <- border_inner_h(ft, border = fp_border(color = "black", width = 1))
  ft <- border_inner_v(ft, border = fp_border(color = "black", width = 1))

  # Merge cells for narrative rows
  for (i in idx_label) {
    ft <- merge_at(ft, i = i,     j = 2:7)
    ft <- merge_at(ft, i = i + 1, j = 2:7)
  }

  # Add thick separators between SMUs
  if (length(idx_label) > 1) {
    idx_separators <- idx_label[-1]
    for (i in idx_separators) {
      ft <- border(ft, i = i, j = 1:ncol(big_df),
                   border.top = fp_border(color = "black", width = 2), part = "body")
    }
  }

  ft <- autofit(ft)
  ft <- set_table_properties(ft, layout = "autofit")
  ft
}

# Build descriptive caption text
make_caption <- function(species, area, year = format(Sys.Date(), "%Y")) {
  area_titleCase <- area %>%
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

# Main function: create formatted table for one area × species
make_table <- function(area, species, data = tabPrep) {
  df_filtered <- data %>% filter(smu_area == area, smu_species == species)
  if (nrow(df_filtered) == 0) stop(paste("No data found for:", area, "/", species))

  # Build separate blocks for each SMU
  smu_list   <- split(df_filtered, df_filtered$smu_name)
  block_list <- lapply(smu_list, build_block)
  big_df     <- bind_rows(block_list)

  # Style table and attach caption
  ft <- style_smu_table(big_df)
  caption_text <- make_caption(species = species, area = area)
  ft <- set_caption(ft, caption = caption_text)
  ft
}

############################################################
# Example usage
make_table("FRASER AND INTERIOR", "Chinook")
make_table("SOUTH COAST", "Chum")
