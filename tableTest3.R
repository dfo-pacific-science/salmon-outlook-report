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
## Code for making tables that give outlooks/narratives for each SMU
################################################################################

## Prep the data

# Crosswalk with full CU names
crosswalkList = read.csv("data/phase1culookup.csv")

# Input Excel file
dummyPath = "data/testData3.xlsx"
sheet_names = readxl::excel_sheets(dummyPath)

# Read sheets into list and export
df_list = map(sheet_names, ~ read_excel(dummyPath, sheet = .x)) %>%
  set_names(sheet_names)
list2env(df_list, envir = .GlobalEnv)

# Columns to keep
keep_cols_repeat = c(
  "globalid", "uniquerowid",
  "smu_area", "smu_species", "smu_name",
  "outlook_narrative", "smu_outlook_assignment", "smu_prelim_forecast"
)
keep_cols_cu = c(
  "cu_outlook_selection", "cu_outlook_assignment", "cu_prelim_forecast", "cu_count",
  "parentrowid"
)
keep_cols_other = c(
  "other_outlook_selection", "other_outlook_assignment", "other_prelim_forecast", "parentrowid"
)

# SMU data
Outlook_Repeat_Test = Salmon_Outlook_Report %>%
  select(all_of(keep_cols_repeat))

# CU data checks
cu_outlook_records <- cu_outlook_records %>%
  select(all_of(keep_cols_cu)) %>%
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
        !(is.na(cu_outlook_assignment) | cu_outlook_assignment == ""),
      "Outlook assigned but no CU specified",
      cu_outlook_selection
    ),
    cu_count = case_when(
      cu_outlook_selection %in% c("CHECK: No data entered", "Outlook assigned but no CU specified") ~ 1L,
      TRUE ~ str_count(cu_outlook_selection, ",") + 1L
    )
  )

# Hatchery/other
# other_outlook_records = other_outlook_records %>%
#   select(all_of(keep_cols_other))

# Enrich CU and other data with SMU info
cu_outlook_records_enriched = full_join(
  cu_outlook_records,
  Outlook_Repeat_Test,
  by = c("parentrowid" = "uniquerowid")
)

# other_outlook_records_enriched = full_join(
#   other_outlook_records,
#   Outlook_Repeat_Test %>%
#     select(uniquerowid, smu_area, smu_species, smu_name),
#   by = c("parentrowid" = "uniquerowid")
# )
#
# stacked_df = bind_rows(Outlook_Repeat_Test, cu_outlook_records_enriched, other_outlook_records_enriched)

################################################################################
## Prep table for report
################################################################################

tabPrep = cu_outlook_records_enriched %>%
  mutate(
    `Avg Run/Avg Spawners` = "50,000",
    `LRP/LBB` = "n/a",
    `Mgmt Target` = "10,000"
  )

# Function to replace CU codes with names
get_labels = function(cu_string, ref_df) {
  cu_codes = str_split(cu_string, ",")[[1]]
  labels = ref_df %>%
    filter(cu %in% cu_codes) %>%
    arrange(match(cu, cu_codes)) %>%
    pull(label)
  paste(labels, collapse = ", ")
}

tabPrep = tabPrep %>%
  mutate(CU_Names = map_chr(cu_outlook_selection, ~ get_labels(.x, crosswalkList)))

################################################################################
## SMU/CU checks and error logic
################################################################################

tabPrep = tabPrep %>%
  group_by(smu_name) %>%
  mutate(
    # Detect empty/out of bounds
    smu_empty = is.na(smu_outlook_assignment) | smu_outlook_assignment %in% c("", "N/A", "NA", "n/a", "na"),
    cu_empty  = is.na(cu_outlook_assignment)  | cu_outlook_assignment  %in% c("", "N/A", "NA", "n/a", "na"),
    smu_dd    = tolower(smu_outlook_assignment) == "data deficient",
    cu_dd     = tolower(cu_outlook_assignment)  == "data deficient",
    smu_numeric = suppressWarnings(!is.na(as.numeric(smu_outlook_assignment))),
    cu_numeric  = suppressWarnings(!is.na(as.numeric(cu_outlook_assignment))),
    needs_check = case_when(
      smu_empty & cu_empty ~ TRUE,
      smu_numeric & cu_numeric ~ TRUE,
      smu_numeric & cu_dd ~ TRUE,
      smu_dd & cu_numeric ~ TRUE,
      TRUE ~ FALSE
    ),
    check_reason = case_when(
      smu_empty & cu_empty ~ "Only NA entered",
      smu_numeric & cu_numeric ~ "SMU and CU both numeric",
      smu_numeric & cu_dd ~ "SMU numeric but CU data deficient",
      smu_dd & cu_numeric ~ "CU numeric but SMU data deficient",
      TRUE ~ NA_character_
    )
  ) %>%
  ungroup()

################################################################################
## Resolution assignment
################################################################################

tabPrep = tabPrep %>%
  group_by(smu_name) %>%
  mutate(
    smu_has_value = any(!is.na(smu_outlook_assignment) & smu_outlook_assignment != "" & tolower(smu_outlook_assignment) != "n/a"),
    cu_has_value  = any(!is.na(cu_outlook_assignment)  & cu_outlook_assignment  != ""),
    needs_check_original = smu_has_value & cu_has_value
  ) %>%
  ungroup() %>%
  mutate(
    # Determine Resolution
    Resolution = case_when(
      needs_check ~ "CHECK VALUE — SMU + CU outlooks assigned",
      smu_has_value & !cu_has_value ~ "SMU",
      cu_has_value & !smu_has_value & cu_count > 1 ~ "CU (aggregate)",
      cu_has_value & !smu_has_value & cu_count == 1 ~ "CU (singular)",
      TRUE ~ "CHECK VALUE — unexpected combination"
    )
  )

################################################################################
## Construct Name, Forecast, Outlook with CU names and aggregation
################################################################################

tabPrep = tabPrep %>%
  group_by(smu_name, Resolution, smu_prelim_forecast, smu_outlook_assignment) %>%
  mutate(
    # Combine CU forecasts/outlooks for one row per Resolution
    CU_Forecast = if (any(Resolution %in% c("CU (aggregate)", "CU (singular)", "CHECK VALUE — SMU + CU outlooks assigned"))) {
      paste(cu_prelim_forecast, collapse = ", ")
    } else NA_character_,
    CU_Outlook  = if (any(Resolution %in% c("CU (aggregate)", "CU (singular)", "CHECK VALUE — SMU + CU outlooks assigned"))) {
      paste0("[", paste(CU_Names, "=", cu_outlook_assignment, collapse = ", "), "]")
    } else NA_character_
  ) %>%
  ungroup() %>%
  mutate(
    # Name column
    Name = case_when(
      Resolution == "SMU" ~ smu_name,
      Resolution %in% c("CU (aggregate)", "CU (singular)") ~ CU_Names,
      str_detect(Resolution, "CHECK VALUE") ~ CU_Names,
      TRUE ~ NA_character_
    ),
    # Forecast
    Forecast = case_when(
      Resolution == "SMU" ~ smu_prelim_forecast,
      Resolution %in% c("CU (aggregate)", "CU (singular)") ~ CU_Forecast,
      str_detect(Resolution, "CHECK VALUE") ~ paste0(
        "CHECK VALUE: Both SMU and CU forecasts entered. ",
        "SMU = ", coalesce(smu_prelim_forecast, "NA"), "; CU = ", CU_Forecast
      ),
      TRUE ~ NA_character_
    ),
    # Outlook
    Outlook = case_when(
      Resolution == "SMU" ~ smu_outlook_assignment,
      Resolution %in% c("CU (aggregate)", "CU (singular)") ~ cu_outlook_assignment,
      str_detect(Resolution, "CHECK VALUE") ~ paste0(
        "CHECK VALUE: Both SMU and CU outlooks entered. SMU = ", coalesce(smu_outlook_assignment, "NA"),
        "; CU = ", CU_Outlook
      ),
      TRUE ~ NA_character_
    )
  )

################################################################################
## Remove duplicates (after combining CU values)
################################################################################

tabPrep = tabPrep %>%
  distinct(
    smu_area, smu_species, smu_name, Narrative, Resolution, Name,
    `Avg Run/Avg Spawners`, `LRP/LBB`, `Mgmt Target`, Forecast, Outlook
  )

################################################################################
## Clean Forecast formatting
################################################################################

tabPrep = tabPrep %>%
  mutate(
    Forecast = str_replace_all(Forecast, "-", "–") %>%
      str_replace_all("\\s*–\\s*", "–")
  ) %>%
  rename(Narrative = outlook_narrative) %>%
  select(
    smu_area, smu_species, smu_name, Narrative, Resolution, Name,
    `Avg Run/Avg Spawners`, `LRP/LBB`, `Mgmt Target`, Forecast, Outlook
  )

################################################################################
## Table-building functions
################################################################################

build_block = function(df_smu) {
  smu  = unique(df_smu$smu_name)
  narr = unique(df_smu$Narrative)

  row1 = data.frame(
    Resolution = "SMU", Name = "Narrative",
    `Avg Run/Avg Spawners` = "", `LRP/LBB` = "", `Mgmt Target` = "",
    Forecast = "", Outlook = "", check.names = FALSE
  )
  row2 = data.frame(
    Resolution = smu, Name = narr,
    `Avg Run/Avg Spawners` = "", `LRP/LBB` = "", `Mgmt Target` = "",
    Forecast = "", Outlook = "", check.names = FALSE
  )
  row3 = data.frame(
    Resolution = "Resolution", Name = "Name",
    `Avg Run/Avg Spawners` = "Avg Run/Avg Spawners",
    `LRP/LBB` = "LRP/LBB", `Mgmt Target` = "Mgmt Target",
    Forecast = "Forecast", Outlook = "Outlook", check.names = FALSE
  )
  data_rows = df_smu %>%
    select(Resolution, Name, `Avg Run/Avg Spawners`, `LRP/LBB`, `Mgmt Target`, Forecast, Outlook)
  bind_rows(row1, row2, row3, data_rows)
}

style_smu_table = function(big_df) {
  ft = flextable(big_df)
  ft = delete_part(ft, part = "header")
  idx_label  = which(big_df$Resolution == "SMU" & big_df$Name == "Narrative")
  idx_header = which(big_df$Resolution == "Resolution")
  ft = bg(ft, i = idx_label, bg = "gray90")
  ft = bold(ft, i = idx_label, bold = TRUE)
  ft = bg(ft, i = idx_header, bg = "gray90")
  ft = bold(ft, i = idx_header, bold = TRUE)
  ft = border_remove(ft)
  ft = border_outer(ft, border = fp_border(color="black", width=1))
  ft = border_inner_h(ft, border = fp_border(color="black", width=1))
  ft = border_inner_v(ft, border = fp_border(color="black", width=1))
  for (i in idx_label) {
    ft = merge_at(ft, i = i, j = 2:7)
    ft = merge_at(ft, i = i + 1, j = 2:7)
  }
  if (length(idx_label) > 1) {
    idx_separators <- idx_label[-1]
    for (i in idx_separators) {
      ft <- border(ft, i = i, j = 1:ncol(big_df), border.top = fp_border(color="black", width=2), part = "body")
    }
  }
  ft = autofit(ft)
  ft = set_table_properties(ft, layout = "autofit")
  ft
}

make_caption = function(species, area, year) {
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

make_table = function(area, species, data = tabPrep) {
  df_filtered = data %>% filter(smu_area == area, smu_species == species)
  if (nrow(df_filtered) == 0) stop(paste("No data found for:", area, "/", species))
  smu_list = split(df_filtered, df_filtered$smu_name)
  block_list = lapply(smu_list, build_block)
  big_df = bind_rows(block_list)
  ft = style_smu_table(big_df)
  caption_text = make_caption(species = species, area = area, year = 2026)
  ft = set_caption(ft, caption = caption_text)
  ft
}

# Test tables
make_table("FRASER AND INTERIOR", "Chinook")
make_table("SOUTH COAST", "Chinook")