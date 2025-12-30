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
## Prep the data

crosswalkList = read.csv("data/phase1culookup.csv")
dummyPath = "data/testData3.xlsx"

sheet_names = readxl::excel_sheets(dummyPath)
df_list = map(sheet_names, ~ read_excel(dummyPath, sheet = .x)) %>%
  set_names(sheet_names)
list2env(df_list, envir = .GlobalEnv)

# Columns to keep
keep_cols_repeat = c("globalid", "uniquerowid",
                     "smu_area", "smu_species", "smu_name",
                     "outlook_narrative", "smu_outlook_assignment", "smu_prelim_forecast")
keep_cols_cu = c("cu_outlook_selection", "cu_outlook_assignment", "cu_prelim_forecast", "cu_count",
                 "parentrowid")
keep_cols_other = c("other_outlook_selection", "other_outlook_assignment", "other_prelim_forecast",
                    "parentrowid")

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

# other_outlook_records = other_outlook_records %>%
#   select(all_of(keep_cols_other))

# Join with SMU info
cu_outlook_records_enriched = full_join(Outlook_Repeat_Test, cu_outlook_records,
          by = c("uniquerowid" = "parentrowid"))

# other_outlook_records_enriched = other_outlook_records %>%
#   left_join(Outlook_Repeat_Test %>% select(uniquerowid, smu_area, smu_species, smu_name),
#             by = c("parentrowid" = "uniquerowid"))
#
# stacked_df = bind_rows(Outlook_Repeat_Test, cu_outlook_records_enriched, other_outlook_records_enriched)

################################################################################
## Prep final table
tabPrep = cu_outlook_records_enriched %>%
  mutate(
    `Avg Run/Avg Spawners` = "50,000",
    `LRP/LBB`        = "n/a",
    `Mgmt Target`    = "10,000"
  )

# Replace CU codes with names
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
## SMU/CU checking & Resolution logic
tabPrep = tabPrep %>%
  group_by(smu_name) %>%
  mutate(
    # Identify Data Deficient and empty entries
    smu_dd = tolower(smu_outlook_assignment) == "data deficient",
    cu_dd = tolower(cu_outlook_assignment) == "data deficient",
    smu_empty = is.na(smu_outlook_assignment) | smu_outlook_assignment %in% c("", "N/A", "NA", "n/a", "na"),
    cu_empty = is.na(cu_outlook_assignment) | cu_outlook_assignment %in% c("", "N/A", "NA", "n/a", "na"),

    # Need check if problematic combination
    needs_check = case_when(
      smu_empty & cu_empty ~ TRUE,
      (!smu_empty & !cu_empty & !smu_dd & !cu_dd & smu_outlook_assignment != cu_outlook_assignment) ~ TRUE,
      smu_dd & !cu_dd ~ TRUE,
      !smu_dd & cu_dd ~ TRUE,
      TRUE ~ FALSE
    ),

    # Construct reason message
    check_reason = case_when(
      smu_empty & cu_empty ~ "Only NA entered",
      (!smu_empty & !cu_empty & !smu_dd & !cu_dd & smu_outlook_assignment != cu_outlook_assignment) ~
        paste0("SMU and CU values differ"),
      smu_dd & !cu_dd ~ "SMU Data Deficient but CU entered",
      !smu_dd & cu_dd ~ "CU Data Deficient but SMU entered",
      TRUE ~ NA_character_
    ),

    # Aggregate all CU values for CHECK VALUE display
    cu_all = paste(cu_outlook_assignment, collapse = ", ")
  ) %>%
  ungroup() %>%
  group_by(smu_name) %>%
  mutate(
    # Determine SMU/CU Resolution
    Resolution = case_when(
      needs_check ~ paste0("CHECK VALUE: Both SMU and CU outlooks entered. SMU = ", smu_outlook_assignment[1],
                           ", CU = [", paste(cu_outlook_assignment[!is.na(cu_outlook_assignment)], collapse = ", "), "]"),
      n() == 1 & !smu_empty ~ "SMU",
      n() > 1 & cu_count > 1 & !cu_empty ~ "CU (aggregate)",
      n() > 1 & cu_count == 1 & !cu_empty ~ "CU (singular)",
      TRUE ~ "Unknown"
    ),
    # Name, Forecast, Outlook
    Name = case_when(
      Resolution == "SMU" ~ smu_name,
      Resolution %in% c("CU (aggregate)", "CU (singular)") ~ CU_Names,
      str_detect(Resolution, "CHECK VALUE") ~ "Check value",
      TRUE ~ NA_character_
    ),
    Forecast = case_when(
      Resolution == "SMU" ~ smu_prelim_forecast,
      Resolution %in% c("CU (aggregate)", "CU (singular)") ~ cu_prelim_forecast,
      str_detect(Resolution, "CHECK VALUE") ~ paste0(
        "SMU Forecast = ", coalesce(smu_prelim_forecast[1], "NA"), "; CU Forecast = ",
        paste(coalesce(cu_prelim_forecast[!is.na(cu_prelim_forecast)], "NA"), collapse = ", ")
      ),
      TRUE ~ NA_character_
    ),
    Outlook = case_when(
      Resolution == "SMU" ~ smu_outlook_assignment,
      Resolution %in% c("CU (aggregate)", "CU (singular)") ~ cu_outlook_assignment,
      str_detect(Resolution, "CHECK VALUE") ~ paste0(
        "SMU = ", coalesce(smu_outlook_assignment[1], "NA"), "; CU = [",
        paste(coalesce(cu_outlook_assignment[!is.na(cu_outlook_assignment)], "NA"), collapse = ", "), "]"
      ),
      TRUE ~ NA_character_
    )
  ) %>%
  ungroup() %>%
  # Clean Forecast formatting
  mutate(Forecast = str_replace_all(Forecast, "-", "–") %>%
           str_replace_all("\\s*–\\s*", "–")) %>%
  rename(Narrative = outlook_narrative) %>%
  select(
    smu_area, smu_species, smu_name, Narrative, Resolution, Name,
    `Avg Run/Avg Spawners`, `LRP/LBB`, `Mgmt Target`,
    Forecast, Outlook
  )


################################################################################
## Function to build SMU blocks
build_block = function(df_smu) {
  smu  = unique(df_smu$smu_name)
  narr = unique(df_smu$Narrative)

  row1 = data.frame(
    Resolution = "SMU",
    Name       = "Narrative",
    `Avg Run/Avg Spawners` = "",
    `LRP/LBB` = "",
    `Mgmt Target` = "",
    Forecast   = "",
    Outlook    = "",
    check.names = FALSE
  )

  row2 = data.frame(
    Resolution = smu,
    Name       = narr,
    `Avg Run/Avg Spawners` = "",
    `LRP/LBB` = "",
    `Mgmt Target` = "",
    Forecast   = "",
    Outlook    = "",
    check.names = FALSE
  )

  row3 = data.frame(
    Resolution = "Resolution",
    Name       = "Name",
    `Avg Run/Avg Spawners` = "Avg Run/Avg Spawners",
    `LRP/LBB` = "LRP/LBB",
    `Mgmt Target` = "Mgmt Target",
    Forecast   = "Forecast",
    Outlook    = "Outlook",
    check.names = FALSE
  )

  data_rows = df_smu %>%
    select(Resolution, Name, `Avg Run/Avg Spawners`,
           `LRP/LBB`, `Mgmt Target`, Forecast, Outlook)

  bind_rows(row1, row2, row3, data_rows)
}

################################################################################
## Style table function
style_smu_table = function(big_df) {
  ft = flextable(big_df) %>% delete_part(part = "header")
  idx_label  = which(big_df$Resolution == "SMU" & big_df$Name == "Narrative")
  idx_header = which(big_df$Resolution == "Resolution")
  ft = bg(ft, i = idx_label, bg = "gray90") %>% bold(i = idx_label, bold = TRUE)
  ft = bg(ft, i = idx_header, bg = "gray90") %>% bold(i = idx_header, bold = TRUE)
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
      ft <- border(ft, i = i, j = 1:ncol(big_df), border.top = fp_border(color="black", width=2),
                   part = "body")
    }
  }
  ft = autofit(ft)
  ft = set_table_properties(ft, layout = "autofit")
  ft
}

################################################################################
## Caption
make_caption = function(species, area, year){
  area_titleCase = area %>% tolower() %>% tools::toTitleCase() %>% gsub("\\bAnd\\b", "and", .)
  paste0(
    "Summary of metrics informing the annual status evaluation for ",
    species, " in the ", area_titleCase, " area during the ", year, " management cycle. ",
    "Values are presented for each stock management unit (SMU), and where applicable, for associated singular conservation units (CUs), CU aggregations, and Hatchery or Indicator stocks. ",
    "Reported values include recent average run size or spawner abundance, biological reference points (limit reference point [LRP] and lower biological benchmark [LBB]), and management (mgmt.) targets or operational control points used to interpret current conditions. ",
    "The forecast provides a numerical abundance estimate for the return year, while the outlook indicates the categorical status classification (see definitions in Table 1)."
  )
}

################################################################################
## Main table function
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

################################################################################
# Example test
make_table("FRASER AND INTERIOR", "Chum")
make_table("SOUTH COAST", "Chinook")
make_table("SOUTH COAST", "Sockeye Lake Type")
make_table("YUKON TRANSBOUNDARY", "Chinook")
