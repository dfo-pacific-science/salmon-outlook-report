### Code for practicing creating flextables

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
## Code for making Table 1 with Outlook classes

# Read raw sheet without interpreting the first row as names
raw = read_excel("data/outlookClasses.xlsx", sheet = 1, col_names = FALSE)

# Extract top and bottom header rows (first two rows of the sheet)
header_top = as.character(unlist(raw[1, ], use.names = FALSE))
header_bottom = as.character(unlist(raw[2, ], use.names = FALSE))

# Body is everything after those two header rows
df_body = raw[-c(1,2), , drop = FALSE]

# Give the body sensible column names using the 'bottom' header,
# making unique names if needed (avoids duplicate-names issues)
colnames(df_body) = make.unique(header_bottom)

# Build the flextable and set a two-row header using set_header_df
mapping = data.frame(
  col_keys = colnames(df_body),
  top = header_top,
  bottom = header_bottom,
  stringsAsFactors = FALSE
)

ft = flextable(df_body) %>%
  set_header_df(mapping = mapping, key = "col_keys") %>%
  # merge cells horizontally where top header labels are the same
  merge_h(part = "header") %>%
  # formatting: remove default inner borders then add the three main horizontal rules
  merge_v(part = "header", j = 1) %>%
  border_remove() %>%
  hline_top(part = "header", border = fp_border(color = "black", width = 1)) %>%
  hline_bottom(part = "header", border = fp_border(color = "black", width = 1)) %>%
  hline_bottom(part = "body",   border = fp_border(color = "black", width = 1)) %>%
  # aesthetic touches
  bold(part = "header") %>%
  align(align = "center", part = "header") %>%
  align(j = 1, align = "left", part = "all") %>%   # left-align first column
  autofit()

# 7) Show the table
ft

#################################################################################
################################################################################
## Code for making tables that give outlooks/narratives for each SMU

## Prep the data
# Current format has the data in an Excel file divided into 3 sheets
# One is data for each SMU, one is data for CUs, one is data for Other (hatchery/indicator)
# These need to be merged together in various ways

# For now, I have practice data stored here
dummyPath = "data/finnis_outlook_phase1_test_dummy_data_20250905.xlsx"


# List all sheet names.
sheet_names = readxl::excel_sheets(dummyPath)

# Read each sheet and name the list elements
df_list = map(sheet_names, ~ read_excel(dummyPath, sheet = .x)) %>%
  set_names(sheet_names)

# Export each dataframe to the global environment
list2env(df_list, envir = .GlobalEnv)

# Columns we want to keep (now using parentrowid instead of uniquerowid)
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

Outlook_Repeat_Test = Outlook_Repeat_Test %>%
  select(all_of(keep_cols_repeat))

cu_outlook_records = cu_outlook_records %>%
  select(all_of(keep_cols_cu))

other_outlook_records = other_outlook_records %>%
  select(all_of(keep_cols_other))



# Step 2: Join smu_area, smu_species, smu_name from Outlook_Repeat_Test to cu_outlook_records
cu_outlook_records_enriched = cu_outlook_records %>%
  left_join(
    Outlook_Repeat_Test,
      #select(uniquerowid, smu_area, smu_species, smu_name, outlook_narrative),
    by = c("parentrowid" = "uniquerowid")
  )

# Step 2: Join smu_area, smu_species, smu_name from Outlook_Repeat_Test to cu_outlook_records
other_outlook_records_enriched = other_outlook_records %>%
  left_join(
    Outlook_Repeat_Test %>%
      select(uniquerowid, smu_area, smu_species, smu_name),
    by = c("parentrowid" = "uniquerowid")
  )


#test = full_join(Outlook_Repeat_Test, cu_outlook_records_enriched, by = c("uniquerowid" = "parentrowid"))

stacked_df = bind_rows(Outlook_Repeat_Test, cu_outlook_records_enriched, other_outlook_records_enriched)


################################################################################
## Prep the data a bit more to make tables for the report

# Tables will be labelled by DFO Area and Species (for later reporting)

# Add fake data to columns for final tables
# These values will come from the second survey
tabPrep = cu_outlook_records_enriched %>%
  mutate(
    `Avg Run/Avg Spawners` = "50,000",
    `LRP/LBB`        = "n/a",
    `Mgmt Target`    = "10,000"
    )




tabPrep <- tabPrep %>%
  group_by(smu_name) %>%
  mutate(
    # Determine resolution type
    Resolution = case_when(
      n() == 1 & cu_count == 1 ~ "SMU",
      n() > 1 & cu_count > 1 ~ "CU (aggregate)",
      n() > 1 & cu_count == 1 ~ "CU (singular)",
      n() == 1 & cu_count > 1 ~ paste0("CHECK VALUE — single SMU but cu_count = ", cu_count),
      TRUE ~ "CHECK VALUE — unexpected combination"
    )
  ) %>%
  ungroup() %>%
  mutate(
    # Assign Name based on resolution
    Name = case_when(
      Resolution == "SMU" ~ smu_name,
      Resolution %in% c("CU (aggregate)", "CU (singular)") ~ cu_outlook_selection,
      str_detect(Resolution, "CHECK VALUE") ~ "Check value",
      TRUE ~ NA_character_
    ),

    # Assign Forecast based on resolution
    Forecast = case_when(
      Resolution == "SMU" ~ smu_prelim_forecast,
      Resolution %in% c("CU (aggregate)", "CU (singular)") ~ cu_prelim_forecast,
      str_detect(Resolution, "CHECK VALUE") ~ "Check value",
      TRUE ~ NA_character_
    ),

    # Assign Outlook based on resolution
    Outlook = case_when(
      Resolution == "SMU" ~ smu_outlook_assignment,
      Resolution %in% c("CU (aggregate)", "CU (singular)") ~ cu_outlook_assignment,
      str_detect(Resolution, "CHECK VALUE") ~ "Check value",
      TRUE ~ NA_character_
    )
  )



# Fix up the Forecasts
tabPrep = tabPrep %>%
  mutate(
    Forecast = Forecast %>%
      # Replace hyphen with en dash
      str_replace_all("-", "–") %>%
      # Remove extra spaces around the dash
      str_replace_all("\\s*–\\s*", "–")
  )



# Prepare the data
tabPrep = tabPrep %>%
  rename(Narrative = outlook_narrative) %>%
  select(
    smu_area,
    smu_species,
    smu_name,
    Narrative,
    Resolution,
    Name,
    `Avg Run/Avg Spawners`,
    `LRP/LBB`,
    `Mgmt Target`,
    Forecast,
    Outlook
  )

# Nest the data by smu_area and smu_species
nested_tabs = tabPrep %>%
  group_by(smu_area, smu_species, smu_name, Narrative) %>%
  group_split()

# Create flextables
ft_list = lapply(nested_tabs, function(df_smu) {
  smu = unique(df_smu$smu_name)
  narrative = unique(df_smu$Narrative)
  area = unique(df_smu$smu_area)
  species = unique(df_smu$smu_species)

  df_smu_clean = df_smu %>% select(-smu_area, -smu_species, -smu_name, -Narrative)
  n_cols = ncol(df_smu_clean)

  ft = flextable(df_smu_clean) %>%
    add_header_row(values = c(smu, narrative), colwidths = c(1, n_cols - 1)) %>%
    add_header_row(values = c("SMU", "Narrative"), colwidths = c(1, n_cols - 1)) %>%
    bg(i = 1, bg = "grey90", part = "header") %>%
    bold(i = 1, bold = TRUE, part = "header") %>%
    color(i = 1, color = "black", part = "header") %>%
    bg(i = 2, bg = "white", part = "header") %>%
    bold(i = 2, bold = FALSE, part = "header") %>%
    bg(i = 3, bg = "grey90", part = "header") %>%
    bold(i = 3, bold = TRUE, part = "header") %>%
    border_remove() %>%
    border_outer(border = fp_border(color = "black", width = 1), part = "all") %>%
    border_inner_h(border = fp_border(color = "black", width = 1), part = "all") %>%
    border_inner_v(border = fp_border(color = "black", width = 1), part = "all") %>%
    autofit() %>%
    set_table_properties(layout = "autofit")

  list(area = area, species = species, table = ft)
})


save(ft_list, file = "ft_list.RData")
