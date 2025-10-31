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


#### STEPHEN REMOVE THIS BUT JUST FOR TESTING
tabPrep = tabPrep%>%
  mutate(
    smu_area = "SOUTH COAST",
    smu_species = "Chum"
  )



tabPrep = tabPrep %>%
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

################################################################################
## NOW TEST ADDING DOING EVERYTHING WITH MY DATA

## RESTRUCTURE THE DATA FRAME
# To get the format we want, the data needs to be rearranged
# Where SMU and Narrative are placed above the remaining data columns

# ---- Split into SMU blocks ----
smu_list = split(tabPrep, tabPrep$smu_name)

# ---- Build each SMU block as a data frame with fake header rows ----
build_block = function(df_smu) {

  smu = unique(df_smu$smu_name)
  narr = unique(df_smu$Narrative)

  # Row 1: SMU | Narrative label row
  row1 = data.frame(
    Resolution = "SMU",
    Name = "Narrative",
    `Avg Run/Avg Spawners` = "",
    `LRP/LBB` = "",
    `Mgmt Target` = "",
    Forecast = "",
    Outlook = "",
    check.names = FALSE
  )

  # Row 2: SMU value | Narrative value
  row2 = data.frame(
    Resolution = smu,
    Name = narr,
    `Avg Run/Avg Spawners` = "",
    `LRP/LBB` = "",
    `Mgmt Target` = "",
    Forecast = "",
    Outlook = "",
    check.names = FALSE
  )

  # Row 3: Proper table header row (as part of body)
  row3 = data.frame(
    Resolution = "Resolution",
    Name = "Name",
    `Avg Run/Avg Spawners` = "Avg Run/Avg Spawners",
    `LRP/LBB` = "LRP/LBB",
    `Mgmt Target` = "Mgmt Target",
    Forecast = "Forecast",
    Outlook = "Outlook",
    check.names = FALSE
  )

  # Actual data rows
  data_rows = df_smu %>%
    select(Resolution, Name, `Avg Run/Avg Spawners`,
           `LRP/LBB`, `Mgmt Target`, Forecast, Outlook)

  # Combine
  block = bind_rows(row1, row2, row3, data_rows)
  block
}

# Build all SMU blocks
block_list = lapply(smu_list, build_block)

# Combine all blocks
big_df = bind_rows(block_list)

# ---- Convert to flextable ----
ft = flextable(big_df)

# Notice though how that created an extra column row at the top.
# Need to remove this default header row
ft = delete_part(ft, part = "header")

## DO MORE STYLISTIC CHANGES

# Make columns with SMU & Narrative grey and bold
# Style Row 1 (SMU/Narrative labels)
ft = bg(ft, i = which(big_df$Resolution == "SMU" & big_df$Name == "Narrative"),
        bg = "gray90")
ft = bold(ft, i = which(big_df$Resolution == "SMU" & big_df$Name == "Narrative"),
          bold = TRUE)

# Style Row 3 (column headers)
# i.e., Make columns with labels Resolution --> Outlook bold (just need to find the ones with Resolution)
ft = bg(ft, i = which(big_df$Resolution == "Resolution"),
        bg = "gray90")
ft = bold(ft, i = which(big_df$Resolution == "Resolution"),
          bold = TRUE)

# Apply borders + autofit
ft = border_remove(ft)
ft = border_outer(ft, border = fp_border(color = "black", width = 1))
ft = border_inner_h(ft, border = fp_border(color = "black", width = 1))
ft = border_inner_v(ft, border = fp_border(color = "black", width = 1))
ft = autofit(ft)
ft

## Now do some final merging. The Narrative Column should extend  from row 2-->7
# Find all rows where Resolution == "SMU" and Name == "Narrative"
narrative_rows = which(big_df$Resolution == "SMU" & big_df$Name == "Narrative")

# For each of those rows, also merge the row immediately below it
for (i in narrative_rows) {
  # Merge the label row
  ft = merge_at(ft, i = i, j = 2:7)
  # Merge the narrative content row (i + 1)
  ft = merge_at(ft, i = i + 1, j = 2:7)
}

ft

ft = ft %>%
  set_table_properties(layout = "autofit")

## Now add thicker border between each SMU
# We've already identified where these rows are above
# Skip the first one because that's at the top
border_rows = narrative_rows[-1]

# Apply thicker top border to each of those rows (i.e., above where it says SMU and Narrative)
for (i in border_rows) {
  ft = border(ft, i = i, j = 1:ncol(big_df), border.top = fp_border(color = "black", width = 2), part = "body")
}

ft

ftSMU = ft


### REMINDER THIS IS NOT ACTUALLY CORRECT

# area = "SOUTH COAST"
# species = "Chum"
# year = 2026
#
# caption_text <- paste0(
#   "Summary of status metrics informing the annual status evaluation for ",
#   species, " in the ", area, " area during the ", year, " management cycle. ",
#   "Values are presented for each stock management unit (SMU), and where applicable, for associated singular conservation units (CUs), CU aggregations, and Hatchery or Indicator stocks. ",
#   "Reported values include recent average run size or spawner abundance, biological reference points (limit reference point [LRP] and lower biological benchmark [LBB]), and management (mgmt) targets or operational control points used to interpret current conditions. ",
#   "The forecast provides a numerical abundance estimate for the return year, while the outlook indicates the categorical status classification (see definitions in Table 1)."
# )
#
# ftSMU <- set_caption(ftSMU, caption = caption_text)

save(ftSMU, file = "ftSMU.RData")


################################################################################
# Code for making 1 single flextable per SMU

# # Nest the data by smu_area and smu_species
# nested_tabs = tabPrep %>%
#   group_by(smu_area, smu_species, smu_name, Narrative) %>%
#   group_split()
#
#
# # Create flextables
# ft_list = lapply(nested_tabs, function(df_smu) {
#   smu = unique(df_smu$smu_name)
#   narrative = unique(df_smu$Narrative)
#   area = unique(df_smu$smu_area)
#   species = unique(df_smu$smu_species)
#
#   df_smu_clean = df_smu %>% select(-smu_area, -smu_species, -smu_name, -Narrative)
#   n_cols = ncol(df_smu_clean)
#
#   ft = flextable(df_smu_clean) %>%
#     add_header_row(values = c(smu, narrative), colwidths = c(1, n_cols - 1)) %>%
#     add_header_row(values = c("SMU", "Narrative"), colwidths = c(1, n_cols - 1)) %>%
#     bg(i = 1, bg = "grey90", part = "header") %>%
#     bold(i = 1, bold = TRUE, part = "header") %>%
#     color(i = 1, color = "black", part = "header") %>%
#     bg(i = 2, bg = "white", part = "header") %>%
#     bold(i = 2, bold = FALSE, part = "header") %>%
#     bg(i = 3, bg = "grey90", part = "header") %>%
#     bold(i = 3, bold = TRUE, part = "header") %>%
#     border_remove() %>%
#     border_outer(border = fp_border(color = "black", width = 1), part = "all") %>%
#     border_inner_h(border = fp_border(color = "black", width = 1), part = "all") %>%
#     border_inner_v(border = fp_border(color = "black", width = 1), part = "all") %>%
#     autofit() %>%
#     set_table_properties(layout = "autofit")
#
#   list(area = area, species = species, table = ft)
# })
#
#
#
# # Find the first match and extract its table
# areas = c("FRASER AND INTERIOR", "NORTH COAST", "SOUTH COAST", "YUKON")
#
# species_list = c("Chum", "Chinook", "Coho", "Pink-Even", "Sockeye (Lake Type)", "Sockeye (River Type)")
#
#
# ################################################################################
# ################################################################################
# ################################################################################
#
# ### PRACTICE MAKING TABLES WITH FAKE DATA
# # Should be one flextable for all SMUs within ONE area and species
#
# # ---- Fake data with same column names ----
# tabPrep = data.frame(
#   smu_name = c("North Coast – Chum", "North Coast – Chum",
#                "Central Coast – Chum", "Central Coast – Chum"),
#   Narrative = c("Another narrative for this SMU", "Another narrative for this SMU",
#                 "Narrative for Central Coast", "Narrative for Central Coast"),
#   Resolution = c("SMU", "CU (...)", "SMU", "CU (...)"),
#   Name = c("Skeena", "Nass", "Bella Coola", "Ahta"),
#   `Avg Run/Avg Spawners` = c("48,000", "55,000", "22,000", "10,000"),
#   `LRP/LBB` = c("n/a", "n/a", "n/a", "n/a"),
#   `Mgmt Target` = c("9,000", "11,000", "5,000", "2,000"),
#   Forecast = c("Good", "Excellent", "Fair", "Poor"),
#   Outlook = c("Neutral", "Optimistic", "Cautious", "Pessimistic"),
#   stringsAsFactors = FALSE,
#   check.names = FALSE
# )
#
# ## RESTRUCTURE THE DATA FRAME
# # To get the format we want, the data needs to be rearranged
# # Where SMU and Narrative are placed above the remaining data columns
#
# # ---- Split into SMU blocks ----
# smu_list = split(tabPrep, tabPrep$smu_name)
#
# # ---- Build each SMU block as a data frame with fake header rows ----
# build_block = function(df_smu) {
#
#   smu = unique(df_smu$smu_name)
#   narr = unique(df_smu$Narrative)
#
#   # Row 1: SMU | Narrative label row
#   row1 = data.frame(
#     Resolution = "SMU",
#     Name = "Narrative",
#     `Avg Run/Avg Spawners` = "",
#     `LRP/LBB` = "",
#     `Mgmt Target` = "",
#     Forecast = "",
#     Outlook = "",
#     check.names = FALSE
#   )
#
#   # Row 2: SMU value | Narrative value
#   row2 = data.frame(
#     Resolution = smu,
#     Name = narr,
#     `Avg Run/Avg Spawners` = "",
#     `LRP/LBB` = "",
#     `Mgmt Target` = "",
#     Forecast = "",
#     Outlook = "",
#     check.names = FALSE
#   )
#
#   # Row 3: Proper table header row (as part of body)
#   row3 = data.frame(
#     Resolution = "Resolution",
#     Name = "Name",
#     `Avg Run/Avg Spawners` = "Avg Run/Avg Spawners",
#     `LRP/LBB` = "LRP/LBB",
#     `Mgmt Target` = "Mgmt Target",
#     Forecast = "Forecast",
#     Outlook = "Outlook",
#     check.names = FALSE
#   )
#
#   # Actual data rows
#   data_rows = df_smu %>%
#     select(Resolution, Name, `Avg Run/Avg Spawners`,
#            `LRP/LBB`, `Mgmt Target`, Forecast, Outlook)
#
#   # Combine
#   block = bind_rows(row1, row2, row3, data_rows)
#   block
# }
#
# # Build all SMU blocks
# block_list = lapply(smu_list, build_block)
#
# # Combine all blocks
# big_df = bind_rows(block_list)
#
# # ---- Convert to flextable ----
# ft = flextable(big_df)
#
# # Notice though how that created an extra column row at the top.
# # Need to remove this default header row
# ft = delete_part(ft, part = "header")
#
# ## DO MORE STYLISTIC CHANGES
#
# # Make columns with SMU & Narrative grey and bold
# # Style Row 1 (SMU/Narrative labels)
# ft = bg(ft, i = which(big_df$Resolution == "SMU" & big_df$Name == "Narrative"),
#          bg = "gray90")
# ft = bold(ft, i = which(big_df$Resolution == "SMU" & big_df$Name == "Narrative"),
#            bold = TRUE)
#
# # Style Row 3 (column headers)
# # i.e., Make columns with labels Resolution --> Outlook bold (just need to find the ones with Resolution)
# ft = bg(ft, i = which(big_df$Resolution == "Resolution"),
#          bg = "gray90")
# ft = bold(ft, i = which(big_df$Resolution == "Resolution"),
#            bold = TRUE)
#
# # Apply borders + autofit
# ft = border_remove(ft)
# ft = border_outer(ft, border = fp_border(color = "black", width = 1))
# ft = border_inner_h(ft, border = fp_border(color = "black", width = 1))
# ft = border_inner_v(ft, border = fp_border(color = "black", width = 1))
# ft = autofit(ft)
# ft
#
# ## Now do some final merging. The Narrative Column should extend  from row 2-->7
# # Find all rows where Resolution == "SMU" and Name == "Narrative"
# narrative_rows = which(big_df$Resolution == "SMU" & big_df$Name == "Narrative")
#
# # For each of those rows, also merge the row immediately below it
# for (i in narrative_rows) {
#   # Merge the label row
#   ft = merge_at(ft, i = i, j = 2:7)
#   # Merge the narrative content row (i + 1)
#   ft = merge_at(ft, i = i + 1, j = 2:7)
# }
#
# ft
#
# ## Now add thicker border between each SMU
# # We've already identified where these rows are above
# # Skip the first one because that's at the top
# border_rows = narrative_rows[-1]
#
# # Apply thicker top border to each of those rows (i.e., above where it says SMU and Narrative)
# for (i in border_rows) {
#   ft = border(ft, i = i, j = 1:ncol(big_df), border.top = fp_border(color = "black", width = 2), part = "body")
# }
#
# ft
#
#
# ###

