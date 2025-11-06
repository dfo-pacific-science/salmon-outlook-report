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


### NEW CODE FOR ADDING CU NAMES BETTER


fullList = read.csv("data/phase1culookup.csv")


# Add fake data to columns for final tables
# These values will come from the second survey
tabPrep = cu_outlook_records_enriched %>%
  mutate(
    `Avg Run/Avg Spawners` = "50,000",
    `LRP/LBB`        = "n/a",
    `Mgmt Target`    = "10,000"
  )


#### STEPHEN REMOVE THIS BUT JUST FOR TESTING
# tabPrep = tabPrep%>%
#   mutate(
#     smu_area = "SOUTH COAST",
#     smu_species = "Chum"
#   )




# Function to map CU codes to labels
get_labels <- function(cu_string, ref_df) {
  cu_codes <- str_split(cu_string, ",")[[1]]
  labels <- ref_df %>%
    filter(cu %in% cu_codes) %>%
    arrange(match(cu, cu_codes)) %>%  # preserve original order
    pull(label)
  paste(labels, collapse = ", ")
}

# Apply function to tabData
tabPrep = tabPrep %>%
  mutate(CU_Names = map_chr(cu_outlook_selection, ~ get_labels(.x, fullList)))



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
# 1. Function: build_block()
#    Builds the SMU → Narrative → Header → Data block for one SMU
################################################################################

build_block <- function(df_smu) {

  # Unique values for the SMU
  smu  <- unique(df_smu$smu_name)
  narr <- unique(df_smu$Narrative)

  # Row 1: label row (gets grey + bold styling later)
  row1 <- data.frame(
    Resolution = "SMU",
    Name       = "Narrative",
    `Avg Run/Avg Spawners` = "",
    `LRP/LBB` = "",
    `Mgmt Target` = "",
    Forecast   = "",
    Outlook    = "",
    check.names = FALSE
  )

  # Row 2: actual SMU values + narrative content
  row2 <- data.frame(
    Resolution = smu,
    Name       = narr,
    `Avg Run/Avg Spawners` = "",
    `LRP/LBB` = "",
    `Mgmt Target` = "",
    Forecast   = "",
    Outlook    = "",
    check.names = FALSE
  )

  # Row 3: pseudo-header row (becomes part of body)
  row3 <- data.frame(
    Resolution = "Resolution",
    Name       = "Name",
    `Avg Run/Avg Spawners` = "Avg Run/Avg Spawners",
    `LRP/LBB` = "LRP/LBB",
    `Mgmt Target` = "Mgmt Target",
    Forecast   = "Forecast",
    Outlook    = "Outlook",
    check.names = FALSE
  )

  # Actual rows from input
  data_rows <- df_smu %>%
    select(Resolution, Name, `Avg Run/Avg Spawners`,
           `LRP/LBB`, `Mgmt Target`, Forecast, Outlook)

  # Combine into final block
  bind_rows(row1, row2, row3, data_rows)
}

################################################################################
# 2. Function: style_smu_table()
#    Applies all styling, merging, borders, bolding, etc.
################################################################################

style_smu_table <- function(big_df) {

  ft <- flextable(big_df)

  # Remove automatically generated header — we already encoded headers manually
  ft <- delete_part(ft, part = "header")

  # Identify key rows for styling
  idx_label  <- which(big_df$Resolution == "SMU"       & big_df$Name == "Narrative")  # row 1 of each block
  idx_header <- which(big_df$Resolution == "Resolution")                               # row 3 of each block

  # Style label row (row 1)
  ft <- bg(ft, i = idx_label, bg = "gray90")
  ft <- bold(ft, i = idx_label, bold = TRUE)

  # Style pseudo-header row (row 3)
  ft <- bg(ft, i = idx_header, bg = "gray90")
  ft <- bold(ft, i = idx_header, bold = TRUE)

  # Borders
  ft <- border_remove(ft)
  ft <- border_outer(ft, border = fp_border(color="black", width=1))
  ft <- border_inner_h(ft, border = fp_border(color="black", width=1))
  ft <- border_inner_v(ft, border = fp_border(color="black", width=1))

  # Merge narrative across columns (rows 1 + 2)
  for (i in idx_label) {
    ft <- merge_at(ft, i = i,     j = 2:7)
    ft <- merge_at(ft, i = i + 1, j = 2:7)
  }

  # Thicker border between SMU blocks (skip first block)
  if (length(idx_label) > 1) {
    idx_separators <- idx_label[-1]
    for (i in idx_separators) {
      ft <- border(
        ft,
        i = i,
        j = 1:ncol(big_df),
        border.top = fp_border(color="black", width=2),
        part = "body"
      )
    }
  }

  ft <- autofit(ft)
  ft <- set_table_properties(ft, layout = "autofit")

  ft
}

################################################################################
# 3. Function: make_table(area, species, data)
#    This is the function you call in Results.Rmd
################################################################################

make_table <- function(area, species, data = tabPrep) {

  # Filter to SMUs that match both area and species
  df_filtered <- data %>%
    filter(smu_area == area,
           smu_species == species)

  # Handle cases where no data exists
  if (nrow(df_filtered) == 0) {
    stop(paste("No data found for:", area, "/", species))
  }

  # Split into SMU-specific dataframes
  smu_list <- split(df_filtered, df_filtered$smu_name)

  # Build each SMU block
  block_list <- lapply(smu_list, build_block)

  # Combine into one long table for this area × species
  big_df <- bind_rows(block_list)

  # Apply styling
  ft <- style_smu_table(big_df)

  return(ft)
}

################################################################################
# End of master script
################################################################################

make_table("FRASER AND INTERIOR", "Chinook")
make_table("SOUTH COAST", "Sockeye (Lake Type)")


















