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

# Read raw sheet without interpreting the first row as names
raw = read_excel("data/outlookClasses.xlsx", sheet = 1, col_names = FALSE)

# Extract top and bottom header rows (first two rows of the sheet)
header_top    = as.character(unlist(raw[1, ], use.names = FALSE))
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





####################################################################################




dummyPath = "data/finnis_outlook_phase1_test_dummy_data_20250905.xlsx"


# List all sheet names
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
    Outlook_Repeat_Test %>%
      select(uniquerowid, smu_area, smu_species, smu_name),
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


#####

df = cu_outlook_records_enriched

df <- df %>%
  group_by(smu_name) %>%
  mutate(
    Resolution = case_when(

      # Case 1: Only one row for this SMU, and no CU info
      # → standalone SMU (no CU breakdown)
      n() == 1 & cu_count ==1 ~ "SMU",

      # Case 2: Multiple rows for this SMU, and more than one CU per SMU
      # → aggregate view across multiple CUs
      n() > 1 & cu_count > 1 ~ "CU (aggregate)",

      # Case 3: Multiple rows for this SMU, each representing a single CU
      # → individual CU(s) listed separately
      n() > 1 & cu_count == 1 ~ "CU (singular)",

      # Case 4: Single SMU row, but cu_count is unexpectedly defined
      # → flag as a potential data issue for review
      n() == 1 & cu_count >1 ~ paste0(
        "CHECK VALUE — single SMU but cu_count = ", cu_count
      ),

      # Catch-all: anything that doesn’t match expected patterns
      TRUE ~ "CHECK VALUE — unexpected combination"
    )) %>%

  ungroup()%>%


  mutate(

    # Assign a name based on the resolution type
    Name = case_when(

      # If resolution is SMU, use the smu_name
      Resolution == "SMU" ~ smu_name,

      # If resolution is CU (aggregate) or CU (singular), use cu_outlook_selection
      Resolution %in% c("CU (aggregate)", "CU (singular)") ~ cu_outlook_selection,

      # If resolution is any kind of CHECK VALUE, assign a generic label
      str_detect(Resolution, "CHECK VALUE") ~ "Check value",

      # Fallback: assign NA if none of the above match
      TRUE ~ NA_character_
    ))



df2 = df %>%
  mutate(
    `Avg Run/Avg Spawners` = "50,000",
    `LRP/LBB`        = "n/a",
    `Mgmt Target`    = "10,000",
    `Narrative Text` = paste(
      rep("Stephen slays the house down boots. ", 5),
      collapse = ""
    )
  )



df3 = df2 %>%
  select(
         `Resolution`,
         `smu_name`,
         `Avg Run/Avg Spawners`,
         `LRP/LBB`,
         `Mgmt Target`,
         `cu_prelim_forecast`,
         `cu_outlook_assignment`
         )



df_list <- df2 %>% group_split(smu_name)

ft_list <- lapply(df_list, function(df_smu) {
  smu <- unique(df_smu$smu_name)
  narrative <- unique(df_smu$`Narrative Text`)

  df3 <- df_smu %>%
    select(
      `Resolution`,
      `Name`,
      `Avg Run/Avg Spawners`,
      `LRP/LBB`,
      `Mgmt Target`,
      `cu_prelim_forecast`,
      `cu_outlook_assignment`
    )

  # keep column count dynamic
  n_cols <- ncol(df3)

  ft <- flextable(df3) %>%

    # Add header rows (values first, labels second so labels sit above values)
    add_header_row(values = c(smu, narrative), colwidths = c(1, n_cols - 1)) %>%
    add_header_row(values = c("SMU", "Narrative"), colwidths = c(1, n_cols - 1)) %>%

    # Now style header rows AFTER both header rows exist:
    # top header (labels) -> row 1: bold + grey background
    bg(i = 1, bg = "grey90", part = "header") %>%
    bold(i = 1, bold = TRUE, part = "header") %>%
    color(i = 1, color = "black", part = "header") %>%

    # second header (values) -> row 2: plain text, white background
    bg(i = 2, bg = "white", part = "header") %>%
    bold(i = 2, bold = FALSE, part = "header") %>%

    # third header row is the actual column names -> style as desired
    bg(i = 3, bg = "grey90", part = "header") %>%
    bold(i = 3, bold = TRUE, part = "header") %>%

    # Ensure all cells (header + body) have black borders
    border_remove() %>%
    border_outer(border = fp_border(color = "black", width = 1), part = "all") %>%
    border_inner_h(border = fp_border(color = "black", width = 1), part = "all") %>%
    border_inner_v(border = fp_border(color = "black", width = 1), part = "all") %>%

    autofit() %>%
    set_table_properties(layout = "autofit")

  ft
})



# Name each flextable by its SMU
names(ft_list) <- sapply(df_list, function(x) unique(x$smu_name))

# Print all tables
for (ft in ft_list) {
  print(ft)
}
ft_list[3]

