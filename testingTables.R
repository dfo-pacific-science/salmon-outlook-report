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
    )
  ) %>%
  ungroup()


df2 = df %>%
  mutate(
    `Avg Run/Avg Spawners` = c("50,000"),
    `LRP/LBB`        = c("n/a"),
    `Mgmt Target`    = c("10,000"),
    `Narrative Text` = paste(
      rep("Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ", 5),
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



library(dplyr)
library(flextable)


###########################


library(dplyr)
library(flextable)

# Split df2 by smu_name
df_list <- df2 %>%
  group_split(smu_name)

# Create flextables with correct header block and full black borders
ft_list <- lapply(df_list, function(df_smu) {
  # Extract values
  smu <- unique(df_smu$smu_name)
  narrative <- unique(df_smu$`Narrative Text`)

  # Create the main table
  df3 <- df_smu %>%
    select(
      `Resolution`,
      `smu_name`,
      `Avg Run/Avg Spawners`,
      `LRP/LBB`,
      `Mgmt Target`,
      `cu_prelim_forecast`,
      `cu_outlook_assignment`
    )

  # Create flextable
  ft <- flextable(df3) %>%

    # Add top block: SMU and Narrative values (plain row FIRST)
    add_header_row(
      values = c(smu, narrative),
      colwidths = c(1, 6)
    ) %>%

    bold(i = 2, bold = TRUE, part = "header") %>%
    bg(i = 2, bg = "grey90", part = "header") %>%



    # Add top block: SMU and Narrative labels (styled row SECOND)
    add_header_row(
      values = c("SMU", "Narrative"),
      colwidths = c(1, 6)
    ) %>%

    bold(i = 1, bold = FALSE, part = "header") %>%
    bg(i = 1, bg = "white", part = "header") %>%

    bold(i = 2, bold = TRUE, part = "header") %>%
    bg(i = 2, bg = "grey90", part = "header") %>%
    color(i = 2, color = "black", part = "header") %>%

    # Style the actual table header (starts at row 3)
    bold(i = 3, bold = TRUE, part = "header") %>%
    bg(i = 3, bg = "grey90", part = "header") %>%
    color(i = 3, color = "black", part = "header") %>%

    # Apply full black borders to all cells
    border_remove() %>%
    border_outer(border = officer::fp_border(color = "black", width = 1)) %>%
    border_inner_h(border = officer::fp_border(color = "black", width = 1)) %>%
    border_inner_v(border = officer::fp_border(color = "black", width = 1)) %>%

    # Final layout
    autofit() %>%
    set_table_properties(layout = "autofit", width = 1)
})

# Name each flextable by its SMU
names(ft_list) <- sapply(df_list, function(x) unique(x$smu_name))

# Print all tables
for (ft in ft_list) {
  print(ft)
}
