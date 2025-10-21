### Code for practicing creating flextables


library(readxl)
library(flextable)
library(dplyr)
library(officer)

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


library("purrr")

dummyPath = "data/finnis_outlook_phase1_test_dummy_data_20250905.xlsx"

# Locate the file by its relative path
#file = doclib$get_item("General/Pacific Salmon Data Community of Practice/Projects & Task Teams/Reproducible Data Products/Salmon Outlook Reporting/Background Research/outlook_phase1_test_dummy_data_20250905.xlsx")


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














