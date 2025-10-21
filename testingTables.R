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