### Code for practicing creating flextables

install.packages("flextable")
library("flextable")

install.packages("readxl")
library("readxl")

install.packages("dplyr")
library("dplyr")

ft <- flextable(airquality[ sample.int(10),])
ft <- add_header_row(ft,
                     colwidths = c(4, 2),
                     values = c("Air quality", "Time")
)
ft <- theme_vanilla(ft)
ft <- add_footer_lines(ft, "Daily air quality measurements in New York, May to September 1973.")
ft <- color(ft, part = "footer", color = "#666666")
ft <- set_caption(ft, caption = "New York Air Quality Measurements")


myft <- flextable(head(mtcars),
                  col_keys = c("am", "carb", "gear", "mpg", "drat" ))
myft = set_caption(myft, caption = "testing if this works")


myft <- italic(myft, j = 3)
myft <- color(myft, ~ drat > 3.5, ~ drat, color = "red")
myft <- bold(myft, ~ drat > 3.5, ~ drat, bold = TRUE)
myft

myft <- add_header_row(
  x = myft, values = c("some measures", "other measures"),
  colwidths = c(3, 2))
myft <- align(myft, i = 1, part = "header", align = "center")

myft

myft = add_header_row(
  x = myft, values = c("stephen's data", "andrea's data"),
  colwidths = c(1,4)
)
myft



dummy_df <- data.frame(
  col1 = letters,
  col2 = letters, stringsAsFactors = FALSE
)
ft_merge <- flextable(dummy_df)
ft_merge <- merge_h(x = ft_merge)
ft_merge










df = read_xlsx("data/outlookClasses.xlsx", col_names = F)


# Load the cleaned CSV
#df <- read_csv("path/to/outlook_table_multilevel.csv")

# Create flextable
ft <- flextable(df)

# Optional: merge vertically repeated values
ft <- ft %>%
  merge_v(j = colnames(df)[1:3]) %>%
  merge_h() %>%
  theme_booktabs() %>%
  autofit()

ft


library(readxl)
library(flextable)
library(dplyr)
library(officer)

# Read the Excel file without headers
df <- read_xlsx("data/outlookClasses.xlsx", col_names = FALSE)

# Create flextable
ft <- flextable(df)

# Merge vertically repeated values in first three columns
ft <- ft %>%
  merge_v(j = 1:3) %>%
  merge_h() %>%
  theme_booktabs() %>%
  autofit()


# Define a border style
my_border <- border(width = 2)


# Add horizontal lines
ft <- ft %>%
  hline_top(border = my_border(width = 2)) %>%  # Top border
  hline(i = which(df[[1]] == "1"), border = my_border(width = 2)) %>%  # Line above row where Outlook Category == 1
  hline_bottom(border = my_border(width = 2))  # Bottom border

ft



