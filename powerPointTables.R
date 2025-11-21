
library(officer)
library(flextable)
library(dplyr)
library(magrittr)

# Slide dimensions
slide_w <- 10
slide_h <- 5.63

# Margins
left_margin  <- 0.5
right_margin <- 0.5

# Map size (unchanged)
map_width    <- 7.95
map_height   <- 5.14

# Compute positions
map_x <- slide_w - map_width - right_margin
map_y <- (slide_h - map_height) / 2

# Table size and position (fixed width)
table_width  <- 2.0   # fixed width for 4 columns
table_height <- map_height
table_x <- left_margin
table_y <- map_y

# Styling function
style_table <- function(ft) {
  ft <- bg(ft, part = "header", bg = "#3B4A5A") |>
    color(part = "header", color = "white") |>
    bold(part = "header", bold = TRUE) |>
    bg(j = 1, bg = "#A9B1B7", part = "body") |>
    bg(j = 2:(ncol(ft$body$dataset)), bg = "#E6ECF3", part = "body") |>
    border_remove() |>
    border_outer(part = "all", border = fp_border(color = "#3B4A5A", width = 2)) |>
    border_inner(part = "all", border = fp_border(color = "white", width = 1)) |>
    align(j = 1, align = "left", part = "all") |>
    align(j = 2:(ncol(ft$body$dataset)), align = "center", part = "all") |>
    autofit()
  return(ft)
}

# Font size adjustment
adjust_font <- function(ft) {
  n_rows <- nrow(ft$body$dataset)
  size <- if (n_rows <= 5) 18 else if (n_rows <= 10) 14 else 10
  ft <- fontsize(ft, size = size, part = "all")
  ft <- style_table(ft)
  return(ft)
}

# Example workflow
ppt <- read_pptx("PracticePowerPoint.pptx")
tables <- list(head(mtcars), head(iris))  # raw data frames
images <- c("data/maps/southCoastMap.png", "data/maps/fraserMap.png")

for (i in seq_along(tables)) {
  # Keep only first 4 columns
  df <- tables[[i]][, 1:4]
  ft <- flextable(df)
  ft <- adjust_font(ft)

  ppt <- on_slide(ppt, index = 4 + i)

  # Add table on the left
  ppt <- ph_with(
    ppt,
    value = ft,
    location = ph_location(
      left = table_x,
      top  = table_y,
      width = table_width,
      height = table_height
    )
  )

  # Add map on the right
  ppt <- ph_with(
    ppt,
    value = external_img(images[i], width = map_width, height = map_height),
    location = ph_location(
      left = map_x,
      top  = map_y
    )
  )
}

print(ppt, target = "updated_presentation.pptx")
