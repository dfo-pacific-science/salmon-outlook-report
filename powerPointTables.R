library(officer)
library(flextable)

# Custom flextable styling function
style_table <- function(ft) {
  # Header style
  ft <- bg(ft, part = "header", bg = "#3B4A5A") # dark blue-grey
  ft <- color(ft, part = "header", color = "white")
  ft <- bold(ft, part = "header", bold = TRUE)

  # Column 1 background (except header)
  ft <- bg(ft, j = 1, bg = "#A9B1B7", part = "body") # grey

  # Other cells background
  ft <- bg(ft, j = 2:(ncol(ft$body$dataset)), bg = "#E6ECF3", part = "body") # light blue

  # Borders
  ft <- border_remove(ft)
  ft <- border_outer(ft, part = "all", border = fp_border(color = "#3B4A5A", width = 2)) # thick outer border
  ft <- border_inner(ft, part = "all", border = fp_border(color = "white", width = 1))   # thin inner borders

  # Alignment
  ft <- align(ft, j = 1, align = "left", part = "all")
  ft <- align(ft, j = 2:(ncol(ft$body$dataset)), align = "center", part = "all")

  # Font family
  ft <- set_flextable_defaults(font.family = "Segoe UI Semilight")

  # Autofit columns
  ft <- autofit(ft)

  return(ft)
}

# Adjust font size based on row count
adjust_font <- function(ft) {
  n_rows <- nrow(ft$body$dataset)
  size <- if (n_rows <= 5) 18 else if (n_rows <= 10) 14 else 10
  ft <- fontsize(ft, size = size, part = "all")
  ft <- style_table(ft)
  return(ft)
}

# Example workflow
ppt <- read_pptx("PracticePowerPoint.pptx")
tables <- list(flextable(head(mtcars)), flextable(head(iris)))

for (i in seq_along(tables)) {
  ft <- adjust_font(tables[[i]])
  ppt <- on_slide(ppt, index = 4 + i)
  ppt <- ph_with(
    ppt,
    value = ft,
    location = ph_location(left = 0.5, top = 1.5, width = 3.3, height = 5)
  )
}

print(ppt, target = "updated_presentation.pptx")