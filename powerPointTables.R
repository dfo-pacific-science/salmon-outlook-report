############ STPEHEN I THINK THE CHECK THING STOPPED WORKING??????
# MAKE SURE SMU SHOWS UP AS -
# ALSO REMEMBER TO ADD NOTES






library(officer)
library(flextable)
library(dplyr)
library(magrittr)

## ----------------------------
## Load real table data
## ----------------------------
source("tableTest4.R")   # defines tabPrep and builds table_list

## ----------------------------
## Helpers: styling functions
## ----------------------------
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
    align(j = 2:(ncol(ft$body$dataset)), align = "center", part = "all")
  return(ft)
}

adjust_font <- function(ft) {
  n_rows <- nrow(ft$body$dataset)
  size <- if (n_rows <= 5) 14 else if (n_rows <= 10) 10 else 8
  ft <- fontsize(ft, size = size, part = "all")
  ft <- style_table(ft)
  return(ft)
}

## ----------------------------
## Read pptx & pick a layout
## ----------------------------
ppt <- read_pptx("draftPrelimPres.pptx")

lsum <- layout_summary(ppt)
layout_name <- lsum$layout[1]      # Use first available layout
master_name <- lsum$master[1]

## ----------------------------
## Slide + object dimensions
## ----------------------------
sz <- slide_size(ppt)
slide_w <- sz$width
slide_h <- sz$height

left_margin  <- 0.5
right_margin <- 0.5

map_width    <- 7.95
map_height   <- 5.14
table_width  <- 4.2   # ✅ Exact width you want
table_height <- map_height

map_y   <- (slide_h - map_height) / 2
table_y <- map_y
table_x <- left_margin
map_x   <- slide_w - map_width - right_margin

fraser_map <- "data/maps/fraserMap.png"

## ----------------------------
## Loop to add slides for real tables
## ----------------------------
for (nm in names(table_list)) {
  df <- table_list[[nm]]

  # ✅ Apply transformations for CHECK
  df <- df %>%
    mutate(
      SMU = if_else(grepl("CHECK", SMU),
                    sub("(CHECK).*", "\\1", SMU),
                    SMU),
      Outlook = if_else(grepl("CHECK", Outlook),
                        "CHECK",
                        Outlook)
    )

  # ✅ Create flextable and force width
  ft <- flextable(df) |> adjust_font()
  col_count <- ncol(df)
  col_width <- table_width / col_count
  ft <- width(ft, j = 1:col_count, width = col_width)  # ✅ Force total width

  ppt <- add_slide(ppt, layout = layout_name, master = master_name)

  ppt <- ph_with(
    ppt,
    value = ft,
    location = ph_location(left = table_x, top = table_y, width = table_width)
  )

  if (file.exists(fraser_map)) {
    ppt <- ph_with(
      ppt,
      value = external_img(fraser_map, width = map_width, height = map_height),
      location = ph_location(left = map_x, top = map_y, width = map_width, height = map_height)
    )
  } else {
    warning(paste("Image not found:", fraser_map))
  }
}

## ----------------------------
## Save updated presentation
## ----------------------------
print(ppt, target = "updated_presentation.pptx")
