
# ----------------------------
# Load libraries
# ----------------------------
library(officer)
library(flextable)
library(dplyr)
library(magrittr)


# ----------------------------
# Load data
# ----------------------------
source("tableTest4.R")   # defines tabPrep and builds table_list

# ----------------------------
# Styling helpers
# ----------------------------
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

# ----------------------------
# Read PPTX and set layout
# ----------------------------
ppt <- read_pptx("draftPrelimPres.pptx")
layout_name <- "1_Title and Content"
master_name <- "DFO-pp-template"

# Inspect layout of slides
layout_summary(ppt)
layout_properties(ppt, layout = "1_Title and Content", master = "DFO-pp-template")









# ----------------------------
# Slide dimensions
# ----------------------------
sz <- slide_size(ppt)
slide_w <- sz$width
slide_h <- sz$height

left_margin  <- 0.5
right_margin <- 0.5
map_width    <- 7.95
map_height   <- 5.14
table_width  <- 4.2
table_height <- map_height

map_y   <- (slide_h - map_height) / 2
table_y <- map_y
table_x <- left_margin
map_x   <- slide_w - map_width - right_margin

fraser_map <- "data/maps/fraserMap.png"

# ----------------------------
# Loop to add slides
# ----------------------------
for (nm in names(table_list)) {
  # Split nm into area and species
  area_species <- strsplit(nm, "_")[[1]]
  area <- area_species[1]
  species <- area_species[2]

  # Prepare table data
  df <- table_list[[nm]] %>%
    mutate(
      SMU = if_else(is.na(SMU) | SMU == "", "-", SMU),
      SMU = if_else(grepl("CHECK", SMU), "CHECK", SMU),
      Outlook = if_else(grepl("CHECK", Outlook), "CHECK", Outlook)
    )

  # Create flextable and adjust width
  ft <- flextable(df) |> adjust_font()
  col_count <- ncol(df)
  col_width <- table_width / col_count
  ft <- width(ft, j = 1:col_count, width = col_width)

  # Add new slide
  ppt <- add_slide(ppt, layout = layout_name, master = master_name)

  # ✅ Add title
  ppt <- ph_with(
    ppt,
    value = paste0("OUTLOOKS – ", area),  # en dash between OUTLOOKS and area
    location = ph_location_type(type = "title")
  )

  # ✅ Add table
  ppt <- ph_with(
    ppt,
    value = ft,
    location = ph_location(left = table_x, top = table_y, width = table_width, height = table_height)
  )

  # Add species name
  ppt <- ph_with(
    ppt,
    value = species,
    location = ph_location_type(type = "body", type_idx = 5)  # assumes the text box is the body placeholder
  )



  # ✅ Add map
  if (file.exists(fraser_map)) {
    ppt <- ph_with(
      ppt,
      value = external_img(fraser_map, width = map_width, height = map_height),
      location = ph_location(left = map_x, top = map_y, width = map_width, height = map_height)
    )
  } else {
    warning(paste("Image not found:", fraser_map))
  }


  # ✅ Add notes from tabPrep
  if (all(c("smu_area", "smu_species", "Narrative") %in% names(tabPrep))) {
    narratives_df <- tabPrep %>%
      filter(trimws(smu_area) == trimws(area),
             trimws(smu_species) == trimws(species)) %>%
      distinct(smu_name, Narrative)  # ✅ Remove duplicates based on SMU and Narrative

    if (nrow(narratives_df) > 0) {
      notes_text <- narratives_df %>%
        mutate(note_line = paste0(smu_name, ": ", coalesce(Narrative, ""))) %>%
        pull(note_line) %>%
        unique() %>%  # ✅ Extra safeguard
        paste(collapse = "\n")
    } else {
      notes_text <- "No notes available"
    }


    }
    ppt <- set_notes(ppt, value = notes_text, location = notes_location_type("body"))
  }


# ----------------------------
# Save presentation
# ----------------------------
print(ppt, target = "updated_presentation.pptx")
