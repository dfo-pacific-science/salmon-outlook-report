################################################################################
### PRELIMINARY OUTLOOK PRESENTATION
###
### PURPOSE
### -------
### This script automates the creation of the Preliminary Outlook Presentation.
### It uses a PowerPoint template and generates slides for each species and area
### using survey results from the `tabPrep` dataset.
###
### SLIDE CONTENT
### -------------
### • Top-right banner: "OUTLOOKS – AREA NAME"
### • Subtitle above table: "Area – Species"
### • Table: SMU | Resolution | CU | Outlook | Forecast
### • Map image (placeholder, if available)
### • Narrative text added to slide notes
###
### SPECIAL CASE
### ------------
### • Fraser and Interior Chinook and Sockeye: one slide per SMU
################################################################################


################################################################################
## READ IN THE DATA
################################################################################

# Load the pre-processed survey data frame `tabPrep`
# This contains the fields:
# smu_area, smu_species, smu_name, Resolution, Name, Outlook, Forecast, Narrative
source("statusTableWithOthers.R")


################################################################################
## PREPARE THE DATA
################################################################################

# Keep only columns relevant for tables
# smu_area & smu_species are kept for grouping only; they are not in the table
tp_min = tabPrep %>%
  dplyr::select(smu_area, smu_species, smu_name, Resolution, Name, Outlook, Forecast)

# Adjust the data frame for table display
tp_long = tp_min %>%
  dplyr::mutate(
    # Rename SMU column
    SMU = smu_name,
    # Create CU column: dash for SMU-level rows, else copy CU name
    CU = dplyr::if_else(Resolution == "SMU", "-", Name),
    # Replace NA in Forecast with blank for table clarity
    Forecast = dplyr::coalesce(Forecast, "")
  ) %>%
  # Keep only the columns that appear in the slide table
  dplyr::select(smu_area, smu_species, SMU, Resolution, CU, Outlook, Forecast)


################################################################################
### SET SLIDE ORDER
################################################################################

# Define the order of areas for presentation (North to South)
area_order = c("YUKON TRANSBOUNDARY", "NORTH COAST", "SOUTH COAST", "FRASER AND INTERIOR")

# Define the order of species for presentation
# Notes:
# • Some species are lake type or river type sockeye
# • Pink Even is not used for North Coast
species_order = c(
  "Sockeye Lake Type",
  "Sockeye River Type",
  "Sockeye",
  "Pink Even",
  "Pink",
  "Chinook",
  "Coho",
  "Chum"
)

# Split table into list by "AREA_SPECIES"
keys = paste(tp_long$smu_area, tp_long$smu_species, sep = "_")
table_list = split(tp_long[, c("SMU", "Resolution", "CU", "Outlook", "Forecast")], keys)

# Reorder tables according to area and species preferences
reorder_table_list = function(tbl) {
  nm_split = do.call(rbind, strsplit(names(tbl), "_"))
  df_order = data.frame(
    name = names(tbl),
    area = nm_split[,1],
    species = nm_split[,2],
    stringsAsFactors = FALSE
  )
  # Rank areas and species
  df_order$area_rank = match(df_order$area, area_order)
  df_order$species_rank = match(df_order$species, species_order)
  # Any species not in species_order get pushed to end
  df_order$species_rank[is.na(df_order$species_rank)] = max(df_order$species_rank, na.rm = TRUE) + 1
  # Reorder table list
  df_order = df_order %>% arrange(area_rank, species_rank)
  tbl[df_order$name]
}
table_list = reorder_table_list(table_list)


################################################################################
### TABLE STYLING FUNCTIONS
################################################################################

# Style table headers, borders, background colors, and alignment
style_table = function(ft) {
  ft = bg(ft, part = "header", bg = "#3B4A5A") %>%
    color(part = "header", color = "white") %>%
    bold(part = "header", bold = TRUE) %>%
    bg(j = 1, bg = "#A9B1B7", part = "body") %>%
    bg(j = 2:(ncol(ft$body$dataset)), bg = "#E6ECF3", part = "body") %>%
    border_remove() %>%
    border_outer(part = "all", border = fp_border(color = "#3B4A5A", width = 2)) %>%
    border_inner(part = "all", border = fp_border(color = "white", width = 1)) %>%
    align(j = 1, align = "left", part = "all") %>%
    align(j = 1:(ncol(ft$body$dataset)), align = "center", part = "all")
  return(ft)
}

# Adjust font size depending on number of rows
adjust_font = function(ft) {
  n_rows = nrow(ft$body$dataset)
  size = if (n_rows <= 5) 14 else if (n_rows <= 10) 10 else 8
  ft = fontsize(ft, size = size, part = "all")
  ft = style_table(ft)
  return(ft)
}


################################################################################
### READ IN POWERPOINT TEMPLATE AND SET DIMENSIONS
################################################################################

# Load the PPTX template with intro slides already in place
ppt = read_pptx("draftPrelimPres.pptx")

# Map placeholder (will be replaced with final GIS map)
fraser_map = "data/maps/fraserMap.png"

# Slide layout names
layout_name = "1_Title and Content"  # Layout for new slides
master_name = "DFO-pp-template"      # Reference master slide

# Slide dimensions
sz = slide_size(ppt)
slide_w = sz$width
slide_h = sz$height

# Layout positioning parameters (trial-and-error)
left_margin  = 0.5
right_margin = 0.5
map_width    = 7.95
map_height   = 5.14
table_width  = 4.2
table_height = map_height

# Vertical positions
map_y   = (slide_h - map_height) / 2
table_y = map_y
table_x = left_margin
map_x   = slide_w - map_width - right_margin

# Shift table down for species text box
species_height = 0.4
species_y = map_y
table_y = table_y + species_height


################################################################################
### LOOP THROUGH TABLES TO ADD SLIDES
################################################################################

for (nm in names(table_list)) {
  df_full = table_list[[nm]]
  area_species = strsplit(nm, "_")[[1]]
  area = area_species[1]
  species = area_species[2]

  # SPECIAL CASE: Fraser/Interior Chinook or Sockeye → one slide per SMU
  if (area == "FRASER AND INTERIOR" & species %in% c("Chinook", "Sockeye")) {
    smus = unique(df_full$SMU)
    for (smu_val in smus) {
      df = df_full %>% filter(SMU == smu_val)
      # Create table for slide
      ft = flextable(df) %>% adjust_font()
      col_count = ncol(df)
      col_width = table_width / col_count
      ft = width(ft, j = 1:col_count, width = col_width)

      # Add new slide
      ppt = add_slide(ppt, layout = layout_name, master = master_name)

      # Add banner at top of slide: "OUTLOOKS – AREA NAME"
      ppt = ph_with(ppt, value = paste0("OUTLOOKS – ", area),
                    location = ph_location_type(type = "title"))

      # Add table
      ppt = ph_with(ppt, value = ft,
                    location = ph_location(left = table_x, top = table_y,
                                           width = table_width, height = table_height))

      # Add subtitle above table: "Area – Species"
      # Note: This was previously only species; now includes area first
      # Convert area to title case for subtitle bove table
      area_title = tools::toTitleCase(tolower(area))
      subtitle_text = paste0(area_title, " – ", species)
      species_text = fpar(
        ftext(subtitle_text,
              prop = fp_text(color = "black", font.size = 20,
                             font.family = "Segoe UI Semibold", bold = TRUE)),
        fp_p = fp_par(text.align = "left")
      )
      ppt = ph_with(ppt, value = species_text,
                    location = ph_location(left = table_x, top = species_y,
                                           width = table_width, height = species_height))

      # Add map if exists
      if (file.exists(fraser_map)) {
        ppt = ph_with(ppt,
                      value = external_img(fraser_map, width = map_width, height = map_height),
                      location = ph_location(left = map_x, top = map_y,
                                             width = map_width, height = map_height))
      }

      # Add slide notes (narrative text)
      if (all(c("smu_area", "smu_species", "Narrative") %in% names(tabPrep))) {
        narratives_df = tabPrep %>%
          filter(trimws(smu_area) == trimws(area),
                 trimws(smu_species) == trimws(species),
                 trimws(smu_name) == trimws(smu_val)) %>%
          distinct(smu_name, Narrative)
        notes_text = if (nrow(narratives_df) > 0) {
          narratives_df %>%
            mutate(note_line = paste0(smu_name, ": ", coalesce(Narrative, ""))) %>%
            pull(note_line) %>% unique() %>% paste(collapse = "\n")
        } else {
          "No notes available"
        }
        ppt = set_notes(ppt, value = notes_text, location = notes_location_type("body"))
      }
    }

  } else {

    # REGULAR SLIDE: All other species/areas (not Fraser/Interior Chinook/Sockeye)
    # ----------------------------
    # df_full contains all SMUs/CUs for this area/species
    df = df_full

    # Create table for slide and adjust font
    ft = flextable(df) %>% adjust_font()
    col_count = ncol(df)
    col_width = table_width / col_count
    ft = width(ft, j = 1:col_count, width = col_width)

    # Add new slide
    ppt = add_slide(ppt, layout = layout_name, master = master_name)

    # Add banner at top: "OUTLOOKS – AREA NAME"
    ppt = ph_with(ppt, value = paste0("OUTLOOKS – ", area),
                  location = ph_location_type(type = "title"))

    # Add table
    ppt = ph_with(ppt, value = ft,
                  location = ph_location(left = table_x, top = table_y,
                                         width = table_width, height = table_height))

    # Add subtitle above table: "Area – Species"
    # This replaces the old behavior where only species was shown
    area_title = tools::toTitleCase(tolower(area))
    subtitle_text = paste0(area_title, " – ", species)
    species_text = fpar(
      ftext(subtitle_text,
            prop = fp_text(color = "black", font.size = 20,
                           font.family = "Segoe UI Semibold", bold = TRUE)),
      fp_p = fp_par(text.align = "left")
    )
    ppt = ph_with(ppt, value = species_text,
                  location = ph_location(left = table_x, top = species_y,
                                         width = table_width, height = species_height))

    # Add map if available
    if (file.exists(fraser_map)) {
      ppt = ph_with(ppt,
                    value = external_img(fraser_map, width = map_width, height = map_height),
                    location = ph_location(left = map_x, top = map_y,
                                           width = map_width, height = map_height))
    }

    # Add slide notes (narrative text)
    if (all(c("smu_area", "smu_species", "Narrative") %in% names(tabPrep))) {
      narratives_df = tabPrep %>%
        filter(trimws(smu_area) == trimws(area),
               trimws(smu_species) == trimws(species)) %>%
        distinct(smu_name, Narrative)
      notes_text = if (nrow(narratives_df) > 0) {
        narratives_df %>%
          mutate(note_line = paste0(smu_name, ": ", coalesce(Narrative, ""))) %>%
          pull(note_line) %>% unique() %>% paste(collapse = "\n")
      } else {
        "No notes available"
      }
      ppt = set_notes(ppt, value = notes_text, location = notes_location_type("body"))
    }
  }
}


################################################################################
### SAVE PRESENTATION
################################################################################

print(ppt, target = "updated_presentation.pptx")