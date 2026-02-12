################################################################################
### PRELIMINARY OUTLOOK PRESENTATION

# This code puts together slides for the Preliminary Outlook Presentation
# that is presented in mid-January

# This code builds on a PowerPoint template, where the intro slides are already
# created, and the overall formatting is already set

# It creates tables from the Outlook data survey results, and adds tables for
# each species and area as new slides
# e.g., 1 table (slide) for Yukon Transboundary Sockeye, another slide for
# Yukon Transboundary Chinook, etc.
# Tables include the following information:
# DFO area, Species, SMU name, Resolution, CU Name, Outlook, Forecast

# Fraser and Interior Chinook/Sockeye need to be split out by SMU

# Some manual re-formatting can be done after

### PowerPoint notes from Dawn:
# North Coast Coho: Order should be Nass, Skeena, Haida, Central
# South Coast: order should be wcvi group, ecvi group for all south coast slides
# Okanagan Sockeye and Chinook on same table.
# 1) OKANAGAN SOCKEYE SALMON
# 2) NO DESIGNATED SMU (CK-02) AND
# 3) OKANAGAN CHINOOK SALMON (OKANAGAN_.X)

# Remember to also add in the area names to the title

# Yukon Transboundary Chinook order: alsek stikine taku at top, then yukon and porcupine for this area


# Something about many lines in the table - wont fit
# Fraser Map - added as placeholder
# Maps are created by Chelsea or other GIS tech


################################################################################
## READ IN THE DATA
# Read in the source data from the code that generates the report
# This code uses the data frame tabPrep
source("statusTableWithOthers.R")

################################################################################
## MAKE ADJUSTMENTS TO THE DATAFRAME
# Select only relevant columns that will be displayed on each slide
# Remember that Name is either the name of the SMU or CU
tp_min = tabPrep %>%
  dplyr::select(smu_area, smu_species, smu_name, Resolution, Name, Outlook, Forecast)

# Make some adjustments to the data frame
tp_long = tp_min %>%
  dplyr::mutate(
    # Rename smu_name to just "SMU"
    SMU      = smu_name,
    # Create a new CU column. e
    # If data is provided for only an entire SMU (i.e., Resolution == SMU), fill in
    # a dash for the CU column
    # Otherwise, copy over the CU name
    CU       = dplyr::if_else(Resolution == "SMU", "-", Name),
    # If forecast info is blank, replace it with a space (instead of NA)
    Forecast = dplyr::coalesce(Forecast, "")
  ) %>%
  dplyr::select(smu_area, smu_species, SMU, Resolution, CU, Outlook, Forecast)

################################################################################
### SET THE ORDER OF AREAS/SPECIES

# Dawn likes to present info in the following order (generally North -->)
area_order = c("YUKON TRANSBOUNDARY", "NORTH COAST", "SOUTH COAST", "FRASER AND INTERIOR")

# This is the order for reporting species info
### COME BACK TO THIS... I THINK LAKE TYPE VS RIVER & PINK / EVEN GOT ADJUSTED EARLIER
species_order = c(
  "Sockeye Lake Type",
  "Sockeye River Type",
  "Sockeye",             # FRASER AND INTERIOR only
  "Pink Even",           # not for NORTH COAST
  "Pink",                # NORTH COAST
  "Chinook",
  "Coho",
  "Chum"
)

# Build a named list keyed by "AREA_SPECIES" and sort it
keys = paste(tp_long$smu_area, tp_long$smu_species, sep = "_")
table_list = split(tp_long[, c("SMU", "Resolution", "CU", "Outlook", "Forecast")], keys)


# Reorder table_list by area and species
reorder_table_list = function(tbl) {
  nm_split = do.call(rbind, strsplit(names(tbl), "_"))
  df_order = data.frame(
    name = names(tbl),
    area = nm_split[,1],
    species = nm_split[,2],
    stringsAsFactors = FALSE
  )
  df_order$area_rank = match(df_order$area, area_order)
  df_order$species_rank = match(df_order$species, species_order)
  df_order$species_rank[is.na(df_order$species_rank)] = max(df_order$species_rank, na.rm = TRUE) + 1
  df_order = df_order %>%
    arrange(area_rank, species_rank)
  tbl[df_order$name]
}
table_list = reorder_table_list(table_list)

################################################################################
### ADJUST STYLING

## Table styling
# Make the header dark grey/blue. The border surrounding the table should be the same.
# The first column (SMU) should be dark grey, the rest a light grey
# Centre all the text except what's in the first column
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

## Adjust the font size based on the number of rows so all the text fits in the slide
# If tables have many rows, they might extend beyond slide limits and need to be
# manually fixed
adjust_font = function(ft) {
  n_rows = nrow(ft$body$dataset)
  size = if (n_rows <= 5) 14 else if (n_rows <= 10) 10 else 8
  ft = fontsize(ft, size = size, part = "all")
  ft = style_table(ft)
  return(ft)
}

################################################################################
### READ IN POWERPOINT AND SET DIMENSIONS FOR ADDING THINGS

# Read in the PowerPoint that has the intro slides and basic template
ppt = read_pptx("draftPrelimPres.pptx")

# Placeholder map - will need to be replaced with true map when ready
fraser_map = "data/maps/fraserMap.png"

# These are the names of the slide layouts e.g., if you go View --> Slide Master
# This is the name of the slide with default layout view that gets filled in
layout_name = "1_Title and Content"
# This is the master slide. Needs to be there for reference for some reason, but
# doesn't get adjusted
master_name = "DFO-pp-template"


# Get the slide dimensions
sz = slide_size(ppt)
slide_w = sz$width
slide_h = sz$height

# These are for setting the dimensions of the tables/margins on the sides etc
# Were set through trial and error
# Note that the table won't always match the height of the map... it depends on
# how many rows of data there are
left_margin  = 0.5
right_margin = 0.5
map_width    = 7.95
map_height   = 5.14
table_width  = 4.2
table_height = map_height

map_y   = (slide_h - map_height) / 2
table_y = map_y
table_x = left_margin
map_x = slide_w - map_width - right_margin

# Shift table down for species text box
species_height = 0.4
species_y = map_y
table_y = table_y + species_height


################################################################################
# Loop to add slides

# It loops through each  species/area to generate a new slide
#

# This makes 1 slide for each Area/SMU. However, Fraser and Interior data, the
# Chinook and Sockeye data are so detailed that 1 slide per SMU is created




for (nm in names(table_list)) {
  df_full = table_list[[nm]]
  area_species = strsplit(nm, "_")[[1]]
  area = area_species[1]
  species = area_species[2]

  # Split FRASER AND INTERIOR Chinook/Sockeye by SMU
  if (area == "FRASER AND INTERIOR" & species %in% c("Chinook", "Sockeye")) {
    smus = unique(df_full$SMU)
    for (smu_val in smus) {
      df = df_full %>% filter(SMU == smu_val)
      # create slide
      ft = flextable(df) %>% adjust_font()
      col_count = ncol(df)
      col_width = table_width / col_count
      ft = width(ft, j = 1:col_count, width = col_width)

      ppt = add_slide(ppt, layout = layout_name, master = master_name)
      ppt = ph_with(ppt, value = paste0("OUTLOOKS – ", area), location = ph_location_type(type = "title"))
      ppt = ph_with(ppt, value = ft, location = ph_location(left = table_x, top = table_y, width = table_width, height = table_height))

      species_text = fpar(
        ftext(species, prop = fp_text(color = "black", font.size = 20, font.family = "Segoe UI Semibold", bold = TRUE)),
        fp_p = fp_par(text.align = "left")
      )
      ppt = ph_with(ppt, value = species_text, location = ph_location(left = table_x, top = species_y, width = table_width, height = species_height))

      if (file.exists(fraser_map)) {
        ppt = ph_with(
          ppt,
          value = external_img(fraser_map, width = map_width, height = map_height),
          location = ph_location(left = map_x, top = map_y, width = map_width, height = map_height)
        )
      }

      # Add notes
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

    # Regular slide for all other species/areas
    df = df_full
    ft = flextable(df) %>% adjust_font()
    col_count = ncol(df)
    col_width = table_width / col_count
    ft = width(ft, j = 1:col_count, width = col_width)

    ppt = add_slide(ppt, layout = layout_name, master = master_name)
    ppt = ph_with(ppt, value = paste0("OUTLOOKS – ", area), location = ph_location_type(type = "title"))
    ppt = ph_with(ppt, value = ft, location = ph_location(left = table_x, top = table_y, width = table_width, height = table_height))

    species_text = fpar(
      ftext(species, prop = fp_text(color = "black", font.size = 20, font.family = "Segoe UI Semibold", bold = TRUE)),
      fp_p = fp_par(text.align = "left")
    )
    ppt = ph_with(ppt, value = species_text, location = ph_location(left = table_x, top = species_y, width = table_width, height = species_height))

    if (file.exists(fraser_map)) {
      ppt = ph_with(
        ppt,
        value = external_img(fraser_map, width = map_width, height = map_height),
        location = ph_location(left = map_x, top = map_y, width = map_width, height = map_height)
      )
    }

    # Add notes
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

print(ppt, target = "updated_presentation.pptx")