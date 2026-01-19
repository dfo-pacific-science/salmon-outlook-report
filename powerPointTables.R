# ----------------------------
# Load data
# ----------------------------

### PowerPoint notes from Dawn:
# North Coast Coho: Order should be Nass, Skeena, Haida, Central
# South Coast: order should be wcvi group, ecvi group for all south coast slides
# Okanagan Sockeye and Chinook on same table.
# 1) OKANAGAN SOCKEYE SALMON
# 2) NO DESIGNATED SMU (CK-02) AND
# 3) OKANAGAN CHINOOK SALMON (OKANAGAN_.X)


source("statusTableWithOthers.R")   # defines tabPrep and builds table_list


# ----------------------------
# Build table_list: SMU, Resolution, CU, Outlook, Forecast
# ----------------------------
need <- c("smu_area", "smu_species", "smu_name", "Resolution", "Name", "Outlook", "Forecast")
if (!exists("tabPrep")) stop("tabPrep not found after sourcing statusTables.R")
if (!all(need %in% names(tabPrep))) {
  stop(paste("tabPrep is missing columns:",
             paste(setdiff(need, names(tabPrep)), collapse = ", ")))
}

tp_min <- tabPrep %>%
  dplyr::select(smu_area, smu_species, smu_name, Resolution, Name, Outlook, Forecast)

# Keep SMU rows and CU-equivalent rows: CU, PFMA, Hatchery or Indicator Stock
tp_long <- tp_min %>%
  dplyr::filter(
    Resolution == "SMU" |
      grepl("^CU", Resolution) |
      Resolution %in% c("PFMA", "Hatchery or Indicator Stock")
  ) %>%
  dplyr::mutate(
    SMU      = smu_name,
    CU       = dplyr::if_else(Resolution == "SMU", "-", Name),
    Outlook  = dplyr::coalesce(Outlook, ""),
    Forecast = dplyr::coalesce(Forecast, "")
  ) %>%
  dplyr::select(smu_area, smu_species, SMU, Resolution, CU, Outlook, Forecast)

# ----------------------------
# Define custom slide order
# ----------------------------
area_order <- c("YUKON TRANSBOUNDARY", "NORTH COAST", "SOUTH COAST", "FRASER AND INTERIOR")

species_order <- c(
  "Sockeye Lake Type",
  "Sockeye River Type",
  "Sockeye",             # FRASER AND INTERIOR only
  "Pink Even",           # not for NORTH COAST
  "Pink",                # NORTH COAST
  "Chinook",
  "Coho",
  "Chum"
)

# ----------------------------
# Build a named list keyed by "AREA_SPECIES" and sort it
# ----------------------------
keys <- paste(tp_long$smu_area, tp_long$smu_species, sep = "_")
table_list <- split(tp_long[, c("SMU", "Resolution", "CU", "Outlook", "Forecast")], keys)

# Drop empties
table_list <- Filter(function(df) is.data.frame(df) && nrow(df) > 0, table_list)

# Reorder table_list by area and species
reorder_table_list <- function(tbl) {
  nm_split <- do.call(rbind, strsplit(names(tbl), "_"))
  df_order <- data.frame(
    name = names(tbl),
    area = nm_split[,1],
    species = nm_split[,2],
    stringsAsFactors = FALSE
  )
  df_order$area_rank <- match(df_order$area, area_order)
  df_order$species_rank <- match(df_order$species, species_order)
  df_order$species_rank[is.na(df_order$species_rank)] <- max(df_order$species_rank, na.rm = TRUE) + 1
  df_order <- df_order %>%
    arrange(area_rank, species_rank)
  tbl[df_order$name]
}
table_list <- reorder_table_list(table_list)

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

# Shift table down for species text box
species_height <- 0.4
species_y <- map_y
table_y <- table_y + species_height

fraser_map <- "data/maps/fraserMap.png"

# ----------------------------
# Loop to add slides
# ----------------------------
for (nm in names(table_list)) {
  df_full <- table_list[[nm]]
  area_species <- strsplit(nm, "_")[[1]]
  area <- area_species[1]
  species <- area_species[2]

  # Split FRASER AND INTERIOR Chinook/Sockeye by SMU
  if (area == "FRASER AND INTERIOR" & species %in% c("Chinook", "Sockeye")) {
    smus <- unique(df_full$SMU)
    for (smu_val in smus) {
      df <- df_full %>% filter(SMU == smu_val)
      # create slide
      ft <- flextable(df) |> adjust_font()
      col_count <- ncol(df)
      col_width <- table_width / col_count
      ft <- width(ft, j = 1:col_count, width = col_width)

      ppt <- add_slide(ppt, layout = layout_name, master = master_name)
      ppt <- ph_with(ppt, value = paste0("OUTLOOKS – ", area), location = ph_location_type(type = "title"))
      ppt <- ph_with(ppt, value = ft, location = ph_location(left = table_x, top = table_y, width = table_width, height = table_height))

      species_text <- fpar(
        ftext(species, prop = fp_text(color = "black", font.size = 20, font.family = "Segoe UI Semibold", bold = TRUE)),
        fp_p = fp_par(text.align = "left")
      )
      ppt <- ph_with(ppt, value = species_text, location = ph_location(left = table_x, top = species_y, width = table_width, height = species_height))

      if (file.exists(fraser_map)) {
        ppt <- ph_with(
          ppt,
          value = external_img(fraser_map, width = map_width, height = map_height),
          location = ph_location(left = map_x, top = map_y, width = map_width, height = map_height)
        )
      }

      # Add notes
      if (all(c("smu_area", "smu_species", "Narrative") %in% names(tabPrep))) {
        narratives_df <- tabPrep %>%
          filter(trimws(smu_area) == trimws(area),
                 trimws(smu_species) == trimws(species),
                 trimws(smu_name) == trimws(smu_val)) %>%
          distinct(smu_name, Narrative)
        notes_text <- if (nrow(narratives_df) > 0) {
          narratives_df %>%
            mutate(note_line = paste0(smu_name, ": ", coalesce(Narrative, ""))) %>%
            pull(note_line) %>% unique() %>% paste(collapse = "\n")
        } else {
          "No notes available"
        }
        ppt <- set_notes(ppt, value = notes_text, location = notes_location_type("body"))
      }
    }
  } else {
    # Regular slide for all other species/areas
    df <- df_full
    ft <- flextable(df) |> adjust_font()
    col_count <- ncol(df)
    col_width <- table_width / col_count
    ft <- width(ft, j = 1:col_count, width = col_width)

    ppt <- add_slide(ppt, layout = layout_name, master = master_name)
    ppt <- ph_with(ppt, value = paste0("OUTLOOKS – ", area), location = ph_location_type(type = "title"))
    ppt <- ph_with(ppt, value = ft, location = ph_location(left = table_x, top = table_y, width = table_width, height = table_height))

    species_text <- fpar(
      ftext(species, prop = fp_text(color = "black", font.size = 20, font.family = "Segoe UI Semibold", bold = TRUE)),
      fp_p = fp_par(text.align = "left")
    )
    ppt <- ph_with(ppt, value = species_text, location = ph_location(left = table_x, top = species_y, width = table_width, height = species_height))

    if (file.exists(fraser_map)) {
      ppt <- ph_with(
        ppt,
        value = external_img(fraser_map, width = map_width, height = map_height),
        location = ph_location(left = map_x, top = map_y, width = map_width, height = map_height)
      )
    }

    # Add notes
    if (all(c("smu_area", "smu_species", "Narrative") %in% names(tabPrep))) {
      narratives_df <- tabPrep %>%
        filter(trimws(smu_area) == trimws(area),
               trimws(smu_species) == trimws(species)) %>%
        distinct(smu_name, Narrative)
      notes_text <- if (nrow(narratives_df) > 0) {
        narratives_df %>%
          mutate(note_line = paste0(smu_name, ": ", coalesce(Narrative, ""))) %>%
          pull(note_line) %>% unique() %>% paste(collapse = "\n")
      } else {
        "No notes available"
      }
      ppt <- set_notes(ppt, value = notes_text, location = notes_location_type("body"))
    }
  }
}

# ----------------------------
# Save presentation
# ----------------------------
print(ppt, target = "updated_presentation.pptx")