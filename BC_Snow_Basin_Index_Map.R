# ============================================================
# BC Snow Basin Index — RFC divisions + neighbors (grey) + zoom to BC
# ============================================================

library(sf)
library(dplyr)
library(ggplot2)
library(rnaturalearth)

# 1) Read FeatureServer data and project to BC Albers
url <- "https://services6.arcgis.com/ubm4tcTYICKBpist/arcgis/rest/services/Snow_Basins_Indices_View/FeatureServer/0/query?where=1%3D1&outFields=*&f=geojson"
basins <- st_read(url, quiet = TRUE) |> st_transform(3005)

# 2) RFC bins (exact)
legend_levels <- c(
  "No Data", "0-49%", "50-59%", "60-69%", "70-79%", "80-89%",
  "90-109%", "110-125%", "125-140%", "> 140%"
)

basins <- basins |>
  mutate(
    sbi_cat = case_when(
      is.na(Snow_Basin_Index) ~ "No Data",
      Snow_Basin_Index < 50   ~ "0-49%",
      Snow_Basin_Index < 60   ~ "50-59%",
      Snow_Basin_Index < 70   ~ "60-69%",
      Snow_Basin_Index < 80   ~ "70-79%",
      Snow_Basin_Index < 90   ~ "80-89%",
      Snow_Basin_Index < 110  ~ "90-109%",
      Snow_Basin_Index < 125  ~ "110-125%",
      Snow_Basin_Index < 140  ~ "125-140%",
      TRUE                    ~ "> 140%"
    ),
    sbi_cat = factor(sbi_cat, levels = legend_levels)
  )

# 2b) Ensure all legend bins show even if unused
missing_bins <- setdiff(legend_levels, as.character(unique(basins$sbi_cat)))
if (length(missing_bins) > 0) {
  dummy <- st_sf(
    sbi_cat  = factor(missing_bins, levels = legend_levels),
    geometry = st_sfc(rep(st_geometrycollection(), length(missing_bins)), crs = 3005)
  )
  basins <- dplyr::bind_rows(basins, dummy)
}

# 3) Colours (unchanged style—adjust only if you want exact RFC swatches)


sbi_cols <- c(
  "No Data"   = "#ffffff",
  "0-49%"     = "#c23b30",  # a bit deeper red for contrast
  "50-59%"    = "#dd6b3b",
  "60-69%"    = "#f19b52",
  "70-79%"    = "#f7cf75",
  "80-89%"    = "#d6eaa0",
  "90-109%"   = "#a9d1a6",
  "110-125%"  = "#7bc0c0",  # slightly brighter than before
  "125-140%"  = "#4f83b6",
  "> 140%"    = "#275f9e"   # a bit darker to bookend the ramp
)


# 4) Neighboring boundaries from rnaturalearth
can_states <- ne_states(country = "canada", returnclass = "sf") |> st_transform(3005)
us_states  <- ne_states(country = "united states of america", returnclass = "sf") |> st_transform(3005)

bc_poly <- can_states |> dplyr::filter(name == "British Columbia")
neighbors_grey <- dplyr::bind_rows(
  can_states |> dplyr::filter(name != "British Columbia"),
  us_states  |> dplyr::filter(name %in% c("Washington", "Idaho", "Montana", "Alaska"))
)

# 5) Zoom to BC extent (tight bbox with padding)
#    Padding = 20 km on each side (adjust as needed)
pad_m <- 20000
bc_bbox <- st_bbox(st_buffer(st_union(st_geometry(bc_poly)), pad_m))

# 6) Plot: neighbors in grey + basins on top (your theme/legend preserved)
ggplot() +
  geom_sf(data = bc_poly, fill = NA, colour = "grey40", linewidth = 0.4)+
  geom_sf(data = neighbors_grey, fill = "grey88", colour = "grey60", linewidth = 0.25) +
 # geom_sf(data = bc_poly, fill = NA, colour = "grey60", linewidth = 0.25) +  # optional BC outline
  geom_sf(data = basins, aes(fill = sbi_cat), linewidth = 0.25,) +
  scale_fill_manual(
    values = sbi_cols,
    breaks = legend_levels,
    limits = legend_levels,
    name = "Percent of normal",
    drop = FALSE
  ) +
  coord_sf(
    xlim = c(bc_bbox["xmin"], bc_bbox["xmax"]),
    ylim = c(bc_bbox["ymin"], bc_bbox["ymax"]),
    expand = FALSE, datum = NA
  ) +
  theme_minimal(base_size = 11) +
  theme(
    panel.grid = element_blank(),
    axis.text = element_blank(),
    axis.title = element_blank(),
    legend.position = "right",
    legend.title = element_text(size = 10),
    legend.text = element_text(size = 9)
  )