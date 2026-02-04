################################################################################
#### BC Snow Basin Index Map

# Made by Stephen Finnis in February 2026
# This is to make a map used in the "Major Stocks/Infographic Presentation"
# The original is a bit blurry and crowded
# See here for original: https://www2.gov.bc.ca/assets/gov/environment/air-land-water/water/river-forecast/2026_jan1.pdf
# There will be a new one every year

# You can search for the data online in the BC provincial data catalogue: https://catalogue.data.gov.bc.ca/dataset/snow-basin-indices
# And then choose the ESRI Rest Server data

################################################################################

# Load appropriate libraries
library(sf)
library(dplyr)
library(ggplot2)
library(rnaturalearth)

################################################################################
# 1) Read Snow Basin Index data and project to BC Albers (crs = 3005)

# Read in the data from the ESRI Rest server (this will be for 2026 data)
url = "https://services6.arcgis.com/ubm4tcTYICKBpist/arcgis/rest/services/Snow_Basins_Indices_View/FeatureServer/0/query?where=1%3D1&outFields=*&f=geojson"

# Read in the basins data and transform
basins = st_read(url, quiet = TRUE) %>%
  st_transform(3005)

################################################################################

# Define legend bins
# This just copies the ones that the province used
# Could probably reduce number of bins or justify other choices

# These are the
legend_levels = c(
  "No Data", "0-49%", "50-59%", "60-69%", "70-79%", "80-89%",
  "90-109%", "110-125%", "125-140%", "> 140%"
)

# Add new column (sbi_cat) that has the legend categories I want
# Reclassifies the data into these bins below
basins = basins |>
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

################################################################################
# Force all legend bins to appear (even if unused)

# Trick to get all the items to show up in the legend
# Some data bins aren't present in data, so you need to have a workaround
# Adds these as dummy data to the basins data frame, so they still show up
missing_bins = setdiff(legend_levels, as.character(unique(basins$sbi_cat)))

basins = basins %>%
  bind_rows(
    st_sf(
      sbi_cat  = factor(missing_bins, levels = legend_levels),
      geometry = st_sfc(rep(st_geometrycollection(), length(missing_bins)), crs = 3005)
    )
  )
################################################################################
# Choose a colour scheme for these bins
sbi_cols = c(
  "No Data"   = "#f0f0f0",
  "0-49%"     = "#b22222",
  "50-59%"    = "#d95f0e",
  "60-69%"    = "#f18f3b",
  "70-79%"    = "#f6c96b",
  "80-89%"    = "#d9e49c",
  "90-109%"   = "#8fc38a",
  "110-125%"  = "#5fb3b3",
  "125-140%"  = "#3f7fbf",
  "> 140%"    = "#1f4e99"
)

################################################################################
# Get data for neighbouring states/provinces and reproject to BC Albers

# Get Canadian province boundaries. Exclude BC so it doesnt cause overlap issues
can_prov = ne_states(country = "canada", returnclass = "sf") %>%
  filter(name != "British Columbia") %>%
  st_transform(3005)

# Get American States
us_states  = ne_states(country = "united states of america", returnclass = "sf") %>%
  st_transform(3005)

################################################################################
# Get map extents of the BC map
# Idk why it wasn't working, but i have to "make the geometries valid" because
# there were topology gaps or something

basins_valid = basins %>%
  st_make_valid() %>%
  st_buffer(0)

# Then dissolve the boundaries, and create a 20km (20000m) buffer
# If you don't add the 20km buffer, map extent will go right to corners of BC map
bc_bbox = basins_valid %>%
  st_union() %>%
  st_buffer(20000) %>%
  st_bbox()

################################################################################
# Make the map

ggplot() +
  geom_sf(data = can_prov, fill = "grey85", colour = "grey60", linewidth = 0.25) +
  geom_sf(data = us_states, fill = "grey85", colour = "grey60", linewidth = 0.25) +
  geom_sf(data = basins, aes(fill = sbi_cat), linewidth = 0.25) +
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
    panel.border = element_rect(colour = "black", fill = NA, linewidth = 1),
    axis.text = element_blank(),
    axis.title = element_blank(),
    legend.position = "right",
    legend.title = element_text(size = 17),
    legend.text = element_text(size = 14)
  )