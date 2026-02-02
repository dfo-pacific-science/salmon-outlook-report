

# --- 1. Read data from the FeatureServer as GeoJSON ---
library(sf)
library(dplyr)
library(ggplot2)

# 1. Read the FeatureServer as GeoJSON
url <- "https://services6.arcgis.com/ubm4tcTYICKBpist/arcgis/rest/services/Snow_Basins_Indices_View/FeatureServer/0/query?where=1%3D1&outFields=*&f=geojson"

basins <- st_read(url)


# 2) Project to BC Albers
basins <- st_transform(basins, 3005)  # EPSG:3005


# 2. Check fields (optional)
names(basins)

# 3. Pick field to map (correct field from your screenshot)
field <- "Snow_Basin_Index"

# 4. Make a basic map
ggplot(basins) +
  geom_sf(aes(fill = .data[[field]])) +
  labs(
    title = "BC Snow Basin Index",
    fill = field
  ) +
  theme_minimal()