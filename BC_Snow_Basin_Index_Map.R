# ============================================================
# BC Snow Basin Index Map — RFC-aligned, softened colours
# ============================================================

library(sf)
library(dplyr)
library(ggplot2)

# ------------------------------------------------------------
# 1. Read Snow Basin Index data
# ------------------------------------------------------------
url <- "https://services6.arcgis.com/ubm4tcTYICKBpist/arcgis/rest/services/Snow_Basins_Indices_View/FeatureServer/0/query?where=1%3D1&outFields=*&f=geojson"

basins <- st_read(url, quiet = TRUE)

# ------------------------------------------------------------
# 2. Project to BC Albers
# ------------------------------------------------------------
basins <- st_transform(basins, 3005)

# ------------------------------------------------------------
# 3. RFC-style bins + explicit No Data category
#    (matches original legend structure)
# ------------------------------------------------------------
basins <- basins %>%
  mutate(
    sbi_cat = case_when(
      is.na(Snow_Basin_Index) ~ "No Data",
      Snow_Basin_Index < 50   ~ "0–49%",
      Snow_Basin_Index < 70   ~ "50–69%",
      Snow_Basin_Index < 90   ~ "70–89%",
      Snow_Basin_Index < 110  ~ "90–109%",
      Snow_Basin_Index < 125  ~ "110–124%",
      TRUE                    ~ "≥ 125%"
    ),
    sbi_cat = factor(
      sbi_cat,
      levels = c(
        "No Data",
        "0–49%",
        "50–69%",
        "70–89%",
        "90–109%",
        "110–124%",
        "≥ 125%"
      )
    )
  )

# ------------------------------------------------------------
# 4. Softer, professional colour palette
#    (RFC logic, less aggressive saturation)
# ------------------------------------------------------------
sbi_cols <- c(
  "No Data"  = "#e6e6e6",
  "0–49%"    = "#d08b80",
  "50–69%"   = "#e6b27d",
  "70–89%"   = "#f2e0a1",
  "90–109%"  = "#b9dcb5",
  "110–124%" = "#7fc9c2",
  "≥ 125%"   = "#4f83b6"
)

# ------------------------------------------------------------
# 5. Build map (PowerPoint-ready)
# ------------------------------------------------------------
ggplot(basins) +
  geom_sf(aes(fill = sbi_cat), linewidth = 0.25) +
  scale_fill_manual(
    values = sbi_cols,
    name = "Percent of normal",
    drop = FALSE
  ) +
  coord_sf(datum = NA) +
  theme_minimal(base_size = 11) +
  theme(
    panel.grid = element_blank(),
    axis.text = element_blank(),
    axis.title = element_blank(),
    legend.position = "right",
    legend.title = element_text(size = 10),
    legend.text = element_text(size = 9),
    plot.caption = element_text(size = 8, colour = "grey40")
  )