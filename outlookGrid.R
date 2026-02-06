################################################################################
### OUTLOOK OVERVIEW GRID FIGURE

# Created by Stephen Finnis 2026

# This is the code to make the Outlook Overview figure for the Outlook report
# It gives a breakdown of all the Outlooks in one image
# Has a grid format: Areas make up the columns, salmon species as rows
# Within each grid is the info for Stock Management Units (SMU)
# The full SMU name is written out. Beside it is colours representing the Outlook
# The colours are displayed as squares/rectangles - one for each SMU, or multiple
# if Outlooks were provided per Conservation Unit (CU)

# This grid goes in the Preliminary Outlook presentation (January of each year)
# and the written report that is also submitted at the same time.

# It is easiest to not have it add to the PowerPoint/reports automatically, since
# it's a bit unpredicatable what the ideal size is for the image.
# Instead, just figure out the exact size in the plotting window, and then save it
# as an image. Then copy/paste into the Reports/PowerPoints as required.

################################################################################

# Read in the script that creates the tables for the PowerPoint
# It uses the data from that
source("powerPointTables.R")

################################################################################
## Load required libraries

library(dplyr)
library(ggplot2)
library(stringr)
library(grid)

################################################################################
## Make edits to the data
# Ideally this would be

tp_clean = tp_long %>%


group_by(SMU) %>%
  mutate(
    CU = case_when(
      SMU == "MIDDLE GEORGIA STRAIT CHINOOK SALMON" & row_number() == 1 ~
        "MGS CHINOOK – PLACEHOLDER A",
      SMU == "MIDDLE GEORGIA STRAIT CHINOOK SALMON" & row_number() == 2 ~
        "MGS CHINOOK – PLACEHOLDER B",
      SMU == "MIDDLE GEORGIA STRAIT CHINOOK SALMON" & row_number() == 3 ~
        "MGS CHINOOK – PLACEHOLDER C",
      TRUE ~ CU
    ),
    Outlook = case_when(
      SMU == "MIDDLE GEORGIA STRAIT CHINOOK SALMON" & row_number() == 1 ~ "DD",
      SMU == "MIDDLE GEORGIA STRAIT CHINOOK SALMON" & row_number() == 2 ~ "1 to 2",
      SMU == "MIDDLE GEORGIA STRAIT CHINOOK SALMON" & row_number() == 3 ~ "4",
      TRUE ~ Outlook
    )
  ) %>%
  ungroup() %>%

  # =========================
# SPLIT FINICKY SMUs / CUs
# =========================
{
  base <- .

  # ---- HAIDA GWAII PINK SALMON ----
  haida_template <- base %>%
    filter(SMU == "HAIDA GWAII PINK SALMON") %>%
    slice(1) %>%
    select(-CU, -Outlook, -Resolution)

  haida_split <- bind_rows(
    haida_template %>%
      mutate(
        CU = "EAST HAIDA GWAII",
        Resolution = "CU (singular)",
        Outlook = "4"
      ),
    haida_template %>%
      mutate(
        CU = "NORTH HAIDA GWAII, WEST HAIDA GWAII",
        Resolution = "CU (aggregate)",
        Outlook = "DD"
      )
  )

  # ---- CENTRAL COAST PINK SALMON ----
  central_template <- base %>%
    filter(SMU == "CENTRAL COAST PINK SALMON") %>%
    slice(1) %>%
    select(-CU, -Outlook, -Resolution)

  central_coast_split <- bind_rows(
    central_template %>%
      mutate(
        CU = "HECATE LOWLANDS, HECATE STRAIT-FJORDS",
        Resolution = "CU (aggregate)",
        Outlook = "4"
      ),
    central_template %>%
      mutate(
        CU = "HOMATHKO-KLINAKLINI-SMITH-RIVERS-BELLA-COOLA-DEAN",
        Resolution = "CU (singular)",
        Outlook = "DD"
      )
  )

  # ---- REMOVE ORIGINALS + ADD SPLITS ----
  base %>%
    filter(
      !SMU %in% c(
        "HAIDA GWAII PINK SALMON",
        "CENTRAL COAST PINK SALMON"
      )
    ) %>%
    bind_rows(
      haida_split,
      central_coast_split
    )
} %>%

  # =========================
# STANDARD CLEANING
# =========================
mutate(
  SMU = str_remove(SMU, regex("\\bSALMON\\b", ignore_case = TRUE)),
  SMU = str_remove(SMU, regex("\\p{Pd}\\s*CHECK:?\\s*.*$", ignore_case = TRUE)),
  SMU = str_squish(SMU),
  Outlook = recode(Outlook, "Data Deficient" = "DD"),

  smu_species = case_when(
    smu_species %in% c("Sockeye Lake Type", "Sockeye River Type") ~ "Sockeye",
    smu_species == "Pink Even" ~ "Pink",
    TRUE ~ smu_species
  )
) %>%

  mutate(
    Outlook = case_when(
      CU == "MIDDLE FRASER-FRASER CANYON_SP_1.3, LOWER FRASER RIVER_SP_1.3" ~ "1",
      CU == "MIDDLE FRASER RIVER_SP_1.3, NORTH THOMPSON_SP_1.3" ~ "2",
      CU == "MIDDLE FRASER RIVER-PORTAGE_FA_1.3, LOWER FRASER RIVER-UPPER PITT_SU_1.3" ~ "1",
      CU == "MIDDLE FRASER RIVER_SU_1.3, NORTH THOMPSON_SU_1.3" ~ "2 to 3",
      TRUE ~ as.character(Outlook)
    )
  ) %>%
  filter(
    CU != "CK-02",
    CU != "OKANAGAN_0.X"
  )  %>%

  filter(
    !(smu_area == "SOUTH COAST" &
        smu_species == "Sockeye" &
        SMU == "NO DESIGNATED SMU")
  ) %>%
  mutate(
    SMU = if_else(
      smu_area == "FRASER AND INTERIOR" &
        smu_species == "Chinook" &
        SMU == "NO DESIGNATED SMU",
      "NO DESIGNATED SMU*",
      SMU
    )
  )





################################################################################
## FACET ORDERING

# This is the preferred order for Areas (so results are North --> South)
area_levels = c(
  "YUKON TRANSBOUNDARY",
  "NORTH COAST",
  "SOUTH COAST",
  "FRASER AND INTERIOR"
)

# This is the preferred order for species (idk why)
species_levels = c(
  "Sockeye",
  "Pink",
  "Chinook",
  "Coho",
  "Chum"
)

# Reorganize the data into the specified order above
tp_clean = tp_clean %>%
  mutate(
    smu_area    = factor(smu_area, levels = area_levels),
    smu_species = factor(smu_species, levels = species_levels)
  )

# Extract the SMU names for later plotting


smu_levels = tp_clean %>%
  distinct(SMU) %>% # Get unique SMu names
  arrange(SMU) %>% # Sort them alphabetically
  pull(SMU) # extract them as a vector

################################################################################
# COMPUTE SEGMENTS




TOTAL_WIDTH <- 3
LABEL_PAD   <- 0.2
ROW_HEIGHT  <- 0.8


# Each SMU can have one more Outlooks
# To draw them as side-by-side rectangles:
# -TOTAL_WIDTH is divided evenly by the number of Outlooks
# -Each Outlook gets a segment (seg_id) within the SMU row
# -xmin/xmax are computed manually (not via geom_col)
# This gives better control over spacing and alignment

tp_plot = tp_clean %>%
  mutate(SMU = factor(SMU, levels = smu_levels)) %>%
  arrange(smu_area, smu_species, SMU, Outlook) %>%
  group_by(smu_area, smu_species, SMU) %>%
  mutate(
    n_outlooks = n(),
    seg_width  = TOTAL_WIDTH / n_outlooks,
    seg_id     = row_number()
  ) %>%
  ungroup() %>%
  group_by(smu_area, smu_species) %>%
  mutate(row_id = dense_rank(SMU)) %>%
  ungroup() %>%

################################################################################
## Fix how the "__ to ___" labels show up in the graphic

# For FIA Sockeye there are so many CUs reporting results, that instead of having
# them written out as e.g., "1 to 2", the numbers 1 and 2 are stacked vertically
# I visually identified them on the grid (they belonged to early summer SMU)
# and then corrected them manually

mutate(
  Outlook_label = Outlook,
  Outlook_label = if_else(
    SMU == "FRASER SOCKEYE - EARLY SUMMER" & Outlook %in% c("1 to 2", "3 to 4"),
    str_replace_all(Outlook, " to ", "\n"),
    Outlook_label
  )
)

################################################################################
## Some SMU names are too long to fit on one line
# Did some trial and error to see what is the max length they should be before
# splitting into 2 lines

#
WRAP_WIDTH  = 18


smu_labels = tp_plot %>%
  distinct(smu_area, smu_species, SMU, row_id) %>%
  mutate(
    SMU_wrapped = str_wrap(as.character(SMU), WRAP_WIDTH)
  )

################################################################################
# GLOBAL LABEL WIDTH


pushViewport(viewport())
GLOBAL_LABEL_MM <- max(
  convertUnit(stringWidth(smu_labels$SMU_wrapped),
              "mm", valueOnly = TRUE)
)
popViewport()

MM_PER_TILE <- 20
GLOBAL_LABEL_WIDTH <- GLOBAL_LABEL_MM / MM_PER_TILE + LABEL_PAD


################################################################################
# RECT COORDINATES


tp_plot = tp_plot %>%
  mutate(
    xmax   = GLOBAL_LABEL_WIDTH + TOTAL_WIDTH - (seg_id - 1) * seg_width,
    xmin   = xmax - seg_width,
    x_mid  = (xmin + xmax) / 2
  )

smu_labels <- smu_labels %>%
  mutate(x_label = GLOBAL_LABEL_WIDTH - LABEL_PAD)

################################################################################
## Outlook labels

# Only the whole numbers have textual descriptions. Make sure to clarify them
# so they show up in the legend

outlook_labels = c(
  "1"  = "1: Well below average",
  "2"  = "2: Below average",
  "3"  = "3: Near average",
  "4"  = "4: Abundant",
  "DD" = "DD: Data Deficient"
)

################################################################################
## Make the final plot


p = ggplot(tp_plot) +
  geom_rect(
    aes(
      xmin = xmin,
      xmax = xmax,
      ymin = row_id - ROW_HEIGHT / 2,
      ymax = row_id + ROW_HEIGHT / 2,
      fill = Outlook
    ),
    color = NA,
    # Need to adjust transparency, since they are slightly transparent on the map
    alpha = 0.7
  ) +

  # Add the Outlook numeric value in bold over each square
  geom_text(
    aes(
      x = x_mid,
      y = row_id,
      label = Outlook_label
    ),
    fontface = "bold",
    size = 3.1,
    vjust = 0.5,
    hjust = 0.5,
    lineheight = 0.9
  ) +

  # Add the SMU names
  geom_text(
    data = smu_labels,
    aes(
      x = x_label,
      y = row_id,
      label = SMU_wrapped
    ),
    inherit.aes = FALSE,
    hjust = 1,
    vjust = 0.5,
    lineheight = 1,
    size = 2.8
  ) +

  # Area as columns, species as rows
  facet_grid(
    smu_species ~ smu_area,
    scales = "free_y",
    space  = "free_y"
  ) +
  scale_y_reverse() +
  scale_x_continuous(
    limits = c(0, GLOBAL_LABEL_WIDTH + TOTAL_WIDTH),
    expand = c(0, 0)
  ) +

  # Hex codes are from Chelsea. Set to 70% transparency, as with the maps
  scale_fill_manual(
    values = c(
      "1"  = "#E4453C",
      "1 to 2" = "#FF8181",
      "2"  = "#F3953C",
      "2 to 3" = "#F0D27D",
      "3"  = "#D4EEC7",
      "3 to 4" = "#8BCE69",
      "4"  = "#1C854F",
      "DD" = "#9E9E9E"
    ),
    labels = outlook_labels, # This has the text description beside the whole numbers
    name   = "Outlook"
  ) +

  # Visual updates
  labs(x = NULL, y = NULL) +
  theme_minimal(base_size = 11) +
  theme(
    panel.grid       = element_blank(),
    axis.text        = element_blank(),
    axis.ticks       = element_blank(),

    panel.border = element_rect(
      colour = "grey40",
      fill   = NA,
      linewidth = 0.4
    ),
    strip.background = element_blank(),

    strip.text.x = element_text(
      face = "bold",
      size = 13
    ),
    strip.text.y = element_text(
      face = "bold",
      size = 13,
      lineheight = 0.9
    ),
    legend.position  = "right",
    panel.spacing.x  = unit(0.8, "lines"),
    panel.spacing.y  = unit(0.8, "lines")
  )

# View the plot
p
