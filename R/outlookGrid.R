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

# For the final image, see page 52 in the 2026 Outlook Report:
# https://waves-vagues.dfo-mpo.gc.ca/library-bibliotheque/41316605.pdf

# It is easiest to not have it add to the PowerPoint/reports automatically, since
# it's a bit unpredicatable what the ideal size is for the image.
# Instead, just figure out the exact size in the plotting window, and then save it
# as an image. Then copy/paste into the Reports/PowerPoints as required.

################################################################################

# Read in the script that creates the tables for the PowerPoint
# It uses the data from that
source("R/powerPointTables.R")

################################################################################
## Load required libraries
library(dplyr)
library(ggplot2)
library(stringr)
library(grid)

################################################################################
## Make a few edits to the data

tp_clean = tp_long %>%

# Make some edits to the data frame
mutate(
  # Remove the word "SALMON" from the SMU name
  SMU = str_remove(SMU, regex("\\bSALMON\\b", ignore_case = TRUE)),
  SMU = str_squish(SMU),
  # Change "Data Deficient" to "DD" for display over top of grid cells
  Outlook = recode(Outlook, "Data Deficient" = "DD")
)

################################################################################
## FACET ORDERING

# This is the preferred order for Areas (so results are North --> South)
area_levels = c(
  "YUKON AND TRANSBOUNDARY",
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

# These variables control the physical layout of each SMU row

# Total horizontal space allocated to Outlook squares per SMU
TOTAL_WIDTH = 3
# space between SMU text labels and Outlook squares
LABEL_PAD   = 0.2
# vertical height of each SMU row
ROW_HEIGHT  = 0.8

# Each SMU can have one or more Outlook (e.g.,  multiple CUs)
# Instead of using geom_col(), rectangle positions are manually computed. This allows:
# -Multiple Outlooks for the same SMU to appear side-by-side
# -consistent spacing across facets

# For each SMU:
# -TOTAL_WIDTH is divided evenly by the number of Outlooks
# -Each Outlook is assigned a segment ID (seg_id)
# -xmin/xmax are computed later using these segment widths
tp_plot = tp_clean %>%
  mutate(SMU = factor(SMU, levels = smu_levels)) %>%
  arrange(smu_area, smu_species, SMU, Outlook) %>%
  group_by(smu_area, smu_species, SMU) %>%
  # Add 3 new columns
  mutate(
    n_outlooks = n(), # how many Outlooks this SMU has
    seg_width  = TOTAL_WIDTH / n_outlooks, # width of each Outlook segment
    seg_id     = row_number() # left-to-right position within SMU
  ) %>%
  ungroup() %>%
  # row_id controls vertical placement of SMUs within each facet
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

# WRAP_WIDTH was chosen by trial and error so labels generally fit within the
# left margin without overlapping with the Outlook squares
WRAP_WIDTH  = 18

# If they exceed WRAP_LENGTH, the text will be written over 2 lines instead of 1
smu_labels = tp_plot %>%
  distinct(smu_area, smu_species, SMU, row_id) %>%
  mutate(
    SMU_wrapped = str_wrap(as.character(SMU), WRAP_WIDTH)
  )

################################################################################
# GLOBAL LABEL WIDTH

# Measure the widest SMU Label (after wrapping) in millimeters
# This lets us reserve enough horizontal space so that:
# - no SMU label overlaps the Outlook squares
# -all facets line up cleanly, regardless of label length

# grid::stringWidth() requires a viewport to measure text size
pushViewport(viewport())
GLOBAL_LABEL_MM = max(
  convertUnit(stringWidth(smu_labels$SMU_wrapped),
              "mm", valueOnly = TRUE)
)
popViewport()

MM_PER_TILE = 20
GLOBAL_LABEL_WIDTH <- GLOBAL_LABEL_MM / MM_PER_TILE + LABEL_PAD


################################################################################
# Rectangle coordinates

# Compute the exact x-positions for each Outlook rectangle
# Rectangles are drawn from left to right so that:
# -single-Outlook SMUs align with multi-Outlook SMUs
# -segment ordering is visually consistent

# xmax/xmin are the coordinate positions of the outer edges of each rectangle
tp_plot = tp_plot %>%
  mutate(
    xmax   = GLOBAL_LABEL_WIDTH + TOTAL_WIDTH - (seg_id - 1) * seg_width,
    xmin   = xmax - seg_width,
    x_mid  = (xmin + xmax) / 2 # used to center the Outlook text label
  )

# SMU text labels are right-aligned just to the left of the Outlook squares
smu_labels = smu_labels %>%
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
