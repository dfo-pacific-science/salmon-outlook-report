
library(dplyr)
library(ggplot2)
library(stringr)
library(grid)

# =========================
# SETTINGS
# =========================

TOTAL_WIDTH <- 3
LABEL_PAD   <- 0.2
WRAP_WIDTH  <- 18
ROW_HEIGHT  <- 0.8

# =========================
# CLEAN DATA
# =========================

tp_clean <- tp_long %>%
  mutate(
    SMU = str_remove(SMU, regex("\\bSALMON\\b", ignore_case = TRUE)),
    SMU = str_remove(SMU, regex("\\p{Pd}\\s*CHECK:?\\s*.*$", ignore_case = TRUE)),
    SMU = str_squish(SMU),
    Outlook = recode(Outlook, "Data Deficient" = "DD"),

  ### Update Species
    smu_species = case_when(
      smu_species %in% c("Sockeye Lake Type", "Sockeye River Type") ~ "Sockeye",
      smu_species == "Pink Even" ~ "Pink",
      TRUE ~ smu_species)
    ) %>%

  mutate(
    Outlook = case_when(
      CU == "MIDDLE FRASER-FRASER CANYON_SP_1.3, LOWER FRASER RIVER_SP_1.3" ~ 1,
      CU == "MIDDLE FRASER RIVER_SP_1.3, NORTH THOMPSON_SP_1.3" ~ 2,
      CU == "MIDDLE FRASER RIVER-PORTAGE_FA_1.3, LOWER FRASER RIVER-UPPER PITT_SU_1.3" ~ 1,
      CU == "MIDDLE FRASER RIVER_SU_1.3, NORTH THOMPSON_SU_1.3" ~ "2 to 3",
      TRUE ~ Outlook
    )
  )


# =========================
# FACET ORDERING
# =========================

area_levels <- c(
  "YUKON TRANSBOUNDARY",
  "NORTH COAST",
  "SOUTH COAST",
  "FRASER AND INTERIOR"
)

species_levels <- c(
  "Sockeye",
  "Pink",
  "Chinook",
  "Coho",
  "Chum"
)

tp_clean <- tp_clean %>%
  mutate(
    smu_area    = factor(smu_area, levels = area_levels),
    smu_species = factor(smu_species, levels = species_levels)
  )

# =========================
# GLOBAL SMU ORDER
# =========================

smu_levels <- tp_clean %>%
  distinct(SMU) %>%
  arrange(SMU) %>%
  pull(SMU)

# =========================
# COMPUTE SEGMENTS
# =========================

tp_plot <- tp_clean %>%
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
  ungroup()

# =========================
# LABEL DATA
# =========================

smu_labels <- tp_plot %>%
  distinct(smu_area, smu_species, SMU, row_id) %>%
  mutate(
    SMU_wrapped = str_wrap(as.character(SMU), WRAP_WIDTH)
  )

# =========================
# GLOBAL LABEL WIDTH
# =========================

pushViewport(viewport())
GLOBAL_LABEL_MM <- max(
  convertUnit(stringWidth(smu_labels$SMU_wrapped),
              "mm", valueOnly = TRUE)
)
popViewport()

MM_PER_TILE <- 20
GLOBAL_LABEL_WIDTH <- GLOBAL_LABEL_MM / MM_PER_TILE + LABEL_PAD

# =========================
# RECT COORDINATES
# =========================

tp_plot <- tp_plot %>%
  mutate(
    xmax   = GLOBAL_LABEL_WIDTH + TOTAL_WIDTH - (seg_id - 1) * seg_width,
    xmin   = xmax - seg_width,
    x_mid  = (xmin + xmax) / 2
  )

smu_labels <- smu_labels %>%
  mutate(x_label = GLOBAL_LABEL_WIDTH - LABEL_PAD)

# =========================
# OUTLOOK LABELS
# =========================

outlook_labels <- c(
  "1"  = "1: Well below average",
  "2"  = "2: Below average",
  "3"  = "3: Near average",
  "4"  = "4: Abundant",
  "DD" = "DD: Data Deficient"
)

# =========================
# PLOT
# =========================

p <- ggplot(tp_plot) +
  geom_rect(
    aes(
      xmin = xmin,
      xmax = xmax,
      ymin = row_id - ROW_HEIGHT / 2,
      ymax = row_id + ROW_HEIGHT / 2,
      fill = Outlook
    ),
    color = NA
  ) +
  geom_text(
    aes(
      x = x_mid,
      y = row_id,
      label = Outlook
    ),
    fontface = "bold",
    size = 2.8,
    vjust = 0.5,
    hjust = 0.5
  ) +
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
    size = 2.6
  ) +
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
  scale_fill_manual(
    values = c(
      "1"  = "#d73027",
      "1 to 2" = "#f46d43",
      "2"  = "#fdae61",
      "2 to 3" = "#fee08b",
      "3"  = "#a6d96a",
      "3 to 4" = "#66bd63",
      "4"  = "#1a9850",   # darkest green
      "DD" = "grey70"
    ),
    labels = outlook_labels,
    name   = "Outlook"
  ) +
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
    # strip.background = element_rect(
    #   fill = "grey85",
    #   colour = "grey40",
    #   linewidth = 0.6
    # ),
    strip.text.x = element_text(
      face = "bold",
      size = 11
    ),
    strip.text.y = element_text(
      face = "bold",
      size = 11,
      lineheight = 0.9
    ),

    legend.position  = "right",

    panel.spacing.x  = unit(0.8, "lines"),
    panel.spacing.y  = unit(0.8, "lines")
  )

p


