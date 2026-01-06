
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
    Outlook = recode(Outlook, "Data Deficient" = "DD")
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
  "Sockeye Lake Type",
  "Sockeye River Type",
  "Pink Even",
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

    legend.position  = "bottom",

    panel.spacing.x  = unit(0.8, "lines"),
    panel.spacing.y  = unit(0.8, "lines")
  )

p




# ============================================================
# SECOND PLOT: Vertically listed (stacked) regional facets
# ============================================================

library(dplyr)
library(ggplot2)
library(stringr)
library(forcats)

# -------------------------
# SETTINGS (independent)
# -------------------------
TOTAL_WIDTH_STACKED <- 3
LABEL_PAD_STACKED   <- 0.5

# -------------------------
# PREP DATA (new objects)
# -------------------------
tp_stacked <- tp_long %>%
  mutate(
    SMU_clean = str_remove(SMU, " Chinook| Chum| Coho| Pink| Sockeye Lake Type| Sockeye River Type "),

    # count outlooks per SMU
    n_outlooks = n(),

    # fixed-width slices
    slice_width = TOTAL_WIDTH_STACKED / n_outlooks
  ) %>%
  group_by(smu_area, smu_species, SMU_clean) %>%
  mutate(
    xmin = (row_number() - 1) * slice_width,
    xmax = row_number() * slice_width
  ) %>%
  ungroup()

# -------------------------
# ORDER REGIONS TOP → BOTTOM
# (edit this to match desired order)
# -------------------------



# ============================================================
# SECOND PLOT: Stacked areas + species (CSAS-style layout)
# ============================================================


library(dplyr)
library(ggplot2)
library(stringr)
library(grid)

TOTAL_WIDTH2 <- TOTAL_WIDTH
LABEL_PAD2   <- LABEL_PAD
WRAP_WIDTH2  <- WRAP_WIDTH

ROW_HALF_HEIGHT <- 0.45

# ------------------------------------------------------------
# Clean + factor levels
# ------------------------------------------------------------
tp_clean2 <- tp_long %>%
  mutate(
    SMU = str_remove(SMU, regex("\\bSALMON\\b", ignore_case = TRUE)),
    SMU = str_remove(SMU, regex("\\p{Pd}\\s*CHECK:?\\s*.*$", ignore_case = TRUE)),
    SMU = str_squish(SMU),
    smu_area    = factor(smu_area, levels = area_levels),
    smu_species = factor(smu_species, levels = species_levels)
  )

# ------------------------------------------------------------
# SMU rows (outlook segment logic unchanged)
# ------------------------------------------------------------
smu_rows2 <- tp_clean2 %>%
  arrange(smu_area, smu_species, SMU, Outlook) %>%
  group_by(smu_area, smu_species, SMU) %>%
  mutate(
    n_outlooks = n(),
    seg_width  = TOTAL_WIDTH2 / n_outlooks,
    seg_id     = row_number(),
    row_type   = "smu"
  ) %>%
  ungroup()

# ------------------------------------------------------------
# Species header rows (text only)
# ------------------------------------------------------------
species_rows2 <- smu_rows2 %>%
  distinct(smu_area, smu_species) %>%
  mutate(
    SMU        = as.character(smu_species),
    Outlook    = NA,
    seg_width  = NA,
    seg_id     = NA,
    row_type   = "species"
  )

# ------------------------------------------------------------
# Stack rows per AREA (species ABOVE SMUs)
# ------------------------------------------------------------
stacked_df2 <- bind_rows(species_rows2, smu_rows2) %>%
  arrange(
    smu_area,
    smu_species,
    row_type,   # species first
    SMU,
    seg_id
  ) %>%
  group_by(smu_area) %>%
  mutate(row_id = row_number()) %>%
  ungroup()

# ------------------------------------------------------------
# Wrap labels (no height distortion)
# ------------------------------------------------------------
label_df2 <- stacked_df2 %>%
  distinct(smu_area, SMU, row_id, row_type) %>%
  mutate(
    SMU_wrapped = str_wrap(SMU, WRAP_WIDTH2)
  )

# ------------------------------------------------------------
# Measure label width per AREA (same as first plot)
# ------------------------------------------------------------
pushViewport(viewport())
label_widths2 <- label_df2 %>%
  group_by(smu_area) %>%
  summarise(
    LABEL_MM = max(convertUnit(stringWidth(SMU_wrapped), "mm", valueOnly = TRUE)),
    .groups = "drop"
  )
popViewport()

MM_PER_TILE <- 20

label_widths2 <- label_widths2 %>%
  mutate(LABEL_WIDTH = LABEL_MM / MM_PER_TILE + LABEL_PAD2)

# ------------------------------------------------------------
# Compute rectangle positions
# ------------------------------------------------------------
plot_df2 <- stacked_df2 %>%
  left_join(label_widths2, by = "smu_area") %>%
  mutate(
    xmax = if_else(
      row_type == "smu",
      LABEL_WIDTH + TOTAL_WIDTH2 - (seg_id - 1) * seg_width,
      NA_real_
    ),
    xmin = if_else(
      row_type == "smu",
      xmax - seg_width,
      NA_real_
    )
  )

label_df2 <- label_df2 %>%
  left_join(label_widths2, by = "smu_area") %>%
  mutate(x_label = LABEL_WIDTH - LABEL_PAD2)




stacked_df2 <- bind_rows(
  species_rows2,
  smu_rows2
) %>%
  arrange(
    smu_area,
    smu_species,
    row_type,   # species first
    SMU,
    seg_id
  ) %>%
  group_by(smu_area) %>%
  mutate(row_id = row_number()) %>%
  ungroup()



p_stacked <- ggplot(plot_df2) +

  # Species header boxes
  geom_rect(
    data = subset(plot_df2, row_type == "species"),
    aes(
      xmin = 0,
      xmax = LABEL_WIDTH + TOTAL_WIDTH2,
      ymin = row_id - ROW_HALF_HEIGHT,
      ymax = row_id + ROW_HALF_HEIGHT
    ),
    fill = "grey85",
    color = "grey40",
    linewidth = 0.3
  ) +

  # Outlook squares
  geom_rect(
    data = subset(plot_df2, row_type == "smu"),
    aes(
      xmin = xmin,
      xmax = xmax,
      ymin = row_id - ROW_HALF_HEIGHT,
      ymax = row_id + ROW_HALF_HEIGHT,
      fill = Outlook
    ),
    color = NA
  ) +

  # Outlook text
  geom_text(
    data = subset(plot_df2, row_type == "smu"),
    aes(
      x = (xmin + xmax) / 2,
      y = row_id,
      label = Outlook
    ),
    fontface = "bold",
    size = 2.6
  ) +

  # SMU labels
  geom_text(
    data = subset(label_df2, row_type == "smu"),
    aes(
      x = x_label,
      y = row_id,
      label = SMU_wrapped
    ),
    hjust = 1,
    size = 2.8,
    inherit.aes = FALSE
  ) +

  # Species labels (centered)
  geom_text(
    data = subset(label_df2, row_type == "species"),
    aes(
      x = (LABEL_WIDTH + TOTAL_WIDTH2) / 2,
      y = row_id,
      label = SMU_wrapped
    ),
    hjust = 0.5,
    fontface = "bold",
    size = 3.2,
    inherit.aes = FALSE
  ) +

  # ✅ CORRECT FACETING
  facet_grid(
    . ~ smu_area,
    scales = "free_y",
    space  = "free_y"
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
    ))+
  scale_y_continuous(expand = c(0, 0)) +
  scale_x_continuous(
    limits = c(0, max(label_widths2$LABEL_WIDTH) + TOTAL_WIDTH2),
    expand = c(0, 0)
  ) +

  theme_minimal(base_size = 11) +
  theme(
    panel.grid       = element_blank(),
    axis.text        = element_blank(),
    axis.ticks       = element_blank(),
    strip.background = element_rect(fill = "grey80", color = "grey40"),
    strip.text       = element_text(face = "bold"),
    panel.spacing.x  = unit(1, "lines"),
    plot.margin      = margin(10, 10, 10, 80)
  )

p_stacked
