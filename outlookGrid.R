library(dplyr)
library(ggplot2)

# =========================
# Prep data
# =========================

# cut off "SALMON" from SMU name
# =========================
# Prep data
# =========================

# Clean SMU names
tp_long <- tp_long %>%
  mutate(
    SMU = str_remove(SMU, regex("\\bSALMON\\b", ignore_case = TRUE)),
    SMU = str_remove(SMU, regex("\\p{Pd}\\s*CHECK:?\\s*.*$", ignore_case = TRUE)),
    SMU = str_squish(SMU)
  )

# Add row/column IDs
tp_plot <- tp_long %>%
  arrange(smu_area, smu_species, SMU) %>%
  group_by(smu_area, smu_species, SMU) %>%
  mutate(
    entry = row_number(),
    x_id  = entry
  ) %>%
  ungroup() %>%
  group_by(smu_area, smu_species) %>%
  mutate(
    row_id = dense_rank(SMU)
  ) %>%
  ungroup()

# Right-align tiles: compute reversed x positions per area
tp_plot <- tp_plot %>%
  group_by(smu_area) %>%
  mutate(
    max_x = max(x_id),
    x_rev = max_x - x_id + 1
  ) %>%
  ungroup()

# SMU labels
smu_labels <- tp_plot %>%
  distinct(smu_area, smu_species, SMU, row_id, max_x) %>%
  mutate(
    SMU_wrapped = str_wrap(SMU, width = 18),
    x_label = max_x + 0.3  # small nudge right of last tile
  )

# =========================
# Plot
# =========================

p <- ggplot(tp_plot, aes(x = x_rev, y = row_id, fill = Outlook)) +
  geom_tile(
    width  = 1,   # width = 1 unit
    height = 1,   # height = 1 unit
    color  = NA   # no gaps between tiles
  ) +
  geom_text(
    data = smu_labels,
    aes(x = x_label, y = row_id, label = SMU_wrapped),
    inherit.aes = FALSE,
    hjust = 1,
    size  = 2.5,
    lineheight = 0.95
  ) +
  facet_grid(
    smu_species ~ smu_area,
    scales = "free_y",
    space  = "free_y"
  ) +
  scale_y_reverse() +
  scale_fill_manual(
    values = c(
      "1"              = "#d73027",
      "1 to 2"         = "#f46d43",
      "2"              = "#fdae61",
      "2 to 3"         = "#fee08b",
      "3"              = "#a6d96a",
      "3 to 4"         = "#1a9850",
      "Data Deficient" = "grey70"
    ),
    name = "Outlook"
  ) +
  labs(x = NULL, y = NULL) +
  theme_minimal(base_size = 11) +
  theme(
    panel.grid       = element_blank(),
    axis.text        = element_blank(),
    axis.ticks       = element_blank(),
    strip.background = element_rect(fill = "grey80"),
    strip.text.x     = element_text(face = "bold", color = "grey30"),
    strip.text.y     = element_text(face = "bold", color = "grey30"),
    legend.position  = "bottom",
    plot.margin      = margin(5.5, 5.5, 5.5, 140),
    panel.spacing.x  = unit(1, "lines"),   # space between area facets
    panel.spacing.y  = unit(0.8, "lines")
  )

p