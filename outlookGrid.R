library(dplyr)
library(ggplot2)

# =========================
# Prep data
# =========================

# cut off "SALMON" from SMU name

tp_long <- tp_long %>%
  mutate(
    # remove standalone word SALMON (case-insensitive)
    SMU = str_remove(SMU, regex("\\bSALMON\\b", ignore_case = TRUE)),

    # remove “<dash> CHECK …” to end-of-line; dash can be -, –, —
    # colon after CHECK is optional
    SMU = str_remove(SMU, regex("\\p{Pd}\\s*CHECK:?\\s*.*$", ignore_case = TRUE)),

    # squeeze leftover whitespace
    SMU = str_squish(SMU)
  )


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

smu_labels <- tp_plot %>%
  distinct(smu_area, smu_species, SMU, row_id) %>%
  mutate(
    SMU_wrapped = str_wrap(SMU, width = 18)
  )

# =========================
# Plot
# =========================

p <- ggplot(
  tp_plot,
  aes(
    x    = x_id,
    y    = row_id,
    fill = Outlook
  )
) +
  geom_tile(
    width  = 0.7,
    height = 0.7,
    color  = "white"
  ) +
  geom_text(
    data = smu_labels,
    aes(
      x     = 0,
      y     = row_id,
      label = SMU_wrapped
    ),
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
  scale_x_continuous(
    limits = c(-0.1, NA),
    expand = c(0, 0)
  ) +
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
  coord_cartesian(clip = "off")+
  theme_minimal(base_size = 11) +
  theme(
    panel.grid      = element_blank(),
    axis.text       = element_blank(),
    axis.ticks      = element_blank(),
    strip.text.x    = element_text(face = "bold"),
    strip.text.y    = element_text(face = "bold"),
    legend.position = "bottom",
    plot.margin     = margin(5.5, 5.5, 5.5, 140)
  )

p
