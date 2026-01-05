
library(grid)      # <-- needed for stringWidth / convertUnit

TOTAL_WIDTH <- 3
LABEL_PAD   <- 0.2
WRAP_WIDTH  <- 18
LINE_HEIGHT <- 1

# Clean
tp_clean <- tp_long %>%
  mutate(
    SMU = str_remove(SMU, regex("\\bSALMON\\b", ignore_case = TRUE)),
    SMU = str_remove(SMU, regex("\\p{Pd}\\s*CHECK:?\\s*.*$", ignore_case = TRUE)),
    SMU = str_squish(SMU)
  )

# Compute segments
tp_plot <- tp_clean %>%
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

# Labels
smu_labels <- tp_plot %>%
  distinct(smu_area, smu_species, SMU, row_id) %>%
  mutate(
    SMU_wrapped = str_wrap(SMU, WRAP_WIDTH),
    n_lines     = str_count(SMU_wrapped, "\n") + 1,
    y_label     = row_id + (n_lines - 1)/2 * LINE_HEIGHT
  )

# --- Compute label width per facet (in mm), then convert to "tile units" if you want ---
# Open a temporary viewport so convertUnit has context
pushViewport(viewport())
label_widths <- smu_labels %>%
  group_by(smu_area, smu_species) %>%
  summarise(
    # Max width (mm) for the wrapped label in this facet
    LABEL_MM = max(convertUnit(stringWidth(SMU_wrapped), "mm", valueOnly = TRUE)),
    .groups = "drop"
  )
popViewport()

# If you want to stay purely in "tile units", just treat 1 tile = LABEL_MM per tile ratio.
# A quick pragmatic approach: assume 1 tile ~ 20 mm (tune as needed).
MM_PER_TILE <- 20  # tweak to taste based on your device size
label_widths <- label_widths %>%
  mutate(LABEL_WIDTH = LABEL_MM / MM_PER_TILE + LABEL_PAD)

# Join widths and compute rects
tp_plot <- tp_plot %>%
  left_join(label_widths, by = c("smu_area", "smu_species")) %>%
  mutate(
    xmax = LABEL_WIDTH + TOTAL_WIDTH - (seg_id - 1) * seg_width,
    xmin = xmax - seg_width
  )

smu_labels <- smu_labels %>%
  left_join(label_widths, by = c("smu_area", "smu_species")) %>%
  mutate(x_label = LABEL_WIDTH - LABEL_PAD)

# Plot (unchanged)
p <- ggplot(tp_plot) +
  geom_rect(aes(xmin = xmin, xmax = xmax,
                ymin = row_id - 0.45, ymax = row_id + 0.45,
                fill = Outlook),
            color = NA) +
  geom_text(data = smu_labels,
            aes(x = x_label, y = y_label, label = SMU_wrapped),
            inherit.aes = FALSE, hjust = 1, size = 2.6) +
  facet_grid(smu_species ~ smu_area, scales = "free_y", space = "free_y") +
  scale_y_reverse() +
  scale_x_continuous(
    limits = c(0, max(label_widths$LABEL_WIDTH) + TOTAL_WIDTH),
    expand = c(0, 0)
  ) +
  scale_fill_manual(values = c(
    "1"              = "#d73027",
    "1 to 2"         = "#f46d43",
    "2"              = "#fdae61",
    "2 to 3"         = "#fee08b",
    "3"              = "#a6d96a",
    "3 to 4"         = "#1a9850",
    "Data Deficient" = "grey70"
  ), name = "Outlook") +
  labs(x = NULL, y = NULL) +
  theme_minimal(base_size = 11) +
  theme(
    panel.grid       = element_blank(),
    axis.text        = element_blank(),
    axis.ticks       = element_blank(),
    strip.background = element_rect(fill = "grey80", color = "grey40"),
    strip.text.x     = element_text(face = "bold"),
    strip.text.y     = element_text(face = "bold"),
    legend.position  = "bottom",
    panel.spacing.x  = unit(1, "lines"),
    panel.spacing.y  = unit(0.8, "lines")
  )

p

