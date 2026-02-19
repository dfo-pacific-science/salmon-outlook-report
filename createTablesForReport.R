

source("dataPreprocessing.R")


################################################################################
### Table building and styling functions

# Each SMU gets its own “block” consisting of:
#  1) a grey SMU label row
#  2) the narrative text
#  3) a column header row
#  4) the actual data rows


build_block = function(df_smu) {
  smu  = unique(df_smu$smu_name)
  narr = unique(df_smu$Narrative)

  # Row that just labels the start of a new SMU block
  row1 = data.frame(
    Resolution = "SMU",
    Name = "Narrative",
    Forecast = "",
    Outlook  = "",
    stringsAsFactors = FALSE
  )

  # Row that actually contains the narrative text
  row2 = data.frame(
    Resolution = smu,
    Name = narr,
    Forecast = "",
    Outlook  = "",
    stringsAsFactors = FALSE
  )

  # Row with next set of column names
  row3 = data.frame(
    Resolution = "Resolution",
    Name = "Name",
    Forecast = "Forecast",
    Outlook  = "Outlook",
    stringsAsFactors = FALSE
  )

  # Actual SMU / CU data rows
  data_rows = df_smu %>%
    select(Resolution, Name, Forecast, Outlook)

  bind_rows(row1, row2, row3, data_rows)
}

################################################################################
# This function defines the visual styling of the SMU tables.
# It controls borders, shading, merging, and layout.
# The input (big_df) is already structured in blocks by SMU.
#
# Note: A lot of this ordering matters. Flextable can behave strangely
# if you merge cells before applying borders, so the sequence below is deliberate.

################################################################################

style_smu_table = function(big_df) {

  # Identify rows where a new SMU block begins.
  # These are the rows where:
  #   - Resolution == "SMU"
  #   - Name == "Narrative"
  # These act like section headers inside the body of the table.
  idx_label  = which(big_df$Resolution == "SMU" & big_df$Name == "Narrative")

  # Identify the row that acts as an internal column header.
  # In this table design, we are not using the native flextable header.
  # Instead, one row in the body acts as the header row.
  idx_header = which(big_df$Resolution == "Resolution")

  # We want a thicker line *before* each new SMU block,
  # but not before the very first one.
  # So we take all SMU start rows except the first,
  # and subtract 1 to place the border on the row above.
  idx_separators = idx_label[-1] - 1

  # Create the flextable object from the prepared data frame.
  ft = flextable(big_df)

  # We intentionally delete the default header.
  # This gives us full control over formatting,
  # since we are treating one body row as the header instead.
  ft = delete_part(ft, part = "header")

  # Start from a clean slate: remove all default borders.
  # Add everything back manually for consistency.
  ft = border_remove(ft)

  # Add thick horizontal lines to visually separate SMU blocks.
  # These create clear section breaks between SMUs.
  if (length(idx_separators) > 0) {
    ft = border(
      ft,
      i = idx_separators,
      j = 1:ncol(big_df),
      border.bottom = fp_border(width = 2),
      part = "body"
    )
  }

  # Add standard outer border around the whole table.
  ft = border_outer(ft, fp_border(width = 1))

  # Add vertical internal borders between columns.
  ft = border_inner_v(ft, fp_border(width = 1))

  # Now add thin horizontal borders to all other rows.
  # We exclude the rows that already received thick separators.
  all_rows  = seq_len(nrow(big_df))
  thin_rows = setdiff(all_rows, idx_separators)

  ft = border(
    ft,
    i = thin_rows,
    j = 1:ncol(big_df),
    border.bottom = fp_border(width = 1),
    part = "body"
  )

  # IMPORTANT: Merge cells only AFTER borders are finalized.
  # If you merge first, borders can disappear or behave unpredictably.
  for (i in idx_label) {
    ft = merge_at(ft, i = i,     j = 2:4)
    ft = merge_at(ft, i = i + 1, j = 2:4)
  }

  # Apply shading and bold formatting to SMU section header rows.
  ft = bg(ft,   i = idx_label,  bg = "gray90")
  ft = bold(ft, i = idx_label,  bold = TRUE)

  # Apply same styling to the internal column header row.
  ft = bg(ft,   i = idx_header, bg = "gray90")
  ft = bold(ft, i = idx_header, bold = TRUE)

  # Automatically adjust column widths to fit content.
  ft = autofit(ft)

  # Ensure layout adapts nicely in Word/HTML output.
  ft = set_table_properties(ft, layout = "autofit")

  ft
}


################################################################################
# This function builds the caption text.
# It:
#   - Normalizes area name formatting
#   - Looks up the scientific name for the species
#   - Returns a single character string
#
# Note: Scientific names are NOT italicized here.
# Italic styling is handled later at the flextable level so we can
# selectively turn italics on/off for specific parts.
################################################################################

################################################################################
# Main wrapper function.
# This:
#   1. Filters the data to a specific area + species
#   2. Builds SMU blocks
#   3. Styles the table
#   4. Applies a caption with controlled italics
################################################################################

make_table = function(area, species, data = tabPrep) {

  # Filter to requested area and species.
  # Assumes smu_area and smu_species already match the provided strings.
  df_filtered = data %>%
    filter(smu_area == area, smu_species == species)

  # Defensive check: stop if no rows found.
  if (nrow(df_filtered) == 0) {
    stop(paste("No data found for:", area, "/", species))
  }

  # Split data into a list by SMU name.
  # Each SMU becomes its own block.
  smu_list   = split(df_filtered, df_filtered$smu_name)

  # Build formatted blocks using external helper function.
  block_list = lapply(smu_list, build_block)

  # Combine blocks back into a single data frame.
  big_df     = bind_rows(block_list)

  # Apply table styling.
  ft = style_smu_table(big_df)

  # Build full caption string.
  caption_text = make_caption(species, area)

  # Recreate scientific name lookup (kept local for clarity).
  sci_lookup = c(
    chinook = "Oncorhynchus tshawytscha",
    coho    = "Oncorhynchus kisutch",
    sockeye = "Oncorhynchus nerka",
    chum    = "Oncorhynchus keta",
    pink    = "Oncorhynchus gorbuscha"
  )

  sp_key = tolower(trimws(species))
  sci = sci_lookup[[sp_key]]

  # Split caption into three parts:
  #   1. Text before scientific name
  #   2. Scientific name
  #   3. Text after scientific name
  #
  # This allows us to control italics at a granular level.
  parts = strsplit(caption_text, sci)[[1]]

  # Apply caption with custom formatting:
  # Entire caption italicized EXCEPT scientific name.
  ft = set_caption(
    ft,
    caption = as_paragraph(
      as_chunk(parts[1], props = fp_text_default(italic = TRUE)),
      as_chunk(sci,      props = fp_text_default(italic = FALSE)),
      as_chunk(parts[2], props = fp_text_default(italic = TRUE))
    )
  )

  ft
}

################################################################################
# Example usage (unchanged)
# make_table("FRASER AND INTERIOR", "Chinook")
# make_table("SOUTH COAST", "Chinook")

# End of script



