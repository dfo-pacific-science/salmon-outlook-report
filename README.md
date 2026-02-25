# Salmon Outlook Report

**Note:** This repository is a work in progress. Instructions and code may change as development continues.

------------------------------------------------------------------------

## Overview

The Salmon Outlook is an annual process that provides categorical and
numeric estimates of salmon abundance for the upcoming year by Stock
Management Unit (SMU) and/or Conservation Unit (CU). These estimates
inform harvest planning decisions. As part of this process, various presentations and 
reports are created, many of which previously required substantial manual editing, emailing, and copying and pasting.


In 2025–2026 several steps were taken to modernize this workflow, including:

- Streamlining data collection (e.g., using Survey123)
- Improving final products (e.g., reorganized tables based on [updated SMU-CU-DU
  Crosswalk data](https://github.com/dfo-pacific-science/Salmon-SMU-CU-DU-Crosswalk), more consistent formatting)
- Reducing manual work

This repository contains R code and supporting files used in this improved data workflow. 
The code helps organize the processing steps, standardize the approach, and automate portions of the analysis.

------------------------------------------------------------------------

## Features

This repository supports the following workflows:

1. **Automated report generation** using R Markdown and `csasdown`.  
   The output is a Word document suitable for submission to the DFO
   Library.

   See the 2026 Preliminary Outlook Report for an example:  
   <https://waves-vagues.dfo-mpo.gc.ca/library-bibliotheque/41316605.pdf>

   **Note:** This report uses the `csasdown` R package, specifically the
   Technical Report styling. While the Salmon Outlook is not a true
   DFO Technical Report, this format is the closest available template to what the 
   DFO Library typically requires (e.g., spacing, font sizes, etc.).

2. **Semi-automated PowerPoint creation** for the annual Preliminary
   Outlook Presentation (January). Tables are programmatically inserted
   into an existing template to reduce manual copying and pasting.

   See the 2026 reference presentation for an example of final output:  
   <https://github.com/dfo-pacific-science/salmon-outlook-report/blob/main/PowerPoints/2026_PreliminaryPresentation_Reference.pptx>

3. **Code for one figure in the Major Stocks Infographic (PowerPoint)**.
Most of the infographic is created manually, but one map is generated using R code in this repository.
(A link to the final PowerPoint will be added when available.)

------------------------------------------------------------------------

## Data

Input data comes from Survey123 and is stored in the `data/` folder
(e.g., `09Jan2026Data.xlsx`).

The Excel file must contain three sheets:

- `Salmon_Outlook_Report` — SMU-level outlooks and related information
- `cu_outlook_records` — CU-level outlooks
- `other_records` — Hatchery and indicator stock outlooks

Additional supporting files:

- `lookupTables.xlsx`  
  Contains CU codes and full names, which are appended to the tables.
  These were obtained from the SMU-CU-DU Crosswalk

- `outlookClasses.xlsx`  
  Contains tabular data used to generate Table 1 in the Outlook Report
  (see p.5 of the 2026 report).

------------------------------------------------------------------------

## R Scripts (in `R/` folder)

These scripts handle data validation, formatting, and output creation.
Comments describing the overall process and steps within the code are
provided in each script.

### `dataPreprocessing.R`

- Reads in raw Outlook data (all three sheets) from Survey123.
- Performs error checks (e.g., duplicate submissions, missing fields,
  mismatched SMU/CU assignments).
- Assigns resolution category (SMU, CU aggregate, or CU singular).
- Outputs a single cleaned data frame.

This data frame feeds into both the report tables and PowerPoint tables.

------------------------------------------------------------------------

### `createTablesForReport.R`

- Reads in the cleaned data frame.
- Creates formatted `flextable` objects for each area/species.
- Applies consistent formatting (headers, bolding, layout).

These tables are used in the final report.

------------------------------------------------------------------------

### `outlookGrid.R`

- Creates the Outlook summary grid used in the report and presentation.
- Displays Outlook categories as coloured, numbered squares.
- Organized by DFO Area (columns) and Pacific salmon species (rows).

See p.52 of the 2026 report for an example:  
<https://waves-vagues.dfo-mpo.gc.ca/library-bibliotheque/41316605.pdf>

While it may be possible to automatically insert these into the report
and PowerPoint, plot sizing can be finicky depending on the RStudio
plotting window and graphics device. It also increases rendering time.
For that reason, it is easiest to generate this separately and insert
the images manually.

------------------------------------------------------------------------

### bib

Contains `refs.bib`, which stores references used in the report.
Only a few citations appear in the final Outlook Report. Additional references from the `csasdown` template were included for completeness.

Note that most citation software can generate `.bib` files automatically. Although it's unlikely 
you'll need to add more, there are various online guides that provide more information on these files, see for example:  
<https://www.ub.uzh.ch/en/unterstuetzung-erhalten/tutorials/publizieren/How-to-Cite-in-LaTeX.html>

------------------------------------------------------------------------

## _book

After knitting the report, the technical report Word document is
generated here as:

`techreport.docx`

------------------------------------------------------------------------

## R Markdown Files

Contains the `.Rmd` files that generate each section of the report
(Introduction, Results, Appendix, Bibliography).

The main file is `Index.Rmd`.

`02_results.Rmd` is where most R code is executed. It sources:

- `dataPreprocessing.R`
- `createTablesForReport.R`

The order of sections is controlled by `_bookdown.yml`.

------------------------------------------------------------------------

## How to Use

### Initial Setup (Important)

If this is your first time using `csasdown`, complete the setup
instructions here:

<https://github.com/pbs-assess/csasdown/tree/main>

It is strongly recommended to do a practice run using their example
template before working with this repository.

For example, run:

`csasdown::draft("techreport")`

This will generate a sample report structure. Review the generated
`.Rmd` files to understand report structure, citation handling, and
formatting expectations.

------------------------------------------------------------------------

### Generate the Report

1. Add the updated Survey123 Excel file to the `data/` folder.
2. Update the file path in `dataPreprocessing.R` (the `dummyPath`
   variable).
3. Open `Index.Rmd` in RStudio.
4. Click **Knit**.

Rendering takes ~30 seconds. The output (`techreport.docx`) will appear
in the `_book` folder.

**Important:**  
Some final, manual edits will be required after knitting. For example:

- Removing the automatically generated title pages and replacing them
  with the format used in previous years.
- Adding the citation page as the final page of the report.
- Inserting finalized maps (created separately in ArcGIS).

Maps should be inserted above the relevant table captions.

------------------------------------------------------------------------

### Create the PowerPoint

1. Confirm the data file path is updated in `dataPreprocessing.R`.
2. Open `powerPointTables.R`.
3. Run the script.

This will:
- Add tables (and placeholder maps) to the template.
- Save the output as `updated_presentation.pptx`.

Manual formatting adjustments may still be required, including:
- Inserting finalized maps (created separately in ArcGIS).
- Adjusting the table size/formatting. For example, while the forecasts are added
 as a column, only the numeric values are required. This is challenging to automate.
 Instead, keep only the numbers, remove unneeded text, then resize the tables as necessary.

------------------------------------------------------------------------

### Create the Outlook Grid

1. Open `outlookGrid.R`.
2. Run the script.
3. Insert the generated image manually into the report and presentation.

------------------------------------------------------------------------

## Git Setup

If unfamiliar with connecting RStudio to GitHub, see:

<https://ucd-r-davis.github.io/R-DAVIS/setting_up_git.html>

Additional R/Git/GitHub guidance:

<https://happygitwithr.com/>