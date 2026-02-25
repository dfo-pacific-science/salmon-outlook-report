# Salmon Outlook Report

**Note:** This repository is a work in progress. Instructions and code may change as development continues.

------------------------------------------------------------------------

## Overview

The Salmon Outlook is an annual process that provides categorical and numeric estimates of salmon abundance for the upcoming year by Stock Management Unit (SMU) and/or Conservation Unit (CU). These estimates inform harvest planning decisions.

This repository contains R code and supporting files used in a modernized data flow for Salmon Outlook to create reproducible work flow, and automate parts of the process.
For example, in 2025-2026 undertook steps to:

-   Streamline data collection (e.g., using Survey123)
-   Improve final products (e.g., reorganized tables based on updated Crosswalk data, more consistent formatting)
-   to reduce manual work

This GitHub repository includes code to undertake those steps.

------------------------------------------------------------------------

## Features

This repository supports the following workflows:

1.  **Automated report generation** using R Markdown and `csasdown`.\
    The output is a Word document suitable for submission to the DFO Library.

    See the 2026 Preliminary Outlook Report for an example:\
    <https://waves-vagues.dfo-mpo.gc.ca/library-bibliotheque/41316605.pdf>

    **Note** that this report was based on the `csasdown` R package, specifically the styling for Technical Report. 
    While the Salmon Outlook is not a true CSAS technical report, this format is the closest available template. 

2.  **Semi-automated PowerPoint creation** for the annual Preliminary Outlook Presentation (January). Tables are programmatically inserted into an existing template to reduce manual copying and pasting.

    See the 2026 reference presentation for example of final output:\
    <https://github.com/dfo-pacific-science/salmon-outlook-report/blob/main/PowerPoints/2026_PreliminaryPresentation_Reference.pptx>

------------------------------------------------------------------------

## Data

Input data comes from Survey123 and is stored in the `data/` folder (e.g., `09Jan2026Data.xlsx`).

The Excel file must contain three sheets:

-   `Salmon_Outlook_Report` — SMU-level outlooks and related information
-   `cu_outlook_records` — CU-level outlooks
-   `other_records` — Hatchery and indicator stock outlooks

Additional supporting files:

-   `lookupTables.xlsx`\
    Contains CU codes and full names, which get added to the tables.

-   `outlookClasses.xlsx`\
    Contains tabular data used to generate Table 1 in the Outlook Report (see p.5 of the 2026 report).

------------------------------------------------------------------------

## R Scripts (in R folder)

These scripts handle data validation, formatting, and output creation.
Comments describing the overall process and steps within code are provided in each script.

### `dataPreprocessing.R`

-   Reads in raw Outlook data (all three sheets) from survey123.
-   Performs error checks (e.g., duplicate submissions, missing fields, mismatched SMU/CU assignments).
-   Assigns resolution category (SMU, CU aggregate, or CU singular).
-   Outputs a single cleaned data frame.

This data frame feeds into both the report tables and PowerPoint tables.

------------------------------------------------------------------------

### `createTablesForReport.R`

-   Reads in the cleaned data frame.
-   Creates formatted `flextable` objects for each area/species.
-   Applies consistent formatting (headers, bolding, layout).

These tables are used in the final report.

------------------------------------------------------------------------

### `outlookGrid.R`

-   Creates the Outlook summary grid used in the report and presentation.
-   Displays Outlook categories as coloured, numbered squares.
-   Organized by DFO Area (columns) and Pacific salmon species (rows).

See p.52 of the 2026 report for an example:\
<https://waves-vagues.dfo-mpo.gc.ca/library-bibliotheque/41316605.pdf>

While it could be possible to automatically add these to the Report and Powerpoint, Plot sizing can be finicky depending on your R Studio plotting window.It also adds rendering time. 
Instead, it is easiest to generate this separately and insert the images manually into the report and PowerPoint.

------------------------------------------------------------------------

### `powerPointTables.R`

-   Reads in the cleaned data.
-   Inserts formatted tables into an existing PowerPoint template.
-   Saves an updated presentation file.

------------------------------------------------------------------------

## figs

Contains PNG files used to generate Figures 1 and 2 in the Outlook Report.

Also includes a placeholder map image used in the PowerPoint. Replace this with finalized maps when available.

------------------------------------------------------------------------

## PowerPoints

-   `2026_PreliminaryPresentation_Reference.pptx`\
    Final presentation after manual edits (for reference only).

-   `2026PreliminaryOutlookTemplate.pptx`\
    Template file used by the automation script.

-   `updated_presentation.pptx`\
    Generated file created by `powerPointTables.R`.\
    Overwritten each time the script is run.

------------------------------------------------------------------------

## csl

Contains the Citation Style Language (CSL) file used by the report.

CSAS uses its own citation style. It is recommended not to modify this file unless an updated version is officially released.

More information on CSL files:\
<https://citationstyles.org/>

------------------------------------------------------------------------

## bib

Contains `refs.bib`, which stores references used in the report.

Only a few citations appear in the final Outlook Report. Additional references from the csasdown generation were included as reference..

Note that most citation softwares will generate .bib files. If you are unfamiliar with these files, see for example, https://www.ub.uzh.ch/en/unterstuetzung-erhalten/tutorials/publizieren/How-to-Cite-in-LaTeX.html

------------------------------------------------------------------------

## \_book

After knitting the report, the technical report Word document is generated here as:

`techreport.docx`

------------------------------------------------------------------------

## R Markdown Files

Contains the `.Rmd` files that generate each section of the report (Introduction, Results, Appendix, Bibliography).

The main file is `Index.Rmd`.

`02_results.Rmd` is where most R code is executed. It sources:

-   `dataPreprocessing.R`
-   `createTablesForReport.R`

The order of sections is controlled by `_bookdown.yml`.

------------------------------------------------------------------------

## How to Use

### Initial Setup (Important)

If this is your first time using `csasdown`, complete the setup instructions here:

<https://github.com/pbs-assess/csasdown/tree/main>

It is strongly recommended to do a practice run using their example template before working with this repository.
Follow their instructions online (e.g., using the `csasdown::draft("techreport")` command, which will create similar files.
e.g., read through their Rmd files that are generated - helpful guidance for how to structure report, reference citations, etc.

------------------------------------------------------------------------

### Generate the Report

1.  Add the updated Survey123 Excel file to the `data/` folder.
2.  Update the file path in `dataPreprocessing.R` (the `dummyPath` variable).
3.  Open `Index.Rmd` in RStudio.
4.  Click **Knit**.

Rendering takes \~30 seconds. The output (`techreport.docx`) will appear in the `_book` folder.

**Important:**\
Because the Outlook is not a true CSAS technical report, manual edits are required after knitting. For example:

-   Removing the automatically-generated title pages, and replacing them with what has been used in previous years.
-   Adding the citation page as the last page to the report.
-   Inserting finalized maps (created separately in ArcGIS)

Maps should be inserted above the relevant table captions.

------------------------------------------------------------------------

### Create the PowerPoint

1.  Confirm the data file path is updated in `dataPreprocessing.R`.
2.  Open `powerPointTables.R`.
3.  Run the script.

This will: - Add tables (and placeholder maps) to the template - Save the output as `updated_presentation.pptx`

Manual formatting adjustments may still be required.

------------------------------------------------------------------------

### Create the Outlook Grid

1.  Open `outlookGrid.R`.
2.  Run the script.
3.  Insert the generated image manually into the report and presentation.

------------------------------------------------------------------------

## Git Setup

If unfamiliar with connecting RStudio to GitHub, see:

<https://ucd-r-davis.github.io/R-DAVIS/setting_up_git.html>

Additional R/Git/GitHub guidance:

<https://happygitwithr.com/>
