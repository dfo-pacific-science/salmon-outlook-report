---
editor_options: 
  markdown: 
    wrap: 72
---

# Salmon Outlook Report

**Note:** This repository is a work in progress. Instructions and code
may change as development continues.

## Overview

The Salmon Outlook is an annual process that provides categorical and
numeric estimates of salmon abundance for the upcoming year by Stock
Management Unit (SMU) and/or Conservation Unit (CU). These estimates
help inform harvest planning decisions.

This repository contains R code and supporting files as part of a
modernized data flow for Salmon Outlooks, which includes:

-   Streamlining data collection (e.g., using surveys)
-   Improving final products (e.g.,reorganized tables based on new
    Crosswalk data, more visually engaging presentations)
-   Automating parts of the process to reduce manual work

------------------------------------------------------------------------

## Features

This repository includes code that supports the following workflows:

1.  Automated report generation using R Markdown, following the
    technical report format provided by the `csasdown` R package. The
    code is used to produce a finalized report for submission to the DFO
    Library. See the [2026 Preliminary Outlook
    Report](https://waves-vagues.dfo-mpo.gc.ca/library-bibliotheque/41316605.pdf)
    for an example of the final output.
2.  Semi‑automated creation of presentation materials for the annual
    Preliminary Outlook Presentation held each January. The code
    populates an existing PowerPoint template with updated tables to
    reduce manual copying and pasting. See the [2026 Preliminary Outlook
    Presentation](https://github.com/dfo-pacific-science/salmon-outlook-report/blob/main/PowerPoints/2026_PreliminaryPresentation_Reference.pptx)
    for the final output.

------------------------------------------------------------------------

## Data

-   Input data comes from Survey123 and is stored in the `data/` folder
    (currently `09Jan2026Data.xlsx`).
-   For scripts to run correctly, the Excel file should include three
    sheets:
    -   `Salmon_Outlook_Report`: Outlooks + related info for SMUs
    -   `cu_outlook_records`: Outlooks + related info for CUs
    -   `other_records`: Outlooks + related info for hatchery &
        indicator stocks
-   `lookupTables.xlsx` includes the CU codes and full names, which get
    added to the data file (currently only includes CU codes).
-   `outlookClasses.xlsx` which contains the tabular data to generate
    Table 1 in the Outlook Report (see p.5 in the [2026 Preliminary
    Outlook
    Report](https://waves-vagues.dfo-mpo.gc.ca/library-bibliotheque/41316605.pdf))

------------------------------------------------------------------------

## R

This folder contains the R code for reading in, manipulating the Excel survey data.

-   dataPreprocessing.R:
    -   Reads in raw outlook data. Contains 3 sheets (SMU-level info, CU, and Hatchery/Indicator)
    -   Completes error checks to ensure e.g., no duplicate submissions, missing data, Outlooks assigned to SMU and CU, etc.
    -   Assigns resolution category (e.g., SMU, CU (aggregate), CU (singular) according to data).
    -   Final output is 1 data frame. Gets read into table making script for Outlook Report, also used for PowerPoint creation tables.
-   createTablesForReport.R
    -   reads in data frame from above, creates “flextables” for each area/species. Flextables are fancy R tables used for reports. These has all the formatting, like grey headers, bold font, etc. in a make_table() function
-   outlookGrid.R:
    -   Creates outlook summary image for PowerPoint and final report. Outlooks shown as squares, either for entire SMU or individual CUs, coloured and numbered
        to show summary for the SMU. Shown in a grid for each DFO area (columns) and Pacific salmon species (rows).
    -   See page (see p.52 in the [2026 Preliminary
    Outlook
    Report](https://waves-vagues.dfo-mpo.gc.ca/library-bibliotheque/41316605.pdf)) for reference
-   powerPointTables.R: 
    -   Adds tables to existing PowerPoint template to create Preliminary Outlook Presentation
    

------------------------------------------------------------------------

## figs

This folder contains the png files for creating Figures 1 and 2 in the [2026 Preliminary
    Outlook Report](https://waves-vagues.dfo-mpo.gc.ca/library-bibliotheque/41316605.pdf))
An additional placeholder map (as png) is included to populate the PowerPoint presentation; needs to be updated with actual maps when they are ready.


------------------------------------------------------------------------

## PowerPoints

Contains PowerPoint files for Preliminary Outlook, including:

-   2026_PreliminaryPresentation_Reference.pptx: the final presentation created, after manual editing. Shown for reference
-   2026PreliminaryOutlookTemplate.pptx: the template file, to which new tables/slides get added with Outlook data
-   updated_presentation.pptx: the file that is generated from powerPointTables.R. Will get replaced as this file is re-run.


------------------------------------------------------------------------

## csl

This includes the Citation Style Language (CSL) for the Report. For example, could specify APA, Chicago, MLA format. 
CSAS seems to follow its own citation style - always a bit unclear what it uses. Recommend not to change, unless new file comes out, or you're really invested in referencing. 
DFO Library does not seem to carefully check citations.

See [https://citationstyles.org](https://citationstyles.org/) if you would like to learn more.

------------------------------------------------------------------------

## bib

Contains `refs.bib` file, which contains the citations used in Report. Note that only three citations are in final outlook report.
Several others included for reference. See [link] for how using .bib files and how to add new ones


------------------------------------------------------------------------
## _book

After knitting the Report, this is where your tech report Word Document file will be generated. 
As `techreport.docx`


------------------------------------------------------------------------
## .Rmd files

Contains the files in Markdown format that generate each of the sections in the [2026 Preliminary Outlook
    Report](https://waves-vagues.dfo-mpo.gc.ca/library-bibliotheque/41316605.pdf) including Introduction, Results, Bibliography, and Appendix.

Most notably, the 02_results.Rmd is where the R code is referenced. Generates the code from dataPreprocessing.R which feeds into createTablesForReport.R to create the flextable objects used for making figures.

The order of these are specified in _bookdown.yml. Note that in many R Markdown documents, yml is where you specify styling requirements. However, csasdown hides many of these stylistic components for you, and can be a bit difficult to figure out.

------------------------------------------------------------------------
## .gitignore

Use this file if there is anything you do not want to appear on the public github repository (e.g., could add the data file, if do not want people to see it.) 
Internet has good documentation, e.g., see [Git Documentation](https://git-scm.com/docs/gitignore) for more information.


## How to Use

### Generate the Report

1.  Add the updated data file from survey123 to `data/` folder.
2.  Update the file path in in the dummyPath variable in `dataPreprocessing.R`.
3.  Open `Index.Rmd` in RStudio.
    -   Click **Knit** to create the report. The document will take ~30 seconds to render.
    -   When done, will create techreport.docx in `_book` folder.

Note that maps for each SMU are created as a separate process in ArcGIS. When these are ready, 
will need to manually add them to the report. 
Report creates auto-generated table captions. Insert maps above each table caption.


### Create the PowerPoint

1.  Ensure the new data file path was updated in `dataPreprocessing.R`
2.  Open `powerPointTables.R` in the `data/` folder.
2.  Run the script to:
    -   Add tables and maps from the data file
    -   Insert them into the template PowerPoint
        (`draftPrelimPres.pptx`)
    -   Save the new file as `updated_presentation.pptx`


### Create the Outlook Grid
1.  Open the outlookGrid.R file and run it.
    It is too finicky to get this to generate at correct siez, depending on your plotting window.
    Easiest to just create this separately, and add it in.
2.  Manually add in to the Report and PowerPoint when ready.



------------------------------------------------------------------------

## Additional Notes

-   You will need R, RStudio, Git and GitHub to connect to this repository and make future additions.
    Various helpful online resources to do so. 
- This site very common guide for getting started. (https://happygitwithr.com)]
-   If you’re unfamiliar with
The [UC Davis Git/GitHub
    Guide](https://ucd-r-davis.github.io/R-DAVIS/setting_up_git.html) also contains helpful advice for connecting RStudio to GitHub.
    -
-   If you are new to using csasdown, it is strongly recommended you view and follow
their initial setup instructions in the
    [csasdown guide](https://github.com/pbs-assess/csasdown/tree/main). Generate the report
    and very helpful comments throughout each of the files. If you follow their steps from a blank R project, it will generate files that I have set up already. But with sample demo, and more background on how everything fits together. E.g., tells you how to make a table, how to reference things. V helpful!!!
At times, very specialized, and may not follow regular logic of R Markdown.Very helpufl!
------------------------------------------------------------------------
