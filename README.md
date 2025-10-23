# Garments QC Workflow Automation System (Python)

## Overview

This Python-based system automates the entire Quality Control (QC) data processing workflow for the garments/textile industry. It was developed to eliminate hours of manual work involved in processing individual inspection reports, consolidating data, filtering critical issues, preparing email summaries, and archiving files.

## Features

* **Automated Data Extraction & Consolidation:** Reads data from multiple Excel inspection reports (`.xlsx`, `.xlsm`).
* **Intelligent Data Processing:** Extracts over 30 data points, calculates defect points, and validates data based on configurable rules.
* **Master File Update:** Populates a central master Excel workbook (`.xlsm`) with processed and sorted data using libraries like `xlwings` to preserve existing formats and features.
* **Smart Report Filtering:** Analyzes reports based on rules defined in `master.json` (e.g., FAIL status, width shortage, avg. defect points, shading %) to identify reports requiring management attention.
* **Automated Email Drafting:** Groups critical reports by buyer/supplier and generates ready-to-send draft emails in Microsoft Outlook with formatted summaries and relevant reports attached using `win32com`.
* **Automated File Organization:** Renames processed reports to a standard format and archives them into a structured folder system (`Buyer/Consignment_Date`).
* **User-Friendly Menu:** Provides a simple command-line interface to run the full process or individual tasks (Data Entry, Emailing, Organizing).
* **Configurable:** All paths, mappings, email recipients, and business rules are managed externally in `master.json` for easy updates without code changes.
* **Logging:** Records all actions and errors in `automation_log.txt` for easy troubleshooting.

## Technologies Used

* **Python 3**
* **Pandas:** For efficient data reading and manipulation.
* **Openpyxl:** For reading data from Excel files quickly.
* **Xlwings:** For writing data to the master `.xlsm` file while preserving macros, tables, and shapes.
* **pywin32 (win32com):** For interacting with Microsoft Outlook to create email drafts.
* **JSON:** For configuration management.
* **Standard Libraries:** `os`, `sys`, `logging`, `re`, `shutil`, `datetime`, `pathlib`.

## How it Works (Conceptual Flow)

1.  Place raw Excel inspection reports in the "Pending Reports" folder.
2.  Run the Python script and choose an option from the menu.
3.  The script processes files, updates the master workbook, creates email drafts in Outlook for critical reports, and moves processed files to the "Ongoing Work" archive, organized by buyer and date.
4.  Standard PASS reports are moved to a "Manual Review" folder.

*(This project demonstrates skills in process automation, data manipulation, file management, and integrating Python with Microsoft Office applications within a specific industry context.)*
