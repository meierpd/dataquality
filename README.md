# Data Quality Control Tool

## Overview

This project provides an automated way to validate Excel files submitted by institutes. Each file is parsed and run through a collection of modular checks. The results are stored in an MSSQL table and later used to generate per-institute Excel reports and a Power BI dashboard.

## Features

* Process multiple Excel files with institute identifiers.
* Automatic versioning based on file hashes.
* Flexible rule system implemented as Python functions.
* One row per check result stored in the database.
* Regenerable outputs, including a forced reprocessing option.
* Fixed-format Excel report generation per institute.
* Power BI-ready denormalized result table.

## Architecture

1. **Ingestion**: Files are downloaded or collected, each paired with an institute ID.
2. **Versioning**: Each file is hashed. A new version is created whenever a new hash appears for the same institute.
3. **Rule Engine**: Checks are simple Python functions registered in a list. Each receives a workbook and returns:

   * Boolean outcome
   * Optional numeric value
   * Description string
4. **Database Writer**: Each result is written as a separate row, including metadata.
5. **Excel Report Generator**: A formatted workbook is created for each institute showing results plus empty fields for manual assessment.

## Directory Structure

```
project/
  checks/
    __init__.py
    rules.py
  core/
    processor.py
    reader.py
    versioning.py
    db.py
    report.py
  templates/
    institute_report_template.xlsx
  main.py
  README.md
```

## Adding a New Check

To add a check, open `checks/rules.py` and define a new function:

```python
from openpyxl import Workbook

def check_example(wb: Workbook):
    desc = "Example description"
    outcome = True
    value = None
    return outcome, value, desc
```

Then add the function to the `REGISTERED_CHECKS` list in the same file.

## Database Schema

Single denormalized table containing all results:

**Table: qc_results**

* id (PK)
* institute_id
* file_name
* file_hash
* version_number
* check_name
* check_description
* outcome_bool
* outcome_numeric
* processed_at (timestamp)

## Processing Workflow

1. Load all incoming Excel files.
2. Compute hash for each.
3. Check the database for existing hashes.
4. If not processed, assign version and run all checks.
5. Store results in the database.
6. Generate or update the Excel report for that institute.

A forced reprocessing mode skips the hash check.

## Excel Report

For each institute, an Excel sheet is generated based on a template. Columns include:

* Description
* Check result
* Assessment (free text)
* Resolved (yes/no)

## Power BI

Power BI connects directly to `qc_results`, which already contains all metadata required for aggregated views.
