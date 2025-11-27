# Product Requirements Document (PRD)

## 1. Goal

Create a Python-based tool to perform automated quality control on Excel files submitted by institutes. The tool validates data against a flexible set of rules, stores results in a denormalized MSSQL table, generates per-institute Excel reports, and supports a Power BI dashboard.

## 2. Users

* Data owner performing quality checks.
* Analysts reviewing aggregated results.
* Power BI consumers.

## 3. Inputs

* Excel file for each institute.
* Institute identifier.
* Directory or list of downloaded files.
* Berichtsjahr (reporting year) - configurable parameter to filter documents for a specific year (default: 2026).

## 4. Core Requirements

### 4.1 Rule System

* Rules implemented as standalone Python functions.
* Each rule receives a workbook object.
* Each rule returns:

  * Boolean outcome
  * Numeric value (optional)
  * Description text
* Rules stored in a registry list for auto-discovery.
* Easy to add new rules without modifying core workflow.

### 4.2 Processing Logic

* Iterate over all provided Excel files.
* Compute hash of each file.
* Check MSSQL table for existing hash.
* If hash exists:

  * Skip unless force mode is active.
* If hash is new:

  * Assign next version number for that institute.
* Run all registered rules.
* Store results.

### 4.3 Versioning

* Version is per institute.
* Version increments only for new unique hashes.
* Hash stored as metadata.

## 5. Outputs

### 5.1 Database

Denormalized table **qc_results** with columns:

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

### 5.2 Excel per Institute

Generated based on fixed template. Columns:

* Description
* Result (bool)
* Value (numeric)
* Assessment (free text)
* Resolved (yes/no)
  Include metadata fields inside the file.
  Checks listed in fixed order.

## 6. Integration with Power BI

* Power BI connects to qc_results.
* Dashboard shows aggregate results per institute:

  * All good
  * Number of issues
  * Trends across versions

## 7. Technical Design

### 7.1 Components

* `reader.py`: load Excel files.
* `versioning.py`: compute hash, assign versions.
* `rules.py`: rule definitions and registry.
* `processor.py`: orchestrate processing.
* `db.py`: MSSQL connector and insert logic.
* `report.py`: create Excel outputs.
* `main.py`: CLI or main entry point.

### 7.2 Excel Handling

* Using openpyxl.
* Template stored in `templates/`.

### 7.3 Document Sourcing

* ORSADocumentSourcer handles automatic document retrieval from FINMA database.
* Filtering by Berichtsjahr (reporting year) to select relevant documents.
* Default Berichtsjahr is 2026, configurable via command-line or API.
* GeschaeftsNr provides unique identification per institute and year.

## 8. Processing Modes

### Standard

* Skip files with known hashes.

### Forced

* Reprocess regardless of hash.

## 9. Performance Requirements

* Able to process dozens of files in one run.
* Rules must run within seconds each.

## 10. Error Handling

* Log invalid Excel files.
* Log rule failures but continue with others.

## 11. Future Extensions

* More advanced rule definitions.
* Additional metadata exports.
* Notification system for severe issues.
