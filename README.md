# Data Quality Control Tool

## Overview

This project provides an automated way to validate Excel files submitted by institutes. Each file is parsed and run through a collection of modular checks. The results are stored in an MSSQL table and later used to generate per-institute Excel reports and a Power BI dashboard.

## Features

* **File Processing & Caching**: SHA-256 hash-based caching to avoid reprocessing identical files
* **Automatic Versioning**: Version assignment per institute based on file content changes
* **Modular Check System**: Extensible Python-based quality checks with simple registration
* **Database Output**: Denormalized qc_results table ready for MSSQL storage
* **Force Reprocess Mode**: Option to reprocess files regardless of cache status
* **ORSADocumentSourcer Integration**: Direct integration with document sourcing system
* **Comprehensive Testing**: Full unit test coverage for all modules

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
orsa_analysis/
  src/
    orsa_analysis/
      __init__.py
      cli.py              # Command-line interface
      core/               # Core processing logic
        __init__.py
        processor.py      # Main document processor
        reader.py         # Excel file reader
        versioning.py     # File versioning & caching
        db.py            # Database writers
        report.py        # Report generation
      checks/            # Quality check rules
        __init__.py
        rules.py          # Check implementations
      sourcing/          # Document sourcing
        __init__.py
        document_sourcer.py  # ORSADocumentSourcer
  templates/
    institute_report_template.xlsx
  sql/
    source_orsa_dokument_metadata.sql
  tests/              # Comprehensive unit tests
  pyproject.toml      # Modern Python packaging
  README.md
```

## Installation

```bash
# Install in editable mode for development
pip install -e .

# Or install from repository
pip install git+https://github.com/meierpd/dataquality.git
```

## Quick Start

The package can be used as a library or via the command-line interface.

### Command-Line Usage

```bash
# Process Excel files with CLI
orsa-qc --institute-id INS001 --input-files file1.xlsx file2.xlsx
```

### Library Usage with ORSADocumentSourcer

```python
from orsa_analysis import DocumentProcessor
from orsa_analysis.core.db import InMemoryDatabaseWriter
from orsa_analysis.sourcing import ORSADocumentSourcer

# Initialize database writer
db_writer = InMemoryDatabaseWriter()

# Initialize processor
processor = DocumentProcessor(db_writer, force_reprocess=False)

# Source documents from database
sourcer = ORSADocumentSourcer()
documents = sourcer.load()  # Returns List[Tuple[str, Path]]

# Process all documents
for name, path in documents:
    results = processor.process_file(path, institute_id="INS001")

# Get processing summary
summary = processor.get_processing_summary()
print(f"Total files: {summary['total_files']}")
print(f"Processed: {summary['processed']}")
print(f"Cached: {summary['cached']}")
print(f"Pass rate: {summary['pass_rate']:.1%}")

# Access stored results
all_results = db_writer.get_results()
```

### ORSADocumentSourcer Setup

To use the document sourcing functionality:

1. Create a `credentials.env` file in the project root:
```bash
FINMA_USERNAME=your_username
FINMA_PASSWORD=your_password
```

2. Ensure the SQL query file exists at `sql/source_orsa_dokument_metadata.sql`

3. The sourcer will automatically:
   - Download documents from the FINMA database
   - Filter for ORSA documents from 2026 onwards
   - Store files in `orsa_response_files/` directory
   - Return document paths for processing

## Adding a New Check

To add a check, open `src/orsa_analysis/checks/rules.py` and define a new function:

```python
from openpyxl.workbook.workbook import Workbook
from typing import Optional, Tuple

def check_example(wb: Workbook) -> Tuple[bool, Optional[float], str]:
    """Your custom check logic.
    
    Args:
        wb: Workbook to check
        
    Returns:
        Tuple of (outcome, numeric_value, description)
    """
    desc = "Example description"
    outcome = True
    value = 42.0  # Optional numeric outcome
    return outcome, value, desc
```

Then add the function to the `REGISTERED_CHECKS` list in the same file:

```python
REGISTERED_CHECKS.append(("example_check", check_example))
```

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

## Testing

Run the comprehensive test suite (103 tests):

```bash
# Run all tests
pytest

# Run with verbose output
pytest -v

# Run with coverage report
pytest --cov=orsa_analysis --cov-report=html

# Run specific test modules
pytest tests/test_document_sourcer.py -v
```

All modules have full unit test coverage:
- `tests/test_reader.py` - ExcelReader tests
- `tests/test_versioning.py` - VersionManager tests
- `tests/test_db.py` - Database writer tests
- `tests/test_rules.py` - Check function tests (15 checks)
- `tests/test_processor.py` - DocumentProcessor integration tests
- `tests/test_document_sourcer.py` - ORSADocumentSourcer tests (22 tests)

## Module Documentation

### core/reader.py
`ExcelReader` class for loading Excel files with openpyxl.
- `load_file(path)` - Loads an Excel file
- `get_sheet_names(workbook)` - Gets list of sheet names
- `close_workbook(workbook)` - Closes the workbook

### core/versioning.py
`VersionManager` handles file hashing and version assignment.
- `compute_file_hash(path)` - Computes SHA-256 hash
- `get_version(institute_id, path)` - Gets or assigns version
- `is_processed(institute_id, hash)` - Checks if file was processed
- `load_existing_versions(data)` - Loads cache from database

### core/db.py
Database models and writers.
- `CheckResult` - Dataclass for check results
- `DatabaseWriter` - Abstract base class for DB operations
- `InMemoryDatabaseWriter` - In-memory implementation

### core/processor.py
`DocumentProcessor` orchestrates the processing workflow.
- `process_file(institute_id, path)` - Processes single file
- `process_documents(documents)` - Batch processing
- `should_process_file(institute_id, path)` - Cache check
- `get_processing_summary()` - Get statistics

### checks/rules.py
Quality check registry with pre-built checks:
- `check_has_sheets` - Verify workbook has sheets
- `check_no_empty_sheets` - Verify no empty sheets
- `check_first_sheet_has_data` - Verify A1 has data
- `check_sheet_names_unique` - Verify unique sheet names
- `check_row_count_reasonable` - Verify row count limits
- `check_has_expected_headers` - Verify headers exist
- `check_no_merged_cells` - Detect merged cells

## Integration Notes

### ORSADocumentSourcer Output Format
The processor expects `List[Tuple[str, Path]]` from `sourcer.load()`:
```python
[
    ("INST001_report.xlsx", Path("/path/to/file1.xlsx")),
    ("INST002_report.xlsx", Path("/path/to/file2.xlsx")),
]
```

### Institute ID Extraction
By default, extracts ID from filename before first `_`, `-`, or space.
Override `DocumentProcessor._extract_institute_id()` for custom logic.

### Custom Database Writer
Implement `DatabaseWriter` for your database backend:

```python
from dataquality.core.db import DatabaseWriter

class SQLDatabaseWriter(DatabaseWriter):
    def write_results(self, results):
        # Write to SQL database
        return len(results)
    
    def get_existing_versions(self):
        # Query existing versions
        return query_results
```
