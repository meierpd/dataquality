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
dataquality/
  src/
    orsa_analysis/
      __init__.py
      cli.py              # Command-line interface  
      core/               # Core processing logic
        __init__.py
        processor.py      # Document processor with versioning
        orchestrator.py   # Pipeline orchestration & caching control
        reader.py         # Excel file reader
        versioning.py     # File versioning & hash-based caching
        database_manager.py  # Simplified database manager & CheckResult
      checks/             # Quality check rules
        __init__.py
        rules.py          # Check implementations
      sourcing/           # Document sourcing
        __init__.py
        document_sourcer.py  # ORSADocumentSourcer
  data/
    orsa_response_files/  # Downloaded documents storage
  sql/
    source_orsa_dokument_metadata.sql    # Query for document metadata
    create_table_orsa_analysis_data.sql  # Database table schema
  tests/                # 109 comprehensive unit tests
  main.py               # Main entry point for production use
  pyproject.toml        # Modern Python packaging
  README.md
  PRD.md
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
# Process ORSA documents from sourcer (writes to MSSQL)
orsa-qc --verbose

# Process documents for a specific reporting year (default: 2026)
orsa-qc --berichtsjahr 2026 --verbose

# Force reprocess all files
orsa-qc --force --verbose

# Use custom credentials file
orsa-qc --credentials /path/to/credentials.env

# Combine options: process 2027 documents with verbose output
orsa-qc --berichtsjahr 2027 --verbose
```

Or run main.py directly:

```bash
python main.py --verbose

# Specify reporting year
python main.py --berichtsjahr 2026 --verbose
```

### Library Usage

**Simple Pipeline Usage:**

```python
from orsa_analysis import ORSAPipeline, DatabaseManager
from orsa_analysis.sourcing import ORSADocumentSourcer
from pathlib import Path

# Initialize pipeline with database
db_manager = DatabaseManager("mssql+pyodbc://server/db?driver=ODBC+Driver+17+for+SQL+Server")
pipeline = ORSAPipeline(db_manager, force_reprocess=False)

# Option 1: Process from document sourcer
sourcer = ORSADocumentSourcer(berichtsjahr=2026)
summary = pipeline.process_from_sourcer(sourcer)

# Option 2: Process specific files
documents = [
    ("INST001_report.xlsx", Path("data/INST001_report.xlsx")),
    ("INST002_report.xlsx", Path("data/INST002_report.xlsx")),
]
summary = pipeline.process_documents(documents)

print(f"Processed: {summary['files_processed']}, Skipped: {summary['files_skipped']}")
print(f"Institutes: {summary['institutes']}")

# Get detailed summary
full_summary = pipeline.generate_summary()
pipeline.close()
```

**Advanced Caching Control:**

```python
from orsa_analysis import CachedDocumentProcessor, DatabaseManager
from pathlib import Path

# Initialize with caching control
db_manager = DatabaseManager("mssql+pyodbc://server/db")
processor = CachedDocumentProcessor(db_manager, cache_enabled=True)

# Check cache status for a file
file_path = Path("data/INST001_report.xlsx")
status = processor.get_cache_status("INST001", file_path)
if status["is_cached"]:
    print(f"File cached with version {status['version_number']}")

# Get cache statistics
stats = processor.get_cache_statistics()
print(f"Cached: {stats['total_institutes']} institutes, {stats['total_versions']} versions")

# Invalidate cache for specific institute
processor.invalidate_cache("INST001")

# Or invalidate entire cache
processor.invalidate_cache()
```

**Direct Processor Usage (Lower Level):**

```python
from orsa_analysis import DocumentProcessor, DatabaseManager
from pathlib import Path

# Initialize components
db_manager = DatabaseManager("mssql+pyodbc://server/db")
processor = DocumentProcessor(db_manager, force_reprocess=False)

# Process single file
file_path = Path("data/INST001_report.xlsx")
version_info, results = processor.process_file("INST001", file_path)

print(f"Processed version {version_info.version_number}")
print(f"Ran {len(results)} checks")

# Close database connection
db_manager.close()
```

## Caching & Versioning

The system uses SHA-256 file hashing for intelligent caching:

- **Automatic Caching**: Files are hashed on first processing. Subsequent runs skip unchanged files.
- **Version Tracking**: Each unique file content gets a new version number per institute.
- **Force Reprocess**: Use `force_reprocess=True` to override caching and reprocess all files.
- **Cache Inspection**: `CachedDocumentProcessor` provides detailed cache introspection and control.

**How it works:**

1. File is hashed (SHA-256) on first encounter
2. Hash is checked against database of processed versions
3. If hash exists → Skip processing (unless forced)
4. If hash is new → Assign new version number and process
5. Results stored in database with version metadata

**Cache Control:**

```python
from orsa_analysis import CachedDocumentProcessor, DatabaseManager

processor = CachedDocumentProcessor(DatabaseManager("..."))

# Check if file is cached
status = processor.get_cache_status("INST001", file_path)

# View cache statistics
stats = processor.get_cache_statistics()

# Invalidate specific institute cache
processor.invalidate_cache("INST001")
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
   - Filter for ORSA documents matching the specified reporting year (Berichtsjahr)
   - Default reporting year is 2026, configurable via `--berichtsjahr` argument
   - Store files in `data/orsa_response_files/` directory
   - Return document paths for processing
   
**Note on Reporting Year:**
The `berichtsjahr` parameter allows you to specify which reporting year to process. 
While the GeschaeftsNr (business number) is unique per institute and year, filtering by 
berichtsjahr makes it easier to focus on specific reporting periods. Over time, as multiple 
years accumulate, the GeschaeftsNr remains the unique identifier.

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

## Database Integration

The system writes check results to an MSSQL database for reporting and analysis.

### Database Schema

**Database**: `GBI_REPORTING`  
**Schema**: `gbi`  
**Table**: `orsa_analysis_data`

Columns:
* id (INT, IDENTITY PK)
* institute_id (NVARCHAR(50))
* file_name (NVARCHAR(255))
* file_hash (NVARCHAR(64))
* version (INT)
* check_name (NVARCHAR(100))
* check_description (NVARCHAR(MAX))
* outcome_bool (BIT)
* outcome_numeric (FLOAT, NULL)
* processed_timestamp (DATETIME2, DEFAULT GETDATE())

Indexes:
* Clustered index on id
* Non-clustered index on institute_id
* Non-clustered index on file_hash
* Composite index on (institute_id, version)

### Database Views

Two convenience views are provided:

1. **vw_orsa_analysis_latest**: Shows only the latest version per institute
2. **vw_orsa_analysis_summary**: Aggregates pass rates by institute and check

### Setup Database

To create the database table and views:

```bash
# Run the SQL script on your MSSQL server
sqlcmd -S your_server -d GBI_REPORTING -i sql/create_table_orsa_analysis_data.sql
```

### Database Connection

The system uses environment variables for database credentials:

```bash
# Set environment variables
export DB_USER=your_username
export DB_PASSWORD=your_password
```

Alternatively, use Windows authentication (default on Windows systems).

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
The processor expects `List[Tuple[str, Path, str]]` from `sourcer.load()`:
```python
[
    ("INST001_report.xlsx", Path("/path/to/file1.xlsx"), "GNR123"),
    ("INST002_report.xlsx", Path("/path/to/file2.xlsx"), "GNR456"),
]
```

### Using ORSADocumentSourcer with Berichtsjahr

```python
from orsa_analysis.sourcing import ORSADocumentSourcer

# Process documents for 2026 (default)
sourcer = ORSADocumentSourcer()
documents = sourcer.load()

# Process documents for a different year
sourcer_2027 = ORSADocumentSourcer(berichtsjahr=2027)
documents_2027 = sourcer_2027.load()

# Use with pipeline
pipeline.process_from_sourcer(sourcer_2027)
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
