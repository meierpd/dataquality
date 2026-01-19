# Data Quality Control Tool

## Overview

This project provides an automated way to validate Excel files submitted by institutes. Each file is parsed and run through a collection of modular checks. The results are stored in an MSSQL table and later used to generate per-institute Excel reports and a Power BI dashboard.

## Features

* **File Processing & Caching**: SHA-256 hash-based caching to avoid reprocessing identical files
* **Automatic Versioning**: Version assignment per institute based on file content changes
* **Modular Check System**: Extensible Python-based quality checks with simple registration
* **Multi-Language Support**: Automatic detection and handling of German, English, and French Excel sheet names
* **Report Generation**: Automated standalone Excel report creation with check results from templates
* **Database Output**: Denormalized qc_results table ready for MSSQL storage
* **Force Reprocess Mode**: Option to reprocess files regardless of cache status
* **ORSADocumentSourcer Integration**: Direct integration with document sourcing system
* **Comprehensive Testing**: Full unit test coverage for all modules (175 tests)

## Architecture

1. **Ingestion**: Files are downloaded or collected, each paired with an institute ID.
2. **Versioning**: Each file is hashed. A new version is created whenever a new hash appears for the same institute.
3. **Rule Engine**: Checks are simple Python functions registered in a list. Each receives a workbook and returns:

   * Boolean outcome (True/False)
   * Outcome string (e.g., 'genügend', 'zu prüfen', or a numeric value as string)
   * Description string (explains why the outcome is as it is)
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
        sheet_mapper.py   # Multi-language sheet name mapping
      sourcing/           # Document sourcing
        __init__.py
        document_sourcer.py  # ORSADocumentSourcer
      reporting/          # Report generation
        __init__.py
        check_mapper.py   # Check to cell mapping
        template_manager.py  # Excel template operations
        report_generator.py  # Report orchestration
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

**Basic Processing:**

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

**Report Generation:**

```bash
# Process documents AND generate reports in one command
orsa-qc --generate-reports --verbose

# Specify custom output directory and template
orsa-qc --generate-reports --output-dir ./my_reports --template ./my_template.xlsx

# Generate reports only (without running checks) from existing database results
orsa-qc --reports-only --verbose

# Generate report for a specific institute
orsa-qc --reports-only --institute INST001 --verbose

# Force overwrite existing reports
orsa-qc --reports-only --force-overwrite --verbose

# Combine: process 2027 documents and generate reports
orsa-qc --berichtsjahr 2027 --generate-reports --verbose
```

Or run main.py directly:

```bash
python main.py --verbose

# Process and generate reports
python main.py --berichtsjahr 2026 --generate-reports --verbose

# Generate reports only
python main.py --reports-only --verbose
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

## Report Generation

The reporting module generates standalone Excel reports containing check results from templates. Reports are saved as separate files with the naming format: `Auswertung_{institute_id}_{source_file_name}.xlsx`

### Report Structure

Generated reports are standalone files containing only template sheets with populated check results:
1. **Template Sheets Only**: The "Auswertung" sheet and other template sheets (source file content is NOT included)
2. **Institut Metadata**: Institute-specific metadata is automatically populated:
   - Cell C3: FinmaID (Institute identifier)
   - Cell C4: FinmaObjektName (Institute name)
   - Cell C5: MitarbeiterName (Employee name)
3. **Check Results**: Check outcomes are written to specific cells (e.g., SST check → C8)

The institut metadata is sourced from the `institut_metadata.sql` query and merged with the report data using FinmaID as the key. The source ORSA file remains unchanged and separate from the generated report.

### Report Generation Workflow

**Via CLI:**

```bash
# Option 1: Process and generate reports in one command
orsa-qc --berichtsjahr 2026 --verbose

# Option 2: Generate reports separately from existing database results
orsa-qc --reports-only --verbose

# Generate report for specific institute
orsa-qc --reports-only --institute INST001 --verbose

# Specify custom output directory and template
orsa-qc --reports-only --output-dir ./my_reports --template ./my_template.xlsx
```

**Via Library:**

```python
from pathlib import Path
from orsa_analysis import DatabaseManager
from orsa_analysis.reporting import ReportGenerator

# Initialize database and report generator
db_manager = DatabaseManager()
report_gen = ReportGenerator(
    db_manager=db_manager,
    template_path=Path("data/auswertungs_template.xlsx"),
    output_dir=Path("reports")
)

# Generate report for single institute
report_path = report_gen.generate_report(
    institute_id="INST001",
    source_file_path=Path("data/orsa_response_files/INST001_orsa.xlsx"),
    force_overwrite=False
)

# Generate reports for all institutes with results
source_files = {
    "INST001": Path("data/orsa_response_files/INST001_orsa.xlsx"),
    "INST002": Path("data/orsa_response_files/INST002_orsa.xlsx"),
}
report_paths = report_gen.generate_all_reports(source_files=source_files)

print(f"Generated {len(report_paths)} reports")
```

### Check to Cell Mapping

The system maps check results to specific cells in the output Excel file. This mapping is defined in `check_mapper.py`:

```python
from orsa_analysis.reporting import CheckMapper, CellMapping

# View all mapped checks
mapper = CheckMapper()
checks = mapper.get_mapped_checks()
print(f"Mapped checks: {checks}")

# Get cell location for a check
mapping = mapper.get_cell_location("check_sst_three_years_filled")
print(f"SST check -> {mapping.sheet_name}!{mapping.cell_address}")

# Add custom mapping
mapper.add_mapping(
    "my_custom_check",
    CellMapping(
        cell_address="D15",
        value_type="outcome_bool",
        format_rule="boolean_to_text",
        sheet_name="Auswertung"
    )
)
```

### Supported Format Rules

The system includes several built-in format rules for transforming check results:

- `boolean_to_text`: True → "Pass", False → "Fail"
- `boolean_to_yes_no`: True → "Yes", False → "No"
- `boolean_inverse`: True → "Fail", False → "Pass" (for negative checks)
- `numeric_with_decimals`: Formats numbers with 2 decimal places
- `percentage`: Formats as percentage (0.856 → "85.6%")
- `count`: Converts to integer count
- `raw`: No transformation (passes value as-is)

### Report Versioning

Reports are automatically versioned based on the check results version in the database:
- Filename format: `{institute_id}_ORSA_Report_v{version}.xlsx`
- By default, existing reports are not overwritten
- Use `force_overwrite=True` to regenerate existing reports

### Customizing Reports

**Custom Cell Mappings:**

```python
from orsa_analysis.reporting import CheckMapper, CellMapping

# Create custom mappings
custom_mappings = {
    "check_sst_three_years_filled": CellMapping(
        cell_address="C8",
        value_type="outcome_bool",
        format_rule="boolean_inverse"
    ),
    "check_has_sheets": CellMapping(
        cell_address="C10",
        value_type="outcome_bool",
        format_rule="boolean_to_yes_no"
    ),
}

mapper = CheckMapper(mappings=custom_mappings)
```

**Custom Template:**

Place your custom template Excel file with an "Auswertung" sheet and specify its path:

```bash
orsa-qc --reports-only --template /path/to/custom_template.xlsx
```

## Multi-Language Support

The system automatically detects and handles ORSA Excel files in German, English, or French. 

### How it Works

1. **Automatic Language Detection**: When processing a file, the system analyzes sheet names to determine the language
2. **Sheet Name Mapping**: Checks can reference German sheet names (as the reference standard), and the system automatically translates to the correct language
3. **Transparent Access**: Check functions use `SheetNameMapper` to access sheets regardless of the file's language

### Supported Languages

| German (Reference) | English | French |
|-------------------|---------|--------|
| Ergebnisse_AVO-FINMA | Results_ISO-FINMA | Résultats_OS-FINMA |
| Auswertung | General details | Info. générales |
| Allgem. Angaben | Risks | Risques |
| Risiken | Measures | Mesures |
| Massnahmen | Scenarios | Scénarios |
| Ergebnisse_IFRS | Results_IFRS | Résultats_IFRS |
| Qual. & langfr. Risiken | Qual. & long-term risks | Risques qual. & à long terme |
| Schlussfolgerungen, Dokument. | Conclusions, documentation | Conclusions, document. |

### Example: Writing Multi-Language Checks

```python
from openpyxl.workbook.workbook import Workbook
from typing import Optional, Tuple
from orsa_analysis.checks.sheet_mapper import SheetNameMapper

def check_example(wb: Workbook) -> Tuple[bool, str, str]:
    """Example check using multi-language sheet mapping.
    
    Args:
        wb: Workbook to check
        
    Returns:
        Tuple of (outcome_bool, outcome_str, description)
    """
    # Create mapper - automatically detects language
    mapper = SheetNameMapper(wb)
    
    # Access sheet using German reference name
    # Works for German, English, or French files
    results_sheet = mapper.get_sheet("Ergebnisse_AVO-FINMA")
    
    if results_sheet is None:
        return False, "zu prüfen", "Required sheet not found"
    
    # Check your data
    value = results_sheet["A1"].value
    outcome = value is not None
    outcome_str = "genügend" if outcome else "zu prüfen"
    
    return outcome, outcome_str, "Example check passed"
```

## Adding a New Check

To add a check, open `src/orsa_analysis/checks/rules.py` and define a new function:

```python
from openpyxl.workbook.workbook import Workbook
from typing import Tuple

def check_example(wb: Workbook) -> Tuple[bool, str, str]:
    """Your custom check logic.
    
    Args:
        wb: Workbook to check
        
    Returns:
        Tuple of (outcome_bool, outcome_str, description)
    """
    desc = "Example check verifies data quality"
    outcome = True
    outcome_str = "genügend"  # Could also be 'zu prüfen' or a numeric value like '5.1'
    return outcome, outcome_str, desc
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
* outcome_str (NVARCHAR(100), NULL) - String outcome (e.g., 'genügend', 'zu prüfen', or numeric value as string)
* berichtsjahr (INT, NULL) - Reporting year for the document
* geschaeft_nr (NVARCHAR(50), NULL) - Business number (unique per institute/year)
* processed_timestamp (DATETIME2, DEFAULT GETDATE())

Indexes:
* Clustered index on id
* Non-clustered index on institute_id
* Non-clustered index on file_hash
* Non-clustered index on berichtsjahr
* Composite index on (institute_id, version)

### Database Views

Two convenience views are provided:

1. **vw_orsa_analysis_latest**: Shows only the latest version per institute (includes berichtsjahr)
2. **vw_orsa_analysis_summary**: Aggregates pass rates by institute and check (includes berichtsjahr)

### Institut Metadata Query

The system includes a query (`sql/institut_metadata.sql`) that retrieves institute metadata from the `DWHMart.dbo.Sachbearbeiter` table. This metadata is automatically integrated into generated reports and includes:

- **FINMAID**: Institute identifier (used as merge key)
- **FinmaObjektName**: Official institute name
- **MitarbeiterName**: Assigned employee name
- **MitarbeiterKuerzel**: Employee abbreviation
- **MitarbeiterOrgEinheit**: Organizational unit
- **ZulassungName**: License/authorization type
- **SachbearbeiterTypName**: Processor type

The metadata is merged with report data using FinmaID as the key and automatically populated in the "Auswertung" sheet (cells C3-C5).

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

For each institute, a standalone Excel report file is generated based on a template with the naming format: `Auswertung_{institute_id}_{source_file_name}.xlsx`

The report contains only template sheets (source file content is not included) with columns:

* Description
* Check result
* Assessment (free text)
* Resolved (yes/no)

The source ORSA files remain separate and unchanged.

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
- `tests/test_rules.py` - Check function tests (including multi-language SST check)
- `tests/test_sheet_mapper.py` - Multi-language sheet mapping tests (14 tests)
- `tests/test_processor.py` - DocumentProcessor integration tests
- `tests/test_document_sourcer.py` - ORSADocumentSourcer tests (25 tests)

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
- `check_sst_three_years_filled` - Verify SST data filled for three years (cells E42:G42, E43:G43, E45:G45)

### checks/sheet_mapper.py
Multi-language sheet name mapping system:
- `SheetNameMapper` - Maps German reference names to actual sheet names in any language
- `SHEET_NAME_MAPPING` - Dictionary containing all supported sheet name translations
- `get_sheet(german_ref)` - Get worksheet by German reference name
- `get_sheet_name(german_ref)` - Get actual sheet name in file's language
- `has_sheet(german_ref)` - Check if sheet exists

## Integration Notes

### ORSADocumentSourcer Output Format
The processor expects `List[Tuple[str, Path, str, str, int]]` from `sourcer.load()`:
```python
[
    ("INST001_report.xlsx", Path("/path/to/file1.xlsx"), "GNR123", "10001", 2026),
    ("INST002_report.xlsx", Path("/path/to/file2.xlsx"), "GNR456", "10002", 2026),
]
```

The tuple format is: `(document_name, file_path, geschaeft_nr, finma_id, berichtsjahr)`
- `document_name`: Name of the document file
- `file_path`: Path to the downloaded file
- `geschaeft_nr`: Business number (GeschaeftNr)
- `finma_id`: Institute identifier (FinmaID)
- `berichtsjahr`: Reporting year

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
