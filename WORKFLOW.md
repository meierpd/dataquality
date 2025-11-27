# ORSA Quality Control System - Complete Workflow

## System Flow Diagram

```
┌─────────────────────────────────────────────────────────────────┐
│                    ORSA Quality Control System                   │
└─────────────────────────────────────────────────────────────────┘

                              START
                                │
                                ▼
┌───────────────────────────────────────────────────────────────┐
│  1. DOCUMENT SOURCING (ORSADocumentSourcer)                   │
├───────────────────────────────────────────────────────────────┤
│  • Load credentials from credentials.env                       │
│  • Connect to GBB_Reporting database (DatabaseManager)        │
│  • Execute SQL query (source_orsa_dokument_metadata.sql)      │
│  • Filter: _ORSA-Formular documents, year >= 2026             │
│  • Download files with NTLM auth                              │
│  • Save to orsa_response_files/                               │
│                                                                │
│  Output: List[Tuple[filename, Path]]                          │
└───────────────────────────────────────────────────────────────┘
                                │
                                ▼
┌───────────────────────────────────────────────────────────────┐
│  2. DOCUMENT PROCESSING (DocumentProcessor)                    │
├───────────────────────────────────────────────────────────────┤
│  For each document:                                           │
│    ┌─────────────────────────────────────────────────────┐   │
│    │ 2a. VERSIONING (VersionManager)                     │   │
│    ├─────────────────────────────────────────────────────┤   │
│    │ • Compute SHA-256 hash of file content              │   │
│    │ • Check cache: Is this hash already processed?      │   │
│    │ • If cached and not force_reprocess: SKIP           │   │
│    │ • If new or changed: Increment version number       │   │
│    │ • Store version mapping in memory                   │   │
│    └─────────────────────────────────────────────────────┘   │
│                            │                                   │
│                            ▼                                   │
│    ┌─────────────────────────────────────────────────────┐   │
│    │ 2b. EXCEL READING (ExcelReader)                     │   │
│    ├─────────────────────────────────────────────────────┤   │
│    │ • Load workbook with openpyxl                       │   │
│    │ • Extract sheets and metadata                       │   │
│    │ • Validate file format (.xlsx/.xlsm)                │   │
│    └─────────────────────────────────────────────────────┘   │
│                            │                                   │
│                            ▼                                   │
│    ┌─────────────────────────────────────────────────────┐   │
│    │ 2c. QUALITY CHECKS (checks/rules.py)                │   │
│    ├─────────────────────────────────────────────────────┤   │
│    │ Run all 7 checks:                                   │   │
│    │  ✓ check_has_sheets                                 │   │
│    │  ✓ check_no_empty_sheets                            │   │
│    │  ✓ check_first_sheet_has_data                       │   │
│    │  ✓ check_sheet_names_unique                         │   │
│    │  ✓ check_row_count_reasonable                       │   │
│    │  ✓ check_has_expected_headers                       │   │
│    │  ✓ check_no_merged_cells                            │   │
│    │                                                      │   │
│    │ Each check returns CheckResult with:                │   │
│    │  - passed (bool)                                    │   │
│    │  - error_message (optional)                         │   │
│    │  - timestamp                                        │   │
│    └─────────────────────────────────────────────────────┘   │
│                            │                                   │
│                            ▼                                   │
│    ┌─────────────────────────────────────────────────────┐   │
│    │ 2d. RESULT ACCUMULATION                             │   │
│    ├─────────────────────────────────────────────────────┤   │
│    │ • Collect all CheckResult objects                   │   │
│    │ • Pass to DatabaseWriter.write()                    │   │
│    │ • Store in memory for batch processing              │   │
│    └─────────────────────────────────────────────────────┘   │
│                                                                │
└───────────────────────────────────────────────────────────────┘
                                │
                                ▼
┌───────────────────────────────────────────────────────────────┐
│  3. DATABASE WRITING (MSSQLDatabaseWriter)                     │
├───────────────────────────────────────────────────────────────┤
│  • Connect to GBI_REPORTING database (DatabaseManager)        │
│  • Convert CheckResults to DataFrame                          │
│  • Batch insert to gbi.orsa_analysis_data table               │
│  • Commit transaction                                         │
│                                                                │
│  Table columns:                                               │
│    - file_name, check_name, institute_id, version             │
│    - passed, error_message, timestamp                         │
│    - file_hash, sheet_name, severity                          │
└───────────────────────────────────────────────────────────────┘
                                │
                                ▼
┌───────────────────────────────────────────────────────────────┐
│  4. SUMMARY & REPORTING                                        │
├───────────────────────────────────────────────────────────────┤
│  • Total files processed                                      │
│  • Total checks executed                                      │
│  • Checks passed / failed                                     │
│  • Pass rate percentage                                       │
│  • List of institutes                                         │
│  • Processing time                                            │
└───────────────────────────────────────────────────────────────┘
                                │
                                ▼
                               END
```

## Data Flow

### Input: ORSADocumentSourcer.load()
```python
[
    ("INS001_2026_ORSA-Formular.xlsx", Path("/path/to/file1.xlsx")),
    ("INS002_2026_ORSA-Formular.xlsx", Path("/path/to/file2.xlsx")),
    ...
]
```

### Processing: CheckResult Objects
```python
CheckResult(
    file_name="INS001_2026_ORSA-Formular.xlsx",
    check_name="check_has_sheets",
    institute_id="INS001",
    version=1,
    passed=True,
    error_message=None,
    timestamp=datetime(2026, 1, 15, 10, 30, 0),
    file_hash="a1b2c3d4...",
    sheet_name=None,
    severity="error"
)
```

### Output: Database Records
```sql
-- Table: GBI_REPORTING.gbi.orsa_analysis_data
file_name                          | check_name            | institute_id | version | passed | error_message | timestamp           | file_hash  | sheet_name | severity
-----------------------------------|-----------------------|--------------|---------|--------|---------------|---------------------|------------|------------|----------
INS001_2026_ORSA-Formular.xlsx     | check_has_sheets      | INS001       | 1       | 1      | NULL          | 2026-01-15 10:30:00 | a1b2c3d4.. | NULL       | error
INS001_2026_ORSA-Formular.xlsx     | check_no_empty_sheets | INS001       | 1       | 1      | NULL          | 2026-01-15 10:30:00 | a1b2c3d4.. | NULL       | error
INS001_2026_ORSA-Formular.xlsx     | check_first_sheet...  | INS001       | 1       | 0      | First sheet.. | 2026-01-15 10:30:00 | a1b2c3d4.. | Sheet1     | error
...
```

## Caching Behavior

### First Run (No Cache)
```
File: INS001_2026_ORSA-Formular.xlsx
Hash: abc123...
Cache: Not found
Action: PROCESS
Version: 1
Result: 7 checks executed, stored in DB
```

### Second Run (Unchanged File)
```
File: INS001_2026_ORSA-Formular.xlsx
Hash: abc123... (same)
Cache: Found (version 1)
Action: SKIP
Result: 0 checks executed
```

### Third Run (Modified File)
```
File: INS001_2026_ORSA-Formular.xlsx
Hash: xyz789... (different)
Cache: Found old version, new hash detected
Action: PROCESS
Version: 2 (incremented)
Result: 7 checks executed, stored in DB
```

### Force Reprocess Mode
```
File: INS001_2026_ORSA-Formular.xlsx
Hash: abc123... (same)
Cache: Found (version 1)
Action: PROCESS (force_reprocess=True)
Version: 2 (incremented anyway)
Result: 7 checks executed, stored in DB
```

## Database Queries for Monitoring

### Get Latest Results Per File
```sql
SELECT * FROM GBI_REPORTING.gbi.vw_orsa_analysis_latest
WHERE institute_id = 'INS001'
ORDER BY timestamp DESC;
```

### Get Pass Rate Summary
```sql
SELECT 
    institute_id,
    CAST(analysis_date AS DATE) as date,
    total_checks,
    passed_checks,
    failed_checks,
    pass_rate
FROM GBI_REPORTING.gbi.vw_orsa_analysis_summary
WHERE analysis_date >= DATEADD(day, -30, GETDATE())
ORDER BY institute_id, analysis_date DESC;
```

### Find Failing Checks
```sql
SELECT 
    file_name,
    check_name,
    error_message,
    timestamp
FROM GBI_REPORTING.gbi.orsa_analysis_data
WHERE passed = 0
  AND timestamp >= DATEADD(day, -7, GETDATE())
ORDER BY timestamp DESC;
```

### Track Version History
```sql
SELECT 
    file_name,
    version,
    COUNT(*) as check_count,
    SUM(CAST(passed AS INT)) as passed_count,
    MIN(timestamp) as first_check_time
FROM GBI_REPORTING.gbi.orsa_analysis_data
WHERE file_name = 'INS001_2026_ORSA-Formular.xlsx'
GROUP BY file_name, version
ORDER BY version DESC;
```

## Error Handling Flow

```
┌─────────────────────────────────────┐
│  Exception Occurs                   │
└─────────────────────────────────────┘
                │
                ▼
┌─────────────────────────────────────┐
│  Where did it occur?                │
└─────────────────────────────────────┘
                │
    ┌───────────┼───────────┐
    ▼           ▼           ▼
┌────────┐  ┌────────┐  ┌────────┐
│Document│  │Quality │  │Database│
│Sourcing│  │Check   │  │Write   │
└────────┘  └────────┘  └────────┘
    │           │           │
    ▼           ▼           ▼
┌────────┐  ┌────────┐  ┌────────┐
│Log     │  │Create  │  │Log     │
│error   │  │failed  │  │error   │
│        │  │Check   │  │        │
│Skip    │  │Result  │  │Raise   │
│file    │  │        │  │        │
│        │  │Continue│  │        │
│Continue│  │        │  │        │
└────────┘  └────────┘  └────────┘
```

## Performance Optimization

### Parallelization Opportunity
```python
# Current: Sequential processing
for file_name, file_path in documents:
    process_single_file(file_name, file_path)

# Future: Parallel processing
from concurrent.futures import ThreadPoolExecutor

with ThreadPoolExecutor(max_workers=4) as executor:
    futures = [
        executor.submit(process_single_file, name, path)
        for name, path in documents
    ]
    results = [f.result() for f in futures]
```

### Database Optimization
```python
# Current: Batch write at end
results = []
for doc in documents:
    results.extend(process_document(doc))
db_writer.write_all(results)  # Single transaction

# Alternative: Streaming writes
for doc in documents:
    results = process_document(doc)
    db_writer.write(results)  # Multiple transactions
    results.clear()  # Free memory
```

## Integration Points

### External Systems
1. **GBB_Reporting Database** (Read)
   - Source: Document metadata
   - Connection: DatabaseManager with NTLM/credentials
   - SQL: `source_orsa_dokument_metadata.sql`

2. **GBI_REPORTING Database** (Write)
   - Target: Quality check results
   - Connection: DatabaseManager with NTLM/credentials
   - Table: `gbi.orsa_analysis_data`

3. **NTLM File Server** (Read)
   - Source: ORSA Excel files
   - Authentication: requests-ntlm
   - Target: `orsa_response_files/` directory

### Configuration Sources
1. **Environment Variables**
   - `DB_USER`, `DB_PASSWORD`: Database credentials
   - `FINMA_USERNAME`, `FINMA_PASSWORD`: Document server credentials
   - `HTTP_PROXY`, `HTTPS_PROXY`: Proxy settings

2. **credentials.env File**
   - FINMA credentials
   - Located in project root
   - Excluded from version control

3. **SQL Files**
   - `source_orsa_dokument_metadata.sql`: Document query
   - `create_table_orsa_analysis_data.sql`: Schema definition

## Deployment Checklist

### Prerequisites
- [ ] Python 3.10+ installed
- [ ] Network access to `frbdata.finma.ch`
- [ ] Database credentials available
- [ ] FINMA credentials available
- [ ] Proxy configured (if needed)

### Setup Steps
1. [ ] Clone repository: `git clone https://github.com/meierpd/dataquality.git`
2. [ ] Install package: `pip install -e .`
3. [ ] Create `credentials.env` with FINMA credentials
4. [ ] Set environment variables: `DB_USER`, `DB_PASSWORD`
5. [ ] Run database schema: Execute `sql/create_table_orsa_analysis_data.sql`
6. [ ] Test connection: `python -c "from orsa_analysis.sourcing import ORSADocumentSourcer; s = ORSADocumentSourcer()"`
7. [ ] Run tests: `pytest tests/ -v`
8. [ ] Test processing: `orsa-qc --verbose`

### Scheduling (Production)
```bash
# Linux (crontab)
# Run daily at 6 AM
0 6 * * * cd /path/to/dataquality && /path/to/venv/bin/orsa-qc >> /var/log/orsa-qc.log 2>&1

# Windows (Task Scheduler)
# Task Name: ORSA Quality Control
# Trigger: Daily at 6:00 AM
# Action: Start a program
#   Program: C:\Python310\python.exe
#   Arguments: -m orsa_analysis.cli --verbose
#   Start in: C:\Projects\dataquality
```

## Monitoring & Alerts

### Health Check Query
```sql
-- Run this daily to check system health
SELECT 
    CAST(GETDATE() AS DATE) as check_date,
    COUNT(DISTINCT file_name) as files_processed,
    COUNT(*) as total_checks,
    SUM(CAST(passed AS INT)) as passed_checks,
    CAST(SUM(CAST(passed AS INT)) * 100.0 / COUNT(*) AS DECIMAL(5,2)) as pass_rate
FROM GBI_REPORTING.gbi.orsa_analysis_data
WHERE timestamp >= DATEADD(day, -1, GETDATE())
GROUP BY CAST(GETDATE() AS DATE);
```

### Alert Conditions
1. **No files processed in last 24 hours** → Email alert
2. **Pass rate < 90%** → Email alert
3. **Specific check failing repeatedly** → Email alert
4. **Database connection failure** → Email alert
5. **Disk space < 1GB** → Email alert

---

**Document Version**: 1.0  
**Last Updated**: 2025-11-26  
**Status**: Production Ready ✓
