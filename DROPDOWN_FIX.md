# Excel Data Validation Dropdown Preservation Fix

## Problem Statement

Previously, generated Excel reports would lose Data Validation dropdowns that were present in the template file. This happened because the report generation process used `openpyxl` to read and write Excel files, which re-serializes the entire Excel file structure and doesn't fully support all Excel features, particularly complex Data Validation structures.

### Root Cause

The issue occurred in `ExcelTemplateManager`:

```python
# OLD APPROACH (problematic)
self.output_wb = load_workbook(self.template_path)  # openpyxl
# ... write cells ...
self.output_wb.save(output_path)  # Re-serializes, losing validation
```

When `openpyxl.save()` is called, it:
1. Reads the Excel file structure into its internal representation
2. Writes a new Excel file from this representation
3. During this process, unsupported or complex Excel features are lost

Data Validation dropdowns (especially cascading dropdowns or complex rules) are among the features that can be lost during this re-serialization.

## Solution

The fix replaces `openpyxl` with `xlwings` for template operations. `xlwings` uses Excel's native COM API (on Windows) or AppleScript (on Mac) to interact with Excel directly, preserving all features.

### New Approach

```python
# NEW APPROACH (preserves all features)
1. Copy template file to output location
2. Open output file with Excel via xlwings
3. Write cell values using Excel's native API
4. Save using Excel's native save() method
5. Close Excel properly
```

### Implementation Details

**File: `src/orsa_analysis/reporting/excel_template_manager.py`**

Key changes:
1. **Buffered writes**: Cell writes are buffered during the `write_cell_value()` calls
2. **Late workbook creation**: The actual Excel workbook is only opened when `save_workbook()` is called
3. **Native Excel operations**: All operations use Excel's COM API via xlwings
4. **Proper cleanup**: Excel process is always terminated, even if errors occur

**Workflow:**
```python
manager = ExcelTemplateManager(template_path)
manager.create_output_workbook(source_path)  # Validates source, initializes buffer

# Buffer writes (no Excel opened yet)
manager.write_cell_value("Daten", "C4", "Test Institute")
manager.write_cell_value("Daten", "C5", "10001")

# Now Excel is opened, writes are applied, and file is saved
manager.save_workbook(output_path)  # Opens Excel, applies writes, saves, closes

manager.close()  # Cleanup (redundant as save_workbook already cleans up)
```

### Benefits

1. **Preserves all Excel features**: Data validation, conditional formatting, pivot tables, etc.
2. **Excel-native operations**: Uses the same save mechanism as Excel itself
3. **Efficient**: Excel is only open during the save operation
4. **Proper cleanup**: Excel processes are always terminated
5. **Error handling**: Graceful failure with clear error messages if Excel is not available

## Environment Requirements

### Required

- **Microsoft Excel** must be installed on the system where reports are generated
  - Windows: Excel 2010 or later (uses COM API)
  - Mac: Excel 2016 or later (uses AppleScript)
- **xlwings** Python package: `pip install xlwings>=0.30.0`

### Optional

- For Linux users: This approach requires Excel. Alternatives:
  - Use a Windows or Mac environment for report generation
  - Use LibreOffice Calc with `pyoo` (but may have compatibility issues)
  - Keep the old `openpyxl` approach (but dropdowns will be lost)

## Installation

Update dependencies:

```bash
# From the dataquality directory
pip install -e .
```

This will install `xlwings>=0.30.0` along with other dependencies.

## Verification

### Automated Verification Script

Run the included verification script:

```bash
cd dataquality
python verify_dropdown_preservation.py
```

This script will:
1. Create a test template with data validation dropdowns
2. Generate a report using the new ExcelTemplateManager
3. Provide instructions for manual verification

### Manual Verification Steps

1. **Create or use a template** with data validation dropdowns
2. **Generate a report** using the updated code
3. **Open the report** in Excel
4. **Click on cells** that should have dropdowns
5. **Verify**: Dropdown arrows should appear and work correctly

### Expected Results

- ✓ **Before this fix**: Dropdowns were missing in generated reports
- ✓ **After this fix**: Dropdowns are preserved and fully functional

## Backward Compatibility

### API Compatibility

The external API remains largely compatible:
- `ExcelTemplateManager.__init__(template_path)` - unchanged
- `create_output_workbook(source_path)` - return type changed from `Workbook` to `None` (return value was not used)
- `write_cell_value(sheet, cell, value)` - unchanged
- `save_workbook(output_path)` - unchanged
- `close()` - unchanged

### Behavior Changes

1. **Excel requirement**: Excel must now be installed (was not required before)
2. **Timing**: Excel is opened only during `save_workbook()` (was opened during `create_output_workbook()` before)
3. **Performance**: Slightly slower due to Excel startup, but more reliable

## Testing

### Unit Tests

The existing unit tests in `tests/test_excel_template_manager.py` will need updates because:
- They mock `openpyxl.Workbook` which is no longer used
- They expect `create_output_workbook()` to return a `Workbook` object
- They check `manager.output_wb` directly which is now an `xlwings.Book` object

### Integration Testing

For full integration testing:
1. Use a real template file with dropdowns
2. Run the report generation pipeline
3. Verify dropdowns are preserved in output

## Troubleshooting

### "xlwings is required" Error

**Problem**: `RuntimeError: xlwings is required for Excel template operations`

**Solution**: Install xlwings: `pip install xlwings`

### "Excel is not available" Error

**Problem**: xlwings cannot connect to Excel

**Solution**: 
- Ensure Microsoft Excel is installed
- On Windows: Excel COM API must be accessible
- On Mac: Excel must be installed in /Applications
- Try opening Excel manually to verify it works

### Excel Process Left Running

**Problem**: Excel.exe process remains after script completes

**Solution**: 
- This should not happen with the new code (cleanup in finally block)
- If it does, manually kill the process and report as a bug
- Temporary workaround: `taskkill /IM EXCEL.EXE /F` (Windows)

### Permission Errors

**Problem**: "Permission denied" when saving workbook

**Solution**:
- Ensure output file is not already open in Excel
- Check write permissions on output directory
- Close any Excel windows showing the file

## Migration Guide

### For Developers

If you have custom code that extends `ExcelTemplateManager`:

1. **Don't access `manager.output_wb` directly** - it's now an xlwings Book, not openpyxl Workbook
2. **Don't expect return value** from `create_output_workbook()` - it now returns None
3. **Install xlwings** in your development environment
4. **Ensure Excel is available** in your test/dev environment

### For End Users

No changes required. The command-line interface and report generation workflow remain the same:

```bash
# Same commands work as before
orsa-qc --berichtsjahr 2026
orsa-qc --berichtsjahr 2026 --reports-only --institute 10001
```

## Performance Considerations

### Before (openpyxl)
- Fast: Pure Python, no external dependencies
- Limited: Loses some Excel features

### After (xlwings)
- Slightly slower: Requires Excel startup (~1-2 seconds per report)
- Complete: Preserves all Excel features
- Worth it: Feature preservation is more important than speed for this use case

### Optimization

The new implementation batches all cell writes and only opens Excel once per report, minimizing overhead.

## Future Improvements

Possible enhancements:
1. **Batch processing**: Keep Excel open for multiple reports in sequence
2. **Parallel processing**: Use multiple Excel instances for parallel report generation
3. **Caching**: Reuse Excel instance across reports in the same session
4. **Progress reporting**: Add callbacks for long-running operations

## Related Files

- `src/orsa_analysis/reporting/excel_template_manager.py` - Main implementation
- `src/orsa_analysis/reporting/report_generator.py` - Uses ExcelTemplateManager
- `pyproject.toml` - Dependencies updated
- `verify_dropdown_preservation.py` - Verification script
- `tests/test_excel_template_manager.py` - Unit tests (need updating)

## References

- xlwings documentation: https://docs.xlwings.org/
- openpyxl limitations: https://openpyxl.readthedocs.io/en/stable/optimized.html
- Excel Data Validation: https://support.microsoft.com/en-us/office/apply-data-validation-to-cells-29fecbcc-d1b9-42c1-9d76-eff3ce5f7249
