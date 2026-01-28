# How to Debug a Specific Check

## Quick Usage

1. **Edit the configuration** in `debug_check.py`:
   ```python
   FILE_PATH = "/path/to/your/file.xlsx"  # Your Excel file path
   INSTITUTE_ID = "10001"                  # Your institute ID (optional)
   ```

2. **Run the script**:
   ```bash
   cd /workspace/project/dataquality
   python debug_check.py
   ```

## What it shows:

✅ All available sheets in the workbook
✅ Which sheet is being used (Zweigniederlassungs vs Standard)
✅ The exact cell range being checked
✅ The value in EACH cell (or "EMPTY" if empty)
✅ Which specific cells are causing the failure
✅ The actual check result

## Example Output:

```
Loading file: /data/institute_10001.xlsx
Institute: 10001
================================================================================

📋 AVAILABLE SHEETS:
  - Ergebnisse
  - Other Sheet 1
  - Other Sheet 2

🔍 DETERMINING SHEET TO USE:
  ✓ Zweigniederlassungs version detected
  ✓ Using sheet: 'Ergebnisse'

📊 CHECKING TIED ASSETS (Gebundenes Vermögen):
  Range to check (Zweigniederlassungs): E38:G40
  ✓ E38: 1000
  ✓ F38: 1100
  ✓ G38: 1200
  ✗ E39: EMPTY        <-- Problem here!
  ✓ F39: 2100
  ✓ G39: 2200
  ✓ E40: 3000
  ✓ F40: 3100
  ✓ G40: 3200

🎯 RUNNING ACTUAL CHECK:
  Result: FAIL ✗
  Outcome: Prüfen
  Description: Gebundenes Vermögen ist nicht vollständig für drei Jahre ausgefüllt...

📈 SUMMARY:
  Total cells checked: 9
  Filled cells: 8
  Empty cells: 1

  ⚠️  Empty cells found: E39
  → Check needs ALL cells in range to be filled!
```

