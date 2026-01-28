"""Quick 5-line debug - just copy and paste this into Python interpreter"""

# ============= CHANGE THESE =============
FILE_PATH = "/path/to/your/file.xlsx"
# =========================================

from openpyxl import load_workbook
from orsa_analysis.checks.rules import check_tied_assets_filled_three_years

wb = load_workbook(FILE_PATH, data_only=True, read_only=False)
result = check_tied_assets_filled_three_years(wb)
print(f"Result: {result[0]} (Pass={result[0]})")
print(f"Outcome: {result[1]}")
print(f"Description: {result[2]}")

# To see actual cell values for Zweigniederlassungs (E38:G40):
sheet = wb["Ergebnisse"]  # or whatever your sheet name is
print("\n=== Cell Values (E38:G40) ===")
for row in range(38, 41):
    for col in ["E", "F", "G"]:
        cell = f"{col}{row}"
        val = sheet[cell].value
        print(f"{cell}: {val if val is not None else 'EMPTY'}")
