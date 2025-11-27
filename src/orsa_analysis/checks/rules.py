"""Quality check rules and registry.

Each check function receives a Workbook object and returns a tuple:
(outcome_bool, outcome_numeric, description)
"""

from typing import Callable, Optional, Tuple
from openpyxl.workbook.workbook import Workbook
import logging

logger = logging.getLogger(__name__)

CheckFunction = Callable[[Workbook], Tuple[bool, Optional[float], str]]


def check_has_sheets(wb: Workbook) -> Tuple[bool, Optional[float], str]:
    """Verify that the workbook contains at least one sheet.

    Args:
        wb: Workbook to check

    Returns:
        Tuple of (outcome, sheet_count, description)
    """
    sheet_count = len(wb.sheetnames)
    outcome = sheet_count > 0
    description = f"Workbook contains {sheet_count} sheet(s)"
    return outcome, float(sheet_count), description


def check_no_empty_sheets(wb: Workbook) -> Tuple[bool, Optional[float], str]:
    """Verify that no sheets are completely empty.

    Args:
        wb: Workbook to check

    Returns:
        Tuple of (outcome, empty_count, description)
    """
    empty_count = 0
    for sheet in wb.worksheets:
        # A sheet with max_row == 1 and max_column == 1 might just have the default empty cell
        # Check if it actually has any data
        is_empty = True
        if sheet.max_row > 0 and sheet.max_column > 0:
            # Check if there's any actual data in the sheet
            for row in sheet.iter_rows(max_row=min(sheet.max_row, 100)):  # Check first 100 rows
                for cell in row:
                    if cell.value is not None:
                        is_empty = False
                        break
                if not is_empty:
                    break
        
        if is_empty:
            empty_count += 1
            logger.warning(f"Empty sheet found: {sheet.title}")

    outcome = empty_count == 0
    description = (
        f"Found {empty_count} empty sheet(s)"
        if empty_count > 0
        else "All sheets contain data"
    )
    return outcome, float(empty_count), description


def check_first_sheet_has_data(wb: Workbook) -> Tuple[bool, Optional[float], str]:
    """Verify that the first sheet has data in cell A1.

    Args:
        wb: Workbook to check

    Returns:
        Tuple of (outcome, None, description)
    """
    if not wb.worksheets:
        return False, None, "No worksheets found"

    first_sheet = wb.worksheets[0]
    cell_a1 = first_sheet["A1"].value
    outcome = cell_a1 is not None
    description = (
        f"First sheet '{first_sheet.title}' has data in A1"
        if outcome
        else f"First sheet '{first_sheet.title}' has no data in A1"
    )
    return outcome, None, description


def check_sheet_names_unique(wb: Workbook) -> Tuple[bool, Optional[float], str]:
    """Verify that all sheet names are unique (should always pass in openpyxl).

    Args:
        wb: Workbook to check

    Returns:
        Tuple of (outcome, unique_count, description)
    """
    sheet_names = wb.sheetnames
    unique_names = set(sheet_names)
    outcome = len(sheet_names) == len(unique_names)
    description = (
        "All sheet names are unique"
        if outcome
        else f"Duplicate sheet names found: {len(sheet_names)} total, {len(unique_names)} unique"
    )
    return outcome, float(len(unique_names)), description


def check_row_count_reasonable(wb: Workbook) -> Tuple[bool, Optional[float], str]:
    """Verify that sheets don't have an unreasonable number of rows.

    Args:
        wb: Workbook to check

    Returns:
        Tuple of (outcome, max_rows, description)
    """
    max_allowed_rows = 1_000_000
    max_rows = 0

    for sheet in wb.worksheets:
        if sheet.max_row > max_rows:
            max_rows = sheet.max_row

    outcome = max_rows <= max_allowed_rows
    description = (
        f"Maximum row count is {max_rows:,} (within limit)"
        if outcome
        else f"Maximum row count is {max_rows:,} (exceeds limit of {max_allowed_rows:,})"
    )
    return outcome, float(max_rows), description


def check_has_expected_headers(wb: Workbook) -> Tuple[bool, Optional[float], str]:
    """Check if the first sheet has headers in the first row.

    This is a simplified example that just checks if first row has values.

    Args:
        wb: Workbook to check

    Returns:
        Tuple of (outcome, header_count, description)
    """
    if not wb.worksheets:
        return False, 0.0, "No worksheets found"

    first_sheet = wb.worksheets[0]
    if first_sheet.max_row == 0:
        return False, 0.0, f"Sheet '{first_sheet.title}' is empty"

    header_count = 0
    for cell in first_sheet[1]:
        if cell.value is not None and str(cell.value).strip():
            header_count += 1

    outcome = header_count > 0
    description = (
        f"Found {header_count} non-empty header cell(s) in first row"
        if outcome
        else "No headers found in first row"
    )
    return outcome, float(header_count), description


def check_no_merged_cells(wb: Workbook) -> Tuple[bool, Optional[float], str]:
    """Check if workbook contains any merged cells.

    Note: This check does not work with read-only mode workbooks.

    Args:
        wb: Workbook to check

    Returns:
        Tuple of (outcome, merged_count, description)
    """
    merged_count = 0
    try:
        for sheet in wb.worksheets:
            if hasattr(sheet, 'merged_cells'):
                merged_count += len(sheet.merged_cells.ranges)
            else:
                logger.warning(f"Sheet '{sheet.title}' does not support merged_cells check (likely read-only mode)")
                return True, None, "Merged cells check skipped (read-only mode)"
    except AttributeError:
        logger.warning("Workbook does not support merged_cells check (likely read-only mode)")
        return True, None, "Merged cells check skipped (read-only mode)"

    outcome = merged_count == 0
    description = (
        "No merged cells found"
        if outcome
        else f"Found {merged_count} merged cell range(s)"
    )
    return outcome, float(merged_count), description


REGISTERED_CHECKS: list[Tuple[str, CheckFunction]] = [
    ("has_sheets", check_has_sheets),
    ("no_empty_sheets", check_no_empty_sheets),
    ("first_sheet_has_data", check_first_sheet_has_data),
    ("sheet_names_unique", check_sheet_names_unique),
    ("row_count_reasonable", check_row_count_reasonable),
    ("has_expected_headers", check_has_expected_headers),
    ("no_merged_cells", check_no_merged_cells),
]


def get_all_checks() -> list[Tuple[str, CheckFunction]]:
    """Get all registered check functions.

    Returns:
        List of tuples (check_name, check_function)
    """
    return REGISTERED_CHECKS.copy()


def run_check(
    check_name: str, check_function: CheckFunction, workbook: Workbook
) -> Tuple[bool, Optional[float], str]:
    """Execute a single check with error handling.

    Args:
        check_name: Name of the check
        check_function: Check function to execute
        workbook: Workbook to check

    Returns:
        Tuple of (outcome, numeric_value, description)
    """
    try:
        return check_function(workbook)
    except Exception as e:
        logger.error(f"Check '{check_name}' failed with error: {e}")
        return False, None, f"Check failed with error: {str(e)}"
