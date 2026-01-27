"""Quality check rules and registry.

Each check function receives a Workbook object and returns a tuple:
(outcome_bool, outcome_numeric, description)
"""

import logging
import re
from datetime import date, datetime
from typing import Callable, Optional, Tuple

from openpyxl.workbook.workbook import Workbook
from openpyxl.utils import column_index_from_string, get_column_letter

from .sheet_mapper import SheetNameMapper

logger = logging.getLogger(__name__)

CheckFunction = Callable[[Workbook], Tuple[bool, str, str]]

STR_GUT = "gut"
STR_UNGENUEGEND = "ungenügend"

### Allgemeine Angaben

def check_responsible_person(wb: Workbook) -> Tuple[bool, str, str]:
    """
    Get the responsible person from the workbook, to fill it into the Auswertungs file.
    This is not a real check, but needed to get the value into the report.
    """
    try:
        mapper = SheetNameMapper(wb)
        sheet = mapper.get_sheet("Allgem. Angaben")

        if sheet is None:
            return False, "NA", "Sheet 'Allgem. Angaben' not found in workbook"

        value = sheet["C8"].value
        value = "" if value is None else str(value)

        return True, value, f"responsible person: {value}"

    except Exception as e:
        logger.error(f"Error in check_responsible_person: {e}")
        return False, "NA", f"Check failed with error: {str(e)}"


def _to_date(v) -> Optional[date]:
    if v is None:
        return None
    if isinstance(v, date) and not isinstance(v, datetime):
        return v
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, str):
        return datetime.strptime(v.strip(), "%d.%m.%Y").date()
    return None


def _months_diff(start: date, end: date) -> int:
    months = (end.year - start.year) * 12 + (end.month - start.month)
    if end.day < start.day:
        months -= 1
    return months


def check_data_recency_geschaeftsplanung(wb: Workbook) -> Tuple[bool, str, str]:
    """
    Check that the data which the orsa is based on is recent enough.
    """
    try:
        mapper = SheetNameMapper(wb)
        sheet = mapper.get_sheet("Allgem. Angaben")

        if sheet is None:
            return (
                False,
                STR_UNGENUEGEND,
                "Sheet 'Allgem. Angaben' not found in workbook",
            )

        approved = _to_date(sheet["C17"].value)
        snapshot = _to_date(sheet["C14"].value)

        if approved is None:
            return (
                False,
                STR_UNGENUEGEND,
                "Approval through board date (C17) is missing/invalid",
            )
        if snapshot is None:
            return False, STR_UNGENUEGEND, "Snapshot date (C14) is missing/invalid"
        if approved < snapshot:
            return False, STR_UNGENUEGEND, "Approval date is before snapshot date"

        months = _months_diff(snapshot, approved)
        ok = months <= 6

        return ok, STR_GUT if ok else STR_UNGENUEGEND, f"Data recency: {months} months"

    except Exception as e:
        logger.error(f"Error in check_data_recency_geschaeftsplanung: {e}")
        return False, STR_UNGENUEGEND, f"Check failed with error: {str(e)}"


def check_data_recency_risikoidentifikation(wb: Workbook) -> Tuple[bool, str, str]:
    """
    Check that the data which the orsa is based on is recent enough.
    """
    try:
        mapper = SheetNameMapper(wb)
        sheet = mapper.get_sheet("Allgem. Angaben")

        if sheet is None:
            return (
                False,
                STR_UNGENUEGEND,
                "Sheet 'Allgem. Angaben' not found in workbook",
            )

        approved = _to_date(sheet["C17"].value)
        snapshot = _to_date(sheet["C15"].value)

        if approved is None:
            return (
                False,
                STR_UNGENUEGEND,
                "Approval through board date (C17) is missing/invalid",
            )
        if snapshot is None:
            return False, STR_UNGENUEGEND, "Snapshot date (C15) is missing/invalid"
        if approved < snapshot:
            return False, STR_UNGENUEGEND, "Approval date is before snapshot date"

        months = _months_diff(snapshot, approved)
        ok = months <= 6

        return ok, STR_GUT if ok else STR_UNGENUEGEND, f"Data recency: {months} months"

    except Exception as e:
        logger.error(f"Error in check_data_recency_risikoidentifikation: {e}")
        return False, STR_UNGENUEGEND, f"Check failed with error: {str(e)}"


def check_data_recency_szenarien(wb: Workbook) -> Tuple[bool, str, str]:
    """
    Check that the data which the orsa is based on is recent enough.
    """
    try:
        mapper = SheetNameMapper(wb)
        sheet = mapper.get_sheet("Allgem. Angaben")

        if sheet is None:
            return (
                False,
                STR_UNGENUEGEND,
                "Sheet 'Allgem. Angaben' not found in workbook",
            )

        approved = _to_date(sheet["C17"].value)
        snapshot = _to_date(sheet["C16"].value)

        if approved is None:
            return (
                False,
                STR_UNGENUEGEND,
                "Approval through board date (C17) is missing/invalid",
            )
        if snapshot is None:
            return False, STR_UNGENUEGEND, "Snapshot date (C16) is missing/invalid"
        if approved < snapshot:
            return False, STR_UNGENUEGEND, "Approval date is before snapshot date"

        months = _months_diff(snapshot, approved)
        ok = months <= 6

        return ok, STR_GUT if ok else STR_UNGENUEGEND, f"Data recency: {months} months"

    except Exception as e:
        logger.error(f"Error in check_data_recency_szenarien: {e}")
        return False, STR_UNGENUEGEND, f"Check failed with error: {str(e)}"


def check_board_approved_orsa(wb: Workbook) -> Tuple[bool, str, str]:
    # Check if this is a Zweigniederlassungs version
    if _is_zweigniederlassungs_version(wb):
        return False, "kein Rating", "Kein Rating da es sich um eine Zweigniederlassung handelt"
    
    mapper = SheetNameMapper(wb)
    sheet = mapper.get_sheet("Allgem. Angaben")
    
    if sheet is None:
        return False, STR_UNGENUEGEND, "Sheet 'Allgem. Angaben' not found in workbook"

    value_who_approved = sheet["C18"].value or ""
    is_approved_through_board = value_who_approved in {
        "(1) gesamter VR",
        "(2) ein VR-Ausschuss",
    }

    return (
        is_approved_through_board,
        STR_GUT if is_approved_through_board else STR_UNGENUEGEND,
        value_who_approved,
    )

## Risiken

def check_risikobeurteilung_method(wb: Workbook) -> Tuple[bool, str, str]:
    mapper = SheetNameMapper(wb)
    sheet = mapper.get_sheet("Risiken")

    value_method = sheet["E7"].value or ""

    if value_method.startswith("(1)"):
        result_str = "gut"
    elif value_method.startswith("(2)"):
        result_str = "mangelhaft"
    elif value_method.startswith("(3)"):
        result_str = "ungenügend"
    elif value_method.startswith("(4)"):
        result_str = "kein Rating"
    else:
        result_str = ""

    return result_str == "gut", result_str, value_method


def check_risk_criteria_sufficient(wb: Workbook) -> Tuple[bool, str, str]:
    mapper = SheetNameMapper(wb)
    sheet = mapper.get_sheet("Risiken")

    required = {"(1)", "(2)", "(3)", "(4)", "(5)"}
    found = set()

    for row in range(10, 17):
        value = sheet[f"E{row}"].value or ""
        for r in required:
            if value.startswith(r):
                found.add(r)

    is_ok = required.issubset(found)
    result_str = "ausreichend" if is_ok else "nicht ausreichend"

    return is_ok, result_str, ", ".join(sorted(found))


def _count_risk(wb: Workbook, prefix: str) -> str:
    mapper = SheetNameMapper(wb)
    sheet = mapper.get_sheet("Risiken")

    count = 0
    for row in range(22, 52):
        value = sheet[f"E{row}"].value or ""
        if value.startswith(prefix):
            count += 1

    return str(count)


def check_finanzmarktrisiko_count(wb: Workbook) -> Tuple[bool, str, str]:
    count_str = _count_risk(wb, "(1)")
    return True, count_str, count_str


def check_versicherungsrisiko_count(wb: Workbook) -> Tuple[bool, str, str]:
    count_str = _count_risk(wb, "(2)")
    return True, count_str, count_str


def check_kreditrisiko_count(wb: Workbook) -> Tuple[bool, str, str]:
    count_str = _count_risk(wb, "(3)")
    return True, count_str, count_str


def check_liquiditaetsrisiko_count(wb: Workbook) -> Tuple[bool, str, str]:
    count_str = _count_risk(wb, "(4)")
    return True, count_str, count_str


def check_operationelles_risiko_count(wb: Workbook) -> Tuple[bool, str, str]:
    count_str = _count_risk(wb, "(5)")
    return True, count_str, count_str


def check_strategisches_umfeld_risiko_count(wb: Workbook) -> Tuple[bool, str, str]:
    count_str = _count_risk(wb, "(6)")
    return True, count_str, count_str


def check_anderes_risiko_count(wb: Workbook) -> Tuple[bool, str, str]:
    count_str = _count_risk(wb, "(7)")
    return True, count_str, count_str

## Massnahmen

def check_count_number_mitigating_measures(wb: Workbook) -> Tuple[bool, str, str]:
    mapper = SheetNameMapper(wb)
    sheet = mapper.get_sheet("Massnahmen")

    count = 0
    for row in range(9, 31):
        if (sheet[f"C{row}"].value or "") != "":
            count += 1

    count_str = str(count)
    return True, count_str, count_str


def check_count_number_potential_mitigating_measures(
    wb: Workbook,
) -> Tuple[bool, str, str]:
    mapper = SheetNameMapper(wb)
    sheet = mapper.get_sheet("Massnahmen")

    count = 0
    for row in range(9, 39):
        if (sheet[f"F{row}"].value or "") != "":
            count += 1

    count_str = str(count)
    return True, count_str, count_str


def check_count_other_measures(wb: Workbook) -> Tuple[bool, str, str]:
    mapper = SheetNameMapper(wb)
    sheet = mapper.get_sheet("Massnahmen")

    count = 0
    for row in range(44, 54):
        if (sheet[f"C{row}"].value or "") != "":
            count += 1

    count_str = str(count)
    return True, count_str, count_str


def check_count_potential_other_measures(wb: Workbook) -> Tuple[bool, str, str]:
    mapper = SheetNameMapper(wb)
    sheet = mapper.get_sheet("Massnahmen")

    count = 0
    for row in range(57, 67):
        if (sheet[f"C{row}"].value or "") != "":
            count += 1

    count_str = str(count)
    return True, count_str, count_str


def check_risks_are_all_mitigated(wb: Workbook) -> Tuple[bool, str, str]:
    mapper = SheetNameMapper(wb)
    sheet_risiken = mapper.get_sheet("Risiken")
    sheet_massnahmen = mapper.get_sheet("Massnahmen")

    required_ids = set()
    for row in range(22, 52):
        if (sheet_risiken[f"C{row}"].value or "") != "":
            rid = sheet_risiken[f"B{row}"].value
            if rid is not None:
                required_ids.add(str(rid))

    mitigated_ids = set()
    # regex extracts a leading number like "1" from strings starting with "(1) ..."
    risk_id_re = re.compile(r"^\s*\(?\s*(\d+)\s*\)?")

    for row in range(9, 39):
        value = sheet_massnahmen[f"C{row}"].value or ""
        m = risk_id_re.match(str(value))
        if m:
            mitigated_ids.add(m.group(1))

    missing = sorted(required_ids - mitigated_ids, key=int)

    if not missing:
        return True, "OK", "Jedes Risiko hat eine Massnahme"

    not_mitigated = ", ".join(missing)
    return False, "NOK", f"Risiko {not_mitigated} hat keine Massnahme"


def check_any_nonmitigating_measures(wb: Workbook) -> Tuple[bool, str, str]:
    mapper = SheetNameMapper(wb)
    sheet = mapper.get_sheet("Massnahmen")

    count = 0
    for row in range(9, 39):
        value = sheet[f"E{row}"].value or ""
        if str(value).startswith("(5)"):
            count += 1

    count_str = str(count)
    is_ok = count == 0
    outcome_str = "OK" if is_ok else "NOK"

    return is_ok, outcome_str, count_str


#### Szenarien
def _scenario_type_cells() -> list[str]:
    return [
        "C10",
        "C34",
        "C58",
        "C82",
        "C106",
        "C130",
        "C154",
        "C178",
        "C202",
        "C226",
        "C250",
        "C274",
        "C298",
        "C322",
        "C346",
    ]


def check_count_all_scenarios(wb: Workbook) -> Tuple[bool, str, str]:
    mapper = SheetNameMapper(wb)
    sheet = mapper.get_sheet("Szenarien")

    count = 0
    for addr in _scenario_type_cells():
        if (sheet[addr].value or "") != "":
            count += 1

    count_str = str(count)
    return True, count_str, count_str


def check_count_advers_scenarios(wb: Workbook) -> Tuple[bool, str, str]:
    mapper = SheetNameMapper(wb)
    sheet = mapper.get_sheet("Szenarien")

    count = 0
    for addr in _scenario_type_cells():
        value = str(sheet[addr].value or "")
        if value.startswith("(1)"):
            count += 1

    count_str = str(count)
    return True, count_str, count_str


def check_count_existenzbedrohend_scenarios(wb: Workbook) -> Tuple[bool, str, str]:
    mapper = SheetNameMapper(wb)
    sheet = mapper.get_sheet("Szenarien")

    count = 0
    for addr in _scenario_type_cells():
        value = str(sheet[addr].value or "")
        if value.startswith("(2)"):
            count += 1

    count_str = str(count)
    return True, count_str, count_str


def check_count_other_scenarios(wb: Workbook) -> Tuple[bool, str, str]:
    mapper = SheetNameMapper(wb)
    sheet = mapper.get_sheet("Szenarien")

    count = 0
    for addr in _scenario_type_cells():
        value = str(sheet[addr].value or "")
        if value.startswith("(3)"):
            count += 1

    count_str = str(count)
    return True, count_str, count_str

def check_every_scenario_has_event(wb: Workbook) -> Tuple[bool, str, str]:
    mapper = SheetNameMapper(wb)
    sheet = mapper.get_sheet("Szenarien")

    ok = True
    for i, type_addr in enumerate(_scenario_type_cells()):
        scenario_type = sheet[type_addr].value or ""
        if scenario_type == "":
            continue

        start_row = 14 + i * 24
        has_event = False
        for r in range(start_row, start_row + 5):  # C14-C18, then +24 each scenario
            if (sheet[f"C{r}"].value or "") != "":
                has_event = True
                break

        if not has_event:
            ok = False
            break

    if ok:
        return True, "OK", "Every scenario has at least one event"
    return False, "NOK", "Not every scenario has at least one event"

def check_count_scenarios_only_one_event(wb: Workbook) -> Tuple[bool, str, str]:
    mapper = SheetNameMapper(wb)
    sheet = mapper.get_sheet("Szenarien")

    count = 0
    for i, type_addr in enumerate(_scenario_type_cells()):
        scenario_type = sheet[type_addr].value or ""
        if scenario_type == "":
            continue

        start_row = 14 + i * 24
        event_count = 0
        for r in range(start_row, start_row + 5):
            if (sheet[f"C{r}"].value or "") != "":
                event_count += 1

        if event_count == 1:
            count += 1

    count_str = str(count)
    return True, count_str, count_str

def check_count_scenarios_multiple_events(wb: Workbook) -> Tuple[bool, str, str]:
    mapper = SheetNameMapper(wb)
    sheet = mapper.get_sheet("Szenarien")

    count = 0
    for i, type_addr in enumerate(_scenario_type_cells()):
        scenario_type = sheet[type_addr].value or ""
        if scenario_type == "":
            continue

        start_row = 14 + i * 24
        event_count = 0
        for r in range(start_row, start_row + 5):
            if (sheet[f"C{r}"].value or "") != "":
                event_count += 1

        if event_count > 1:
            count += 1

    count_str = str(count)
    return True, count_str, count_str

def check_every_scenario_has_risk(wb: Workbook) -> Tuple[bool, str, str]:
    mapper = SheetNameMapper(wb)
    sheet = mapper.get_sheet("Szenarien")

    ok = True
    for i, type_addr in enumerate(_scenario_type_cells()):
        scenario_type = sheet[type_addr].value or ""
        if scenario_type == "":
            continue

        start_row = 20 + i * 24  # C20-C29, then +24 each scenario
        has_risk = False
        for r in range(start_row, start_row + 10):
            if (sheet[f"C{r}"].value or "") != "":
                has_risk = True
                break

        if not has_risk:
            ok = False
            break

    if ok:
        return True, "OK", "Every scenario has at least one risk"
    return False, "NOK", "Not every scenario has at least one risk"

def check_count_scenarios_only_one_risk(wb: Workbook) -> Tuple[bool, str, str]:
    mapper = SheetNameMapper(wb)
    sheet = mapper.get_sheet("Szenarien")

    count = 0
    for i, type_addr in enumerate(_scenario_type_cells()):
        scenario_type = sheet[type_addr].value or ""
        if scenario_type == "":
            continue

        start_row = 20 + i * 24
        risk_count = 0
        for r in range(start_row, start_row + 10):
            if (sheet[f"C{r}"].value or "") != "":
                risk_count += 1

        if risk_count == 1:
            count += 1

    count_str = str(count)
    return True, count_str, count_str

def check_count_scenrios_multiple_risks(wb: Workbook) -> Tuple[bool, str, str]:
    mapper = SheetNameMapper(wb)
    sheet = mapper.get_sheet("Szenarien")

    count = 0
    for i, type_addr in enumerate(_scenario_type_cells()):
        scenario_type = sheet[type_addr].value or ""
        if scenario_type == "":
            continue

        start_row = 20 + i * 24
        risk_count = 0
        for r in range(start_row, start_row + 10):
            if (sheet[f"C{r}"].value or "") != "":
                risk_count += 1

        if risk_count > 1:
            count += 1

    count_str = str(count)
    return True, count_str, count_str

### Resultate AVO-FINMA / IFRS

def _is_any_filled(sheet, cells: list[str]) -> bool:
    for addr in cells:
        v = sheet[addr].value
        if v is not None and str(v).strip() != "":
            return True
    return False


def _is_zweigniederlassungs_version(wb: Workbook) -> bool:
    """Detect if this is a Zweigniederlassungs version of the Excel file.
    
    The zweigniederlassungs_version has the sheet 'Ergebnisse' (and language equivalents),
    whereas the standard version has 'Ergebnisse_IFRS' and 'Ergebnisse_AVO-FINMA'.
    
    Args:
        wb: Workbook to check
        
    Returns:
        True if this is a Zweigniederlassungs version, False otherwise
    """
    # Check directly in workbook sheet names to avoid logging warnings
    # when the sheet doesn't exist (normal for standard version)
    sheet_names = set(wb.sheetnames)
    
    # Check for 'Ergebnisse' in all supported languages
    zweigniederlassungs_sheets = {
        "Ergebnisse",  # German
        "Results",      # English
        "Résultats"    # French
    }
    
    return bool(zweigniederlassungs_sheets & sheet_names)


def _get_filled_results_sheet(wb: Workbook) -> Tuple[bool, str, str, object]:
    mapper = SheetNameMapper(wb)
    
    # Check if this is a Zweigniederlassungs version
    if _is_zweigniederlassungs_version(wb):
        sheet_ergebnisse = mapper.get_sheet("Ergebnisse")
        if sheet_ergebnisse is None:
            msg = "Error: Ergebnisse sheet not found in Zweigniederlassungs version"
            return False, msg, msg, None
        return True, "OK", "OK", sheet_ergebnisse
    
    # Standard version logic
    sheet_avo = mapper.get_sheet("Ergebnisse_AVO-FINMA")
    sheet_ifrs = mapper.get_sheet("Ergebnisse_IFRS")

    check_cells = ["E26", "F26", "G26"]
    avo_filled = _is_any_filled(sheet_avo, check_cells)
    ifrs_filled = _is_any_filled(sheet_ifrs, check_cells)

    if avo_filled and ifrs_filled:
        msg = "Error: Both result sheets are filled out (AVO-FINMA and IFRS)"
        return False, msg, msg, None

    if not avo_filled and not ifrs_filled:
        msg = "Error: No results are filled out (AVO-FINMA or IFRS)"
        return False, msg, msg, None

    return True, "OK", "OK", (sheet_avo if avo_filled else sheet_ifrs)


def _range_has_no_empty_cells(sheet, start_col: str, end_col: str, start_row: int, end_row: int) -> bool:
    for row in range(start_row, end_row + 1):
        for col_ord in range(ord(start_col), ord(end_col) + 1):
            addr = f"{chr(col_ord)}{row}"
            v = sheet[addr].value
            if v is None or str(v).strip() == "":
                return False
    return True

def check_business_planning_filled_three_years(wb: Workbook) -> Tuple[bool, str, str]:
    ok, outcome_str, details_str, sheet = _get_filled_results_sheet(wb)
    if not ok:
        return False, outcome_str, details_str

    # Check if Zweigniederlassungs version (by checking if sheet name contains "Ergebnisse" without suffix)
    is_zweigniederlassung = "Ergebnisse_" not in sheet.title and "Ergebnisse" in sheet.title
    is_avo = sheet.title == "Ergebnisse_AVO-FINMA"

    if is_zweigniederlassung:
        # Zweigniederlassungs version: E10-G20 and E23-E30
        ok_1 = _range_has_no_empty_cells(sheet, "E", "G", 10, 20)
        ok_2 = _range_has_no_empty_cells(sheet, "E", "E", 23, 30)
    elif is_avo:
        ok_1 = _range_has_no_empty_cells(sheet, "E", "G", 11, 21)
        ok_2 = _range_has_no_empty_cells(sheet, "E", "G", 24, 35)
    else:
        ok_1 = _range_has_no_empty_cells(sheet, "E", "G", 11, 23)
        ok_2 = _range_has_no_empty_cells(sheet, "E", "G", 26, 38)

    if ok_1 and ok_2:
        return True, "OK", "Business planning is filled for three years"

    return False, "Prüfen", "Business planning is not fully filled for three years"


def check_sst_filled_three_years(wb: Workbook) -> Tuple[bool, str, str]:
    # Check if this is a Zweigniederlassungs version
    if _is_zweigniederlassungs_version(wb):
        return False, "Kein Rating", "Kein Rating da es sich um eine Zweigniederlassung handelt"
    
    ok, outcome_str, details_str, sheet = _get_filled_results_sheet(wb)
    if not ok:
        return False, outcome_str, details_str

    is_avo = sheet.title == "Ergebnisse_AVO-FINMA"

    if is_avo:
        ok = _range_has_no_empty_cells(sheet, "E", "G", 42, 45)
    else:
        ok = _range_has_no_empty_cells(sheet, "E", "G", 44, 47)

    if ok:
        return True, "OK", "SST is filled for three years"

    return False, "Prüfen", "SST is not fully filled for three years"

def check_tied_assets_filled_three_years(wb: Workbook) -> Tuple[bool, str, str]:
    ok, outcome_str, details_str, sheet = _get_filled_results_sheet(wb)
    if not ok:
        return False, outcome_str, details_str

    # Check if Zweigniederlassungs version
    is_zweigniederlassung = "Ergebnisse_" not in sheet.title and "Ergebnisse" in sheet.title
    is_avo = sheet.title == "Ergebnisse_AVO-FINMA"

    if is_zweigniederlassung:
        # Zweigniederlassungs version: E38 to G40
        ok = _range_has_no_empty_cells(sheet, "E", "G", 38, 40)
    elif is_avo:
        ok = _range_has_no_empty_cells(sheet, "E", "G", 49, 51)
    else:
        ok = _range_has_no_empty_cells(sheet, "E", "G", 51, 54)

    if ok:
        return True, "OK", "Tied assets are filled for three years"

    return False, "Prüfen", "Tied assets are not fully filled for three years"
def check_provisions_filled_three_years(wb: Workbook) -> Tuple[bool, str, str]:
    ok, outcome_str, details_str, sheet = _get_filled_results_sheet(wb)
    if not ok:
        return False, outcome_str, details_str

    # Check if Zweigniederlassungs version
    is_zweigniederlassung = "Ergebnisse_" not in sheet.title and "Ergebnisse" in sheet.title
    is_avo = sheet.title == "Ergebnisse_AVO-FINMA"

    if is_zweigniederlassung:
        # Zweigniederlassungs version: E60 to G60
        ok_range = _range_has_no_empty_cells(sheet, "E", "G", 60, 60)
    elif is_avo:
        ok_range = _range_has_no_empty_cells(sheet, "E", "G", 71, 71)
    else:
        ok_range = _range_has_no_empty_cells(sheet, "E", "G", 73, 73)

    if ok_range:
        return True, "OK", "Provisions are filled for three years"

    return False, "Prüfen", "Provisions are not fully filled for three years"

def check_other_perspective_filled_three_years(wb: Workbook) -> Tuple[bool, str, str]:
    ok, outcome_str, details_str, sheet = _get_filled_results_sheet(wb)
    if not ok:
        return False, outcome_str, details_str

    # Check if Zweigniederlassungs version
    is_zweigniederlassung = "Ergebnisse_" not in sheet.title and "Ergebnisse" in sheet.title
    is_avo = sheet.title == "Ergebnisse_AVO-FINMA"

    if is_zweigniederlassung:
        # Zweigniederlassungs version: E63 to G63, E66 to G66, E69 to G69
        row_ranges = [(63, 63), (66, 66), (69, 69)]
    else:
        shift = 0 if is_avo else 2
        row_ranges = [
            (74, 74),
            (77, 77),
            (82, 84),
            (86, 86),
            (89, 89),
            (93, 93),
            (97, 97),
        ]

    for r1, r2 in row_ranges:
        if is_zweigniederlassung:
            for row in range(r1, r2 + 1):
                e_val = sheet[f"E{row}"].value
                f_val = sheet[f"F{row}"].value
                g_val = sheet[f"G{row}"].value

                e_filled = e_val is not None and str(e_val).strip() != ""
                f_filled = f_val is not None and str(f_val).strip() != ""
                g_filled = g_val is not None and str(g_val).strip() != ""

                if e_filled and not (f_filled and g_filled):
                    return False, "Prüfen", "Other perspective is partially filled (E filled but F/G missing)"
        else:
            for row in range(r1 + shift, r2 + shift + 1):
                e_val = sheet[f"E{row}"].value
                f_val = sheet[f"F{row}"].value
                g_val = sheet[f"G{row}"].value

                e_filled = e_val is not None and str(e_val).strip() != ""
                f_filled = f_val is not None and str(f_val).strip() != ""
                g_filled = g_val is not None and str(g_val).strip() != ""

                if e_filled and not (f_filled and g_filled):
                    return False, "Prüfen", "Other perspective is partially filled (E filled but F/G missing)"

    return True, "OK", "Other perspective is consistently filled (rows are either empty or fully filled)"



def _range_has_no_empty_cells_cols(sheet, start_col: str, end_col: str, start_row: int, end_row: int) -> bool:
    c1 = column_index_from_string(start_col)
    c2 = column_index_from_string(end_col)
    for row in range(start_row, end_row + 1):
        for c in range(c1, c2 + 1):
            addr = f"{get_column_letter(c)}{row}"
            v = sheet[addr].value
            if v is None or str(v).strip() == "":
                return False
    return True


def _scenario_cols(scenario_index_zero_based: int) -> Tuple[str, str]:
    start_idx = column_index_from_string("K") + scenario_index_zero_based * 6
    return get_column_letter(start_idx), get_column_letter(start_idx + 2)


def check_scenarios_business_planning_filled_three_years(wb: Workbook) -> Tuple[bool, str, str]:
    ok, outcome_str, details_str, results_sheet = _get_filled_results_sheet(wb)
    if not ok:
        return False, outcome_str, details_str

    mapper = SheetNameMapper(wb)
    szenarien_sheet = mapper.get_sheet("Szenarien")

    # Check if Zweigniederlassungs version
    is_zweigniederlassung = "Ergebnisse_" not in results_sheet.title and "Ergebnisse" in results_sheet.title
    is_avo = results_sheet.title == "Ergebnisse_AVO-FINMA"

    for i, type_addr in enumerate(_scenario_type_cells()):
        if (szenarien_sheet[type_addr].value or "") == "":
            continue

        start_col, end_col = _scenario_cols(i)

        if is_zweigniederlassung:
            # Zweigniederlassungs version: K10-M20 and K23 to M30
            ok_1 = _range_has_no_empty_cells_cols(results_sheet, start_col, end_col, 10, 20)
            ok_2 = _range_has_no_empty_cells_cols(results_sheet, start_col, end_col, 23, 30)
        elif is_avo:
            ok_1 = _range_has_no_empty_cells_cols(results_sheet, start_col, end_col, 11, 21)
            ok_2 = _range_has_no_empty_cells_cols(results_sheet, start_col, end_col, 24, 35)
        else:
            ok_1 = _range_has_no_empty_cells_cols(results_sheet, start_col, end_col, 11, 23)
            ok_2 = _range_has_no_empty_cells_cols(results_sheet, start_col, end_col, 26, 38)

        if not (ok_1 and ok_2):
            return False, "Prüfen", "Business planning is not fully filled for all scenarios"

    return True, "OK", "Business planning is filled for three years for all scenarios"


def check_scenarios_sst_filled_three_years(wb: Workbook) -> Tuple[bool, str, str]:
    # Check if this is a Zweigniederlassungs version
    if _is_zweigniederlassungs_version(wb):
        return False, "Kein Rating", "Kein Rating da es sich um eine Zweigniederlassung handelt"
    
    ok, outcome_str, details_str, results_sheet = _get_filled_results_sheet(wb)
    if not ok:
        return False, outcome_str, details_str

    mapper = SheetNameMapper(wb)
    szenarien_sheet = mapper.get_sheet("Szenarien")

    is_avo = results_sheet.title == "Ergebnisse_AVO-FINMA"

    for i, type_addr in enumerate(_scenario_type_cells()):
        if (szenarien_sheet[type_addr].value or "") == "":
            continue

        start_col, end_col = _scenario_cols(i)

        if is_avo:
            ok_range = _range_has_no_empty_cells_cols(results_sheet, start_col, end_col, 42, 45)
        else:
            ok_range = _range_has_no_empty_cells_cols(results_sheet, start_col, end_col, 44, 47)

        if not ok_range:
            return False, "Prüfen", "SST is not fully filled for all scenarios"

    return True, "OK", "SST is filled for three years for all scenarios"


def check_scenarios_tied_assets_filled_three_years(wb: Workbook) -> Tuple[bool, str, str]:
    ok, outcome_str, details_str, results_sheet = _get_filled_results_sheet(wb)
    if not ok:
        return False, outcome_str, details_str

    mapper = SheetNameMapper(wb)
    szenarien_sheet = mapper.get_sheet("Szenarien")

    # Check if Zweigniederlassungs version
    is_zweigniederlassung = "Ergebnisse_" not in results_sheet.title and "Ergebnisse" in results_sheet.title
    is_avo = results_sheet.title == "Ergebnisse_AVO-FINMA"

    for i, type_addr in enumerate(_scenario_type_cells()):
        if (szenarien_sheet[type_addr].value or "") == "":
            continue

        start_col, end_col = _scenario_cols(i)

        if is_zweigniederlassung:
            # Zweigniederlassungs version: K38 to M40
            ok_range = _range_has_no_empty_cells_cols(results_sheet, start_col, end_col, 38, 40)
        elif is_avo:
            ok_range = _range_has_no_empty_cells_cols(results_sheet, start_col, end_col, 49, 51)
        else:
            ok_range = _range_has_no_empty_cells_cols(results_sheet, start_col, end_col, 51, 54)

        if not ok_range:
            return False, "Prüfen", "Tied assets are not fully filled for all scenarios"

    return True, "OK", "Tied assets are filled for three years for all scenarios"


def check_scenarios_provisions_filled_three_years(wb: Workbook) -> Tuple[bool, str, str]:
    ok, outcome_str, details_str, results_sheet = _get_filled_results_sheet(wb)
    if not ok:
        return False, outcome_str, details_str

    mapper = SheetNameMapper(wb)
    szenarien_sheet = mapper.get_sheet("Szenarien")

    # Check if Zweigniederlassungs version
    is_zweigniederlassung = "Ergebnisse_" not in results_sheet.title and "Ergebnisse" in results_sheet.title
    is_avo = results_sheet.title == "Ergebnisse_AVO-FINMA"
    
    if is_zweigniederlassung:
        row = 60
    else:
        row = 71 if is_avo else 73

    for i, type_addr in enumerate(_scenario_type_cells()):
        if (szenarien_sheet[type_addr].value or "") == "":
            continue

        start_col, end_col = _scenario_cols(i)
        ok_range = _range_has_no_empty_cells_cols(results_sheet, start_col, end_col, row, row)

        if not ok_range:
            return False, "Prüfen", "Provisions are not fully filled for all scenarios"

    return True, "OK", "Provisions are filled for three years for all scenarios"


def check_scenarios_other_perspective_filled_three_years(wb: Workbook) -> Tuple[bool, str, str]:
    ok, outcome_str, details_str, results_sheet = _get_filled_results_sheet(wb)
    if not ok:
        return False, outcome_str, details_str

    mapper = SheetNameMapper(wb)
    szenarien_sheet = mapper.get_sheet("Szenarien")

    # Check if Zweigniederlassungs version
    is_zweigniederlassung = "Ergebnisse_" not in results_sheet.title and "Ergebnisse" in results_sheet.title
    is_avo = results_sheet.title == "Ergebnisse_AVO-FINMA"

    if is_zweigniederlassung:
        # Zweigniederlassungs version: K63 to M63, K66 to M66, K69 to M69
        row_ranges = [(63, 63), (66, 66), (69, 69)]
        shift = 0
    else:
        shift = 0 if is_avo else 2
        row_ranges = [
            (74, 74),
            (77, 77),
            (82, 84),
            (86, 86),
            (89, 89),
            (93, 93),
            (97, 97),
        ]

    for i, type_addr in enumerate(_scenario_type_cells()):
        if (szenarien_sheet[type_addr].value or "") == "":
            continue

        c_e, c_g = _scenario_cols(i)
        c_f = get_column_letter(column_index_from_string(c_e) + 1)

        for r1, r2 in row_ranges:
            for row in range(r1 + shift, r2 + shift + 1):
                e_val = results_sheet[f"{c_e}{row}"].value
                f_val = results_sheet[f"{c_f}{row}"].value
                g_val = results_sheet[f"{c_g}{row}"].value

                e_filled = e_val is not None and str(e_val).strip() != ""
                f_filled = f_val is not None and str(f_val).strip() != ""
                g_filled = g_val is not None and str(g_val).strip() != ""

                if e_filled and not (f_filled and g_filled):
                    return False, "Prüfen", "Other perspective is partially filled for at least one scenario"

    return True, "OK", "Other perspective is consistently filled for all scenarios"

###### Qual. & langfr. Risiken

def check_count_longterm_risks(wb: Workbook) -> Tuple[bool, str, str]:
    # Check if this is a Zweigniederlassungs version
    if _is_zweigniederlassungs_version(wb):
        return False, "Kein Rating", "Kein Rating da es sich um eine Zweigniederlassung handelt"
    
    mapper = SheetNameMapper(wb)
    sheet = mapper.get_sheet("Qual. & langfr. Risiken")

    count = 0
    for row in range(25, 40):
        if (sheet[f"C{row}"].value or "") != "":
            count += 1

    count_str = str(count)
    return True, count_str, count_str

def check_treatment_of_qual_risks(wb: Workbook) -> Tuple[bool, str, str]:
    # Check if this is a Zweigniederlassungs version
    if _is_zweigniederlassungs_version(wb):
        return False, "Kein Rating", "Kein Rating da es sich um eine Zweigniederlassung handelt"
    
    mapper = SheetNameMapper(wb)
    sheet = mapper.get_sheet("Qual. & langfr. Risiken")

    value = str(sheet["E4"].value or "")

    ok = value.startswith("(1)") or value.startswith("(2)") or value.startswith("(3)")
    outcome_str = "OK" if ok else "Prüfen"

    return ok, outcome_str, value

#### Schlussfolgerungen, Dokument.

def check_orsa_dokumentation_sufficient(wb: Workbook) -> Tuple[bool, str, str]:
    mapper = SheetNameMapper(wb)
    sheet = mapper.get_sheet("Schlussfolgerungen, Dokument.")

    values = [str(sheet[f"C{row}"].value or "") for row in range(24, 31)]

    if any(v.startswith("(3)") for v in values):
        result_str = "ungenügend"
    elif all(v.startswith("(2)") for v in values if v != "") and any(v != "" for v in values):
        result_str = "gut"
    elif any(v.startswith("(2)") for v in values):
        result_str = "mangelhaft"
    else:
        result_str = "Prüfen"

    return result_str == "gut", result_str, result_str


##########################



REGISTERED_CHECKS: list[Tuple[str, CheckFunction]] = [
    ("check_responsible_person", check_responsible_person),
    ("check_data_recency_geschaeftsplanung", check_data_recency_geschaeftsplanung),
    ("check_data_recency_risikoidentifikation",check_data_recency_risikoidentifikation,),
    ("check_data_recency_szenarien", check_data_recency_szenarien),
    ("check_board_approved_orsa", check_board_approved_orsa),
    ("check_risikobeurteilung_method", check_risikobeurteilung_method),
    ("check_risk_criteria_sufficient", check_risk_criteria_sufficient),
    ("check_finanzmarktrisiko_count", check_finanzmarktrisiko_count),
    ("check_versicherungsrisiko_count", check_versicherungsrisiko_count),
    ("check_kreditrisiko_count", check_kreditrisiko_count),
    ("check_liquiditaetsrisiko_count", check_liquiditaetsrisiko_count),
    ("check_operationelles_risiko_count", check_operationelles_risiko_count),
    ("check_strategisches_umfeld_risiko_count",check_strategisches_umfeld_risiko_count,),
    ("check_anderes_risiko_count", check_anderes_risiko_count),
    ("check_count_number_mitigating_measures", check_count_number_mitigating_measures),
    ("check_count_number_potential_mitigating_measures",check_count_number_potential_mitigating_measures,),
    ("check_count_other_measures", check_count_other_measures),
    ("check_count_potential_other_measures", check_count_potential_other_measures),
    ("check_risks_are_all_mitigated", check_risks_are_all_mitigated),
    ("check_any_nonmitigating_measures", check_any_nonmitigating_measures),
    ("check_count_all_scenarios", check_count_all_scenarios),
    ("check_count_advers_scenarios", check_count_advers_scenarios),
    ("check_count_existenzbedrohend_scenarios",check_count_existenzbedrohend_scenarios,),
    ("check_count_other_scenarios", check_count_other_scenarios),
    ("check_every_scenario_has_event", check_every_scenario_has_event),
    ("check_count_scenarios_only_one_event", check_count_scenarios_only_one_event),
    ("check_count_scenarios_multiple_events", check_count_scenarios_multiple_events),
    ("check_every_scenario_has_risk", check_every_scenario_has_risk),
    ("check_count_scenarios_only_one_risk", check_count_scenarios_only_one_risk),
    ("check_count_scenrios_multiple_risks", check_count_scenrios_multiple_risks),
    ("check_business_planning_filled_three_years", check_business_planning_filled_three_years),
    ("check_sst_filled_three_years", check_sst_filled_three_years),
    ("check_tied_assets_filled_three_years", check_tied_assets_filled_three_years),
    ("check_provisions_filled_three_years", check_provisions_filled_three_years),
    ("check_other_perspective_filled_three_years", check_other_perspective_filled_three_years),
    ("check_scenarios_business_planning_filled_three_years", check_scenarios_business_planning_filled_three_years),
    ("check_scenarios_sst_filled_three_years", check_scenarios_sst_filled_three_years),
    ("check_scenarios_tied_assets_filled_three_years", check_scenarios_tied_assets_filled_three_years),
    ("check_scenarios_provisions_filled_three_years", check_scenarios_provisions_filled_three_years),
    ("check_scenarios_other_perspective_filled_three_years", check_scenarios_other_perspective_filled_three_years),
    ("check_count_longterm_risks", check_count_longterm_risks),
    ("check_treatment_of_qual_risks", check_treatment_of_qual_risks),
    ("check_orsa_dokumentation_sufficient", check_orsa_dokumentation_sufficient),


]


def get_all_checks() -> list[Tuple[str, CheckFunction]]:
    """Get all registered check functions.

    Returns:
        List of tuples (check_name, check_function)
    """
    return REGISTERED_CHECKS.copy()


def run_check(
    check_name: str, check_function: CheckFunction, workbook: Workbook
) -> Tuple[bool, str, str]:
    """Execute a single check with error handling.

    Args:
        check_name: Name of the check
        check_function: Check function to execute
        workbook: Workbook to check

    Returns:
        Tuple of (outcome, outcome_str, description)
    """
    try:
        return check_function(workbook)
    except Exception as e:
        logger.error(f"Check '{check_name}' failed with error: {e}")
        return False, "zu prüfen", f"Check failed with error: {str(e)}"
