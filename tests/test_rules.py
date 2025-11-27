"""Unit tests for the rules module."""

import pytest
from openpyxl import Workbook

from orsa_analysis.checks.rules import (
    check_has_sheets,
    check_no_empty_sheets,
    check_first_sheet_has_data,
    check_sheet_names_unique,
    check_row_count_reasonable,
    check_has_expected_headers,
    check_no_merged_cells,
    get_all_checks,
    run_check,
)


@pytest.fixture
def basic_workbook():
    """Create a basic workbook for testing."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "Header1"
    ws["B1"] = "Header2"
    ws["A2"] = "Data1"
    ws["B2"] = "Data2"
    return wb


@pytest.fixture
def empty_workbook():
    """Create an empty workbook."""
    wb = Workbook()
    return wb


@pytest.fixture
def multi_sheet_workbook():
    """Create a workbook with multiple sheets."""
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Sheet1"
    ws1["A1"] = "Data"

    ws2 = wb.create_sheet("Sheet2")
    ws2["A1"] = "More data"

    return wb


class TestCheckHasSheets:
    """Test cases for check_has_sheets."""

    def test_workbook_with_sheets(self, basic_workbook):
        """Test workbook with sheets passes check."""
        outcome, value, description = check_has_sheets(basic_workbook)

        assert outcome is True
        assert value == 1.0
        assert "1" in description

    def test_workbook_multiple_sheets(self, multi_sheet_workbook):
        """Test workbook with multiple sheets."""
        outcome, value, description = check_has_sheets(multi_sheet_workbook)

        assert outcome is True
        assert value == 2.0
        assert "2" in description


class TestCheckNoEmptySheets:
    """Test cases for check_no_empty_sheets."""

    def test_workbook_with_data(self, basic_workbook):
        """Test workbook with data passes check."""
        outcome, value, description = check_no_empty_sheets(basic_workbook)

        assert outcome is True
        assert value == 0.0
        assert "All sheets contain data" in description

    def test_workbook_with_empty_sheet(self):
        """Test workbook with an empty sheet fails check."""
        wb = Workbook()
        wb.active["A1"] = "Data"
        wb.create_sheet("EmptySheet")

        outcome, value, description = check_no_empty_sheets(wb)

        assert outcome is False
        assert value == 1.0
        assert "1 empty sheet" in description


class TestCheckFirstSheetHasData:
    """Test cases for check_first_sheet_has_data."""

    def test_first_sheet_has_data(self, basic_workbook):
        """Test first sheet with data passes check."""
        outcome, value, description = check_first_sheet_has_data(basic_workbook)

        assert outcome is True
        assert value is None
        assert "has data in A1" in description

    def test_first_sheet_no_data(self, empty_workbook):
        """Test first sheet without data fails check."""
        outcome, value, description = check_first_sheet_has_data(empty_workbook)

        assert outcome is False
        assert value is None
        assert "no data in A1" in description


class TestCheckSheetNamesUnique:
    """Test cases for check_sheet_names_unique."""

    def test_unique_sheet_names(self, basic_workbook):
        """Test workbook with unique sheet names passes check."""
        outcome, value, description = check_sheet_names_unique(basic_workbook)

        assert outcome is True
        assert value == 1.0
        assert "unique" in description

    def test_multiple_unique_sheets(self, multi_sheet_workbook):
        """Test multiple sheets with unique names."""
        outcome, value, description = check_sheet_names_unique(multi_sheet_workbook)

        assert outcome is True
        assert value == 2.0


class TestCheckRowCountReasonable:
    """Test cases for check_row_count_reasonable."""

    def test_reasonable_row_count(self, basic_workbook):
        """Test workbook with reasonable row count passes check."""
        outcome, value, description = check_row_count_reasonable(basic_workbook)

        assert outcome is True
        assert value == 2.0
        assert "within limit" in description

    def test_empty_sheet_row_count(self, empty_workbook):
        """Test empty sheet has reasonable row count."""
        outcome, value, description = check_row_count_reasonable(empty_workbook)

        assert outcome is True
        assert value == 1.0


class TestCheckHasExpectedHeaders:
    """Test cases for check_has_expected_headers."""

    def test_workbook_with_headers(self, basic_workbook):
        """Test workbook with headers passes check."""
        outcome, value, description = check_has_expected_headers(basic_workbook)

        assert outcome is True
        assert value == 2.0
        assert "2 non-empty header" in description

    def test_workbook_without_headers(self, empty_workbook):
        """Test empty workbook fails header check."""
        outcome, value, description = check_has_expected_headers(empty_workbook)

        assert outcome is False
        assert "No headers found" in description or "empty" in description.lower()


class TestCheckNoMergedCells:
    """Test cases for check_no_merged_cells."""

    def test_no_merged_cells(self, basic_workbook):
        """Test workbook without merged cells passes check."""
        outcome, value, description = check_no_merged_cells(basic_workbook)

        assert outcome is True
        assert value == 0.0
        assert "No merged cells" in description

    def test_with_merged_cells(self):
        """Test workbook with merged cells fails check."""
        wb = Workbook()
        ws = wb.active
        ws.merge_cells("A1:B2")

        outcome, value, description = check_no_merged_cells(wb)

        assert outcome is False
        assert value == 1.0
        assert "merged cell" in description


class TestGetAllChecks:
    """Test cases for get_all_checks."""

    def test_get_all_checks_returns_list(self):
        """Test that get_all_checks returns a list."""
        checks = get_all_checks()

        assert isinstance(checks, list)
        assert len(checks) > 0

    def test_get_all_checks_format(self):
        """Test that each check is a tuple of (name, function)."""
        checks = get_all_checks()

        for check in checks:
            assert isinstance(check, tuple)
            assert len(check) == 2
            assert isinstance(check[0], str)
            assert callable(check[1])

    def test_get_all_checks_contains_expected(self):
        """Test that expected checks are in the registry."""
        checks = get_all_checks()
        check_names = [name for name, _ in checks]

        assert "has_sheets" in check_names
        assert "no_empty_sheets" in check_names
        assert "first_sheet_has_data" in check_names


class TestRunCheck:
    """Test cases for run_check wrapper function."""

    def test_run_check_success(self, basic_workbook):
        """Test successful check execution."""
        outcome, value, description = run_check(
            "test_check", check_has_sheets, basic_workbook
        )

        assert isinstance(outcome, bool)
        assert isinstance(description, str)

    def test_run_check_with_error(self, basic_workbook):
        """Test check execution with error handling."""

        def failing_check(wb):
            raise ValueError("Test error")

        outcome, value, description = run_check(
            "failing_check", failing_check, basic_workbook
        )

        assert outcome is False
        assert value is None
        assert "error" in description.lower()

    def test_run_check_all_registered(self, basic_workbook):
        """Test running all registered checks."""
        checks = get_all_checks()

        for check_name, check_function in checks:
            outcome, value, description = run_check(
                check_name, check_function, basic_workbook
            )

            assert isinstance(outcome, bool)
            assert isinstance(description, str)
            assert (value is None) or isinstance(value, float)
