"""Tests for CheckMapper class."""

import pytest
from orsa_analysis.reporting.check_mapper import (
    CheckMapper,
    CellMapping,
    CHECK_CELL_MAPPINGS,
    FORMAT_RULES
)


class TestCellMapping:
    """Test CellMapping dataclass."""
    
    def test_cell_mapping_creation(self):
        """Test creating a CellMapping."""
        mapping = CellMapping(
            cell_address="C8",
            value_type="outcome_bool",
            format_rule="boolean_to_text",
            sheet_name="Auswertung"
        )
        
        assert mapping.cell_address == "C8"
        assert mapping.value_type == "outcome_bool"
        assert mapping.format_rule == "boolean_to_text"
        assert mapping.sheet_name == "Auswertung"
    
    def test_cell_mapping_defaults(self):
        """Test CellMapping default values."""
        mapping = CellMapping(
            cell_address="D10",
            value_type="outcome_numeric"
        )
        
        assert mapping.format_rule is None
        assert mapping.sheet_name == "Auswertung"


class TestFormatRules:
    """Test format rules."""
    
    def test_boolean_to_text(self):
        """Test boolean to text conversion."""
        rule = FORMAT_RULES["boolean_to_text"]
        assert rule(True) == "Pass"
        assert rule(False) == "Fail"
    
    def test_boolean_to_yes_no(self):
        """Test boolean to yes/no conversion."""
        rule = FORMAT_RULES["boolean_to_yes_no"]
        assert rule(True) == "Yes"
        assert rule(False) == "No"
    
    def test_boolean_inverse(self):
        """Test inverse boolean conversion."""
        rule = FORMAT_RULES["boolean_inverse"]
        assert rule(True) == "Fail"
        assert rule(False) == "Pass"
    
    def test_numeric_with_decimals(self):
        """Test numeric formatting."""
        rule = FORMAT_RULES["numeric_with_decimals"]
        assert rule(123.456) == "123.46"
        assert rule(10) == "10.00"
        assert rule(None) == "N/A"
    
    def test_count(self):
        """Test count formatting."""
        rule = FORMAT_RULES["count"]
        assert rule(5.7) == 5
        assert rule(10) == 10
        assert rule(None) == 0
    
    def test_percentage(self):
        """Test percentage formatting."""
        rule = FORMAT_RULES["percentage"]
        assert rule(0.856) == "85.6%"
        assert rule(1.0) == "100.0%"
        assert rule(None) == "N/A"
    
    def test_raw(self):
        """Test raw (no transformation)."""
        rule = FORMAT_RULES["raw"]
        assert rule("test") == "test"
        assert rule(123) == 123
        assert rule(None) is None


class TestCheckMapper:
    """Test CheckMapper class."""
    
    def test_initialization_default(self):
        """Test CheckMapper initialization with defaults."""
        mapper = CheckMapper()
        assert mapper.mappings == CHECK_CELL_MAPPINGS
        assert mapper.format_rules == FORMAT_RULES
    
    def test_initialization_custom(self):
        """Test CheckMapper initialization with custom mappings."""
        custom_mappings = {
            "test_check": CellMapping(
                cell_address="A1",
                value_type="outcome_bool"
            )
        }
        mapper = CheckMapper(mappings=custom_mappings)
        assert mapper.mappings == custom_mappings
    
    def test_get_cell_location_exists(self):
        """Test getting cell location for existing check."""
        mapper = CheckMapper()
        mapping = mapper.get_cell_location("check_sst_three_years_filled")
        
        assert mapping is not None
        assert mapping.cell_address == "C8"
        assert mapping.value_type == "outcome_bool"
    
    def test_get_cell_location_not_exists(self):
        """Test getting cell location for non-existing check."""
        mapper = CheckMapper()
        mapping = mapper.get_cell_location("nonexistent_check")
        assert mapping is None
    
    def test_has_mapping_true(self):
        """Test has_mapping returns True for existing check."""
        mapper = CheckMapper()
        assert mapper.has_mapping("check_sst_three_years_filled") is True
    
    def test_has_mapping_false(self):
        """Test has_mapping returns False for non-existing check."""
        mapper = CheckMapper()
        assert mapper.has_mapping("nonexistent_check") is False
    
    def test_format_value_with_rule(self):
        """Test formatting value with a rule."""
        mapper = CheckMapper()
        result = mapper.format_value(True, "boolean_to_text")
        assert result == "Pass"
    
    def test_format_value_without_rule(self):
        """Test formatting value without a rule."""
        mapper = CheckMapper()
        result = mapper.format_value("test_value", None)
        assert result == "test_value"
    
    def test_format_value_unknown_rule(self):
        """Test formatting with unknown rule raises error."""
        mapper = CheckMapper()
        with pytest.raises(ValueError, match="Unknown format rule"):
            mapper.format_value(True, "unknown_rule")
    
    def test_get_mapped_checks(self):
        """Test getting list of mapped checks."""
        mapper = CheckMapper()
        checks = mapper.get_mapped_checks()
        assert isinstance(checks, list)
        assert "check_sst_three_years_filled" in checks
    
    def test_add_mapping(self):
        """Test adding a new mapping."""
        mapper = CheckMapper()
        new_mapping = CellMapping(
            cell_address="E10",
            value_type="outcome_numeric"
        )
        
        mapper.add_mapping("new_check", new_mapping)
        assert mapper.has_mapping("new_check")
        assert mapper.get_cell_location("new_check") == new_mapping
    
    def test_get_value_from_result_outcome_bool(self):
        """Test extracting outcome_bool from result."""
        mapper = CheckMapper()
        result = {
            "check_name": "test_check",
            "outcome_bool": True,
            "outcome_numeric": 5.0,
            "check_description": "Test description"
        }
        mapping = CellMapping(
            cell_address="A1",
            value_type="outcome_bool",
            format_rule="boolean_to_text"
        )
        
        value = mapper.get_value_from_result(result, mapping)
        assert value == "Pass"
    
    def test_get_value_from_result_outcome_numeric(self):
        """Test extracting outcome_numeric from result."""
        mapper = CheckMapper()
        result = {
            "outcome_numeric": 123.456,
        }
        mapping = CellMapping(
            cell_address="A1",
            value_type="outcome_numeric",
            format_rule="numeric_with_decimals"
        )
        
        value = mapper.get_value_from_result(result, mapping)
        assert value == "123.46"
    
    def test_get_value_from_result_description(self):
        """Test extracting description from result."""
        mapper = CheckMapper()
        result = {
            "check_description": "Test check description",
        }
        mapping = CellMapping(
            cell_address="A1",
            value_type="description"
        )
        
        value = mapper.get_value_from_result(result, mapping)
        assert value == "Test check description"
    
    def test_get_value_from_result_no_formatting(self):
        """Test extracting value without formatting."""
        mapper = CheckMapper()
        result = {
            "outcome_bool": False,
        }
        mapping = CellMapping(
            cell_address="A1",
            value_type="outcome_bool",
            format_rule=None
        )
        
        value = mapper.get_value_from_result(result, mapping)
        assert value is False


class TestCheckCellMappings:
    """Test the global CHECK_CELL_MAPPINGS configuration."""
    
    def test_sst_check_mapping_exists(self):
        """Test that SST check mapping is configured."""
        assert "check_sst_three_years_filled" in CHECK_CELL_MAPPINGS
    
    def test_sst_check_mapping_correct(self):
        """Test that SST check mapping has correct configuration."""
        mapping = CHECK_CELL_MAPPINGS["check_sst_three_years_filled"]
        
        assert mapping.cell_address == "C8"
        assert mapping.value_type == "outcome_bool"
        assert mapping.format_rule == "boolean_inverse"
        assert mapping.sheet_name == "Auswertung"
