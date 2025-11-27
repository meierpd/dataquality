"""Tests for CheckToCellMapper class."""

import pytest
from orsa_analysis.reporting.check_to_cell_mapper import CheckToCellMapper, CHECK_MAPPINGS


class TestCheckToCellMapperInitialization:
    """Test CheckToCellMapper initialization."""
    
    def test_init_with_defaults(self):
        """Test initialization with default mappings."""
        mapper = CheckToCellMapper()
        assert mapper.mappings == CHECK_MAPPINGS
    
    def test_init_with_custom_mappings(self):
        """Test initialization with custom mappings."""
        custom = {
            "test_check": ("Sheet1", "A1", "bool")
        }
        mapper = CheckToCellMapper(custom)
        assert mapper.mappings == custom


class TestHasMapping:
    """Test has_mapping method."""
    
    def test_has_mapping_true(self):
        """Test has_mapping returns True for existing check."""
        mapper = CheckToCellMapper({"test_check": ("Sheet1", "A1", "bool")})
        assert mapper.has_mapping("test_check") is True
    
    def test_has_mapping_false(self):
        """Test has_mapping returns False for non-existing check."""
        mapper = CheckToCellMapper({})
        assert mapper.has_mapping("nonexistent") is False


class TestGetCellLocation:
    """Test get_cell_location method."""
    
    def test_get_cell_location_exists(self):
        """Test getting cell location for existing check."""
        mapper = CheckToCellMapper({"test_check": ("Sheet1", "A1", "bool")})
        location = mapper.get_cell_location("test_check")
        
        assert location is not None
        assert location == ("Sheet1", "A1", "bool")
    
    def test_get_cell_location_not_exists(self):
        """Test getting cell location for non-existing check."""
        mapper = CheckToCellMapper({})
        location = mapper.get_cell_location("nonexistent")
        assert location is None


class TestGetValueFromResult:
    """Test get_value_from_result method."""
    
    def test_bool_value_fulfilled(self):
        """Test extracting bool value when check is fulfilled (1)."""
        mapper = CheckToCellMapper()
        result = {"outcome_bool": 1, "outcome_numeric": None}
        
        value = mapper.get_value_from_result(result, "Sheet1", "A1", "bool")
        assert value == "erfüllt"
    
    def test_bool_value_not_fulfilled(self):
        """Test extracting bool value when check is not fulfilled (0)."""
        mapper = CheckToCellMapper()
        result = {"outcome_bool": 0, "outcome_numeric": None}
        
        value = mapper.get_value_from_result(result, "Sheet1", "A1", "bool")
        assert value == "nicht erfüllt"
    
    def test_metric_value(self):
        """Test extracting metric value."""
        mapper = CheckToCellMapper()
        result = {"outcome_bool": 1, "outcome_numeric": 42.5}
        
        value = mapper.get_value_from_result(result, "Sheet1", "A1", "metric")
        assert value == 42.5
    
    def test_metric_value_none(self):
        """Test extracting metric value when it's None."""
        mapper = CheckToCellMapper()
        result = {"outcome_bool": 1, "outcome_numeric": None}
        
        value = mapper.get_value_from_result(result, "Sheet1", "A1", "metric")
        assert value is None
    
    def test_unknown_value_type(self):
        """Test extracting value with unknown type returns None."""
        mapper = CheckToCellMapper()
        result = {"outcome_bool": 1, "outcome_numeric": 42}
        
        value = mapper.get_value_from_result(result, "Sheet1", "A1", "unknown")
        assert value is None


class TestAddMapping:
    """Test add_mapping method."""
    
    def test_add_new_mapping(self):
        """Test adding a new mapping."""
        mapper = CheckToCellMapper({})
        mapper.add_mapping("new_check", "Sheet2", "B5", "bool")
        
        assert mapper.has_mapping("new_check")
        assert mapper.get_cell_location("new_check") == ("Sheet2", "B5", "bool")
    
    def test_update_existing_mapping(self):
        """Test updating an existing mapping."""
        mapper = CheckToCellMapper({"test_check": ("Sheet1", "A1", "bool")})
        mapper.add_mapping("test_check", "Sheet2", "B5", "metric")
        
        assert mapper.get_cell_location("test_check") == ("Sheet2", "B5", "metric")


class TestGetMappedChecks:
    """Test get_mapped_checks method."""
    
    def test_get_mapped_checks_empty(self):
        """Test getting checks when no mappings exist."""
        mapper = CheckToCellMapper({})
        checks = mapper.get_mapped_checks()
        assert checks == []
    
    def test_get_mapped_checks(self):
        """Test getting list of mapped checks."""
        mappings = {
            "check1": ("Sheet1", "A1", "bool"),
            "check2": ("Sheet1", "A2", "metric"),
        }
        mapper = CheckToCellMapper(mappings)
        checks = mapper.get_mapped_checks()
        
        assert len(checks) == 2
        assert "check1" in checks
        assert "check2" in checks


class TestDefaultMappings:
    """Test default CHECK_MAPPINGS configuration."""
    
    def test_default_mappings_exist(self):
        """Test that default mappings exist."""
        assert isinstance(CHECK_MAPPINGS, dict)
    
    def test_default_mapping_format(self):
        """Test that default mappings have correct format."""
        for check_name, mapping in CHECK_MAPPINGS.items():
            assert isinstance(check_name, str)
            assert isinstance(mapping, tuple)
            assert len(mapping) == 3
            sheet_name, cell_address, value_type = mapping
            assert isinstance(sheet_name, str)
            assert isinstance(cell_address, str)
            assert value_type in ["bool", "metric"]
