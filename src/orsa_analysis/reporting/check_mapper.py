"""Check result to cell mapping configuration for report generation."""

from dataclasses import dataclass
from typing import Any, Callable, Dict, Optional


@dataclass
class CellMapping:
    """Configuration for mapping a check result to a cell.
    
    Attributes:
        cell_address: Excel cell address (e.g., "C8")
        value_type: Type of value to use ("outcome_bool", "outcome_numeric", "description")
        format_rule: Optional name of formatting rule to apply
        sheet_name: Name of sheet where cell is located
    """
    cell_address: str
    value_type: str
    format_rule: Optional[str] = None
    sheet_name: str = "Auswertung"


# Format rules for transforming check results to cell values
FORMAT_RULES: Dict[str, Callable[[Any], Any]] = {
    "boolean_to_text": lambda x: "Pass" if x else "Fail",
    "boolean_to_yes_no": lambda x: "Yes" if x else "No",
    "boolean_inverse": lambda x: "Fail" if x else "Pass",  # For checks where False is good
    "numeric_with_decimals": lambda x: f"{x:.2f}" if x is not None else "N/A",
    "count": lambda x: int(x) if x is not None else 0,
    "percentage": lambda x: f"{x*100:.1f}%" if x is not None else "N/A",
    "raw": lambda x: x,  # No transformation
}


# Global mapping configuration: check_name -> cell location
CHECK_CELL_MAPPINGS: Dict[str, CellMapping] = {
    "check_sst_three_years_filled": CellMapping(
        cell_address="C8",
        value_type="outcome_bool",
        format_rule="boolean_inverse",  # False means fail (data not filled)
        sheet_name="Auswertung"
    ),
    # Add more check mappings here as needed
    # Example:
    # "check_has_expected_headers": CellMapping(
    #     cell_address="C9",
    #     value_type="outcome_bool",
    #     format_rule="boolean_to_text",
    # ),
}


class CheckMapper:
    """Map check results to cell locations in output reports."""
    
    def __init__(self, mappings: Optional[Dict[str, CellMapping]] = None):
        """Initialize check mapper.
        
        Args:
            mappings: Optional custom mapping dictionary. If None, uses CHECK_CELL_MAPPINGS.
        """
        self.mappings = mappings if mappings is not None else CHECK_CELL_MAPPINGS
        self.format_rules = FORMAT_RULES
    
    def get_cell_location(self, check_name: str) -> Optional[CellMapping]:
        """Get cell mapping configuration for a check.
        
        Args:
            check_name: Name of the check
            
        Returns:
            CellMapping if mapping exists, None otherwise
        """
        return self.mappings.get(check_name)
    
    def has_mapping(self, check_name: str) -> bool:
        """Check if a mapping exists for this check.
        
        Args:
            check_name: Name of the check
            
        Returns:
            True if mapping exists, False otherwise
        """
        return check_name in self.mappings
    
    def format_value(self, value: Any, format_rule: Optional[str]) -> Any:
        """Apply formatting rule to a value.
        
        Args:
            value: Value to format
            format_rule: Name of format rule to apply, or None for no formatting
            
        Returns:
            Formatted value
            
        Raises:
            ValueError: If format_rule is not recognized
        """
        if format_rule is None:
            return value
        
        if format_rule not in self.format_rules:
            raise ValueError(f"Unknown format rule: {format_rule}")
        
        return self.format_rules[format_rule](value)
    
    def get_mapped_checks(self) -> list[str]:
        """Get list of all check names that have mappings.
        
        Returns:
            List of check names
        """
        return list(self.mappings.keys())
    
    def add_mapping(self, check_name: str, mapping: CellMapping) -> None:
        """Add or update a mapping.
        
        Args:
            check_name: Name of the check
            mapping: Cell mapping configuration
        """
        self.mappings[check_name] = mapping
    
    def get_value_from_result(self, check_result: Dict[str, Any], 
                            mapping: CellMapping) -> Any:
        """Extract and format the appropriate value from a check result.
        
        Args:
            check_result: Dictionary containing check result data
            mapping: Cell mapping configuration
            
        Returns:
            Formatted value ready for cell insertion
        """
        # Extract the requested value type
        if mapping.value_type == "outcome_bool":
            raw_value = check_result.get("outcome_bool")
        elif mapping.value_type == "outcome_numeric":
            raw_value = check_result.get("outcome_numeric")
        elif mapping.value_type == "description":
            raw_value = check_result.get("check_description")
        else:
            raw_value = None
        
        # Apply formatting if specified
        return self.format_value(raw_value, mapping.format_rule)
