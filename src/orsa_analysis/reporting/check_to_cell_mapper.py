"""Simple check result to cell mapping for report generation."""

from typing import Any, Dict, Optional, Tuple


# Simple mapping: check_name -> (sheet_name, cell_address, value_type)
# value_type can be: "bool" or "metric"
# Note: check_name should match the name registered in REGISTERED_CHECKS (rules.py)
CHECK_MAPPINGS: Dict[str, Tuple[str, str, str]] = {
    "sst_three_years_filled": ("Auswertung", "C8", "bool"),
    # Add more mappings here as needed:
    # "check_name": ("SheetName", "A1", "bool"),
    # "check_name_with_metric": ("SheetName", "A2", "metric"),
}


class CheckToCellMapper:
    """Map check results to cell locations in output reports."""
    
    def __init__(self, mappings: Optional[Dict[str, Tuple[str, str, str]]] = None):
        """Initialize check mapper.
        
        Args:
            mappings: Optional custom mapping dict. If None, uses CHECK_MAPPINGS.
        """
        self.mappings = mappings if mappings is not None else CHECK_MAPPINGS
    
    def has_mapping(self, check_name: str) -> bool:
        """Check if a mapping exists for this check.
        
        Args:
            check_name: Name of the check
            
        Returns:
            True if mapping exists, False otherwise
        """
        return check_name in self.mappings
    
    def get_cell_location(self, check_name: str) -> Optional[Tuple[str, str, str]]:
        """Get cell location for a check.
        
        Args:
            check_name: Name of the check
            
        Returns:
            Tuple of (sheet_name, cell_address, value_type) or None
        """
        return self.mappings.get(check_name)
    
    def get_value_from_result(self, check_result: Dict[str, Any],
                            sheet_name: str, cell_address: str, 
                            value_type: str) -> Any:
        """Extract and format value from check result.
        
        Args:
            check_result: Dictionary containing check result data
            sheet_name: Sheet name (not used, for API compatibility)
            cell_address: Cell address (not used, for API compatibility)
            value_type: Type of value to extract ("bool" or "metric")
            
        Returns:
            Formatted value ready for cell insertion
        """
        if value_type == "bool":
            outcome_bool = check_result.get("outcome_bool")
            if outcome_bool == 1:
                return "erfüllt"
            else:
                return "nicht erfüllt"
        elif value_type == "metric":
            return check_result.get("outcome_numeric")
        else:
            return None
    
    def add_mapping(self, check_name: str, sheet_name: str, 
                   cell_address: str, value_type: str) -> None:
        """Add or update a mapping.
        
        Args:
            check_name: Name of the check
            sheet_name: Target sheet name
            cell_address: Target cell address
            value_type: Type of value ("bool" or "metric")
        """
        self.mappings[check_name] = (sheet_name, cell_address, value_type)
    
    def get_mapped_checks(self) -> list[str]:
        """Get list of all check names that have mappings.
        
        Returns:
            List of check names
        """
        return list(self.mappings.keys())
