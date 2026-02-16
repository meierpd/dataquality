"""Simple check result to cell mapping for report generation."""

from typing import Any, Dict, Optional, Tuple


# Simple mapping: check_name -> (sheet_name, outcome_cell, value_type, description_cell)
# value_type can be: "outcome_str" or "outcome_bool"
# description_cell is optional (can be None if no description mapping is needed)
# Note: check_name should match the name registered in REGISTERED_CHECKS (rules.py)
# 
# Daten sheet structure:
# C4-C7: Metadata (Versicherungsunternehmen, FINMA-ID, Aufsichtskategorie, FINMA-Sachbearbeiter)
# C8: check_orsa_version
# C9 onwards: All other checks
CHECK_MAPPINGS: Dict[str, Tuple[str, str, str, Optional[str]]] = {
    # Metadata will be at C4-C7 (handled separately in report_generator.py)
    # check_orsa_version is at C8
    "check_orsa_version": ("Daten", "C8", "outcome_str", "D8"),
    
    # All other checks start at C9 - order matches template
    "check_responsible_person": ("Daten", "C9", "outcome_str", "D9"),
    "check_data_recency_geschaeftsplanung": ("Daten", "C10", "outcome_str", "D10"),
    "check_data_recency_risikoidentifikation": ("Daten", "C11", "outcome_str", "D11"),
    "check_data_recency_szenarien": ("Daten", "C12", "outcome_str", "D12"),
    "check_board_approved_orsa": ("Daten", "C13", "outcome_str", "D13"),
    "check_risikobeurteilung_method": ("Daten", "C14", "outcome_str", "D14"),
    "check_risk_criteria_sufficient": ("Daten", "C15", "outcome_str", "D15"),
    "check_finanzmarktrisiko_count": ("Daten", "C16", "outcome_str", "D16"),
    "check_versicherungsrisiko_count": ("Daten", "C17", "outcome_str", "D17"),
    "check_kreditrisiko_count": ("Daten", "C18", "outcome_str", "D18"),
    "check_liquiditaetsrisiko_count": ("Daten", "C19", "outcome_str", "D19"),
    "check_operationelles_risiko_count": ("Daten", "C20", "outcome_str", "D20"),
    "check_strategisches_umfeld_risiko_count": ("Daten", "C21", "outcome_str", "D21"),
    "check_anderes_risiko_count": ("Daten", "C22", "outcome_str", "D22"),
    "check_count_number_mitigating_measures": ("Daten", "C23", "outcome_str", "D23"),
    "check_count_number_potential_mitigating_measures": ("Daten", "C24", "outcome_str", "D24"),
    "check_count_other_measures": ("Daten", "C25", "outcome_str", "D25"),
    "check_count_potential_other_measures": ("Daten", "C26", "outcome_str", "D26"),
    "check_risks_are_all_mitigated": ("Daten", "C27", "outcome_str", "D27"),
    "check_any_nonmitigating_measures": ("Daten", "C28", "outcome_str", "D28"),
    "check_count_number_mitigating_measures_other_effect": ("Daten", "C29", "outcome_str", "D29"),
    "check_count_number_mitigating_measures_risk_accepted": ("Daten", "C30", "outcome_str", "D30"),
    "check_count_number_other_measures_other_effect": ("Daten", "C31", "outcome_str", "D31"),
    "check_count_advers_scenarios": ("Daten", "C32", "outcome_str", "D32"),
    "check_count_existenzbedrohend_scenarios": ("Daten", "C33", "outcome_str", "D33"),
    "check_count_other_scenarios": ("Daten", "C34", "outcome_str", "D34"),
    "check_count_all_scenarios": ("Daten", "C35", "outcome_str", "D35"),
    "check_every_scenario_has_event": ("Daten", "C36", "outcome_str", "D36"),
    "check_count_scenarios_only_one_event": ("Daten", "C37", "outcome_str", "D37"),
    "check_count_scenarios_multiple_events": ("Daten", "C38", "outcome_str", "D38"),
    "check_every_scenario_has_risk": ("Daten", "C39", "outcome_str", "D39"),
    "check_count_scenarios_only_one_risk": ("Daten", "C40", "outcome_str", "D40"),
    "check_count_scenrios_multiple_risks": ("Daten", "C41", "outcome_str", "D41"),
    "check_business_planning_filled_three_years": ("Daten", "C42", "outcome_str", "D42"),
    "check_sst_filled_three_years": ("Daten", "C43", "outcome_str", "D43"),
    "check_tied_assets_filled_three_years": ("Daten", "C44", "outcome_str", "D44"),
    "check_provisions_filled_three_years": ("Daten", "C45", "outcome_str", "D45"),
    "check_liquidity_filled_three_years": ("Daten", "C46", "outcome_str", "D46"),
    "check_other_perspective_filled_three_years": ("Daten", "C47", "outcome_str", "D47"),
    "check_scenarios_business_planning_filled_three_years": ("Daten", "C48", "outcome_str", "D48"),
    "check_scenarios_sst_filled_three_years": ("Daten", "C49", "outcome_str", "D49"),
    "check_scenarios_tied_assets_filled_three_years": ("Daten", "C50", "outcome_str", "D50"),
    "check_scenarios_provisions_filled_three_years": ("Daten", "C51", "outcome_str", "D51"),
    "check_scenarios_liquidity_filled_three_years": ("Daten", "C52", "outcome_str", "D52"),
    "check_scenarios_other_perspective_filled_three_years": ("Daten", "C53", "outcome_str", "D53"),
    "check_count_longterm_risks": ("Daten", "C54", "outcome_str", "D54"),
    "check_treatment_of_qual_risks": ("Daten", "C55", "outcome_str", "D55"),
    "check_orsa_dokumentation_sufficient": ("Daten", "C56", "outcome_str", "D56"),

    # Add more mappings here as needed:
    # "check_name": ("SheetName", "A1", "outcome_str", "B1"),  # with description
    # "check_name_bool": ("SheetName", "A2", "outcome_bool", None),  # without description
}


class CheckToCellMapper:
    """Map check results to cell locations in output reports."""
    
    def __init__(self, mappings: Optional[Dict[str, Tuple[str, str, str, Optional[str]]]] = None):
        """Initialize check mapper.
        
        Args:
            mappings: Optional custom mapping dict. If None, uses CHECK_MAPPINGS.
                      Format: check_name -> (sheet_name, outcome_cell, value_type, description_cell)
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
    
    def get_cell_location(self, check_name: str) -> Optional[Tuple[str, str, str, Optional[str]]]:
        """Get cell location for a check.
        
        Args:
            check_name: Name of the check
            
        Returns:
            Tuple of (sheet_name, outcome_cell, value_type, description_cell) or None
        """
        return self.mappings.get(check_name)
    
    def get_value_from_result(self, check_result: Dict[str, Any],
                            sheet_name: str, cell_address: str, 
                            value_type: str) -> Any:
        """Extract value from check result.
        
        Args:
            check_result: Dictionary containing check result data
            sheet_name: Sheet name (not used, for API compatibility)
            cell_address: Cell address (not used, for API compatibility)
            value_type: Type of value to extract ("outcome_str", "outcome_bool", or "check_description")
            
        Returns:
            Value ready for cell insertion
        """
        if value_type == "outcome_str":
            return check_result.get("outcome_str")
        elif value_type == "outcome_bool":
            return check_result.get("outcome_bool")
        elif value_type == "check_description":
            return check_result.get("check_description")
        else:
            return None
    
    def add_mapping(self, check_name: str, sheet_name: str, 
                   cell_address: str, value_type: str, 
                   description_cell: Optional[str] = None) -> None:
        """Add or update a mapping.
        
        Args:
            check_name: Name of the check
            sheet_name: Target sheet name
            cell_address: Target cell address for outcome
            value_type: Type of value ("outcome_str" or "outcome_bool")
            description_cell: Optional cell address for check description
        """
        self.mappings[check_name] = (sheet_name, cell_address, value_type, description_cell)
    
    def get_mapped_checks(self) -> list[str]:
        """Get list of all check names that have mappings.
        
        Returns:
            List of check names
        """
        return list(self.mappings.keys())
