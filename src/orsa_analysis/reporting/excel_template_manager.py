"""Template manager for Excel report generation."""

import logging
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook

logger = logging.getLogger(__name__)


class ExcelTemplateManager:
    """Manages Excel template operations for report generation.
    
    This manager loads the template file directly and modifies it in place.
    This approach preserves all formatting including conditional formatting,
    data validation, and other advanced Excel features that would be lost
    when copying sheets cell by cell.
    """

    def __init__(self, template_path: Path):
        """Initialize template manager.

        Args:
            template_path: Path to the Excel template file

        Raises:
            FileNotFoundError: If template file doesn't exist
        """
        if not template_path.exists():
            raise FileNotFoundError(f"Template file not found: {template_path}")

        self.template_path = template_path
        self.output_wb = None
        logger.info(f"Template manager initialized with: {template_path}")

    def create_output_workbook(self, source_path: Path) -> Workbook:
        """Load template workbook for modification.

        This method loads the template file directly. The workbook can then be
        modified and saved to a different location, preserving all formatting.
        The source file path is validated but its content is not included in the output.

        Args:
            source_path: Path to source ORSA Excel file (for validation only)

        Returns:
            Template workbook ready for modification
            
        Raises:
            FileNotFoundError: If source file doesn't exist
        """
        if not source_path.exists():
            raise FileNotFoundError(f"Source file not found: {source_path}")

        # Load template file directly - this preserves all formatting including
        # conditional formatting, data validation, and other advanced features
        self.output_wb = load_workbook(self.template_path)
        logger.debug(f"Loaded template with sheets: {self.output_wb.sheetnames}")

        logger.info(
            f"Loaded template workbook with {len(self.output_wb.sheetnames)} sheet(s)"
        )
        return self.output_wb

    def write_cell_value(
        self, sheet_name: str, cell_address: str, value: Any
    ) -> bool:
        """Write a value to a specific cell in the output workbook.

        Args:
            sheet_name: Name of the worksheet
            cell_address: Cell address (e.g., "A1", "B5")
            value: Value to write

        Returns:
            True if successful, False otherwise
        """
        if self.output_wb is None:
            logger.error("Output workbook not created yet")
            return False

        if sheet_name not in self.output_wb.sheetnames:
            logger.error(f"Sheet not found: {sheet_name}")
            return False

        try:
            sheet = self.output_wb[sheet_name]
            sheet[cell_address] = value
            logger.debug(f"Wrote value to {sheet_name}!{cell_address}: {value}")
            return True
        except Exception as e:
            logger.error(f"Error writing to {sheet_name}!{cell_address}: {e}")
            return False

    def save_workbook(self, output_path: Path) -> None:
        """Save the output workbook to file.

        Args:
            output_path: Path where to save the workbook
        """
        if self.output_wb is None:
            logger.error("No output workbook to save")
            return

        # Create output directory if needed
        output_path.parent.mkdir(parents=True, exist_ok=True)

        self.output_wb.save(output_path)
        logger.info(f"Saved output workbook to: {output_path}")

    def close(self) -> None:
        """Close all open workbooks."""
        self.output_wb = None
        logger.debug("Closed template manager workbooks")
