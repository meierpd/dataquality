"""Template manager for Excel report generation."""

import logging
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

logger = logging.getLogger(__name__)


class ExcelTemplateManager:
    """Manages Excel template operations for report generation.
    
    This manager creates standalone report files from templates:
    1. Load the template file
    2. Copy all template sheets to a new workbook
    3. The source file is referenced but not included in the output
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
        """Create standalone output workbook from template.

        This method creates a new workbook containing only the template sheets.
        The source file path is validated but its content is not included in the output.

        Args:
            source_path: Path to source ORSA Excel file (for validation only)

        Returns:
            Output workbook containing only template sheets
            
        Raises:
            FileNotFoundError: If source file doesn't exist
        """
        if not source_path.exists():
            raise FileNotFoundError(f"Source file not found: {source_path}")

        # Load template file
        template_wb = load_workbook(self.template_path)
        logger.debug(f"Loaded template with sheets: {template_wb.sheetnames}")

        # Create a new workbook for output
        self.output_wb = Workbook()
        
        # Remove default sheet if it exists
        if "Sheet" in self.output_wb.sheetnames:
            self.output_wb.remove(self.output_wb["Sheet"])
        
        # Copy all template sheets to output workbook
        template_sheet_names = list(template_wb.sheetnames)
        for sheet_name in template_sheet_names:
            # Copy the template sheet
            self._copy_sheet(template_wb[sheet_name], self.output_wb, sheet_name)
            logger.debug(f"Copied template sheet: {sheet_name}")

        logger.info(
            f"Created standalone output workbook with {len(self.output_wb.sheetnames)} template sheet(s)"
        )
        return self.output_wb

    def _copy_sheet(
        self, source_sheet: Worksheet, target_wb: Workbook, new_name: str
    ) -> Worksheet:
        """Copy a worksheet to target workbook with all styles and formatting.

        Args:
            source_sheet: Source worksheet to copy
            target_wb: Target workbook
            new_name: Name for the new sheet

        Returns:
            Newly created worksheet
        """
        target_sheet = target_wb.create_sheet(new_name)

        # Copy cell values and styles
        for row in source_sheet.iter_rows():
            for cell in row:
                target_cell = target_sheet[cell.coordinate]
                target_cell.value = cell.value

                # Copy styles if present
                if cell.has_style:
                    target_cell.font = cell.font.copy()
                    target_cell.border = cell.border.copy()
                    target_cell.fill = cell.fill.copy()
                    target_cell.number_format = cell.number_format
                    target_cell.alignment = cell.alignment.copy()

        # Copy column dimensions
        for col_letter, dim in source_sheet.column_dimensions.items():
            target_sheet.column_dimensions[col_letter].width = dim.width

        # Copy row dimensions
        for row_num, dim in source_sheet.row_dimensions.items():
            target_sheet.row_dimensions[row_num].height = dim.height

        # Copy merged cells
        for merged_range in source_sheet.merged_cells.ranges:
            target_sheet.merge_cells(str(merged_range))

        return target_sheet

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
