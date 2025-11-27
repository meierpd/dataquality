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
    
    This manager uses a simple approach:
    1. Load the source ORSA file as the base workbook
    2. Copy template sheets and insert them at the beginning
    3. Keep all original source sheets intact
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
        """Create output workbook by loading source file and prepending template sheets.

        This method:
        1. Loads the source ORSA file as the base workbook
        2. Copies template sheets and inserts them at the beginning
        3. Keeps all original source sheets intact

        Args:
            source_path: Path to source ORSA Excel file

        Returns:
            Output workbook with template sheets first, then source sheets
            
        Raises:
            FileNotFoundError: If source file doesn't exist
        """
        if not source_path.exists():
            raise FileNotFoundError(f"Source file not found: {source_path}")

        # Load source file as base
        self.output_wb = load_workbook(source_path)
        logger.debug(f"Loaded source file with sheets: {self.output_wb.sheetnames}")

        # Load template file
        template_wb = load_workbook(self.template_path)
        logger.debug(f"Loaded template with sheets: {template_wb.sheetnames}")

        # Copy template sheets and insert them at the beginning
        template_sheet_names = list(template_wb.sheetnames)
        for i, sheet_name in enumerate(template_sheet_names):
            # Handle potential name conflicts
            target_name = sheet_name
            if target_name in self.output_wb.sheetnames:
                target_name = f"{sheet_name}_template"
                logger.debug(
                    f"Sheet name conflict: renaming template sheet {sheet_name} to {target_name}"
                )
            
            # Copy the template sheet
            self._copy_sheet(template_wb[sheet_name], self.output_wb, target_name)
            logger.debug(f"Copied template sheet: {target_name}")

            # Move to position i (template sheets at beginning)
            self.output_wb.move_sheet(target_name, offset=-(len(self.output_wb.sheetnames) - 1 - i))

        logger.info(
            f"Created output workbook with {len(self.output_wb.sheetnames)} sheets "
            f"(template sheets first)"
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
