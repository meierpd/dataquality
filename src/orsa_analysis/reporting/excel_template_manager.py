"""Template manager for Excel report generation."""

import logging
import shutil
from pathlib import Path
from typing import Any, Union, Optional

try:
    import xlwings as xw
    XLWINGS_AVAILABLE = True
except ImportError:
    XLWINGS_AVAILABLE = False
    xw = None

logger = logging.getLogger(__name__)


class ExcelTemplateManager:
    """Manages Excel template operations for report generation.
    
    This manager uses xlwings to interact with Excel natively, which preserves
    all formatting including conditional formatting, data validation dropdowns,
    and other advanced Excel features that would be lost when using openpyxl.
    
    The workflow is:
    1. Copy template to output location
    2. Open output file with Excel (via xlwings)
    3. Write values to cells
    4. Save using Excel's native save (preserves all features)
    5. Close Excel properly
    """

    def __init__(self, template_path: Path):
        """Initialize template manager.

        Args:
            template_path: Path to the Excel template file

        Raises:
            FileNotFoundError: If template file doesn't exist
            RuntimeError: If xlwings is not available
        """
        if not XLWINGS_AVAILABLE:
            raise RuntimeError(
                "xlwings is required for Excel template operations. "
                "Install it with: pip install xlwings"
            )
        
        if not template_path.exists():
            raise FileNotFoundError(f"Template file not found: {template_path}")

        self.template_path = template_path
        self.output_wb: Optional[xw.Book] = None
        self.output_path: Optional[Path] = None
        self.app: Optional[xw.App] = None
        logger.info(f"Template manager initialized with: {template_path}")

    def create_output_workbook(self, source_path: Path) -> None:
        """Prepare for workbook creation.

        This method validates the source file and initializes internal state.
        The actual workbook is created when save_workbook() is called, using
        xlwings to preserve all Excel features including data validation.

        Args:
            source_path: Path to source ORSA Excel file (for validation only)
            
        Raises:
            FileNotFoundError: If source file doesn't exist
        """
        if not source_path.exists():
            raise FileNotFoundError(f"Source file not found: {source_path}")

        # Initialize write buffer for batching cell writes
        self._write_buffer = []
        
        logger.info("Prepared for workbook creation from template")
        logger.debug(f"Source file validated: {source_path}")

    @staticmethod
    def _convert_numeric_string(value: Any) -> Union[int, float, Any]:
        """Convert string representations of numbers to actual numeric types.
        
        This ensures that numeric values stored as strings are converted to
        actual numbers before being written to Excel cells. This is important
        because Excel formulas work correctly with numeric types but may fail
        with string representations of numbers.
        
        Args:
            value: The value to potentially convert
            
        Returns:
            - int: If value is a string representing an integer (e.g., "42")
            - float: If value is a string representing a decimal (e.g., "42.5")
            - original value: If value is not a numeric string or already a number
            
        Examples:
            >>> _convert_numeric_string("42")
            42
            >>> _convert_numeric_string("42.5")
            42.5
            >>> _convert_numeric_string("genügend")
            "genügend"
            >>> _convert_numeric_string(42)
            42
        """
        # Only process string values
        if not isinstance(value, str):
            return value
        
        # Skip empty strings
        if not value or value.isspace():
            return value
        
        # Try to convert to numeric type
        try:
            # Check if it's an integer (no decimal point)
            if '.' not in value:
                # Try converting to int
                return int(value)
            else:
                # Try converting to float
                return float(value)
        except (ValueError, TypeError):
            # If conversion fails, return original value
            return value

    def write_cell_value(
        self, sheet_name: str, cell_address: str, value: Any
    ) -> bool:
        """Buffer a value to write to a specific cell.
        
        String representations of numbers are automatically converted to numeric
        types before writing to ensure Excel formulas work correctly.
        
        The actual write happens when save_workbook() is called.

        Args:
            sheet_name: Name of the worksheet
            cell_address: Cell address (e.g., "A1", "B5")
            value: Value to write (numeric strings are auto-converted)

        Returns:
            True if successfully buffered, False otherwise
        """
        if not hasattr(self, '_write_buffer'):
            logger.error("Output workbook not created yet - call create_output_workbook() first")
            return False

        try:
            # Convert numeric strings to actual numbers
            converted_value = self._convert_numeric_string(value)
            
            # Buffer the write operation
            self._write_buffer.append((sheet_name, cell_address, converted_value))
            
            # Log conversion if it occurred
            if converted_value != value and isinstance(value, str):
                logger.debug(
                    f"Buffered write: converted string '{value}' to {type(converted_value).__name__} "
                    f"{converted_value} for {sheet_name}!{cell_address}"
                )
            else:
                logger.debug(f"Buffered write to {sheet_name}!{cell_address}: {converted_value}")
            
            return True
        except Exception as e:
            logger.error(f"Error buffering write to {sheet_name}!{cell_address}: {e}")
            return False

    def save_workbook(self, output_path: Path) -> None:
        """Save the output workbook to file using Excel's native save.
        
        This method:
        1. Copies the template to the output location
        2. Opens the copy with Excel (via xlwings)
        3. Applies all buffered cell writes
        4. Saves using Excel's native save (preserves all features)
        5. Closes Excel properly

        Args:
            output_path: Path where to save the workbook
            
        Raises:
            RuntimeError: If Excel is not available
        """
        if not hasattr(self, '_write_buffer'):
            logger.error("No output workbook prepared - call create_output_workbook() first")
            return

        # Create output directory if needed
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Copy template to output location
        logger.info(f"Copying template to output: {output_path}")
        shutil.copy2(self.template_path, output_path)
        
        try:
            # Start Excel application (invisible, without alerts)
            logger.debug("Starting Excel application via xlwings")
            self.app = xw.App(visible=False, add_book=False)
            self.app.display_alerts = False
            self.app.screen_updating = False
            
            # Open the output file
            logger.debug(f"Opening output file: {output_path}")
            self.output_wb = self.app.books.open(str(output_path.absolute()))
            self.output_path = output_path
            
            # Apply all buffered writes
            logger.info(f"Applying {len(self._write_buffer)} buffered writes")
            for sheet_name, cell_address, value in self._write_buffer:
                try:
                    # Access the sheet
                    sheet = self.output_wb.sheets[sheet_name]
                    
                    # Write the value
                    sheet.range(cell_address).value = value
                    
                    logger.debug(f"Applied: {sheet_name}!{cell_address} = {value}")
                except Exception as e:
                    logger.error(f"Failed to write {sheet_name}!{cell_address}: {e}")
                    # Continue with other writes even if one fails
            
            # Save using Excel's native save (preserves all features)
            logger.debug("Saving workbook using Excel's native save")
            self.output_wb.save()
            
            logger.info(f"Saved output workbook to: {output_path}")
            
        except Exception as e:
            logger.error(f"Error during workbook save: {e}")
            raise
        finally:
            # Always clean up Excel resources
            self._cleanup()

    def _cleanup(self) -> None:
        """Clean up Excel resources (internal method).
        
        Closes the workbook and quits the Excel application.
        This ensures no Excel processes are left running.
        """
        try:
            if self.output_wb is not None:
                logger.debug("Closing workbook")
                self.output_wb.close()
                self.output_wb = None
        except Exception as e:
            logger.warning(f"Error closing workbook: {e}")
        
        try:
            if self.app is not None:
                logger.debug("Quitting Excel application")
                self.app.quit()
                self.app = None
        except Exception as e:
            logger.warning(f"Error quitting Excel application: {e}")
    
    def close(self) -> None:
        """Close all open workbooks and clean up Excel resources.
        
        This method ensures Excel processes are properly terminated.
        """
        self._cleanup()
        logger.debug("Closed template manager resources")
