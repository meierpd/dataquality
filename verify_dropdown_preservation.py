#!/usr/bin/env python3
"""
Verification script for Excel Data Validation dropdown preservation.

This script demonstrates that the updated ExcelTemplateManager preserves
Excel Data Validation dropdowns when generating reports.

REQUIREMENTS:
- Microsoft Excel must be installed on the system
- xlwings must be installed (pip install xlwings)
- A template file with data validation dropdowns

USAGE:
    python verify_dropdown_preservation.py

VERIFICATION STEPS:
1. Run this script to generate a test report
2. Open the generated report in Excel
3. Click on cells that should have dropdowns
4. Verify that the dropdown arrows appear and work correctly

If dropdowns are preserved: ✓ Fix is working
If dropdowns are missing: ✗ Issue still exists
"""

import logging
import sys
from pathlib import Path

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def create_template_with_dropdown(template_path: Path) -> None:
    """Create a test template with a data validation dropdown.
    
    Args:
        template_path: Path where to save the template
    """
    try:
        import xlwings as xw
    except ImportError:
        logger.error("xlwings is required. Install it with: pip install xlwings")
        sys.exit(1)
    
    logger.info(f"Creating test template with dropdown at: {template_path}")
    
    # Create template directory if needed
    template_path.parent.mkdir(parents=True, exist_ok=True)
    
    # Start Excel
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    
    try:
        # Create a new workbook
        wb = app.books.add()
        
        # Add a sheet named "Daten"
        if wb.sheets.count > 0:
            sheet = wb.sheets[0]
            sheet.name = "Daten"
        else:
            sheet = wb.sheets.add("Daten")
        
        # Add some header text
        sheet.range("A1").value = "Test Template with Dropdown"
        sheet.range("A2").value = "Data Validation Cell:"
        sheet.range("B2").value = "Select an option"
        
        # Add a data validation dropdown to cell B2
        # Options: genügend, zu prüfen, nicht genügend
        validation = sheet.range("B2").api.Validation
        validation.Delete()  # Clear any existing validation
        validation.Add(
            Type=3,  # xlValidateList
            AlertStyle=1,  # xlValidAlertStop
            Operator=1,  # xlBetween (not used for lists)
            Formula1="genügend,zu prüfen,nicht genügend"
        )
        validation.IgnoreBlank = True
        validation.InCellDropdown = True
        
        # Add another cell with dropdown (for testing multiple dropdowns)
        sheet.range("A4").value = "Another Dropdown:"
        sheet.range("B4").value = "Option 1"
        validation2 = sheet.range("B4").api.Validation
        validation2.Delete()
        validation2.Add(
            Type=3,
            AlertStyle=1,
            Operator=1,
            Formula1="Option 1,Option 2,Option 3"
        )
        validation2.IgnoreBlank = True
        validation2.InCellDropdown = True
        
        # Add metadata cells (similar to real template)
        sheet.range("C4").value = "FinmaObjektName"
        sheet.range("C5").value = "FinmaID"
        sheet.range("C6").value = "Aufsichtskategorie"
        sheet.range("C7").value = "MitarbeiterName"
        
        # Save the template
        wb.save(str(template_path.absolute()))
        wb.close()
        
        logger.info(f"✓ Template created successfully with dropdowns in cells B2 and B4")
        
    finally:
        app.quit()

def create_dummy_source_file(source_path: Path) -> None:
    """Create a dummy source file for testing.
    
    Args:
        source_path: Path where to save the source file
    """
    try:
        import xlwings as xw
    except ImportError:
        logger.error("xlwings is required. Install it with: pip install xlwings")
        sys.exit(1)
    
    logger.info(f"Creating dummy source file at: {source_path}")
    
    # Create source directory if needed
    source_path.parent.mkdir(parents=True, exist_ok=True)
    
    # Start Excel
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    
    try:
        # Create a new workbook
        wb = app.books.add()
        sheet = wb.sheets[0]
        sheet.name = "Data"
        sheet.range("A1").value = "Dummy Source File"
        
        # Save the source file
        wb.save(str(source_path.absolute()))
        wb.close()
        
        logger.info("✓ Dummy source file created")
        
    finally:
        app.quit()

def generate_test_report(template_path: Path, source_path: Path, output_path: Path) -> bool:
    """Generate a test report using ExcelTemplateManager.
    
    Args:
        template_path: Path to the template file
        source_path: Path to the source file
        output_path: Path where to save the report
        
    Returns:
        True if report was generated successfully, False otherwise
    """
    try:
        from orsa_analysis.reporting.excel_template_manager import ExcelTemplateManager
    except ImportError:
        logger.error("Could not import ExcelTemplateManager. Make sure orsa_analysis is installed.")
        return False
    
    logger.info("Generating test report...")
    
    try:
        # Create template manager
        manager = ExcelTemplateManager(template_path)
        
        # Create output workbook
        manager.create_output_workbook(source_path)
        
        # Write some test values (including to the dropdown cells)
        manager.write_cell_value("Daten", "C4", "Test Institute Ltd.")
        manager.write_cell_value("Daten", "C5", "10001")
        manager.write_cell_value("Daten", "C6", "Kategorie 1")
        manager.write_cell_value("Daten", "C7", "John Doe")
        
        # Write value to dropdown cell (should preserve the dropdown)
        manager.write_cell_value("Daten", "B2", "genügend")
        manager.write_cell_value("Daten", "B4", "Option 2")
        
        # Save the workbook
        manager.save_workbook(output_path)
        
        # Clean up
        manager.close()
        
        logger.info(f"✓ Test report generated successfully: {output_path}")
        return True
        
    except Exception as e:
        logger.error(f"✗ Failed to generate report: {e}", exc_info=True)
        return False

def main():
    """Main verification workflow."""
    logger.info("=" * 60)
    logger.info("Excel Data Validation Dropdown Preservation Verification")
    logger.info("=" * 60)
    
    # Set up paths
    test_dir = Path("test_verification")
    template_path = test_dir / "template_with_dropdown.xlsx"
    source_path = test_dir / "dummy_source.xlsx"
    output_path = test_dir / "generated_report.xlsx"
    
    # Step 1: Create template with dropdown
    logger.info("\nStep 1: Creating template with data validation dropdown...")
    try:
        create_template_with_dropdown(template_path)
    except Exception as e:
        logger.error(f"✗ Failed to create template: {e}")
        logger.error("Make sure Microsoft Excel is installed and accessible")
        sys.exit(1)
    
    # Step 2: Create dummy source file
    logger.info("\nStep 2: Creating dummy source file...")
    try:
        create_dummy_source_file(source_path)
    except Exception as e:
        logger.error(f"✗ Failed to create source file: {e}")
        sys.exit(1)
    
    # Step 3: Generate report
    logger.info("\nStep 3: Generating report using ExcelTemplateManager...")
    success = generate_test_report(template_path, source_path, output_path)
    
    if not success:
        logger.error("\n✗ Report generation failed")
        sys.exit(1)
    
    # Step 4: Verification instructions
    logger.info("\n" + "=" * 60)
    logger.info("VERIFICATION INSTRUCTIONS")
    logger.info("=" * 60)
    logger.info(f"\n1. Open the generated report in Excel:")
    logger.info(f"   {output_path.absolute()}")
    logger.info(f"\n2. Navigate to the 'Daten' sheet")
    logger.info(f"\n3. Click on cell B2")
    logger.info(f"   - You should see a dropdown arrow on the right side of the cell")
    logger.info(f"   - Click the arrow to verify options: genügend, zu prüfen, nicht genügend")
    logger.info(f"\n4. Click on cell B4")
    logger.info(f"   - You should see a dropdown arrow")
    logger.info(f"   - Click the arrow to verify options: Option 1, Option 2, Option 3")
    logger.info(f"\n5. Verification Results:")
    logger.info(f"   ✓ If dropdowns are present and working: FIX IS WORKING")
    logger.info(f"   ✗ If dropdowns are missing: Issue still exists")
    logger.info("\n" + "=" * 60)
    logger.info("✓ Verification script completed successfully")
    logger.info("=" * 60)

if __name__ == "__main__":
    main()
