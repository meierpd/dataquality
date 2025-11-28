"""Main report generation orchestrator."""

import logging
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from orsa_analysis.core.database_manager import DatabaseManager
from orsa_analysis.reporting.excel_template_manager import ExcelTemplateManager
from orsa_analysis.reporting.check_to_cell_mapper import CheckToCellMapper


logger = logging.getLogger(__name__)


class ReportGenerator:
    """Generate Excel reports from check results stored in database."""
    
    def __init__(self, 
                 db_manager: DatabaseManager,
                 template_path: Path,
                 output_dir: Path,
                 check_mapper: Optional[CheckToCellMapper] = None):
        """Initialize report generator.
        
        Args:
            db_manager: Database manager for querying results
            template_path: Path to template Excel file
            output_dir: Directory for output files
            check_mapper: Optional custom check mapper. If None, uses default.
        """
        self.db_manager = db_manager
        self.template_manager = ExcelTemplateManager(template_path)
        self.output_dir = Path(output_dir)
        self.check_mapper = check_mapper if check_mapper is not None else CheckToCellMapper()
        
        # Ensure output directory exists
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        logger.info(f"ReportGenerator initialized")
        logger.info(f"  Template: {template_path}")
        logger.info(f"  Output dir: {output_dir}")
    
    def generate_report(self, 
                       institute_id: str,
                       source_file_path: Optional[Path] = None,
                       force_overwrite: bool = False) -> Optional[Path]:
        """Generate report for a single institute.
        
        Args:
            institute_id: Institute identifier (FinmaID)
            source_file_path: Optional path to source ORSA file
            force_overwrite: If True, overwrite existing report file
            
        Returns:
            Path to generated report file, or None if generation failed or skipped
        """
        logger.info(f"Generating report for institute: {institute_id}")
        
        # Get check results from database
        check_results = self.db_manager.get_latest_results_for_institute(institute_id)
        if not check_results:
            logger.warning(f"No check results found for institute {institute_id}")
            return None
        
        logger.info(f"Found {len(check_results)} check results for {institute_id}")
        
        # Get metadata for file naming
        metadata = self.db_manager.get_institute_metadata(institute_id)
        if not metadata:
            logger.error(f"Could not retrieve metadata for {institute_id}")
            return None
        
        # Determine output file path
        output_path = self._get_output_path(institute_id, metadata)
        
        # Check if file already exists
        if output_path.exists() and not force_overwrite:
            logger.info(f"Report already exists: {output_path}")
            logger.info(f"Skipping generation (use force_overwrite=True to regenerate)")
            return None
        
        # Validate source file path
        if not source_file_path or not source_file_path.exists():
            logger.error(f"Source file required and must exist: {source_file_path}")
            return None
        
        # Create output workbook (source file as base, template sheets prepended)
        try:
            self.template_manager.create_output_workbook(source_file_path)
        except Exception as e:
            logger.error(f"Failed to create output workbook: {e}")
            return None
        
        # Apply check results to output
        applied_count = self._apply_check_results(check_results)
        logger.info(f"Applied {applied_count} check results to report")
        
        # Apply institut metadata to output (FinmaID, FinmaObjektName, MitarbeiterName)
        self._apply_institut_metadata(institute_id)
        
        # Save output file
        self.template_manager.save_workbook(output_path)
        
        # Clean up
        self.template_manager.close()
        
        logger.info(f"Report generated successfully: {output_path}")
        return output_path
    
    def generate_all_reports(self, 
                           source_files: Optional[Dict[str, Path]] = None,
                           force_overwrite: bool = False) -> List[Path]:
        """Generate reports for all institutes with check results.
        
        Args:
            source_files: Optional mapping of institute_id -> source file path
            force_overwrite: If True, overwrite existing report files
            
        Returns:
            List of paths to generated report files
        """
        institutes = self.db_manager.get_all_institutes_with_results()
        logger.info(f"Generating reports for {len(institutes)} institutes")
        
        generated_reports = []
        skipped_count = 0
        error_count = 0
        
        for institute_id in institutes:
            logger.info(f"Processing {institute_id}...")
            
            # Get source file path if provided
            source_path = None
            if source_files and institute_id in source_files:
                source_path = source_files[institute_id]
            
            # Generate report
            report_path = self.generate_report(
                institute_id, 
                source_path,
                force_overwrite
            )
            
            if report_path:
                generated_reports.append(report_path)
            elif self._report_exists(institute_id):
                skipped_count += 1
            else:
                error_count += 1
        
        logger.info(f"Report generation complete:")
        logger.info(f"  Generated: {len(generated_reports)}")
        logger.info(f"  Skipped (already exists): {skipped_count}")
        logger.info(f"  Errors: {error_count}")
        
        return generated_reports
    
    def get_institutes_with_results(self) -> List[str]:
        """Get list of institutes that have check results.
        
        Returns:
            List of institute IDs
        """
        return self.db_manager.get_all_institutes_with_results()
    
    def _apply_check_results(self, check_results: List[Dict]) -> int:
        """Apply check results to workbook cells.
        
        Args:
            check_results: List of check result dictionaries
            
        Returns:
            Number of check results successfully applied
        """
        applied_count = 0
        
        for result in check_results:
            check_name = result.get('check_name')
            
            # Check if we have a mapping for this check
            if not self.check_mapper.has_mapping(check_name):
                logger.debug(f"No mapping for check: {check_name}")
                continue
            
            # Get cell mapping (sheet_name, cell_address, value_type)
            mapping = self.check_mapper.get_cell_location(check_name)
            if not mapping:
                continue
            
            sheet_name, cell_address, value_type = mapping
            
            # Get and format value
            try:
                value = self.check_mapper.get_value_from_result(
                    result, sheet_name, cell_address, value_type
                )
                
                # Write to cell
                success = self.template_manager.write_cell_value(
                    sheet_name,
                    cell_address,
                    value
                )
                
                if success:
                    applied_count += 1
                    logger.debug(
                        f"Applied {check_name} -> "
                        f"{sheet_name}!{cell_address} = {value}"
                    )
            except Exception as e:
                logger.error(f"Failed to apply check {check_name}: {e}")
        
        return applied_count
    
    def _apply_institut_metadata(self, institute_id: str) -> bool:
        """Apply institut metadata to workbook cells.
        
        Writes institute metadata to cells C3, C4, and C5 on the Auswertung sheet:
        - C3: FinmaID
        - C4: FinmaObjektName
        - C5: MitarbeiterName
        
        Args:
            institute_id: Institute identifier (FinmaID)
            
        Returns:
            True if metadata was successfully applied, False otherwise
        """
        try:
            # Fetch institut metadata from database
            institut_metadata = self.db_manager.get_institut_metadata_by_finmaid(institute_id)
            
            if not institut_metadata:
                logger.warning(
                    f"No institut metadata found for {institute_id}. "
                    f"Metadata cells will be empty."
                )
                return False
            
            # Define the target sheet and cell mappings
            sheet_name = "Auswertung"
            metadata_mappings = [
                ("C3", "FINMAID", "FinmaID"),
                ("C4", "FinmaObjektName", "FinmaObjektName"),
                ("C5", "MitarbeiterName", "MitarbeiterName")
            ]
            
            # Write each metadata field to its corresponding cell
            success_count = 0
            for cell_address, field_key, field_name in metadata_mappings:
                value = institut_metadata.get(field_key)
                
                if value is not None:
                    success = self.template_manager.write_cell_value(
                        sheet_name,
                        cell_address,
                        value
                    )
                    
                    if success:
                        success_count += 1
                        logger.debug(
                            f"Applied {field_name} -> "
                            f"{sheet_name}!{cell_address} = {value}"
                        )
                    else:
                        logger.warning(
                            f"Failed to write {field_name} to {sheet_name}!{cell_address}"
                        )
                else:
                    logger.warning(f"{field_name} is None for institute {institute_id}")
            
            logger.info(f"Applied {success_count}/3 institut metadata fields to report")
            return success_count == 3
            
        except Exception as e:
            logger.error(f"Failed to apply institut metadata: {e}")
            return False
    
    def _get_output_path(self, institute_id: str, metadata: Dict) -> Path:
        """Determine output file path for a report.
        
        Args:
            institute_id: Institute identifier
            metadata: Institute metadata dictionary
            
        Returns:
            Path for output file
        """
        # Extract version and other metadata
        version = metadata.get('version', 1)
        
        # Build filename: {institute_id}_ORSA_Report[_v{version}].xlsx
        # Only include version suffix for v2+, omit for v1
        if version == 1:
            filename = f"{institute_id}_ORSA_Report.xlsx"
        else:
            filename = f"{institute_id}_ORSA_Report_v{version}.xlsx"
        
        return self.output_dir / filename
    
    def _report_exists(self, institute_id: str) -> bool:
        """Check if a report already exists for an institute.
        
        Args:
            institute_id: Institute identifier
            
        Returns:
            True if report exists, False otherwise
        """
        metadata = self.db_manager.get_institute_metadata(institute_id)
        if not metadata:
            return False
        
        output_path = self._get_output_path(institute_id, metadata)
        return output_path.exists()
