"""Main report generation orchestrator."""

import logging
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from orsa_analysis.core.database_manager import DatabaseManager
from orsa_analysis.reporting.excel_template_manager import ExcelTemplateManager
from orsa_analysis.reporting.check_to_cell_mapper import CheckToCellMapper
from orsa_analysis.reporting.sharepoint_uploader import SharePointUploader


logger = logging.getLogger(__name__)


class ReportGenerator:
    """Generate Excel reports from check results stored in database."""
    
    def __init__(self, 
                 db_manager: DatabaseManager,
                 template_path: Path,
                 output_dir: Path,
                 check_mapper: Optional[CheckToCellMapper] = None,
                 enable_upload: bool = False,
                 download_links: Optional[Dict[str, str]] = None):
        """Initialize report generator.
        
        Args:
            db_manager: Database manager for querying results
            template_path: Path to template Excel file
            output_dir: Directory for output files
            check_mapper: Optional custom check mapper. If None, uses default.
            enable_upload: If True, upload reports to SharePoint (default: False)
            download_links: Optional mapping of institute_id -> download_link for uploads
        """
        self.db_manager = db_manager
        self.template_manager = ExcelTemplateManager(template_path)
        self.output_dir = Path(output_dir)
        self.check_mapper = check_mapper if check_mapper is not None else CheckToCellMapper()
        self.enable_upload = enable_upload
        self.download_links = download_links or {}
        self.uploader = None
        
        # Initialize SharePoint uploader if upload is enabled
        if self.enable_upload:
            try:
                self.uploader = SharePointUploader()
                logger.info("SharePoint upload enabled")
            except Exception as e:
                logger.warning(f"Failed to initialize SharePoint uploader: {e}")
                logger.warning("Reports will be generated but not uploaded")
                self.enable_upload = False
        
        # Ensure output directory exists
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        logger.info(f"ReportGenerator initialized")
        logger.info(f"  Template: {template_path}")
        logger.info(f"  Output dir: {output_dir}")
        logger.info(f"  Upload enabled: {self.enable_upload}")
    
    def generate_report(self, 
                       institute_id: str,
                       source_file_path: Optional[Path] = None) -> Optional[Path]:
        """Generate report for a single institute.
        
        Local report files are always overwritten if they exist.
        SharePoint uploads skip if file already exists (see enable_upload parameter).
        
        Args:
            institute_id: Institute identifier (FinmaID)
            source_file_path: Optional path to source ORSA file
            
        Returns:
            Path to generated report file, or None if generation failed
        """
        logger.info(f"Generating report for institute: {institute_id}")
        
        # Get check results from database
        check_results = self.db_manager.get_latest_results_for_institute(institute_id)
        if not check_results:
            logger.warning(f"No check results found for institute {institute_id}")
            return None
        
        logger.info(f"Found {len(check_results)} check results for {institute_id}")
        
        # Validate source file path
        if not source_file_path or not source_file_path.exists():
            logger.error(f"Source file required and must exist: {source_file_path}")
            return None
        
        # Determine output file path
        output_path = self._get_output_path(institute_id, source_file_path)
        
        # Always overwrite local reports if they exist
        if output_path.exists():
            logger.info(f"Report already exists locally, will overwrite: {output_path}")
        
        # Create standalone output workbook from template
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
        try:
            self.template_manager.save_workbook(output_path)
            logger.info(f"Report saved successfully: {output_path}")
        except Exception as e:
            logger.error(f"Failed to save report to {output_path}: {e}")
            # Clean up before re-raising
            self.template_manager.close()
            raise
        
        # Clean up
        self.template_manager.close()
        
        logger.info(f"Report generated successfully: {output_path}")
        
        # Upload to SharePoint if enabled
        if self.enable_upload and self.uploader:
            try:
                self._upload_report(institute_id, output_path)
            except Exception as e:
                logger.error(f"Failed to upload report for {institute_id}: {e}", exc_info=True)
                # Don't fail the entire report generation if upload fails
        
        return output_path
    
    def generate_all_reports(self, 
                           source_files: Optional[Dict[str, Path]] = None) -> List[Path]:
        """Generate reports for all institutes with check results.
        
        Local report files are always overwritten if they exist.
        SharePoint uploads skip if file already exists (see enable_upload parameter).
        
        Args:
            source_files: Optional mapping of institute_id -> source file path
            
        Returns:
            List of paths to generated report files
        """
        institutes = self.db_manager.get_all_institutes_with_results()
        logger.info(f"Generating reports for {len(institutes)} institutes")
        
        generated_reports = []
        failed_reports = []
        
        for institute_id in institutes:
            logger.info(f"Processing {institute_id}...")
            
            # Get source file path if provided
            source_path = None
            if source_files and institute_id in source_files:
                source_path = source_files[institute_id]
            
            # Generate report with error handling
            try:
                report_path = self.generate_report(
                    institute_id, 
                    source_path
                )
                
                if report_path:
                    generated_reports.append(report_path)
                    logger.info(f"✓ Report generated successfully for {institute_id}")
                else:
                    failed_reports.append({
                        'institute_id': institute_id,
                        'error': 'Report generation returned None'
                    })
                    logger.warning(f"✗ Report generation returned None for {institute_id}")
            except Exception as e:
                failed_reports.append({
                    'institute_id': institute_id,
                    'error': str(e),
                    'error_type': type(e).__name__
                })
                logger.error(
                    f"✗ Report generation failed for {institute_id}: {type(e).__name__}: {e}",
                    exc_info=True
                )
        
        # Log comprehensive summary
        logger.info("=" * 60)
        logger.info(f"Report generation complete:")
        logger.info(f"  Total institutes: {len(institutes)}")
        logger.info(f"  Successfully generated: {len(generated_reports)}")
        logger.info(f"  Failed: {len(failed_reports)}")
        
        if failed_reports:
            logger.warning("Failed reports:")
            for failed in failed_reports:
                logger.warning(
                    f"  - {failed['institute_id']}: {failed.get('error_type', 'Error')}: {failed['error']}"
                )
        logger.info("=" * 60)
        
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
            
            # Get cell mapping (sheet_name, outcome_cell, value_type, description_cell)
            mapping = self.check_mapper.get_cell_location(check_name)
            if not mapping:
                continue
            
            sheet_name, cell_address, value_type, description_cell = mapping
            
            # Get and format value
            try:
                value = self.check_mapper.get_value_from_result(
                    result, sheet_name, cell_address, value_type
                )
                
                # Write outcome value to cell
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
                
                # Write description to cell if mapping exists
                if description_cell:
                    description = self.check_mapper.get_value_from_result(
                        result, sheet_name, description_cell, "check_description"
                    )
                    
                    if description:
                        desc_success = self.template_manager.write_cell_value(
                            sheet_name,
                            description_cell,
                            description
                        )
                        
                        if desc_success:
                            logger.debug(
                                f"Applied {check_name} description -> "
                                f"{sheet_name}!{description_cell} = {description}"
                            )
                        
            except Exception as e:
                logger.error(f"Failed to apply check {check_name}: {e}")
        
        return applied_count
    
    def _apply_institut_metadata(self, institute_id: str) -> bool:
        """Apply institut metadata to workbook cells.
        
        Writes institute metadata to cells on the Auswertung sheet:
        - E2: FinmaObjektName
        - E3: FinmaID
        - E4: Aufsichtskategorie
        - E6: MitarbeiterName
        
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
                ("E2", "FinmaObjektName", "FinmaObjektName"),
                ("E3", "FINMAID", "FinmaID"),
                ("E4", "Aufsichtskategorie", "Aufsichtskategorie"),
                ("E6", "MitarbeiterName", "MitarbeiterName")
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
            
            logger.info(f"Applied {success_count}/4 institut metadata fields to report")
            return success_count == 4
            
        except Exception as e:
            logger.error(f"Failed to apply institut metadata: {e}")
            return False
    
    def _upload_report(self, institute_id: str, report_path: Path) -> bool:
        """Upload report to SharePoint.
        
        Args:
            institute_id: Institute identifier
            report_path: Path to generated report file
            
        Returns:
            True if upload was successful or skipped, False otherwise
        """
        if institute_id not in self.download_links:
            logger.warning(
                f"No download link found for institute {institute_id}. "
                f"Report will not be uploaded."
            )
            return False
        
        download_link = self.download_links[institute_id]
        
        try:
            logger.info(f"Uploading report for {institute_id} to SharePoint...")
            result = self.uploader.upload(
                download_link=download_link,
                filepath=str(report_path),
                skip_if_exists=True
            )
            
            if result["success"]:
                if result["skipped"]:
                    logger.info(f"Report already exists on SharePoint: {report_path.name}")
                else:
                    logger.info(f"✓ Report uploaded successfully: {report_path.name}")
                return True
            else:
                logger.error(f"✗ Upload failed: {result['message']}")
                return False
                
        except Exception as e:
            logger.error(f"Upload failed with exception: {e}", exc_info=True)
            return False
    
    def _get_output_path(self, institute_id: str, source_file_path: Path) -> Path:
        """Determine output file path for a report.
        
        Args:
            institute_id: Institute identifier
            source_file_path: Path to the source ORSA file
            
        Returns:
            Path for output file with format: Auswertung_{institute_id}_{source_file_name}.xlsx
        """
        # Extract source file name (without path)
        source_file_name = source_file_path.name
        
        # Build filename: Auswertung_{institute_id}_{source_file_name}
        filename = f"Auswertung_{institute_id}_{source_file_name}"
        
        return self.output_dir / filename
