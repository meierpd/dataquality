"""Main processing orchestration with caching."""

from pathlib import Path
from typing import List, Tuple, Optional
from datetime import datetime
import logging

from orsa_analysis.core.reader import ExcelReader
from orsa_analysis.core.versioning import VersionManager, FileVersion
from orsa_analysis.core.database_manager import CheckResult, DatabaseManager
from orsa_analysis.checks.rules import get_all_checks, run_check

logger = logging.getLogger(__name__)


class DocumentProcessor:
    """Orchestrates the processing of Excel documents with caching."""

    def __init__(
        self,
        db_manager: DatabaseManager,
        force_reprocess: bool = False,
    ):
        """Initialize the document processor.

        Args:
            db_manager: Database manager for storing results
            force_reprocess: If True, reprocess files even if hash exists
        """
        self.reader = ExcelReader(data_only=True, read_only=True)
        self.version_manager = VersionManager()
        self.db_manager = db_manager
        self.force_reprocess = force_reprocess
        self._load_existing_versions()

    def _load_existing_versions(self) -> None:
        """Load existing version information from database."""
        existing_versions = self.db_manager.get_existing_versions()
        self.version_manager.load_existing_versions(existing_versions)
        logger.info(f"Loaded {len(existing_versions)} existing versions from database")

    def should_process_file(
        self, institute_id: str, file_path: Path
    ) -> Tuple[bool, str]:
        """Determine if a file should be processed.

        Args:
            institute_id: Identifier for the institute
            file_path: Path to the file

        Returns:
            Tuple of (should_process, reason)
        """
        if self.force_reprocess:
            return True, "Force reprocess mode enabled"

        file_hash = self.version_manager.compute_file_hash(file_path)
        if self.version_manager.is_processed(institute_id, file_hash):
            return False, f"File hash {file_hash[:8]}... already processed"

        return True, "New file hash"

    def process_file(
        self, institute_id: str, file_path: Path, geschaeft_nr: Optional[str] = None, berichtsjahr: Optional[int] = None
    ) -> Tuple[FileVersion, List[CheckResult]]:
        """Process a single Excel file and run all checks.

        Args:
            institute_id: Identifier for the institute
            file_path: Path to the Excel file
            geschaeft_nr: Optional business case number (Geschäftsnummer)
            berichtsjahr: Optional reporting year (Berichtsjahr)

        Returns:
            Tuple of (FileVersion, List of CheckResults)

        Raises:
            FileNotFoundError: If file doesn't exist
            ValueError: If file is not valid Excel format
        """
        logger.info(f"Processing file: {file_path.name} for institute: {institute_id}")

        version_info = self.version_manager.get_version(institute_id, file_path)
        workbook = None
        results = []

        try:
            workbook = self.reader.load_file(file_path)
            processed_at = datetime.now()

            all_checks = get_all_checks()
            logger.info(f"Running {len(all_checks)} checks on {file_path.name}")

            for check_name, check_function in all_checks:
                outcome, numeric_value, description = run_check(
                    check_name, check_function, workbook
                )

                result = CheckResult(
                    institute_id=institute_id,
                    file_name=version_info.file_name,
                    file_hash=version_info.file_hash,
                    version_number=version_info.version_number,
                    check_name=check_name,
                    check_description=description,
                    outcome_bool=outcome,
                    outcome_numeric=numeric_value,
                    processed_at=processed_at,
                    geschaeft_nr=geschaeft_nr,
                    berichtsjahr=berichtsjahr,
                )
                results.append(result)

                log_level = logging.INFO if outcome else logging.WARNING
                logger.log(
                    log_level,
                    f"Check '{check_name}': {'PASS' if outcome else 'FAIL'} - {description}",
                )

            logger.info(
                f"Completed processing {file_path.name}: "
                f"{sum(1 for r in results if r.outcome_bool)}/{len(results)} checks passed"
            )

            # Write results to database
            self.db_manager.write_results(results)

        finally:
            if workbook:
                self.reader.close_workbook(workbook)

        return version_info, results

    def process_documents(
        self, documents: List[Tuple[str, Path]]
    ) -> List[Tuple[str, FileVersion, List[CheckResult]]]:
        """Process multiple documents from ORSADocumentSourcer.

        Args:
            documents: List of tuples (file_name, file_path) from ORSADocumentSourcer.load()

        Returns:
            List of tuples (institute_id, FileVersion, List[CheckResult])
        """
        all_results = []
        processed_count = 0
        skipped_count = 0

        logger.info(f"Starting batch processing of {len(documents)} documents")

        for file_name, file_path in documents:
            institute_id = self._extract_institute_id(file_name)

            should_process, reason = self.should_process_file(institute_id, file_path)

            if not should_process:
                logger.info(f"Skipping {file_name}: {reason}")
                skipped_count += 1
                continue

            try:
                version_info, results = self.process_file(institute_id, file_path)
                all_results.append((institute_id, version_info, results))
                processed_count += 1
            except Exception as e:
                logger.error(f"Failed to process {file_name}: {e}")
                continue

        logger.info(
            f"Batch processing complete: {processed_count} processed, {skipped_count} skipped"
        )
        return all_results

    def _extract_institute_id(self, file_name: str) -> str:
        """Extract institute ID from file name.

        This is a simple implementation that uses the first part of the filename
        before any underscore or dash. Override this method for custom logic.

        Args:
            file_name: Name of the file

        Returns:
            Institute identifier
        """
        base_name = Path(file_name).stem
        for separator in ["_", "-", " "]:
            if separator in base_name:
                return base_name.split(separator)[0]
        return base_name

    def get_processing_summary(self) -> dict:
        """Get summary of processing statistics.

        Returns:
            Dictionary with processing statistics
        """
        # Note: This method requires stored results from the database
        # For now, return basic statistics from version manager
        all_versions = self.version_manager._version_cache
        if not all_versions:
            return {
                "total_files": 0,
                "total_checks": 0,
                "checks_passed": 0,
                "checks_failed": 0,
                "institutes": [],
            }

        institutes = set(all_versions.keys())

        total_files = sum(len(hashes) for hashes in all_versions.values())

        return {
            "total_files": total_files,
            "total_checks": 0,
            "checks_passed": 0,
            "checks_failed": 0,
            "institutes": sorted(list(institutes)),
            "pass_rate": "N/A",
        }
