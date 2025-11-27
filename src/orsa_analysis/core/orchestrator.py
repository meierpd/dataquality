"""Orchestrator module for ORSA document quality control pipeline.

This module coordinates the entire workflow:
1. Document sourcing from ORSADocumentSourcer
2. File processing with hash-based caching
3. Quality check execution
4. Result storage to database
"""

import logging
from pathlib import Path
from typing import List, Tuple, Optional, Dict, Any
from datetime import datetime

from orsa_analysis.core.processor import DocumentProcessor
from orsa_analysis.core.database_manager import DatabaseManager, CheckResult

logger = logging.getLogger(__name__)


class ORSAPipeline:
    """Orchestrates the complete ORSA document quality control pipeline.
    
    This class coordinates:
    - Document sourcing from external systems
    - File processing with intelligent caching
    - Quality check execution
    - Database result storage
    
    Args:
        db_manager: DatabaseManager instance for result storage
        force_reprocess: If True, reprocess all files regardless of cache status
    
    Example:
        >>> db_manager = DatabaseManager(connection_string="mssql+pyodbc://...")
        >>> pipeline = ORSAPipeline(db_manager, force_reprocess=False)
        >>> results = pipeline.process_documents([("INST001", Path("file1.xlsx"))])
        >>> pipeline.generate_summary()
    """
    
    def __init__(self, db_manager: DatabaseManager, force_reprocess: bool = False):
        """Initialize the pipeline with database connection and processing options.
        
        Args:
            db_manager: DatabaseManager for storing results
            force_reprocess: Whether to reprocess already-seen files
        """
        self.db_manager = db_manager
        self.processor = DocumentProcessor(db_manager, force_reprocess=force_reprocess)
        self.processing_stats = {
            "files_processed": 0,
            "files_skipped": 0,
            "files_failed": 0,
            "checks_run": 0,
            "checks_passed": 0,
            "checks_failed": 0,
            "start_time": None,
            "end_time": None,
        }
        logger.info(
            f"Initialized ORSA pipeline (force_reprocess={force_reprocess})"
        )
    
    def process_documents(
        self, documents: List[Tuple[str, Path]]
    ) -> Dict[str, Any]:
        """Process a list of documents through the quality control pipeline.
        
        This method:
        1. Validates input documents exist
        2. Extracts institute IDs from filenames
        3. Processes each document through quality checks
        4. Handles caching based on file hashes
        5. Stores results in the database
        6. Returns processing statistics
        
        Args:
            documents: List of tuples (document_name, file_path)
        
        Returns:
            Dictionary containing:
                - files_processed: Number of files successfully processed
                - files_skipped: Number of files skipped (cached)
                - files_failed: Number of files that failed processing
                - total_checks: Total number of quality checks executed
                - checks_passed: Number of checks that passed
                - checks_failed: Number of checks that failed
                - pass_rate: Ratio of passed checks (0.0 to 1.0)
                - institutes: List of unique institute IDs processed
                - processing_time: Time taken in seconds
        
        Raises:
            FileNotFoundError: If a document file doesn't exist
            
        Example:
            >>> documents = [("INST001_report.xlsx", Path("data/file1.xlsx"))]
            >>> results = pipeline.process_documents(documents)
            >>> print(f"Processed {results['files_processed']} files")
        """
        self.processing_stats["start_time"] = datetime.now()
        institutes_seen = set()
        
        logger.info(f"Starting pipeline processing for {len(documents)} documents")
        
        for doc_name, file_path in documents:
            try:
                # Validate file exists
                if not file_path.exists():
                    logger.error(f"File not found: {file_path}")
                    self.processing_stats["files_failed"] += 1
                    continue
                
                # Extract institute ID from filename
                institute_id = self.processor._extract_institute_id(doc_name)
                institutes_seen.add(institute_id)
                
                logger.info(
                    f"Processing {doc_name} for institute {institute_id}"
                )
                
                # Check if file should be processed (caching logic)
                should_process, reason = self.processor.should_process_file(
                    institute_id, file_path
                )
                
                if not should_process:
                    logger.info(
                        f"Skipping {doc_name}: {reason}"
                    )
                    self.processing_stats["files_skipped"] += 1
                    continue
                
                # Process file through quality checks
                version_info, check_results = self.processor.process_file(
                    institute_id, file_path
                )
                
                # Update statistics
                self.processing_stats["files_processed"] += 1
                self.processing_stats["checks_run"] += len(check_results)
                
                # Count passed/failed checks
                for check_result in check_results:
                    if check_result.outcome_bool:
                        self.processing_stats["checks_passed"] += 1
                    else:
                        self.processing_stats["checks_failed"] += 1
                
                logger.info(
                    f"Completed {doc_name}: version {version_info.version_number}, "
                    f"{len(check_results)} checks run"
                )
                
            except Exception as e:
                logger.error(
                    f"Failed to process {doc_name}: {str(e)}", exc_info=True
                )
                self.processing_stats["files_failed"] += 1
        
        self.processing_stats["end_time"] = datetime.now()
        self.processing_stats["institutes"] = sorted(list(institutes_seen))
        
        # Calculate processing time
        processing_time = (
            self.processing_stats["end_time"] - self.processing_stats["start_time"]
        ).total_seconds()
        
        # Calculate pass rate
        total_checks = self.processing_stats["checks_run"]
        checks_passed = self.processing_stats["checks_passed"]
        pass_rate = checks_passed / total_checks if total_checks > 0 else 0.0
        
        summary = {
            "files_processed": self.processing_stats["files_processed"],
            "files_skipped": self.processing_stats["files_skipped"],
            "files_failed": self.processing_stats["files_failed"],
            "total_checks": total_checks,
            "checks_passed": checks_passed,
            "checks_failed": self.processing_stats["checks_failed"],
            "pass_rate": pass_rate,
            "institutes": self.processing_stats["institutes"],
            "processing_time": processing_time,
        }
        
        logger.info(
            f"Pipeline completed: {summary['files_processed']} files processed, "
            f"{summary['files_skipped']} skipped, {summary['files_failed']} failed "
            f"in {processing_time:.2f}s"
        )
        
        return summary
    
    def process_from_sourcer(self, sourcer: Any) -> Dict[str, Any]:
        """Process documents directly from an ORSADocumentSourcer.
        
        This is a convenience method that:
        1. Calls sourcer.load() to download documents
        2. Processes the downloaded files
        3. Returns processing statistics
        
        Args:
            sourcer: ORSADocumentSourcer instance with load() method
        
        Returns:
            Processing summary dictionary (see process_documents)
        
        Example:
            >>> from orsa_analysis.sourcing import ORSADocumentSourcer
            >>> sourcer = ORSADocumentSourcer()
            >>> results = pipeline.process_from_sourcer(sourcer)
        """
        logger.info("Loading documents from ORSADocumentSourcer")
        documents = sourcer.load()
        logger.info(f"Retrieved {len(documents)} documents from sourcer")
        return self.process_documents(documents)
    
    def generate_summary(self) -> Dict[str, Any]:
        """Generate a comprehensive summary of pipeline execution.
        
        Returns:
            Dictionary with detailed statistics including:
                - Overall processing statistics
                - Database storage info
                - Cache performance metrics
        """
        processor_summary = self.processor.get_processing_summary()
        
        return {
            **processor_summary,
            "pipeline_stats": self.processing_stats,
        }
    
    def close(self):
        """Close database connections and clean up resources.
        
        Should be called when done with the pipeline to ensure
        proper cleanup of database connections.
        """
        logger.info("Closing pipeline and database connections")
        self.db_manager.close()


class CachedDocumentProcessor:
    """Processor with explicit caching control for document processing.
    
    This class provides fine-grained control over the caching behavior,
    allowing manual cache invalidation and inspection.
    
    Args:
        db_manager: DatabaseManager for result storage
        cache_enabled: Whether caching is enabled (default True)
    """
    
    def __init__(self, db_manager: DatabaseManager, cache_enabled: bool = True):
        """Initialize processor with caching control."""
        self.db_manager = db_manager
        self.cache_enabled = cache_enabled
        self.processor = DocumentProcessor(
            db_manager, force_reprocess=not cache_enabled
        )
        logger.info(f"Initialized cached processor (cache_enabled={cache_enabled})")
    
    def get_cache_status(self, institute_id: str, file_path: Path) -> Dict[str, Any]:
        """Get detailed cache status for a specific file.
        
        Args:
            institute_id: Institute identifier
            file_path: Path to the document file
        
        Returns:
            Dictionary containing:
                - is_cached: Whether file is in cache
                - file_hash: SHA-256 hash of the file
                - version_number: Current version number (if cached)
                - last_processed: Timestamp of last processing (if cached)
        """
        file_hash = self.processor.version_manager.compute_file_hash(file_path)
        is_cached = self.processor.version_manager.is_processed(
            institute_id, file_hash
        )
        
        cache_info = {
            "is_cached": is_cached,
            "file_hash": file_hash,
        }
        
        if is_cached:
            # Get version from cache
            version_cache = self.processor.version_manager._version_cache.get(
                institute_id, {}
            )
            version_number = version_cache.get(file_hash)
            if version_number:
                cache_info["version_number"] = version_number
        
        return cache_info
    
    def invalidate_cache(self, institute_id: Optional[str] = None):
        """Invalidate cache for specific institute or all institutes.
        
        Args:
            institute_id: If provided, only invalidate cache for this institute.
                         If None, invalidate entire cache.
        """
        if institute_id:
            # Invalidate specific institute
            if institute_id in self.processor.version_manager._version_cache:
                del self.processor.version_manager._version_cache[institute_id]
                logger.info(f"Invalidated cache for institute {institute_id}")
        else:
            # Invalidate all
            self.processor.version_manager._version_cache.clear()
            logger.info("Invalidated entire cache")
    
    def get_cache_statistics(self) -> Dict[str, Any]:
        """Get statistics about the current cache state.
        
        Returns:
            Dictionary with cache metrics:
                - total_institutes: Number of institutes in cache
                - total_versions: Total number of cached versions
                - institutes: List of institute IDs in cache
        """
        cache = self.processor.version_manager._version_cache
        
        total_versions = sum(len(hashes) for hashes in cache.values())
        
        return {
            "total_institutes": len(cache),
            "total_versions": total_versions,
            "institutes": sorted(list(cache.keys())),
        }
