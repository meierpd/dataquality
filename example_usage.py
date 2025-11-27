#!/usr/bin/env python3
"""Example usage of the ORSA Quality Control system with ORSADocumentSourcer.

This script demonstrates the complete workflow:
1. Source documents from the FINMA database using ORSADocumentSourcer
2. Process files through quality checks
3. Store results in the database
4. Generate summary reports

Requirements:
- credentials.env file with FINMA_USERNAME and FINMA_PASSWORD
- sql/source_orsa_dokument_metadata.sql query file
- Database connection configured (or use InMemoryDatabaseWriter for testing)
"""

import logging
from pathlib import Path
from typing import List, Tuple

from orsa_analysis import DocumentProcessor
from orsa_analysis.core.db import InMemoryDatabaseWriter, DatabaseWriter
from orsa_analysis.sourcing import ORSADocumentSourcer

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


def example_with_in_memory_db():
    """Example using in-memory database (for testing/demo)."""
    logger.info("=" * 60)
    logger.info("Example 1: Using In-Memory Database")
    logger.info("=" * 60)
    
    # Initialize in-memory database writer
    db_writer = InMemoryDatabaseWriter()
    logger.info("✓ Initialized in-memory database writer")
    
    # Initialize document processor
    processor = DocumentProcessor(db_writer, force_reprocess=False)
    logger.info("✓ Initialized document processor")
    
    # Source documents from FINMA database
    logger.info("\nSourcing documents from FINMA database...")
    try:
        sourcer = ORSADocumentSourcer()
        documents = sourcer.load()
        logger.info(f"✓ Retrieved {len(documents)} documents")
    except FileNotFoundError as e:
        logger.warning(f"⚠ Could not source documents: {e}")
        logger.info("  Using sample files instead for demonstration...")
        documents = get_sample_documents()
    
    # Process all documents
    logger.info("\nProcessing documents...")
    for name, path in documents:
        logger.info(f"  Processing: {name}")
        try:
            results = processor.process_file(path, institute_id="INS001")
            logger.info(f"    ✓ {len(results)} checks completed")
        except Exception as e:
            logger.error(f"    ✗ Failed: {e}")
    
    # Get and display processing summary
    logger.info("\nProcessing Summary:")
    summary = processor.get_processing_summary()
    logger.info(f"  Total files: {summary['total_files']}")
    logger.info(f"  Processed: {summary['processed']}")
    logger.info(f"  Cached: {summary['cached']}")
    logger.info(f"  Pass rate: {summary['pass_rate']:.1%}")
    
    # Access stored results
    all_results = db_writer.get_results()
    logger.info(f"\n✓ Stored {len(all_results)} check results in database")
    
    # Show some example results
    if all_results:
        logger.info("\nSample Check Results:")
        for result in all_results[:5]:  # Show first 5
            status = "✓ PASS" if result.outcome_bool else "✗ FAIL"
            logger.info(f"  {status} - {result.check_name}")
            logger.info(f"           {result.check_description}")


def example_with_mssql_db():
    """Example using actual MSSQL database connection."""
    logger.info("\n" + "=" * 60)
    logger.info("Example 2: Using MSSQL Database")
    logger.info("=" * 60)
    
    # Initialize database writer with connection string
    # Adjust connection string according to your setup
    connection_string = (
        "mssql+pymssql://username:password@server/database"
    )
    
    try:
        db_writer = DatabaseWriter(
            connection_string=connection_string,
            table_name="qc_results"
        )
        logger.info("✓ Connected to MSSQL database")
    except Exception as e:
        logger.error(f"✗ Failed to connect to database: {e}")
        logger.info("  Using in-memory database as fallback")
        db_writer = InMemoryDatabaseWriter()
    
    # Initialize processor
    processor = DocumentProcessor(db_writer, force_reprocess=False)
    
    # Source and process documents
    sourcer = ORSADocumentSourcer()
    documents = sourcer.load()
    
    for name, path in documents:
        results = processor.process_file(path, institute_id="INS001")
        logger.info(f"Processed {name}: {len(results)} checks")
    
    # Write results to database
    db_writer.write_results()
    logger.info("✓ Results written to MSSQL database")


def example_with_force_reprocess():
    """Example demonstrating force reprocess mode."""
    logger.info("\n" + "=" * 60)
    logger.info("Example 3: Force Reprocess Mode")
    logger.info("=" * 60)
    
    db_writer = InMemoryDatabaseWriter()
    
    # First pass: normal processing
    logger.info("\nFirst pass (normal processing):")
    processor = DocumentProcessor(db_writer, force_reprocess=False)
    
    try:
        sourcer = ORSADocumentSourcer()
        documents = sourcer.load()
    except Exception:
        documents = get_sample_documents()
    
    for name, path in documents[:1]:  # Process just one file
        results = processor.process_file(path, institute_id="INS001")
        logger.info(f"  Processed {name}: {len(results)} checks")
    
    summary1 = processor.get_processing_summary()
    logger.info(f"  Processed: {summary1['processed']}, Cached: {summary1['cached']}")
    
    # Second pass: with caching (should skip)
    logger.info("\nSecond pass (with caching):")
    processor2 = DocumentProcessor(db_writer, force_reprocess=False)
    
    for name, path in documents[:1]:
        results = processor2.process_file(path, institute_id="INS001")
        logger.info(f"  Processed {name}: {len(results)} checks")
    
    summary2 = processor2.get_processing_summary()
    logger.info(f"  Processed: {summary2['processed']}, Cached: {summary2['cached']}")
    
    # Third pass: force reprocess
    logger.info("\nThird pass (force reprocess):")
    processor3 = DocumentProcessor(db_writer, force_reprocess=True)
    
    for name, path in documents[:1]:
        results = processor3.process_file(path, institute_id="INS001")
        logger.info(f"  Processed {name}: {len(results)} checks")
    
    summary3 = processor3.get_processing_summary()
    logger.info(f"  Processed: {summary3['processed']}, Cached: {summary3['cached']}")


def example_custom_checks():
    """Example showing how to access individual check results."""
    logger.info("\n" + "=" * 60)
    logger.info("Example 4: Analyzing Individual Check Results")
    logger.info("=" * 60)
    
    db_writer = InMemoryDatabaseWriter()
    processor = DocumentProcessor(db_writer, force_reprocess=False)
    
    try:
        sourcer = ORSADocumentSourcer()
        documents = sourcer.load()
    except Exception:
        documents = get_sample_documents()
    
    # Process one file
    if documents:
        name, path = documents[0]
        logger.info(f"\nProcessing: {name}")
        results = processor.process_file(path, institute_id="INS001")
        
        # Analyze results by outcome
        passed = [r for r in results if r.outcome_bool]
        failed = [r for r in results if not r.outcome_bool]
        
        logger.info(f"\n✓ Passed checks: {len(passed)}")
        for result in passed:
            logger.info(f"  - {result.check_name}")
        
        logger.info(f"\n✗ Failed checks: {len(failed)}")
        for result in failed:
            logger.info(f"  - {result.check_name}: {result.check_description}")
        
        # Show checks with numeric values
        numeric_results = [r for r in results if r.outcome_numeric is not None]
        logger.info(f"\nChecks with numeric values: {len(numeric_results)}")
        for result in numeric_results:
            logger.info(f"  - {result.check_name}: {result.outcome_numeric}")


def get_sample_documents() -> List[Tuple[str, Path]]:
    """Get sample documents for demonstration (when sourcer is unavailable).
    
    Returns:
        List of (name, path) tuples
    """
    # In a real scenario, you might have some test files
    # For now, return empty list
    logger.info("  No sample documents available")
    return []


def main():
    """Run all examples."""
    logger.info("ORSA Quality Control System - Example Usage")
    logger.info("=" * 60)
    
    try:
        # Example 1: Basic usage with in-memory database
        example_with_in_memory_db()
        
        # Example 2: MSSQL database (commented out by default)
        # example_with_mssql_db()
        
        # Example 3: Force reprocess demonstration
        # example_with_force_reprocess()
        
        # Example 4: Analyzing individual check results
        # example_custom_checks()
        
    except KeyboardInterrupt:
        logger.info("\n\nInterrupted by user")
    except Exception as e:
        logger.error(f"\n\nError: {e}", exc_info=True)
    
    logger.info("\n" + "=" * 60)
    logger.info("Examples completed")
    logger.info("=" * 60)


if __name__ == "__main__":
    main()
