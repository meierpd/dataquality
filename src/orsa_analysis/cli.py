"""Command-line interface for the data quality control tool."""

import argparse
import logging
import sys
from pathlib import Path

from orsa_analysis.core.processor import DocumentProcessor
from orsa_analysis.core.db import MSSQLDatabaseWriter
from orsa_analysis.sourcing import ORSADocumentSourcer

logger = logging.getLogger(__name__)


def setup_logging(verbose: bool = False) -> None:
    """Configure logging for the application.

    Args:
        verbose: If True, set log level to DEBUG, otherwise INFO
    """
    log_level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=log_level,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def process_from_sourcer(
    force_reprocess: bool = False,
    verbose: bool = False,
    credentials_file: str = "credentials.env"
) -> None:
    """Process documents from ORSADocumentSourcer and write to MSSQL database.

    Args:
        force_reprocess: If True, reprocess all files regardless of cache
        verbose: If True, enable verbose logging
        credentials_file: Path to credentials.env file
    """
    setup_logging(verbose)

    logger.info("Starting ORSA data quality control processing")
    logger.info(f"Force reprocess: {force_reprocess}")

    # Initialize database writer for MSSQL
    db_writer = MSSQLDatabaseWriter(
        server="frbdata.finma.ch",
        database="GBI_REPORTING"
    )
    
    processor = DocumentProcessor(db_writer, force_reprocess=force_reprocess)

    try:
        # Initialize document sourcer
        sourcer = ORSADocumentSourcer()
        
        documents = sourcer.load()
        logger.info(f"Loaded {len(documents)} documents from sourcer")

        results = processor.process_documents(documents)

        # Write results to database
        logger.info("Writing results to MSSQL database...")
        db_writer.write_results()
        logger.info("Results successfully written to database")

        summary = processor.get_processing_summary()
        logger.info("=" * 60)
        logger.info("PROCESSING SUMMARY")
        logger.info("=" * 60)
        logger.info(f"Total files processed: {summary['total_files']}")
        logger.info(f"Total checks executed: {summary['total_checks']}")
        logger.info(f"Checks passed: {summary['checks_passed']}")
        logger.info(f"Checks failed: {summary['checks_failed']}")
        logger.info(f"Pass rate: {summary['pass_rate']:.1%}")
        logger.info(f"Institutes: {', '.join(summary['institutes'])}")
        logger.info("=" * 60)

    except Exception as e:
        logger.error(f"Processing failed: {e}", exc_info=True)
        sys.exit(1)


def main():
    """Command-line interface entry point."""
    parser = argparse.ArgumentParser(
        description="ORSA Data Quality Control Tool - Process documents from sourcer and write to MSSQL"
    )
    parser.add_argument(
        "--force",
        "-f",
        action="store_true",
        help="Force reprocessing of all files, ignoring cache",
    )
    parser.add_argument(
        "--verbose",
        "-v",
        action="store_true",
        help="Enable verbose logging",
    )
    parser.add_argument(
        "--credentials",
        "-c",
        default="credentials.env",
        help="Path to credentials file (default: credentials.env)",
    )

    args = parser.parse_args()

    process_from_sourcer(
        force_reprocess=args.force,
        verbose=args.verbose,
        credentials_file=args.credentials,
    )


if __name__ == "__main__":
    main()
