"""Main entry point for the data quality control tool."""

import argparse
import logging
import sys
from pathlib import Path
from typing import List, Tuple

from core.processor import DocumentProcessor
from core.db import InMemoryDatabaseWriter

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
    sourcer_instance,
    force_reprocess: bool = False,
    verbose: bool = False,
) -> None:
    """Process documents from an ORSADocumentSourcer instance.

    Args:
        sourcer_instance: Instance of ORSADocumentSourcer with load() method
        force_reprocess: If True, reprocess all files regardless of cache
        verbose: If True, enable verbose logging
    """
    setup_logging(verbose)

    logger.info("Starting data quality control processing")
    logger.info(f"Force reprocess: {force_reprocess}")

    db_writer = InMemoryDatabaseWriter()
    processor = DocumentProcessor(db_writer, force_reprocess=force_reprocess)

    try:
        documents = sourcer_instance.load()
        logger.info(f"Loaded {len(documents)} documents from sourcer")

        results = processor.process_documents(documents)

        summary = processor.get_processing_summary()
        logger.info("=" * 60)
        logger.info("PROCESSING SUMMARY")
        logger.info("=" * 60)
        logger.info(f"Total files processed: {summary['total_files']}")
        logger.info(f"Total checks executed: {summary['total_checks']}")
        logger.info(f"Checks passed: {summary['checks_passed']}")
        logger.info(f"Checks failed: {summary['checks_failed']}")
        logger.info(f"Pass rate: {summary['pass_rate']}")
        logger.info(f"Institutes: {', '.join(summary['institutes'])}")
        logger.info("=" * 60)

        return results

    except Exception as e:
        logger.error(f"Processing failed: {e}", exc_info=True)
        sys.exit(1)


def process_from_directory(
    directory: Path,
    force_reprocess: bool = False,
    verbose: bool = False,
) -> None:
    """Process Excel files from a directory.

    Args:
        directory: Path to directory containing Excel files
        force_reprocess: If True, reprocess all files regardless of cache
        verbose: If True, enable verbose logging
    """
    setup_logging(verbose)

    if not directory.exists():
        logger.error(f"Directory not found: {directory}")
        sys.exit(1)

    if not directory.is_dir():
        logger.error(f"Not a directory: {directory}")
        sys.exit(1)

    logger.info(f"Processing files from directory: {directory}")

    excel_files = list(directory.glob("*.xlsx")) + list(directory.glob("*.xlsm"))
    if not excel_files:
        logger.warning(f"No Excel files found in {directory}")
        sys.exit(0)

    documents: List[Tuple[str, Path]] = [(f.name, f) for f in excel_files]

    db_writer = InMemoryDatabaseWriter()
    processor = DocumentProcessor(db_writer, force_reprocess=force_reprocess)

    try:
        results = processor.process_documents(documents)

        summary = processor.get_processing_summary()
        logger.info("=" * 60)
        logger.info("PROCESSING SUMMARY")
        logger.info("=" * 60)
        logger.info(f"Total files processed: {summary['total_files']}")
        logger.info(f"Total checks executed: {summary['total_checks']}")
        logger.info(f"Checks passed: {summary['checks_passed']}")
        logger.info(f"Checks failed: {summary['checks_failed']}")
        logger.info(f"Pass rate: {summary['pass_rate']}")
        logger.info(f"Institutes: {', '.join(summary['institutes'])}")
        logger.info("=" * 60)

    except Exception as e:
        logger.error(f"Processing failed: {e}", exc_info=True)
        sys.exit(1)


def main():
    """Command-line interface entry point."""
    parser = argparse.ArgumentParser(
        description="Data Quality Control Tool for Excel files"
    )
    parser.add_argument(
        "directory",
        type=Path,
        help="Directory containing Excel files to process",
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

    args = parser.parse_args()

    process_from_directory(
        directory=args.directory,
        force_reprocess=args.force,
        verbose=args.verbose,
    )


if __name__ == "__main__":
    main()
