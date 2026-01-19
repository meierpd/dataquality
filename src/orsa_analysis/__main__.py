"""Main entry point for the data quality control tool.

This module can be run directly as:
    python -m orsa_analysis

Or via the installed command:
    orsa-qc

Quick Start:
    To run with default settings:
    
        python -m orsa_analysis
    
    This will:
    - Connect to the database using credentials from credentials.env
    - Process ORSA documents for reporting year 2026
    - Skip already processed files (use --force to reprocess all)
"""

import argparse
import logging
import sys
from pathlib import Path

from orsa_analysis import ORSAPipeline, DatabaseManager
from orsa_analysis.sourcing import ORSADocumentSourcer
from orsa_analysis.reporting import ReportGenerator

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
    credentials_file: str = "credentials.env",
    berichtsjahr: int = 2026,
    generate_reports: bool = True,
    output_dir: str = "reports",
    template_file: str = "data/auswertungs_template.xlsx"
) -> None:
    """Process documents from ORSADocumentSourcer and write to MSSQL database.

    Args:
        force_reprocess: If True, reprocess all files regardless of cache
        verbose: If True, enable verbose logging
        credentials_file: Path to credentials.env file
        berichtsjahr: Reporting year to filter documents (default: 2026)
        generate_reports: If True, generate Excel reports after processing (default: True)
        output_dir: Directory for output reports (default: reports)
        template_file: Path to template file (default: data/auswertungs_template.xlsx)
    """
    setup_logging(verbose)

    logger.info("Starting ORSA data quality control processing")
    logger.info(f"Force reprocess: {force_reprocess}")
    logger.info(f"Berichtsjahr: {berichtsjahr}")
    logger.info(f"Generate reports: {generate_reports}")

    try:
        # Initialize database manager - it will handle credentials automatically
        db_manager = DatabaseManager()
        
        # Initialize pipeline
        pipeline = ORSAPipeline(db_manager, force_reprocess=force_reprocess)
        
        # Initialize document sourcer
        sourcer = ORSADocumentSourcer(cred_file=credentials_file, berichtsjahr=berichtsjahr)
        
        # Process documents through pipeline
        logger.info("Loading documents from FINMA database...")
        summary = pipeline.process_from_sourcer(sourcer)

        # Display summary
        logger.info("=" * 60)
        logger.info("PROCESSING SUMMARY")
        logger.info("=" * 60)
        logger.info(f"Files processed: {summary['files_processed']}")
        logger.info(f"Files skipped: {summary['files_skipped']}")
        logger.info(f"Total checks: {summary['total_checks']}")
        logger.info(f"Checks passed: {summary['checks_passed']}")
        logger.info(f"Pass rate: {summary['pass_rate']:.1%}")
        logger.info(f"Institutes: {', '.join(summary['institutes'])}")
        logger.info("=" * 60)
        
        # Generate reports if requested
        if generate_reports:
            logger.info("")
            logger.info("=" * 60)
            logger.info("GENERATING REPORTS")
            logger.info("=" * 60)
            
            # Get source files from sourcer
            documents = sourcer.load()
            source_files = {}
            for document_name, file_path, geschaeft_nr, finma_id, berichtsjahr in documents:
                # Use FinmaID from database as the institute_id
                source_files[finma_id] = Path(file_path)
            
            # Initialize report generator
            report_gen = ReportGenerator(
                db_manager=db_manager,
                template_path=Path(template_file),
                output_dir=Path(output_dir)
            )
            
            # Generate reports
            report_paths = report_gen.generate_all_reports(
                source_files=source_files,
                force_overwrite=False
            )
            
            logger.info("=" * 60)
            logger.info(f"Generated {len(report_paths)} reports in {output_dir}")
            logger.info("=" * 60)
        
        # Close pipeline
        pipeline.close()
        logger.info("Processing completed successfully")

        return summary

    except Exception as e:
        logger.error(f"Processing failed: {e}", exc_info=True)
        sys.exit(1)


def generate_reports_only(
    verbose: bool = False,
    credentials_file: str = "credentials.env",
    output_dir: str = "reports",
    template_file: str = "data/auswertungs_template.xlsx",
    institute_id: str = None,
    berichtsjahr: int = 2026,
    force_overwrite: bool = False
) -> None:
    """Generate reports from existing check results in database.

    Args:
        verbose: If True, enable verbose logging
        credentials_file: Path to credentials.env file
        output_dir: Directory for output reports
        template_file: Path to template file
        institute_id: Optional specific institute to generate report for
        berichtsjahr: Reporting year for sourcing files (default: 2026)
        force_overwrite: If True, overwrite existing reports
    """
    setup_logging(verbose)

    logger.info("Starting ORSA report generation")
    logger.info(f"Output directory: {output_dir}")
    logger.info(f"Template file: {template_file}")

    try:
        # Initialize database manager
        db_manager = DatabaseManager()
        
        # Initialize document sourcer to get source files
        sourcer = ORSADocumentSourcer(cred_file=credentials_file, berichtsjahr=berichtsjahr)
        documents = sourcer.load()
        source_files = {}
        for document_name, file_path, geschaeft_nr, finma_id, berichtsjahr_val in documents:
            # Use FinmaID from database as the institute_id
            source_files[finma_id] = Path(file_path)
        
        logger.info(f"Found {len(source_files)} source files")
        
        # Initialize report generator
        report_gen = ReportGenerator(
            db_manager=db_manager,
            template_path=Path(template_file),
            output_dir=Path(output_dir)
        )
        
        # Generate reports
        if institute_id:
            logger.info(f"Generating report for institute: {institute_id}")
            source_path = source_files.get(institute_id)
            report_path = report_gen.generate_report(
                institute_id=institute_id,
                source_file_path=source_path,
                force_overwrite=force_overwrite
            )
            if report_path:
                logger.info(f"Report generated: {report_path}")
            else:
                logger.warning(f"No report generated for {institute_id}")
        else:
            logger.info("Generating reports for all institutes")
            report_paths = report_gen.generate_all_reports(
                source_files=source_files,
                force_overwrite=force_overwrite
            )
            logger.info(f"Generated {len(report_paths)} reports")
        
        logger.info("Report generation completed successfully")

    except Exception as e:
        logger.error(f"Report generation failed: {e}", exc_info=True)
        sys.exit(1)


def main():
    """Command-line interface entry point."""
    parser = argparse.ArgumentParser(
        description="ORSA Data Quality Control Tool - Process documents and generate reports"
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
        type=str,
        default="credentials.env",
        help="Path to credentials.env file (default: credentials.env)",
    )
    parser.add_argument(
        "--berichtsjahr",
        "-b",
        type=int,
        default=2026,
        help="Reporting year to filter documents (default: 2026)",
    )
    parser.add_argument(
        "--no-reports",
        action="store_true",
        help="Skip Excel report generation after processing checks (reports are generated by default)",
    )
    parser.add_argument(
        "--reports-only",
        action="store_true",
        help="Only generate reports without running checks (requires existing results in database)",
    )
    parser.add_argument(
        "--output-dir",
        "-o",
        default="reports",
        help="Directory for output reports (default: reports)",
    )
    parser.add_argument(
        "--template",
        "-t",
        default="data/auswertungs_template.xlsx",
        help="Path to template Excel file (default: data/auswertungs_template.xlsx)",
    )
    parser.add_argument(
        "--institute",
        "-i",
        default=None,
        help="Generate report for specific institute only (use with --reports-only)",
    )
    parser.add_argument(
        "--force-overwrite",
        action="store_true",
        help="Overwrite existing report files (use with --reports-only)",
    )

    args = parser.parse_args()

    # Run in reports-only mode
    if args.reports_only:
        generate_reports_only(
            verbose=args.verbose,
            credentials_file=args.credentials,
            output_dir=args.output_dir,
            template_file=args.template,
            institute_id=args.institute,
            berichtsjahr=args.berichtsjahr,
            force_overwrite=args.force_overwrite,
        )
    else:
        # Run normal processing (with optional report generation)
        process_from_sourcer(
            force_reprocess=args.force,
            verbose=args.verbose,
            credentials_file=args.credentials,
            berichtsjahr=args.berichtsjahr,
            generate_reports=not args.no_reports,
            output_dir=args.output_dir,
            template_file=args.template,
        )


if __name__ == "__main__":
    main()
