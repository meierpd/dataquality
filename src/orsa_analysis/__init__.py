"""ORSA Analysis - Data Quality Control System for Excel Files.

This package provides a comprehensive data quality control system for Excel file
validation with file processing, SHA-256-based caching, and a modular check system.

Main components:
- core: Core functionality (reader, versioning, database, processor)
- checks: Quality check rules and registry
- sourcing: Document sourcing from external systems
"""

__version__ = "0.1.0"

from orsa_analysis.core.processor import DocumentProcessor
from orsa_analysis.core.db import CheckResult, DatabaseWriter, InMemoryDatabaseWriter
from orsa_analysis.core.versioning import VersionManager
from orsa_analysis.core.reader import ExcelReader

__all__ = [
    "DocumentProcessor",
    "CheckResult",
    "DatabaseWriter",
    "InMemoryDatabaseWriter",
    "VersionManager",
    "ExcelReader",
]
