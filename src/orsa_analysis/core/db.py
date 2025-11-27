"""Database models and output structures."""

from dataclasses import dataclass, asdict
from datetime import datetime
from typing import Optional, Any, Dict
import logging

logger = logging.getLogger(__name__)


@dataclass
class CheckResult:
    """Represents a single check result that can be stored in the database.

    This corresponds to one row in the qc_results table.
    """

    institute_id: str
    file_name: str
    file_hash: str
    version_number: int
    check_name: str
    check_description: str
    outcome_bool: bool
    outcome_numeric: Optional[float]
    processed_at: datetime

    def to_dict(self) -> Dict[str, Any]:
        """Convert the result to a dictionary suitable for database insertion.

        Returns:
            Dictionary representation of the check result
        """
        data = asdict(self)
        data["processed_at"] = self.processed_at.isoformat()
        return data

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "CheckResult":
        """Create a CheckResult from a dictionary.

        Args:
            data: Dictionary containing check result data

        Returns:
            CheckResult instance
        """
        if isinstance(data["processed_at"], str):
            data["processed_at"] = datetime.fromisoformat(data["processed_at"])
        return cls(**data)


class DatabaseWriter:
    """Handles writing check results to the database.

    This is an abstract interface. Concrete implementations would
    handle specific database types (MSSQL, SQLite, etc.)
    """

    def __init__(self):
        """Initialize the database writer."""
        self.results_buffer = []

    def add_result(self, result: CheckResult) -> None:
        """Add a check result to the buffer.

        Args:
            result: CheckResult to be added
        """
        self.results_buffer.append(result)
        logger.debug(f"Added result for check: {result.check_name}")

    def get_results(self) -> list[CheckResult]:
        """Get all buffered results.

        Returns:
            List of CheckResult objects
        """
        return self.results_buffer.copy()

    def clear_buffer(self) -> None:
        """Clear the results buffer."""
        count = len(self.results_buffer)
        self.results_buffer.clear()
        logger.info(f"Cleared {count} results from buffer")

    def get_existing_versions(self) -> list[Dict[str, Any]]:
        """Retrieve existing version information from the database.

        Returns:
            List of dicts with institute_id, file_hash, and version_number

        Note:
            This is a stub method. Concrete implementations should
            query the actual database.
        """
        return []

    def write_results(self, results: list[CheckResult]) -> int:
        """Write check results to the database.

        Args:
            results: List of CheckResult objects to write

        Returns:
            Number of records written

        Note:
            This is a stub method. Concrete implementations should
            handle actual database insertion.
        """
        self.results_buffer.extend(results)
        logger.info(f"Buffered {len(results)} results for writing")
        return len(results)

    def get_results_for_institute(
        self, institute_id: str, version: Optional[int] = None
    ) -> list[CheckResult]:
        """Retrieve check results for a specific institute.

        Args:
            institute_id: Identifier for the institute
            version: Optional version number to filter by

        Returns:
            List of CheckResult objects

        Note:
            This is a stub method. Concrete implementations should
            query the actual database.
        """
        results = [r for r in self.results_buffer if r.institute_id == institute_id]
        if version is not None:
            results = [r for r in results if r.version_number == version]
        return results


class InMemoryDatabaseWriter(DatabaseWriter):
    """In-memory implementation of DatabaseWriter for testing and development."""

    def __init__(self):
        """Initialize the in-memory database writer."""
        super().__init__()
        self.stored_results = []

    def write_results(self, results: list[CheckResult]) -> int:
        """Write check results to in-memory storage.

        Args:
            results: List of CheckResult objects to write

        Returns:
            Number of records written
        """
        self.stored_results.extend(results)
        logger.info(f"Wrote {len(results)} results to in-memory storage")
        return len(results)
    
    def get_results(self) -> list[CheckResult]:
        """Get all stored results.

        Returns:
            List of CheckResult objects
        """
        return self.stored_results.copy()

    def get_existing_versions(self) -> list[Dict[str, Any]]:
        """Retrieve existing version information from in-memory storage.

        Returns:
            List of dicts with institute_id, file_hash, and version_number
        """
        unique_versions = {}
        for result in self.stored_results:
            key = (result.institute_id, result.file_hash)
            if key not in unique_versions:
                unique_versions[key] = {
                    "institute_id": result.institute_id,
                    "file_hash": result.file_hash,
                    "version_number": result.version_number,
                }
        return list(unique_versions.values())

    def get_results_for_institute(
        self, institute_id: str, version: Optional[int] = None
    ) -> list[CheckResult]:
        """Retrieve check results for a specific institute.

        Args:
            institute_id: Identifier for the institute
            version: Optional version number to filter by

        Returns:
            List of CheckResult objects
        """
        results = [r for r in self.stored_results if r.institute_id == institute_id]
        if version is not None:
            results = [r for r in results if r.version_number == version]
        return results

    def clear_all(self) -> None:
        """Clear all stored results."""
        self.stored_results.clear()
        self.results_buffer.clear()
        logger.info("Cleared all in-memory storage")
