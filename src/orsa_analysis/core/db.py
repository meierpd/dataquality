"""Database models and output structures."""

from dataclasses import dataclass, asdict
from datetime import datetime
from typing import Optional, Any, Dict
from pathlib import Path
import logging
import pandas as pd

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


class MSSQLDatabaseWriter(DatabaseWriter):
    """MSSQL implementation of DatabaseWriter using DatabaseManager.
    
    Writes check results to GBI_REPORTING.gbi.orsa_analysis_data table.
    """
    
    def __init__(
        self,
        server: str = "dwhdata.finma.ch",
        database: str = "GBI_REPORTING",
        schema: str = "gbi",
        table_name: str = "orsa_analysis_data",
        credentials_file: Optional[Path] = None
    ):
        """Initialize the MSSQL database writer.
        
        Args:
            server: Database server hostname
            database: Database name
            schema: Database schema
            table_name: Table name for storing results
            credentials_file: Path to credentials.env file
        """
        super().__init__()
        from orsa_analysis.core.database_manager import DatabaseManager
        
        self.db_manager = DatabaseManager(
            server=server,
            database=database,
            credentials_file=credentials_file
        )
        self.schema = schema
        self.table_name = table_name
        
        # Test connection
        if not self.db_manager.test_connection():
            logger.warning("Database connection test failed")
    
    def write_results(self, results: list[CheckResult]) -> int:
        """Write check results to MSSQL database.
        
        Args:
            results: List of CheckResult objects to write
            
        Returns:
            Number of records written
            
        Raises:
            Exception: If database write fails
        """
        if not results:
            logger.info("No results to write")
            return 0
        
        # Convert results to DataFrame
        data = []
        for result in results:
            data.append({
                "institute_id": result.institute_id,
                "file_name": result.file_name,
                "file_hash": result.file_hash,
                "version": result.version_number,
                "check_name": result.check_name,
                "check_description": result.check_description,
                "outcome_bool": int(result.outcome_bool),  # Convert bool to int for BIT type
                "outcome_numeric": result.outcome_numeric,
                "processed_timestamp": result.processed_at,
            })
        
        df = pd.DataFrame(data)
        
        try:
            self.db_manager.write_dataframe(
                df=df,
                table_name=self.table_name,
                schema=self.schema,
                if_exists="append"
            )
            logger.info(f"Successfully wrote {len(results)} results to {self.schema}.{self.table_name}")
            return len(results)
        except Exception as e:
            logger.error(f"Failed to write results to database: {e}")
            raise
    
    def get_existing_versions(self) -> list[Dict[str, Any]]:
        """Retrieve existing version information from the database.
        
        Returns:
            List of dicts with institute_id, file_hash, and version_number
        """
        try:
            query = f"""
                SELECT DISTINCT 
                    institute_id,
                    file_hash,
                    MAX(version) as version_number
                FROM {self.schema}.{self.table_name}
                GROUP BY institute_id, file_hash
            """
            df = self.db_manager.execute_query(query)
            return df.to_dict('records')
        except Exception as e:
            logger.error(f"Failed to retrieve existing versions: {e}")
            return []
    
    def get_results_for_institute(
        self, institute_id: str, version: Optional[int] = None
    ) -> list[CheckResult]:
        """Retrieve check results for a specific institute from database.
        
        Args:
            institute_id: Identifier for the institute
            version: Optional version number to filter by
            
        Returns:
            List of CheckResult objects
        """
        try:
            query = f"""
                SELECT 
                    institute_id,
                    file_name,
                    file_hash,
                    version as version_number,
                    check_name,
                    check_description,
                    outcome_bool,
                    outcome_numeric,
                    processed_timestamp as processed_at
                FROM {self.schema}.{self.table_name}
                WHERE institute_id = '{institute_id}'
            """
            
            if version is not None:
                query += f" AND version = {version}"
            
            df = self.db_manager.execute_query(query)
            
            results = []
            for _, row in df.iterrows():
                results.append(CheckResult(
                    institute_id=row["institute_id"],
                    file_name=row["file_name"],
                    file_hash=row["file_hash"],
                    version_number=int(row["version_number"]),
                    check_name=row["check_name"],
                    check_description=row["check_description"],
                    outcome_bool=bool(row["outcome_bool"]),
                    outcome_numeric=row["outcome_numeric"] if pd.notna(row["outcome_numeric"]) else None,
                    processed_at=row["processed_at"],
                ))
            
            return results
        except Exception as e:
            logger.error(f"Failed to retrieve results for institute {institute_id}: {e}")
            return []
    
    def close(self) -> None:
        """Close the database connection."""
        if self.db_manager:
            self.db_manager.close()
