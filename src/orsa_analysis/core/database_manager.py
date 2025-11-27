"""Simple database manager for FINMA MSSQL databases."""

import os
import logging
from dataclasses import dataclass
from datetime import datetime
from typing import Optional, List, Dict, Any
from pathlib import Path
from dotenv import load_dotenv
from sqlalchemy import create_engine
import pandas as pd

logger = logging.getLogger(__name__)


@dataclass
class CheckResult:
    """Check result that can be stored in the database."""

    institute_id: str
    file_name: str
    file_hash: str
    version_number: int
    check_name: str
    check_description: str
    outcome_bool: bool
    outcome_numeric: Optional[float]
    processed_at: datetime
    geschaeft_nr: Optional[str] = None


class DatabaseManager:
    """Simple database manager for MSSQL connections and result storage.
    
    Supports Windows authentication or credential-based authentication.
    """
    
    def __init__(
        self,
        server: str = "dwhdata.finma.ch",
        database: str = "GBI_REPORTING",
        schema: str = "gbi",
        table_name: str = "orsa_analysis_data",
        credentials_file: Optional[Path] = None
    ):
        """Initialize database manager.
        
        Args:
            server: Database server hostname
            database: Database name
            schema: Database schema for writing results
            table_name: Table name for storing results
            credentials_file: Path to credentials.env file (optional)
                Expected to contain 'username' and 'password' fields.
                For GBB_Reporting server, username will be prefixed with 'Finma\\'.
        """
        self.server = server
        self.database = database
        self.schema = schema
        self.table_name = table_name
        
        # Load credentials if provided
        if credentials_file and Path(credentials_file).exists():
            load_dotenv(credentials_file)
            username = os.getenv("username", "")
            password = os.getenv("password", "")
            
            if username and password:
                # Store credentials in environment variables (without domain prefix)
                # This allows other components (like document_sourcer) to use them
                os.environ["DB_USER"] = username
                os.environ["DB_PASSWORD"] = password
                logger.info(f"Loaded credentials from {credentials_file}")
        
        self.engine = self._create_engine()
        
        logger.info(f"Database manager initialized for {server}/{database}")
    
    def _create_engine(self):
        """Create SQLAlchemy engine with Windows auth or credentials."""
        if "DB_USER" in os.environ and "DB_PASSWORD" in os.environ:
            db_user = os.environ["DB_USER"]
            db_password = os.environ["DB_PASSWORD"]
            
            # Prefix username with domain for GBB_Reporting server
            if self.server == "frbdata.finma.ch":
                db_user = "Finma\\" + db_user
            
            conn_str = f"mssql+pymssql://{db_user}:{db_password}@{self.server}/{self.database}"
            logger.info(f"Using pymssql driver with credential-based auth for {self.server}/{self.database}")
            logger.debug(f"Connection string: mssql+pymssql://{db_user}:***@{self.server}/{self.database}")
            # pymssql driver doesn't support use_setinputsizes parameter
            return create_engine(conn_str)
        else:
            conn_str = f"mssql+pyodbc://{self.server}/{self.database}?driver=SQL+Server"
            logger.info(f"Using pyodbc driver with Windows authentication for {self.server}/{self.database}")
            logger.debug(f"Connection string: {conn_str}")
            # pyodbc driver supports use_setinputsizes parameter
            return create_engine(conn_str, use_setinputsizes=False)
    
    def execute_query(self, query: str) -> pd.DataFrame:
        """Execute SQL query and return DataFrame."""
        with self.engine.connect() as conn:
            return pd.read_sql(query, con=conn)
    
    def execute_query_from_file(self, sql_file: Path) -> pd.DataFrame:
        """Execute SQL query from file."""
        if not sql_file.exists():
            raise FileNotFoundError(f"SQL file not found: {sql_file}")
        query = sql_file.read_text(encoding="utf-8")
        return self.execute_query(query)
    
    def write_results(self, results: List[CheckResult]) -> int:
        """Write check results to database.
        
        Args:
            results: List of CheckResult objects
            
        Returns:
            Number of records written
        """
        if not results:
            return 0
        
        # Convert to DataFrame
        data = [{
            "institute_id": r.institute_id,
            "file_name": r.file_name,
            "file_hash": r.file_hash,
            "version": r.version_number,
            "check_name": r.check_name,
            "check_description": r.check_description,
            "outcome_bool": int(r.outcome_bool),
            "outcome_numeric": r.outcome_numeric,
            "processed_timestamp": r.processed_at,
            "geschaeft_nr": r.geschaeft_nr,
        } for r in results]
        
        df = pd.DataFrame(data)
        
        with self.engine.connect() as conn:
            df.to_sql(
                name=self.table_name,
                con=conn,
                schema=self.schema,
                if_exists="append",
                index=False
            )
        
        logger.info(f"Wrote {len(results)} results to {self.schema}.{self.table_name}")
        return len(results)
    
    def get_existing_versions(self) -> List[Dict[str, Any]]:
        """Get existing file versions from database."""
        try:
            query = f"""
                SELECT DISTINCT 
                    institute_id,
                    file_hash,
                    MAX(version) as version_number
                FROM {self.schema}.{self.table_name}
                GROUP BY institute_id, file_hash
            """
            df = self.execute_query(query)
            return df.to_dict('records')
        except Exception as e:
            logger.warning(f"Could not load existing versions: {e}")
            return []
    
    def close(self) -> None:
        """Close database connection."""
        if self.engine:
            self.engine.dispose()
