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
    berichtsjahr: Optional[int] = None


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
            "berichtsjahr": r.berichtsjahr,
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
    
    def get_latest_results_for_institute(self, institute_id: str) -> List[Dict[str, Any]]:
        """Get latest check results for a specific institute.
        
        Args:
            institute_id: Institute identifier
            
        Returns:
            List of check result dictionaries
        """
        try:
            query = f"""
                SELECT 
                    institute_id,
                    file_name,
                    file_hash,
                    version,
                    check_name,
                    check_description,
                    outcome_bool,
                    outcome_numeric,
                    processed_timestamp,
                    geschaeft_nr,
                    berichtsjahr
                FROM {self.schema}.vw_orsa_analysis_latest
                WHERE institute_id = '{institute_id}'
                ORDER BY check_name
            """
            df = self.execute_query(query)
            logger.debug(f"Retrieved {len(df)} results for institute {institute_id}")
            return df.to_dict('records')
        except Exception as e:
            logger.error(f"Failed to get results for institute {institute_id}: {e}")
            return []
    
    def get_all_institutes_with_results(self) -> List[str]:
        """Get list of all institutes that have check results.
        
        Returns:
            List of institute IDs
        """
        try:
            query = f"""
                SELECT DISTINCT institute_id
                FROM {self.schema}.vw_orsa_analysis_latest
                ORDER BY institute_id
            """
            df = self.execute_query(query)
            institutes = df['institute_id'].tolist()
            logger.debug(f"Found {len(institutes)} institutes with results")
            return institutes
        except Exception as e:
            logger.error(f"Failed to get institutes with results: {e}")
            return []
    
    def get_latest_version_for_institute(self, institute_id: str) -> Optional[int]:
        """Get latest version number for an institute.
        
        Args:
            institute_id: Institute identifier
            
        Returns:
            Latest version number, or None if not found
        """
        try:
            query = f"""
                SELECT MAX(version) as max_version
                FROM {self.schema}.{self.table_name}
                WHERE institute_id = '{institute_id}'
            """
            df = self.execute_query(query)
            if not df.empty and df['max_version'].iloc[0] is not None:
                version = int(df['max_version'].iloc[0])
                logger.debug(f"Latest version for {institute_id}: {version}")
                return version
            return None
        except Exception as e:
            logger.error(f"Failed to get latest version for {institute_id}: {e}")
            return None
    
    def get_institute_metadata(self, institute_id: str) -> Optional[Dict[str, Any]]:
        """Get metadata for an institute (file name, hash, version, etc.).
        
        Args:
            institute_id: Institute identifier
            
        Returns:
            Dictionary with metadata, or None if not found
        """
        try:
            query = f"""
                SELECT TOP 1
                    institute_id,
                    file_name,
                    file_hash,
                    version,
                    geschaeft_nr,
                    berichtsjahr,
                    processed_timestamp
                FROM {self.schema}.vw_orsa_analysis_latest
                WHERE institute_id = '{institute_id}'
            """
            df = self.execute_query(query)
            if not df.empty:
                metadata = df.iloc[0].to_dict()
                logger.debug(f"Retrieved metadata for {institute_id}")
                return metadata
            return None
        except Exception as e:
            logger.error(f"Failed to get metadata for {institute_id}: {e}")
            return None
    
    def get_institut_metadata_by_finmaid(self, finma_id: str) -> Optional[Dict[str, Any]]:
        """Get institut metadata from DWHMart database.
        
        This method queries the institut_metadata.sql to retrieve additional
        institute information including FinmaObjektName and MitarbeiterName.
        
        Args:
            finma_id: FINMA ID (FinmaObjektNr) of the institute
            
        Returns:
            Dictionary with keys: FINMAID, FinmaObjektName, MitarbeiterName, etc.
            Returns None if not found or on error.
        """
        try:
            # Load the SQL query from file
            sql_file = Path(__file__).parent.parent.parent.parent / "sql" / "institut_metadata.sql"
            if not sql_file.exists():
                logger.error(f"SQL file not found: {sql_file}")
                return None
            
            query = sql_file.read_text(encoding="utf-8")
            
            # Execute query and filter by FINMAID
            df = self.execute_query(query)
            
            # Filter for the specific FINMA ID
            filtered_df = df[df['FINMAID'] == finma_id]
            
            if not filtered_df.empty:
                metadata = filtered_df.iloc[0].to_dict()
                logger.debug(f"Retrieved institut metadata for FINMAID {finma_id}")
                return metadata
            else:
                logger.warning(f"No institut metadata found for FINMAID {finma_id}")
                return None
        except Exception as e:
            logger.error(f"Failed to get institut metadata for FINMAID {finma_id}: {e}")
            return None
    
    def close(self) -> None:
        """Close database connection."""
        if self.engine:
            self.engine.dispose()
