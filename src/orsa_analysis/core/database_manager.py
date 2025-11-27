"""Database connection manager for FINMA MSSQL databases."""

import os
import logging
from typing import Optional
from pathlib import Path
from dotenv import load_dotenv
from sqlalchemy import create_engine
from sqlalchemy.engine import Engine
import pandas as pd

logger = logging.getLogger(__name__)


class DatabaseManager:
    """Manages database connections with FINMA MSSQL servers.
    
    Supports both Windows authentication (local) and credential-based authentication.
    Credentials are loaded from environment variables or credentials.env file.
    
    Attributes:
        server: Database server hostname
        database: Database name
        engine: SQLAlchemy engine instance
    """
    
    def __init__(
        self,
        server: Optional[str] = None,
        database: Optional[str] = None,
        credentials_file: Optional[Path] = None
    ):
        """Initialize database manager.
        
        Args:
            server: Database server hostname (default: dwhdata.finma.ch)
            database: Database name (default: DWHMart)
            credentials_file: Path to credentials.env file (optional)
        """
        # Load credentials from file if provided
        if credentials_file:
            if Path(credentials_file).exists():
                load_dotenv(credentials_file)
                logger.debug(f"Loaded credentials from {credentials_file}")
            else:
                logger.warning(f"Credentials file not found: {credentials_file}")
        
        # Set defaults
        self.server = server or "dwhdata.finma.ch"
        self.database = database or "DWHMart"
        
        # Create engine
        self.engine = self._create_engine()
        logger.info(f"Database manager initialized for {self.server}/{self.database}")
    
    def _create_engine(self) -> Engine:
        """Create SQLAlchemy engine based on available credentials.
        
        Returns:
            SQLAlchemy Engine instance
        """
        # Check if environment variables for database connection are set
        if "DB_USER" in os.environ and "DB_PASSWORD" in os.environ:
            # Use credentials login for connection with database server
            db_user = os.environ["DB_USER"]
            db_password = os.environ["DB_PASSWORD"]
            connection_string = f"mssql+pymssql://{db_user}:{db_password}@{self.server}/{self.database}"
            logger.info(f"Using credential-based authentication for {self.server}")
        else:
            # Use Windows login for local connection
            connection_string = f"mssql+pyodbc://{self.server}/{self.database}?driver=SQL+Server"
            logger.info(f"Using Windows authentication for {self.server}")
        
        # In newer versions of sqlalchemy, use_setinputsizes=False must be set 
        # in order to avoid a precision error while writing to the db
        engine = create_engine(connection_string, use_setinputsizes=False)
        
        return engine
    
    def execute_query(self, query: str) -> pd.DataFrame:
        """Execute a SQL query and return results as DataFrame.
        
        Args:
            query: SQL query string
            
        Returns:
            DataFrame with query results
            
        Raises:
            Exception: If query execution fails
        """
        try:
            with self.engine.connect() as connection:
                df = pd.read_sql(query, con=connection)
            logger.debug(f"Query returned {len(df)} rows")
            return df
        except Exception as e:
            logger.error(f"Query execution failed: {e}")
            raise
    
    def execute_query_from_file(self, sql_file: Path) -> pd.DataFrame:
        """Execute a SQL query from a file.
        
        Args:
            sql_file: Path to SQL file
            
        Returns:
            DataFrame with query results
            
        Raises:
            FileNotFoundError: If SQL file doesn't exist
            Exception: If query execution fails
        """
        if not sql_file.exists():
            raise FileNotFoundError(f"SQL file not found: {sql_file}")
        
        query = sql_file.read_text(encoding="utf-8")
        logger.debug(f"Executing query from {sql_file}")
        return self.execute_query(query)
    
    def write_dataframe(
        self,
        df: pd.DataFrame,
        table_name: str,
        schema: Optional[str] = None,
        if_exists: str = "append"
    ) -> None:
        """Write a DataFrame to a database table.
        
        Args:
            df: DataFrame to write
            table_name: Name of the target table
            schema: Database schema (optional)
            if_exists: How to behave if table exists {'fail', 'replace', 'append'}
                      Default is 'append'
        
        Raises:
            Exception: If write operation fails
        """
        try:
            with self.engine.connect() as connection:
                df.to_sql(
                    name=table_name,
                    con=connection,
                    schema=schema,
                    if_exists=if_exists,
                    index=False
                )
            logger.info(f"Wrote {len(df)} rows to {schema}.{table_name}" if schema else f"Wrote {len(df)} rows to {table_name}")
        except Exception as e:
            logger.error(f"Failed to write DataFrame to database: {e}")
            raise
    
    def test_connection(self) -> bool:
        """Test the database connection.
        
        Returns:
            True if connection successful, False otherwise
        """
        try:
            with self.engine.connect() as connection:
                result = connection.execute("SELECT 1 AS test")
                row = result.fetchone()
                if row and row[0] == 1:
                    logger.info("Database connection test successful")
                    return True
            return False
        except Exception as e:
            logger.error(f"Database connection test failed: {e}")
            return False
    
    def get_table_row_count(self, table_name: str, schema: Optional[str] = None) -> int:
        """Get the number of rows in a table.
        
        Args:
            table_name: Name of the table
            schema: Database schema (optional)
            
        Returns:
            Number of rows in the table
        """
        table_ref = f"{schema}.{table_name}" if schema else table_name
        query = f"SELECT COUNT(*) as count FROM {table_ref}"
        df = self.execute_query(query)
        return int(df.iloc[0]["count"])
    
    def close(self) -> None:
        """Close the database connection."""
        if self.engine:
            self.engine.dispose()
            logger.info("Database connection closed")


# Factory functions for common database configurations

def create_dwh_mart_manager(credentials_file: Optional[Path] = None) -> DatabaseManager:
    """Create a DatabaseManager for DWHMart database.
    
    Args:
        credentials_file: Path to credentials.env file (optional)
        
    Returns:
        DatabaseManager instance
    """
    return DatabaseManager(
        server="dwhdata.finma.ch",
        database="DWHMart",
        credentials_file=credentials_file
    )


def create_gbi_reporting_manager(credentials_file: Optional[Path] = None) -> DatabaseManager:
    """Create a DatabaseManager for GBI_REPORTING database.
    
    Args:
        credentials_file: Path to credentials.env file (optional)
        
    Returns:
        DatabaseManager instance
    """
    return DatabaseManager(
        server="dwhdata.finma.ch",
        database="GBI_REPORTING",
        credentials_file=credentials_file
    )
