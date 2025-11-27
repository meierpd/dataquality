"""Document sourcer for retrieving ORSA documents from the database."""

import os
import logging
from pathlib import Path
from typing import List, Tuple

import pandas as pd
import requests
from requests_ntlm import HttpNtlmAuth

logger = logging.getLogger(__name__)

# Configure environment for FINMA network
os.environ["NO_PROXY"] = "finma.ch"
os.environ["REQUESTS_CA_BUNDLE"] = "SwisscomRootCore.crt"


class ORSADocumentSourcer:
    """Source ORSA documents from the database and download them locally.
    
    This class handles:
    - Querying the database for document metadata
    - Filtering relevant ORSA documents
    - Downloading documents from SharePoint links
    
    Note:
        Credentials are managed by DatabaseManager and accessed via environment variables.
        The DatabaseManager must be initialized before downloading documents to ensure
        credentials are loaded into the environment.
    
    Attributes:
        base_dir: Base directory of the sourcing module
        cred_file: Path to credentials file
        default_target_dir: Directory where documents are downloaded
        berichtsjahr: Reporting year to filter documents for
    """

    def __init__(self, cred_file: str = "credentials.env", berichtsjahr: int = 2026):
        """Initialize the document sourcer.
        
        Args:
            cred_file: Name of credentials file (relative to project root)
                This is passed to DatabaseManager for credential loading.
            berichtsjahr: Reporting year to filter documents (default: 2026)
        """
        self.base_dir = Path(__file__).resolve().parent
        self.cred_file = self.base_dir.parent.parent.parent / cred_file
        self.default_target_dir = self.base_dir.parent.parent.parent / "data" / "orsa_response_files"
        self.berichtsjahr = berichtsjahr
        logger.info(f"Initialized ORSADocumentSourcer for Berichtsjahr {berichtsjahr}")

    def _load_query(self, name: str) -> str:
        """Load SQL query from file.
        
        Args:
            name: Name of the query file (without .sql extension)
            
        Returns:
            Query string
            
        Raises:
            FileNotFoundError: If query file doesn't exist
        """
        query_file = self.base_dir.parent.parent.parent / "sql" / f"{name}.sql"
        if not query_file.exists():
            raise FileNotFoundError(f"Query file not found: {query_file}")
        return query_file.read_text()

    def _run_query(self, query: str) -> pd.DataFrame:
        """Execute SQL query against the database.
        
        Args:
            query: SQL query string
            
        Returns:
            Query results as DataFrame
        """
        from orsa_analysis.core.database_manager import DatabaseManager
        
        # Create database manager for GBB_Reporting
        # DatabaseManager will handle loading credentials and prefixing username
        db_manager = DatabaseManager(
            server="frbdata.finma.ch",
            database="GBB_Reporting",
            credentials_file=self.cred_file
        )
        
        logger.info("Executing query against GBB_Reporting database")
        return db_manager.execute_query(query)

    def get_document_metadata(self) -> pd.DataFrame:
        """Retrieve document metadata from the database for the specified reporting year.
        
        The SQL query is parameterized with berichtsjahr to filter at the database level.
        
        Returns:
            DataFrame with columns: DokumentName, DokumentLink, Berichtsjahr, etc.
        """
        q = self._load_query("source_orsa_dokument_metadata")
        # Substitute berichtsjahr parameter in query
        q = q.format(berichtsjahr=self.berichtsjahr)
        document_metadata_df = self._run_query(q)
        logger.info(
            f"Retrieved {len(document_metadata_df)} documents from database "
            f"for Berichtsjahr {self.berichtsjahr}"
        )
        return document_metadata_df

    def filter_relevant(self, df: pd.DataFrame) -> pd.DataFrame:
        """Filter for relevant ORSA documents.
        
        Note: Most filtering is now done at the database level via the SQL query.
        This method serves as a safety check and for backward compatibility.
        
        Filters for:
        - Documents containing "_ORSA-Formular" in name (redundant with SQL WHERE clause)
        - Reporting year matching the specified berichtsjahr (redundant with SQL WHERE clause)
        
        Args:
            df: DataFrame with document metadata
            
        Returns:
            Filtered DataFrame with reporting_year column added
        """
        df = df.copy()
        
        # If empty, still add the expected column for backward compatibility
        if len(df) == 0:
            logger.warning("No documents returned from database query")
            df["reporting_year"] = pd.Series(dtype=float)
            return df
        
        # Safety check: filter by year if Berichtsjahr column exists
        if "Berichtsjahr" in df.columns:
            mask = df["Berichtsjahr"] == self.berichtsjahr
            filtered_df = df[mask]
            # Add reporting_year column for backward compatibility
            filtered_df["reporting_year"] = filtered_df["Berichtsjahr"]
        else:
            # Fallback: extract year from DokumentName
            df["reporting_year"] = (
                df["DokumentName"].str.extract(r"(\d{4})")[0].astype(float)
            )
            mask = df["DokumentName"].str.lower().str.contains("_orsa-formular") & (
                df["reporting_year"] == self.berichtsjahr
            )
            filtered_df = df[mask]
            
        logger.info(
            f"Verified {len(filtered_df)} relevant ORSA documents for Berichtsjahr {self.berichtsjahr}"
        )
        return filtered_df

    def download_documents(
        self, document_df: pd.DataFrame, target_dir: Path = None
    ) -> List[Tuple[str, Path, str, str]]:
        """Download documents from SharePoint links.
        
        Args:
            document_df: DataFrame with DokumentName, DokumentLink, GeschaeftNr, and FinmaID columns
            target_dir: Directory to save files (default: orsa_response_files/)
            
        Returns:
            List of tuples (document_name, file_path, geschaeft_nr, finma_id)
            
        Note:
            Requires DB_USER and DB_PASSWORD to be set in environment variables.
            These are set by DatabaseManager when credentials_file is provided.
        """
        if target_dir is None:
            target_dir = self.default_target_dir
            
        target_dir.mkdir(exist_ok=True)
        logger.info(f"Downloading documents to: {target_dir}")
        
        # Get credentials from environment (set by DatabaseManager)
        username = os.getenv("DB_USER", "")
        password = os.getenv("DB_PASSWORD", "")
        
        if not username or not password:
            logger.warning("No credentials found in environment. Downloads may fail.")
        
        results = []
        for idx, row in document_df.iterrows():
            name = row["DokumentName"]
            link = row["DokumentLink"]
            geschaeft_nr = row.get("GeschaeftNr", None)
            finma_id = row.get("FinmaID", None)
            out = target_dir / name
            
            try:
                logger.info(f"Downloading: {name}")
                r = requests.get(
                    link,
                    auth=HttpNtlmAuth(username, password),
                    allow_redirects=True,
                )
                r.raise_for_status()
                out.write_bytes(r.content)
                results.append((name, out, geschaeft_nr, finma_id))
                logger.info(f"  ✓ Saved to: {out}")
            except Exception as e:
                logger.error(f"  ✗ Failed to download {name}: {e}")
                
        logger.info(f"Successfully downloaded {len(results)}/{len(document_df)} documents")
        return results

    def load(self, target_dir: Path = None) -> List[Tuple[str, Path, str, str]]:
        """Load all relevant ORSA documents.
        
        This is the main entry point that:
        1. Retrieves document metadata from database
        2. Filters for relevant documents
        3. Downloads documents locally
        
        Args:
            target_dir: Directory to save files (default: orsa_response_files/)
            
        Returns:
            List of tuples (document_name, file_path, geschaeft_nr, finma_id)
        """
        logger.info("Starting ORSA document loading process")
        all_document_metadata_df = self.get_document_metadata()
        relevant_document_metadata_df = self.filter_relevant(all_document_metadata_df)
        documents = self.download_documents(relevant_document_metadata_df, target_dir)
        logger.info(f"Document loading complete: {len(documents)} files ready")
        return documents


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s %(message)s"
    )
    sourcer = ORSADocumentSourcer()
    documents = sourcer.load()
    logger.info(f"Loaded documents: {documents}")
