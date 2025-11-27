"""Document sourcer for retrieving ORSA documents from the database."""

import os
import logging
from pathlib import Path
from typing import List, Tuple

import pandas as pd
import requests
from dotenv import load_dotenv
from requests_ntlm import HttpNtlmAuth

logger = logging.getLogger(__name__)

# Configure environment for FINMA network
os.environ["NO_PROXY"] = "finma.ch"
os.environ["REQUESTS_CA_BUNDLE"] = "SwisscomRootCore.crt"


class ORSADocumentSourcer:
    """Source ORSA documents from the database and download them locally.
    
    This class handles:
    - Loading credentials from environment file
    - Querying the database for document metadata
    - Filtering relevant ORSA documents
    - Downloading documents from SharePoint links
    
    Attributes:
        username: Database username
        password: Database password
        base_dir: Base directory of the sourcing module
        cred_file: Path to credentials file
        default_target_dir: Directory where documents are downloaded
    """

    def __init__(self, cred_file: str = "credentials.env"):
        """Initialize the document sourcer.
        
        Args:
            cred_file: Name of credentials file (relative to project root)
        """
        self.base_dir = Path(__file__).resolve().parent
        self.cred_file = self.base_dir.parent.parent.parent / cred_file
        self.default_target_dir = self.base_dir.parent.parent.parent / "data" / "orsa_response_files"
        self._load_credentials()

    def _load_credentials(self) -> None:
        """Load username and password from credentials file."""
        if not self.cred_file.exists():
            logger.warning(f"Credentials file not found: {self.cred_file}")
            self.username = os.getenv("username", "")
            self.password = os.getenv("password", "")
            return
            
        load_dotenv(self.cred_file.as_posix())
        self.username = os.getenv("username", "")
        self.password = os.getenv("password", "")
        
        if not self.username or not self.password:
            logger.warning("Username or password not set in credentials file")

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
        """Retrieve all document metadata from the database.
        
        Returns:
            DataFrame with columns: DokumentName, DokumentLink, etc.
        """
        q = self._load_query("source_orsa_dokument_metadata")
        document_metadata_df = self._run_query(q)
        logger.info(f"Retrieved {len(document_metadata_df)} documents from database")
        return document_metadata_df

    def filter_relevant(self, df: pd.DataFrame) -> pd.DataFrame:
        """Filter for relevant ORSA documents.
        
        Filters for:
        - Documents containing "_ORSA-Formular" in name
        - Reporting year >= 2026
        
        Args:
            df: DataFrame with document metadata
            
        Returns:
            Filtered DataFrame
        """
        df = df.copy()
        df["reporting_year"] = (
            df["DokumentName"].str.extract(r"(\d{4})")[0].astype(float)
        )

        # Filter ORSA docs and year >= 2026
        mask = df["DokumentName"].str.lower().str.contains("_orsa-formular") & (
            df["reporting_year"] >= 2026
        )
        filtered_df = df[mask]
        logger.info(f"Filtered to {len(filtered_df)} relevant ORSA documents")
        return filtered_df

    def download_documents(
        self, document_df: pd.DataFrame, target_dir: Path = None
    ) -> List[Tuple[str, Path, str]]:
        """Download documents from SharePoint links.
        
        Args:
            document_df: DataFrame with DokumentName, DokumentLink, and GeschaeftNr columns
            target_dir: Directory to save files (default: orsa_response_files/)
            
        Returns:
            List of tuples (document_name, file_path, geschaeft_nr)
        """
        if target_dir is None:
            target_dir = self.default_target_dir
            
        target_dir.mkdir(exist_ok=True)
        logger.info(f"Downloading documents to: {target_dir}")
        
        results = []
        for idx, row in document_df.iterrows():
            name = row["DokumentName"]
            link = row["DokumentLink"]
            geschaeft_nr = row.get("GeschaeftNr", None)
            out = target_dir / name
            
            try:
                logger.info(f"Downloading: {name}")
                r = requests.get(
                    link,
                    auth=HttpNtlmAuth(self.username, self.password),
                    allow_redirects=True,
                )
                r.raise_for_status()
                out.write_bytes(r.content)
                results.append((name, out, geschaeft_nr))
                logger.info(f"  ✓ Saved to: {out}")
            except Exception as e:
                logger.error(f"  ✗ Failed to download {name}: {e}")
                
        logger.info(f"Successfully downloaded {len(results)}/{len(document_df)} documents")
        return results

    def load(self, target_dir: Path = None) -> List[Tuple[str, Path, str]]:
        """Load all relevant ORSA documents.
        
        This is the main entry point that:
        1. Retrieves document metadata from database
        2. Filters for relevant documents
        3. Downloads documents locally
        
        Args:
            target_dir: Directory to save files (default: orsa_response_files/)
            
        Returns:
            List of tuples (document_name, file_path, geschaeft_nr)
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
