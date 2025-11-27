"""Tests for the ORSADocumentSourcer class."""

import os
from pathlib import Path
from unittest.mock import Mock, patch, MagicMock

import pandas as pd
import pytest

from orsa_analysis.sourcing.document_sourcer import ORSADocumentSourcer


@pytest.fixture
def mock_env_file(tmp_path):
    """Create a temporary credentials file."""
    cred_file = tmp_path / "test_credentials.env"
    cred_file.write_text("username=testuser\npassword=testpass\n")
    return cred_file


@pytest.fixture
def mock_sql_file(tmp_path):
    """Create a temporary SQL file."""
    sql_dir = tmp_path / "sql"
    sql_dir.mkdir()
    sql_file = sql_dir / "source_orsa_dokument_metadata.sql"
    sql_file.write_text("SELECT * FROM test_table;")
    return sql_file


@pytest.fixture
def sample_metadata_df():
    """Create sample document metadata DataFrame."""
    return pd.DataFrame(
        {
            "DokumentID": [1, 2, 3, 4],
            "DokumentName": [
                "2026_Company_A_ORSA-Formular.xlsx",
                "2025_Company_B_ORSA-Formular.xlsx",
                "2027_Company_C_ORSA-Formular.xlsx",
                "2026_Company_D_Other_Document.xlsx",
            ],
            "DokumentLink": [
                "http://example.com/doc1.xlsx",
                "http://example.com/doc2.xlsx",
                "http://example.com/doc3.xlsx",
                "http://example.com/doc4.xlsx",
            ],
        }
    )


class TestORSADocumentSourcerInit:
    """Tests for ORSADocumentSourcer initialization."""

    @patch("orsa_analysis.sourcing.document_sourcer.load_dotenv")
    @patch("os.getenv")
    def test_init_with_existing_credentials_file(
        self, mock_getenv, mock_load_dotenv, tmp_path
    ):
        """Test initialization with existing credentials file."""
        cred_file = tmp_path / "credentials.env"
        cred_file.write_text("username=testuser\npassword=testpass\n")

        mock_getenv.side_effect = lambda key, default="": {
            "username": "testuser",
            "password": "testpass",
        }.get(key, default)

        with patch.object(Path, "exists", return_value=True):
            with patch.object(
                ORSADocumentSourcer, "_load_credentials", return_value=None
            ):
                sourcer = ORSADocumentSourcer(cred_file="credentials.env")
                assert sourcer.cred_file.name == "credentials.env"

    @patch("orsa_analysis.sourcing.document_sourcer.load_dotenv")
    @patch("os.getenv")
    def test_init_without_credentials_file(self, mock_getenv, mock_load_dotenv):
        """Test initialization without credentials file."""
        mock_getenv.side_effect = lambda key, default="": {
            "username": "",
            "password": "",
        }.get(key, default)

        with patch.object(Path, "exists", return_value=False):
            sourcer = ORSADocumentSourcer()
            assert sourcer.username == ""
            assert sourcer.password == ""

    def test_directory_structure(self):
        """Test that directory paths are correctly calculated."""
        sourcer = ORSADocumentSourcer()
        assert sourcer.base_dir.name == "sourcing"
        assert sourcer.default_target_dir.name == "orsa_response_files"


class TestORSADocumentSourcerCredentials:
    """Tests for credential loading."""

    @patch("os.getenv")
    @patch("orsa_analysis.sourcing.document_sourcer.load_dotenv")
    def test_load_credentials_success(self, mock_load_dotenv, mock_getenv, tmp_path):
        """Test successful credential loading."""
        cred_file = tmp_path / "credentials.env"
        cred_file.write_text("username=user123\npassword=pass456\n")

        mock_getenv.side_effect = lambda key, default="": {
            "username": "user123",
            "password": "pass456",
        }.get(key, default)

        with patch.object(Path, "exists", return_value=True):
            sourcer = ORSADocumentSourcer()
            sourcer.cred_file = cred_file
            sourcer._load_credentials()

            assert sourcer.username == "user123"
            assert sourcer.password == "pass456"

    @patch("os.getenv")
    @patch("orsa_analysis.sourcing.document_sourcer.load_dotenv")
    def test_load_credentials_missing_values(
        self, mock_load_dotenv, mock_getenv, tmp_path
    ):
        """Test credential loading with missing values."""
        cred_file = tmp_path / "credentials.env"
        cred_file.write_text("username=\npassword=\n")

        mock_getenv.side_effect = lambda key, default="": default

        with patch.object(Path, "exists", return_value=True):
            sourcer = ORSADocumentSourcer()
            sourcer.cred_file = cred_file
            sourcer._load_credentials()

            assert sourcer.username == ""
            assert sourcer.password == ""


class TestORSADocumentSourcerQuery:
    """Tests for SQL query loading and execution."""

    def test_load_query_success(self, tmp_path):
        """Test successful query loading."""
        sql_dir = tmp_path / "sql"
        sql_dir.mkdir()
        query_file = sql_dir / "source_orsa_dokument_metadata.sql"
        query_content = "SELECT * FROM documents WHERE active = 1;"
        query_file.write_text(query_content)

        sourcer = ORSADocumentSourcer()
        sourcer.base_dir = tmp_path / "src" / "orsa_analysis" / "sourcing"

        with patch.object(
            Path, "parent", new_callable=lambda: property(lambda self: tmp_path)
        ):
            query = sourcer._load_query("source_orsa_dokument_metadata")
            assert query == query_content

    def test_load_query_file_not_found(self):
        """Test query loading when file doesn't exist."""
        sourcer = ORSADocumentSourcer()

        with pytest.raises(FileNotFoundError, match="Query file not found"):
            sourcer._load_query("nonexistent_query")

    @patch("orsa_analysis.core.database_manager.DatabaseManager")
    def test_run_query_success(self, mock_db_manager_class):
        """Test successful query execution."""
        sourcer = ORSADocumentSourcer()
        sourcer.username = "testuser"
        sourcer.password = "testpass"

        mock_df = pd.DataFrame({"col1": [1, 2], "col2": ["a", "b"]})
        
        # Mock the DatabaseManager instance
        mock_db_instance = MagicMock()
        mock_db_instance.execute_query.return_value = mock_df
        mock_db_manager_class.return_value = mock_db_instance

        result = sourcer._run_query("SELECT * FROM test;")

        assert result.equals(mock_df)
        mock_db_manager_class.assert_called_once_with(
            server="frbdata.finma.ch",
            database="GBB_Reporting",
            username="Finma\\testuser",
            password="testpass"
        )
        mock_db_instance.execute_query.assert_called_once_with("SELECT * FROM test;")


class TestORSADocumentSourcerMetadata:
    """Tests for document metadata retrieval."""

    @patch.object(ORSADocumentSourcer, "_run_query")
    @patch.object(ORSADocumentSourcer, "_load_query")
    def test_get_document_metadata_success(
        self, mock_load_query, mock_run_query, sample_metadata_df
    ):
        """Test successful metadata retrieval."""
        mock_load_query.return_value = "SELECT * FROM documents;"
        mock_run_query.return_value = sample_metadata_df

        sourcer = ORSADocumentSourcer()
        result = sourcer.get_document_metadata()

        assert len(result) == 4
        assert "DokumentName" in result.columns
        assert "DokumentLink" in result.columns
        mock_load_query.assert_called_once_with("source_orsa_dokument_metadata")
        mock_run_query.assert_called_once()


class TestORSADocumentSourcerFiltering:
    """Tests for document filtering."""

    def test_filter_relevant_orsa_2026_and_above(self, sample_metadata_df):
        """Test filtering for ORSA documents year >= 2026."""
        sourcer = ORSADocumentSourcer()
        result = sourcer.filter_relevant(sample_metadata_df)

        assert len(result) == 2  # Only 2026 and 2027 ORSA documents
        assert all("_ORSA-Formular" in name for name in result["DokumentName"])
        assert all(year >= 2026 for year in result["reporting_year"])

    def test_filter_relevant_empty_dataframe(self):
        """Test filtering with empty DataFrame."""
        sourcer = ORSADocumentSourcer()
        empty_df = pd.DataFrame(columns=["DokumentName", "DokumentLink"])

        result = sourcer.filter_relevant(empty_df)

        assert len(result) == 0
        assert "reporting_year" in result.columns

    def test_filter_relevant_no_orsa_documents(self):
        """Test filtering when no ORSA documents match criteria."""
        df = pd.DataFrame(
            {
                "DokumentName": [
                    "2024_Company_A_Other.xlsx",
                    "2025_Company_B_Report.xlsx",
                ],
                "DokumentLink": [
                    "http://example.com/doc1.xlsx",
                    "http://example.com/doc2.xlsx",
                ],
            }
        )

        sourcer = ORSADocumentSourcer()
        result = sourcer.filter_relevant(df)

        assert len(result) == 0

    def test_filter_relevant_year_extraction(self):
        """Test year extraction from document names."""
        df = pd.DataFrame(
            {
                "DokumentName": [
                    "2026_ORSA-Formular_Test.xlsx",
                    "Test_2027_ORSA-Formular.xlsx",
                    "2028_Test_ORSA-Formular.xlsx",
                ],
                "DokumentLink": ["http://example.com/doc.xlsx"] * 3,
            }
        )

        sourcer = ORSADocumentSourcer()
        result = sourcer.filter_relevant(df)

        assert len(result) == 3
        assert list(result["reporting_year"]) == [2026.0, 2027.0, 2028.0]


class TestORSADocumentSourcerDownload:
    """Tests for document downloading."""

    @patch("requests.get")
    def test_download_documents_success(self, mock_get, tmp_path):
        """Test successful document download."""
        mock_response = Mock()
        mock_response.content = b"file content"
        mock_response.raise_for_status = Mock()
        mock_get.return_value = mock_response

        df = pd.DataFrame(
            {
                "DokumentName": ["test_doc.xlsx"],
                "DokumentLink": ["http://example.com/test_doc.xlsx"],
            }
        )

        sourcer = ORSADocumentSourcer()
        sourcer.username = "testuser"
        sourcer.password = "testpass"

        results = sourcer.download_documents(df, target_dir=tmp_path)

        assert len(results) == 1
        assert results[0][0] == "test_doc.xlsx"
        assert results[0][1] == tmp_path / "test_doc.xlsx"
        assert (tmp_path / "test_doc.xlsx").exists()
        assert (tmp_path / "test_doc.xlsx").read_bytes() == b"file content"

    @patch("requests.get")
    def test_download_documents_multiple(self, mock_get, tmp_path):
        """Test downloading multiple documents."""
        mock_response = Mock()
        mock_response.content = b"file content"
        mock_response.raise_for_status = Mock()
        mock_get.return_value = mock_response

        df = pd.DataFrame(
            {
                "DokumentName": ["doc1.xlsx", "doc2.xlsx", "doc3.xlsx"],
                "DokumentLink": [
                    "http://example.com/doc1.xlsx",
                    "http://example.com/doc2.xlsx",
                    "http://example.com/doc3.xlsx",
                ],
            }
        )

        sourcer = ORSADocumentSourcer()
        sourcer.username = "testuser"
        sourcer.password = "testpass"

        results = sourcer.download_documents(df, target_dir=tmp_path)

        assert len(results) == 3
        assert all(name in ["doc1.xlsx", "doc2.xlsx", "doc3.xlsx"] for name, _ in results)
        assert all((tmp_path / name).exists() for name, _ in results)

    @patch("requests.get")
    def test_download_documents_with_failure(self, mock_get, tmp_path):
        """Test download with some failures."""
        def side_effect(*args, **kwargs):
            if "doc2" in args[0]:
                raise Exception("Download failed")
            mock_response = Mock()
            mock_response.content = b"file content"
            mock_response.raise_for_status = Mock()
            return mock_response

        mock_get.side_effect = side_effect

        df = pd.DataFrame(
            {
                "DokumentName": ["doc1.xlsx", "doc2.xlsx", "doc3.xlsx"],
                "DokumentLink": [
                    "http://example.com/doc1.xlsx",
                    "http://example.com/doc2.xlsx",
                    "http://example.com/doc3.xlsx",
                ],
            }
        )

        sourcer = ORSADocumentSourcer()
        sourcer.username = "testuser"
        sourcer.password = "testpass"

        results = sourcer.download_documents(df, target_dir=tmp_path)

        assert len(results) == 2  # Only successful downloads
        assert all(name in ["doc1.xlsx", "doc3.xlsx"] for name, _ in results)

    @patch("requests.get")
    def test_download_documents_default_directory(self, mock_get, tmp_path):
        """Test download to default directory."""
        mock_response = Mock()
        mock_response.content = b"file content"
        mock_response.raise_for_status = Mock()
        mock_get.return_value = mock_response

        df = pd.DataFrame(
            {
                "DokumentName": ["test_doc.xlsx"],
                "DokumentLink": ["http://example.com/test_doc.xlsx"],
            }
        )

        sourcer = ORSADocumentSourcer()
        sourcer.username = "testuser"
        sourcer.password = "testpass"
        sourcer.default_target_dir = tmp_path / "test_orsa_files"

        results = sourcer.download_documents(df)

        assert len(results) == 1
        assert results[0][1].parent == sourcer.default_target_dir
        assert (sourcer.default_target_dir / "test_doc.xlsx").exists()

    @patch("requests.get")
    def test_download_documents_auth_used(self, mock_get, tmp_path):
        """Test that NTLM authentication is used."""
        from requests_ntlm import HttpNtlmAuth

        mock_response = Mock()
        mock_response.content = b"file content"
        mock_response.raise_for_status = Mock()
        mock_get.return_value = mock_response

        df = pd.DataFrame(
            {
                "DokumentName": ["test_doc.xlsx"],
                "DokumentLink": ["http://example.com/test_doc.xlsx"],
            }
        )

        sourcer = ORSADocumentSourcer()
        sourcer.username = "testuser"
        sourcer.password = "testpass"

        sourcer.download_documents(df, target_dir=tmp_path)

        mock_get.assert_called_once()
        call_kwargs = mock_get.call_args[1]
        assert "auth" in call_kwargs
        assert isinstance(call_kwargs["auth"], HttpNtlmAuth)
        assert call_kwargs["allow_redirects"] is True


class TestORSADocumentSourcerLoad:
    """Tests for the main load method."""

    @patch.object(ORSADocumentSourcer, "download_documents")
    @patch.object(ORSADocumentSourcer, "filter_relevant")
    @patch.object(ORSADocumentSourcer, "get_document_metadata")
    def test_load_integration(
        self,
        mock_get_metadata,
        mock_filter,
        mock_download,
        sample_metadata_df,
        tmp_path,
    ):
        """Test the full load workflow."""
        filtered_df = sample_metadata_df[sample_metadata_df["DokumentID"].isin([1, 3])]
        expected_results = [
            ("2026_Company_A_ORSA-Formular.xlsx", tmp_path / "doc1.xlsx"),
            ("2027_Company_C_ORSA-Formular.xlsx", tmp_path / "doc3.xlsx"),
        ]

        mock_get_metadata.return_value = sample_metadata_df
        mock_filter.return_value = filtered_df
        mock_download.return_value = expected_results

        sourcer = ORSADocumentSourcer()
        results = sourcer.load(target_dir=tmp_path)

        assert results == expected_results
        mock_get_metadata.assert_called_once()
        mock_filter.assert_called_once_with(sample_metadata_df)
        mock_download.assert_called_once_with(filtered_df, tmp_path)

    @patch.object(ORSADocumentSourcer, "download_documents")
    @patch.object(ORSADocumentSourcer, "filter_relevant")
    @patch.object(ORSADocumentSourcer, "get_document_metadata")
    def test_load_default_target_dir(
        self, mock_get_metadata, mock_filter, mock_download, sample_metadata_df
    ):
        """Test load with default target directory."""
        mock_get_metadata.return_value = sample_metadata_df
        mock_filter.return_value = sample_metadata_df
        mock_download.return_value = []

        sourcer = ORSADocumentSourcer()
        sourcer.load()

        mock_download.assert_called_once_with(sample_metadata_df, None)

    @patch.object(ORSADocumentSourcer, "download_documents")
    @patch.object(ORSADocumentSourcer, "filter_relevant")
    @patch.object(ORSADocumentSourcer, "get_document_metadata")
    def test_load_empty_results(
        self, mock_get_metadata, mock_filter, mock_download, tmp_path
    ):
        """Test load when no documents match criteria."""
        empty_df = pd.DataFrame(columns=["DokumentName", "DokumentLink"])

        mock_get_metadata.return_value = empty_df
        mock_filter.return_value = empty_df
        mock_download.return_value = []

        sourcer = ORSADocumentSourcer()
        results = sourcer.load(target_dir=tmp_path)

        assert results == []


class TestORSADocumentSourcerEnvironment:
    """Tests for environment configuration."""

    def test_environment_variables_set(self):
        """Test that required environment variables are set."""
        assert os.environ.get("NO_PROXY") == "finma.ch"
        assert os.environ.get("REQUESTS_CA_BUNDLE") == "SwisscomRootCore.crt"
