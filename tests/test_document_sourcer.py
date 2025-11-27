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
            "FinmaID": [
                "10001",
                "10002",
                "10003",
                "10004",
            ],
            "GeschaeftNr": [
                "GNR001",
                "GNR002",
                "GNR003",
                "GNR004",
            ],
        }
    )


class TestORSADocumentSourcerInit:
    """Tests for ORSADocumentSourcer initialization."""

    def test_init_with_existing_credentials_file(self, tmp_path):
        """Test initialization with existing credentials file."""
        cred_file = tmp_path / "credentials.env"
        cred_file.write_text("username=testuser\npassword=testpass\n")

        sourcer = ORSADocumentSourcer(cred_file="credentials.env")
        assert sourcer.cred_file.name == "credentials.env"

    def test_init_without_credentials_file(self):
        """Test initialization without credentials file."""
        sourcer = ORSADocumentSourcer()
        # Should initialize without errors even if file doesn't exist
        assert sourcer.cred_file.name == "credentials.env"
        assert sourcer.berichtsjahr == 2026  # Default berichtsjahr

    def test_init_with_custom_berichtsjahr(self):
        """Test initialization with custom berichtsjahr."""
        sourcer = ORSADocumentSourcer(berichtsjahr=2027)
        assert sourcer.berichtsjahr == 2027

    def test_directory_structure(self):
        """Test that directory paths are correctly calculated."""
        sourcer = ORSADocumentSourcer()
        assert sourcer.base_dir.name == "sourcing"
        assert sourcer.default_target_dir.name == "orsa_response_files"


class TestORSADocumentSourcerCredentials:
    """Tests for credential usage from environment variables."""

    def test_credentials_from_environment(self, monkeypatch):
        """Test that credentials are read from environment variables."""
        monkeypatch.setenv("DB_USER", "user123")
        monkeypatch.setenv("DB_PASSWORD", "pass456")

        sourcer = ORSADocumentSourcer()
        
        # Verify credentials are available in environment
        assert os.getenv("DB_USER") == "user123"
        assert os.getenv("DB_PASSWORD") == "pass456"

    def test_missing_credentials_in_environment(self, monkeypatch):
        """Test handling of missing credentials."""
        monkeypatch.delenv("DB_USER", raising=False)
        monkeypatch.delenv("DB_PASSWORD", raising=False)

        sourcer = ORSADocumentSourcer()
        
        # Should initialize without errors
        assert os.getenv("DB_USER", "") == ""
        assert os.getenv("DB_PASSWORD", "") == ""


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
    def test_run_query_success(self, mock_db_manager_class, monkeypatch):
        """Test successful query execution."""
        sourcer = ORSADocumentSourcer()
        monkeypatch.setenv("DB_USER", "testuser")
        monkeypatch.setenv("DB_PASSWORD", "testpass")

        mock_df = pd.DataFrame({"col1": [1, 2], "col2": ["a", "b"]})
        
        # Mock the DatabaseManager instance
        mock_db_instance = MagicMock()
        mock_db_instance.execute_query.return_value = mock_df
        mock_db_manager_class.return_value = mock_db_instance

        result = sourcer._run_query("SELECT * FROM test;")

        assert result.equals(mock_df)
        # Verify DatabaseManager was called with the credentials file
        mock_db_manager_class.assert_called_once_with(
            server="frbdata.finma.ch",
            database="GBB_Reporting",
            credentials_file=sourcer.cred_file
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


class TestORSADocumentSourcerDownload:
    """Tests for document downloading."""

    @patch("requests.get")
    def test_download_documents_success(self, mock_get, tmp_path, monkeypatch):
        """Test successful document download."""
        mock_response = Mock()
        mock_response.content = b"file content"
        mock_response.raise_for_status = Mock()
        mock_get.return_value = mock_response

        df = pd.DataFrame(
            {
                "DokumentName": ["test_doc.xlsx"],
                "DokumentLink": ["http://example.com/test_doc.xlsx"],
                "GeschaeftNr": ["GNR123"],
                "FinmaID": ["10001"],
            }
        )

        sourcer = ORSADocumentSourcer()
        monkeypatch.setenv("DB_USER", "testuser")
        monkeypatch.setenv("DB_PASSWORD", "testpass")

        results = sourcer.download_documents(df, target_dir=tmp_path)

        assert len(results) == 1
        assert results[0][0] == "test_doc.xlsx"
        assert results[0][1] == tmp_path / "test_doc.xlsx"
        assert results[0][2] == "GNR123"
        assert results[0][3] == "10001"
        assert (tmp_path / "test_doc.xlsx").exists()
        assert (tmp_path / "test_doc.xlsx").read_bytes() == b"file content"

    @patch("requests.get")
    def test_download_documents_multiple(self, mock_get, tmp_path, monkeypatch):
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
        monkeypatch.setenv("DB_USER", "testuser")
        monkeypatch.setenv("DB_PASSWORD", "testpass")

        results = sourcer.download_documents(df, target_dir=tmp_path)

        assert len(results) == 3
        assert all(name in ["doc1.xlsx", "doc2.xlsx", "doc3.xlsx"] for name, _, _, _ in results)
        assert all((tmp_path / name).exists() for name, _, _, _ in results)

    @patch("requests.get")
    def test_download_documents_with_failure(self, mock_get, tmp_path, monkeypatch):
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
        monkeypatch.setenv("DB_USER", "testuser")
        monkeypatch.setenv("DB_PASSWORD", "testpass")

        results = sourcer.download_documents(df, target_dir=tmp_path)

        assert len(results) == 2  # Only successful downloads
        assert all(name in ["doc1.xlsx", "doc3.xlsx"] for name, _, _, _ in results)

    @patch("requests.get")
    def test_download_documents_default_directory(self, mock_get, tmp_path, monkeypatch):
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
        monkeypatch.setenv("DB_USER", "testuser")
        monkeypatch.setenv("DB_PASSWORD", "testpass")
        sourcer.default_target_dir = tmp_path / "test_orsa_files"

        results = sourcer.download_documents(df)

        assert len(results) == 1
        assert results[0][1].parent == sourcer.default_target_dir
        assert (sourcer.default_target_dir / "test_doc.xlsx").exists()

    @patch("requests.get")
    def test_download_documents_auth_used(self, mock_get, tmp_path, monkeypatch):
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
        monkeypatch.setenv("DB_USER", "testuser")
        monkeypatch.setenv("DB_PASSWORD", "testpass")

        sourcer.download_documents(df, target_dir=tmp_path)

        mock_get.assert_called_once()
        call_kwargs = mock_get.call_args[1]
        assert "auth" in call_kwargs
        assert isinstance(call_kwargs["auth"], HttpNtlmAuth)
        assert call_kwargs["allow_redirects"] is True


class TestORSADocumentSourcerLoad:
    """Tests for the main load method."""

    @patch.object(ORSADocumentSourcer, "download_documents")
    @patch.object(ORSADocumentSourcer, "get_document_metadata")
    def test_load_integration(
        self,
        mock_get_metadata,
        mock_download,
        sample_metadata_df,
        tmp_path,
    ):
        """Test the full load workflow."""
        expected_results = [
            ("2026_Company_A_ORSA-Formular.xlsx", tmp_path / "doc1.xlsx", "GNR001", "10001"),
            ("2027_Company_C_ORSA-Formular.xlsx", tmp_path / "doc3.xlsx", "GNR003", "10003"),
        ]

        mock_get_metadata.return_value = sample_metadata_df
        mock_download.return_value = expected_results

        sourcer = ORSADocumentSourcer()
        results = sourcer.load(target_dir=tmp_path)

        assert results == expected_results
        mock_get_metadata.assert_called_once()
        mock_download.assert_called_once()
        # Check that download_documents was called with correct target_dir
        call_args = mock_download.call_args[0]
        assert call_args[1] == tmp_path

    @patch.object(ORSADocumentSourcer, "download_documents")
    @patch.object(ORSADocumentSourcer, "get_document_metadata")
    def test_load_default_target_dir(
        self, mock_get_metadata, mock_download, sample_metadata_df
    ):
        """Test load with default target directory."""
        mock_get_metadata.return_value = sample_metadata_df
        mock_download.return_value = []

        sourcer = ORSADocumentSourcer()
        sourcer.load()

        mock_download.assert_called_once()
        # Check that download_documents was called with correct target_dir
        call_args = mock_download.call_args[0]
        assert call_args[1] is None

    @patch.object(ORSADocumentSourcer, "download_documents")
    @patch.object(ORSADocumentSourcer, "get_document_metadata")
    def test_load_empty_results(
        self, mock_get_metadata, mock_download, tmp_path
    ):
        """Test load when no documents match criteria."""
        empty_df = pd.DataFrame(columns=["DokumentName", "DokumentLink"])

        mock_get_metadata.return_value = empty_df
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
