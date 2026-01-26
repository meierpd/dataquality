"""Unit tests for the SharePoint uploader module."""

import pytest
from unittest.mock import Mock, patch, MagicMock
from pathlib import Path

from orsa_analysis.reporting.sharepoint_uploader import SharePointUploader


@pytest.fixture
def mock_env_vars(monkeypatch):
    """Mock environment variables for testing."""
    monkeypatch.setenv("DB_USER", "test_user")
    monkeypatch.setenv("DB_PASSWORD", "test_password")


@pytest.fixture
def uploader(mock_env_vars, tmp_path):
    """Create a SharePointUploader instance for testing."""
    # Create a mock certificate file
    cert_path = tmp_path / "test_cert.crt"
    cert_path.write_text("MOCK CERTIFICATE")
    
    return SharePointUploader(ca_cert_path=cert_path)


class TestSharePointUploader:
    """Test cases for SharePointUploader."""
    
    def test_init_with_credentials(self, mock_env_vars):
        """Test initialization with environment credentials."""
        uploader = SharePointUploader()
        assert uploader.user == "test_user"
        assert uploader.password == "test_password"
        assert uploader.auth is not None
    
    def test_init_without_credentials(self, monkeypatch):
        """Test initialization without credentials logs warning."""
        monkeypatch.delenv("DB_USER", raising=False)
        monkeypatch.delenv("DB_PASSWORD", raising=False)
        
        uploader = SharePointUploader()
        assert uploader.user is None
        assert uploader.password is None
    
    @patch('orsa_analysis.reporting.sharepoint_uploader.requests.get')
    def test_resolve_folder_from_link(self, mock_get, uploader):
        """Test folder resolution from download link."""
        # Mock the response
        mock_response = Mock()
        mock_response.url = "https://stb.finma.ch/sharepoint/documents/G01410166/test.pdf"
        mock_response.raise_for_status = Mock()
        mock_get.return_value = mock_response
        
        download_link = "https://stb.finma.ch:30017/redirectToSharePoint/documents/G01410166/%2F%5Btest.pdf"
        folder_url = uploader.resolve_folder_from_link(download_link)
        
        assert folder_url == "https://stb.finma.ch/sharepoint/documents/G01410166"
        mock_get.assert_called_once()
    
    @patch('orsa_analysis.reporting.sharepoint_uploader.requests.head')
    def test_file_exists_returns_true(self, mock_head, uploader):
        """Test file_exists returns True when file exists."""
        # Mock the response
        mock_response = Mock()
        mock_response.status_code = 200
        mock_head.return_value = mock_response
        
        folder_url = "https://stb.finma.ch/sharepoint/documents/G01410166"
        exists = uploader.file_exists(folder_url, "test.xlsx")
        
        assert exists is True
        mock_head.assert_called_once()
    
    @patch('orsa_analysis.reporting.sharepoint_uploader.requests.head')
    def test_file_exists_returns_false(self, mock_head, uploader):
        """Test file_exists returns False when file doesn't exist."""
        # Mock the response
        mock_response = Mock()
        mock_response.status_code = 404
        mock_head.return_value = mock_response
        
        folder_url = "https://stb.finma.ch/sharepoint/documents/G01410166"
        exists = uploader.file_exists(folder_url, "test.xlsx")
        
        assert exists is False
    
    @patch('orsa_analysis.reporting.sharepoint_uploader.requests.put')
    @patch('orsa_analysis.reporting.sharepoint_uploader.requests.get')
    def test_upload_new_file(self, mock_get, mock_put, uploader, tmp_path):
        """Test uploading a new file (file doesn't exist)."""
        # Create a test file
        test_file = tmp_path / "test_report.xlsx"
        test_file.write_bytes(b"test content")
        
        # Mock folder resolution
        mock_get_response = Mock()
        mock_get_response.url = "https://stb.finma.ch/sharepoint/documents/G01410166/original.pdf"
        mock_get_response.raise_for_status = Mock()
        mock_get.return_value = mock_get_response
        
        # Mock PUT response (file created)
        mock_put_response = Mock()
        mock_put_response.status_code = 201
        mock_put.return_value = mock_put_response
        
        download_link = "https://stb.finma.ch:30017/redirectToSharePoint/documents/G01410166/%2F%5Boriginal.pdf"
        result = uploader.upload(download_link, str(test_file), skip_if_exists=False)
        
        assert result["success"] is True
        assert result["skipped"] is False
        assert result["message"] == "File created."
        mock_put.assert_called_once()
    
    @patch('orsa_analysis.reporting.sharepoint_uploader.requests.head')
    @patch('orsa_analysis.reporting.sharepoint_uploader.requests.get')
    def test_upload_skip_if_exists(self, mock_get, mock_head, uploader, tmp_path):
        """Test that upload is skipped when file already exists."""
        # Create a test file
        test_file = tmp_path / "test_report.xlsx"
        test_file.write_bytes(b"test content")
        
        # Mock folder resolution
        mock_get_response = Mock()
        mock_get_response.url = "https://stb.finma.ch/sharepoint/documents/G01410166/original.pdf"
        mock_get_response.raise_for_status = Mock()
        mock_get.return_value = mock_get_response
        
        # Mock HEAD response (file exists)
        mock_head_response = Mock()
        mock_head_response.status_code = 200
        mock_head.return_value = mock_head_response
        
        download_link = "https://stb.finma.ch:30017/redirectToSharePoint/documents/G01410166/%2F%5Boriginal.pdf"
        result = uploader.upload(download_link, str(test_file), skip_if_exists=True)
        
        assert result["success"] is True
        assert result["skipped"] is True
        assert "already exists" in result["message"]
        mock_head.assert_called_once()
    
    @patch('orsa_analysis.reporting.sharepoint_uploader.requests.put')
    @patch('orsa_analysis.reporting.sharepoint_uploader.requests.get')
    def test_upload_unauthorized(self, mock_get, mock_put, uploader, tmp_path):
        """Test handling of unauthorized upload."""
        # Create a test file
        test_file = tmp_path / "test_report.xlsx"
        test_file.write_bytes(b"test content")
        
        # Mock folder resolution
        mock_get_response = Mock()
        mock_get_response.url = "https://stb.finma.ch/sharepoint/documents/G01410166/original.pdf"
        mock_get_response.raise_for_status = Mock()
        mock_get.return_value = mock_get_response
        
        # Mock PUT response (unauthorized)
        mock_put_response = Mock()
        mock_put_response.status_code = 401
        mock_put.return_value = mock_put_response
        
        download_link = "https://stb.finma.ch:30017/redirectToSharePoint/documents/G01410166/%2F%5Boriginal.pdf"
        result = uploader.upload(download_link, str(test_file), skip_if_exists=False)
        
        assert result["success"] is False
        assert result["message"] == "Unauthorized."
    
    @patch('orsa_analysis.reporting.sharepoint_uploader.requests.get')
    def test_upload_folder_resolution_error(self, mock_get, uploader, tmp_path):
        """Test handling of folder resolution errors."""
        # Create a test file
        test_file = tmp_path / "test_report.xlsx"
        test_file.write_bytes(b"test content")
        
        # Mock folder resolution failure
        mock_get.side_effect = Exception("Network error")
        
        download_link = "https://stb.finma.ch:30017/redirectToSharePoint/documents/G01410166/%2F%5Boriginal.pdf"
        result = uploader.upload(download_link, str(test_file), skip_if_exists=False)
        
        assert result["success"] is False
        assert "Upload failed" in result["message"]
