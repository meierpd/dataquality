"""Unit tests for the database module."""

import pytest
from datetime import datetime
from unittest.mock import Mock, MagicMock, patch
from pathlib import Path
import pandas as pd

from orsa_analysis.core.database_manager import CheckResult, DatabaseManager


@pytest.fixture
def sample_check_result():
    """Create a sample CheckResult for testing."""
    return CheckResult(
        institute_id="INST001",
        file_name="test_file.xlsx",
        file_hash="abcd1234" * 8,
        version_number=1,
        check_name="test_check",
        check_description="Test check description",
        outcome_bool=True,
        outcome_numeric=42.0,
        processed_at=datetime(2023, 1, 1, 12, 0, 0),
    )


@pytest.fixture
def mock_db_manager():
    """Create a mock DatabaseManager for testing."""
    mock = Mock(spec=DatabaseManager)
    mock.execute_query = Mock()
    return mock


@pytest.fixture
def sample_institut_metadata():
    """Create sample institut metadata."""
    return {
        'MitarbeiterNummer': 12345,
        'MitarbeiterKuerzel': 'JD',
        'MitarbeiterName': 'John Doe',
        'MitarbeiterOrgEinheit': 'Unit A',
        'FINMAID': 'INST001',
        'FinmaObjektName': 'Test Institute Ltd.',
        'ZulassungName': 'Banking License',
        'SachbearbeiterTypName': 'Type A'
    }


class TestCheckResult:
    """Test cases for CheckResult dataclass."""

    def test_initialization(self, sample_check_result):
        """Test CheckResult initialization."""
        assert sample_check_result.institute_id == "INST001"
        assert sample_check_result.file_name == "test_file.xlsx"
        assert sample_check_result.check_name == "test_check"
        assert sample_check_result.outcome_bool is True
        assert sample_check_result.outcome_numeric == 42.0

    def test_with_optional_numeric(self):
        """Test CheckResult with optional numeric value."""
        result = CheckResult(
            institute_id="INST002",
            file_name="file.xlsx",
            file_hash="hash123",
            version_number=2,
            check_name="check_name",
            check_description="Description",
            outcome_bool=False,
            outcome_numeric=None,
            processed_at=datetime.now(),
        )
        
        assert result.institute_id == "INST002"
        assert result.version_number == 2
        assert result.outcome_bool is False
        assert result.outcome_numeric is None


class TestDatabaseManagerInstitutMetadata:
    """Test cases for institut metadata functionality in DatabaseManager."""

    @patch('orsa_analysis.core.database_manager.create_engine')
    def test_get_institut_metadata_by_finmaid_success(
        self, mock_create_engine, sample_institut_metadata, tmp_path
    ):
        """Test successful retrieval of institut metadata."""
        # Create a temporary SQL file
        sql_dir = tmp_path / "sql"
        sql_dir.mkdir()
        sql_file = sql_dir / "institut_metadata.sql"
        sql_file.write_text("SELECT * FROM DWHMart.dbo.Sachbearbeiter")
        
        # Mock database response
        mock_df = pd.DataFrame([sample_institut_metadata])
        
        # Create DatabaseManager instance
        db = DatabaseManager()
        db.execute_query = Mock(return_value=mock_df)
        
        # Patch Path to point to our temp SQL file
        with patch.object(Path, 'exists', return_value=True):
            with patch.object(Path, 'read_text', return_value="SELECT * FROM test"):
                result = db.get_institut_metadata_by_finmaid("INST001")
        
        assert result is not None
        assert result['FINMAID'] == 'INST001'
        assert result['FinmaObjektName'] == 'Test Institute Ltd.'
        assert result['MitarbeiterName'] == 'John Doe'

    @patch('orsa_analysis.core.database_manager.create_engine')
    def test_get_institut_metadata_by_finmaid_not_found(self, mock_create_engine):
        """Test when no metadata is found for the given FinmaID."""
        # Mock database response with no matching records
        mock_df = pd.DataFrame()
        
        db = DatabaseManager()
        db.execute_query = Mock(return_value=mock_df)
        
        with patch.object(Path, 'exists', return_value=True):
            with patch.object(Path, 'read_text', return_value="SELECT * FROM test"):
                result = db.get_institut_metadata_by_finmaid("NONEXISTENT")
        
        assert result is None

    @patch('orsa_analysis.core.database_manager.create_engine')
    def test_get_institut_metadata_by_finmaid_sql_file_missing(self, mock_create_engine):
        """Test when SQL file doesn't exist."""
        db = DatabaseManager()
        
        with patch.object(Path, 'exists', return_value=False):
            result = db.get_institut_metadata_by_finmaid("INST001")
        
        assert result is None

    @patch('orsa_analysis.core.database_manager.create_engine')
    def test_get_institut_metadata_by_finmaid_query_error(self, mock_create_engine):
        """Test error handling when query execution fails."""
        db = DatabaseManager()
        db.execute_query = Mock(side_effect=Exception("Database connection error"))
        
        with patch.object(Path, 'exists', return_value=True):
            with patch.object(Path, 'read_text', return_value="SELECT * FROM test"):
                result = db.get_institut_metadata_by_finmaid("INST001")
        
        assert result is None
