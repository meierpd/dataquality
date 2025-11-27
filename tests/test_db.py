"""Unit tests for the database module."""

import pytest
from datetime import datetime

from orsa_analysis.core.database_manager import CheckResult


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
