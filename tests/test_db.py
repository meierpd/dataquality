"""Unit tests for the database module."""

import pytest
from datetime import datetime

from core.db import CheckResult, DatabaseWriter, InMemoryDatabaseWriter


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

    def test_to_dict(self, sample_check_result):
        """Test converting CheckResult to dictionary."""
        result_dict = sample_check_result.to_dict()

        assert isinstance(result_dict, dict)
        assert result_dict["institute_id"] == "INST001"
        assert result_dict["check_name"] == "test_check"
        assert result_dict["processed_at"] == "2023-01-01T12:00:00"

    def test_from_dict(self):
        """Test creating CheckResult from dictionary."""
        data = {
            "institute_id": "INST002",
            "file_name": "file.xlsx",
            "file_hash": "hash123",
            "version_number": 2,
            "check_name": "check_name",
            "check_description": "Description",
            "outcome_bool": False,
            "outcome_numeric": None,
            "processed_at": "2023-01-01T12:00:00",
        }

        result = CheckResult.from_dict(data)

        assert result.institute_id == "INST002"
        assert result.version_number == 2
        assert result.outcome_bool is False
        assert result.outcome_numeric is None
        assert isinstance(result.processed_at, datetime)

    def test_round_trip_conversion(self, sample_check_result):
        """Test converting to dict and back."""
        result_dict = sample_check_result.to_dict()
        reconstructed = CheckResult.from_dict(result_dict)

        assert reconstructed.institute_id == sample_check_result.institute_id
        assert reconstructed.check_name == sample_check_result.check_name
        assert reconstructed.outcome_bool == sample_check_result.outcome_bool


class TestDatabaseWriter:
    """Test cases for DatabaseWriter base class."""

    def test_initialization(self):
        """Test DatabaseWriter initialization."""
        writer = DatabaseWriter()
        assert writer.results_buffer == []

    def test_add_result(self, sample_check_result):
        """Test adding a result to the buffer."""
        writer = DatabaseWriter()
        writer.add_result(sample_check_result)

        assert len(writer.results_buffer) == 1
        assert writer.results_buffer[0] == sample_check_result

    def test_add_multiple_results(self, sample_check_result):
        """Test adding multiple results."""
        writer = DatabaseWriter()
        writer.add_result(sample_check_result)
        writer.add_result(sample_check_result)

        assert len(writer.results_buffer) == 2

    def test_get_results(self, sample_check_result):
        """Test getting buffered results."""
        writer = DatabaseWriter()
        writer.add_result(sample_check_result)

        results = writer.get_results()
        assert len(results) == 1
        assert results[0] == sample_check_result

    def test_clear_buffer(self, sample_check_result):
        """Test clearing the results buffer."""
        writer = DatabaseWriter()
        writer.add_result(sample_check_result)
        writer.clear_buffer()

        assert len(writer.results_buffer) == 0

    def test_get_existing_versions(self):
        """Test getting existing versions (stub method)."""
        writer = DatabaseWriter()
        versions = writer.get_existing_versions()

        assert isinstance(versions, list)
        assert len(versions) == 0

    def test_write_results(self, sample_check_result):
        """Test writing results (stub method)."""
        writer = DatabaseWriter()
        count = writer.write_results([sample_check_result])

        assert count == 1
        assert len(writer.results_buffer) == 1


class TestInMemoryDatabaseWriter:
    """Test cases for InMemoryDatabaseWriter class."""

    def test_initialization(self):
        """Test InMemoryDatabaseWriter initialization."""
        writer = InMemoryDatabaseWriter()
        assert writer.stored_results == []
        assert writer.results_buffer == []

    def test_write_results(self, sample_check_result):
        """Test writing results to in-memory storage."""
        writer = InMemoryDatabaseWriter()
        count = writer.write_results([sample_check_result])

        assert count == 1
        assert len(writer.stored_results) == 1
        assert writer.stored_results[0] == sample_check_result

    def test_get_existing_versions(self, sample_check_result):
        """Test retrieving existing versions from in-memory storage."""
        writer = InMemoryDatabaseWriter()
        writer.write_results([sample_check_result])

        versions = writer.get_existing_versions()

        assert len(versions) == 1
        assert versions[0]["institute_id"] == "INST001"
        assert versions[0]["file_hash"] == sample_check_result.file_hash
        assert versions[0]["version_number"] == 1

    def test_get_existing_versions_deduplication(self):
        """Test that duplicate versions are not returned."""
        writer = InMemoryDatabaseWriter()

        result1 = CheckResult(
            institute_id="INST001",
            file_name="file.xlsx",
            file_hash="hash1",
            version_number=1,
            check_name="check1",
            check_description="Desc1",
            outcome_bool=True,
            outcome_numeric=None,
            processed_at=datetime.now(),
        )

        result2 = CheckResult(
            institute_id="INST001",
            file_name="file.xlsx",
            file_hash="hash1",
            version_number=1,
            check_name="check2",
            check_description="Desc2",
            outcome_bool=False,
            outcome_numeric=None,
            processed_at=datetime.now(),
        )

        writer.write_results([result1, result2])
        versions = writer.get_existing_versions()

        assert len(versions) == 1

    def test_get_results_for_institute(self):
        """Test retrieving results for a specific institute."""
        writer = InMemoryDatabaseWriter()

        result1 = CheckResult(
            institute_id="INST001",
            file_name="file.xlsx",
            file_hash="hash1",
            version_number=1,
            check_name="check1",
            check_description="Desc1",
            outcome_bool=True,
            outcome_numeric=None,
            processed_at=datetime.now(),
        )

        result2 = CheckResult(
            institute_id="INST002",
            file_name="file.xlsx",
            file_hash="hash2",
            version_number=1,
            check_name="check1",
            check_description="Desc1",
            outcome_bool=True,
            outcome_numeric=None,
            processed_at=datetime.now(),
        )

        writer.write_results([result1, result2])
        results = writer.get_results_for_institute("INST001")

        assert len(results) == 1
        assert results[0].institute_id == "INST001"

    def test_get_results_for_institute_with_version(self):
        """Test retrieving results filtered by version."""
        writer = InMemoryDatabaseWriter()

        result1 = CheckResult(
            institute_id="INST001",
            file_name="file1.xlsx",
            file_hash="hash1",
            version_number=1,
            check_name="check1",
            check_description="Desc1",
            outcome_bool=True,
            outcome_numeric=None,
            processed_at=datetime.now(),
        )

        result2 = CheckResult(
            institute_id="INST001",
            file_name="file2.xlsx",
            file_hash="hash2",
            version_number=2,
            check_name="check1",
            check_description="Desc1",
            outcome_bool=True,
            outcome_numeric=None,
            processed_at=datetime.now(),
        )

        writer.write_results([result1, result2])
        results = writer.get_results_for_institute("INST001", version=1)

        assert len(results) == 1
        assert results[0].version_number == 1

    def test_clear_all(self, sample_check_result):
        """Test clearing all stored data."""
        writer = InMemoryDatabaseWriter()
        writer.write_results([sample_check_result])
        writer.add_result(sample_check_result)

        writer.clear_all()

        assert len(writer.stored_results) == 0
        assert len(writer.results_buffer) == 0
