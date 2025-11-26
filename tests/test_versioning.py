"""Unit tests for the versioning module."""

import pytest
from pathlib import Path

from core.versioning import VersionManager, FileVersion


@pytest.fixture
def sample_file(tmp_path):
    """Create a sample file for testing."""
    file_path = tmp_path / "test_file.txt"
    file_path.write_text("Sample content for hashing")
    return file_path


@pytest.fixture
def version_manager():
    """Create a VersionManager instance."""
    return VersionManager()


class TestVersionManager:
    """Test cases for VersionManager class."""

    def test_initialization(self, version_manager):
        """Test VersionManager initialization."""
        assert version_manager._version_cache == {}

    def test_compute_file_hash(self, version_manager, sample_file):
        """Test computing file hash."""
        file_hash = version_manager.compute_file_hash(sample_file)

        assert isinstance(file_hash, str)
        assert len(file_hash) == 64

    def test_compute_file_hash_consistent(self, version_manager, sample_file):
        """Test that file hash is consistent across multiple calls."""
        hash1 = version_manager.compute_file_hash(sample_file)
        hash2 = version_manager.compute_file_hash(sample_file)

        assert hash1 == hash2

    def test_compute_file_hash_different_content(self, version_manager, tmp_path):
        """Test that different files produce different hashes."""
        file1 = tmp_path / "file1.txt"
        file2 = tmp_path / "file2.txt"
        file1.write_text("Content 1")
        file2.write_text("Content 2")

        hash1 = version_manager.compute_file_hash(file1)
        hash2 = version_manager.compute_file_hash(file2)

        assert hash1 != hash2

    def test_compute_file_hash_not_found(self, version_manager):
        """Test computing hash of non-existent file."""
        with pytest.raises(FileNotFoundError):
            version_manager.compute_file_hash(Path("/nonexistent/file.txt"))

    def test_get_version_new_file(self, version_manager, sample_file):
        """Test getting version for a new file."""
        version_info = version_manager.get_version("INST001", sample_file)

        assert isinstance(version_info, FileVersion)
        assert version_info.institute_id == "INST001"
        assert version_info.file_name == sample_file.name
        assert version_info.version_number == 1
        assert len(version_info.file_hash) == 64

    def test_get_version_existing_hash(self, version_manager, sample_file):
        """Test getting version for an existing file hash."""
        version1 = version_manager.get_version("INST001", sample_file)
        version2 = version_manager.get_version("INST001", sample_file)

        assert version1.version_number == version2.version_number
        assert version1.file_hash == version2.file_hash

    def test_get_version_new_hash(self, version_manager, tmp_path):
        """Test getting version for a new file hash."""
        file1 = tmp_path / "file1.txt"
        file2 = tmp_path / "file2.txt"
        file1.write_text("Content 1")
        file2.write_text("Content 2")

        version1 = version_manager.get_version("INST001", file1)
        version2 = version_manager.get_version("INST001", file2)

        assert version1.version_number == 1
        assert version2.version_number == 2

    def test_get_version_different_institutes(self, version_manager, sample_file):
        """Test that versions are tracked separately per institute."""
        version1 = version_manager.get_version("INST001", sample_file)
        version2 = version_manager.get_version("INST002", sample_file)

        assert version1.version_number == 1
        assert version2.version_number == 1

    def test_load_existing_versions(self, version_manager):
        """Test loading existing version data."""
        existing_data = [
            {"institute_id": "INST001", "file_hash": "hash1", "version_number": 1},
            {"institute_id": "INST001", "file_hash": "hash2", "version_number": 2},
            {"institute_id": "INST002", "file_hash": "hash3", "version_number": 1},
        ]

        version_manager.load_existing_versions(existing_data)

        assert "INST001" in version_manager._version_cache
        assert "INST002" in version_manager._version_cache
        assert version_manager._version_cache["INST001"]["hash1"] == 1
        assert version_manager._version_cache["INST001"]["hash2"] == 2

    def test_is_processed(self, version_manager):
        """Test checking if a file has been processed."""
        existing_data = [
            {"institute_id": "INST001", "file_hash": "hash1", "version_number": 1},
        ]
        version_manager.load_existing_versions(existing_data)

        assert version_manager.is_processed("INST001", "hash1") is True
        assert version_manager.is_processed("INST001", "hash2") is False
        assert version_manager.is_processed("INST002", "hash1") is False

    def test_get_latest_version(self, version_manager):
        """Test getting the latest version number for an institute."""
        existing_data = [
            {"institute_id": "INST001", "file_hash": "hash1", "version_number": 1},
            {"institute_id": "INST001", "file_hash": "hash2", "version_number": 3},
            {"institute_id": "INST001", "file_hash": "hash3", "version_number": 2},
        ]
        version_manager.load_existing_versions(existing_data)

        latest = version_manager.get_latest_version("INST001")
        assert latest == 3

    def test_get_latest_version_no_data(self, version_manager):
        """Test getting latest version when no data exists."""
        latest = version_manager.get_latest_version("INST999")
        assert latest is None

    def test_version_increments_correctly(self, version_manager, tmp_path):
        """Test that versions increment correctly for new files."""
        file1 = tmp_path / "file1.txt"
        file2 = tmp_path / "file2.txt"
        file3 = tmp_path / "file3.txt"
        file1.write_text("Content 1")
        file2.write_text("Content 2")
        file3.write_text("Content 3")

        v1 = version_manager.get_version("INST001", file1)
        v2 = version_manager.get_version("INST001", file2)
        v3 = version_manager.get_version("INST001", file3)

        assert v1.version_number == 1
        assert v2.version_number == 2
        assert v3.version_number == 3
