"""File versioning and hashing module."""

import hashlib
from pathlib import Path
from typing import Dict, Optional
from dataclasses import dataclass
import logging

logger = logging.getLogger(__name__)


@dataclass
class FileVersion:
    """Represents version metadata for a file."""

    institute_id: str
    file_name: str
    file_hash: str
    version_number: int


class VersionManager:
    """Manages file versioning based on hash computation."""

    def __init__(self):
        """Initialize the version manager with empty state."""
        self._version_cache: Dict[str, Dict[str, int]] = {}

    def compute_file_hash(self, file_path: Path) -> str:
        """Compute SHA-256 hash of a file.

        Args:
            file_path: Path to the file

        Returns:
            Hexadecimal string representation of the file hash

        Raises:
            FileNotFoundError: If the file does not exist
        """
        if not file_path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")

        sha256_hash = hashlib.sha256()
        try:
            with open(file_path, "rb") as f:
                for byte_block in iter(lambda: f.read(4096), b""):
                    sha256_hash.update(byte_block)
            file_hash = sha256_hash.hexdigest()
            logger.debug(f"Computed hash for {file_path.name}: {file_hash}")
            return file_hash
        except Exception as e:
            logger.error(f"Failed to compute hash for {file_path}: {e}")
            raise

    def load_existing_versions(
        self, existing_data: list[Dict[str, any]]
    ) -> None:
        """Load existing version information from database.

        Args:
            existing_data: List of dicts containing institute_id, file_hash, and version_number
        """
        self._version_cache.clear()
        for record in existing_data:
            institute_id = record["institute_id"]
            file_hash = record["file_hash"]
            version = record["version_number"]

            if institute_id not in self._version_cache:
                self._version_cache[institute_id] = {}

            self._version_cache[institute_id][file_hash] = version

        logger.info(f"Loaded {len(existing_data)} existing version records")

    def get_version(
        self, institute_id: str, file_path: Path
    ) -> FileVersion:
        """Get version information for a file.

        If the file hash is new for this institute, assigns a new version number.

        Args:
            institute_id: Identifier for the institute
            file_path: Path to the file

        Returns:
            FileVersion object with version metadata
        """
        file_hash = self.compute_file_hash(file_path)
        file_name = file_path.name

        if institute_id not in self._version_cache:
            self._version_cache[institute_id] = {}

        if file_hash in self._version_cache[institute_id]:
            version_number = self._version_cache[institute_id][file_hash]
            logger.info(
                f"Found existing version {version_number} for {institute_id}/{file_name}"
            )
        else:
            current_versions = list(self._version_cache[institute_id].values())
            version_number = max(current_versions) + 1 if current_versions else 1
            self._version_cache[institute_id][file_hash] = version_number
            logger.info(
                f"Assigned new version {version_number} for {institute_id}/{file_name}"
            )

        return FileVersion(
            institute_id=institute_id,
            file_name=file_name,
            file_hash=file_hash,
            version_number=version_number,
        )

    def is_processed(self, institute_id: str, file_hash: str) -> bool:
        """Check if a file with given hash has already been processed.

        Args:
            institute_id: Identifier for the institute
            file_hash: Hash of the file

        Returns:
            True if the file has been processed, False otherwise
        """
        return (
            institute_id in self._version_cache
            and file_hash in self._version_cache[institute_id]
        )

    def get_latest_version(self, institute_id: str) -> Optional[int]:
        """Get the latest version number for an institute.

        Args:
            institute_id: Identifier for the institute

        Returns:
            Latest version number or None if no versions exist
        """
        if institute_id not in self._version_cache or not self._version_cache[institute_id]:
            return None
        return max(self._version_cache[institute_id].values())
