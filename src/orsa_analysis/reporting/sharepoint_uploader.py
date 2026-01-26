"""SharePoint uploader for uploading generated reports back to SharePoint."""

import logging
import os
from pathlib import Path
from urllib.parse import urlparse

import requests
from requests_ntlm import HttpNtlmAuth

logger = logging.getLogger(__name__)


class SharePointUploader:
    """Uploads report files to SharePoint folders.

    The upload folder is automatically resolved by following the redirect
    embedded in a download link. Files with the same name are not overwritten.
    
    Note:
        Credentials are retrieved from environment variables DB_USER and DB_PASSWORD,
        which are set by DatabaseManager.
    """

    def __init__(self, ca_cert_path: Path = None):
        """Initialize the SharePoint uploader.
        
        Args:
            ca_cert_path: Optional path to CA certificate file.
                         If None, looks for SwisscomRootCore.crt in project root.
        """
        # Get credentials from environment (set by DatabaseManager)
        self.user = os.getenv("DB_USER")
        self.password = os.getenv("DB_PASSWORD")
        
        if not self.user or not self.password:
            logger.warning("DB_USER or DB_PASSWORD not found in environment. Upload may fail.")
        
        self.auth = HttpNtlmAuth(self.user, self.password)
        
        # Determine CA certificate path
        if ca_cert_path is None:
            # Look for certificate in project root
            project_root = Path(__file__).resolve().parent.parent.parent.parent
            ca_cert_path = project_root / "SwisscomRootCore.crt"
        
        self.ca_cert = ca_cert_path
        
        if not self.ca_cert.exists():
            logger.warning(f"CA certificate not found at {self.ca_cert}. Using default SSL verification.")
            self.ca_cert = None

    def resolve_folder_from_link(self, download_link: str) -> str:
        """Resolve the WebDAV folder URL from a download link.

        Args:
            download_link: Redirect-style SharePoint download link.

        Returns:
            Real WebDAV folder URL.
            
        Raises:
            Exception: If folder resolution fails.
        """
        try:
            logger.debug(f"Resolving folder from link: {download_link}")
            
            verify = str(self.ca_cert) if self.ca_cert else True
            
            r = requests.get(
                download_link,
                auth=self.auth,
                verify=verify,
                allow_redirects=True
            )
            r.raise_for_status()
            
            final_url = r.url
            parsed = urlparse(final_url)
            parts = parsed.path.split("/")
            folder = "/".join(parts[:-1])
            folder_url = f"{parsed.scheme}://{parsed.netloc}{folder}"
            
            logger.debug(f"Resolved folder URL: {folder_url}")
            return folder_url
            
        except Exception as e:
            logger.error(f"Failed to resolve folder from link: {e}")
            raise

    def file_exists(self, folder_url: str, filename: str) -> bool:
        """Check if a file already exists in the SharePoint folder.
        
        Args:
            folder_url: WebDAV folder URL.
            filename: Name of the file to check.
            
        Returns:
            True if file exists, False otherwise.
        """
        try:
            target_url = f"{folder_url}/{filename}"
            
            verify = str(self.ca_cert) if self.ca_cert else True
            
            r = requests.head(
                target_url,
                auth=self.auth,
                verify=verify
            )
            
            # HEAD request returns 200 if file exists
            return r.status_code == 200
            
        except Exception as e:
            logger.debug(f"Error checking file existence: {e}")
            return False

    def upload(self, download_link: str, filepath: str, skip_if_exists: bool = True) -> dict:
        """Upload a file to SharePoint folder inferred from download link.

        Args:
            download_link: Redirect-style download link of any file in target folder.
            filepath: Local file path to upload.
            skip_if_exists: If True, skip upload if file already exists (default: True).

        Returns:
            Dictionary with status information:
                - success: Boolean indicating if upload was successful
                - message: Status message
                - skipped: Boolean indicating if upload was skipped
        """
        try:
            # Resolve folder URL
            folder_url = self.resolve_folder_from_link(download_link)
            
            filename = Path(filepath).name
            target_url = f"{folder_url}/{filename}"
            
            # Check if file already exists
            if skip_if_exists and self.file_exists(folder_url, filename):
                logger.info(f"File already exists on SharePoint, skipping: {filename}")
                return {
                    "success": True,
                    "message": "File already exists, skipped upload.",
                    "skipped": True
                }
            
            # Read file content
            content = Path(filepath).read_bytes()
            
            verify = str(self.ca_cert) if self.ca_cert else True
            
            # Upload file
            logger.info(f"Uploading {filename} to SharePoint...")
            r = requests.put(
                target_url,
                data=content,
                auth=self.auth,
                verify=verify
            )
            
            # Handle response
            if r.status_code == 201:
                logger.info(f"✓ File created: {filename}")
                return {"success": True, "message": "File created.", "skipped": False}
            elif r.status_code == 200 or r.status_code == 204:
                logger.info(f"✓ File updated: {filename}")
                return {"success": True, "message": "File updated.", "skipped": False}
            elif r.status_code == 401:
                logger.error(f"✗ Unauthorized: {filename}")
                return {"success": False, "message": "Unauthorized.", "skipped": False}
            elif r.status_code == 403:
                logger.error(f"✗ Forbidden: {filename}")
                return {"success": False, "message": "Forbidden.", "skipped": False}
            elif r.status_code == 404:
                logger.error(f"✗ Folder not found: {filename}")
                return {"success": False, "message": "Folder not found.", "skipped": False}
            else:
                logger.error(f"✗ Unexpected status {r.status_code}: {r.text}")
                return {
                    "success": False,
                    "message": f"Unexpected status {r.status_code}: {r.text}",
                    "skipped": False
                }
                
        except Exception as e:
            logger.error(f"Failed to upload file: {e}")
            return {"success": False, "message": f"Upload failed: {e}", "skipped": False}
