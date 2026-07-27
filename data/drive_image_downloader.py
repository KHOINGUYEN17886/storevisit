import os
import re
import io
import logging
from typing import Optional
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2.service_account import Credentials

logger = logging.getLogger(__name__)

class DriveImageDownloader:
    def __init__(self, credentials_path: str):
        self.credentials_path = credentials_path
        self.service = None
        
    def _authenticate(self):
        if not self.service:
            if not os.path.exists(self.credentials_path):
                raise FileNotFoundError(f"Google credentials file not found at {self.credentials_path}")
            scopes = ["https://www.googleapis.com/auth/drive.readonly"]
            creds = Credentials.from_service_account_file(self.credentials_path, scopes=scopes)
            self.service = build("drive", "v3", credentials=creds)
            
    def extract_file_id(self, url: str) -> Optional[str]:
        if not url:
            return None
        url = url.strip()
        if "id=" in url:
            match = re.search(r"id=([^&/]+)", url)
            if match:
                return match.group(1)
        elif "/d/" in url:
            match = re.search(r"/d/([^&/]+)", url)
            if match:
                return match.group(1)
        if "/" not in url and "." not in url and len(url) > 10:
            return url
        return None
        
    def download_image(self, drive_url: str, dest_dir: str) -> str:
        file_id = self.extract_file_id(drive_url)
        if not file_id:
            logger.warning(f"Could not extract file ID from URL: {drive_url}")
            return ""
            
        os.makedirs(dest_dir, exist_ok=True)
        
        try:
            self._authenticate()
            file_metadata = self.service.files().get(fileId=file_id, fields="name, mimeType").execute()
            filename = file_metadata.get("name", f"{file_id}.jpg")
            
            # Sanitise filename
            filename = "".join(c for c in filename if c.isalnum() or c in "._- ")
            dest_path = os.path.join(dest_dir, f"{file_id}_{filename}")
            
            if os.path.exists(dest_path) and os.path.getsize(dest_path) > 0:
                logger.info(f"File {dest_path} already exists. Skipping download.")
                return dest_path
                
            logger.info(f"Downloading file ID {file_id} to {dest_path}...")
            
            request = self.service.files().get_media(fileId=file_id)
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                status, done = downloader.next_chunk()
                
            with open(dest_path, "wb") as f:
                f.write(fh.getvalue())
                
            logger.info(f"Successfully downloaded file ID {file_id}")
            return dest_path
            
        except Exception as e:
            logger.error(f"Error downloading Google Drive file ID {file_id}: {e}")
            return ""
