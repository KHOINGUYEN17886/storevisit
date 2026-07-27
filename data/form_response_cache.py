import os
import json
import logging
from typing import List, Dict, Any, Optional
from data.models import StoreFormResponse, FormPhoto, MarketSurveyResponse, SurveyPhoto
from data.google_sheets_reader import GoogleSheetsReader
from data.drive_image_downloader import DriveImageDownloader

logger = logging.getLogger(__name__)

class FormResponseCache:
    def __init__(self, cache_path: str, credentials_path: str, spreadsheet_id: str, photo_cache_dir: str):
        self.cache_path = cache_path
        self.credentials_path = credentials_path
        self.spreadsheet_id = spreadsheet_id
        self.photo_cache_dir = photo_cache_dir
        self.responses: Dict[str, StoreFormResponse] = {}
        self.load()

    def load(self):
        if not os.path.exists(self.cache_path):
            self.responses = {}
            return
        try:
            with open(self.cache_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                self.responses = {
                    k: StoreFormResponse.model_validate(v) for k, v in data.items()
                }
            logger.info(f"Loaded {len(self.responses)} responses from cache.")
        except Exception as e:
            logger.error(f"Error loading form cache: {e}")
            self.responses = {}

    def save(self):
        try:
            os.makedirs(os.path.dirname(self.cache_path), exist_ok=True)
            with open(self.cache_path, "w", encoding="utf-8") as f:
                json.dump(
                    {k: v.model_dump() for k, v in self.responses.items()},
                    f,
                    ensure_ascii=False,
                    indent=2
                )
            logger.info("Saved form cache successfully.")
        except Exception as e:
            logger.error(f"Error saving form cache: {e}")

    def get_all(self) -> List[StoreFormResponse]:
        def get_sort_key(r: StoreFormResponse):
            try:
                return int(r.response_id)
            except ValueError:
                return 0
        return sorted(self.responses.values(), key=get_sort_key, reverse=True)

    def get_by_id(self, response_id: str) -> Optional[StoreFormResponse]:
        return self.responses.get(response_id)

    def sync_from_google(self, sheet_name: str = "Form Responses 1", progress_callback=None) -> int:
        if not self.spreadsheet_id:
            logger.warning("Spreadsheet ID is not set. Sync skipped.")
            return 0
            
        reader = GoogleSheetsReader(self.credentials_path, self.spreadsheet_id)
        downloader = DriveImageDownloader(self.credentials_path)
        
        if progress_callback:
            progress_callback("Connecting to Google Sheets...", 10)
            
        try:
            pending_sheets = reader.get_pending_responses(sheet_name)
        except Exception as e:
            logger.error(f"Failed to fetch pending responses: {e}")
            if progress_callback:
                progress_callback(f"Error: {e}", -1)
            raise e

        if progress_callback:
            progress_callback(f"Found {len(pending_sheets)} pending responses.", 30)

        new_count = 0
        total_photos = sum(len(r.photos) for r in pending_sheets)
        downloaded_photos = 0
        
        for idx, sheet_resp in enumerate(pending_sheets):
            rid = sheet_resp.response_id
            
            if rid not in self.responses:
                new_count += 1
                
            # Download photos for this response
            for photo in sheet_resp.photos:
                if photo.drive_url:
                    if progress_callback:
                        progress_callback(
                            f"Downloading photo {downloaded_photos + 1}/{total_photos} for store {sheet_resp.store_code}...",
                            30 + int(60 * (downloaded_photos / max(total_photos, 1)))
                        )
                    
                    local_p = downloader.download_image(photo.drive_url, self.photo_cache_dir)
                    photo.local_path = local_p
                    downloaded_photos += 1

            self.responses[rid] = sheet_resp
            
        self.save()
        if progress_callback:
            progress_callback(f"Sync complete. Added/updated {new_count} records.", 100)
            
        return new_count

    def mark_done(self, response_id: str, sheet_name: str = "Form Responses 1"):
        if response_id in self.responses:
            self.responses[response_id].status = "done"
            self.save()
            
            if self.spreadsheet_id:
                try:
                    reader = GoogleSheetsReader(self.credentials_path, self.spreadsheet_id)
                    reader.update_row_status(response_id, "done", sheet_name)
                except Exception as e:
                    logger.error(f"Failed to mark row {response_id} as done in Google Sheets: {e}")


class MarketSurveyCache:
    def __init__(self, cache_path: str, credentials_path: str, spreadsheet_id: str, photo_cache_dir: str):
        self.cache_path = cache_path
        self.credentials_path = credentials_path
        self.spreadsheet_id = spreadsheet_id
        self.photo_cache_dir = photo_cache_dir
        self.responses: Dict[str, MarketSurveyResponse] = {}
        self.load()

    def load(self):
        if not os.path.exists(self.cache_path):
            self.responses = {}
            return
        try:
            with open(self.cache_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                self.responses = {
                    k: MarketSurveyResponse.model_validate(v) for k, v in data.items()
                }
            logger.info(f"Loaded {len(self.responses)} survey responses from cache.")
        except Exception as e:
            logger.error(f"Error loading survey cache: {e}")
            self.responses = {}

    def save(self):
        try:
            os.makedirs(os.path.dirname(self.cache_path), exist_ok=True)
            with open(self.cache_path, "w", encoding="utf-8") as f:
                json.dump(
                    {k: v.model_dump() for k, v in self.responses.items()},
                    f,
                    ensure_ascii=False,
                    indent=2
                )
            logger.info("Saved survey cache successfully.")
        except Exception as e:
            logger.error(f"Error saving survey cache: {e}")

    def get_all(self) -> List[MarketSurveyResponse]:
        def get_sort_key(r: MarketSurveyResponse):
            try:
                return int(r.response_id)
            except ValueError:
                return 0
        return sorted(self.responses.values(), key=get_sort_key, reverse=True)

    def get_by_id(self, response_id: str) -> Optional[MarketSurveyResponse]:
        return self.responses.get(response_id)

    def sync_from_google(self, sheet_name: str = "MarketSurvey_Responses", progress_callback=None) -> int:
        if not self.spreadsheet_id:
            logger.warning("Spreadsheet ID is not set. Sync skipped.")
            return 0
            
        reader = GoogleSheetsReader(self.credentials_path, self.spreadsheet_id)
        downloader = DriveImageDownloader(self.credentials_path)
        
        if progress_callback:
            progress_callback("Connecting to Google Sheets (Survey)...", 10)
            
        try:
            pending_sheets = reader.get_pending_survey_responses(sheet_name)
        except Exception as e:
            logger.error(f"Failed to fetch pending survey responses: {e}")
            if progress_callback:
                progress_callback(f"Error: {e}", -1)
            raise e

        if progress_callback:
            progress_callback(f"Found {len(pending_sheets)} pending survey responses.", 30)

        new_count = 0
        total_photos = sum(len(r.photos) for r in pending_sheets)
        downloaded_photos = 0
        
        for idx, sheet_resp in enumerate(pending_sheets):
            rid = sheet_resp.response_id
            
            if rid in self.responses and self.responses[rid].status == "done":
                continue
                
            if rid not in self.responses:
                new_count += 1
                
            # Download photos for this survey
            for photo in sheet_resp.photos:
                if photo.drive_url:
                    if progress_callback:
                        progress_callback(
                            f"Downloading photo {downloaded_photos + 1}/{total_photos} for store {sheet_resp.store_code}...",
                            30 + int(60 * (downloaded_photos / max(total_photos, 1)))
                        )
                    
                    local_p = downloader.download_image(photo.drive_url, self.photo_cache_dir)
                    photo.local_path = local_p
                    downloaded_photos += 1

            self.responses[rid] = sheet_resp
            
        self.save()
        if progress_callback:
            progress_callback(f"Sync complete. Added/updated {new_count} survey records.", 100)
            
        return new_count
