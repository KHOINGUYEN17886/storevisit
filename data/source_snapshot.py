import os
import hashlib
import json
from typing import List, Dict

class SourceFileSnapshot:
    def __init__(self, path: str):
        self.path = path
        self.size = 0
        self.modified_at = ""
        self.sha256 = ""
        if os.path.exists(path):
            self.capture()

    def capture(self):
        stat = os.stat(self.path)
        self.size = stat.st_size
        # Get ISO formatted modification time
        from datetime import datetime, timezone
        mtime = datetime.fromtimestamp(stat.st_mtime, tz=timezone.utc)
        self.modified_at = mtime.isoformat()
        
        # Calculate SHA256 in chunks to avoid memory issues with large files
        sha256_hash = hashlib.sha256()
        with open(self.path, "rb") as f:
            for byte_block in iter(lambda: f.read(4096), b""):
                sha256_hash.update(byte_block)
        self.sha256 = sha256_hash.hexdigest()

    def to_dict(self) -> dict:
        return {
            "path": self.path,
            "size": self.size,
            "modified_at": self.modified_at,
            "sha256": self.sha256
        }

class JobDataSnapshot:
    def __init__(self, paths: List[str]):
        self.paths = paths
        self.initial_snapshots: Dict[str, SourceFileSnapshot] = {}
        
    def record_before(self):
        """Record file states before reading."""
        for path in self.paths:
            if os.path.exists(path):
                self.initial_snapshots[path] = SourceFileSnapshot(path)

    def verify_after_and_save(self, output_json_path: str):
        """Record file states after reading, verify no changes, and write to JSON."""
        current_snapshots = {}
        for path in self.paths:
            if os.path.exists(path):
                current_snapshots[path] = SourceFileSnapshot(path)
                
        # Compare before and after
        for path, initial in self.initial_snapshots.items():
            current = current_snapshots.get(path)
            if not current:
                raise RuntimeError(f"Source file was deleted during execution: {path}")
            if initial.sha256 != current.sha256 or initial.size != current.size:
                raise RuntimeError(
                    f"Data concurrency violation: File {path} was modified during job execution!\n"
                    f"Before: SHA256={initial.sha256}, Size={initial.size}\n"
                    f"After: SHA256={current.sha256}, Size={current.size}"
                )
                
        # If valid, serialize snapshot to job temp folder
        snapshot_data = {
            "source_files": [snap.to_dict() for snap in current_snapshots.values()]
        }
        os.makedirs(os.path.dirname(output_json_path), exist_ok=True)
        with open(output_json_path, "w", encoding="utf-8") as f:
            json.dump(snapshot_data, f, indent=2, ensure_ascii=False)
