import sys
import json
import os
from datetime import datetime

class IPCMessageSender:
    def __init__(self, job_id: str, log_dir: str = "logs"):
        self.job_id = job_id
        self.seq = 0
        self.root_dir = os.path.dirname(os.path.abspath(__file__))
        self.log_path = os.path.join(self.root_dir, log_dir, f"job_{job_id}.log")
        os.makedirs(os.path.dirname(self.log_path), exist_ok=True)

    def _send(self, msg_type: str, data: dict = None):
        """Build message, write to stdout as a single JSON line, flush, and write to log file."""
        self.seq += 1
        message = {
            "job_id": self.job_id,
            "seq": self.seq,
            "timestamp": datetime.now().isoformat(),
            "type": msg_type,
            "payload": data or {}
        }
        
        # Write to stdout strictly as single line
        json_line = json.dumps(message, ensure_ascii=False)
        sys.stdout.write(json_line + "\n")
        sys.stdout.flush()

        # Log locally for diagnostics
        try:
            with open(self.log_path, "a", encoding="utf-8") as f:
                f.write(json_line + "\n")
        except Exception:
            pass

    def send_job_started(self, params: dict):
        self._send("job_started", params)

    def send_stage_started(self, stage_name: str, total_steps: int = 100):
        self._send("stage_started", {"stage": stage_name, "total_steps": total_steps})

    def send_progress(self, current_step: int, message: str):
        self._send("progress", {"current_step": current_step, "message": message})

    def send_warning(self, warn_msg: str):
        self._send("warning", {"warning": warn_msg})

    def send_stage_completed(self, stage_name: str):
        self._send("stage_completed", {"stage": stage_name})

    def send_job_cancelled(self, reason: str = "User requested cancellation"):
        self._send("job_cancelled", {"reason": reason})

    def send_job_failed(self, error_msg: str, traceback_str: str = ""):
        self._send("job_failed", {"error": error_msg, "traceback": traceback_str})

    def send_job_completed(self, pptx_path: str, pdf_path: str, manifest_path: str):
        self._send("job_completed", {
            "pptx": pptx_path,
            "pdf": pdf_path,
            "manifest": manifest_path
        })
