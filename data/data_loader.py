import os
import yaml
import pandas as pd
import json
from typing import Dict, Any

class DataLoader:
    def __init__(self, config_path: str = "config/app_config.yaml"):
        # Resolve config path relative to the StoreVisit root directory
        self.root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        abs_config_path = os.path.join(self.root_dir, config_path)
        
        with open(abs_config_path, "r", encoding="utf-8") as f:
            self.config = yaml.safe_load(f)
            
        # Load field mapping
        abs_mapping_path = os.path.join(self.root_dir, "config/field_mapping.yaml")
        with open(abs_mapping_path, "r", encoding="utf-8") as f:
            self.field_mapping = yaml.safe_load(f)
            
        self.paths = self.config["paths"]
        # Convert relative paths to absolute if necessary
        for key, path in self.paths.items():
            if not os.path.isabs(path):
                self.paths[key] = os.path.join(self.root_dir, path)
                
        self._cache: Dict[str, Any] = {}

    def get_path(self, name: str) -> str:
        return self.paths.get(name, "")

    def load_dim_store(self) -> pd.DataFrame:
        if "dim_store" not in self._cache:
            path = self.get_path("dim_store")
            self._cache["dim_store"] = pd.read_excel(path)
        return self._cache["dim_store"]

    def load_revenue(self) -> pd.DataFrame:
        if "revenue" not in self._cache:
            path = self.get_path("revenue")
            # Parse Date column
            self._cache["revenue"] = pd.read_csv(path, parse_dates=["Date"])
        return self._cache["revenue"]

    def load_target(self) -> pd.DataFrame:
        if "target" not in self._cache:
            path = self.get_path("target")
            self._cache["target"] = pd.read_csv(path)
        return self._cache["target"]

    def load_health_master(self) -> pd.DataFrame:
        if "health_master" not in self._cache:
            path = self.get_path("health_master")
            # Read all columns as string to avoid type inference warnings on SKU or names
            self._cache["health_master"] = pd.read_csv(path, dtype=str)
        return self._cache["health_master"]

    def load_stock(self) -> pd.DataFrame:
        if "stock" not in self._cache:
            path = self.get_path("stock")
            # Read required columns for quantities, SKU counts, and distribution periods
            self._cache["stock"] = pd.read_csv(
                path, 
                usecols=["StoreCode", "Quantity", "ProductPeriod", "ProductCode", "Năm phân phối", "Tháng phân phối"]
            )
        return self._cache["stock"]

    def load_staff_list(self) -> pd.DataFrame:
        if "staff_list" not in self._cache:
            path = self.get_path("staff_list")
            # Staff list has 1 header row to skip as seen in generate_full_report.py
            df = pd.read_excel(path, skiprows=1)
            
            # Dynamically calculate "Thâm niên năm" if missing but "Ngày vào làm" is present
            if "Thâm niên năm" not in df.columns and "Ngày vào làm" in df.columns:
                try:
                    hire_date = pd.to_datetime(df["Ngày vào làm"], unit='D', origin='1899-12-30', errors='coerce')
                    today = pd.Timestamp.now()
                    df["Thâm niên năm"] = (today - hire_date).dt.days / 365.25
                except Exception as e:
                    df["Thâm niên năm"] = 0.0
                    
            self._cache["staff_list"] = df
        return self._cache["staff_list"]

    def load_weekly_json(self) -> dict:
        if "weekly_json" not in self._cache:
            path = self.get_path("weekly_json")
            if os.path.exists(path):
                with open(path, "r", encoding="utf-8") as f:
                    self._cache["weekly_json"] = json.load(f)
            else:
                self._cache["weekly_json"] = {}
        return self._cache["weekly_json"]

    def get_asm_email(self, asm_name: str) -> str:
        """Lookup email for given ASM name from dim_store or config fallback."""
        if not asm_name:
            return ""
            
        import re
        def remove_accents(str_val: str) -> str:
            if not str_val: return ""
            s = str(str_val).strip()
            s = re.sub(r'[àáạảãâầấậẩẫăằắặẳẵ]', 'a', s)
            s = re.sub(r'[ÀÁẠẢÃÂẦẤẬẨẪĂẰẮẶẲẴ]', 'A', s)
            s = re.sub(r'[èéẹẻẽêềếệểễ]', 'e', s)
            s = re.sub(r'[ÈÉẸẺẼÊỀẾỆỂỄ]', 'E', s)
            s = re.sub(r'[ìíịỉĩ]', 'i', s)
            s = re.sub(r'[ÌÍỊỈĨ]', 'I', s)
            s = re.sub(r'[òóọỏõôồốộổỗơờớợởỡ]', 'o', s)
            s = re.sub(r'[ÒÓỌỎÕÔỒỐỘỔỖƠỜỚỢỞỠ]', 'O', s)
            s = re.sub(r'[ùúụủũưừứựửữ]', 'u', s)
            s = re.sub(r'[ÙÚỤỦŨƯỪỨỰỬỮ]', 'U', s)
            s = re.sub(r'[ỳýỵỷỹ]', 'y', s)
            s = re.sub(r'[ỲÝỴỶỸ]', 'Y', s)
            s = re.sub(r'[đ]', 'd', s)
            s = re.sub(r'[Đ]', 'D', s)
            s = re.sub(r'[^a-zA-Z0-9]', '', s)
            return s.lower()

        norm_input = remove_accents(asm_name)
        df_store = self.load_dim_store()
        
        if not df_store.empty:
            for _, r in df_store.iterrows():
                raw_asm = str(r.get("ASM", "")).strip()
                email = str(r.get("Email", r.get("ASM_Email", ""))).strip()
                if raw_asm and email and "@" in email:
                    if remove_accents(raw_asm) == norm_input:
                        return email

        # Config fallback mapping if available
        asm_emails_config = self.config.get("asm_emails", {})
        for k, v in asm_emails_config.items():
            if remove_accents(k) == norm_input:
                return v

        return f"{norm_input}@anphuoc.com.vn"

    def clear_cache(self):
        self._cache.clear()
