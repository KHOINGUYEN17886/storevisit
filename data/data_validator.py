import os
import pandas as pd
from typing import List, Dict

class DataValidationError(Exception):
    pass

class DataValidator:
    def __init__(self, loader):
        self.loader = loader

    def validate_files_exist(self) -> List[str]:
        """Verify that all required files exist on disk."""
        missing = []
        for name, path in self.loader.paths.items():
            # temp and output dirs don't need to exist beforehand (will be created)
            if name in ["temp_dir", "output_dir", "templates_dir", "weekly_json"]:
                continue
            if not os.path.exists(path):
                missing.append(f"{name} (expected at: {path})")
        if missing:
            raise DataValidationError(f"Missing required source file(s):\n" + "\n".join(missing))
        input_files = []
        for name, path in self.loader.paths.items():
            if name not in ["temp_dir", "output_dir", "templates_dir"]:
                input_files.append(path)
        return input_files

    def validate_schemas(self):
        """Validate essential column names in loaded files."""
        # 1. DimStore validation
        df_dim = self.loader.load_dim_store()
        required_dim = ["StoreCode", "StoreName", "Address", "Region", "ASM"]
        self._check_cols("DimStore", df_dim, required_dim)

        # 2. Revenue validation
        df_rev = self.loader.load_revenue()
        required_rev = ["StoreCode", "Date", "SalesType", "Revenue"]
        self._check_cols("Revenue", df_rev, required_rev)

        # 3. Target validation
        df_tgt = self.loader.load_target()
        required_tgt = ["StoreCode", "Year", "MonthNo", "BrandGroup", "Target"]
        self._check_cols("Target", df_tgt, required_tgt)

        # 4. Stock validation
        df_stock = self.loader.load_stock()
        required_stock = ["StoreCode", "Quantity", "ProductPeriod"]
        self._check_cols("Stock", df_stock, required_stock)

        # 5. Health Master validation
        df_hm = self.loader.load_health_master()
        required_hm = ["StoreCode", "SKU", "ProductName", "Brand", "Stock_Qty", "Weekly_Sales_Store", "Sales_4W", "Stock_Age_Days"]
        self._check_cols("Health Master", df_hm, required_hm)

        # 6. Staff List validation
        df_staff = self.loader.load_staff_list()
        required_staff = ["Nơi làm việc", "Tên nhân viên", "Chức danh", "Thâm niên năm"]
        self._check_cols("Staff List", df_staff, required_staff)

    def _check_cols(self, label: str, df: pd.DataFrame, expected_cols: List[str]):
        missing = [col for col in expected_cols if col not in df.columns]
        if missing:
            raise DataValidationError(
                f"Schema validation failed for {label}: Missing required column(s): {', '.join(missing)}\n"
                f"Available columns: {', '.join(df.columns[:15])}"
            )
