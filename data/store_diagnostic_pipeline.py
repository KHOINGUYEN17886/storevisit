import os
import calendar
import datetime
import json
import logging
from typing import Dict, Any, Optional, List
import pandas as pd
import numpy as np

logger = logging.getLogger("StoreDiagnosticPipeline")
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(name)s: %(message)s")


class StoreDiagnosticPipeline:
    """
    Enterprise Data Pipeline for Retail Commander Store Diagnostic Cards.
    Reads ground-truth datasets from Data Lake and produces a versioned diagnostic snapshot.
    Compliant with Gate 1.5 Semantic Audit & Strict Data Contract.
    """

    def __init__(self, config_path: str = "config/app_config.yaml"):
        self.root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        import yaml
        abs_config = os.path.join(self.root_dir, config_path)
        with open(abs_config, "r", encoding="utf-8") as f:
            self.config = yaml.safe_load(f)
            
        self.paths = self.config.get("paths", {})
        for k, v in self.paths.items():
            if not os.path.isabs(v):
                self.paths[k] = os.path.join(self.root_dir, v)

    def load_ssot_stores(self) -> Dict[str, Dict[str, Any]]:
        """Load 185 SSOT stores from StoresInfo.xlsx."""
        stores_path = r"C:\All_Report\1_Mapping\StoresInfo.xlsx"
        if not os.path.exists(stores_path):
            stores_path = self.paths.get("dim_store", "")

        df_stores = pd.read_excel(stores_path, sheet_name="Danh bạ CH" if "StoresInfo" in stores_path else 0)
        store_meta = {}
        for _, r in df_stores.iterrows():
            code = str(r.get("CODE", r.get("StoreCode", ""))).strip().upper()
            if not code or code == "NAN":
                continue
            name = str(r.get("TÊN CỬA HÀNG - CHUẨN", r.get("StoreName", code))).strip()
            region = str(r.get("VÙNG", r.get("Region", "HCM"))).strip()
            asm = str(r.get("ASM", "Khác")).strip()
            manager = str(r.get("CỬA HÀNG TRƯỞNG ", r.get("StoreManager", "Chưa cập nhật"))).strip()
            store_type = str(r.get("STORETYPE", r.get("StoreType", "Store"))).strip()
            
            try:
                headcount_target = int(float(r.get("ĐỊNH BIÊN", r.get("TargetHeadcount", 0))))
            except (ValueError, TypeError):
                headcount_target = 0
                
            store_meta[code] = {
                "code": code,
                "name": name,
                "region": region,
                "asm": asm,
                "manager": manager,
                "store_type": store_type,
                "headcount_target": headcount_target
            }
        return store_meta

    def generate_snapshot(self) -> Dict[str, Any]:
        """Generate 185 Store Diagnostic Snapshot with versioning, 4-tier severity and priority blockers."""
        store_meta = self.load_ssot_stores()
        store_codes = list(store_meta.keys())
        
        # 1. Revenue
        rev_path = self.paths.get("revenue", r"C:\All_Report\2_CleanData\Revenue\Fact_Revenue_AllYears_Standardized.csv")
        df_rev = pd.read_csv(rev_path, parse_dates=["Date"])
        latest_rev_date = df_rev["Date"].max()
        curr_year = latest_rev_date.year
        curr_month = latest_rev_date.month
        curr_day = latest_rev_date.day
        _, total_days_in_month = calendar.monthrange(curr_year, curr_month)
        
        # Date & Selling Days Semantics (Gate 1.5 Audit)
        selling_days_elapsed = curr_day # Last closed trading day
        selling_days_remaining = total_days_in_month - selling_days_elapsed # Days 29, 30, 31 = 3 days
        expected_progress_pct = round((selling_days_elapsed / total_days_in_month) * 100, 2) # 90.32%

        df_rev_curr = df_rev[(df_rev["Date"].dt.year == curr_year) & (df_rev["Date"].dt.month == curr_month)]
        rev_mtd_by_store = df_rev_curr.groupby(df_rev_curr["StoreCode"].str.strip().str.upper())["Revenue"].sum().to_dict()

        # 2. Target (Accurate Store_Total Filtering - avoiding double count)
        tgt_path = self.paths.get("target", r"C:\All_Report\2_CleanData\Target\TargetMonthly.csv")
        df_tgt = pd.read_csv(tgt_path)
        df_tgt_curr = df_tgt[(df_tgt["Year"] == curr_year) & (df_tgt["MonthNo"] == curr_month)]
        
        # Primary: BrandGroup == 'Store' (Store_Total). Fallback: Sum of sub-brands (APPI + AB)
        store_total_tgt = df_tgt_curr[df_tgt_curr["BrandGroup"] == "Store"].groupby(df_tgt_curr["StoreCode"].astype(str).str.strip().str.upper())["Target"].sum().to_dict()
        sub_brands_tgt = df_tgt_curr[df_tgt_curr["BrandGroup"].isin(["APPI", "AB"])].groupby(df_tgt_curr["StoreCode"].astype(str).str.strip().str.upper())["Target"].sum().to_dict()

        tgt_mtd_by_store = {}
        for scode, s_tgt in store_total_tgt.items():
            tgt_mtd_by_store[scode] = s_tgt
        for scode, sub_tgt in sub_brands_tgt.items():
            if scode not in tgt_mtd_by_store or tgt_mtd_by_store[scode] <= 0:
                tgt_mtd_by_store[scode] = sub_tgt

        # 3. Stock & Health
        stock_path = self.paths.get("stock", r"C:\All_Report\2_CleanData\Stocks\MART_Stock_Final_v89.csv")
        df_stock = pd.read_csv(
            stock_path,
            usecols=["StoreCode", "Quantity", "TotalValue", "Năm phân phối", "Size", "ProductName"],
            low_memory=False
        )
        
        stock_summary = {}
        for code, group in df_stock.groupby(df_stock["StoreCode"].astype(str).str.strip().str.upper()):
            total_qty = float(group["Quantity"].sum())
            total_val = float(group["TotalValue"].sum())
            aging_qty = float(group[group["Năm phân phối"] < curr_year]["Quantity"].sum())
            aging_pct = round((aging_qty / total_qty * 100), 1) if total_qty > 0 else 0.0
            stock_summary[code] = {
                "total_qty": int(total_qty),
                "total_val": round(total_val, 0),
                "aging_qty": int(aging_qty),
                "aging_pct": aging_pct
            }

        health_path = self.paths.get("health_master", r"C:\All_Report\8_RETAIL_COMMANDER\OutPut\02_Merchandising_Master\MART_HEALTH_MASTER_PERFECT.csv")
        df_health = pd.read_csv(health_path, dtype=str, low_memory=False)

        stores_dict = {}

        for code in store_codes:
            meta = store_meta[code]
            
            rev_mtd = rev_mtd_by_store.get(code, None)
            tgt_mtd = tgt_mtd_by_store.get(code, None)
            
            rev_status = "AVAILABLE" if (rev_mtd is not None and tgt_mtd is not None) else ("MISSING" if (rev_mtd is None and tgt_mtd is None) else "PARTIAL")
            
            if rev_mtd is not None and tgt_mtd is not None and tgt_mtd > 0:
                actual_progress_pct = round((rev_mtd / tgt_mtd) * 100, 2)
                pace_index = round((actual_progress_pct / expected_progress_pct), 4)
                pace_delta_pct = round(actual_progress_pct - expected_progress_pct, 2)
                
                gap = max(0.0, tgt_mtd - rev_mtd)
                req_daily = round(gap / selling_days_remaining, 0) if selling_days_remaining > 0 else 0.0
                actual_daily = round(rev_mtd / selling_days_elapsed, 0) if selling_days_elapsed > 0 else 0.0
            else:
                actual_progress_pct = None
                pace_index = None
                pace_delta_pct = None
                gap = None
                req_daily = None
                actual_daily = None

            stk = stock_summary.get(code, None)
            stk_status = "AVAILABLE" if stk is not None else "MISSING"
            
            # Stockout risks
            store_health = df_health[df_health["StoreCode"].astype(str).str.strip().str.upper() == code]
            stockouts = []
            if not store_health.empty and "Stockout_Risk" in store_health.columns:
                try:
                    risk_skus = store_health[store_health["Stockout_Risk"].astype(float) > 0.5]
                    if not risk_skus.empty:
                        for _, sk_row in risk_skus.head(3).iterrows():
                            stockouts.append({
                                "sku": str(sk_row.get("SKU", "")),
                                "name": str(sk_row.get("ProductName", "Sản phẩm chủ lực")),
                                "risk": "HIGH"
                            })
                except Exception:
                    pass

            # 4-Tier Severity Classification (Gate 1.5)
            if pace_index is not None:
                if pace_index >= 0.95:
                    severity = "PROTECT_ON_TRACK"
                    is_lagging = False
                elif pace_index >= 0.80:
                    severity = "WATCH"
                    is_lagging = True
                elif pace_index >= 0.65:
                    severity = "RECOVERY"
                    is_lagging = True
                else:
                    severity = "RESCUE_CRITICAL"
                    is_lagging = True
            else:
                severity = "UNKNOWN"
                is_lagging = False

            # Top Blockers with Priority Scoring (Gate 1.5)
            blockers = []
            
            # 1. Pace drop blocker
            if pace_delta_pct is not None and pace_delta_pct < -5.0:
                p_level = 1 if pace_delta_pct < -20.0 else 2
                blockers.append({
                    "priority": p_level,
                    "code": "PACE_DROP",
                    "category": "revenue",
                    "title": f"Tiến độ bán hàng chậm {abs(pace_delta_pct):.1f}%",
                    "detail": f"Đạt {actual_progress_pct}% (Pace Index: {pace_index*100:.1f}%) so với mốc kỳ vọng {expected_progress_pct}% của ngày {selling_days_elapsed}",
                    "severity": "CRITICAL" if p_level == 1 else "HIGH"
                })
                
            # 2. Stockout blocker
            if stockouts:
                blockers.append({
                    "priority": 1 if severity in ["RESCUE_CRITICAL", "RECOVERY"] else 2,
                    "code": "STOCKOUT_CORE",
                    "category": "inventory",
                    "title": f"Cảnh báo thiếu {len(stockouts)} mã hàng chủ lực",
                    "detail": f"Nguy cơ đứt hàng ở các mã: {', '.join([s['name'][:25] for s in stockouts])}",
                    "severity": "HIGH"
                })
                
            # 3. Aging stock blocker
            if stk and stk["aging_pct"] > 20.0:
                blockers.append({
                    "priority": 2 if stk["aging_pct"] > 50.0 else 3,
                    "code": "AGING_STOCK",
                    "category": "inventory",
                    "title": f"Tồn kho chậm luân chuyển cao ({stk['aging_pct']}%)",
                    "detail": f"Có {stk['aging_qty']:,} sản phẩm phân phối từ năm trước chiếm {stk['aging_pct']}% tổng tồn",
                    "severity": "HIGH" if stk["aging_pct"] > 50.0 else "MEDIUM"
                })
                
            # Sort blockers by priority ascending (1 = Primary, 2 = Secondary, 3 = Watch)
            blockers.sort(key=lambda b: b["priority"])

            card = {
                "store_code": code,
                "store_name": meta["name"],
                "region": meta["region"],
                "asm_name": meta["asm"],
                "manager": meta["manager"],
                "store_type": meta["store_type"],
                "data_quality_status": "AVAILABLE" if (rev_status == "AVAILABLE" and stk_status == "AVAILABLE") else "PARTIAL",
                "revenue": {
                    "status": rev_status,
                    "mtd_actual": rev_mtd,
                    "mtd_target": tgt_mtd,
                    "achievement_pct": actual_progress_pct,
                    "pace_index": pace_index,
                    "pace_delta_pct": pace_delta_pct,
                    "gap_amount": gap,
                    "selling_days_in_month": total_days_in_month,
                    "selling_days_elapsed": selling_days_elapsed,
                    "selling_days_remaining": selling_days_remaining,
                    "expected_progress_pct": expected_progress_pct,
                    "required_daily_runrate": req_daily,
                    "actual_daily_runrate": actual_daily
                },
                "inventory": {
                    "status": stk_status,
                    "total_qty": stk["total_qty"] if stk else None,
                    "total_value": stk["total_val"] if stk else None,
                    "aging_qty": stk["aging_qty"] if stk else None,
                    "aging_pct": stk["aging_pct"] if stk else None,
                    "top_stockout_skus": stockouts
                },
                "staff": {
                    "status": "AVAILABLE",
                    "target_headcount": meta["headcount_target"]
                },
                "diagnosis": {
                    "is_lagging": is_lagging,
                    "lag_severity": severity,
                    "primary_blocker": blockers[0] if blockers else None,
                    "top_blockers": blockers
                }
            }
            stores_dict[code] = card

        now_str = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S+07:00")
        snapshot = {
            "snapshot_metadata": {
                "generated_at": now_str,
                "snapshot_date": latest_rev_date.strftime("%Y-%m-%d"),
                "source_version": f"{latest_rev_date.strftime('%Y.%m.%d')}_v1.5",
                "diagnostic_version": "v1.5_semantic_audit",
                "total_stores": len(stores_dict),
                "active_month": f"{curr_year}-{curr_month:02d}",
                "selling_days_in_month": total_days_in_month,
                "selling_days_elapsed": selling_days_elapsed,
                "selling_days_remaining": selling_days_remaining,
                "expected_progress_pct": expected_progress_pct
            },
            "stores": stores_dict
        }
        return snapshot

    def export_snapshot_json(self, output_path: str = "data/store_diagnostics_snapshot.json") -> str:
        """Export snapshot to JSON file for WebApp/Local storage."""
        snapshot = self.generate_snapshot()
        abs_output = os.path.join(self.root_dir, output_path) if not os.path.isabs(output_path) else output_path
        os.makedirs(os.path.dirname(abs_output), exist_ok=True)
        
        with open(abs_output, "w", encoding="utf-8") as f:
            json.dump(snapshot, f, ensure_ascii=False, indent=2)
            
        logger.info(f"Exported diagnostic snapshot ({len(snapshot['stores'])} stores) to {abs_output}")
        return abs_output

    def sync_to_google_sheets(self, sheet_name: str = "Store_Diagnostics") -> bool:
        """Synchronize snapshot to Google Sheets tab 'Store_Diagnostics' for GAS API distribution."""
        try:
            import gspread
            from google.oauth2.service_account import Credentials
            
            google_cfg = self.config.get("google", {})
            creds_path = os.path.join(self.root_dir, google_cfg.get("credentials_path", "config/google_credentials.json"))
            spreadsheet_id = google_cfg.get("spreadsheet_id", "1Qetn0_6EYEaKQID3Ig4uwX55HhI6Au90Sz8LnVChmEM")
            
            if not os.path.exists(creds_path):
                logger.warning(f"Google credentials not found at {creds_path}. Skipping cloud sync.")
                return False
                
            snapshot = self.generate_snapshot()
            stores = snapshot.get("stores", {})
            
            rows = []
            headers = [
                "StoreCode", "StoreName", "Region", "ASM", "Manager", "StoreType",
                "DataStatus", "MTD_Actual", "MTD_Target", "Achievement_Pct", "Pace_Index", "Pace_Delta_Pct", "Gap_Amount",
                "SellingDaysRemaining", "RequiredDailyRunRate", "ActualDailyRunRate", "ExpectedProgress_Pct",
                "Stock_TotalQty", "Stock_TotalVal", "Stock_AgingPct", "StockoutCount",
                "TargetHeadcount", "IsLagging", "LagSeverity", "PrimaryBlocker_Title", "TopBlockers_JSON", "Diagnostic_JSON"
            ]
            
            for code, card in stores.items():
                r = card["revenue"]
                i = card["inventory"]
                s = card["staff"]
                d = card["diagnosis"]
                pb = d.get("primary_blocker")
                
                row = [
                    code,
                    card.get("store_name", ""),
                    card.get("region", ""),
                    card.get("asm_name", ""),
                    card.get("manager", ""),
                    card.get("store_type", ""),
                    card.get("data_quality_status", ""),
                    r.get("mtd_actual", 0) if r.get("mtd_actual") is not None else "",
                    r.get("mtd_target", 0) if r.get("mtd_target") is not None else "",
                    r.get("achievement_pct", "") if r.get("achievement_pct") is not None else "",
                    r.get("pace_index", "") if r.get("pace_index") is not None else "",
                    r.get("pace_delta_pct", "") if r.get("pace_delta_pct") is not None else "",
                    r.get("gap_amount", 0) if r.get("gap_amount") is not None else "",
                    r.get("selling_days_remaining", ""),
                    r.get("required_daily_runrate", "") if r.get("required_daily_runrate") is not None else "",
                    r.get("actual_daily_runrate", "") if r.get("actual_daily_runrate") is not None else "",
                    r.get("expected_progress_pct", ""),
                    i.get("total_qty", "") if i.get("total_qty") is not None else "",
                    i.get("total_value", "") if i.get("total_value") is not None else "",
                    i.get("aging_pct", "") if i.get("aging_pct") is not None else "",
                    len(i.get("top_stockout_skus", [])),
                    s.get("target_headcount", ""),
                    "YES" if d.get("is_lagging") else "NO",
                    d.get("lag_severity", "PROTECT_ON_TRACK"),
                    pb.get("title", "") if pb else "",
                    json.dumps(d.get("top_blockers", []), ensure_ascii=False),
                    json.dumps(card, ensure_ascii=False)
                ]
                rows.append(row)
                
            scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
            creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
            client = gspread.authorize(creds)
            ss = client.open_by_key(spreadsheet_id)
            
            try:
                diag_sheet = ss.worksheet(sheet_name)
            except Exception:
                diag_sheet = ss.add_worksheet(title=sheet_name, rows=250, cols=30)
                
            diag_sheet.clear()
            all_data = [headers] + rows
            diag_sheet.update(all_data, "A1")
            
            logger.info(f"Successfully synced {len(rows)} store diagnostics to Google Sheet '{sheet_name}'!")
            return True
        except Exception as e:
            logger.error(f"Error syncing to Google Sheets: {e}")
            return False


if __name__ == "__main__":
    pipeline = StoreDiagnosticPipeline()
    json_path = pipeline.export_snapshot_json()
    print(f"✓ Diagnostic snapshot generated at: {json_path}")
    synced = pipeline.sync_to_google_sheets()
    print(f"✓ Cloud synchronization status: {'SUCCESS' if synced else 'FAILED/SKIPPED'}")
