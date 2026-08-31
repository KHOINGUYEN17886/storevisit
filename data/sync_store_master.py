import os
import calendar
import datetime
import json
import logging
from typing import Dict, Any, Optional, List
import pandas as pd
import numpy as np

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(name)s: %(message)s")
logger = logging.getLogger("SyncStoreMaster")

def build_store_master_ssot():
    root_dir = r"c:\All_Report\8_RETAIL_COMMANDER\StoreVisit"
    stores_info_path = r"C:\All_Report\1_Mapping\StoresInfo.xlsx"
    
    if not os.path.exists(stores_info_path):
        raise FileNotFoundError(f"StoresInfo.xlsx not found at {stores_info_path}")
        
    logger.info(f"Reading StoresInfo.xlsx from {stores_info_path}...")
    df_stores = pd.read_excel(stores_info_path, sheet_name="Danh bạ CH")
    
    # 1. Parse Store Meta
    store_meta = {}
    asm_set = set()
    region_set = set()
    
    for _, r in df_stores.iterrows():
        code = str(r.get("CODE", r.get("STORECODE", ""))).strip().upper()
        if not code or code == "NAN":
            continue
            
        name = str(r.get("TÊN CỬA HÀNG - CHUẨN", r.get("CỬA HÀNG ", code))).strip()
        region = str(r.get("VÙNG", "HCM")).strip()
        asm = str(r.get("ASM", "Khác")).strip()
        manager = str(r.get("CỬA HÀNG TRƯỞNG ", "")).strip()
        phone = str(r.get("ĐIỆN THOẠI", "")).strip()
        email = str(r.get("EMAIL", "")).strip()
        address = str(r.get("ĐỊA CHỈ", "")).strip()
        brand_str = str(r.get("BRAND", "AP-PIE-A-B")).strip().upper()
        store_type = str(r.get("STORETYPE", "Standalone Boutique")).strip()
        
        try:
            headcount = int(float(r.get("ĐỊNH BIÊN", 4)))
        except:
            headcount = 4
            
        asm_set.add(asm)
        region_set.add(region)
        
        # Brand presence mapping
        has_ap = ("AP" in brand_str) or (brand_str == "ALL") or ("APPI" in brand_str)
        has_pie = ("PIE" in brand_str) or (brand_str == "ALL") or ("APPI" in brand_str)
        has_anamai = ("A-B" in brand_str) or ("AN" in brand_str) or ("-A-" in brand_str) or (brand_str == "AP-PIE-A-B") or (brand_str == "ALL")
        has_bonjour = ("A-B" in brand_str) or ("BJ" in brand_str) or ("-B" in brand_str) or (brand_str == "AP-PIE-A-B") or (brand_str == "ALL")
        
        # Concession / Mall Format rules
        is_concession = ("Concession" in store_type) or ("Mall" in store_type) or ("TTTM" in name) or ("Plaza" in name) or ("Vincom" in name) or ("Takashimaya" in name) or ("Diamond" in name)
        has_guard = not is_concession
        has_private_toilet = not is_concession
        
        store_meta[code] = {
            "code": code,
            "name": name,
            "display_name": f"{code} - {name}",
            "region": region,
            "asm": asm,
            "manager": manager,
            "phone": phone,
            "email": email,
            "address": address,
            "brand_raw": brand_str,
            "has_ap": has_ap,
            "has_pie": has_pie,
            "has_anamai": has_anamai,
            "has_bonjour": has_bonjour,
            "store_type": store_type,
            "is_concession": is_concession,
            "has_guard": has_guard,
            "has_private_toilet": has_private_toilet,
            "headcount_target": headcount
        }
        
    logger.info(f"Loaded {len(store_meta)} stores across {len(asm_set)} ASMs and {len(region_set)} Regions.")
    
    # 2. Revenue & Target from Data Lake
    rev_path = r"C:\All_Report\2_CleanData\Revenue\Fact_Revenue_AllYears_Standardized.csv"
    tgt_path = r"C:\All_Report\2_CleanData\Target\TargetMonthly.csv"
    stock_path = r"C:\All_Report\2_CleanData\Stocks\MART_Stock_Final_v89.csv"
    health_path = r"C:\All_Report\8_RETAIL_COMMANDER\OutPut\02_Merchandising_Master\MART_HEALTH_MASTER_PERFECT.csv"
    
    rev_mtd_by_store = {}
    rev_yoy_by_store = {}
    tgt_mtd_by_store = {}
    stock_summary = {}
    stockouts_by_store = {}
    
    selling_days_elapsed = 28
    total_days_in_month = 31
    selling_days_remaining = 3
    expected_progress_pct = 90.32
    
    curr_year = 2026
    curr_month = 8
    
    if os.path.exists(rev_path):
        logger.info("Reading Revenue data...")
        df_rev = pd.read_csv(rev_path, parse_dates=["Date"], low_memory=False)
        latest_rev_date = df_rev["Date"].max()
        curr_year = latest_rev_date.year
        curr_month = latest_rev_date.month
        curr_day = latest_rev_date.day
        _, total_days_in_month = calendar.monthrange(curr_year, curr_month)
        selling_days_elapsed = curr_day
        selling_days_remaining = max(1, total_days_in_month - selling_days_elapsed)
        expected_progress_pct = round((selling_days_elapsed / total_days_in_month) * 100, 2)
        
        # MTD
        df_rev_curr = df_rev[(df_rev["Date"].dt.year == curr_year) & (df_rev["Date"].dt.month == curr_month)]
        rev_mtd_by_store = df_rev_curr.groupby(df_rev_curr["StoreCode"].str.strip().str.upper())["Revenue"].sum().to_dict()
        
        # YoY (Same month last year)
        df_rev_yoy = df_rev[(df_rev["Date"].dt.year == (curr_year - 1)) & (df_rev["Date"].dt.month == curr_month)]
        rev_yoy_by_store = df_rev_yoy.groupby(df_rev_yoy["StoreCode"].str.strip().str.upper())["Revenue"].sum().to_dict()

    if os.path.exists(tgt_path):
        logger.info("Reading Target data...")
        df_tgt = pd.read_csv(tgt_path, low_memory=False)
        df_tgt_curr = df_tgt[(df_tgt["Year"] == curr_year) & (df_tgt["MonthNo"] == curr_month)]
        store_total_tgt = df_tgt_curr[df_tgt_curr["BrandGroup"] == "Store"].groupby(df_tgt_curr["StoreCode"].astype(str).str.strip().str.upper())["Target"].sum().to_dict()
        sub_brands_tgt = df_tgt_curr[df_tgt_curr["BrandGroup"].isin(["APPI", "AB"])].groupby(df_tgt_curr["StoreCode"].astype(str).str.strip().str.upper())["Target"].sum().to_dict()
        
        for scode, s_tgt in store_total_tgt.items():
            tgt_mtd_by_store[scode] = s_tgt
        for scode, sub_tgt in sub_brands_tgt.items():
            if scode not in tgt_mtd_by_store or tgt_mtd_by_store[scode] <= 0:
                tgt_mtd_by_store[scode] = sub_tgt

    if os.path.exists(stock_path):
        logger.info("Reading Stock data...")
        df_stock = pd.read_csv(stock_path, usecols=["StoreCode", "Quantity", "TotalValue", "Năm phân phối"], low_memory=False)
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

    if os.path.exists(health_path):
        logger.info("Reading Health Master...")
        df_health = pd.read_csv(health_path, dtype=str, low_memory=False)
        if "Stockout_Risk" in df_health.columns:
            for code, group in df_health.groupby(df_health["StoreCode"].astype(str).str.strip().str.upper()):
                try:
                    risk_skus = group[group["Stockout_Risk"].astype(float) > 0.5]
                    stockouts_by_store[code] = [
                        {"sku": str(r.get("SKU", "")), "name": str(r.get("ProductName", "Mã chủ lực")), "risk": "HIGH"}
                        for _, r in risk_skus.head(3).iterrows()
                    ]
                except:
                    pass

    # 3. Build Full Diagnostic Cards for all 185 stores
    diagnostics = {}
    for code, meta in store_meta.items():
        rev_mtd = rev_mtd_by_store.get(code, 0.0)
        rev_yoy = rev_yoy_by_store.get(code, 0.0)
        tgt_mtd = tgt_mtd_by_store.get(code, 0.0)
        
        if tgt_mtd > 0:
            ach_pct = round((rev_mtd / tgt_mtd) * 100, 1)
            pace_index = round((ach_pct / expected_progress_pct), 2)
            pace_delta = round(ach_pct - expected_progress_pct, 1)
            gap = max(0.0, tgt_mtd - rev_mtd)
            req_daily = round(gap / selling_days_remaining, 0)
            actual_daily = round(rev_mtd / selling_days_elapsed, 0)
        else:
            ach_pct = 0.0
            pace_index = 0.0
            pace_delta = 0.0
            gap = 0.0
            req_daily = 0.0
            actual_daily = 0.0
            
        yoy_growth_pct = round(((rev_mtd - rev_yoy) / rev_yoy * 100), 1) if rev_yoy > 0 else 0.0
        
        # Severity
        if tgt_mtd > 0:
            if pace_index >= 0.95:
                severity = "PROTECT_ON_TRACK"
            elif pace_index >= 0.80:
                severity = "WATCH"
            elif pace_index >= 0.65:
                severity = "RECOVERY"
            else:
                severity = "RESCUE_CRITICAL"
        else:
            severity = "PROTECT_ON_TRACK"
            
        stk = stock_summary.get(code, {"total_qty": 0, "total_val": 0, "aging_qty": 0, "aging_pct": 0.0})
        stockouts = stockouts_by_store.get(code, [])
        
        diagnostics[code] = {
            "meta": meta,
            "severity": severity,
            "revenue": {
                "mtd_actual": rev_mtd,
                "mtd_target": tgt_mtd,
                "mtd_yoy": rev_yoy,
                "yoy_growth_pct": yoy_growth_pct,
                "achievement_pct": ach_pct,
                "pace_index": pace_index,
                "pace_delta_pct": pace_delta,
                "gap_amount": gap,
                "req_daily": req_daily,
                "actual_daily": actual_daily,
                "selling_days_elapsed": selling_days_elapsed,
                "selling_days_remaining": selling_days_remaining,
                "total_days_in_month": total_days_in_month,
                "expected_progress_pct": expected_progress_pct
            },
            "inventory": stk,
            "stockouts": stockouts
        }

    # 4. Group stores by ASM and by Region for Dropdown Hydration
    asm_store_map = {}
    region_store_map = {}
    
    for code, meta in store_meta.items():
        asm = meta["asm"]
        reg = meta["region"]
        
        if asm not in asm_store_map:
            asm_store_map[asm] = []
        asm_store_map[asm].append({"code": code, "name": meta["name"], "display": meta["display_name"]})
        
        if reg not in region_store_map:
            region_store_map[reg] = []
        region_store_map[reg].append({"code": code, "name": meta["name"], "display": meta["display_name"]})

    # Sort ASMs and Regions
    sorted_asms = sorted(list(asm_set))
    sorted_regions = sorted(list(region_set))
    
    # Bundle into SSOT payload
    ssot_bundle = {
        "generated_at": datetime.datetime.now().isoformat(),
        "total_stores": len(store_meta),
        "total_asms": len(sorted_asms),
        "asms": sorted_asms,
        "regions": sorted_regions,
        "stores_by_asm": asm_store_map,
        "stores_by_region": region_store_map,
        "profiles": store_meta,
        "diagnostics": diagnostics
    }
    
    # 5. Output JS bundle and JSON
    js_output_path = os.path.join(root_dir, "webapp", "store_manifest_ssot.js")
    json_output_path = os.path.join(root_dir, "data", "store_manifest_ssot.json")
    
    with open(json_output_path, "w", encoding="utf-8") as f:
        json.dump(ssot_bundle, f, ensure_ascii=False, indent=2)
    logger.info(f"✓ Saved JSON snapshot to {json_output_path}")
    
    js_code = f"""// AUTO-GENERATED BY data/sync_store_master.py - ZERO HARDCODING SSOT MANIFEST
// Generated At: {ssot_bundle['generated_at']} | Total Stores: {len(store_meta)}
window.STORE_DATA_SSOT = {json.dumps(ssot_bundle, ensure_ascii=False)};
window.STORE_PROFILES_SSOT = window.STORE_DATA_SSOT.profiles;
window.STORE_DIAGNOSTICS_SSOT = window.STORE_DATA_SSOT.diagnostics;
"""
    with open(js_output_path, "w", encoding="utf-8") as f:
        f.write(js_code)
    logger.info(f"✓ Saved JS bundle to {js_output_path}")

    # Also save to data/sync_store_master.py
    with open(os.path.join(root_dir, "data", "sync_store_master.py"), "w", encoding="utf-8") as f:
        with open(__file__, "r", encoding="utf-8") as cur_f:
            f.write(cur_f.read())

    print("\n==========================================================================")
    print("STORE MASTER SSOT PIPELINE SUMMARY:")
    print("==========================================================================")
    print(f"Total Stores Processed: {len(store_meta)}")
    print(f"Total ASMs: {len(sorted_asms)} -> {sorted_asms}")
    print(f"Total Regions: {len(sorted_regions)} -> {sorted_regions}")
    print(f"Concession Stores (Auto N/A Guard & WC): {sum(1 for m in store_meta.values() if m['is_concession'])}")
    print(f"AP-PIE-A-B Stores (Full 4 Brands): {sum(1 for m in store_meta.values() if m['has_anamai'] and m['has_bonjour'])}")
    print(f"AP-PIE Only Stores (Auto N/A Anamai/Bonjour): {sum(1 for m in store_meta.values() if m['has_ap'] and m['has_pie'] and not m['has_anamai'])}")
    print("==========================================================================\n")

if __name__ == "__main__":
    build_store_master_ssot()
