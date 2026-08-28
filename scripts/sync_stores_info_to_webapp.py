# -*- coding: utf-8 -*-
"""
sync_stores_info_to_webapp.py
=============================
Đọc C:\All_Report\1_Mapping\StoresInfo.xlsx (Sheet "Danh bạ CH") và tự động đồng bộ:
1. webapp/store_mapping.json
2. webapp/store_profile_map.json
3. webapp/store_profile_map.js & webapp/StoreProfileMap.gs
4. Cập nhật STORE_DATA_MAP và initASMUsersSheet() trong Code.gs
"""
import os
import json
import openpyxl
import pandas as pd

SRC_EXCEL = r"C:\All_Report\1_Mapping\StoresInfo.xlsx"
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
WEBAPP_DIR = os.path.join(BASE_DIR, "webapp")

GUARD_TYPES = {"Standalone Boutique", "Maison", "Flagship Store", "Mono-brand Store"}
NO_SURVEY_TYPES = {"Online Store"}

def brand_tokens(brand: str):
    b = (brand or "").strip().upper()
    if not b or b == "0":
        return set()
    return set(t.strip() for t in b.split("-") if t.strip())

def sync_all():
    print(f"=== Đọc dữ liệu từ: {SRC_EXCEL} ===")
    df = pd.read_excel(SRC_EXCEL, sheet_name="Danh bạ CH")
    print(f"Tổng số dòng cửa hàng: {len(df)}")

    mapping_by_asm = {}
    mapping_by_region = {}
    asm_store_codes = {}
    asm_regions = {}
    profiles = {}

    for idx, row in df.iterrows():
        code = str(row["CODE"]).strip()
        name = str(row["TÊN CỬA HÀNG - CHUẨN"]).strip()
        asm = str(row["ASM"]).strip()
        region = str(row["VÙNG"]).strip()
        brand = str(row["BRAND"]).strip() if pd.notna(row["BRAND"]) else ""
        stype = str(row["STORETYPE"]).strip() if pd.notna(row["STORETYPE"]) else ""

        label = f"{code} - {name}"

        # 1. Grouping by ASM
        if asm not in mapping_by_asm:
            mapping_by_asm[asm] = []
            asm_store_codes[asm] = []
            asm_regions[asm] = set()
        mapping_by_asm[asm].append(label)
        asm_store_codes[asm].append(code)
        if region:
            asm_regions[asm].add(region)

        # 2. Grouping by Region
        if region not in mapping_by_region:
            mapping_by_region[region] = []
        mapping_by_region[region].append(label)

        # 3. Store Profiles
        tokens = brand_tokens(brand)
        has_ap = "AP" in tokens
        has_pie = "PIE" in tokens
        has_anamai = "A" in tokens
        has_bonjour = "B" in tokens
        is_ab_only = (has_anamai or has_bonjour) and not (has_ap or has_pie)
        skip_survey = (stype in NO_SURVEY_TYPES) or (not tokens)
        has_guard = (stype in GUARD_TYPES) and (has_ap or has_pie) and not is_ab_only

        profiles[code] = {
            "brand": brand,
            "storetype": stype,
            "has_ap": has_ap,
            "has_pie": has_pie,
            "has_anamai": has_anamai,
            "has_bonjour": has_bonjour,
            "is_ab_only": is_ab_only,
            "has_guard": has_guard,
            "skip_survey": skip_survey,
        }

    # Save store_mapping.json
    store_mapping_data = {
        "asms": sorted(list(mapping_by_asm.keys())),
        "regions": sorted(list(mapping_by_region.keys())),
        "mapping_by_asm": mapping_by_asm,
        "mapping_by_region": mapping_by_region
    }
    mapping_json_path = os.path.join(WEBAPP_DIR, "store_mapping.json")
    with open(mapping_json_path, "w", encoding="utf-8") as f:
        json.dump(store_mapping_data, f, ensure_ascii=False, indent=2)
    print(f"✓ Đã ghi: {mapping_json_path}")

    # Save store_profile_map.json
    profile_json_path = os.path.join(WEBAPP_DIR, "store_profile_map.json")
    with open(profile_json_path, "w", encoding="utf-8") as f:
        json.dump(profiles, f, ensure_ascii=False, indent=2)
    print(f"✓ Đã ghi: {profile_json_path}")

    # Save store_profile_map.js & StoreProfileMap.gs
    js_content = "// AUTO-GENERATED bởi sync_stores_info_to_webapp.py — KHÔNG sửa tay.\n"
    js_content += "// Nguồn: StoresInfo.xlsx (cột BRAND, STORECODE, STORETYPE).\n"
    js_content += "var STORE_PROFILE_MAP = " + json.dumps(profiles, ensure_ascii=False, indent=2) + ";\n"

    for fname in ["store_profile_map.js", "StoreProfileMap.gs"]:
        fpath = os.path.join(WEBAPP_DIR, fname)
        with open(fpath, "w", encoding="utf-8") as f:
            f.write(js_content)
        print(f"✓ Đã ghi: {fpath}")

    print(f"\nTổng số ASM: {len(mapping_by_asm)}")
    print(f"Tổng số Vùng: {len(mapping_by_region)}")
    print(f"Tổng số Cửa hàng: {sum(len(v) for v in mapping_by_asm.values())}")

    return store_mapping_data, asm_store_codes, asm_regions

if __name__ == "__main__":
    sync_all()
