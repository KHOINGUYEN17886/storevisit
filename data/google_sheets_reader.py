import os
import re
import json
import logging
from typing import List, Dict, Any, Optional
import gspread
from google.oauth2.service_account import Credentials
from data.models import StoreFormResponse, FormPhoto, MarketSurveyResponse, SurveyPhoto

logger = logging.getLogger(__name__)

class GoogleSheetsReader:
    def __init__(self, credentials_path: str, spreadsheet_id: str):
        self.credentials_path = credentials_path
        self.spreadsheet_id = spreadsheet_id
        self.client = None
        
    def _authenticate(self):
        if not self.client:
            if not os.path.exists(self.credentials_path):
                raise FileNotFoundError(f"Google credentials file not found at {self.credentials_path}")
            scopes = [
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"
            ]
            creds = Credentials.from_service_account_file(self.credentials_path, scopes=scopes)
            self.client = gspread.authorize(creds)
            
    def get_pending_responses(self, sheet_name: str = "Form Responses 1") -> List[StoreFormResponse]:
        self._authenticate()
        if not self.spreadsheet_id:
            logger.warning("Spreadsheet ID is empty. Returning empty list.")
            return []
        
        try:
            sheet = self.client.open_by_key(self.spreadsheet_id).worksheet(sheet_name)
        except Exception as e:
            logger.error(f"Error opening sheet {sheet_name}: {e}")
            raise e
            
        records = sheet.get_all_records()
        
        responses = []
        for i, row in enumerate(records):
            # Row index in sheet (1-based, plus header row, so records[i] corresponds to row i + 2)
            row_idx = i + 2
            
            status = str(row.get("Status", "")).strip().lower()
            if status in ["done", "draft"]:
                continue
                
            resp = self._parse_row(row, str(row_idx))
            if resp:
                responses.append(resp)
        return responses
        
    def mark_as_done(self, row_idx: str, sheet_name: str = "Form Responses 1"):
        self._authenticate()
        if not self.spreadsheet_id:
            return
        try:
            sheet = self.client.open_by_key(self.spreadsheet_id).worksheet(sheet_name)
            headers = sheet.row_values(1)
            
            if "Status" not in headers:
                col_idx = len(headers) + 1
                sheet.update_cell(1, col_idx, "Status")
            else:
                col_idx = headers.index("Status") + 1
                
            sheet.update_cell(int(row_idx), col_idx, "done")
            logger.info(f"Successfully marked row {row_idx} as done in Sheet.")
        except Exception as e:
            logger.error(f"Error marking row {row_idx} as done: {e}")
            raise e
        
    def _parse_row(self, row: Dict[str, Any], row_id: str) -> Optional[StoreFormResponse]:
        import json
        def find_val(keywords: List[str], default: Any = "") -> Any:
            for k, v in row.items():
                k_lower = k.lower()
                if any(kw.lower() in k_lower for kw in keywords):
                    return v
            return default

        store_code = str(find_val(["Mã cửa hàng", "Ma cua hang", "Store Code"])).strip()
        if not store_code:
            return None
            
        if " - " in store_code:
            store_code = store_code.split(" - ")[0].strip()
            
        report_date = str(find_val(["Ngày kiểm tra", "Ngay kiem tra", "Date"])).strip()
        asm_name = str(find_val(["QLKD/ASM", "ASM", "Người kiểm tra", "Nguoi kiem tra"])).strip()
        cht_name = str(find_val(["Tên CHT", "Ten CHT", "Cửa hàng trưởng"])).strip()
        time_start = str(find_val(["Giờ bắt đầu", "Gio bat dau", "Time start"])).strip()
        time_end = str(find_val(["Giờ kết thúc", "Gio ket thuc", "Time end"])).strip()
        
        nv_count_raw = find_val(["Số NV", "So NV", "Nhân viên có mặt"], 0)
        try:
            nv_count = int(float(nv_count_raw)) if nv_count_raw else 0
        except ValueError:
            nv_count = 0
            
        rating_frontage = str(find_val(["Đánh giá mặt tiền", "Danh gia mat tien", "Exterior rating"], "Đạt")).strip()
        rating_inner = str(find_val(["Đánh giá bên trong", "Danh gia ben trong", "Inner rating"], "Đạt")).strip()
        rating_merch = str(find_val(["Đánh giá hàng hóa", "Danh gia hang hoa", "Merchandise rating"], "Đạt")).strip()
        rating_staff = str(find_val(["Đánh giá nhân sự", "Danh gia nhan su", "Staff rating"], "Đạt")).strip()
        rating_csvc = str(find_val(["Đánh giá CSVC", "Danh gia csvc", "CSVC rating"], "Đạt")).strip()
        
        comment_frontage = str(find_val(["Nhận xét mặt tiền", "Nhan xet mat tien", "Exterior comments"])).strip()
        comment_inner = str(find_val(["Nhận xét bên trong", "Nhan xet ben trong", "Inner comments"])).strip()
        comment_merch = str(find_val(["Nhận xét hàng hóa", "Nhan xet hang hoa", "Merchandise comments"])).strip()
        comment_staff = str(find_val(["Nhận xét nhân sự", "Nhan xet nhan su", "Staff comments"])).strip()
        comment_csvc = str(find_val(["Nhận xét CSVC", "Nhan xet csvc", "CSVC comments"])).strip()
        
        pending_issues = str(find_val(["Vấn đề tồn đọng", "Van de ton dong", "Issues"])).strip()
        action_plan = str(find_val(["Kế hoạch khắc phục", "Ke hoach khac phuc", "Action plan"])).strip()
        action_deadline = str(find_val(["Thời hạn xử lý", "Thoi han xu ly", "Deadline"])).strip()
        store_recommendation = str(find_val(["Đề xuất phát triển", "store_recommendation", "storeRecommendation"], "")).strip()

        # Đồng bộ webapp 30-07: phân loại lượt kiểm tra + field riêng cho khai trương
        inspection_mode = str(find_val(["inspection_mode"], "own")).strip().lower() or "own"
        opening_type = str(find_val(["opening_type"], "")).strip() or None
        opening_phase = str(find_val(["opening_phase"], "")).strip() or None
        opening_date = str(find_val(["opening_date"], "")).strip() or None
        opening_readiness = str(find_val(["opening_readiness"], "")).strip() or None

        # Parse checklist_json if present in the row
        checklist_json_val = find_val(["checklist_json", "checklist"], "")
        checklist_json_str = str(checklist_json_val).strip() if checklist_json_val else ""
        if checklist_json_str:
            checklist_json_str = re.sub(r"\[object\s+Object\]", "", checklist_json_str, flags=re.IGNORECASE)
            checklist_json_str = re.sub(r"\s*\[o\]?\s*(?=\\n|\"|\'|,)", "", checklist_json_str, flags=re.IGNORECASE)
        checklist_data = {}
        if checklist_json_str:
            try:
                checklist_data = json.loads(checklist_json_str)
                # Sanitize any nested survey answers
                if isinstance(checklist_data, dict) and "survey" in checklist_data:
                    for s_k, s_v in checklist_data["survey"].items():
                        if isinstance(s_v, dict) and "answer" in s_v:
                            ans = str(s_v["answer"])
                            ans = re.sub(r"\[object\s+Object\]", "", ans, flags=re.IGNORECASE).strip()
                            ans = re.sub(r"\s*\[o\]?\s*$", "", ans, flags=re.IGNORECASE).strip()
                            s_v["answer"] = ans
                    checklist_json_str = json.dumps(checklist_data, ensure_ascii=False)
            except Exception as e:
                logger.error(f"Error parsing checklist_json for row {row_id}: {e}")

        if checklist_data:
            sections = checklist_data.get("sections", {})
            
             # 1. Frontage override
            f_sec = sections.get("frontage", {})
            if f_sec.get("rating"):
                rating_frontage = f_sec.get("rating")
            if f_sec.get("comment"):
                comment_frontage = f_sec.get("comment")
                
            # 1.5. Inner override
            inner_sec = sections.get("inner", {})
            if inner_sec.get("rating"):
                rating_inner = inner_sec.get("rating")
            if inner_sec.get("comment"):
                comment_inner = inner_sec.get("comment")
                
            # 2. Merchandise override (combine from merch_ap, merch_pie, merch_ab, merch_anamai, merch_bonjour, merch_pk)
            m_ratings = []
            m_comments = []
            for m_key in ["merch_ap", "merch_pie", "merch_ab", "merch_anamai", "merch_bonjour", "merch_pk"]:
                m_sec = sections.get(m_key, {})
                if m_sec.get("rating"):
                    m_ratings.append(m_sec.get("rating"))
                if m_sec.get("comment"):
                    label_map = {
                        "merch_ap": "An Phước",
                        "merch_pie": "Pierre Cardin",
                        "merch_ab": "Anamai/Bonjour",
                        "merch_anamai": "Anamai",
                        "merch_bonjour": "Bonjour",
                        "merch_pk": "Phụ kiện"
                    }
                    m_comments.append(f"{label_map.get(m_key)}: {m_sec.get('comment')}")
            if m_ratings:
                if "Chưa đạt" in m_ratings:
                    rating_merch = "Chưa đạt"
                elif all(r == "Tốt" for r in m_ratings):
                    rating_merch = "Tốt"
                else:
                    rating_merch = "Đạt"
            if m_comments:
                comment_merch = " | ".join(m_comments)
                
            # 3. Staff override
            s_sec = sections.get("staff", {})
            if s_sec.get("rating"):
                rating_staff = s_sec.get("rating")
            if s_sec.get("comment"):
                comment_staff = s_sec.get("comment")
                
            # 4. CSVC override (combine from warehouse, cashier, stockroom, fitting_room, toilet, fire_safety, packaging_security)
            c_ratings = []
            c_comments = []
            for c_key in ["warehouse", "cashier", "stockroom", "fitting_room", "toilet", "fire_safety", "packaging_security"]:
                c_sec = sections.get(c_key, {})
                if c_sec.get("rating"):
                    c_ratings.append(c_sec.get("rating"))
                if c_sec.get("comment"):
                    c_comments.append(f"{c_key.upper()}: {c_sec.get('comment')}")
            if c_ratings:
                if "Chưa đạt" in c_ratings:
                    rating_csvc = "Chưa đạt"
                elif all(r == "Tốt" for r in c_ratings):
                    rating_csvc = "Tốt"
                else:
                    rating_csvc = "Đạt"
            if c_comments:
                comment_csvc = " | ".join(c_comments)

        photos = []

        def _normalize_drive_url(raw: str) -> str:
            """Normalize a Drive file ID or URL to a canonical https URL."""
            raw = raw.strip()
            if raw.startswith("http"):
                return raw
            # bare file-ID (no slashes, no dots, length >= 15)
            if len(raw) >= 15 and " " not in raw and "/" not in raw and "." not in raw:
                return f"https://drive.google.com/open?id={raw}"
            return ""

        def _is_valid_drive_ref(raw: str) -> bool:
            if not raw:
                return False
            raw = raw.strip()
            if raw.startswith("http"):
                return True
            if len(raw) >= 15 and " " not in raw and "/" not in raw and "." not in raw:
                return True
            return False

        def _get_drive_refs(raw: str) -> list:
            if not raw or not isinstance(raw, str):
                return []
            parts = []
            for part in raw.replace("\n", ",").split(","):
                part = part.strip()
                if not part:
                    continue
                if _is_valid_drive_ref(part):
                    parts.append(part)
            return parts

        if checklist_data:
            # ── SOURCE A: general_photos (canonical positions for frontage / inner / CSVC) ──
            g_photos = checklist_data.get("general_photos", {})
            g_mapping = [
                ("frontage_main",  "frontage",    1),
                ("frontage_left",  "frontage",    2),
                ("frontage_right", "frontage",    3),
                ("inner_entrance", "inner",        1),
                ("inner_left",     "inner",        2),
                ("inner_right",    "inner",        3),
                ("stockroom",      "stockroom",    1),
                ("fitting_room",   "fitting_room", 1),
                ("cashier",        "cashier",      1),
                # Đồng bộ 30-07: ảnh trước/sau sửa chữa cho báo cáo khai trương (tái khai trương)
                ("opening_before", "opening_before", 1),
                ("opening_after",  "opening_after",  1),
            ]
            seen_urls: set = set()
            for g_key, sec, idx in g_mapping:
                na_key = f"{g_key}_na"
                if g_photos.get(na_key) is True:
                    continue
                val = g_photos.get(g_key, "")
                refs = _get_drive_refs(val)
                for ref_idx, ref in enumerate(refs):
                    url = _normalize_drive_url(ref)
                    if url and url not in seen_urls:
                        seen_urls.add(url)
                        photos.append(FormPhoto(section=sec, index=idx + ref_idx, drive_url=url))

            # ── SOURCE B: VM photos from checklist item photo_before / photo_after ──
            sections_data = checklist_data.get("sections", {})
            vm_sec_map = {
                "merch_ap":      "vm_ap",
                "merch_pie":     "vm_pie",
                "merch_ab":      "vm_ab",
                "merch_anamai":  "vm_ab",
                "merch_bonjour": "vm_ab",
                "merch_pk":      "vm_pk",
            }
            for sec_key, sec_val in sections_data.items():
                p_sec = vm_sec_map.get(sec_key)
                if not p_sec:
                    continue
                for item in sec_val.get("items", []):
                    for p_type, p_idx in [("photo_before", 1), ("photo_after", 2), ("photo_detail", 3)]:
                        raw = item.get(p_type, "")
                        refs = _get_drive_refs(raw)
                        for ref_idx, ref in enumerate(refs):
                            url = _normalize_drive_url(ref)
                            if url and url not in seen_urls:
                                seen_urls.add(url)
                                photos.append(FormPhoto(section=p_sec, index=p_idx + ref_idx, drive_url=url))

            # ── SOURCE C: Competitor photos ──
            comp_sec = checklist_data.get("competitor", {})
            for p_key, p_idx in [("photo1", 1), ("photo2", 2), ("photo3", 3)]:
                raw = comp_sec.get(p_key, "")
                refs = _get_drive_refs(raw)
                for ref_idx, ref in enumerate(refs):
                    url = _normalize_drive_url(ref)
                    if url and url not in seen_urls:
                        seen_urls.add(url)
                        photos.append(FormPhoto(section="competitor", index=p_idx + ref_idx, drive_url=url))

            # ── SOURCE D: Issue photos (pending_issues[].photo_before / photo_after) ──
            for issue in checklist_data.get("pending_issues", []):
                src_sec = issue.get("source_section", "inner")
                
                # Before photos
                raw_before = issue.get("photo_before", "")
                refs_before = _get_drive_refs(raw_before)
                for ref_idx, ref in enumerate(refs_before):
                    url = _normalize_drive_url(ref)
                    if url and url not in seen_urls:
                        seen_urls.add(url)
                        photos.append(FormPhoto(section=f"issue_{src_sec}_before", index=1 + ref_idx, drive_url=url))
                        
                # After photos
                raw_after = issue.get("photo_after", "")
                refs_after = _get_drive_refs(raw_after)
                for ref_idx, ref in enumerate(refs_after):
                    url = _normalize_drive_url(ref)
                    if url and url not in seen_urls:
                        seen_urls.add(url)
                        photos.append(FormPhoto(section=f"issue_{src_sec}_after", index=1 + ref_idx, drive_url=url))

        else:
            # ── FALLBACK: No checklist_json — parse from old Google Sheet columns ──
            sections_mapping = {
                "frontage":    ["Ảnh mặt tiền",  "Anh mat tien",  "Exterior photo"],
                "inner":       ["Ảnh bên trong",  "Anh ben trong",  "Inner photo"],
                "merchandise": ["Ảnh hàng hóa",   "Anh hang hoa",   "Merchandise photo"],
                "staff":       ["Ảnh nhân sự",    "Anh nhan su",    "Staff photo"],
                "csvc":        ["Ảnh CSVC",        "Anh CSVC",       "CSVC photo"],
            }
            seen_urls = set()
            for section, keywords in sections_mapping.items():
                matching_keys = [k for k in row.keys() if any(kw.lower() in k.lower() for kw in keywords)]
                matching_keys.sort()
                idx = 1
                for k in matching_keys:
                    urls_raw = str(row[k]).strip()
                    if not urls_raw:
                        continue
                    for url in urls_raw.replace("\n", ",").split(","):
                        url = url.strip()
                        if url and _is_valid_drive_ref(url):
                            norm = _normalize_drive_url(url)
                            if norm and norm not in seen_urls:
                                seen_urls.add(norm)
                                photos.append(FormPhoto(section=section, index=idx, drive_url=norm))
                                idx += 1
                                if idx > 3:
                                    break
                    if idx > 3:
                        break

        # ── Re-index within each section to ensure consecutive 1-based indices ──
        for sec in ["frontage", "inner", "competitor", "stockroom", "fitting_room", "cashier", "csvc",
                    "vm_ap", "vm_pie", "vm_ab", "vm_pk"]:
            sec_photos = [p for p in photos if p.section == sec]
            for i, p in enumerate(sec_photos):
                p.index = i + 1

        return StoreFormResponse(
            response_id=row_id,
            store_code=store_code,
            report_date=report_date,
            asm_name=asm_name,
            cht_name=cht_name,
            time_start=time_start,
            time_end=time_end,
            nv_count=nv_count,
            rating_frontage=rating_frontage,
            rating_inner=rating_inner,
            rating_merch=rating_merch,
            rating_staff=rating_staff,
            rating_csvc=rating_csvc,
            comment_frontage=comment_frontage,
            comment_inner=comment_inner,
            comment_merch=comment_merch,
            comment_staff=comment_staff,
            comment_csvc=comment_csvc,
            pending_issues=pending_issues,
            action_plan=action_plan,
            action_deadline=action_deadline,
            store_recommendation=store_recommendation,
            photos=photos,
            status=str(row.get("Status", "pending")).strip(),
            checklist_json=checklist_json_str,
            inspection_mode=inspection_mode,
            opening_type=opening_type,
            opening_phase=opening_phase,
            opening_date=opening_date,
            opening_readiness=opening_readiness,
        )

    def get_pending_survey_responses(self, sheet_name: str = "MarketSurvey_Responses") -> List[MarketSurveyResponse]:
        self._authenticate()
        if not self.spreadsheet_id:
            logger.warning("Spreadsheet ID is empty. Returning empty list.")
            return []
        
        try:
            sheet = self.client.open_by_key(self.spreadsheet_id).worksheet(sheet_name)
        except Exception as e:
            logger.error(f"Error opening sheet {sheet_name}: {e}")
            raise e
            
        records = sheet.get_all_records()
        
        responses = []
        for i, row in enumerate(records):
            row_idx = i + 2
            status = str(row.get("Status", "")).strip().lower()
            if status in ["done", "processing", "ignored"]:
                continue
                
            resp = self._parse_survey_row(row, str(row_idx))
            if resp:
                responses.append(resp)
        return responses

    def update_row_status(self, row_idx: str, status: str, sheet_name: str, error_msg: str = None):
        self._authenticate()
        if not self.spreadsheet_id:
            return
        try:
            sheet = self.client.open_by_key(self.spreadsheet_id).worksheet(sheet_name)
            headers = [h.strip() for h in sheet.row_values(1)]
            
            def ensure_col(name: str) -> int:
                if name not in headers:
                    c_idx = len(headers) + 1
                    sheet.update_cell(1, c_idx, name)
                    headers.append(name)
                    return c_idx
                return headers.index(name) + 1
                
            status_col = ensure_col("Status")
            sheet.update_cell(int(row_idx), status_col, status)
            
            if error_msg is not None:
                err_col = ensure_col("Error_Message")
                sheet.update_cell(int(row_idx), err_col, error_msg)
                
            if status == "done":
                from datetime import datetime
                time_col = ensure_col("Processed_At")
                sheet.update_cell(int(row_idx), time_col, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                err_col = ensure_col("Error_Message")
                sheet.update_cell(int(row_idx), err_col, "")
                
            logger.info(f"Successfully updated row {row_idx} to status={status} in sheet {sheet_name}")
        except Exception as e:
            logger.error(f"Error updating row {row_idx} status in sheet {sheet_name}: {e}")
            raise e

    def _parse_survey_row(self, row: Dict[str, Any], row_id: str) -> Optional[MarketSurveyResponse]:
        columns_map = {
            "store_code": ["[MS01]", "Mã cửa hàng", "Ma cua hang", "Store Code"],
            "region": ["[MS02]", "Khu vực", "Cum cua hang", "Cụm cửa hàng", "Region"],
            "qlkd_asm": ["[MS03]", "QLKD/ASM", "ASM phụ trách", "ASM phu trach", "Người phụ trách"],
            "respondent_name": ["[MS04]", "Người đại diện", "Nguoi dai dien", "Họ tên người trả lời"],
            "respondent_role": ["[MS05]", "Chức danh", "Chuc danh", "Respondent role"],
            "discussion_count": ["[MS06]", "Số người", "So nguoi", "Discussion count"],
            "survey_date": ["[MS07]", "Ngày thực hiện", "Ngay thuc hien", "Survey date"],
            "customer_change": ["[MS11]", "Khách hàng đang có thay đổi", "Khach hang dang co thay doi", "Customer change"],
            "demand_increase": ["[MS15]", "những nhu cầu nào đang tăng", "nhung nhu cau nao dang tang", "demand increase"],
            "lost_sale_reasons": ["[MS21]", "nguyên nhân khách không mua", "nguyen nhan khach khong mua", "lost sale reasons"],
            "lost_sale_top1": ["[MS22]", "nguyên nhân lớn nhất", "nguyen nhan quan trong nhat", "lost sale top 1"],
            "product_gap": ["[MS31]", "nhóm sản phẩm cần bổ sung", "nhom san pham can bo sung", "product gap"],
            "acceptable_price": ["[MS37]", "khoảng giá khách chấp nhận", "khoang gia khach chap nhan", "acceptable price"],
            "support_categories": ["[MS41]", "nhóm cần công ty hỗ trợ", "nhom can cong ty ho tro", "support categories"],
            "suggested_solution": ["[MS45]", "giải pháp cửa hàng đề xuất", "giai phap cua hang de xuat", "suggested solution"],
            "support_photos": ["[MS48]", "ảnh minh chứng", "anh minh chung", "support photos", "support photo"],
            "local_opportunity": ["[MS51]", "mùa vụ hoặc cơ hội", "mua vu hoac co hoi", "local opportunity"],
            "need_before_date": ["[MS59]", "thời điểm hàng cần có", "thoi diem hang can co", "need before date"],
            "store_recommendation": ["[MS61]", "kiến nghị ưu tiên hàng đầu", "kien nghi uu tien hang dau", "store recommendation"],
        }
        
        def find_val(keywords: List[str], default: Any = "") -> Any:
            tech_code = keywords[0]
            for k, v in row.items():
                if tech_code in str(k):
                    return v
            for k, v in row.items():
                k_lower = str(k).lower()
                if any(kw.lower() in k_lower for kw in keywords[1:]):
                    return v
            return default
            
        store_code = str(find_val(columns_map["store_code"])).strip()
        if not store_code:
            return None
            
        region = str(find_val(columns_map["region"])).strip()
        qlkd_asm = str(find_val(columns_map["qlkd_asm"])).strip()
        respondent_name = str(find_val(columns_map["respondent_name"])).strip()
        respondent_role = str(find_val(columns_map["respondent_role"])).strip()
        
        disc_raw = find_val(columns_map["discussion_count"], 0)
        try:
            discussion_count = int(float(disc_raw)) if disc_raw else 0
        except ValueError:
            discussion_count = 0
            
        survey_date = str(find_val(columns_map["survey_date"])).strip()
        customer_change = str(find_val(columns_map["customer_change"])).strip()
        
        def parse_multi_select(val: Any) -> List[str]:
            if not val:
                return []
            return [item.strip() for item in str(val).split(",") if item.strip()]
            
        demand_increase = parse_multi_select(find_val(columns_map["demand_increase"]))
        lost_sale_reasons = parse_multi_select(find_val(columns_map["lost_sale_reasons"]))
        lost_sale_top1 = str(find_val(columns_map["lost_sale_top1"])).strip()
        product_gap = parse_multi_select(find_val(columns_map["product_gap"]))
        acceptable_price = str(find_val(columns_map["acceptable_price"])).strip()
        support_categories = parse_multi_select(find_val(columns_map["support_categories"]))
        suggested_solution = str(find_val(columns_map["suggested_solution"])).strip()
        
        photos = []
        photo_urls_raw = str(find_val(columns_map["support_photos"])).strip()
        if photo_urls_raw:
            idx = 1
            for url in photo_urls_raw.replace("\n", ",").split(","):
                url = url.strip()
                if url.startswith("http") or "/" in url or "id=" in url:
                    photos.append(SurveyPhoto(
                        index=idx,
                        drive_url=url
                    ))
                    idx += 1
                    if idx > 2:
                        break
                        
        local_opportunity = str(find_val(columns_map["local_opportunity"])).strip()
        need_before_date = str(find_val(columns_map["need_before_date"])).strip()
        store_recommendation = str(find_val(columns_map["store_recommendation"])).strip()
        
        status = str(row.get("Status", "new")).strip()
        qc_status = str(row.get("QC_Status", "pending")).strip()
        
        return MarketSurveyResponse(
            response_id=row_id,
            store_code=store_code,
            region=region,
            qlkd_asm=qlkd_asm,
            respondent_name=respondent_name,
            respondent_role=respondent_role,
            discussion_count=discussion_count,
            survey_date=survey_date,
            customer_change=customer_change,
            demand_increase=demand_increase,
            lost_sale_reasons=lost_sale_reasons,
            lost_sale_top1=lost_sale_top1,
            product_gap=product_gap,
            acceptable_price=acceptable_price,
            support_categories=support_categories,
            suggested_solution=suggested_solution,
            photos=photos,
            local_opportunity=local_opportunity,
            need_before_date=need_before_date,
            store_recommendation=store_recommendation,
            status=status,
            qc_status=qc_status
        )
