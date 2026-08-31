import os
import re
import json
import logging
from typing import List, Dict, Any, Optional
import gspread
from google.oauth2.service_account import Credentials
from data.models import (
    StoreFormResponse, FormPhoto, MarketSurveyResponse, SurveyPhoto,
    UnifiedInspectionRecord, CommonInspectionEnvelope,
    QuickPulsePayload, TargetRescuePayload, DeepAuditPayload,
    CrossInspectionPayload, OpeningInspectionPayload,
    DiagnosticCardModel, DiagnosticBlocker,
    RescueInterventionModel, ReconciliationIncidentModel
)

logger = logging.getLogger(__name__)

class GoogleSheetsReader:
    def __init__(self, credentials_path: str, spreadsheet_id: str, snapshot_path: Optional[str] = None):
        self.credentials_path = credentials_path
        self.spreadsheet_id = spreadsheet_id
        self.snapshot_path = snapshot_path or os.path.join(os.path.dirname(__file__), "store_diagnostics_snapshot.json")
        self.client = None
        self._diagnostics_cache = None
        
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

    def load_diagnostic_snapshot(self) -> Dict[str, DiagnosticCardModel]:
        if self._diagnostics_cache is not None:
            return self._diagnostics_cache
            
        diagnostics = {}
        if os.path.exists(self.snapshot_path):
            try:
                with open(self.snapshot_path, "r", encoding="utf-8") as f:
                    snap_data = json.load(f)
                    stores = snap_data.get("stores", {})
                    for code, s in stores.items():
                        diag_block = s.get("diagnosis", {})
                        prim = diag_block.get("primary_blocker")
                        prim_model = DiagnosticBlocker(**prim) if prim else None
                        sec_models = [DiagnosticBlocker(**b) for b in diag_block.get("secondary_blockers", [])]
                        
                        rev = s.get("revenue", {})
                        card = DiagnosticCardModel(
                            store_code=s.get("store_code", code),
                            store_name=s.get("store_name", code),
                            region=s.get("region", "HCM"),
                            asm_name=s.get("asm_name", ""),
                            manager=s.get("manager", ""),
                            store_type=s.get("store_type", "Standard"),
                            data_quality_status=s.get("data_quality_status", "AVAILABLE"),
                            mtd_actual=float(rev.get("mtd_actual", 0.0) or 0.0),
                            mtd_target=float(rev.get("mtd_target", 0.0) or 0.0),
                            achievement_pct=float(rev.get("achievement_pct", 0.0) or 0.0),
                            pace_index=float(rev.get("pace_index", 1.0) or 1.0),
                            pace_delta_pct=float(rev.get("pace_delta_pct", 0.0) or 0.0),
                            gap_amount=float(rev.get("gap_amount", 0.0) or 0.0),
                            selling_days_in_month=int(rev.get("selling_days_in_month", 31) or 31),
                            selling_days_elapsed=int(rev.get("selling_days_elapsed", 28) or 28),
                            selling_days_remaining=int(rev.get("selling_days_remaining", 3) or 3),
                            expected_progress_pct=float(rev.get("expected_progress_pct", 90.32) or 90.32),
                            required_daily_runrate=float(rev.get("required_daily_runrate", 0.0) or 0.0),
                            actual_daily_runrate=float(rev.get("actual_daily_runrate", 0.0) or 0.0),
                            lag_severity=diag_block.get("lag_severity", "PROTECT_ON_TRACK"),
                            primary_blocker=prim_model,
                            secondary_blockers=sec_models
                        )
                        diagnostics[code] = card
            except Exception as e:
                logger.error(f"Error loading diagnostic snapshot from {self.snapshot_path}: {e}")
                
        self._diagnostics_cache = diagnostics
        return diagnostics

    def get_rescue_interventions(self, sheet_name: str = "Rescue_Interventions") -> Dict[str, RescueInterventionModel]:
        """Reads secondary Rescue_Interventions sheet tab and maps by visit_id."""
        self._authenticate()
        if not self.spreadsheet_id:
            return {}
        try:
            ss = self.client.open_by_key(self.spreadsheet_id)
            try:
                sheet = ss.worksheet(sheet_name)
            except Exception:
                return {}
            records = sheet.get_all_records()
            rescue_map = {}
            for r in records:
                vid = str(r.get("Visit_ID", "")).strip()
                if not vid:
                    continue
                exp_rec = None
                if r.get("ExpectedRecovery"):
                    try:
                        exp_rec = float(str(r.get("ExpectedRecovery")).replace(",", ""))
                    except ValueError:
                        exp_rec = None
                act_res = None
                if r.get("ActualResult"):
                    try:
                        act_res = float(str(r.get("ActualResult")).replace(",", ""))
                    except ValueError:
                        act_res = None
                rescue_map[vid] = RescueInterventionModel(
                    visit_id=vid,
                    store_code=str(r.get("StoreCode", "")).strip(),
                    asm_name=str(r.get("ASM", "")).strip(),
                    report_date=str(r.get("ReportDate", "")).strip(),
                    lag_severity=str(r.get("LagSeverity", "RESCUE_CRITICAL")).strip(),
                    primary_blocker=str(r.get("PrimaryBlocker", "")).strip(),
                    action_plan=str(r.get("ActionPlan", "")).strip(),
                    action_owner=str(r.get("ActionOwner", "")).strip(),
                    action_due_date=str(r.get("ActionDueDate", "")).strip(),
                    expected_recovery=exp_rec,
                    intervention_status=str(r.get("InterventionStatus", "COMMITTED")).strip(),
                    actual_result=act_res,
                    verified_at=str(r.get("VerifiedAt", "")).strip() or None,
                    effectiveness_verdict=str(r.get("EffectivenessVerdict", "PENDING_EVALUATION")).strip(),
                    effectiveness_evidence_id=str(r.get("EffectivenessEvidenceID", "")).strip() or None,
                    submitted_at=str(r.get("SubmittedAt", "")).strip(),
                    payload_json=str(r.get("Payload_JSON", "")).strip()
                )
            return rescue_map
        except Exception as e:
            logger.warning(f"Could not load Rescue_Interventions: {e}")
            return {}

    def get_reconciliation_incidents(self, sheet_name: str = "Reconciliation_Alerts") -> List[ReconciliationIncidentModel]:
        """Reads Reconciliation_Alerts incident sheet."""
        self._authenticate()
        if not self.spreadsheet_id:
            return []
        try:
            ss = self.client.open_by_key(self.spreadsheet_id)
            try:
                sheet = ss.worksheet(sheet_name)
            except Exception:
                return []
            records = sheet.get_all_records()
            incidents = []
            for r in records:
                inc_id = str(r.get("Incident_ID", "")).strip()
                if not inc_id:
                    continue
                incidents.append(ReconciliationIncidentModel(
                    incident_id=inc_id,
                    detected_at=str(r.get("Timestamp", "") or r.get("Detected_At", "")).strip(),
                    visit_id=str(r.get("Visit_ID", "")).strip(),
                    store_code=str(r.get("StoreCode", "")).strip(),
                    asm_name=str(r.get("ASM", "")).strip(),
                    ghost_row_idx=int(r.get("GhostRowIdx", 0) or 0),
                    failure_type=str(r.get("FailureType", "ROLLBACK_DELETION_FAILED")).strip(),
                    owner=str(r.get("Owner", "SYSTEM_ADMIN")).strip(),
                    status=str(r.get("Status", "UNRESOLVED")).strip(),
                    resolution=str(r.get("Resolution", "")).strip(),
                    resolved_at=str(r.get("Resolved_At", "")).strip() or None
                ))
            return incidents
        except Exception as e:
            logger.warning(f"Could not load Reconciliation_Alerts: {e}")
            return []

    def get_unified_inspection_records(self, sheet_name: str = "Form Responses 1") -> List[UnifiedInspectionRecord]:
        """
        Wave 6 Unified Ingestion:
        Reads Form Responses 1, merges with Rescue_Interventions by visit_id,
        and binds DiagnosticCardModel from Data Lake snapshot by store_code.
        """
        self._authenticate()
        diag_map = self.load_diagnostic_snapshot()
        rescue_map = self.get_rescue_interventions()
        
        try:
            sheet = self.client.open_by_key(self.spreadsheet_id).worksheet(sheet_name)
            records = sheet.get_all_records()
        except Exception as e:
            logger.error(f"Error opening sheet {sheet_name}: {e}")
            raise e
            
        unified_records = []
        for i, row in enumerate(records):
            row_idx = i + 2
            resp = self._parse_row(row, str(row_idx))
            if not resp:
                continue
                
            code = resp.store_code
            diag_card = diag_map.get(code)
            vid = resp.visit_id or resp.response_id or f"visit_{row_idx}"
            rescue_item = rescue_map.get(vid)
            
            envelope = CommonInspectionEnvelope(
                visit_id=vid,
                store_code=code,
                asm_name=resp.asm_name,
                report_date=resp.report_date,
                inspection_mode=resp.inspection_mode,
                data_class=resp.data_class,
                timestamp=resp.time_start,
                cht_name=resp.cht_name,
                status=resp.status,
                diagnostic_snapshot_id="SNAPSHOT_2026_08_28"
            )
            
            # Build mode payload
            qp_payload = None
            tr_payload = None
            da_payload = None
            ci_payload = None
            oi_payload = None
            
            mode = resp.inspection_mode
            if mode == "quick_pulse":
                qp_payload = resp.quick_pulse_payload or QuickPulsePayload(photos=resp.photos)
            elif mode == "target_rescue":
                tr_payload = resp.rescue_payload or TargetRescuePayload(
                    lag_severity=rescue_item.lag_severity if rescue_item else (diag_card.lag_severity if diag_card else "RESCUE_CRITICAL"),
                    primary_blocker=rescue_item.primary_blocker if rescue_item else (diag_card.primary_blocker.title if diag_card and diag_card.primary_blocker else ""),
                    action_plan=rescue_item.action_plan if rescue_item else resp.action_plan,
                    action_owner=rescue_item.action_owner if rescue_item else resp.asm_name,
                    action_due_date=rescue_item.action_due_date if rescue_item else resp.action_deadline,
                    expected_recovery=rescue_item.expected_recovery if rescue_item else None,
                    intervention_status=rescue_item.intervention_status if rescue_item else "COMMITTED",
                    actual_result=rescue_item.actual_result if rescue_item else None,
                    verified_at=rescue_item.verified_at if rescue_item else None,
                    effectiveness_verdict=rescue_item.effectiveness_verdict if rescue_item else "PENDING_EVALUATION",
                    effectiveness_evidence_id=rescue_item.effectiveness_evidence_id if rescue_item else None,
                    photos=resp.photos
                )
            elif mode == "opening_inspection":
                oi_payload = OpeningInspectionPayload(
                    opening_type=resp.opening_type or "new",
                    opening_phase=resp.opening_phase or "day",
                    opening_date=resp.opening_date or resp.report_date,
                    opening_readiness=resp.opening_readiness or "ready"
                )
            elif mode == "cross_inspection":
                ci_payload = CrossInspectionPayload(
                    home_asm=resp.asm_name,
                    cross_notes=resp.pending_issues
                )
            else: # deep_audit / standard
                da_payload = DeepAuditPayload(
                    nv_count=resp.nv_count,
                    time_start=resp.time_start,
                    time_end=resp.time_end,
                    rating_frontage=resp.rating_frontage,
                    rating_inner=resp.rating_inner,
                    rating_merch=resp.rating_merch,
                    rating_staff=resp.rating_staff,
                    rating_csvc=resp.rating_csvc,
                    comment_frontage=resp.comment_frontage,
                    comment_inner=resp.comment_inner,
                    comment_merch=resp.comment_merch,
                    comment_staff=resp.comment_staff,
                    comment_csvc=resp.comment_csvc,
                    pending_issues=resp.pending_issues,
                    action_plan=resp.action_plan,
                    action_deadline=resp.action_deadline,
                    store_recommendation=resp.store_recommendation,
                    checklist_json=resp.checklist_json,
                    photos=resp.photos
                )
                
            unified = UnifiedInspectionRecord(
                envelope=envelope,
                quick_pulse=qp_payload,
                target_rescue=tr_payload,
                deep_audit=da_payload,
                cross_inspection=ci_payload,
                opening_inspection=oi_payload,
                diagnostic=diag_card,
                rescue_intervention=rescue_item
            )
            unified_records.append(unified)
            
        return unified_records
            
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
        def find_val(keywords: List[str], default: Any = "") -> Any:
            for k, v in row.items():
                k_lower = k.lower()
                if any(kw.lower() in k_lower for kw in keywords):
                    return v
            return default

        store_code = str(find_val(["Mã cửa hàng", "Ma cua hang", "Store Code", "storeCode", "StoreCode"])).strip()
        if not store_code:
            return None
            
        if " - " in store_code:
            store_code = store_code.split(" - ")[0].strip()
            
        report_date = str(find_val(["Ngày kiểm tra", "Ngay kiem tra", "Date", "reportDate", "ReportDate"])).strip()
        asm_name = str(find_val(["QLKD/ASM", "ASM", "Người kiểm tra", "Nguoi kiem tra", "asmName", "asm"])).strip()
        cht_name = str(find_val(["Tên CHT", "Ten CHT", "Cửa hàng trưởng", "chtName"])).strip()
        time_start = str(find_val(["Giờ bắt đầu", "Gio bat dau", "Time start", "timestamp", "Timestamp"])).strip()
        time_end = str(find_val(["Giờ kết thúc", "Gio ket thuc", "Time end"])).strip()
        
        # Inspection Mode Detection
        mode_val = str(find_val(["Loại kiểm tra", "Phân loại", "inspection_mode", "inspectionMode", "visit_type", "VisitType", "Mode"], "deep_audit")).strip().lower()
        if "pulse" in mode_val or "nhanh" in mode_val:
            canonical_mode = "quick_pulse"
        elif "rescue" in mode_val or "cứu" in mode_val or "cuu" in mode_val or "target" in mode_val:
            canonical_mode = "target_rescue"
        elif "cross" in mode_val or "chéo" in mode_val or "cheo" in mode_val:
            canonical_mode = "cross_inspection"
        elif "opening" in mode_val or "khai trương" in mode_val or "khai truong" in mode_val:
            canonical_mode = "opening_inspection"
        else:
            canonical_mode = "deep_audit"
            
        # Data class & visit_id
        visit_id = str(find_val(["Visit_ID", "visit_id", "visitId", "ID", "Mã lượt ghé", "SubmissionID"], "")).strip() or f"visit_{row_id}"
        data_class = str(find_val(["DATA_CLASS", "data_class", "DataClass", "EvidenceClass"], "REAL_FIELD")).strip()
        
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
        
        pending_issues = str(find_val(["Vấn đề tồn đọng", "Van de ton dong", "Issues", "pending_issues"])).strip()
        action_plan = str(find_val(["Kế hoạch khắc phục", "Ke hoach khac phuc", "Action plan", "action_plan", "ActionPlan"])).strip()
        action_deadline = str(find_val(["Thời hạn xử lý", "Thoi han xu ly", "Deadline", "due_date", "action_deadline"])).strip()
        store_recommendation = str(find_val(["Đề xuất phát triển", "store_recommendation", "storeRecommendation"], "")).strip()
        
        # Photos parser
        photos = []
        photo_fields = [
            ("frontage", ["Ảnh mặt tiền", "Anh mat tien", "Frontage photo", "photo_frontage"]),
            ("merchandise", ["Ảnh hàng hóa", "Anh hang hoa", "Merch photo", "photo_merch"]),
            ("staff", ["Ảnh nhân sự", "Anh nhan su", "Staff photo", "photo_staff"]),
            ("csvc", ["Ảnh CSVC", "Anh CSVC", "CSVC photo", "photo_csvc"]),
            ("rescue", ["Ảnh cứu target", "photo_rescue", "rescue_photo"]),
            ("pulse", ["Ảnh kiểm tra nhanh", "photo_pulse", "pulse_photo"])
        ]
        
        for section, keywords in photo_fields:
            val = str(find_val(keywords)).strip()
            if val:
                idx = 1
                for url in val.replace("\r", "\n").split("\n"):
                    for u in url.split(","):
                        u = u.strip()
                        if u.startswith("http") or "/" in u or "id=" in u:
                            photos.append(FormPhoto(section=section, index=idx, drive_url=u))
                            idx += 1
                            if idx > 2:
                                break
                            
        checklist_json = str(find_val(["Checklist_JSON", "checklist_json", "ChecklistData", "payload_json", "Payload_JSON"], "")).strip() or None
        
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
            status=str(row.get("Status", "pending")).strip().lower(),
            checklist_json=checklist_json,
            inspection_mode=canonical_mode,
            data_class=data_class,
            visit_id=visit_id
        )

    def get_market_surveys(self, sheet_name: str = "MarketSurvey_Responses") -> List[MarketSurveyResponse]:
        self._authenticate()
        if not self.spreadsheet_id:
            return []
        try:
            sheet = self.client.open_by_key(self.spreadsheet_id).worksheet(sheet_name)
            records = sheet.get_all_records()
        except Exception as e:
            logger.warning(f"MarketSurvey sheet not available: {e}")
            return []
            
        surveys = []
        for i, row in enumerate(records):
            row_idx = i + 2
            survey = self._parse_survey_row(row, str(row_idx))
            if survey:
                surveys.append(survey)
        return surveys

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
            for url in photo_urls_raw.replace("\r", "\n").split("\n"):
                for u in url.split(","):
                    u = u.strip()
                    if u.startswith("http") or "/" in u or "id=" in u:
                        photos.append(SurveyPhoto(
                            index=idx,
                            drive_url=u
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
