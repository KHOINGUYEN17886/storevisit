from __future__ import annotations
from pydantic import BaseModel, Field
from typing import List, Dict, Optional, Any, Union
from enum import Enum

class InspectionModeEnum(str, Enum):
    QUICK_PULSE = "quick_pulse"
    TARGET_RESCUE = "target_rescue"
    DEEP_AUDIT = "deep_audit"
    CROSS_INSPECTION = "cross_inspection"
    OPENING_INSPECTION = "opening_inspection"
    # Legacy fallbacks
    OWN = "own"
    CROSS = "cross"
    OPENING = "opening"

class ActionLifecycleEnum(str, Enum):
    COMMITTED = "COMMITTED"
    IN_PROGRESS = "IN_PROGRESS"
    COMPLETED = "COMPLETED"
    VERIFIED = "VERIFIED"
    EFFECTIVE = "EFFECTIVE"

class DataClassEnum(str, Enum):
    REAL_FIELD = "REAL_FIELD"
    CONTROLLED_PILOT_BASELINE = "CONTROLLED_PILOT_BASELINE"
    TEST = "TEST"

class StoreMetadata(BaseModel):
    store_code: str
    store_name: str
    address: str = ""
    region: str = "HCM"
    asm_name: str = ""

class RevenueData(BaseModel):
    revenue_actual: int = Field(default=0, description="MTD actual revenue in VND")
    revenue_target: int = Field(default=0, description="Monthly target in VND")
    attainment_pct: float = Field(default=0.0, description="Attainment percentage vs full monthly target")
    revenue_prev: int = Field(default=0, description="Previous month MTD like-for-like revenue in VND")
    revenue_yoy: int = Field(default=0, description="YoY month MTD like-for-like revenue in VND")
    mom_change_pct: float = Field(default=0.0, description="MoM change percentage")
    yoy_change_pct: float = Field(default=0.0, description="YoY change percentage")
    commentary: str = ""
    cutoff_day: Optional[int] = None
    is_mtd: bool = False
    revenue_target_mtd: int = 0
    attainment_mtd_pct: float = 0.0

class StockInventory(BaseModel):
    total_qty: int = 0
    skus_count: int = 0
    qty_nguyen_gia: int = 0
    qty_sale: int = 0
    qty_thanh_ly: int = 0
    age_groups: Dict[str, int] = Field(default_factory=dict)
    commentary: str = ""

class BestSellerItem(BaseModel):
    rank: int
    sku: str
    product_name: str
    brand: str
    sales_4w: int
    stock_qty: int
    image_path: Optional[str] = None

class SlowSellerItem(BaseModel):
    rank: int
    sku: str
    product_name: str
    brand: str
    stock_qty: int
    age_days: int
    image_path: Optional[str] = None

class StaffItem(BaseModel):
    name: str
    role: str
    seniority: float
    skill_rating: str = "Đạt"
    notes: str = ""

class StaffRoster(BaseModel):
    cht_name: str = "Chưa bổ nhiệm"
    chp_name: str = "Chưa bổ nhiệm"
    staff_list: List[StaffItem] = []

class OperationalIssue(BaseModel):
    index: int
    label: str
    issue: str
    date: str
    assignee: str
    status: str
    notes: str

class FormPhoto(BaseModel):
    section: str          # "frontage", "merchandise", "staff", "csvc", "rescue", "pulse"
    index: int            # 1, 2, 3...
    drive_url: str        # Google Drive shared URL or ID
    local_path: str = ""  # Path to local downloaded photo

# ============================================================================
# WAVE 6: POLYMORPHIC INSPECTION DOMAIN MODELS
# ============================================================================

class QuickPulsePayload(BaseModel):
    staff_on_duty: bool = True
    uniform_grooming: bool = True
    customer_present: bool = False
    cleanliness_lighting: bool = True
    hot_skus_available: bool = True
    pos_system_ok: bool = True
    quick_notes: str = ""
    photos: List[FormPhoto] = []

class TargetRescuePayload(BaseModel):
    lag_severity: str = "RESCUE_CRITICAL"
    primary_blocker: str = ""
    action_plan: str = ""
    action_owner: str = ""
    action_due_date: str = ""
    expected_recovery: Optional[float] = None
    intervention_status: str = "COMMITTED" # COMMITTED -> IN_PROGRESS -> COMPLETED -> VERIFIED -> EFFECTIVE
    actual_result: Optional[float] = None
    verified_at: Optional[str] = None
    effectiveness_verdict: str = "PENDING_EVALUATION" # EFFECTIVE / INEFFECTIVE / PENDING_EVALUATION
    effectiveness_evidence_id: Optional[str] = None
    photos: List[FormPhoto] = []

class DeepAuditPayload(BaseModel):
    nv_count: int = 0
    time_start: str = ""
    time_end: str = ""
    rating_frontage: str = "Đạt"
    rating_inner: str = "Đạt"
    rating_merch: str = "Đạt"
    rating_staff: str = "Đạt"
    rating_csvc: str = "Đạt"
    comment_frontage: str = ""
    comment_inner: str = ""
    comment_merch: str = ""
    comment_staff: str = ""
    comment_csvc: str = ""
    pending_issues: str = ""
    action_plan: str = ""
    action_deadline: str = ""
    store_recommendation: Optional[str] = None
    checklist_json: Optional[str] = None
    photos: List[FormPhoto] = []

class CrossInspectionPayload(BaseModel):
    origin_region: str = "HCM"
    inspected_region: str = ""
    home_asm: str = ""
    cross_notes: str = ""
    deep_audit_data: Optional[DeepAuditPayload] = None

class OpeningInspectionPayload(BaseModel):
    opening_type: str = "new"        # "new" | "reopen"
    opening_phase: str = "day"       # "before" | "day" | "after"
    opening_date: str = ""
    opening_readiness: str = "ready" # "ready" | "minor_fix" | "not_ready"
    deep_audit_data: Optional[DeepAuditPayload] = None

class DiagnosticBlocker(BaseModel):
    priority: int = 1
    code: str = "UNKNOWN"
    category: str = "revenue"
    title: str = ""
    detail: str = ""
    severity: str = "MEDIUM"

class DiagnosticCardModel(BaseModel):
    store_code: str
    store_name: str
    region: str = "HCM"
    asm_name: str = ""
    manager: str = ""
    store_type: str = "Standard"
    data_quality_status: str = "AVAILABLE"
    mtd_actual: float = 0.0
    mtd_target: float = 0.0
    achievement_pct: float = 0.0
    pace_index: float = 1.0
    pace_delta_pct: float = 0.0
    gap_amount: float = 0.0
    selling_days_in_month: int = 31
    selling_days_elapsed: int = 28
    selling_days_remaining: int = 3
    expected_progress_pct: float = 90.32
    required_daily_runrate: float = 0.0
    actual_daily_runrate: float = 0.0
    lag_severity: str = "PROTECT_ON_TRACK" # PROTECT_ON_TRACK / WATCH / RECOVERY / RESCUE_CRITICAL / UNKNOWN
    primary_blocker: Optional[DiagnosticBlocker] = None
    secondary_blockers: List[DiagnosticBlocker] = []

class RescueInterventionModel(BaseModel):
    visit_id: str
    store_code: str
    asm_name: str
    report_date: str
    lag_severity: str = "RESCUE_CRITICAL"
    primary_blocker: str = ""
    action_plan: str = ""
    action_owner: str = ""
    action_due_date: str = ""
    expected_recovery: Optional[float] = None
    intervention_status: str = "COMMITTED"
    actual_result: Optional[float] = None
    verified_at: Optional[str] = None
    effectiveness_verdict: str = "PENDING_EVALUATION"
    effectiveness_evidence_id: Optional[str] = None
    submitted_at: str = ""
    payload_json: str = ""

class ReconciliationIncidentModel(BaseModel):
    incident_id: str
    detected_at: str
    visit_id: str
    store_code: str
    asm_name: str
    ghost_row_idx: int
    failure_type: str = "ROLLBACK_DELETION_FAILED"
    owner: str = "SYSTEM_ADMIN"
    status: str = "UNRESOLVED" # UNRESOLVED / IN_INVESTIGATION / RESOLVED
    resolution: str = ""
    resolved_at: Optional[str] = None

# Common Envelope & Unified Semantic Inspection Record
class CommonInspectionEnvelope(BaseModel):
    visit_id: str
    store_code: str
    asm_name: str
    report_date: str
    inspection_mode: str = "deep_audit"
    data_class: str = "REAL_FIELD"
    timestamp: str = ""
    cht_name: str = ""
    status: str = "pending"
    diagnostic_snapshot_id: str = "SNAPSHOT_2026_08_28"

class UnifiedInspectionRecord(BaseModel):
    envelope: CommonInspectionEnvelope
    quick_pulse: Optional[QuickPulsePayload] = None
    target_rescue: Optional[TargetRescuePayload] = None
    deep_audit: Optional[DeepAuditPayload] = None
    cross_inspection: Optional[CrossInspectionPayload] = None
    opening_inspection: Optional[OpeningInspectionPayload] = None
    diagnostic: Optional[DiagnosticCardModel] = None
    rescue_intervention: Optional[RescueInterventionModel] = None

# Backward compatibility alias
class StoreFormResponse(BaseModel):
    response_id: str
    store_code: str
    report_date: str
    asm_name: str
    cht_name: str = ""
    time_start: str = ""
    time_end: str = ""
    nv_count: int = 0
    rating_frontage: str = "Đạt"
    rating_inner: str = "Đạt"
    rating_merch: str = "Đạt"
    rating_staff: str = "Đạt"
    rating_csvc: str = "Đạt"
    comment_frontage: str = ""
    comment_inner: str = ""
    comment_merch: str = ""
    comment_staff: str = ""
    comment_csvc: str = ""
    pending_issues: str = ""
    action_plan: str = ""
    action_deadline: str = ""
    store_recommendation: Optional[str] = None
    photos: List[FormPhoto] = []
    status: str = "pending"
    checklist_json: Optional[str] = None
    inspection_mode: str = "deep_audit"
    data_class: str = "REAL_FIELD"
    visit_id: str = ""
    # Opening specific
    opening_type: Optional[str] = None
    opening_phase: Optional[str] = None
    opening_date: Optional[str] = None
    opening_readiness: Optional[str] = None
    # Rescue specific
    rescue_payload: Optional[TargetRescuePayload] = None
    quick_pulse_payload: Optional[QuickPulsePayload] = None

class StoreReportData(BaseModel):
    metadata: StoreMetadata
    revenue: RevenueData
    stock: StockInventory
    best_sellers: List[BestSellerItem] = []
    slow_sellers: List[SlowSellerItem] = []
    staff: StaffRoster
    issues: List[OperationalIssue] = []
    csvc_comment: str = ""
    frontage_rating: str = "Đạt"
    frontage_issue: str = ""
    frontage_action: str = ""
    form_response: Optional[StoreFormResponse] = None
    unified_inspection: Optional[UnifiedInspectionRecord] = None

class ClusterStorePerformance(BaseModel):
    store_code: str
    store_name: str
    revenue_actual: int
    revenue_target: int
    attainment_pct: float

class ClusterCriticalIssue(BaseModel):
    store_name: str
    label: str
    issue: str
    assignee: str
    priority: str

class ClusterReportData(BaseModel):
    cluster_name: str
    report_date: str
    revenue_actual: int
    revenue_target: int
    attainment_pct: float
    stores_performance: List[ClusterStorePerformance] = []
    critical_issues: List[ClusterCriticalIssue] = []
    individual_reports: List[StoreReportData] = []

SECTION_LABELS = {
    "frontage": "Mặt tiền", "inner": "Không gian trong",
    "merch_ap": "Trưng bày AP", "merch_pie": "Trưng bày PIE",
    "merch_ab": "Trưng bày AB", "merch_anamai": "Trưng bày Anamai",
    "merch_bonjour": "Trưng bày Bonjour", "merch_pk": "Phụ kiện",
    "warehouse": "Kho/Phòng thử", "stockroom": "Kho hàng",
    "fitting_room": "Phòng thử đồ", "toilet": "Nhà vệ sinh",
    "fire_safety": "PCCC & Thoát hiểm", "cashier": "Thu ngân",
    "packaging_security": "Bao bì & An ninh", "staff": "Nhân sự",
    "security_guard": "Bảo vệ",
}

class SurveyPhoto(BaseModel):
    index: int
    drive_url: str
    local_path: str = ""

class MarketSurveyResponse(BaseModel):
    response_id: str
    store_code: str
    region: str
    qlkd_asm: str
    respondent_name: str
    respondent_role: str
    discussion_count: int
    survey_date: str
    customer_change: str
    demand_increase: List[str] = []
    lost_sale_reasons: List[str] = []
    lost_sale_top1: str
    product_gap: List[str] = []
    acceptable_price: str
    support_categories: List[str] = []
    suggested_solution: str
    photos: List[SurveyPhoto] = []
    local_opportunity: str
    need_before_date: str
    store_recommendation: str
    status: str = "new"
    qc_status: str = "pending"
