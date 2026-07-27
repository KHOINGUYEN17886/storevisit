from __future__ import annotations
from pydantic import BaseModel, Field
from typing import List, Dict, Optional

class StoreMetadata(BaseModel):
    store_code: str
    store_name: str
    address: str
    region: str
    asm_name: str

class RevenueData(BaseModel):
    revenue_actual: int = Field(default=0, description="MTD revenue in VND")
    revenue_target: int = Field(default=0, description="Monthly target in VND")
    attainment_pct: float = Field(default=0.0, description="Attainment percentage")
    revenue_prev: int = Field(default=0, description="Previous month MTD revenue in VND")
    revenue_yoy: int = Field(default=0, description="YoY month MTD revenue in VND")
    mom_change_pct: float = Field(default=0.0)
    yoy_change_pct: float = Field(default=0.0)
    commentary: str = ""

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

class SlowSellerItem(BaseModel):
    rank: int
    sku: str
    product_name: str
    brand: str
    stock_qty: int
    age_days: int

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

class FormPhoto(BaseModel):
    section: str          # "frontage", "merchandise", "staff", "csvc"
    index: int            # 1 or 2
    drive_url: str        # Google Drive shared URL or ID
    local_path: str = ""  # Path to local downloaded photo

class StoreFormResponse(BaseModel):
    response_id: str      # Timestamp or unique row identifier
    store_code: str
    report_date: str
    asm_name: str
    cht_name: str
    time_start: str
    time_end: str
    nv_count: int
    # Ratings ("Tốt" / "Đạt" / "Chưa đạt")
    rating_frontage: str
    rating_inner: str
    rating_merch: str
    rating_staff: str
    rating_csvc: str
    # Comments
    comment_frontage: str
    comment_inner: str
    comment_merch: str
    comment_staff: str
    comment_csvc: str
    pending_issues: str
    action_plan: str
    action_deadline: str
    store_recommendation: Optional[str] = None
    # Photos list
    photos: List[FormPhoto] = []
    status: str = "pending"  # "pending" | "done"
    checklist_json: Optional[str] = None



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
    # Section data
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
    # Status
    status: str = "new"  # new | processing | done | error | ignored
    qc_status: str = "pending"  # pending | approved | rejected


