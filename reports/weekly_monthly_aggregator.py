import os
import json
import re
from datetime import datetime, timedelta
import pandas as pd
from typing import Dict, List, Any, Tuple, Union, Optional
from data.models import UnifiedInspectionRecord, CommonInspectionEnvelope, DiagnosticCardModel

class WeeklyMonthlyAggregator:
    """
    Wave 6 Executive Semantic Aggregation Engine.
    Computes Canonical Metrics for CEO Board Pack (Excel & PPTX):
      - Network Pacing, Gap Amount, Revenue at Risk
      - 5-Mode Inspection Volume & Compliance Breakdown
      - 4 Canonical Action Effectiveness Metrics
      - Quick Pulse 6-Toggle Operational Health Matrix
      - Deep Audit Top 5 Systemic Failures
    """
    def __init__(self, loader=None):
        self.loader = loader
        self.dim_store = loader.load_dim_store() if loader else {}

    @staticmethod
    def remove_accents(str_val: str) -> str:
        if not str_val:
            return ""
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

    def aggregate_unified_records(
        self,
        records: List[UnifiedInspectionRecord],
        period_name: str = "Tháng 08/2026",
        asm_filter: str = "ALL",
        region_filter: str = "ALL"
    ) -> Dict[str, Any]:
        """
        Aggregates UnifiedInspectionRecords into an Executive Data Structure for Excel & PPTX Generators.
        """
        filtered_records = []
        for r in records:
            if asm_filter != "ALL" and self.remove_accents(r.envelope.asm_name) != self.remove_accents(asm_filter):
                continue
            if region_filter != "ALL" and r.diagnostic and r.diagnostic.region.upper() != region_filter.upper():
                continue
            filtered_records.append(r)
            
        total_visits = len(filtered_records)
        
        # 1. Mode Volume Breakdown
        mode_counts = {
            "quick_pulse": sum(1 for r in filtered_records if r.envelope.inspection_mode == "quick_pulse"),
            "target_rescue": sum(1 for r in filtered_records if r.envelope.inspection_mode == "target_rescue"),
            "deep_audit": sum(1 for r in filtered_records if r.envelope.inspection_mode in ["deep_audit", "own"]),
            "cross_inspection": sum(1 for r in filtered_records if r.envelope.inspection_mode in ["cross_inspection", "cross"]),
            "opening_inspection": sum(1 for r in filtered_records if r.envelope.inspection_mode in ["opening_inspection", "opening"])
        }
        
        # 2. Diagnostic & Network Pacing Metrics
        distinct_stores = {}
        for r in filtered_records:
            code = r.envelope.store_code
            if r.diagnostic and code not in distinct_stores:
                distinct_stores[code] = r.diagnostic
                
        total_actual = sum(d.mtd_actual for d in distinct_stores.values())
        total_target = sum(d.mtd_target for d in distinct_stores.values())
        attainment_pct = (total_actual / total_target * 100) if total_target > 0 else 0.0
        gap_total = max(0.0, total_target - total_actual)
        
        # Severity tiers count
        sev_counts = {
            "PROTECT_ON_TRACK": sum(1 for d in distinct_stores.values() if d.lag_severity == "PROTECT_ON_TRACK"),
            "WATCH": sum(1 for d in distinct_stores.values() if d.lag_severity == "WATCH"),
            "RECOVERY": sum(1 for d in distinct_stores.values() if d.lag_severity == "RECOVERY"),
            "RESCUE_CRITICAL": sum(1 for d in distinct_stores.values() if d.lag_severity == "RESCUE_CRITICAL"),
            "UNKNOWN": sum(1 for d in distinct_stores.values() if d.lag_severity == "UNKNOWN")
        }
        
        # 3. Action Lifecycle KPIs (Canonical Denominators DA-07 / INV-05)
        rescue_records = [r for r in filtered_records if r.target_rescue or r.rescue_intervention]
        total_committed = len(rescue_records)
        total_completed = sum(1 for r in rescue_records if (r.target_rescue and r.target_rescue.intervention_status in ["COMPLETED", "VERIFIED", "EFFECTIVE"]) or (r.rescue_intervention and r.rescue_intervention.intervention_status in ["COMPLETED", "VERIFIED", "EFFECTIVE"]))
        total_verified = sum(1 for r in rescue_records if (r.target_rescue and r.target_rescue.intervention_status in ["VERIFIED", "EFFECTIVE"]) or (r.rescue_intervention and r.rescue_intervention.intervention_status in ["VERIFIED", "EFFECTIVE"]))
        
        total_effective = 0
        total_expected_recovery = 0.0
        total_actual_recovery = 0.0
        
        for r in rescue_records:
            tr = r.target_rescue or r.rescue_intervention
            if tr:
                if tr.expected_recovery:
                    total_expected_recovery += tr.expected_recovery
                if tr.actual_result:
                    total_actual_recovery += tr.actual_result
                if tr.effectiveness_verdict == "EFFECTIVE":
                    total_effective += 1
                elif tr.actual_result and tr.expected_recovery and tr.actual_result >= tr.expected_recovery and tr.intervention_status in ["VERIFIED", "EFFECTIVE"]:
                    total_effective += 1

        action_completion_rate = (total_completed / total_committed * 100) if total_committed > 0 else 100.0
        action_verification_rate = (total_verified / total_committed * 100) if total_committed > 0 else 100.0
        recovery_effectiveness_rate = (total_effective / total_verified * 100) if total_verified > 0 else 100.0
        effective_action_rate = (total_effective / total_committed * 100) if total_committed > 0 else 100.0
        
        # 4. Quick Pulse Health Matrix
        pulse_records = [r for r in filtered_records if r.quick_pulse]
        pulse_count = max(1, len(pulse_records))
        pulse_stats = {
            "staff_on_duty_pct": round(sum(1 for r in pulse_records if r.quick_pulse.staff_on_duty) / pulse_count * 100, 1),
            "uniform_grooming_pct": round(sum(1 for r in pulse_records if r.quick_pulse.uniform_grooming) / pulse_count * 100, 1),
            "customer_present_pct": round(sum(1 for r in pulse_records if r.quick_pulse.customer_present) / pulse_count * 100, 1),
            "cleanliness_lighting_pct": round(sum(1 for r in pulse_records if r.quick_pulse.cleanliness_lighting) / pulse_count * 100, 1),
            "hot_skus_available_pct": round(sum(1 for r in pulse_records if r.quick_pulse.hot_skus_available) / pulse_count * 100, 1),
            "pos_system_ok_pct": round(sum(1 for r in pulse_records if r.quick_pulse.pos_system_ok) / pulse_count * 100, 1)
        }
        
        # 5. Top Systemic Failures (Deep Audit)
        audit_records = [r for r in filtered_records if r.deep_audit]
        issue_categories = {}
        for r in audit_records:
            da = r.deep_audit
            if da.rating_frontage in ["Chưa đạt", "Không đạt"]:
                issue_categories["Mặt tiền & Biển hiệu"] = issue_categories.get("Mặt tiền & Biển hiệu", 0) + 1
            if da.rating_inner in ["Chưa đạt", "Không đạt"]:
                issue_categories["Không gian bên trong"] = issue_categories.get("Không gian bên trong", 0) + 1
            if da.rating_merch in ["Chưa đạt", "Không đạt"]:
                issue_categories["Trưng bày hàng hóa"] = issue_categories.get("Trưng bày hàng hóa", 0) + 1
            if da.rating_staff in ["Chưa đạt", "Không đạt"]:
                issue_categories["Tác phong nhân sự"] = issue_categories.get("Tác phong nhân sự", 0) + 1
            if da.rating_csvc in ["Chưa đạt", "Không đạt"]:
                issue_categories["Cơ sở vật chất & PCCC"] = issue_categories.get("Cơ sở vật chất & PCCC", 0) + 1
                
        sorted_issues = sorted(issue_categories.items(), key=lambda x: x[1], reverse=True)
        top_systemic_issues = [{"category": k, "count": v} for k, v in sorted_issues[:5]]
        
        # 6. Detailed Store List for Excel Tables
        store_rows = []
        for r in filtered_records:
            code = r.envelope.store_code
            diag = r.diagnostic
            tr = r.target_rescue or r.rescue_intervention
            store_rows.append({
                "visit_id": r.envelope.visit_id,
                "store_code": code,
                "store_name": diag.store_name if diag else code,
                "region": diag.region if diag else "HCM",
                "asm_name": r.envelope.asm_name,
                "report_date": r.envelope.report_date,
                "mode": r.envelope.inspection_mode,
                "data_class": r.envelope.data_class,
                "mtd_actual": diag.mtd_actual if diag else 0.0,
                "mtd_target": diag.mtd_target if diag else 0.0,
                "achievement_pct": diag.achievement_pct if diag else 0.0,
                "pace_delta_pct": diag.pace_delta_pct if diag else 0.0,
                "lag_severity": tr.lag_severity if tr else (diag.lag_severity if diag else "UNKNOWN"),
                "primary_blocker": tr.primary_blocker if tr else (diag.primary_blocker.title if diag and diag.primary_blocker else ""),
                "action_plan": tr.action_plan if tr else (r.deep_audit.action_plan if r.deep_audit else ""),
                "action_owner": tr.action_owner if tr else r.envelope.asm_name,
                "action_due_date": tr.action_due_date if tr else (r.deep_audit.action_deadline if r.deep_audit else ""),
                "expected_recovery": tr.expected_recovery if tr else None,
                "actual_result": tr.actual_result if tr else None,
                "intervention_status": tr.intervention_status if tr else "N/A",
                "effectiveness_verdict": tr.effectiveness_verdict if tr else "N/A"
            })
            
        return {
            "period_name": period_name,
            "asm_filter": asm_filter,
            "region_filter": region_filter,
            "generated_at": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "kpis": {
                "total_visited": total_visits,
                "unique_stores_count": len(distinct_stores),
                "mode_counts": mode_counts,
                "network_revenue_actual": total_actual,
                "network_revenue_target": total_target,
                "network_attainment_pct": round(attainment_pct, 1),
                "network_gap_total": gap_total,
                "severity_counts": sev_counts,
                "total_committed_actions": total_committed,
                "total_completed_actions": total_completed,
                "total_verified_actions": total_verified,
                "total_effective_actions": total_effective,
                "action_completion_rate_pct": round(action_completion_rate, 1),
                "action_verification_rate_pct": round(action_verification_rate, 1),
                "recovery_effectiveness_rate_pct": round(recovery_effectiveness_rate, 1),
                "effective_action_rate_pct": round(effective_action_rate, 1),
                "total_expected_recovery": total_expected_recovery,
                "total_actual_recovery": total_actual_recovery,
                "pulse_stats": pulse_stats,
                "top_systemic_issues": top_systemic_issues
            },
            "store_rows": store_rows
        }
