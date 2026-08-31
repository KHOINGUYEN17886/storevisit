import os
import hashlib
import datetime
from typing import List, Dict, Any, Optional
from pydantic import BaseModel, Field
from data.models import UnifiedInspectionRecord, ReconciliationIncidentModel, ActionLifecycleEnum

class AdmissionVerdict(BaseModel):
    admissible: bool
    report_run_id: str
    timestamp: str
    evidence_class: str # REAL_FIELD / CONTROLLED_PILOT_BASELINE
    total_inspections: int
    total_rescue_interventions: int
    unresolved_incidents_count: int
    duplicate_count: int
    ghost_records_count: int
    orphan_rescue_count: int
    block_reasons: List[str] = []
    audit_hash: str = ""

class ReportAdmissionGate:
    """
    Wave 6 Pre-flight Fail-Closed Admission Gate (Gate G5 / DA-X06).
    Enforces that NO report (Excel or PPTX) can be generated if:
      - Any UNRESOLVED incident exists in Reconciliation_Alerts.
      - Any duplicate visit_id exists.
      - Any ghost or orphan rescue record exists.
      - Any invalid state machine transition exists.
    """
    @staticmethod
    def evaluate(
        records: List[UnifiedInspectionRecord],
        incidents: Optional[List[ReconciliationIncidentModel]] = None,
        enforce_real_field: bool = False
    ) -> AdmissionVerdict:
        incidents = incidents or []
        timestamp = datetime.datetime.now().isoformat()
        run_id = "RUN_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        block_reasons = []
        
        # 1. Unresolved Incident Check (FAIL-CLOSED)
        unresolved = [inc for inc in incidents if inc.status == "UNRESOLVED"]
        if len(unresolved) > 0:
            block_reasons.append(f"CRITICAL: Found {len(unresolved)} UNRESOLVED incident(s) in Reconciliation_Alerts (e.g., {unresolved[0].incident_id})")
            
        # 2. Duplicate visit_id Check
        visited_ids = set()
        duplicates = []
        for rec in records:
            vid = rec.envelope.visit_id
            if vid in visited_ids:
                duplicates.append(vid)
            else:
                visited_ids.add(vid)
                
        if len(duplicates) > 0:
            block_reasons.append(f"DATA INTEGRITY ERROR: Found {len(duplicates)} duplicate visit_id(s): {duplicates[:3]}")
            
        # 3. Ghost and Orphan Record Checks
        ghost_count = 0
        orphan_count = 0
        total_rescue = 0
        
        for rec in records:
            if rec.envelope.inspection_mode == "target_rescue":
                total_rescue += 1
                if not rec.target_rescue:
                    orphan_count += 1
                    block_reasons.append(f"ORPHAN RESCUE: Visit {rec.envelope.visit_id} marked as target_rescue but missing TargetRescuePayload")
            elif rec.rescue_intervention and rec.envelope.inspection_mode != "target_rescue":
                ghost_count += 1
                block_reasons.append(f"GHOST RESCUE: Visit {rec.envelope.visit_id} has RescueIntervention but inspection_mode is {rec.envelope.inspection_mode}")

        # 4. Action State Machine Validation (DA-07 / G0)
        for rec in records:
            if rec.target_rescue:
                status = rec.target_rescue.intervention_status
                verdict = rec.target_rescue.effectiveness_verdict
                if verdict == "EFFECTIVE" and status not in ["VERIFIED", "EFFECTIVE"]:
                    block_reasons.append(f"ILLEGAL STATE TRANSITION: Visit {rec.envelope.visit_id} claimed EFFECTIVE without VERIFIED status!")

        # 5. Evidence Classification & Isolation (INV-01)
        real_field_count = sum(1 for r in records if r.envelope.data_class == "REAL_FIELD")
        baseline_count = sum(1 for r in records if r.envelope.data_class != "REAL_FIELD")
        
        evidence_class = "REAL_FIELD" if (real_field_count > 0 and baseline_count == 0) else ("CONTROLLED_PILOT_BASELINE" if real_field_count == 0 else "CONTAMINATED_MIXED_DATASET")
        
        if evidence_class == "CONTAMINATED_MIXED_DATASET":
            block_reasons.append("BASELINE CONTAMINATION DETECTED: Dataset contains mixed REAL_FIELD and CONTROLLED_PILOT_BASELINE records!")
            
        if enforce_real_field and evidence_class != "REAL_FIELD":
            block_reasons.append(f"ENFORCEMENT FAILED: Report requires REAL_FIELD data but got {evidence_class}")

        admissible = (len(block_reasons) == 0)
        
        # Calculate Deterministic Audit Hash
        hash_seed = f"{run_id}|{evidence_class}|{len(records)}|{len(incidents)}|{admissible}"
        audit_hash = hashlib.sha256(hash_seed.encode("utf-8")).hexdigest()[:16].upper()
        
        return AdmissionVerdict(
            admissible=admissible,
            report_run_id=run_id,
            timestamp=timestamp,
            evidence_class=evidence_class,
            total_inspections=len(records),
            total_rescue_interventions=total_rescue,
            unresolved_incidents_count=len(unresolved),
            duplicate_count=len(duplicates),
            ghost_records_count=ghost_count,
            orphan_rescue_count=orphan_count,
            block_reasons=block_reasons,
            audit_hash=audit_hash
        )
