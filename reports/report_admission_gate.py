import os
import json
import hashlib
import datetime
from typing import List, Dict, Any, Optional
from pydantic import BaseModel, Field
from data.models import UnifiedInspectionRecord, ReconciliationIncidentModel, ActionLifecycleEnum

class AdmissionVerdict(BaseModel):
    admissible: bool
    report_run_id: str
    timestamp: str
    source_snapshot_id: str = "SNAPSHOT_2026_08_28"
    source_snapshot_date: str = "28/08/2026"
    engine_version: str = "v6.0-PRO-CERTIFIED"
    schema_version: str = "2026.08.31-v1.0"
    evidence_class: str # REAL_FIELD / CONTROLLED_PILOT_BASELINE
    total_inspections: int
    total_rescue_interventions: int
    unresolved_incidents_count: int
    duplicate_count: int
    ghost_records_count: int
    orphan_rescue_count: int
    block_reasons: List[str] = []
    audit_hash_full: str = ""      # Full 64-char SHA-256 Hex Hash
    audit_hash_display: str = ""   # 16-char Display Digest with ellipsis

class ReportAdmissionGate:
    """
    Wave 6 Pre-flight Fail-Closed Admission Gate (Gate G5 / DA-X06 / P18.1 / P18.2).
    Enforces:
      1. Zero Unresolved Incidents in Reconciliation_Alerts.
      2. Zero duplicate visit_id.
      3. Zero ghost and orphan records.
      4. Valid state machine transition.
      5. Strict DATA_CLASS provenance validation.
      6. Canonical 64-char SHA-256 evidence hashing.
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

        # 4. Action State Machine Validation (DA-07 / G0 / P18.4)
        for rec in records:
            if rec.target_rescue:
                status = rec.target_rescue.intervention_status
                verdict = rec.target_rescue.effectiveness_verdict
                if verdict == "EFFECTIVE" and status not in ["VERIFIED", "EFFECTIVE"]:
                    block_reasons.append(f"ILLEGAL STATE TRANSITION: Visit {rec.envelope.visit_id} claimed EFFECTIVE without VERIFIED status!")

        # 5. Evidence Classification & Provenance Validation (P18.1 / INV-01)
        real_field_count = sum(1 for r in records if r.envelope.data_class == "REAL_FIELD")
        baseline_count = sum(1 for r in records if r.envelope.data_class != "REAL_FIELD")
        
        if real_field_count > 0 and baseline_count == 0:
            evidence_class = "REAL_FIELD"
        elif real_field_count == 0 and baseline_count > 0:
            evidence_class = "CONTROLLED_PILOT_BASELINE"
        else:
            evidence_class = "CONTAMINATED_MIXED_DATASET"
            block_reasons.append("BASELINE CONTAMINATION DETECTED: Dataset contains mixed REAL_FIELD and CONTROLLED_PILOT_BASELINE records!")
            
        if enforce_real_field and evidence_class != "REAL_FIELD":
            block_reasons.append(f"ENFORCEMENT FAILED: Report requires REAL_FIELD data but got {evidence_class}")

        admissible = (len(block_reasons) == 0)
        
        # 6. Canonical Full SHA-256 Audit Hash (P18.2 / 64 Hex Characters)
        canonical_payload = {
            "run_id": run_id,
            "timestamp": timestamp,
            "evidence_class": evidence_class,
            "inspections_count": len(records),
            "incidents_count": len(incidents),
            "unresolved_count": len(unresolved),
            "duplicate_count": len(duplicates),
            "ghost_count": ghost_count,
            "orphan_count": orphan_count,
            "admissible": admissible,
            "visit_ids": sorted(list(visited_ids))
        }
        canonical_str = json.dumps(canonical_payload, sort_keys=True, ensure_ascii=False)
        audit_hash_full = hashlib.sha256(canonical_str.encode("utf-8")).hexdigest().upper()
        audit_hash_display = f"{audit_hash_full[:16]}..."
        
        return AdmissionVerdict(
            admissible=admissible,
            report_run_id=run_id,
            timestamp=timestamp,
            source_snapshot_id="SNAPSHOT_2026_08_28",
            source_snapshot_date="28/08/2026",
            engine_version="v6.0-PRO-CERTIFIED",
            schema_version="2026.08.31-v1.0",
            evidence_class=evidence_class,
            total_inspections=len(records),
            total_rescue_interventions=total_rescue,
            unresolved_incidents_count=len(unresolved),
            duplicate_count=len(duplicates),
            ghost_records_count=ghost_count,
            orphan_rescue_count=orphan_count,
            block_reasons=block_reasons,
            audit_hash_full=audit_hash_full,
            audit_hash_display=audit_hash_display
        )
