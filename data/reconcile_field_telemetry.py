import json
import os
import datetime

print("==========================================================================")
print("WAVE 5: FIELD TELEMETRY & BUSINESS EFFECTIVENESS RECONCILIATION ENGINE")
print("==========================================================================")

def reconcile_telemetry(submissions_log, form_responses_rows, rescue_rows, incidents_rows=None):
    """
    Reconciles Submitted Client Intents vs Form Responses 1 vs Rescue_Interventions vs Reconciliation_Alerts.
    Computes 8 Operational KPIs + 4 Business Decision Intelligence KPIs (DA-01 to DA-07 Governance).
    """
    incidents_rows = incidents_rows or []
    sub_count = len(submissions_log)
    form_count = len(form_responses_rows)
    rescue_count = len(rescue_rows)
    incident_count = len(incidents_rows)
    
    # 1. Primary Mapping & Dedup Verification
    primary_by_id = {}
    duplicate_persisted = 0
    for r in form_responses_rows:
        vid = r.get("visit_id")
        if vid in primary_by_id:
            duplicate_persisted += 1
        else:
            primary_by_id[vid] = r
            
    # 2. Secondary Mapping
    rescue_by_id = {}
    for rc in rescue_rows:
        vid = rc.get("visit_id")
        rescue_by_id[vid] = rc
        
    # 3. Ghost Record Detection (Secondary row without Primary row)
    ghost_records = [vid for vid in rescue_by_id if vid not in primary_by_id]
            
    # 4. Orphan Rescue Detection (Primary target_rescue without Secondary row)
    orphan_rescue = [vid for vid, pr in primary_by_id.items() if pr.get("visit_type") == "target_rescue" and vid not in rescue_by_id]
            
    # 5. Business Decision Intelligence KPIs (DA-07 Lifecycle: COMMITTED -> IN_PROGRESS -> COMPLETED -> VERIFIED -> EFFECTIVE)
    total_committed = len(rescue_rows)
    total_in_progress = sum(1 for rc in rescue_rows if rc.get("status") in ["IN_PROGRESS", "COMPLETED", "VERIFIED"])
    total_completed = sum(1 for rc in rescue_rows if rc.get("status") in ["COMPLETED", "VERIFIED"])
    total_verified = sum(1 for rc in rescue_rows if rc.get("status") == "VERIFIED")
    
    # Effectiveness Evaluation: Actual Result >= Expected Recovery
    total_effective = 0
    for rc in rescue_rows:
        act = rc.get("actual_result")
        exp = rc.get("expected_recovery")
        verdict = rc.get("effectiveness_verdict")
        if verdict == "EFFECTIVE" or (act is not None and exp is not None and float(act) >= float(exp)):
            total_effective += 1
            
    # 4 Business KPIs
    action_commitment_rate = (total_committed / max(1, sum(1 for pr in primary_by_id.values() if pr.get("lag_severity") in ["RECOVERY", "RESCUE_CRITICAL"]))) * 100
    action_completion_rate = (total_completed / total_committed * 100) if total_committed > 0 else 100.0
    action_verification_rate = (total_verified / total_committed * 100) if total_committed > 0 else 100.0
    recovery_effectiveness_rate = (total_effective / total_verified * 100) if total_verified > 0 else 100.0
    
    # 6. Unresolved Incidents Tracking (DA-04)
    unresolved_incidents = sum(1 for inc in incidents_rows if inc.get("status") == "UNRESOLVED")
    
    report = {
        "reconciliation_timestamp": datetime.datetime.now().isoformat(),
        "operational_kpis": {
            "submitted_intents": sub_count,
            "primary_persisted_rows": form_count,
            "secondary_rescue_rows": rescue_count,
            "duplicate_rows_detected": duplicate_persisted,
            "ghost_records_detected": len(ghost_records),
            "orphan_rescue_detected": len(orphan_rescue),
            "unexplained_deltas": abs(sub_count - form_count),
            "unresolved_incidents": unresolved_incidents,
            "operational_verdict": "PASS" if (duplicate_persisted == 0 and len(ghost_records) == 0 and len(orphan_rescue) == 0 and unresolved_incidents == 0) else "ALERT"
        },
        "business_decision_kpis": {
            "total_committed_interventions": total_committed,
            "total_completed_interventions": total_completed,
            "total_verified_interventions": total_verified,
            "total_effective_interventions": total_effective,
            "action_completion_rate_pct": round(action_completion_rate, 1),
            "action_verification_rate_pct": round(action_verification_rate, 1),
            "recovery_effectiveness_rate_pct": round(recovery_effectiveness_rate, 1)
        }
    }
    
    return report

if __name__ == "__main__":
    # Simulate 20 Pilot Submissions
    mock_submissions = [{"visit_id": f"pilot_{i:03d}"} for i in range(1, 21)]
    mock_primary = [
        {"visit_id": f"pilot_{i:03d}", "visit_type": "target_rescue" if i <= 5 else "quick_pulse", "lag_severity": "RESCUE_CRITICAL" if i <= 5 else "WATCH"}
        for i in range(1, 21)
    ]
    mock_rescue = [
        {
            "visit_id": f"pilot_{i:03d}",
            "status": "VERIFIED" if i <= 3 else ("COMPLETED" if i == 4 else "IN_PROGRESS"),
            "expected_recovery": 50000000,
            "actual_result": 55000000 if i <= 2 else (40000000 if i == 3 else None),
            "effectiveness_verdict": "EFFECTIVE" if i <= 2 else ("INEFFECTIVE" if i == 3 else "PENDING_EVALUATION")
        }
        for i in range(1, 6)
    ]
    mock_incidents = [] # 0 unresolved incidents

    telemetry_result = reconcile_telemetry(mock_submissions, mock_primary, mock_rescue, mock_incidents)
    print(json.dumps(telemetry_result, ensure_ascii=False, indent=2))

    assert telemetry_result["operational_kpis"]["operational_verdict"] == "PASS"
    assert telemetry_result["business_decision_kpis"]["action_completion_rate_pct"] == 80.0
    assert telemetry_result["business_decision_kpis"]["action_verification_rate_pct"] == 60.0
    assert telemetry_result["business_decision_kpis"]["recovery_effectiveness_rate_pct"] == 66.7
    print("\n✓ Wave 5 Field Telemetry & Business Decision KPIs Engine: 100% VERIFIED!")
