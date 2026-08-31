import json
import os
import datetime

print("==========================================================================")
print("RUNNING WAVE 4 FIELD TELEMETRY & DATA RECONCILIATION ENGINE (GATE P16)")
print("==========================================================================")

def reconcile_telemetry(submissions_log, form_responses_rows, rescue_rows):
    """
    Reconciles Submitted Client Intents vs Primary Form Rows vs Secondary Rescue Rows.
    Calculates Ghost Records, Orphan Records, Idempotency Violations & Action Closure Rate.
    """
    sub_count = len(submissions_log)
    form_count = len(form_responses_rows)
    rescue_count = len(rescue_rows)
    
    primary_by_id = {}
    duplicate_persisted = 0
    for r in form_responses_rows:
        vid = r.get("visit_id")
        if vid in primary_by_id:
            duplicate_persisted += 1
        else:
            primary_by_id[vid] = r
            
    rescue_by_id = {}
    for rc in rescue_rows:
        vid = rc.get("visit_id")
        rescue_by_id[vid] = rc
        
    ghost_records = []
    for vid, rc in rescue_by_id.items():
        if vid not in primary_by_id:
            ghost_records.append(vid)
            
    orphan_rescue = []
    for vid, pr in primary_by_id.items():
        if pr.get("visit_type") == "target_rescue" and vid not in rescue_by_id:
            orphan_rescue.append(vid)
            
    total_committed = len(rescue_rows)
    completed_or_verified = sum(1 for rc in rescue_rows if rc.get("status") in ["COMPLETED", "VERIFIED"])
    verified_count = sum(1 for rc in rescue_rows if rc.get("status") == "VERIFIED")
    
    action_closure_rate = (verified_count / total_committed * 100) if total_committed > 0 else 100.0
    action_progress_rate = (completed_or_verified / total_committed * 100) if total_committed > 0 else 100.0
    
    report = {
        "reconciliation_timestamp": datetime.datetime.now().isoformat(),
        "submitted_intents": sub_count,
        "primary_persisted_rows": form_count,
        "secondary_rescue_rows": rescue_count,
        "duplicate_rows_detected": duplicate_persisted,
        "ghost_records_detected": len(ghost_records),
        "orphan_rescue_detected": len(orphan_rescue),
        "unexplained_deltas": abs(sub_count - form_count),
        "total_committed_interventions": total_committed,
        "verified_interventions": verified_count,
        "action_closure_rate_pct": round(action_closure_rate, 1),
        "action_progress_rate_pct": round(action_progress_rate, 1),
        "reconciliation_verdict": "PASS" if (duplicate_persisted == 0 and len(ghost_records) == 0 and len(orphan_rescue) == 0) else "FAIL"
    }
    
    return report

if __name__ == "__main__":
    mock_submissions = [{"visit_id": f"v_{i:03d}"} for i in range(1, 21)]
    mock_primary = [{"visit_id": f"v_{i:03d}", "visit_type": "target_rescue" if i <= 5 else "quick_pulse"} for i in range(1, 21)]
    mock_rescue = [{"visit_id": f"v_{i:03d}", "status": "VERIFIED" if i <= 3 else "IN_PROGRESS"} for i in range(1, 6)]

    rec_result = reconcile_telemetry(mock_submissions, mock_primary, mock_rescue)
    print(json.dumps(rec_result, ensure_ascii=False, indent=2))

    assert rec_result["duplicate_rows_detected"] == 0
    assert rec_result["ghost_records_detected"] == 0
    assert rec_result["orphan_rescue_detected"] == 0
    assert rec_result["reconciliation_verdict"] == "PASS"
    print("\n✓ Gate P16 [Data Reconciliation & Action Closure Rate Engine]: PASS (100% Reconciled)")
