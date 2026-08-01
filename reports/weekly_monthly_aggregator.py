import os
import json
from datetime import datetime, timedelta
import pandas as pd
from typing import Dict, List, Any, Tuple

class WeeklyMonthlyAggregator:
    """
    Data Aggregation Engine for Weekly & Monthly Executive Reports.
    Period Rules:
      - Weekly: Saturday -> Friday
      - Monthly: 1st of Month -> Last Day of Month
    Scoring Rule:
      - Global Retail Risk-Based Scoring System
    """
    def __init__(self, loader):
        self.loader = loader
        self.dim_store = loader.load_dim_store()

    @staticmethod
    def get_period_date_range(period_type: str, reference_date: datetime = None) -> Tuple[datetime, datetime]:
        if reference_date is None:
            reference_date = datetime.now()
            
        if period_type == "weekly":
            # Find closest previous/current Saturday
            # Python weekday: Mon=0, Tue=1, Wed=2, Thu=3, Fri=4, Sat=5, Sun=6
            offset = (reference_date.weekday() + 2) % 7
            start_date = (reference_date - timedelta(days=offset)).replace(hour=0, minute=0, second=0, microsecond=0)
            end_date = (start_date + timedelta(days=6)).replace(hour=23, minute=59, second=59, microsecond=999999)
        else: # monthly
            start_date = reference_date.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
            # Find last day of month
            if start_date.month == 12:
                next_month = start_date.replace(year=start_date.year + 1, month=1, day=1)
            else:
                next_month = start_date.replace(month=start_date.month + 1, day=1)
            end_date = (next_month - timedelta(seconds=1))
            
        return start_date, end_date

    def aggregate_data(self, period_type: str = "weekly", asm_filter: str = "ALL", reference_date: datetime = None) -> Dict[str, Any]:
        start_date, end_date = self.get_period_date_range(period_type, reference_date)
        start_str = start_date.strftime("%d/%m/%Y")
        end_str = end_date.strftime("%d/%m/%Y")
        period_name = f"Tuần ({start_str} - {end_str})" if period_type == "weekly" else f"Tháng {start_date.strftime('%m/%Y')}"

        # 1. Load StoreVisit Submissions from cache
        form_cache_path = os.path.join(self.loader.root_dir, "data", "form_cache.json")
        submissions = []
        if os.path.exists(form_cache_path):
            try:
                with open(form_cache_path, "r", encoding="utf-8") as f:
                    cache_data = json.load(f)
                    for _, sub in cache_data.items():
                        rdate_str = sub.get("report_date", "")
                        dt = self._parse_date(rdate_str)
                        if dt and start_date <= dt <= end_date:
                            asm_name = sub.get("asm_name", "")
                            if asm_filter == "ALL" or asm_name.strip().lower() == asm_filter.strip().lower():
                                submissions.append(sub)
            except Exception as e:
                print(f"[Aggregator] Error loading form cache: {e}")

        # 2. Load Market Survey Responses from cache
        survey_cache_path = os.path.join(self.loader.root_dir, "data", "survey_cache.json")
        surveys = []
        if os.path.exists(survey_cache_path):
            try:
                with open(survey_cache_path, "r", encoding="utf-8") as f:
                    surv_data = json.load(f)
                    for _, s in surv_data.items():
                        sdate_str = s.get("timestamp", "")
                        dt = self._parse_date(sdate_str)
                        if dt and start_date <= dt <= end_date:
                            surveys.append(s)
            except Exception as e:
                print(f"[Aggregator] Error loading survey cache: {e}")

        # 3. Calculate Store Health Matrix
        store_matrix = []
        systemic_issues = {}
        total_critical_violations = 0

        for sub in submissions:
            store_code = sub.get("store_code", "").upper()
            store_name = sub.get("store_name", store_code)
            asm_name = sub.get("asm_name", "Khác")
            checklist_json = sub.get("checklist_json", "{}")
            
            try:
                checklist = json.loads(checklist_json) if isinstance(checklist_json, str) else checklist_json
            except Exception:
                checklist = {}

            total_items = 0
            passed_items = 0
            failed_items = 0
            na_items = 0
            critical_count = 0
            open_issues = []

            sections = checklist.get("sections", {})
            for sec_key, sec in sections.items():
                items = sec.get("items", [])
                for item in items:
                    eval_val = item.get("eval", "")
                    severity = item.get("severity", "")
                    label = item.get("label", item.get("id", ""))

                    if eval_val == "Đạt":
                        passed_items += 1
                        total_items += 1
                    elif eval_val == "Không đạt":
                        failed_items += 1
                        total_items += 1
                        if severity in ["Khẩn cấp", "Cao"]:
                            critical_count += 1
                            total_critical_violations += 1
                        
                        systemic_issues[label] = systemic_issues.get(label, 0) + 1
                        open_issues.append({
                            "store_code": store_code,
                            "store_name": store_name,
                            "asm_name": asm_name,
                            "issue_label": label,
                            "severity": severity,
                            "assignee": item.get("assignee", "CHT"),
                            "deadline": item.get("deadline", "---"),
                            "note": item.get("note", "")
                        })
                    elif eval_val == "Không áp dụng":
                        na_items += 1

            base_pass_rate = (passed_items / total_items * 100.0) if total_items > 0 else 100.0
            health_score = max(0.0, round(base_pass_rate - (critical_count * 15.0), 1))
            
            if health_score >= 90.0 and critical_count == 0:
                status_label = "Tốt"
            elif health_score >= 80.0 and critical_count == 0:
                status_label = "Đạt"
            else:
                status_label = "Chưa Đạt"

            store_matrix.append({
                "submission_id": sub.get("submission_id", ""),
                "report_date": sub.get("report_date", ""),
                "store_code": store_code,
                "store_name": store_name,
                "asm_name": asm_name,
                "passed_items": passed_items,
                "failed_items": failed_items,
                "na_items": na_items,
                "total_applicable": total_items,
                "base_pass_rate": round(base_pass_rate, 1),
                "critical_violations": critical_count,
                "health_score": health_score,
                "status_label": status_label,
                "open_issues": open_issues
            })

        # Sort store matrix by health score descending
        store_matrix.sort(key=lambda x: x["health_score"], reverse=True)

        # 4. Calculate ASM Leaderboard
        asm_leaderboard = {}
        if not self.dim_store.empty:
            asm_store_counts = self.dim_store.groupby("ASM")["StoreCode"].count().to_dict()
        else:
            asm_store_counts = {}

        for row in store_matrix:
            asm = row["asm_name"]
            if asm not in asm_leaderboard:
                asm_leaderboard[asm] = {
                    "asm_name": asm,
                    "total_assigned_stores": asm_store_counts.get(asm, 0),
                    "visited_count": 0,
                    "scores": [],
                    "critical_count": 0,
                    "passed_stores": 0
                }
            asm_leaderboard[asm]["visited_count"] += 1
            asm_leaderboard[asm]["scores"].append(row["health_score"])
            asm_leaderboard[asm]["critical_count"] += row["critical_violations"]
            if row["status_label"] in ["Tốt", "Đạt"]:
                asm_leaderboard[asm]["passed_stores"] += 1

        leaderboard_list = []
        for asm, data in asm_leaderboard.items():
            tot = data["total_assigned_stores"]
            vis = data["visited_count"]
            coverage_pct = round((vis / tot * 100.0), 1) if tot > 0 else 100.0
            avg_score = round(sum(data["scores"]) / len(data["scores"]), 1) if data["scores"] else 0.0
            leaderboard_list.append({
                "asm_name": asm,
                "assigned_stores": tot,
                "visited_stores": vis,
                "coverage_pct": coverage_pct,
                "avg_health_score": avg_score,
                "critical_violations": data["critical_count"],
                "pass_rate_pct": round((data["passed_stores"] / vis * 100.0), 1) if vis > 0 else 0.0
            })

        leaderboard_list.sort(key=lambda x: (x["avg_health_score"], x["coverage_pct"]), reverse=True)

        # 5. Top 5 Systemic Issues
        sorted_systemic = sorted(systemic_issues.items(), key=lambda x: x[1], reverse=True)[:5]
        top_systemic_issues = [{"label": k, "count": v} for k, v in sorted_systemic]

        # Overall Summary KPIs
        total_visited = len(store_matrix)
        avg_network_score = round(sum(s["health_score"] for s in store_matrix) / total_visited, 1) if total_visited > 0 else 0.0
        good_count = sum(1 for s in store_matrix if s["status_label"] == "Tốt")
        pass_count = sum(1 for s in store_matrix if s["status_label"] == "Đạt")
        fail_count = sum(1 for s in store_matrix if s["status_label"] == "Chưa Đạt")

        return {
            "period_type": period_type,
            "period_name": period_name,
            "start_date": start_str,
            "end_date": end_str,
            "asm_filter": asm_filter,
            "kpis": {
                "total_visited": total_visited,
                "avg_network_score": avg_network_score,
                "good_stores_count": good_count,
                "pass_stores_count": pass_count,
                "fail_stores_count": fail_count,
                "critical_violations": total_critical_violations,
                "market_surveys_count": len(surveys)
            },
            "store_matrix": store_matrix,
            "asm_leaderboard": leaderboard_list,
            "top_systemic_issues": top_systemic_issues,
            "market_surveys": surveys
        }

    def _parse_date(self, dstr: str) -> datetime:
        if not dstr:
            return None
        dstr = str(dstr).strip()
        for fmt in ["%Y-%m-%d", "%d/%m/%Y", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S"]:
            try:
                return datetime.strptime(dstr.split("T")[0], fmt.split("T")[0])
            except Exception:
                pass
        return None
