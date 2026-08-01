import os
import json
import re
from datetime import datetime, timedelta
import pandas as pd
from typing import Dict, List, Any, Tuple

class WeeklyMonthlyAggregator:
    """
    Top 0.1% Data Aggregation Engine for Weekly & Monthly Executive Reports.
    Features:
      - Accent-insensitive & Fuzzy ASM Name Normalization
      - Robust Multi-format Date Parsing
      - Global Retail Risk-Based Scoring System with Keyword & Category Weighting
      - Multi-period calculation (Weekly, Monthly, Quarterly)
    """
    def __init__(self, loader):
        self.loader = loader
        self.dim_store = loader.load_dim_store()

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

    @staticmethod
    def get_period_date_range(period_type: str, reference_date: datetime = None) -> Tuple[datetime, datetime]:
        if reference_date is None:
            reference_date = datetime.now()
            
        if period_type == "weekly":
            # Find closest previous/current Saturday (Saturday -> Friday)
            # Python weekday: Mon=0, Tue=1, Wed=2, Thu=3, Fri=4, Sat=5, Sun=6
            offset = (reference_date.weekday() + 2) % 7
            start_date = (reference_date - timedelta(days=offset)).replace(hour=0, minute=0, second=0, microsecond=0)
            end_date = (start_date + timedelta(days=6)).replace(hour=23, minute=59, second=59, microsecond=999999)
        elif period_type == "monthly":
            start_date = reference_date.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
            if start_date.month == 12:
                next_month = start_date.replace(year=start_date.year + 1, month=1, day=1)
            else:
                next_month = start_date.replace(month=start_date.month + 1, day=1)
            end_date = (next_month - timedelta(seconds=1))
        else: # quarterly
            curr_quarter = (reference_date.month - 1) // 3 + 1
            first_month_of_q = (curr_quarter - 1) * 3 + 1
            start_date = reference_date.replace(month=first_month_of_q, day=1, hour=0, minute=0, second=0, microsecond=0)
            last_month_of_q = first_month_of_q + 2
            if last_month_of_q == 12:
                next_q = start_date.replace(year=start_date.year + 1, month=1, day=1)
            else:
                next_q = start_date.replace(month=last_month_of_q + 1, day=1)
            end_date = (next_q - timedelta(seconds=1))
            
        return start_date, end_date

    def aggregate_data(self, period_type: str = "weekly", asm_filter: str = "ALL", reference_date: datetime = None) -> Dict[str, Any]:
        start_date, end_date = self.get_period_date_range(period_type, reference_date)
        start_str = start_date.strftime("%d/%m/%Y")
        end_str = end_date.strftime("%d/%m/%Y")
        
        if period_type == "weekly":
            period_name = f"Tuần ({start_str} - {end_str})"
        elif period_type == "monthly":
            period_name = f"Tháng {start_date.strftime('%m/%Y')}"
        else:
            q_num = (start_date.month - 1) // 3 + 1
            period_name = f"Quý {q_num}/{start_date.year}"

        asm_filter_norm = self.remove_accents(asm_filter)

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
                            asm_norm = self.remove_accents(asm_name)
                            if asm_filter == "ALL" or asm_filter_norm in asm_norm or asm_norm in asm_filter_norm:
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
                    eval_val = str(item.get("eval", "")).strip()
                    severity = str(item.get("severity", "")).strip()
                    label = str(item.get("label", item.get("id", ""))).strip()
                    sev_lower = severity.lower()

                    if eval_val == "Đạt":
                        passed_items += 1
                        total_items += 1
                    elif eval_val == "Không đạt":
                        failed_items += 1
                        total_items += 1
                        
                        # Robust Critical Detection (Case & Keyword Insensitive)
                        is_critical = (
                            sev_lower in ["khẩn cấp", "cao", "nghiêm trọng", "critical", "high"] or
                            sec_key in ["security_guard", "cash", "fire_safety"] or
                            "pccc" in label.lower() or "quầy thu ngân" in label.lower()
                        )
                        if is_critical:
                            critical_count += 1
                            total_critical_violations += 1
                        
                        systemic_issues[label] = systemic_issues.get(label, 0) + 1
                        open_issues.append({
                            "store_code": store_code,
                            "store_name": store_name,
                            "asm_name": asm_name,
                            "issue_label": label,
                            "severity": severity or ("Khẩn cấp" if is_critical else "Bình thường"),
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

        # 4. Calculate ASM Leaderboard with Normalized Name Matching
        asm_store_counts = {}
        if not self.dim_store.empty:
            for _, r in self.dim_store.iterrows():
                raw_asm = str(r.get("ASM", "")).strip()
                if raw_asm:
                    norm = self.remove_accents(raw_asm)
                    asm_store_counts[norm] = asm_store_counts.get(norm, 0) + 1

        asm_leaderboard = {}
        for row in store_matrix:
            raw_asm = row["asm_name"]
            norm_asm = self.remove_accents(raw_asm)
            
            if norm_asm not in asm_leaderboard:
                asm_leaderboard[norm_asm] = {
                    "display_name": raw_asm,
                    "total_assigned_stores": asm_store_counts.get(norm_asm, 0),
                    "visited_count": 0,
                    "scores": [],
                    "critical_count": 0,
                    "passed_stores": 0
                }
            asm_leaderboard[norm_asm]["visited_count"] += 1
            asm_leaderboard[norm_asm]["scores"].append(row["health_score"])
            asm_leaderboard[norm_asm]["critical_count"] += row["critical_violations"]
            if row["status_label"] in ["Tốt", "Đạt"]:
                asm_leaderboard[norm_asm]["passed_stores"] += 1

        leaderboard_list = []
        for norm, data in asm_leaderboard.items():
            tot = data["total_assigned_stores"]
            vis = data["visited_count"]
            coverage_pct = round((vis / tot * 100.0), 1) if tot > 0 else 100.0
            avg_score = round(sum(data["scores"]) / len(data["scores"]), 1) if data["scores"] else 0.0
            leaderboard_list.append({
                "asm_name": data["display_name"],
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
