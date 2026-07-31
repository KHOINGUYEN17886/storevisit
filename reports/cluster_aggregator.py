from typing import List
from data.models import StoreReportData, ClusterReportData, ClusterStorePerformance, ClusterCriticalIssue
from datetime import datetime

class ClusterAggregator:
    def __init__(self):
        pass

    def aggregate_stores(self, store_reports: List[StoreReportData], cluster_name: str = "Cụm cửa hàng") -> ClusterReportData:
        """
        Aggregate multiple individual store reports into a single ClusterReportData model.
        Crucially, cluster attainment is Sum(Actual) / Sum(Target), not average of percentages.
        """
        total_actual = 0
        total_target = 0
        
        stores_performance: List[ClusterStorePerformance] = []
        critical_issues: List[ClusterCriticalIssue] = []

        for report in store_reports:
            actual = report.revenue.revenue_actual
            target = report.revenue.revenue_target
            attainment = report.revenue.attainment_pct
            
            total_actual += actual
            total_target += target

            # Add to performance table
            stores_performance.append(ClusterStorePerformance(
                store_code=report.metadata.store_code,
                store_name=report.metadata.store_name,
                revenue_actual=actual,
                revenue_target=target,
                attainment_pct=attainment
            ))

            # Collect high priority operational issues
            for issue in report.issues:
                # If priority is High, Urgent, or similar, or contains red status symbols
                is_critical = False
                notes_lower = issue.notes.lower()
                if "cao" in notes_lower or "khẩn" in notes_lower or "priority: a" in notes_lower:
                    is_critical = True
                
                if is_critical or len(critical_issues) < 5: # Fallback: take first few if no high priority ones exist
                    priority_label = "Cao" if is_critical else "Bình thường"
                    critical_issues.append(ClusterCriticalIssue(
                        store_name=report.metadata.store_name,
                        label=issue.label,
                        issue=issue.issue,
                        assignee=issue.assignee,
                        priority=priority_label
                    ))

        # Attainment is Sum(Actual) / Sum(Target)
        cluster_attainment = (total_actual / total_target * 100) if total_target > 0 else 0.0
        
        # Sort performance table descending by attainment
        stores_performance.sort(key=lambda x: x.attainment_pct, reverse=True)
        
        # Limit critical issues table to 10 rows maximum for PPTX table slide height
        critical_issues = critical_issues[:10]

        report_date = datetime.now().strftime("%d/%m/%Y")

        return ClusterReportData(
            cluster_name=cluster_name,
            report_date=report_date,
            revenue_actual=total_actual,
            revenue_target=total_target,
            attainment_pct=cluster_attainment,
            stores_performance=stores_performance,
            critical_issues=critical_issues,
            individual_reports=store_reports
        )
