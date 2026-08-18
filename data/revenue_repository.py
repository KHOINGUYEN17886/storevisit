import pandas as pd
from typing import Optional
from .models import RevenueData

class RevenueRepository:
    def __init__(self, loader, field_mapping: dict):
        self.loader = loader
        self.store_code_mapping = field_mapping.get("store_code_mapping", {})

    def get_revenue_data(self, store_key: str, year: int = 2026, month: int = 6, cutoff_day: Optional[int] = None) -> RevenueData:
        store_code = self.store_code_mapping.get(store_key.upper(), store_key.upper())
        df_rev = self.loader.load_revenue()
        
        # Calculate calendar days in respective months
        days_in_curr_month = pd.Period(f"{year}-{month:02d}").days_in_month
        prev_month = 12 if month == 1 else month - 1
        prev_year = year - 1 if month == 1 else year
        days_in_prev_month = pd.Period(f"{prev_year}-{prev_month:02d}").days_in_month
        days_in_yoy_month = pd.Period(f"{year - 1}-{month:02d}").days_in_month

        # Determine MTD cutoff limits
        is_mtd = cutoff_day is not None and cutoff_day < days_in_curr_month
        curr_limit_day = min(cutoff_day, days_in_curr_month) if cutoff_day else days_in_curr_month
        prev_limit_day = min(cutoff_day, days_in_prev_month) if cutoff_day else days_in_prev_month
        yoy_limit_day = min(cutoff_day, days_in_yoy_month) if cutoff_day else days_in_yoy_month

        # 1. Current Month Retail Revenue (MTD up to curr_limit_day)
        rev_curr_mask = (
            (df_rev["StoreCode"] == store_code) &
            (df_rev["Date"].dt.year == year) &
            (df_rev["Date"].dt.month == month) &
            (df_rev["Date"].dt.day <= curr_limit_day) &
            (df_rev["SalesType"] == "Retail")
        )
        total_rev_curr = int(df_rev[rev_curr_mask]["Revenue"].sum())

        # 2. Previous Month Retail Revenue (MoM Like-for-Like up to prev_limit_day)
        rev_prev_mask = (
            (df_rev["StoreCode"] == store_code) &
            (df_rev["Date"].dt.year == prev_year) &
            (df_rev["Date"].dt.month == prev_month) &
            (df_rev["Date"].dt.day <= prev_limit_day) &
            (df_rev["SalesType"] == "Retail")
        )
        total_rev_prev = int(df_rev[rev_prev_mask]["Revenue"].sum())

        # 3. Same Month Last Year Retail Revenue (YoY Like-for-Like up to yoy_limit_day)
        rev_yoy_mask = (
            (df_rev["StoreCode"] == store_code) &
            (df_rev["Date"].dt.year == year - 1) &
            (df_rev["Date"].dt.month == month) &
            (df_rev["Date"].dt.day <= yoy_limit_day) &
            (df_rev["SalesType"] == "Retail")
        )
        total_rev_yoy = int(df_rev[rev_yoy_mask]["Revenue"].sum())

        # 4. Target for current month
        df_tgt = self.loader.load_target()
        tgt_row = df_tgt[
            (df_tgt["StoreCode"] == store_code) &
            (df_tgt["Year"] == year) &
            (df_tgt["MonthNo"] == month) &
            (df_tgt["BrandGroup"] == "Store")
        ]
        
        if tgt_row.empty:
            # Fallback to sum of targets divided by 2 if BrandGroup Store is missing
            tgt_row_any = df_tgt[
                (df_tgt["StoreCode"] == store_code) &
                (df_tgt["Year"] == year) &
                (df_tgt["MonthNo"] == month)
            ]
            total_tgt_curr = int(tgt_row_any["Target"].sum() / 2.0) if not tgt_row_any.empty else 0
        else:
            total_tgt_curr = int(tgt_row["Target"].sum())

        # Prorated MTD Target
        total_tgt_mtd = int(total_tgt_curr * curr_limit_day / days_in_curr_month) if total_tgt_curr > 0 else 0

        # Calculate percentages
        attainment_full = (total_rev_curr / total_tgt_curr * 100) if total_tgt_curr > 0 else 0.0
        attainment_mtd = (total_rev_curr / total_tgt_mtd * 100) if total_tgt_mtd > 0 else 0.0
        
        diff_mom = total_rev_curr - total_rev_prev
        pct_mom = (diff_mom / total_rev_prev * 100) if total_rev_prev > 0 else 0.0
        
        diff_yoy = total_rev_curr - total_rev_yoy
        pct_yoy = (diff_yoy / total_rev_yoy * 100) if total_rev_yoy > 0 else 0.0

        # Build commentary
        if is_mtd:
            comment = (
                f"Nhận xét tình hình doanh thu bán lẻ lũy kế MTD (01/{month:02d} - {curr_limit_day:02d}/{month:02d}/{year}):\n"
                f"• Doanh thu thực tế đạt {total_rev_curr:,.0f} VNĐ, hoàn thành {attainment_mtd:.1f}% tiến độ MTD ({attainment_full:.1f}% KH cả tháng).\n"
            )
            if diff_mom > 0:
                comment += f"• So với cùng kỳ tháng trước (MoM cùng {curr_limit_day} ngày), tăng trưởng {pct_mom:.1f}% (+{diff_mom:,.0f} VNĐ).\n"
            elif diff_mom == 0 or abs(pct_mom) < 0.05:
                comment += f"• So với cùng kỳ tháng trước (MoM cùng {curr_limit_day} ngày), duy trì ổn định ngang mức (+0.0%).\n"
            else:
                comment += f"• So với cùng kỳ tháng trước (MoM cùng {curr_limit_day} ngày), giảm {abs(pct_mom):.1f}% ({diff_mom:,.0f} VNĐ).\n"
                
            if diff_yoy > 0:
                comment += f"• So với cùng kỳ năm trước (YoY cùng {curr_limit_day} ngày), tăng trưởng {pct_yoy:.1f}% (+{diff_yoy:,.0f} VNĐ)."
            elif diff_yoy == 0 or abs(pct_yoy) < 0.05:
                comment += f"• So với cùng kỳ năm trước (YoY cùng {curr_limit_day} ngày), tương đương cùng kỳ (+0.0%)."
            else:
                comment += f"• So với cùng kỳ năm trước (YoY cùng {curr_limit_day} ngày), giảm {abs(pct_yoy):.1f}% ({diff_yoy:,.0f} VNĐ)."
        else:
            comment = (
                f"Nhận xét tình hình doanh thu bán lẻ tháng {month:02d}/{year}:\n"
                f"• Doanh thu thực tế đạt {total_rev_curr:,.0f} VNĐ, hoàn thành {attainment_full:.1f}% kế hoạch tháng.\n"
            )
            if diff_mom > 0:
                comment += f"• So với tháng trước (MoM), tăng trưởng {pct_mom:.1f}% (+{diff_mom:,.0f} VNĐ).\n"
            elif diff_mom == 0 or abs(pct_mom) < 0.05:
                comment += f"• So với tháng trước (MoM), duy trì ổn định (+0.0%).\n"
            else:
                comment += f"• So với tháng trước (MoM), giảm {abs(pct_mom):.1f}% ({diff_mom:,.0f} VNĐ) do biến động mùa vụ.\n"
                
            if diff_yoy > 0:
                comment += f"• So với cùng kỳ năm trước (YoY), tăng trưởng {pct_yoy:.1f}% (+{diff_yoy:,.0f} VNĐ)."
            elif diff_yoy == 0 or abs(pct_yoy) < 0.05:
                comment += f"• So với cùng kỳ năm trước (YoY), tương đương cùng kỳ (+0.0%)."
            else:
                comment += f"• So với cùng kỳ năm trước (YoY), giảm {abs(pct_yoy):.1f}% ({diff_yoy:,.0f} VNĐ)."

        return RevenueData(
            revenue_actual=total_rev_curr,
            revenue_target=total_tgt_curr,
            attainment_pct=attainment_full,
            revenue_prev=total_rev_prev,
            revenue_yoy=total_rev_yoy,
            mom_change_pct=pct_mom,
            yoy_change_pct=pct_yoy,
            commentary=comment,
            cutoff_day=curr_limit_day if is_mtd else None,
            is_mtd=is_mtd,
            revenue_target_mtd=total_tgt_mtd,
            attainment_mtd_pct=attainment_mtd
        )
