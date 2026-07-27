import pandas as pd
from typing import Optional
from .models import RevenueData

class RevenueRepository:
    def __init__(self, loader, field_mapping: dict):
        self.loader = loader
        self.store_code_mapping = field_mapping.get("store_code_mapping", {})

    def get_revenue_data(self, store_key: str, year: int = 2026, month: int = 6) -> RevenueData:
        store_code = self.store_code_mapping.get(store_key.upper(), store_key.upper())
        df_rev = self.loader.load_revenue()
        
        # 1. June 2026 Retail Revenue
        rev_curr_mask = (
            (df_rev["StoreCode"] == store_code) &
            (df_rev["Date"].dt.year == year) &
            (df_rev["Date"].dt.month == month) &
            (df_rev["SalesType"] == "Retail")
        )
        total_rev_curr = int(df_rev[rev_curr_mask]["Revenue"].sum())

        # 2. May 2026 Retail Revenue (MoM)
        prev_month = 12 if month == 1 else month - 1
        prev_year = year - 1 if month == 1 else year
        rev_prev_mask = (
            (df_rev["StoreCode"] == store_code) &
            (df_rev["Date"].dt.year == prev_year) &
            (df_rev["Date"].dt.month == prev_month) &
            (df_rev["SalesType"] == "Retail")
        )
        total_rev_prev = int(df_rev[rev_prev_mask]["Revenue"].sum())

        # 3. June 2025 Retail Revenue (YoY)
        rev_yoy_mask = (
            (df_rev["StoreCode"] == store_code) &
            (df_rev["Date"].dt.year == year - 1) &
            (df_rev["Date"].dt.month == month) &
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

        # Calculate percentages
        attainment = (total_rev_curr / total_tgt_curr * 100) if total_tgt_curr > 0 else 0.0
        
        diff_mom = total_rev_curr - total_rev_prev
        pct_mom = (diff_mom / total_rev_prev * 100) if total_rev_prev > 0 else 0.0
        
        diff_yoy = total_rev_curr - total_rev_yoy
        pct_yoy = (diff_yoy / total_rev_yoy * 100) if total_rev_yoy > 0 else 0.0

        # Build commentary
        comment = (
            f"Nhận xét tình hình doanh thu bán lẻ:\n"
            f"• Doanh thu thực tế lũy kế tháng đạt {total_rev_curr:,.0f} VNĐ, hoàn thành {attainment:.1f}% kế hoạch tháng.\n"
        )
        if diff_mom > 0:
            comment += f"• So với tháng trước (MoM), tăng trưởng {pct_mom:.1f}% (+{diff_mom:,.0f} VNĐ).\n"
        else:
            comment += f"• So với tháng trước (MoM), giảm {abs(pct_mom):.1f}% ({diff_mom:,.0f} VNĐ) do biến động mùa vụ.\n"
            
        if diff_yoy > 0:
            comment += f"• So với cùng kỳ năm trước (YoY), tăng trưởng {pct_yoy:.1f}% (+{diff_yoy:,.0f} VNĐ)."
        else:
            comment += f"• So với cùng kỳ năm trước (YoY), giảm {abs(pct_yoy):.1f}% ({diff_yoy:,.0f} VNĐ)."

        return RevenueData(
            revenue_actual=total_rev_curr,
            revenue_target=total_tgt_curr,
            attainment_pct=attainment,
            revenue_prev=total_rev_prev,
            revenue_yoy=total_rev_yoy,
            mom_change_pct=pct_mom,
            yoy_change_pct=pct_yoy,
            commentary=comment
        )
