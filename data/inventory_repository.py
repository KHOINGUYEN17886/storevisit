import pandas as pd
from typing import List, Dict
from .models import StockInventory, BestSellerItem, SlowSellerItem

class InventoryRepository:
    def __init__(self, loader, field_mapping: dict):
        self.loader = loader
        self.store_code_mapping = field_mapping.get("store_code_mapping", {})
        self.period_variants = field_mapping.get("product_period_variants", {
            "nguyen_gia": "Nguyên giá",
            "sale": "Sale",
            "thanh_ly": "Thanh lý"
        })

    def get_stock_inventory(self, store_key: str) -> StockInventory:
        store_code = self.store_code_mapping.get(store_key.upper(), store_key.upper())
        df_stock = self.loader.load_stock()
        df_store_stock = df_stock[df_stock["StoreCode"] == store_code]

        total_qty = int(df_store_stock["Quantity"].sum())
        
        # Calculate quantities by product period
        qty_nguyen_gia = int(df_store_stock[df_store_stock["ProductPeriod"] == self.period_variants.get("nguyen_gia", "Nguyên giá")]["Quantity"].sum())
        qty_sale = int(df_store_stock[df_store_stock["ProductPeriod"] == self.period_variants.get("sale", "Sale")]["Quantity"].sum())
        qty_thanh_ly = int(df_store_stock[df_store_stock["ProductPeriod"] == self.period_variants.get("thanh_ly", "Thanh lý")]["Quantity"].sum())

        # Calculate SKU count from stock file
        skus_count = int(df_store_stock["ProductCode"].nunique())

        # Age group classification based on distribution year/month
        df_store_stock_copy = df_store_stock.copy()
        
        def get_bucket(r):
            pp = r.get("ProductPeriod", "Nguyên giá")
            year_pp = r.get("Năm phân phối")
            month_pp = r.get("Tháng phân phối")
            
            if pp == "Sale":
                return "Hàng sale"
            elif pp == "Thanh lý":
                return "Hàng thanh lý"
                
            if pd.isna(year_pp) or pd.isna(month_pp):
                return "Khác/Chưa rõ"
                
            try:
                y = float(year_pp)
                m = float(month_pp)
            except (ValueError, TypeError):
                return "Khác/Chưa rõ"
                
            if y == 2026 and m == 7:
                return "Đợt PP tháng 7/2026"
            elif y == 2026 and m in [4, 5, 6]:
                return "Quý 2/2026"
            elif y == 2026 and m in [1, 2, 3]:
                return "Quý 1/2026"
            elif y == 2025 and m in [10, 11, 12]:
                return "Quý 4/2025"
            elif y == 2025 and m in [7, 8, 9]:
                return "Quý 3/2025"
            else:
                return "Hàng nguyên giá PP > 1 năm"
                
        df_store_stock_copy["AgeBucket"] = df_store_stock_copy.apply(get_bucket, axis=1)
        
        age_groups = {
            "Đợt PP tháng 7/2026": int(df_store_stock_copy[df_store_stock_copy["AgeBucket"] == "Đợt PP tháng 7/2026"]["Quantity"].sum()),
            "Quý 2/2026": int(df_store_stock_copy[df_store_stock_copy["AgeBucket"] == "Quý 2/2026"]["Quantity"].sum()),
            "Quý 1/2026": int(df_store_stock_copy[df_store_stock_copy["AgeBucket"] == "Quý 1/2026"]["Quantity"].sum()),
            "Quý 4/2025": int(df_store_stock_copy[df_store_stock_copy["AgeBucket"] == "Quý 4/2025"]["Quantity"].sum()),
            "Quý 3/2025": int(df_store_stock_copy[df_store_stock_copy["AgeBucket"] == "Quý 3/2025"]["Quantity"].sum()),
            "Hàng nguyên giá PP > 1 năm": int(df_store_stock_copy[df_store_stock_copy["AgeBucket"] == "Hàng nguyên giá PP > 1 năm"]["Quantity"].sum()),
            "Hàng sale": int(df_store_stock_copy[df_store_stock_copy["AgeBucket"] == "Hàng sale"]["Quantity"].sum()),
            "Hàng thanh lý": int(df_store_stock_copy[df_store_stock_copy["AgeBucket"] == "Hàng thanh lý"]["Quantity"].sum()),
            "Khác/Chưa rõ": int(df_store_stock_copy[df_store_stock_copy["AgeBucket"] == "Khác/Chưa rõ"]["Quantity"].sum())
        }
        
        total_valid = sum(age_groups.values()) or 1
        pcts = {k: v / total_valid * 100 for k, v in age_groups.items()}
        
        commentary = (
            f"Phân tích cơ cấu tồn kho theo thời gian phân phối:\n"
            f"• Đợt PP tháng 7/2026 (mới nhất): {age_groups['Đợt PP tháng 7/2026']:,.0f} SP ({pcts['Đợt PP tháng 7/2026']:.1f}%)\n"
            f"• Quý 2/2026: {age_groups['Quý 2/2026']:,.0f} SP ({pcts['Quý 2/2026']:.1f}%)\n"
            f"• Quý 1/2026: {age_groups['Quý 1/2026']:,.0f} SP ({pcts['Quý 1/2026']:.1f}%)\n"
            f"• Quý 4/2025: {age_groups['Quý 4/2025']:,.0f} SP ({pcts['Quý 4/2025']:.1f}%)\n"
            f"• Quý 3/2025: {age_groups['Quý 3/2025']:,.0f} SP ({pcts['Quý 3/2025']:.1f}%)\n"
            f"• Hàng nguyên giá phân phối > 1 năm: {age_groups['Hàng nguyên giá PP > 1 năm']:,.0f} SP ({pcts['Hàng nguyên giá PP > 1 năm']:.1f}%)\n"
            f"• Hàng Sale Off: {age_groups['Hàng sale']:,.0f} SP ({pcts['Hàng sale']:.1f}%)\n"
            f"• Hàng Thanh Lý: {age_groups['Hàng thanh lý']:,.0f} SP ({pcts['Hàng thanh lý']:.1f}%)\n"
            f"Nhận xét: Hàng nguyên giá mới phân phối trong vòng 1 năm (từ Quý 3/2025 đến nay) chiếm tỷ trọng {pcts['Đợt PP tháng 7/2026'] + pcts['Quý 2/2026'] + pcts['Quý 1/2026'] + pcts['Quý 4/2025'] + pcts['Quý 3/2025']:.1f}%. Hàng tồn lâu phân phối trên 1 năm là {pcts['Hàng nguyên giá PP > 1 năm']:.1f}% ({age_groups['Hàng nguyên giá PP > 1 năm']:,.0f} SP). Cần lên kế hoạch điều chuyển hoặc thúc đẩy tiêu thụ nhóm hàng tồn lâu."
        )
        
        return StockInventory(
            total_qty=total_qty,
            skus_count=skus_count,
            qty_nguyen_gia=qty_nguyen_gia,
            qty_sale=qty_sale,
            qty_thanh_ly=qty_thanh_ly,
            age_groups=age_groups,
            commentary=commentary
        )

    def get_best_sellers(self, store_key: str, limit: int = 10) -> List[BestSellerItem]:
        store_code = self.store_code_mapping.get(store_key.upper(), store_key.upper())
        df_hm = self.loader.load_health_master()
        df_hm_store = df_hm[df_hm["StoreCode"] == store_code].copy()
        
        # Filter out gift vouchers
        df_hm_store = df_hm_store[~df_hm_store["ProductName"].str.contains("Phiếu quà tặng|Voucher", case=False, na=False)]
        
        for col in ["Sales_4W", "Stock_Qty"]:
            df_hm_store[col] = pd.to_numeric(df_hm_store[col], errors="coerce").fillna(0)
            
        best_df = df_hm_store.sort_values(by="Sales_4W", ascending=False).head(limit)
        items = []
        for i, (_, row) in enumerate(best_df.iterrows()):
            items.append(BestSellerItem(
                rank=i + 1,
                sku=str(row.get("SKU", "")),
                product_name=str(row.get("ProductName", "")),
                brand=str(row.get("Brand", "")),
                sales_4w=int(row.get("Sales_4W", 0)),
                stock_qty=int(row.get("Stock_Qty", 0))
            ))
        return items

    def get_slow_sellers(self, store_key: str, limit: int = 10) -> List[SlowSellerItem]:
        store_code = self.store_code_mapping.get(store_key.upper(), store_key.upper())
        df_hm = self.loader.load_health_master()
        df_hm_store = df_hm[df_hm["StoreCode"] == store_code].copy()
        
        # Filter out gift vouchers
        df_hm_store = df_hm_store[~df_hm_store["ProductName"].str.contains("Phiếu quà tặng|Voucher", case=False, na=False)]
        
        for col in ["Sales_4W", "Stock_Qty", "Stock_Age_Days"]:
            df_hm_store[col] = pd.to_numeric(df_hm_store[col], errors="coerce").fillna(0)
            
        # Slow sellers: Sales_4W == 0, sorted by Stock_Qty descending
        slow_df = df_hm_store[df_hm_store["Sales_4W"] == 0].sort_values(by="Stock_Qty", ascending=False).head(limit)
        items = []
        for i, (_, row) in enumerate(slow_df.iterrows()):
            items.append(SlowSellerItem(
                rank=i + 1,
                sku=str(row.get("SKU", "")),
                product_name=str(row.get("ProductName", "")),
                brand=str(row.get("Brand", "")),
                stock_qty=int(row.get("Stock_Qty", 0)),
                age_days=int(row.get("Stock_Age_Days", 0))
            ))
        return items
