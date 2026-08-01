import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

class ExecutiveExcelGenerator:
    """
    Top 0.1% Executive Excel Dashboard Generator (.xlsx)
    Brand Palette: An Phước Navy (#1B2A4A), Red Accent (#C41E3A), Gold (#D4AF37)
    Features:
      - Smart Auto-Column Width calculation (Excludes Title Banners to prevent over-stretching)
      - Explicit Cell Number/Percentage Formatting
      - Soft Status Fills (Green/Yellow/Red) & High Contrast Typography
    """
    def __init__(self):
        self.navy_fill = PatternFill(start_color="1B2A4A", end_color="1B2A4A", fill_type="solid")
        self.header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        self.title_font = Font(name="Calibri", size=16, bold=True, color="1B2A4A")
        self.subtitle_font = Font(name="Calibri", size=11, italic=True, color="555555")
        
        self.good_fill = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid") # Soft Green
        self.good_font = Font(name="Calibri", size=11, color="155724", bold=True)
        self.pass_fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid") # Soft Yellow
        self.pass_font = Font(name="Calibri", size=11, color="856404", bold=True)
        self.fail_fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid") # Soft Red
        self.fail_font = Font(name="Calibri", size=11, color="721C24", bold=True)
        
        self.thin_border = Border(
            left=Side(style="thin", color="DDDDDD"),
            right=Side(style="thin", color="DDDDDD"),
            top=Side(style="thin", color="DDDDDD"),
            bottom=Side(style="thin", color="DDDDDD")
        )

    def generate(self, agg_data: dict, output_filepath: str) -> str:
        wb = openpyxl.Workbook()
        
        # Tab 1: Executive Summary
        ws1 = wb.active
        ws1.title = "Executive_Summary"
        self._build_summary_tab(ws1, agg_data)
        
        # Tab 2: ASM Leaderboard
        ws2 = wb.create_sheet("ASM_Leaderboard")
        self._build_leaderboard_tab(ws2, agg_data)
        
        # Tab 3: Store Health Matrix
        ws3 = wb.create_sheet("Store_Health_Matrix")
        self._build_matrix_tab(ws3, agg_data)
        
        # Tab 4: Competitor Survey Analytics
        ws4 = wb.create_sheet("Competitor_Survey")
        self._build_survey_tab(ws4, agg_data)
        
        # Tab 5: CAPA Open Issues
        ws5 = wb.create_sheet("CAPA_Open_Issues")
        self._build_capa_tab(ws5, agg_data)
        
        os.makedirs(os.path.dirname(output_filepath), exist_ok=True)
        wb.save(output_filepath)
        return output_filepath

    def _build_summary_tab(self, ws, data):
        ws.views.sheetView[0].showGridLines = True
        ws.cell(row=2, column=2, value="BÁO CÁO CÔNG TÁC CỬA HÀNG & THỊ TRƯỜNG - BAN GIÁM ĐỐC").font = self.title_font
        ws.cell(row=3, column=2, value=f"Kỳ báo cáo: {data.get('period_name', '')} | Phạm vi: {data.get('asm_filter', 'ALL')}").font = self.subtitle_font
        
        kpis = data.get("kpis", {})
        ws.cell(row=5, column=2, value="CHỈ SỐ TOÀN MẠNG LƯỚI (NETWORK KPIS)").font = Font(name="Calibri", size=12, bold=True, color="1B2A4A")
        
        kpi_headers = ["Lượt kiểm tra", "Điểm Sức khỏe TB", "CH Đạt / Tốt", "CH Chưa Đạt", "Lỗi Nghiêm trọng", "Khảo sát Đối thủ"]
        kpi_values = [
            f"{kpis.get('total_visited', 0)} CH",
            f"{kpis.get('avg_network_score', 0)} / 100",
            f"{kpis.get('good_stores_count', 0) + kpis.get('pass_stores_count', 0)} CH",
            f"{kpis.get('fail_stores_count', 0)} CH",
            f"{kpis.get('critical_violations', 0)} lỗi",
            f"{kpis.get('market_surveys_count', 0)} phiếu"
        ]
        
        for i, (h, v) in enumerate(zip(kpi_headers, kpi_values), start=2):
            cell_h = ws.cell(row=6, column=i, value=h)
            cell_h.fill = self.navy_fill
            cell_h.font = self.header_font
            cell_h.alignment = Alignment(horizontal="center", vertical="center")
            
            cell_v = ws.cell(row=7, column=i, value=v)
            cell_v.font = Font(name="Calibri", size=13, bold=True, color="1B2A4A")
            cell_v.alignment = Alignment(horizontal="center", vertical="center")
            cell_v.border = self.thin_border

        # Top Systemic Issues Table
        ws.cell(row=10, column=2, value="TOP 5 LỖI HỆ THỐNG PHÁT SINH NHIỀU NHẤT").font = Font(name="Calibri", size=12, bold=True, color="1B2A4A")
        c1 = ws.cell(row=11, column=2, value="Hạng mục vi phạm")
        c1.fill = self.navy_fill
        c1.font = self.header_font
        
        c2 = ws.cell(row=11, column=3, value="Số lần phát sinh")
        c2.fill = self.navy_fill
        c2.font = self.header_font
        c2.alignment = Alignment(horizontal="center")
        
        top_issues = data.get("top_systemic_issues", [])
        if not top_issues:
            ws.cell(row=12, column=2, value="Không ghi nhận vi phạm hệ thống.").border = self.thin_border
            ws.cell(row=12, column=3, value="0").border = self.thin_border
        else:
            for idx, issue in enumerate(top_issues, start=12):
                c_lbl = ws.cell(row=idx, column=2, value=issue.get("label", ""))
                c_cnt = ws.cell(row=idx, column=3, value=f"{issue.get('count', 0)} lần")
                c_lbl.border = self.thin_border
                c_cnt.border = self.thin_border
                c_cnt.alignment = Alignment(horizontal="center")
                c_cnt.font = Font(name="Calibri", size=11, bold=True, color="C41E3A")

        self._auto_fit_columns(ws, start_row=5)

    def _build_leaderboard_tab(self, ws, data):
        ws.views.sheetView[0].showGridLines = True
        ws.cell(row=2, column=2, value="BẢNG XẾP HẠNG VẬN HÀNH THEO ASM").font = self.title_font
        
        headers = ["Hạng", "ASM Phụ trách", "CH Phụ trách", "Đã Kiểm tra", "% Tỷ lệ Phủ", "Điểm Sức khỏe TB", "Vi phạm Nghiêm trọng", "% Tỷ lệ Đạt"]
        for c, h in enumerate(headers, start=2):
            cell = ws.cell(row=4, column=c, value=h)
            cell.fill = self.navy_fill
            cell.font = self.header_font
            cell.alignment = Alignment(horizontal="center")
            
        leaderboard = data.get("asm_leaderboard", [])
        if not leaderboard:
            ws.cell(row=5, column=2, value="Chưa có dữ liệu ASM trong kỳ báo cáo này.").border = self.thin_border
        else:
            for r, asm in enumerate(leaderboard, start=5):
                ws.cell(row=r, column=2, value=r-4).alignment = Alignment(horizontal="center")
                ws.cell(row=r, column=3, value=asm.get("asm_name", ""))
                ws.cell(row=r, column=4, value=asm.get("assigned_stores", 0)).alignment = Alignment(horizontal="center")
                ws.cell(row=r, column=5, value=asm.get("visited_stores", 0)).alignment = Alignment(horizontal="center")
                ws.cell(row=r, column=6, value=f"{asm.get('coverage_pct', 0)}%").alignment = Alignment(horizontal="center")
                
                sc = ws.cell(row=r, column=7, value=asm.get("avg_health_score", 0))
                sc.alignment = Alignment(horizontal="center")
                sc.font = Font(name="Calibri", size=11, bold=True)
                
                ws.cell(row=r, column=8, value=asm.get("critical_violations", 0)).alignment = Alignment(horizontal="center")
                ws.cell(row=r, column=9, value=f"{asm.get('pass_rate_pct', 0)}%").alignment = Alignment(horizontal="center")
                
                for c in range(2, 10):
                    ws.cell(row=r, column=c).border = self.thin_border

        self._auto_fit_columns(ws, start_row=4)

    def _build_matrix_tab(self, ws, data):
        ws.views.sheetView[0].showGridLines = True
        ws.cell(row=2, column=2, value="CHI TIẾT MẬT ĐỘ SỨC KHỎE CỬA HÀNG (STORE HEALTH MATRIX)").font = self.title_font
        
        headers = ["Mã CH", "Tên Cửa hàng", "ASM Phụ trách", "Ngày KT", "Số Mục Đạt", "Số Mục Lỗi", "Mục N/A", "Tỷ lệ Đạt Cơ bản", "Lỗi Nghiêm trọng", "Điểm Sức Khỏe", "Đánh Giá"]
        for c, h in enumerate(headers, start=2):
            cell = ws.cell(row=4, column=c, value=h)
            cell.fill = self.navy_fill
            cell.font = self.header_font
            cell.alignment = Alignment(horizontal="center")
            
        matrix = data.get("store_matrix", [])
        if not matrix:
            ws.cell(row=5, column=2, value="Chưa có bản ghi kiểm tra cửa hàng trong kỳ này.").border = self.thin_border
        else:
            for r, row in enumerate(matrix, start=5):
                ws.cell(row=r, column=2, value=row.get("store_code", "")).alignment = Alignment(horizontal="center")
                ws.cell(row=r, column=3, value=row.get("store_name", ""))
                ws.cell(row=r, column=4, value=row.get("asm_name", ""))
                ws.cell(row=r, column=5, value=row.get("report_date", "")).alignment = Alignment(horizontal="center")
                ws.cell(row=r, column=6, value=row.get("passed_items", 0)).alignment = Alignment(horizontal="center")
                ws.cell(row=r, column=7, value=row.get("failed_items", 0)).alignment = Alignment(horizontal="center")
                ws.cell(row=r, column=8, value=row.get("na_items", 0)).alignment = Alignment(horizontal="center")
                ws.cell(row=r, column=9, value=f"{row.get('base_pass_rate', 0)}%").alignment = Alignment(horizontal="center")
                ws.cell(row=r, column=10, value=row.get("critical_violations", 0)).alignment = Alignment(horizontal="center")
                
                sc = ws.cell(row=r, column=11, value=row.get("health_score", 0))
                sc.alignment = Alignment(horizontal="center")
                sc.font = Font(name="Calibri", size=11, bold=True)
                
                st = ws.cell(row=r, column=12, value=row.get("status_label", ""))
                st.alignment = Alignment(horizontal="center")
                if row.get("status_label") == "Tốt":
                    st.fill = self.good_fill
                    st.font = self.good_font
                elif row.get("status_label") == "Đạt":
                    st.fill = self.pass_fill
                    st.font = self.pass_font
                else:
                    st.fill = self.fail_fill
                    st.font = self.fail_font
                    
                for c in range(2, 13):
                    ws.cell(row=r, column=c).border = self.thin_border

        self._auto_fit_columns(ws, start_row=4)

    def _build_survey_tab(self, ws, data):
        ws.views.sheetView[0].showGridLines = True
        ws.cell(row=2, column=2, value="DỮ LIỆU KHẢO SÁT THỊ TRƯỜNG & ĐỐI THỦ CẠNH TRANH").font = self.title_font
        
        headers = ["Mã CH", "Tên Cửa hàng", "Đối thủ", "Tên Chương trình / SP", "Thời gian", "Nội dung Khảo sát", "Ghi chú Khuyến mãi"]
        for c, h in enumerate(headers, start=2):
            cell = ws.cell(row=4, column=c, value=h)
            cell.fill = self.navy_fill
            cell.font = self.header_font
            cell.alignment = Alignment(horizontal="center")
            
        surveys = data.get("market_surveys", [])
        if not surveys:
            ws.cell(row=5, column=2, value="Chưa có dữ liệu khảo sát thị trường.").border = self.thin_border
        else:
            for r, surv in enumerate(surveys, start=5):
                ws.cell(row=r, column=2, value=surv.get("store_code", "")).alignment = Alignment(horizontal="center")
                ws.cell(row=r, column=3, value=surv.get("store_name", ""))
                ws.cell(row=r, column=4, value=surv.get("competitor_name", "---"))
                ws.cell(row=r, column=5, value=surv.get("campaign_name", "---"))
                ws.cell(row=r, column=6, value=surv.get("timestamp", "")).alignment = Alignment(horizontal="center")
                ws.cell(row=r, column=7, value=surv.get("notes", ""))
                ws.cell(row=r, column=8, value=surv.get("promotion_details", ""))
                
                for c in range(2, 9):
                    ws.cell(row=r, column=c).border = self.thin_border

        self._auto_fit_columns(ws, start_row=4)

    def _build_capa_tab(self, ws, data):
        ws.views.sheetView[0].showGridLines = True
        ws.cell(row=2, column=2, value="DANH SÁCH VI PHẠM CẦN KHẮC PHỤC (CAPA ACTION PLAN)").font = self.title_font
        
        headers = ["Mã CH", "Tên Cửa hàng", "ASM", "Hạng mục Vi phạm", "Mức độ", "Người chịu trách nhiệm", "Thời hạn (Deadline)", "Ghi chú Vi phạm"]
        for c, h in enumerate(headers, start=2):
            cell = ws.cell(row=4, column=c, value=h)
            cell.fill = self.navy_fill
            cell.font = self.header_font
            cell.alignment = Alignment(horizontal="center")
            
        matrix = data.get("store_matrix", [])
        row_idx = 5
        has_issues = False
        for store in matrix:
            for issue in store.get("open_issues", []):
                has_issues = True
                ws.cell(row=row_idx, column=2, value=issue.get("store_code", "")).alignment = Alignment(horizontal="center")
                ws.cell(row=row_idx, column=3, value=issue.get("store_name", ""))
                ws.cell(row=row_idx, column=4, value=issue.get("asm_name", ""))
                ws.cell(row=row_idx, column=5, value=issue.get("issue_label", ""))
                
                sev = ws.cell(row=row_idx, column=6, value=issue.get("severity", "Bình thường"))
                sev.alignment = Alignment(horizontal="center")
                if issue.get("severity") in ["Khẩn cấp", "Cao"]:
                    sev.fill = self.fail_fill
                    sev.font = self.fail_font
                    
                ws.cell(row=row_idx, column=7, value=issue.get("assignee", "CHT")).alignment = Alignment(horizontal="center")
                ws.cell(row=row_idx, column=8, value=issue.get("deadline", "---")).alignment = Alignment(horizontal="center")
                ws.cell(row=row_idx, column=9, value=issue.get("note", ""))
                
                for c in range(2, 10):
                    ws.cell(row=row_idx, column=c).border = self.thin_border
                row_idx += 1

        if not has_issues:
            ws.cell(row=5, column=2, value="Không ghi nhận vi phạm tồn đọng chưa giải quyết.").border = self.thin_border

        self._auto_fit_columns(ws, start_row=4)

    def _auto_fit_columns(self, ws, start_row=4):
        """Auto fit column widths, skipping title banner rows (< start_row)"""
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.row >= start_row and cell.value is not None:
                    val_str = str(cell.value)
                    if len(val_str) > max_len:
                        max_len = len(val_str)
            ws.column_dimensions[col_letter].width = max(max_len + 5, 12)
