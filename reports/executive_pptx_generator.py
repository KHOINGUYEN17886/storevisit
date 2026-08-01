import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

class ExecutivePPTXGenerator:
    """
    Generates 5-Slide Executive PowerPoint Presentation Deck (.pptx)
    Brand Colors:
      Navy: RGB(27, 42, 74)
      Crimson: RGB(196, 30, 58)
      Gold: RGB(212, 175, 55)
    """
    def __init__(self):
        self.navy = RGBColor(27, 42, 74)
        self.crimson = RGBColor(196, 30, 58)
        self.gold = RGBColor(212, 175, 55)
        self.dark_gray = RGBColor(60, 60, 60)
        self.light_bg = RGBColor(248, 249, 250)

    def generate(self, agg_data: dict, output_filepath: str) -> str:
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        blank_layout = prs.slide_layouts[6] # Blank layout

        # Slide 1: Executive Overview & Network KPIs
        slide1 = prs.slides.add_slide(blank_layout)
        self._build_slide1(slide1, agg_data)

        # Slide 2: Store Health Heatmap & Ranking
        slide2 = prs.slides.add_slide(blank_layout)
        self._build_slide2(slide2, agg_data)

        # Slide 3: Top Systemic Operational Failures
        slide3 = prs.slides.add_slide(blank_layout)
        self._build_slide3(slide3, agg_data)

        # Slide 4: Market Survey & Competitor Landscape
        slide4 = prs.slides.add_slide(blank_layout)
        self._build_slide4(slide4, agg_data)

        # Slide 5: Executive CAPA & Action Plan
        slide5 = prs.slides.add_slide(blank_layout)
        self._build_slide5(slide5, agg_data)

        os.makedirs(os.path.dirname(output_filepath), exist_ok=True)
        prs.save(output_filepath)
        return output_filepath

    def _add_header(self, slide, title_text: str, subtitle_text: str):
        # Header Banner background
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.1))
        shape.fill.solid()
        shape.fill.fore_color.rgb = self.navy
        shape.line.color.rgb = self.navy

        # Title text
        txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.15), Inches(12.333), Inches(0.5))
        tf = txBox.text_frame
        p = tf.paragraphs[0]
        p.text = title_text
        p.font.size = Pt(22)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.font.name = "Calibri"

        # Subtitle text
        p2 = tf.add_paragraph()
        p2.text = subtitle_text
        p2.font.size = Pt(13)
        p2.font.color.rgb = self.gold
        p2.font.name = "Calibri"

    def _build_slide1(self, slide, data):
        self._add_header(slide, "BÁO CÁO CÔNG TÁC CỬA HÀNG & THỊ TRƯỜNG - BAN GIÁM ĐỐC", f"Kỳ báo cáo: {data.get('period_name', '')} | Phạm vi: {data.get('asm_filter', 'ALL')}")

        kpis = data.get("kpis", {})
        cards = [
            ("LƯỢT KIỂM TRA", f"{kpis.get('total_visited', 0)} CH", "Tổng số cửa hàng đã hoàn tất công tác"),
            ("ĐIỂM SỨC KHỎE TB", f"{kpis.get('avg_network_score', 0)} / 100", "Chỉ số sức khỏe vận hành trung bình"),
            ("CỬA HÀNG ĐẠT / TỐT", f"{kpis.get('good_stores_count', 0) + kpis.get('pass_stores_count', 0)} CH", f"Chưa Đạt: {kpis.get('fail_stores_count', 0)} cửa hàng"),
            ("LỖI NGHIÊM TRỌNG", f"{kpis.get('critical_violations', 0)} lỗi", "Lỗi PCCC / An toàn / Thất thoát"),
            ("KHẢO SÁT THỊ TRƯỜNG", f"{kpis.get('market_surveys_count', 0)} phiếu", "Thông tin đối thủ cạnh tranh đã thu thập")
        ]

        left_start = Inches(0.8)
        top_pos = Inches(1.6)
        card_width = Inches(3.6)
        card_height = Inches(2.2)

        for idx, (title, val, sub) in enumerate(cards):
            row = idx // 3
            col = idx % 3
            left = left_start + col * Inches(3.9)
            top = top_pos + row * Inches(2.6)

            card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, card_width, card_height)
            card.fill.solid()
            card.fill.fore_color.rgb = self.light_bg
            card.line.color.rgb = self.navy
            card.line.width = Pt(1.5)

            tf = card.text_frame
            tf.word_wrap = True

            p1 = tf.paragraphs[0]
            p1.text = title
            p1.font.size = Pt(12)
            p1.font.bold = True
            p1.font.color.rgb = self.navy

            p2 = tf.add_paragraph()
            p2.text = val
            p2.font.size = Pt(28)
            p2.font.bold = True
            p2.font.color.rgb = self.crimson if "Nghiêm trọng" in title else self.navy

            p3 = tf.add_paragraph()
            p3.text = sub
            p3.font.size = Pt(11)
            p3.font.color.rgb = self.dark_gray

    def _build_slide2(self, slide, data):
        self._add_header(slide, "XẾP HẠNG BẢN ĐỒ SỨC KHỎE VẬN HÀNH (STORE HEALTH MATRIX)", "Đánh giá sức khỏe vận hành theo Tiêu chuẩn Bán lẻ Quốc tế")

        matrix = data.get("store_matrix", [])[:10] # Top 10
        rows = len(matrix) + 1
        cols = 6

        table_shape = slide.shapes.add_table(rows, cols, Inches(0.8), Inches(1.5), Inches(11.733), Inches(5.2))
        table = table_shape.table

        headers = ["Mã CH", "Tên Cửa Hàng", "ASM Phụ Trách", "Mục Đạt", "Điểm Sức Khỏe", "Đánh Giá"]
        for idx, h in enumerate(headers):
            cell = table.cell(0, idx)
            cell.text = h
            cell.fill.solid()
            cell.fill.fore_color.rgb = self.navy
            for p in cell.text_frame.paragraphs:
                p.font.bold = True
                p.font.color.rgb = RGBColor(255, 255, 255)
                p.font.size = Pt(12)

        for r_idx, store in enumerate(matrix, start=1):
            vals = [
                store.get("store_code", ""),
                store.get("store_name", ""),
                store.get("asm_name", ""),
                f"{store.get('passed_items', 0)} / {store.get('total_applicable', 0)}",
                f"{store.get('health_score', 0)}",
                store.get("status_label", "")
            ]
            for c_idx, val in enumerate(vals):
                cell = table.cell(r_idx, c_idx)
                cell.text = str(val)
                for p in cell.text_frame.paragraphs:
                    p.font.size = Pt(11)
                    if c_idx == 5:
                        p.font.bold = True
                        if val == "Tốt":
                            p.font.color.rgb = RGBColor(34, 139, 34)
                        elif val == "Đạt":
                            p.font.color.rgb = RGBColor(204, 153, 0)
                        else:
                            p.font.color.rgb = self.crimson

    def _build_slide3(self, slide, data):
        self._add_header(slide, "TOP 5 NGUYÊN NHÂN VI PHẠM HỆ THỐNG (SYSTEMIC FAILURES)", "Phân tích các lỗi lặp lại nhiều nhất tại cụm cửa hàng")

        issues = data.get("top_systemic_issues", [])
        txBox = slide.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(11.733), Inches(5.2))
        tf = txBox.text_frame
        
        p = tf.paragraphs[0]
        p.text = "DANH SÁCH HẠNG MỤC VI PHẠM PHÁT SINH NHIỀU NHẤT:"
        p.font.bold = True
        p.font.size = Pt(14)
        p.font.color.rgb = self.navy

        for idx, issue in enumerate(issues, start=1):
            p_issue = tf.add_paragraph()
            p_issue.text = f"{idx}. {issue.get('label', '')} — Phát sinh: {issue.get('count', 0)} lần"
            p_issue.font.size = Pt(13)
            p_issue.font.color.rgb = self.crimson

    def _build_slide4(self, slide, data):
        self._add_header(slide, "KHẢO SÁT THỊ TRƯỜNG & ĐỐI THỦ CẠNH TRANH (MARKET SURVEY)", "Tổng hợp diễn biến khuyến mãi & sản phẩm đối thủ trên địa bàn")

        surveys = data.get("market_surveys", [])[:5]
        rows = len(surveys) + 1 if surveys else 2
        cols = 5

        table_shape = slide.shapes.add_table(rows, cols, Inches(0.8), Inches(1.5), Inches(11.733), Inches(5.2))
        table = table_shape.table

        headers = ["Cửa Hàng", "Đối Thủ Cạnh Tranh", "Chương Trình / Sản Phẩm", "Ngày Khảo Sát", "Ghi Chú Khuyến Mãi"]
        for idx, h in enumerate(headers):
            cell = table.cell(0, idx)
            cell.text = h
            cell.fill.solid()
            cell.fill.fore_color.rgb = self.navy
            for p in cell.text_frame.paragraphs:
                p.font.bold = True
                p.font.color.rgb = RGBColor(255, 255, 255)
                p.font.size = Pt(12)

        if not surveys:
            cell = table.cell(1, 0)
            cell.text = "Chưa có dữ liệu khảo sát thị trường trong kỳ báo cáo này."
        else:
            for r_idx, surv in enumerate(surveys, start=1):
                vals = [
                    surv.get("store_name", surv.get("store_code", "")),
                    surv.get("competitor_name", "---"),
                    surv.get("campaign_name", "---"),
                    surv.get("timestamp", ""),
                    surv.get("notes", "")
                ]
                for c_idx, val in enumerate(vals):
                    cell = table.cell(r_idx, c_idx)
                    cell.text = str(val)
                    for p in cell.text_frame.paragraphs:
                        p.font.size = Pt(11)

    def _build_slide5(self, slide, data):
        self._add_header(slide, "KẾ HOẠCH KHẮC PHỤC KỲ TIẾP THEO (EXECUTIVE CAPA PLAN)", "Phân công trách nhiệm & Thời hạn giải quyết các vấn đề tồn đọng")

        matrix = data.get("store_matrix", [])
        capa_items = []
        for store in matrix:
            for issue in store.get("open_issues", []):
                capa_items.append(issue)
        capa_items = capa_items[:8] # Top 8 open issues

        rows = len(capa_items) + 1 if capa_items else 2
        cols = 5

        table_shape = slide.shapes.add_table(rows, cols, Inches(0.8), Inches(1.5), Inches(11.733), Inches(5.2))
        table = table_shape.table

        headers = ["Cửa Hàng", "Vi Phạm Tồn Đọng", "Mức Độ", "Người Chịu Trách Nhiệm", "Thời Hạn (Deadline)"]
        for idx, h in enumerate(headers):
            cell = table.cell(0, idx)
            cell.text = h
            cell.fill.solid()
            cell.fill.fore_color.rgb = self.navy
            for p in cell.text_frame.paragraphs:
                p.font.bold = True
                p.font.color.rgb = RGBColor(255, 255, 255)
                p.font.size = Pt(12)

        if not capa_items:
            cell = table.cell(1, 0)
            cell.text = "Không có vi phạm tồn đọng chưa giải quyết."
        else:
            for r_idx, issue in enumerate(capa_items, start=1):
                vals = [
                    issue.get("store_name", issue.get("store_code", "")),
                    issue.get("issue_label", ""),
                    issue.get("severity", "Bình thường"),
                    issue.get("assignee", "CHT"),
                    issue.get("deadline", "---")
                ]
                for c_idx, val in enumerate(vals):
                    cell = table.cell(r_idx, c_idx)
                    cell.text = str(val)
                    for p in cell.text_frame.paragraphs:
                        p.font.size = Pt(11)
                        if c_idx == 2 and val in ["Khẩn cấp", "Cao"]:
                            p.font.color.rgb = self.crimson
                            p.font.bold = True
