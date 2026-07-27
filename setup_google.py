import os
import sys
import yaml
import json
import argparse
from datetime import datetime

def print_walkthrough():
    instructions = """
========================================================================
HƯỚNG DẪN SETUP KẾT NỐI GOOGLE SHEETS & DRIVE CHO RETAIL COMMANDER
========================================================================

BƯỚC 1: TẠO GOOGLE SERVICE ACCOUNT TRÊN GOOGLE CLOUD
------------------------------------------------------------------------
1. Truy cập: https://console.cloud.google.com
2. Tạo một Project mới (ví dụ: "RetailCommanderStoreVisit").
3. Vào "APIs & Services" > "Library".
4. Tìm kiếm và click ENABLE 2 thư viện sau:
   - Google Sheets API
   - Google Drive API
5. Vào "IAM & Admin" > "Service Accounts" > click "CREATE SERVICE ACCOUNT".
6. Nhập tên tài khoản, click "CREATE AND CONTINUE", sau đó click "DONE".
7. Tại danh sách Service Accounts, click vào email của tài khoản vừa tạo.
8. Chọn Tab "KEYS" > "ADD KEY" > "Create new key" > Chọn định dạng "JSON" > click "CREATE".
9. Lưu tệp JSON tải về và đổi tên thành:
   'C:\\All_Report\\8_RETAIL_COMMANDER\\StoreVisit\\config\\google_credentials.json'

BƯỚC 2: CHIA SẺ QUYỀN TRUY CẬP GOOGLE SHEETS & DRIVE
------------------------------------------------------------------------
1. Mở file Google Sheets chứa phản hồi từ Google Forms.
2. Click "Share" (Chia sẻ) ở góc phải màn hình.
3. Paste địa chỉ Email của Service Account vào (email này nằm trong file credentials.json, dạng 'xxxx@xxxx.iam.gserviceaccount.com').
4. Cấp quyền chỉnh sửa (Editor / Người chỉnh sửa) và nhấn "Gửi" (Send).
5. Làm tương tự: Nếu trong Google Sheets có cột ảnh upload từ Drive, hãy chia sẻ thư mục lưu trữ ảnh trên Google Drive cho email Service Account với quyền "Viewer" (Người xem) để app có thể tải ảnh về.

BƯỚC 3: CẤU HÌNH SPREADSHEET ID TRONG APP_CONFIG.YAML
------------------------------------------------------------------------
1. Copy ID của Google Sheets từ thanh địa chỉ trình duyệt:
   Ví dụ: link là https://docs.google.com/spreadsheets/d/1AbCdEfGhIjKlMnOpQrStUvWxYz/edit#gid=0
   ID sẽ là: '1AbCdEfGhIjKlMnOpQrStUvWxYz'
2. Mở file 'C:\\All_Report\\8_RETAIL_COMMANDER\\StoreVisit\\config\\app_config.yaml'
3. Tìm phần 'google:' và điền ID vào dòng 'spreadsheet_id':
   google:
     spreadsheet_id: "1AbCdEfGhIjKlMnOpQrStUvWxYz"

------------------------------------------------------------------------
Sau khi thực hiện, bạn có thể chạy kiểm tra kết nối bằng lệnh:
  .\\.venv\\Scripts\\python.exe setup_google.py --test-connection

Hoặc test schema Khảo sát thị trường:
  .\\.venv\\Scripts\\python.exe setup_google.py --test-connection --schema market_survey

Hoặc tạo dữ liệu giả lập (mock data) để test offline ngay lập tức:
  .\\.venv\\Scripts\\python.exe setup_google.py --mock-data
  .\\.venv\\Scripts\\python.exe setup_google.py --mock-data --schema market_survey
========================================================================
"""
    print(instructions)

def test_connection(schema: str = "store_visit", worksheet_name: str = None):
    root_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(root_dir, "config/app_config.yaml")
    
    if not os.path.exists(config_path):
        print("Lỗi: Không tìm thấy file config/app_config.yaml")
        return
        
    with open(config_path, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)
        
    google_config = config.get("google", {})
    creds_path = google_config.get("credentials_path", "")
    sheet_id = google_config.get("spreadsheet_id", "")
    
    if schema == "market_survey":
        sheet_name = worksheet_name or google_config.get("survey_sheet_name", "MarketSurvey_Responses")
    else:
        sheet_name = worksheet_name or google_config.get("sheet_name", "Form Responses 1")
    
    if not os.path.isabs(creds_path):
        creds_path = os.path.join(root_dir, creds_path)
        
    print(f"Kiểm tra credentials tại: {creds_path}")
    if not os.path.exists(creds_path):
        print("❌ Lỗi: Chưa tìm thấy file google_credentials.json. Vui lòng làm theo Bước 1.")
        return
        
    print(f"Kiểm tra spreadsheet ID: {sheet_id}")
    if not sheet_id:
        print("❌ Lỗi: spreadsheet_id trong app_config.yaml đang để trống. Vui lòng làm theo Bước 3.")
        return
        
    print(f"Đang kết nối tới Google Sheets API (Schema: {schema})...")
    try:
        import gspread
        from google.oauth2.service_account import Credentials
        
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
        client = gspread.authorize(creds)
        
        doc = client.open_by_key(sheet_id)
        print(f"✅ Kết nối thành công tới Spreadsheet: '{doc.title}'")
        
        worksheets = doc.worksheets()
        print("Danh sách các sheet con:")
        for ws in worksheets:
            print(f"  - {ws.title} (Số dòng: {ws.row_count}, Số cột: {ws.col_count})")
            
        try:
            target_ws = doc.worksheet(sheet_name)
            print(f"✅ Đã tìm thấy target worksheet: '{sheet_name}'")
            headers = [h.strip() for h in target_ws.row_values(1)]
            print(f"Headers trong sheet: {headers}")
            
            print("\nKiểm tra cấu trúc tương thích cột:")
            if schema == "market_survey":
                columns_map = {
                    "store_code": ["[MS01]", "Mã cửa hàng", "Ma cua hang", "Store Code"],
                    "region": ["[MS02]", "Khu vực", "Cum cua hang", "Cụm cửa hàng", "Region"],
                    "qlkd_asm": ["[MS03]", "QLKD/ASM", "ASM phụ trách", "ASM phu trach", "Người phụ trách"],
                    "respondent_name": ["[MS04]", "Người đại diện", "Nguoi dai dien", "Họ tên người trả lời"],
                    "respondent_role": ["[MS05]", "Chức danh", "Chuc danh", "Respondent role"],
                    "discussion_count": ["[MS06]", "Số người", "So nguoi", "Discussion count"],
                    "survey_date": ["[MS07]", "Ngày thực hiện", "Ngay thuc hien", "Survey date"],
                    "customer_change": ["[MS11]", "Khách hàng đang có thay đổi", "Khach hang dang co thay doi", "Customer change"],
                    "demand_increase": ["[MS15]", "những nhu cầu nào đang tăng", "nhung nhu cau nao dang tang", "demand increase"],
                    "lost_sale_reasons": ["[MS21]", "nguyên nhân khách không mua", "nguyen nhan khach khong mua", "lost sale reasons"],
                    "lost_sale_top1": ["[MS22]", "nguyên nhân lớn nhất", "nguyen nhan quan trong nhat", "lost sale top 1"],
                    "product_gap": ["[MS31]", "nhóm sản phẩm cần bổ sung", "nhom san pham can bo sung", "product gap"],
                    "acceptable_price": ["[MS37]", "khoảng giá khách chấp nhận", "khoang gia khach chap nhan", "acceptable price"],
                    "support_categories": ["[MS41]", "nhóm cần công ty hỗ trợ", "nhom can cong ty ho tro", "support categories"],
                    "suggested_solution": ["[MS45]", "giải pháp cửa hàng đề xuất", "giai phap cua hang de xuat", "suggested solution"],
                    "support_photos": ["[MS48]", "ảnh minh chứng", "anh minh chung", "support photos", "support photo"],
                    "local_opportunity": ["[MS51]", "mùa vụ hoặc cơ hội", "mua vu hoac co hoi", "local opportunity"],
                    "need_before_date": ["[MS59]", "thời điểm hàng cần có", "thoi diem hang can co", "need before date"],
                    "store_recommendation": ["[MS61]", "kiến nghị ưu tiên hàng đầu", "kien nghi uu tien hang dau", "store recommendation"],
                }
            else:
                columns_map = {
                    "store_code": ["Mã cửa hàng", "Ma cua hang", "Store Code"],
                    "report_date": ["Ngày kiểm tra", "Ngay kiem tra", "Date"],
                    "asm_name": ["QLKD/ASM", "ASM", "Người kiểm tra", "Nguoi kiem tra"],
                    "cht_name": ["Tên CHT", "Ten CHT", "Cửa hàng trưởng"],
                    "time_start": ["Giờ bắt đầu", "Gio bat dau", "Time start"],
                    "time_end": ["Giờ kết thúc", "Gio ket thuc", "Time end"],
                    "nv_count": ["Số NV", "So NV", "Nhân viên có mặt"],
                    "rating_frontage": ["Đánh giá mặt tiền", "Danh gia mat tien", "Exterior rating"],
                    "comment_frontage": ["Nhận xét mặt tiền", "Nhan xet mat tien", "Exterior comments"],
                    "photos_frontage": ["Ảnh mặt tiền", "Anh mat tien", "Exterior photo"],
                    "pending_issues": ["Vấn đề tồn đọng", "Van de ton dong", "Issues"],
                    "action_plan": ["Kế hoạch khắc phục", "Ke hoach khac phuc", "Action plan"],
                    "action_deadline": ["Thời hạn xử lý", "Thoi han xu ly", "Deadline"]
                }
                
            has_errors = False
            for key, keywords in columns_map.items():
                found = False
                matched_col = ""
                if schema == "market_survey" and keywords[0].startswith("["):
                    tech_code = keywords[0]
                    for h in headers:
                        if tech_code in h:
                            found = True
                            matched_col = h
                            break
                if not found:
                    for h in headers:
                        h_lower = h.lower()
                        kws_to_check = keywords[1:] if (schema == "market_survey") else keywords
                        if any(kw.lower() in h_lower for kw in kws_to_check):
                            found = True
                            matched_col = h
                            break
                            
                if found:
                    print(f"  [OK] {key} -> '{matched_col}'")
                else:
                    is_required = key in ["store_code", "qlkd_asm", "asm_name", "survey_date", "report_date"]
                    status_str = "[MISSING]" if is_required else "[WARNING]"
                    if is_required:
                        has_errors = True
                    print(f"  {status_str} Không tìm thấy tiêu đề phù hợp cho biến '{key}' (Từ khóa: {keywords})")
                    
            if "Status" in headers:
                print("  [OK] Status column exists")
            else:
                print("  [WARNING] Cột quản lý 'Status' chưa tồn tại. Python sẽ tự động tạo ở lần chạy đầu tiên.")
                
            if has_errors:
                print("\n❌ Cảnh báo: Biểu mẫu thiếu một số trường thông tin bắt buộc!")
            else:
                print("\n✅ Cấu hình biểu mẫu hợp lệ, sẵn sàng kết nối!")
                
        except gspread.WorksheetNotFound:
            print(f"❌ Cảnh báo: Không tìm thấy sheet con tên '{sheet_name}'.")
            
    except Exception as e:
        print("❌ Lỗi kết nối Google API:")
        print(f"  Chi tiết: {e}")

def create_mock_data(schema: str = "store_visit"):
    root_dir = os.path.dirname(os.path.abspath(__file__))
    
    if schema == "market_survey":
        cache_path = os.path.join(root_dir, "data/survey_cache.json")
        os.makedirs(os.path.dirname(cache_path), exist_ok=True)
        
        photo_cache_dir = os.path.join(root_dir, "temp/survey_photos")
        os.makedirs(photo_cache_dir, exist_ok=True)
        
        mock_paths = []
        try:
            from PIL import Image, ImageDraw
            for name in ["mock_survey_1", "mock_survey_2"]:
                p = os.path.join(photo_cache_dir, f"{name}.jpg")
                img = Image.new("RGB", (800, 600), color=(10, 35, 66))
                d = ImageDraw.Draw(img)
                d.text((100, 250), f"MOCK SURVEY IMAGE: {name.upper()}", fill=(255, 255, 255))
                img.save(p)
                mock_paths.append(p)
            print("✅ Đã tạo 2 ảnh khảo sát giả lập trong temp/survey_photos")
        except Exception as e:
            print(f"Cảnh báo tạo ảnh: {e}. Sẽ dùng fallback placeholder.")
            
        mock_responses = {
            "2001": {
                "response_id": "2001",
                "store_code": "VINCOM",
                "region": "Hồ Chí Minh",
                "qlkd_asm": "Nguyễn Văn Nam (ASM Mock)",
                "respondent_name": "Trần Thị CHT (Vincom)",
                "respondent_role": "Cửa hàng trưởng",
                "discussion_count": 3,
                "survey_date": datetime.now().strftime("%d/%m/%Y"),
                "customer_change": "Khách hàng phân khúc cao cấp hỏi nhiều về dòng suit Pierre Cardin chất liệu nhẹ mát cho mùa hè.",
                "demand_increase": ["Pierre Cardin Suit", "Light Fabric Shirts"],
                "lost_sale_reasons": ["Thiếu size", "Chưa đúng mùa vụ"],
                "lost_sale_top1": "Thiếu size",
                "product_gap": ["Bộ vest Pierre Cardin siêu nhẹ"],
                "acceptable_price": "1.200.000–2.000.000 đồng",
                "support_categories": ["Sản phẩm", "Truyền thông địa phương"],
                "suggested_solution": "Bổ sung gấp các mẫu vest hè mỏng, chạy quảng cáo Target khu vực Quận 1.",
                "photos": [
                    {
                        "index": 1,
                        "drive_url": "https://drive.google.com/open?id=mock_survey_id_1",
                        "local_path": mock_paths[0] if len(mock_paths) > 0 else ""
                    }
                ],
                "local_opportunity": "Lễ hội mùa hè trung tâm thương mại Vincom Q.1",
                "need_before_date": "01/07/2026",
                "store_recommendation": "Cần ưu tiên phân bổ size L và XL cho Vest Pierre Cardin vì lượng khách trung niên thể hình lớn tăng cao.",
                "status": "new",
                "qc_status": "approved"
            },
            "2002": {
                "response_id": "2002",
                "store_code": "CMT8",
                "region": "Hồ Chí Minh",
                "qlkd_asm": "Nguyễn Văn Nam (ASM Mock)",
                "respondent_name": "Lê Văn CHT (CMT8)",
                "respondent_role": "Cửa hàng trưởng",
                "discussion_count": 2,
                "survey_date": datetime.now().strftime("%d/%m/%Y"),
                "customer_change": "Khách hỏi nhiều về quần short và áo thun polo trẻ trung đi du lịch hè.",
                "demand_increase": ["Quần short", "Áo thun Polo"],
                "lost_sale_reasons": ["Không có mẫu phù hợp", "Thiếu màu"],
                "lost_sale_top1": "Không có mẫu phù hợp",
                "product_gap": ["Polo họa tiết năng động"],
                "acceptable_price": "500.000–800.000 đồng",
                "support_categories": ["Sản phẩm", "Công cụ bán hàng"],
                "suggested_solution": "Thêm mẫu short kaki co giãn nhiều màu sắc.",
                "photos": [],
                "local_opportunity": "Đợt du lịch hè gia đình đầu tháng 7",
                "need_before_date": "25/06/2026",
                "store_recommendation": "Đề xuất cho CMT8 thêm các mã short màu sáng năng động.",
                "status": "new",
                "qc_status": "approved"
            }
        }
        
        with open(cache_path, "w", encoding="utf-8") as f:
            json.dump(mock_responses, f, indent=2, ensure_ascii=False)
            
        print(f"✅ Đã ghi dữ liệu khảo sát test giả lập vào: {cache_path}")
        return
        
    cache_path = os.path.join(root_dir, "data/form_cache.json")
    os.makedirs(os.path.dirname(cache_path), exist_ok=True)
    
    photo_cache_dir = os.path.join(root_dir, "temp/drive_photos")
    os.makedirs(photo_cache_dir, exist_ok=True)
    
    mock_paths = []
    try:
        from PIL import Image, ImageDraw
        for name in ["mock_front_1", "mock_front_2", "mock_merch_1"]:
            p = os.path.join(photo_cache_dir, f"{name}.jpg")
            img = Image.new("RGB", (800, 600), color=(10, 35, 66))
            d = ImageDraw.Draw(img)
            d.text((100, 250), f"MOCK IMAGE: {name.upper()}", fill=(255, 255, 255))
            img.save(p)
            mock_paths.append(p)
        print("✅ Đã tạo 3 ảnh giả lập trong temp/drive_photos")
    except Exception as e:
        print(f"Cảnh báo tạo ảnh: {e}. Sẽ dùng fallback placeholder.")
        
    mock_responses = {
        "1002": {
            "response_id": "1002",
            "store_code": "VINCOM",
            "report_date": datetime.now().strftime("%d/%m/%Y"),
            "asm_name": "Nguyễn Văn Nam (ASM Mock)",
            "cht_name": "Trần Thị CHT (Vincom)",
            "time_start": "09:00",
            "time_end": "11:00",
            "nv_count": 4,
            "rating_frontage": "Chưa đạt",
            "rating_merch": "Đạt",
            "rating_staff": "Đạt",
            "rating_csvc": "Đạt",
            "comment_frontage": "Mặt tiền bám bụi nhiều tại vách kính bên trái. Đèn chữ P bị chập chờn nhẹ -> ASM đã nhắc nhở CHT vệ sinh và gọi PTTT thay bóng.",
            "comment_merch": "Trưng bày sảnh chính gọn gàng, đúng layout Pierre Cardin. Tuy nhiên kệ phụ kiện hơi thưa hàng.",
            "comment_staff": "Nhân viên mặc đúng đồng phục, chào hỏi khách hàng tốt.",
            "comment_csvc": "Hệ thống POS chạy tốt. Cửa kính bên trái bị kẹt nhẹ khi đẩy.",
            "pending_issues": "Kính cửa kẹt khó đẩy & Đèn chữ P bảng hiệu bị chập chờn.",
            "action_plan": "Liên hệ PTTT sửa bản lề kính & thay bóng đèn chữ P.",
            "action_deadline": "20/06/2026",
            "photos": [
                {
                    "section": "frontage",
                    "index": 1,
                    "drive_url": "https://drive.google.com/open?id=mock_file_id_1",
                    "local_path": mock_paths[0] if len(mock_paths) > 0 else ""
                },
                {
                    "section": "frontage",
                    "index": 2,
                    "drive_url": "https://drive.google.com/open?id=mock_file_id_2",
                    "local_path": mock_paths[1] if len(mock_paths) > 1 else ""
                },
                {
                    "section": "merchandise",
                    "index": 1,
                    "drive_url": "https://drive.google.com/open?id=mock_file_id_3",
                    "local_path": mock_paths[2] if len(mock_paths) > 2 else ""
                }
            ],
            "status": "pending"
        },
        "1003": {
            "response_id": "1003",
            "store_code": "CMT8",
            "report_date": datetime.now().strftime("%d/%m/%Y"),
            "asm_name": "Nguyễn Văn Nam (ASM Mock)",
            "cht_name": "Lê Văn CHT (CMT8)",
            "time_start": "13:30",
            "time_end": "15:00",
            "nv_count": 3,
            "rating_frontage": "Tốt",
            "rating_merch": "Tốt",
            "rating_staff": "Đạt",
            "rating_csvc": "Chưa đạt",
            "comment_frontage": "Mặt tiền trang trí lộng lẫy, sạch đẹp, thu hút.",
            "comment_merch": "Hàng hóa đầy ắp, ủi phẳng phiu, lên kệ đúng màu sắc.",
            "comment_staff": "Tác phong chuyên nghiệp.",
            "comment_csvc": "Máy lạnh sảnh chính bị rò rỉ nước tại góc trong cùng bên phải -> Cần kỹ thuật khắc phục ngay tránh ướt sàn gỗ.",
            "pending_issues": "Rò rỉ nước điều hòa sảnh chính.",
            "action_plan": "Gọi thợ bảo dưỡng máy lạnh đến thông ống thoát nước thải.",
            "action_deadline": "17/06/2026",
            "photos": [
                {
                    "section": "frontage",
                    "index": 1,
                    "drive_url": "https://drive.google.com/open?id=mock_file_id_4",
                    "local_path": mock_paths[0] if len(mock_paths) > 0 else ""
                }
            ],
            "status": "pending"
        }
    }
    
    with open(cache_path, "w", encoding="utf-8") as f:
        json.dump(mock_responses, f, indent=2, ensure_ascii=False)
        
    print(f"✅ Đã ghi dữ liệu test giả lập vào: {cache_path}")
    print("Mở app StoreVisit lên sẽ thấy 2 dòng phản hồi giả lập để test ngay!")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Google Sheets Setup & Verification Utility")
    parser.add_argument("--test-connection", action="store_true", help="Test connection to Google Sheets API")
    parser.add_argument("--mock-data", action="store_true", help="Generate mock form response cache for offline testing")
    parser.add_argument("--schema", type=str, default="store_visit", choices=["store_visit", "market_survey"], help="The schema to verify/mock")
    parser.add_argument("--worksheet", type=str, default=None, help="Optional custom worksheet name")
    
    args = parser.parse_args()
    if args.test_connection:
        test_connection(schema=args.schema, worksheet_name=args.worksheet)
    elif args.mock_data:
        create_mock_data(schema=args.schema)
    else:
        print_walkthrough()
