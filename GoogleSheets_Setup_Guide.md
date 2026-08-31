# Hướng Dẫn Cấu Hình Bảng Tính Google Sheets (StoreVisit Enterprise Hub)

Tài liệu này quy chuẩn cấu trúc toàn diện các tab trong Google Spreadsheet trung tâm (`SPREADSHEET_ID`), phục vụ đồng bộ dữ liệu 2 chiều giữa WebApp (Vercel/GAS), Google Drive và Python Desktop Command Center.

---

## 1. CẤU TRÚC CÁC TAB CHÍNH (WORKSHEETS ARCHITECTURE)

| Tên Tab (Worksheet) | Mục Đích Sử Dụng | Quyền Ghi / Đọc |
| :--- | :--- | :--- |
| **`Form Responses 1`** | Lưu trữ toàn bộ các phiếu kiểm tra hiện trường gửi từ WebApp | WebApp (Ghi), Python (Đọc & Ghi `Status=done`) |
| **`Draft_StoreVisits`** | Lưu trữ bộ nhớ đệm bản nháp đám mây (Cloud Draft Sync) đa thiết bị | WebApp (Đọc & Ghi theo `key_id`) |
| **`Issues_Register`** | Sổ đăng ký theo dõi và quản lý vòng đời khắc phục lỗi (CAPA) | WebApp & Python |
| **`StoreMapping`** | Danh bạ chuẩn 185 Cửa hàng & 11 ASM phân quyền | Master Sync từ `StoresInfo.xlsx` |
| **`Users`** | Bảng tài khoản, mật khẩu băm, vai trò và phân quyền vùng | Hệ thống xác thực WebApp |

---

## 2. QUY CHUẨN CẤU TRÚC TAB `Draft_StoreVisits` (CLOUD DRAFT SYNC)

Để hỗ trợ tính năng đồng bộ bản nháp 2 chiều giữa Điện thoại và Máy tính (PC $\leftrightarrow$ Mobile), tab `Draft_StoreVisits` được thiết lập với 7 cột chuẩn:

| Cột | Tên Tiêu Đề | Kiểu Dữ Liệu | Mô Tả & Ví Dụ |
| :---: | :--- | :--- | :--- |
| **A** | `key_id` | Text (Unique Key) | Khóa định danh bản nháp: `{username}::{store_code}` (ví dụ: `khoi::so1`, `tien::diamond`) |
| **B** | `username` | Text | Tên đăng nhập của người tạo nháp (ví dụ: `khoi`, `tien`) |
| **C** | `store_code` | Text | Mã cửa hàng đang khảo sát (ví dụ: `SO1`, `NTT`) |
| **D** | `asm_name` | Text | Tên ASM phụ trách (ví dụ: `Nguyễn Đăng Khôi`) |
| **E** | `report_date` | Text / Date | Ngày kiểm tra (ví dụ: `2026-08-31`) |
| **F** | `updated_at` | Text (ISO 8601) | Timestamp lần lưu gần nhất (ví dụ: `2026-08-31T16:15:02.123Z`) |
| **G** | `draft_json` | Long Text (JSON) | Chuỗi JSON chứa toàn bộ dữ liệu form khảo sát đã điền |

---

## 3. QUY CHUẨN CẤU TRÚC TAB `Form Responses 1` (SUBMISSION DATA)

Dữ liệu gửi từ 3 chế độ kiểm tra (`full_audit`, `quick_pulse`, `target_rescue`) đều được chuẩn hóa ghi nhận vào tab này:

| Tên Cột Chuẩn | Từ Khóa Nhận Diện (Fuzzy Keyword) | Kiểu Dữ Liệu | Ghi Chú |
| :--- | :--- | :--- | :--- |
| **Timestamp** | `Timestamp`, `Dấu thời gian` | Date Time | Thời điểm gửi báo cáo |
| **Mã gửi (Submission ID)** | `submission_id`, `Mã gửi` | Text (UUID) | Khóa duy nhất của mỗi lượt gửi |
| **Chế độ kiểm tra** | `inspection_profile`, `Chế độ` | Text | `full_audit`, `quick_pulse`, `target_rescue` |
| **Mã cửa hàng** | `store_code`, `Mã cửa hàng` | Text | Mã viết hoa (ví dụ: `SO1`, `NTT`, `CMT8`) |
| **Ngày kiểm tra** | `report_date`, `Ngày kiểm tra` | Date | `YYYY-MM-DD` |
| **QLKD / ASM** | `asm_name`, `ASM` | Text | Tên người kiểm tra |
| **Đánh giá mặt tiền** | `rating_frontage` | Text | `Đạt` / `Không đạt` |
| **Đánh giá không gian trong**| `rating_inner` | Text | `Đạt` / `Không đạt` |
| **Đánh giá trưng bày AP** | `rating_merch_ap` | Text | `Đạt` / `Không đạt` / `Không áp dụng` |
| **Đánh giá trưng bày PIE** | `rating_merch_pie` | Text | `Đạt` / `Không đạt` / `Không áp dụng` |
| **Đánh giá nhân sự** | `rating_staff` | Text | `Đạt` / `Không đạt` |
| **Đánh giá CSVC & PCCC** | `rating_csvc` | Text | `Đạt` / `Không đạt` |
| **Khảo sát đối thủ** | `survey_competitor` | Text (JSON) | Thông tin giá, CTKM, sản phẩm đối thủ |
| **Kế hoạch hành động** | `action_plan` | Text | Cam kết khắc phục của CHT/ASM |
| **Hạn chót khắc phục** | `action_deadline` | Date | Hạn chót xử lý lỗi |
| **Checklist Chi Tiết** | `checklist_json` | Long Text (JSON) | Toàn bộ 65 tiêu chí chi tiết kèm link ảnh Drive |
| **Trạng Thái Xử Lý** | `Status` | Text | `done` (khi Python đã sinh xong báo cáo) |

---

## 4. QUY TRÌNH KIỂM TRA ĐỒNG BỘ PYTHON

Mở terminal trong thư mục dự án và chạy lệnh kiểm thử kết nối Google API:
```powershell
.venv\Scripts\python.exe setup_google.py --test-connection
```
Lệnh này sẽ quét toàn bộ tiêu đề cột trong Google Sheets và đối soát với hệ thống Data Loader của Python!
