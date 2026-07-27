# Hướng dẫn thiết kế Google Form & Google Sheet đồng bộ với Python (StoreVisit)

Để Python có thể đọc và đồng bộ hóa tự động dữ liệu từ điện thoại của ASM gửi lên thông qua Google Forms, bạn cần thiết kế Google Form với các câu hỏi và cấu trúc cột trong Google Sheet theo đúng quy chuẩn dưới đây.

---

## 1. Cơ chế ánh xạ thông minh của Python (Fuzzy Matching)

Mã nguồn Python trong [google_sheets_reader.py](file:///c:/All_Report/8_RETAIL_COMMANDER/StoreVisit/data/google_sheets_reader.py) sử dụng thuật toán **khớp từ khóa mờ (Fuzzy Keyword Matching)**. 
* Tức là tiêu đề cột trong Google Sheets **không cần chính xác 100%** so với tên biến, mà chỉ cần **chứa từ khóa chính**.
* Ví dụ: Cột có tiêu đề `"Chọn Mã cửa hàng bạn đến kiểm tra"` vẫn sẽ được nhận diện là cột `store_code` vì chứa cụm từ khóa `"Mã cửa hàng"`.

---

## 2. Bảng quy chuẩn câu hỏi & tiêu đề cột

Hãy thiết lập các câu hỏi trên Google Form (hoặc tiêu đề cột trong Google Sheets) theo bảng sau để Python nhận diện chính xác nhất:

| Biến hệ thống | Từ khóa nhận diện trong tiêu đề (Chỉ cần chứa một trong các cụm này) | Kiểu dữ liệu phù hợp trên Google Form | Ghi chú & Ví dụ giá trị |
| :--- | :--- | :--- | :--- |
| **Mã cửa hàng** | `Mã cửa hàng`, `Ma cua hang`, `Store Code` | **Dropdown (Trình đơn thả xuống)** | Bắt buộc chọn đúng mã viết hoa, ví dụ: `VINCOM`, `CMT8`, `NTT`. |
| **Ngày kiểm tra** | `Ngày kiểm tra`, `Ngay kiem tra`, `Date` | **Date (Ngày)** | Định dạng chuẩn: `DD/MM/YYYY`. |
| **QLKD / ASM** | `QLKD/ASM`, `ASM`, `Người kiểm tra`, `Nguoi kiem tra` | **Short Text hoặc Dropdown** | Tên của ASM đi kiểm tra, ví dụ: `Nguyễn Văn Nam`. |
| **Cửa hàng trưởng** | `Tên CHT`, `Ten CHT`, `Cửa hàng trưởng` | **Short Text** | Họ tên CHT có mặt hôm đó. |
| **Giờ bắt đầu** | `Giờ bắt đầu`, `Gio bat dau`, `Time start` | **Time (Thời gian)** | Định dạng: `HH:MM`, ví dụ: `09:00`. |
| **Giờ kết thúc** | `Giờ kết thúc`, `Gio ket thuc`, `Time end` | **Time (Thời gian)** | Định dạng: `HH:MM`, ví dụ: `11:30`. |
| **Nhân viên có mặt** | `Số NV`, `So NV`, `Nhân viên có mặt` | **Number (Số nguyên)** | Số lượng nhân viên trực ca lúc kiểm tra. |
| **Đánh giá mặt tiền** | `Đánh giá mặt tiền`, `Danh gia mat tien`, `Exterior rating` | **Multiple Choice (Trắc nghiệm)** | Chọn một trong: `Tốt`, `Đạt`, `Chưa đạt`. |
| **Nhận xét mặt tiền** | `Nhận xét mặt tiền`, `Nhan xet mat tien`, `Exterior comments` | **Paragraph (Đoạn văn)** | Ghi chú lỗi nếu có, ví dụ: `Bảng hiệu bám bụi bẩn`. |
| **Đánh giá hàng hóa** | `Đánh giá hàng hóa`, `Danh gia hang hoa`, `Merchandise rating` | **Multiple Choice (Trắc nghiệm)** | Chọn một trong: `Tốt`, `Đạt`, `Chưa đạt`. |
| **Nhận xét hàng hóa** | `Nhận xét hàng hóa`, `Nhan xet hang hoa`, `Merchandise comments` | **Paragraph (Đoạn văn)** | Nhận xét cách trưng bày, gấp ủi... |
| **Đánh giá nhân sự** | `Đánh giá nhân sự`, `Danh gia nhan su`, `Staff rating` | **Multiple Choice (Trắc nghiệm)** | Chọn một trong: `Tốt`, `Đạt`, `Chưa đạt`. |
| **Nhận xét nhân sự** | `Nhận xét nhân sự`, `Nhan xet nhan su`, `Staff comments` | **Paragraph (Đoạn văn)** | Nhận xét thái độ phục vụ, đồng phục... |
| **Đánh giá CSVC** | `Đánh giá CSVC`, `Danh gia csvc`, `CSVC rating` | **Multiple Choice (Trắc nghiệm)** | Chọn một trong: `Tốt`, `Đạt`, `Chưa đạt`. |
| **Nhận xét CSVC** | `Nhận xét CSVC`, `Nhan xet csvc`, `CSVC comments` | **Paragraph (Đoạn văn)** | Nhận xét điều hòa, đèn, POS, điện nước... |
| **Vấn đề tồn đọng** | `Vấn đề tồn đọng`, `Van de ton dong`, `Issues` | **Paragraph (Đoạn văn)** | Vấn đề cần CHT khắc phục gấp. |
| **Kế hoạch khắc phục** | `Kế hoạch khắc phục`, `Ke hoach khac phuc`, `Action plan` | **Paragraph (Đoạn văn)** | Hướng giải quyết lỗi đã thảo luận với CHT. |
| **Thời hạn xử lý** | `Thời hạn xử lý`, `Thoi han xu ly`, `Deadline` | **Date hoặc Short Text** | Thời hạn tối đa cho CHT. |

---

## 3. Cấu hình các câu hỏi tải lên hình ảnh (Photos Upload)

Google Forms hỗ trợ tính năng **File Upload** trực tiếp vào Google Drive. Khi người dùng tải ảnh lên, Google Sheets tự động lưu link liên kết đến ảnh dạng:
`https://drive.google.com/open?id=xxxxxxxxxxxxxxxxxxxxxxxxxxxx`

Để Python có thể tự động tải ảnh này về và chèn vào slide:
1. **Tiêu đề câu hỏi tải ảnh** cần chứa một trong các cụm từ khóa sau:
   * **Ảnh mặt tiền** (Frontage slide): `Ảnh mặt tiền`, `Anh mat tien`, `Exterior photo`
   * **Ảnh trưng bày** (Merchandise slide): `Ảnh hàng hóa`, `Anh hang hoa`, `Merchandise photo`
   * **Ảnh cơ sở vật chất** (CSVC/Outstanding Issues): `Ảnh CSVC`, `Anh CSVC`, `CSVC photo`
   * **Ảnh nhân sự** (Roster / Staff): `Ảnh nhân sự`, `Anh nhan su`, `Staff photo`
2. **Cấu hình giới hạn file trên Google Form**:
   * Thiết lập **Maximum number of files (Số lượng tệp tối đa)**: `2` (Phù hợp với thiết kế của slide báo cáo - tối đa 2 ảnh cho mỗi phần).
   * Cấp quyền **Viewer (Người xem)** cho email Service Account đối với thư mục lưu trữ ảnh trên Google Drive (thư mục này được tạo tự động bởi Google Form khi bạn thêm câu hỏi File Upload).

---

## 4. Thiết lập Cột trạng thái xử lý (`Status`) để quản lý luồng dữ liệu

Để tránh việc sinh lại báo cáo trùng lặp cho những dòng dữ liệu cũ đã xử lý xong:
1. Hãy tạo thêm một cột trống ở cuối cùng của Google Sheet (nếu tự làm thủ công) hoặc Python sẽ tự động chèn thêm cột này ở lần chạy đầu tiên.
2. Tiêu đề cột phải viết chính xác: **`Status`** (phân biệt hoa thường).
3. **Cơ chế hoạt động**:
   * Khi ASM gửi Form mới lên $\rightarrow$ Cột `Status` sẽ trống (hoặc có giá trị mặc định không phải `done`). Python sẽ lấy dòng này về để xử lý.
   * Sau khi Python chạy sinh xong báo cáo PPTX/PDF và xác thực QC thành công $\rightarrow$ Python sẽ tự động gửi lệnh API ghi chữ **`done`** vào cột `Status` của dòng tương ứng.
   * Lần đồng bộ tiếp theo, những dòng có chữ `done` sẽ bị bỏ qua.

---

## 5. Các bước kiểm tra nhanh

Sau khi bạn thiết lập xong Google Sheet và chia sẻ cho email Service Account, hãy kiểm tra tính tương thích bằng cách mở terminal trong thư mục `StoreVisit` và chạy:
```powershell
.venv\Scripts\python.exe setup_google.py --test-connection
```
Lệnh này sẽ quét toàn bộ tiêu đề cột trong Google Sheets của bạn và in ra. Bạn sẽ biết ngay cột nào đã được Python nhận diện đúng, cột nào bị thiếu hoặc sai từ khóa!
