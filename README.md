# Banana SmartLink MVP (SmartLink Portal)

**Banana SmartLink Portal** là giải pháp Cổng tổng hợp số liệu và chỉ tiêu Kinh tế - Xã hội dành cho Ủy ban Nhân dân cấp xã (hiện tại đang áp dụng mẫu cho **UBND Xã Bà Nà**). Hệ thống hỗ trợ cán bộ xã thu thập, tự động kiểm định (validate), theo dõi tiến độ nộp báo cáo và tổng hợp dữ liệu từ các thôn một cách thông minh, nhanh chóng và chính xác.

Hệ thống được phát triển bằng ngôn ngữ **Python** trên nền tảng web nhẹ **Streamlit** kết hợp cơ sở dữ liệu **SQLite**.

---

## Các Tính Năng Chính

1. **Dashboard Trực Quan (CT01 - CT14)**:
   - Theo dõi chi tiết 14 chỉ tiêu kinh tế - xã hội quan trọng theo từng kỳ (quý, tháng) và từng thôn.
   - So sánh trực quan dữ liệu giữa các thôn bằng biểu đồ động (Altair).
   - Tính toán nhanh tổng số liệu toàn xã ngay khi thay đổi bộ lọc.

2. **Tự Động Nhập Liệu & Chuẩn Hóa**:
   - Hỗ trợ tải lên file báo cáo riêng lẻ của từng thôn (dạng cột dọc) hoặc file tổng hợp tiến độ chung (dạng bảng ngang).
   - Tự động bóc tách thông tin xã, thôn, người lập báo cáo, số điện thoại, thời điểm nộp và kỳ báo cáo (hỗ trợ chuyển đổi ngôn ngữ/định dạng kỳ báo cáo tiếng Việt linh hoạt như: *"Quý II 2026"*, *"Tháng 06/2026"* thành dạng chuẩn `YYYY-Qx` hoặc `YYYY-MM`).

3. **Kiểm Định Dữ Liệu Tự Động (Data Validation)**:
   - Kiểm tra các lỗi nghiêm trọng (`BLOCKER`) như thiếu các trường bắt buộc (Tên thôn, Kỳ báo cáo, Tổng số hộ, Nhân khẩu, Số hộ nghèo/cận nghèo).
   - Phát hiện các cảnh báo (`WARNING`) như số điện thoại không hợp lệ, dữ liệu trống hoặc bất thường ở các chỉ tiêu CT05 - CT14.
   - Nhật ký lỗi chi tiết hiển thị trực tiếp cho người dùng sửa đổi trước khi lưu chính thức vào cơ sở dữ liệu.

4. **Quản Lý Tiến Độ & Nhắc Nhở**:
   - Giám sát thời gian nộp báo cáo của từng thôn so với hạn nộp (Due Date).
   - Phân loại trạng thái nộp tự động: *Đúng hạn, Nộp trễ (tính rõ số ngày trễ), Chưa nộp*.

5. **Xuất Báo Cáo Chuyên Nghiệp**:
   - Xuất dữ liệu tổng hợp ra file Excel (`.xlsx`) chuẩn định dạng để lưu trữ.
   - Xuất văn bản báo cáo hành chính tổng hợp ra file Word (`.docx`) theo đúng thể thức văn bản hành chính Việt Nam.

6. **Trợ Lý Phân Tích Dữ Liệu AI**:
   - Tích hợp khung Chatbot AI (nút nổi góc màn hình) hỗ trợ cán bộ xã truy vấn dữ liệu nhanh bằng ngôn ngữ tự nhiên, phân tích chất lượng số liệu và đưa ra nhận định kinh tế - xã hội của địa phương.

---

## Cấu Trúc Thư Mục Dự Án

```text
banana_smartlink_mvp/
├── .streamlit/             # Cấu hình giao diện và giao diện người dùng của Streamlit
├── data/                   # Thư mục lưu trữ database SQLite (banana_smartlink.db)
├── exports/                # Thư mục chứa các file báo cáo Excel/Word được xuất ra
├── sample_data/            # Dữ liệu mẫu phục vụ chạy thử nghiệm
│   └── drive_imported/     # Các file báo cáo tải xuống từ Drive và file theo dõi tiến độ
├── templates/              # Thư mục chứa biểu mẫu báo cáo Excel chuẩn (.xlsx)
├── app.py                  # Điểm chạy chính (Streamlit UI, Navigation & AI Chat)
├── database.py             # Quản trị SQLite DB (khởi tạo, truy vấn, lưu trữ dữ liệu)
├── schema.py               # Định nghĩa cấu trúc chỉ tiêu (CT01-CT14) & trường dữ liệu bắt buộc
├── validators.py           # Bộ thư viện chuẩn hóa văn bản, xử lý ngày tháng & kiểm định Excel
├── report_generator.py     # Module phụ trách tạo file xuất Excel và Word
├── requirements.txt        # Danh sách các thư viện Python phụ thuộc
├── run_app.bat             # File script chạy nhanh ứng dụng trên Windows
└── run_app.sh              # File script chạy nhanh ứng dụng trên Linux / macOS
```

---

## Danh Sách 14 Chỉ Tiêu Kinh Tế - Xã Hội (CT01 - CT14)

| Mã Chỉ Tiêu | Tên Chỉ Tiêu | Đơn Vị Tính |
| :--- | :--- | :---: |
| **CT01** | Tổng số hộ dân | Hộ |
| **CT02** | Tổng số nhân khẩu | Người |
| **CT03** | Số hộ nghèo | Hộ |
| **CT04** | Số hộ cận nghèo | Hộ |
| **CT05** | Số người có công với cách mạng | Người |
| **CT06** | Số đối tượng bảo trợ xã hội đang hưởng trợ cấp | Người |
| **CT07** | Số trẻ em dưới 16 tuổi | Người |
| **CT08** | Số trẻ em có hoàn cảnh đặc biệt | Người |
| **CT09** | Số hộ đạt Gia đình văn hóa | Hộ |
| **CT10** | Số người trong độ tuổi lao động | Người |
| **CT11** | Số người tham gia Bảo hiểm y tế (BHYT) | Người |
| **CT12** | Số thành viên Tổ công nghệ số cộng đồng | Người |
| **CT13** | Số người dân được hướng dẫn dùng DVC trực tuyến trong kỳ | Người |
| **CT14** | Số vụ bạo lực gia đình ghi nhận trong kỳ | Vụ |

---

## Hướng Dẫn Cài Đặt & Khởi Chạy

### Cách 1: Chạy nhanh bằng File Script có sẵn
- **Trên Windows**: Kích đúp vào file `run_app.bat`
- **Trên Linux/macOS**: Cấp quyền thực thi và chạy file `run_app.sh`
  ```bash
  chmod +x run_app.sh
  ./run_app.sh
  ```
*Script sẽ tự động tạo môi trường ảo `.venv`, cài đặt thư viện từ `requirements.txt` và khởi động cổng web.*

---

### Cách 2: Khởi chạy thủ công bằng dòng lệnh

**Bước 1: Tạo và kích hoạt môi trường ảo Python**
- **Windows**:
  ```bash
  python -m venv .venv
  .venv\Scripts\activate
  ```
- **Linux/macOS**:
  ```bash
  python3 -m venv .venv
  source .venv/bin/activate
  ```

**Bước 2: Cài đặt các thư viện phụ thuộc**
```bash
pip install -r requirements.txt
```

**Bước 3: Khởi chạy ứng dụng Streamlit**
```bash
python -m streamlit run app.py
```
*Sau khi chạy thành công, trình duyệt web sẽ tự động mở giao diện ứng dụng tại địa chỉ mặc định `http://localhost:8501`.*

---
