# 📊 Báo Cáo Tổng Hợp & Phân Tích Quản Trị Chi Phí Sửa Chữa Lớn (SCL)

Đây là một ứng dụng Dashboard tài chính chuyên nghiệp dựa trên thư viện **Streamlit** (Python). Ứng dụng tự động đọc dữ liệu từ tệp Excel, thiết kế biểu đồ so sánh chi phí Khái toán - Thực hiện và tự động xuất **Báo cáo cảnh báo rủi ro chuyên môn từ Kế Toán Trưởng** định dạng Word (.docx). 

## 1. Yêu cầu hệ thống (Requirements)

Để ứng dụng có thể chạy mượt mà, máy tính của bạn cần cài đặt **Python** (version >= 3.8) và cài đặt đầy đủ các thư viện ở trong file `requirements.txt`.

Chạy dòng lệnh sau trong Terminal hoặc Command Prompt để cài đặt các thư viện:
```bash
pip install -r requirements.txt
```

*Các thư viện chính bao gồm: `streamlit` (UI), `pandas` (Xử lý dữ liệu), `matplotlib` (Biểu đồ), `openpyxl` (Đọc Excel), `python-docx` (Tạo file Word).*

## 2. Cấu trúc thư mục định dạng
```
📁 BC SCL
├── 📄 scl_dashboard.py    <-- File mã nguồn chính của ứng dụng
├── 📄 Tong Hop.xlsx       <-- File số liệu nguồn (Bắt buộc phải có)
├── 📄 requirements.txt    <-- Danh sách thư cài đặt Python
└── 📄 README.md           <-- File hướng dẫn này
```

**⚠️ Lưu ý quan trọng:** File dữ liệu `Tong Hop.xlsx` phải nằm chung một cấp thư mục với `scl_dashboard.py` hoặc sửa lại thành đường dẫn tuyệt đối trong file code. Dữ liệu mặc định sẽ được đọc từ Sheet có tên là `Sheet1`.

## 3. Hướng dẫn chạy ứng dụng

- **Bước 1**: Mở Terminal hoặc Command Prompt, truy cập vào thư mục chứa `scl_dashboard.py`. Hoặc mở trực tiếp thư mục `D:\HOC A.I\KT SCL\BC SCL\` trên CMD.
- **Bước 2**: Gõ lệnh khởi động Server nội bộ của Streamlit:
```bash
streamlit run scl_dashboard.py
```
*(Nếu bạn để đường dẫn tuyệt đối thì chạy lệnh: `streamlit run "D:\HOC A.I\KT SCL\BC SCL\scl_dashboard.py"`)*
- **Bước 3**: Ngay lập tức, trình duyệt mặc định của hệ thống tính toán (thường là Chrome/Edge) sẽ mở ra một cửa sổ (thường là link `http://localhost:8501`) chiếu trực tiếp giao diện.

## 4. Các Chức Năng Chính

* **🔄 Reload Runtime (Tải lại dữ liệu sống):** Bất cứ khi nào bạn chỉnh sửa số liệu bên trong tệp `Tong Hop.xlsx` và bấm Save trên Excel. Ngay trên giao diện Web, ấn nút **"🔄 Cập nhật/Tải lại dữ liệu"**, ứng dụng sẽ nạp số liệu mới tức thời vào biểu đồ.
* **Biểu đồ Cột Vốn & Pie Chart Mảng:** Khủng hoảng tài chính được dễ nhìn chỉ nhờ 2 biểu đồ quan trọng này với tỉ lệ màu chuyên nghiệp.
* **Tự Động Xuất Báo Cáo Word (Auto-reporting):** Ấn nút "📥 Tải Xuất báo cáo chính thức" ở cuối màn hình Web để trích xuất file .docx đóng dấu nội bộ đi nộp luôn cho Ban Giám Đốc mà không tốn công chép tay đánh văn bản lỗi số đếm.

---
*Developed as a specific feature module.*
