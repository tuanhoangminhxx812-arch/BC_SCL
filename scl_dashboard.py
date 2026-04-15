import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import docx
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import os

# --- Cấu hình trang Streamlit ---
st.set_page_config(page_title="Báo Cáo Kế Toán SCL", layout="wide", initial_sidebar_state="collapsed")

# --- Đọc dữ liệu ---
@st.cache_data
def load_data():
    # Ưu tiên tìm file cùng thư mục với script này (để tương thích khi deploy lên Streamlit Cloud)
    base_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()
    file_path = os.path.join(base_dir, "Tong Hop.xlsx")
    
    # Fallback cho chạy local nếu thư mục hiện tại khác
    if not os.path.exists(file_path):
        local_path = r"D:\HOC A.I\KT SCL\BC SCL\Tong Hop.xlsx"
        if os.path.exists(local_path):
            file_path = local_path

    if not os.path.exists(file_path):
        st.error("Không tìm thấy file dữ liệu 'Tong Hop.xlsx'. Vui lòng upload file này vào cùng thư mục với mã nguồn trên Github.")
        return pd.DataFrame()
        
    df = pd.read_excel(file_path, sheet_name="Sheet1")
    
    # Lọc bỏ dòng "Tổng giá trị" (dòng mà Mã công trình bị trống)
    df = df[df['Mã công trình'].notna()]
    
    # Fill NA cho các cột giá trị
    df['Giá trị khái toán'] = df['Giá trị khái toán'].fillna(0)
    df['Giá trị thực hiện'] = df['Giá trị thực hiện'].fillna(0)
    df['Giá trị quyết toán'] = df['Giá trị quyết toán'].fillna(0)
    
    return df

df = load_data()

if df.empty:
    st.stop()

# --- Tính toán các chỉ số ---
tong_khai_toan = df['Giá trị khái toán'].sum()
tong_thuc_hien = df['Giá trị thực hiện'].sum()
tong_quyet_toan = df['Giá trị quyết toán'].sum()
ty_le_giai_ngan = (tong_thuc_hien / tong_khai_toan * 100) if tong_khai_toan > 0 else 0

# --- Giao diện Dashboard ---
st.title("📊 BÁO CÁO TỔNG HỢP & PHÂN TÍCH QUẢN TRỊ CHI PHÍ SCL")

col_title, col_btn = st.columns([8, 2])
with col_btn:
    if st.button("🔄 Cập nhật/Tải lại dữ liệu", type="primary", use_container_width=True):
        load_data.clear()
        st.rerun()

st.markdown("---")

col1, col2, col3, col4 = st.columns(4)
col1.metric("Tổng Số Công Trình", len(df))
col2.metric("Tổng Giá Trị Khái Toán", f"{tong_khai_toan:,.0f} đ")
col3.metric("Tổng Giá Trị Thực Hiện", f"{tong_thuc_hien:,.0f} đ")
col4.metric("Tỷ Lệ Giải Ngân", f"{ty_le_giai_ngan:.2f} %")

st.markdown("---")
st.subheader("📈 Sơ đồ trực quan hóa dữ liệu")

col_chart1, col_chart2 = st.columns(2)

with col_chart1:
    st.markdown("**1. Tỷ trọng trạng thái dự án**")
    status_counts = df['Trạng thái'].value_counts()
    
    fig1, ax1 = plt.subplots(figsize=(7, 4))
    # Sử dụng bảng màu chuyên nghiệp
    colors = ['#ff9999', '#66b3ff', '#99ff99', '#ffcc99']
    ax1.pie(status_counts, labels=status_counts.index, autopct='%1.1f%%', startangle=90, colors=colors[:len(status_counts)], wedgeprops={'edgecolor': 'white'})
    ax1.axis('equal')  
    st.pyplot(fig1)

with col_chart2:
    st.markdown("**2. Top dự án có mức ngân sách cao nhất (Khái toán vs Thực hiện)**")
    df_sorted = df.sort_values(by='Giá trị khái toán', ascending=False).head(5)
    
    fig2, ax2 = plt.subplots(figsize=(8, 4))
    x = range(len(df_sorted))
    width = 0.35
    
    # Đơn vị tỷ đồng để hiển thị đẹp hơn
    khai_toan_ty = df_sorted['Giá trị khái toán'] / 1e9
    thuc_hien_ty = df_sorted['Giá trị thực hiện'] / 1e9
    
    ax2.bar([i - width/2 for i in x], khai_toan_ty, width, label='Khái toán (Tỷ đ)', color='#2171b5')
    ax2.bar([i + width/2 for i in x], thuc_hien_ty, width, label='Thực hiện (Tỷ đ)', color='#fd8d3c')
    
    ax2.set_xticks(x)
    m_cong_trinh = df_sorted['Mã công trình'].tolist()
    ax2.set_xticklabels(m_cong_trinh, rotation=30, ha="right")
    ax2.legend()
    ax2.grid(axis='y', linestyle='--', alpha=0.7)
    
    st.pyplot(fig2)

st.markdown("---")
st.subheader("📋 Bảng số liệu chi tiết các dự án SCL")
# Định dạng số tiền có dấu phẩy để dễ nhìn
df_display = df[['Mã công trình', 'Tên công trình', 'Trạng thái', 'Giá trị khái toán', 'Giá trị thực hiện', 'Giá trị quyết toán']].copy()
for col in ['Giá trị khái toán', 'Giá trị thực hiện', 'Giá trị quyết toán']:
    df_display[col] = df_display[col].apply(lambda x: f"{x:,.0f}")
st.dataframe(df_display, use_container_width=True)

# --- Phân tích của Kế toán trưởng ---
st.markdown("---")
st.subheader("⚠️ Phân tích rủi ro & Đề xuất (Trình Ban Giám Đốc)")

# Đánh giá tỷ lệ giải ngân động theo số liệu
if ty_le_giai_ngan < 30:
    nhan_xet_giai_ngan = "Ở mức **báo động đỏ** (Trễ tiến độ giải ngân)"
    kl_giai_ngan = "Sự chênh lệch lớn giữa Ngân sách và Thực tế cho thấy các thủ tục chuẩn bị hồ sơ thanh toán đang bị đình trệ nghiêm trọng."
elif ty_le_giai_ngan < 70:
    nhan_xet_giai_ngan = "Ở mức **trung bình** (Cần đẩy nhanh hơn)"
    kl_giai_ngan = "Tiến độ giải ngân đang được thực hiện nhưng cần đốc thúc thêm để rải đều trong năm, hoàn thành đúng mục tiêu dòng tiền."
else:
    nhan_xet_giai_ngan = "Ở mức **rất tốt** (Hoàn thành theo bám sát kế hoạch)"
    kl_giai_ngan = "Các công tác thi công và nghiệm thu hồ sơ đang phối hợp rất nhịp nhàng, đảm bảo tính pháp lý và giảm tải rủi ro dồn khối lượng vào cuối năm."

# Tính số dư
so_du_an_0 = len(df[df['Giá trị thực hiện'] == 0])
so_du_an_quyet_toan = len(df[df['Giá trị quyết toán'] > 0])
tong_du_an = len(df)

# Phân tích đồng bộ chứng từ
if so_du_an_0 > 0:
    nhan_xet_0 = f"Báo cáo cho thấy có **{so_du_an_0}** dự án hoàn toàn chưa ghi nhận chứng từ chi phí dở dang ('Giá trị thực hiện' = 0đ)."
    kl_chung_tu = "Cần rà soát chéo lượng công trình này ngay. Xem đây là do thực sự chưa triển khai ngoài hiện trường, hay kỹ thuật đã cho làm nhưng nhà thầu chây ỳ chưa lập hồ sơ nghiệm thu. Tránh tình trạng nợ đọng, thi công xong mà sổ sách không có chứng từ."
else:
    nhan_xet_0 = f"Rất tốt, 100% ({tong_du_an}/{tong_du_an}) dự án đều đã có hồ sơ ghi nhận Khối lượng thực hiện ban đầu."
    kl_chung_tu = "Sự phối hợp cập nhật chứng từ giữa phòng Kỹ thuật và Kế toán đang bám sát thực tế, không có tình trạng bị trễ nhịp hay bỏ quên dự án."

if so_du_an_quyet_toan == 0:
    nhan_xet_qt = "Đồng thời, chưa có dự án nào chuyển sang bước 'Giá trị quyết toán'."
elif so_du_an_quyet_toan == tong_du_an:
    nhan_xet_qt = "Tuyệt vời, tất cả các dự án đều đã có số liệu Quyết Toán! Quá trình khép sổ tài chính SCL gần như đã trọn vẹn."
else:
    nhan_xet_qt = f"Tiến độ quyết toán: Đã có **{so_du_an_quyet_toan}/{tong_du_an}** dự án có số liệu Quyết toán thành công."

# Kiến nghị
if so_du_an_quyet_toan == tong_du_an:
    kien_nghi = "- Hồ sơ tài chính đã đạt mức độ hoàn thiện cao. Đề nghị các phòng ban chuẩn bị đóng luồng hồ sơ cuối năm và báo cáo Giám đốc."
else:
    kien_nghi = "- Đẩy nhanh tiến độ hoàn công chuyển các công trình thành Quyết Toán (QT).\n- Liên tục tổ chức đối chiếu công nợ khối lượng dở dang hàng tháng giữa kế toán và kỹ thuật."

# Sinh nội dung văn bản phân tích
analysis_text = f"Dưới đây là phần trình bày tổng hợp các chỉ số đánh giá chuyên môn về mặt quản trị tài chính doanh nghiệp:\n\n"

analysis_text += f"""**1. Tỷ lệ giải ngân: {nhan_xet_giai_ngan}**
- Tổng quy mô vốn khái toán cho {tong_du_an} công trình là hơn **{tong_khai_toan/1e9:,.1f} tỷ đồng**.
- Tuy nhiên, giá trị khối lượng thực hiện ghi nhận trên sổ sách là **{tong_thuc_hien/1e9:,.2f} tỷ đồng**, đạt mức **{ty_le_giai_ngan:.2f}%**.
=> Kết luận: {kl_giai_ngan}

**2. Công tác đồng bộ giữa thực địa và chứng từ (Phòng Kỹ thuật vs Kế toán)**
- {nhan_xet_0}
- {nhan_xet_qt}
=> Kết luận: {kl_chung_tu}

**🔴 KIẾN NGHỊ TỪ KẾ TOÁN TRƯỞNG:**
{kien_nghi}
"""

st.markdown(analysis_text)

# --- Hàm tạo báo cáo Word ---
def export_word_report():
    doc = docx.Document()
    
    # Cấu hình Font mặc định toàn văn bản
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(13)
    
    # Header
    doc.add_paragraph('CÔNG TY ĐIỆN LỰC VŨNG TÀU\nPHÒNG TÀI CHÍNH KẾ TOÁN')
    
    title = doc.add_paragraph('\nBÁO CÁO PHÂN TÍCH TÌNH HÌNH THỰC HIỆN KẾ HOẠCH TÀI CHÍNH\nCÔNG TÁC SỬA CHỮA LỚN NĂM 2026\n')
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.bold = True
        run.font.size = Pt(16)
        
    p = doc.add_paragraph()
    p.add_run('Kính gửi: ').bold = True
    p.add_run('Ông (Bà) Giám đốc Công ty')
    
    doc.add_paragraph(f"Căn cứ vào dữ liệu tổng hợp về tình hình thực hiện kế hoạch các dự án sửa chữa lớn, trên cương vị Kế toán trưởng, tôi xin báo cáo các số liệu tài chính quan trọng và các điểm bất ổn cần Giám đốc khẩn trương chỉ đạo như sau:")
    
    doc.add_paragraph('I. BẢNG TỔNG HỢP SỐ LIỆU TÀI CHÍNH:', style='Heading 3')
    
    # Tạo bảng
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Chỉ tiêu'
    hdr_cells[1].text = 'Kế hoạch/Khái toán (VNĐ)'
    hdr_cells[2].text = 'Giá trị thực hiện (VNĐ)'
    hdr_cells[3].text = 'Tỉ lệ hoàn thành (%)'
    for cell in hdr_cells:
        cell.paragraphs[0].runs[0].bold = True
    
    row_cells = table.add_row().cells
    row_cells[0].text = 'Toàn bộ công trình SCL'
    row_cells[1].text = f"{tong_khai_toan:,.0f}"
    row_cells[2].text = f"{tong_thuc_hien:,.0f}"
    row_cells[3].text = f"{ty_le_giai_ngan:.2f}%"
    
    doc.add_paragraph()
    
    doc.add_paragraph('II. PHÂN TÍCH ĐÁNH GIÁ & CẢNH BÁO BẤT ỔN:', style='Heading 3')
    
    parts = analysis_text.split('\n')
    for p_text in parts:
        if p_text.strip():
            p_docx = doc.add_paragraph(p_text.strip())
            if p_text.startswith("🔴") or "CẢNH BÁO" in p_text or "Kết luận" in p_text:
                for run in p_docx.runs:
                    run.bold = True
            
    doc.add_paragraph()
    
    # Ký tên
    p_sig = doc.add_paragraph()
    p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_date = p_sig.add_run('Vũng Tàu, ngày ...... tháng ...... năm ...... \n')
    run_date.italic = True
    run_title = p_sig.add_run('KẾ TOÁN TRƯỞNG\n\n\n\n\n')
    run_title.bold = True
    p_sig.add_run('(Đã ký)')
    
    # Chuyển docx buffer sang bytes để Streamlit tải
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

st.markdown("### 📥 Tải Xuất báo cáo chính thức")
st.download_button(
    label="📄 Tải Xuất Báo Cáo Kế Toán Trưởng (File Word .docx)",
    data=export_word_report(),
    file_name="Bao_Cao_SCL_KeToanTruong.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)

# Chạy ứng dụng bằng lệnh: streamlit run "D:\HOC A.I\KT SCL\BC SCL\scl_dashboard.py"
