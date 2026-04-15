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
    file_path = r"D:\HOC A.I\KT SCL\BC SCL\Tong Hop.xlsx"
    if not os.path.exists(file_path):
        st.error(f"Không tìm thấy file: {file_path}")
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

analysis_text = f"""Dưới đây là phần trình bày tổng hợp các chỉ số cảnh báo rủi ro về mặt quản trị tài chính doanh nghiệp:

1. Tỷ lệ giải ngân ở mức báo động đỏ (Rủi ro trễ tiến độ hoàn thành chi phí)
- Hiện tại, tổng quy mô vốn khái toán cho {len(df)} công trình là hơn {tong_khai_toan/1e9:,.1f} tỷ đồng.
- Tuy nhiên, giá trị khối lượng thực hiện mới chỉ ghi nhận trên sổ sách là {tong_thuc_hien/1e9:,.2f} tỷ đồng, tức là tỷ lệ khối lượng hoàn thành mới đạt {ty_le_giai_ngan:.2f}%.
=> Kết luận: Sự chênh lệch khổng lồ giữa Ngân sách và Thực tế giải ngân chỉ ra rằng tiến độ thi công và các thủ tục chuẩn bị hồ sơ thanh/quyết toán đang bị đình trệ nghiêm trọng. Việc dồn ứ khối lượng thanh toán vào cuối năm sẽ gây áp lực rất lớn lên dòng tiền của Công ty và rủi ro bị từ chối thanh toán do hồ sơ làm gấp, sai sót.

2. Thiếu đồng bộ giữa "Thực địa" và "Hồ sơ chứng từ"
- Qua rà soát, có nhiều dự án đang ở trạng thái "Đang thi công" hoặc "Lập PAKT" nhưng chưa hề có xác nhận hay chứng từ ghi nhận chi phí dở dang ("Giá trị thực hiện" = 0đ). 
- Đặc biệt, chưa có dự án nào có "Giá trị quyết toán".
=> Kết luận: Đang có sự tắc nghẽn thông tin giữa phòng Kỹ thuật đang giám sát thi công và bộ phận Tài chính Kế toán. Nhà thầu có thể đã làm xong một phần khối lượng nhưng không chịu làm hồ sơ nghiệm thu giai đoạn (nghiệm thu A-B) để Kế toán hạch toán chi phí và theo dõi hạn mức tín dụng.

3. Kẹt nút thắt ở các Dự án trọng điểm
- Đối chiếu công trình có vốn khái toán cao nhất (13.2 tỷ đồng) là dự án "Sửa chữa lớn máy phát điện Cummins" hiện vẫn đang dậm chân ở bước "Lập kế hoạch đấu thầu".
=> Cảnh báo: Với quy định ngân sách SCL hằng năm, việc chậm chọn nhà thầu đối với dự án trên 10 tỷ đồng này tiềm ẩn nguy cơ phá vỡ kế hoạch tài chính đã duyệt. Nếu không ký được hợp đồng và tạm ứng trong quý này, khả năng không kịp thực hiện trong năm tài chính là rất cao.

🔴 KIẾN NGHỊ TỪ KẾ TOÁN TRƯỞNG:
1. Yêu cầu Phòng Kỹ Thuật / Quản lý dự án phải rà soát các hợp đồng để ép dãn tiến độ thanh toán của nhà thầu. Yêu cầu nhà thầu làm ngay hồ sơ nghiệm thu khối lượng đã hoàn thành. 
2. Yêu cầu Phòng Kế hoạch khẩn trương trình phê duyệt sớm hồ sơ đấu thầu dự án máy Cummins, cam kết mốc thời gian chốt hợp đồng.
3. Thường xuyên tổ chức đối chiếu công nợ dở dang hàng tháng giữa kế toán và kỹ thuật để tránh tình trạng "công trình đã xong nhưng sổ sách kế toán chưa có giấy tờ".
"""

st.error(analysis_text)

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
