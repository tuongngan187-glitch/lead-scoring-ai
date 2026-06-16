import streamlit as st
import pandas as pd
import openpyxl
import io
import requests
import json
import os
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from app_lead_scoring import rule_based_scorer, normalize_dataframe, SHEET_URL

# Page configuration
st.set_page_config(
    page_title="AI Lead Scoring Dashboard",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling (Dark-ish mode alignment)
st.markdown("""
<style>
    .main-title {
        font-family: 'Outfit', sans-serif;
        font-size: 2.5rem;
        font-weight: 800;
        background: linear-gradient(135deg, #F59E0B, #3B82F6);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.5rem;
    }
    .sub-title {
        color: #94A3B8;
        font-size: 1rem;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: rgba(255, 255, 255, 0.05);
        border: 1px solid rgba(255, 255, 255, 0.1);
        padding: 1.5rem;
        border-radius: 12px;
        text-align: center;
    }
</style>
""", unsafe_allow_index=True)

# Session State Initialization
if 'leads' not in st.session_state:
    st.session_state.leads = []

def load_mock_data():
    mock_file = "mock_leads.json"
    if os.path.exists(mock_file):
        with open(mock_file, 'r', encoding='utf-8') as f:
            st.session_state.leads = json.load(f)
            st.success("Đã nạp thành công dữ liệu mẫu!")
    else:
        st.error("Không tìm thấy tệp mock_leads.json")

def fetch_google_sheet(url):
    csv_url = url
    if "/edit" in url:
        # Extract sheet id
        match = re.search(r'\/spreadsheets\/d\/([a-zA-Z0-9-_]+)', url)
        if match:
            sheet_id = match.group(1)
            gid = "0"
            gid_match = re.search(r'gid=(\d+)', url)
            if gid_match:
                gid = gid_match.group(1)
            csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
            
    headers = {'User-Agent': 'Mozilla/5.0'}
    try:
        res = requests.get(csv_url, headers=headers, timeout=15)
        if res.status_code == 200:
            df = pd.read_csv(io.StringIO(res.text))
            st.session_state.leads = normalize_dataframe(df)
            st.success(f"Đã đồng bộ thành công {len(st.session_state.leads)} khách hàng từ Google Sheets!")
        else:
            st.error(f"Không thể tải Google Sheet. Mã lỗi: {res.status_code}. Hãy chắc chắn bảng tính được chia sẻ công khai.")
    except Exception as e:
        st.error(f"Lỗi kết nối: {str(e)}")

def generate_excel_bytes(leads_list):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Báo cáo Chấm Điểm Leads"
    ws.views.sheetView[0].showGridLines = True
    
    headers = [
        "Mã Lead", "Họ Tên Khách Hàng", "Số Điện Thoại", "Email", 
        "Nhu Cầu Chi Tiết", "Điểm Đánh Giá", "Phân Loại", "Lý Do Đánh Giá", 
        "Trạng Thái Duyệt", "Ghi Chú Kiểm Duyệt"
    ]
    ws.append(headers)
    
    font_family = "Segoe UI"
    header_font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    data_font = Font(name=font_family, size=11)
    
    thin = Side(border_style="thin", color="D9D9D9")
    double = Side(border_style="double", color="1F4E79")
    grid_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    
    fill_vip = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_med = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    fill_trash = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    
    fill_approved = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    fill_rejected = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    fill_pending = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = Border(left=thin, right=thin, top=thin, bottom=double)
    ws.row_dimensions[1].height = 28
    
    for lead in leads_list:
        row_data = [
            f"L{lead.get('id', 0):03d}",
            lead.get('name', ''),
            lead.get('phone', ''),
            lead.get('email', ''),
            lead.get('requirement', ''),
            lead.get('score', 0) if lead.get('score') is not None else "",
            lead.get('classification', 'Trung bình'),
            lead.get('reason', ''),
            lead.get('status', 'Chưa duyệt'),
            lead.get('reviewer_notes', '')
        ]
        ws.append(row_data)
        
        row_idx = ws.max_row
        ws.row_dimensions[row_idx].height = 24
        
        for col_idx in range(1, len(row_data) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = data_font
            cell.border = grid_border
            cell.alignment = align_left
            
            if col_idx in [1, 3, 6, 7, 9]:
                cell.alignment = align_center
                
            if col_idx == 7:
                val = str(cell.value)
                if val == "VIP":
                    cell.fill = fill_vip
                    cell.font = Font(name=font_family, size=11, bold=True, color="7F6000")
                elif val == "Trung bình":
                    cell.fill = fill_med
                    cell.font = Font(name=font_family, size=11, bold=False, color="1F4E79")
                elif val == "Rác":
                    cell.fill = fill_trash
                    cell.font = Font(name=font_family, size=11, bold=True, color="C00000")
                    
            if col_idx == 9:
                val = str(cell.value)
                if val == "Đã duyệt":
                    cell.fill = fill_approved
                    cell.font = Font(name=font_family, size=11, bold=True, color="375623")
                elif val == "Từ chối":
                    cell.fill = fill_rejected
                    cell.font = Font(name=font_family, size=11, bold=True, color="C00000")
                else:
                    cell.fill = fill_pending
                    cell.font = Font(name=font_family, size=11, bold=False, color="7F6000")

    for col in ws.columns:
        max_len = 0
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        if col[0].column in [5, 8]:
            ws.column_dimensions[col_letter].width = 40
            continue
        for cell in col:
            val_str = str(cell.value or '')
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)
        
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

# MAIN INTERFACE
st.markdown('<div class="main-title">🤖 AI LEAD SCORING DASHBOARD</div>', unsafe_allow_index=True)
st.markdown('<div class="sub-title">Phân Loại Khách Hàng Tiềm Năng Bất Động Sản & Duyệt Kết Quả (Human-in-the-loop)</div>', unsafe_allow_index=True)

# SIDEBAR CONTROL PANEL
with st.sidebar:
    st.header("⚙️ Cấu Hình Dữ Liệu")
    
    # Google Sheet URL
    sheet_url_input = st.text_input("URL Google Sheets", value=SHEET_URL)
    if st.button("🔄 Đồng Bộ Google Sheets", use_container_width=True):
        import re
        fetch_google_sheet(sheet_url_input)
        
    st.markdown("---")
    st.subheader("Dự phòng & Thử nghiệm")
    
    # CSV / Excel Uploader
    uploaded_file = st.file_uploader("Tải lên tệp CSV/Excel", type=["csv", "xlsx"])
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith(".csv"):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            st.session_state.leads = normalize_dataframe(df)
            st.success(f"Nạp thành công {len(st.session_state.leads)} dòng từ tệp tải lên!")
        except Exception as e:
            st.error(f"Lỗi đọc tệp: {str(e)}")
            
    # Load Mock Data Button
    if st.button("📋 Tải Bộ Dữ Liệu Mẫu", use_container_width=True):
        load_mock_data()
        
    st.markdown("---")
    
    # Process scoring button
    if st.button("✨ Kích Hoạt Chấm Điểm AI", type="primary", use_container_width=True):
        if not st.session_state.leads:
            st.warning("Vui lòng nạp dữ liệu khách hàng trước!")
        else:
            updated_count = 0
            for lead in st.session_state.leads:
                res = rule_based_scorer(lead.get('requirement', ''))
                lead['score'] = res['score']
                lead['classification'] = res['classification']
                lead['reason'] = res['reason']
                if 'status' not in lead or not lead['status']:
                    lead['status'] = 'Chưa duyệt'
                if 'reviewer_notes' not in lead:
                    lead['reviewer_notes'] = ''
                updated_count += 1
            st.success(f"Đã chấm điểm xong {updated_count} khách hàng!")

# STATS COUNTER
if st.session_state.leads:
    vip_c = sum(1 for l in st.session_state.leads if l.get('classification') == 'VIP')
    med_c = sum(1 for l in st.session_state.leads if l.get('classification') == 'Trung bình')
    trash_c = sum(1 for l in st.session_state.leads if l.get('classification') == 'Rác')
    app_c = sum(1 for l in st.session_state.leads if l.get('status') == 'Đã duyệt')
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("🏆 Khách Hàng VIP", vip_c)
    with col2:
        st.metric("🏠 Khách Tầm Trung", med_c)
    with col3:
        st.metric("🗑️ Khách Hàng Rác", trash_c)
    with col4:
        st.metric("✅ Đã Phê Duyệt", app_c)

# INTERACTIVE DATA EDITOR (HUMAN IN THE LOOP)
st.subheader("📋 Bảng Kiểm Duyệt Khách Hàng")
if st.session_state.leads:
    # Convert session state leads to DataFrame for display/edit
    df_leads = pd.DataFrame(st.session_state.leads)
    
    # Make sure all required columns exist
    for col in ['score', 'classification', 'reason', 'status', 'reviewer_notes']:
        if col not in df_leads.columns:
            df_leads[col] = None
            
    # Order columns
    col_order = ['id', 'name', 'phone', 'requirement', 'score', 'classification', 'reason', 'status', 'reviewer_notes']
    df_leads = df_leads[col_order]
    
    # Configure columns for editing
    edited_df = st.data_editor(
        df_leads,
        column_config={
            "id": st.column_config.NumberColumn("Mã", disabled=True),
            "name": st.column_config.TextColumn("Họ tên", required=True),
            "phone": st.column_config.TextColumn("Số điện thoại"),
            "requirement": st.column_config.TextColumn("Nhu cầu", width="large"),
            "score": st.column_config.NumberColumn("Điểm AI", disabled=True),
            "classification": st.column_config.SelectboxColumn(
                "Phân loại AI",
                options=["VIP", "Trung bình", "Rác"],
                disabled=True
            ),
            "reason": st.column_config.TextColumn("Lý do chấm điểm", disabled=True, width="medium"),
            "status": st.column_config.SelectboxColumn(
                "Duyệt (Human)",
                options=["Chưa duyệt", "Đã duyệt", "Từ chối"],
                required=True
            ),
            "reviewer_notes": st.column_config.TextColumn("Ghi chú kiểm duyệt", width="medium")
        },
        disabled=["id", "score", "classification", "reason"],
        hide_index=True,
        use_container_width=True
    )
    
    # Save edits back to session state
    st.session_state.leads = edited_df.to_dict('records')
    
    st.markdown("---")
    
    # EXPORT ACTION
    excel_bytes = generate_excel_bytes(st.session_state.leads)
    st.download_button(
        label="📥 Tải xuống File Excel (.xlsx)",
        data=excel_bytes,
        file_name="Bao_Cao_Khach_Hang_Tiem_Nang.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
else:
    st.info("Chưa có dữ liệu. Hãy sử dụng bảng cấu hình bên trái để nạp Google Sheets hoặc dữ liệu mẫu.")
