import os
import base64
import pandas as pd
import streamlit as st

# Page Config
st.set_page_config(
    page_title="Window Details | Glass Calculator",
    page_icon="🪟",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Helper Function: Base64 Logo
def get_base64_image(image_path: str) -> str | None:
    if os.path.exists(image_path):
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    return None

logo_b64 = get_base64_image("logo.png")

# Data Processing Function
def process_measurement_sheet(file_obj) -> pd.DataFrame:
    # 1. Excel शीट लोड करा आणि 'MEASUREMENT' किंवा 2nd sheet शोधा
    excel_file = pd.ExcelFile(file_obj)
    sheet_names = excel_file.sheet_names
    
    target_sheet = None
    for s in sheet_names:
        if "MEASUREMENT" in s.upper():
            target_sheet = s
            break
    
    if not target_sheet:
        if len(sheet_names) >= 2:
            target_sheet = sheet_names[1]  # 2nd Sheet
        else:
            target_sheet = sheet_names[0]

    # Read sheet without strictly enforcing single header
    df_raw = pd.read_excel(file_obj, sheet_name=target_sheet)

    # Clean Multi-level or Merged Headers
    # Find the header row where "WINDOW" or "PROFILE" or "SL.NO" exists
    header_idx = 0
    for idx, row in df_raw.head(10).iterrows():
        row_str = " ".join([str(val).upper() for val in row.values])
        if "WINDOW" in row_str or "AREA" in row_str or "WIDTH" in row_str:
            header_idx = idx
            break

    df = pd.read_excel(file_obj, sheet_name=target_sheet, header=header_idx)
    
    # Clean Column Names
    df.columns = [str(c).replace('\n', ' ').strip() for c in df.columns]

    # Map necessary columns dynamically
    win_col = None
    width_col = None
    height_col = None
    sqft_col = None
    glass_col = None
    qty_col = None

    for col in df.columns:
        c_upper = col.upper()
        if "WINDOW" in c_upper or "TYPE" in c_upper or "LOCATION" in c_upper:
            if not win_col: win_col = col
        elif "WIDTH" in c_upper:
            width_col = col
        elif "HEIGHT" in c_upper:
            height_col = col
        elif "SQ" in c_upper or "AREA" in c_upper:
            sqft_col = col
        elif "SPECE" in c_upper or "GLASS" in c_upper or "THICKNESS" in c_upper:
            glass_col = col
        elif "QTY" in c_upper or "NOS" in c_upper:
            qty_col = col

    # Fallbacks for critical columns
    win_col = win_col if win_col else df.columns[2]
    sqft_col = sqft_col if sqft_col else df.columns[7]
    glass_col = glass_col if glass_col else df.columns[-1]

    # Filter out invalid/empty rows
    df[sqft_col] = pd.to_numeric(df[sqft_col], errors='coerce')
    df = df.dropna(subset=[sqft_col])
    df = df[df[sqft_col] > 0]

    # Rule: Detect "Frosted Toughened" / "Special Glass" (Ignore plain "Frosted")
    def is_special_glass(val):
        text = str(val).lower()
        if "frosted toughened" in text or "frosted toughen" in text or "special" in text:
            return True
        return False

    df['Is_Special'] = df[glass_col].apply(is_special_glass)

    # Process and Aggregate Window Wise Data
    summary_list = []
    grouped = df.groupby(win_col)

    for win_name, group in grouped:
        all_window_sqft = group[sqft_col].sum()
        special_glass_sqft = group[group['Is_Special']][sqft_col].sum()
        
        # Dimensions and Spec representation
        sample_w = group[width_col].iloc[0] if width_col in group.columns else "-"
        sample_h = group[height_col].iloc[0] if height_col in group.columns else "-"
        sample_qty = len(group)
        glass_type = ", ".join(group[glass_col].astype(str).unique())

        summary_list.append({
            'Window Type / Code': str(win_name),
            'Width (mm)': sample_w,
            'Height (mm)': sample_h,
            'Qty': sample_qty,
            'Glass Specification': glass_type,
            'ALL Window SQFT': round(all_window_sqft, 2),
            'Special glass SQFT': round(special_glass_sqft, 2)
        })

    return pd.DataFrame(summary_list), target_sheet

# Custom UI
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif !important;
        background-color: #f8fafc !important;
        color: #0f172a !important;
    }
    .header-container {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 24px 32px;
        margin-bottom: 24px;
    }
    .main-title { font-size: 22px !important; font-weight: 800 !important; color: #0f172a; }
    .main-subtitle { font-size: 13px !important; color: #64748b; }
    div.stButton > button[kind="primary"] {
        background: #2563eb !important;
        border: none !important;
        color: #ffffff !important;
        font-weight: 600 !important;
    }
    </style>
""", unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    if logo_b64:
        st.markdown(f'<img src="data:image/png;base64,{logo_b64}" style="width:140px;">', unsafe_allow_html=True)
    st.markdown("### 🪟 Window Details Module")
    st.caption("Reads MEASUREMENT / 2nd Sheet to process Window SQFT & Special Frosted Toughened Glass SQFT.")

# Header
st.markdown("""
    <div class="header-container">
        <div class="main-title">Window Details & Glass SQFT Engine</div>
        <div class="main-subtitle">Automated Reader for 'MEASUREMENT' Sheet (ALL Window SQFT & Special Glass SQFT)</div>
    </div>
""", unsafe_allow_html=True)

# File Upload
uploaded_file = st.file_uploader("Upload Excel BOQ File", type=["xlsx", "xls"])

if uploaded_file:
    try:
        result_df, sheet_used = process_measurement_sheet(uploaded_file)
        
        st.success(f"Successfully processed sheet: **'{sheet_used}'**")
        
        st.markdown("### 📑 Window Details Output Table")
        st.dataframe(result_df, use_container_width=True)

        # Totals Display Cards
        st.markdown("<br>", unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        
        total_all_sqft = result_df["ALL Window SQFT"].sum()
        total_special_sqft = result_df["Special glass SQFT"].sum()

        with c1:
            st.metric("Total Window Types", len(result_df))
        with c2:
            st.metric("Total ALL Window SQFT", f"{total_all_sqft:,.2f} sqft")
        with c3:
            st.metric("Total Special Glass SQFT (Frosted Toughened)", f"{total_special_sqft:,.2f} sqft")

    except Exception as e:
        st.error(f"Error parsing sheet: {str(e)}")
