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

# Smart & Safe Measurement Sheet Parser
def process_measurement_sheet(file_obj):
    excel_file = pd.ExcelFile(file_obj)
    sheet_names = excel_file.sheet_names
    
    # 1. Target Sheet Selection
    target_sheet = None
    for s in sheet_names:
        if "MEASUREMENT" in s.upper():
            target_sheet = s
            break
    if not target_sheet:
        target_sheet = sheet_names[1] if len(sheet_names) >= 2 else sheet_names[0]

    # Read raw data without assuming top row as header
    df_raw = pd.read_excel(file_obj, sheet_name=target_sheet, header=None)

    # 2. Find actual data starting row (where S.NO / numbers start)
    start_row = 0
    for idx, row in df_raw.iterrows():
        # Check if first/second column has numeric S.NO (e.g. 1, 2, 3...)
        val0 = str(row[0]).strip()
        val1 = str(row[1]).strip()
        if val0.isdigit() or val1.isdigit():
            start_row = idx
            break

    # Extract Data Rows
    df_data = df_raw.iloc[start_row:].copy().reset_index(drop=True)

    rows = []
    for _, row in df_data.iterrows():
        # Column mappings based on standard MEASUREMENT sheet layout
        # Col 0: S.NO, Col 1: FLAT NO, Col 2: WINDOW TYPE, Col 3: LOCATION
        # Col 5: Width, Col 6: Height, Col 7: Area Sq.Ft
        # Col 11: Glass / Flymesh Spec
        
        win_type = str(row[2]).strip() if pd.notna(row[2]) else ""
        location = str(row[3]).strip() if pd.notna(row[3]) else ""
        
        # Combine Window Type + Location for unique identifier
        window_name = f"{win_type} ({location})" if location and location != "nan" else win_type
        if not window_name or window_name == "nan":
            continue

        width = pd.to_numeric(row[5], errors='coerce')
        height = pd.to_numeric(row[6], errors='coerce')
        sqft = pd.to_numeric(row[7], errors='coerce')
        glass_spec = str(row[11]).strip() if pd.notna(row[11]) else ""

        # Ignore invalid/summary rows
        if pd.isna(sqft) or sqft <= 0:
            continue

        # Rule for Special Glass: Check "frosted toughened" (Ignore plain frosted)
        glass_lower = glass_spec.lower()
        is_special = ("frosted" in glass_lower and "toughened" in glass_lower) or ("satin toughened" in glass_lower)

        rows.append({
            'Window Code / Type': window_name,
            'Width (mm)': width if pd.notna(width) else "-",
            'Height (mm)': height if pd.notna(height) else "-",
            'Glass Specification': glass_spec,
            'SQFT': sqft,
            'Is_Special': is_special
        })

    df_clean = pd.DataFrame(rows)

    if df_clean.empty:
        return pd.DataFrame(), target_sheet

    # Group by Window Code/Type
    summary = []
    for win_code, group in df_clean.groupby('Window Code / Type'):
        all_sqft = group['SQFT'].sum()
        special_sqft = group[group['Is_Special']]['SQFT'].sum()
        
        sample_w = group['Width (mm)'].iloc[0]
        sample_h = group['Height (mm)'].iloc[0]
        glass_type = ", ".join([g for g in group['Glass Specification'].unique() if g])

        summary.append({
            'Window Code / Type': win_code,
            'Width (mm)': sample_w,
            'Height (mm)': sample_h,
            'Qty': len(group),
            'Glass Specification': glass_type if glass_type else "Standard",
            'ALL Window SQFT': round(all_sqft, 2),
            'Special glass SQFT': round(special_sqft, 2)
        })

    return pd.DataFrame(summary), target_sheet

# Custom UI CSS
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
    </style>
""", unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    if logo_b64:
        st.markdown(f'<img src="data:image/png;base64,{logo_b64}" style="width:140px;">', unsafe_allow_html=True)
    st.markdown("### 🪟 Window Details Module")
    st.caption("Reads MEASUREMENT Sheet (ALL Window SQFT & Special Glass SQFT)")

# Header
st.markdown("""
    <div class="header-container">
        <div class="main-title">Window Details & Glass SQFT Engine</div>
        <div class="main-subtitle">Automated Reader for 'MEASUREMENT' Sheet</div>
    </div>
""", unsafe_allow_html=True)

# File Upload
uploaded_file = st.file_uploader("Upload Excel BOQ File", type=["xlsx", "xls"])

if uploaded_file:
    try:
        result_df, sheet_used = process_measurement_sheet(uploaded_file)
        
        st.success(f"Successfully processed sheet: **'{sheet_used}'**")
        
        if not result_df.empty:
            st.markdown("### 📑 Window Details Output Table")
            st.dataframe(result_df, use_container_width=True)

            # Metrics
            st.markdown("<br>", unsafe_allow_html=True)
            c1, c2, c3 = st.columns(3)
            
            tot_all = result_df["ALL Window SQFT"].sum()
            tot_spec = result_df["Special glass SQFT"].sum()

            with c1:
                st.metric("Total Window Types", len(result_df))
            with c2:
                st.metric("Total ALL Window SQFT", f"{tot_all:,.2f} sqft")
            with c3:
                st.metric("Total Special Glass SQFT", f"{tot_spec:,.2f} sqft")
        else:
            st.warning("No valid measurement rows found in the sheet. Please check the Excel file.")

    except Exception as e:
        st.error(f"Error parsing sheet: {str(e)}")
