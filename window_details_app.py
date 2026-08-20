import os
import sys
import base64
import pandas as pd
import streamlit as st

# Page Configuration
st.set_page_config(
    page_title="Window Details | Glass Calculator",
    page_icon="🪟",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# HELPER FUNCTIONS
# ============================================================
def get_base64_image(image_path: str) -> str | None:
    if os.path.exists(image_path):
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    return None

logo_b64 = get_base64_image("logo.png")

def calculate_window_sqft(df: pd.DataFrame) -> pd.DataFrame:
    """
    Window Code/Name प्रमाणे Total Sqft, Frosted Sqft आणि Non-Frosted Sqft कॅल्क्युलेट करतो.
    """
    col_map = {}
    for col in df.columns:
        c_lower = str(col).strip().lower()
        if any(k in c_lower for k in ["window", "tag", "item", "code", "win"]):
            col_map['window'] = col
        elif any(k in c_lower for k in ["sqft", "area", "sq.ft", "sq ft"]):
            col_map['sqft'] = col
        elif any(k in c_lower for k in ["spec", "glass", "description", "type", "frosted"]):
            col_map['spec'] = col

    # Fallbacks if column names don't match standard keywords
    win_col = col_map.get('window', df.columns[0])
    sqft_col = col_map.get('sqft', df.columns[1] if len(df.columns) > 1 else df.columns[0])
    spec_col = col_map.get('spec', df.columns[2] if len(df.columns) > 2 else df.columns[0])

    # Clean data
    df_clean = df.copy()
    df_clean[sqft_col] = pd.to_numeric(df_clean[sqft_col], errors='coerce').fillna(0)
    df_clean['Is_Frosted'] = df_clean[spec_col].astype(str).str.lower().str.contains('frost|frosted|satin|etched|opaque')

    # Safe GroupBy Calculation
    records = []
    for win_name, group in df_clean.groupby(win_col):
        tot_sqft = group[sqft_col].sum()
        frosted_sqft = group[group['Is_Frosted']][sqft_col].sum()
        non_frosted_sqft = group[~group['Is_Frosted']][sqft_col].sum()
        
        records.append({
            'Window Code / Name': str(win_name),
            'Total OC Sqft': round(tot_sqft, 2),
            'Frosted Sqft': round(frosted_sqft, 2),
            'Non-Frosted Sqft': round(non_frosted_sqft, 2),
            'Total Panels': len(group)
        })

    result_df = pd.DataFrame(records)
    return result_df

# ============================================================
# CUSTOM UI CSS
# ============================================================
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
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.02);
    }

    .main-title {
        color: #0f172a !important;
        font-size: 22px !important;
        font-weight: 800 !important;
        margin: 0 0 6px 0 !important;
    }

    .main-subtitle {
        color: #64748b !important;
        font-size: 13px !important;
        font-weight: 500 !important;
        margin: 0 !important;
    }

    .step-heading {
        color: #0f172a;
        font-size: 15px;
        font-weight: 700;
        margin-top: 10px;
        margin-bottom: 12px;
    }

    div.stButton > button[kind="primary"] {
        background: #2563eb !important;
        border: none !important;
        color: #ffffff !important;
        font-weight: 600 !important;
        border-radius: 6px !important;
    }

    .sidebar-logo {
        width: 140px;
        height: auto;
        margin-bottom: 20px;
    }

    [data-testid="stHeader"] { display: none; }
    </style>
""", unsafe_allow_html=True)

# ============================================================
# SIDEBAR
# ============================================================
with st.sidebar:
    if logo_b64:
        st.markdown(f'<img src="data:image/png;base64,{logo_b64}" class="sidebar-logo">', unsafe_allow_html=True)
    else:
        st.markdown("<h2 style='color:#0f172a; font-weight:800;'>WinSquare</h2>", unsafe_allow_html=True)
    
    st.divider()
    st.markdown("### 🪟 Window Details Module")
    st.caption("Auto-calculates total square feet per window code with Frosted vs Non-Frosted glass breakdown.")

# ============================================================
# HEADER
# ============================================================
st.markdown("""
    <div class="header-container">
        <div class="main-title">Window Details & Area Breakdown</div>
        <div class="main-subtitle">Comprehensive window codes directory, total Sqft (All OC), Frosted and Non-Frosted Glass measurement generator.</div>
    </div>
""", unsafe_allow_html=True)

# ============================================================
# FILE UPLOAD
# ============================================================
st.markdown('<div class="step-heading">📁 Upload BOQ / Window Schedule Excel Sheet</div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "Upload BOQ Excel file",
    type=["xlsx", "xls"],
    label_visibility="collapsed"
)

if uploaded_file:
    try:
        raw_df = pd.read_excel(uploaded_file)
        st.success(f"File uploaded successfully! Loaded {len(raw_df)} rows.")
        
        with st.expander("📄 View Raw Excel File", expanded=False):
            st.dataframe(raw_df, use_container_width=True)

        if st.button("📊 Generate Window Details & Breakdown", type="primary"):
            with st.spinner("Calculating Window-Wise Square Feet..."):
                result_df = calculate_window_sqft(raw_df)
                st.session_state["window_result"] = result_df

    except Exception as e:
        st.error(f"Error reading file: {str(e)}")

# ============================================================
# RESULT DASHBOARD & SUMMARY
# ============================================================
if "window_result" in st.session_state and not st.session_state["window_result"].empty:
    res_df = st.session_state["window_result"]

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("### 📊 Overall Area Summary")

    total_windows = len(res_df)
    total_oc_sqft = res_df["Total OC Sqft"].sum()
    total_frosted_sqft = res_df["Frosted Sqft"].sum()
    total_non_frosted_sqft = res_df["Non-Frosted Sqft"].sum()

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.markdown(f"""
            <div style="background: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; padding: 14px 18px;">
                <p style="color: #64748b; font-size: 11px; font-weight: 700; margin: 0;">TOTAL WINDOW TYPES</p>
                <h3 style="color: #0f172a; font-size: 24px; font-weight: 800; margin: 4px 0 0 0;">{total_windows}</h3>
            </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown(f"""
            <div style="background: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; padding: 14px 18px;">
                <p style="color: #2563eb; font-size: 11px; font-weight: 700; margin: 0;">TOTAL OC SQFT</p>
                <h3 style="color: #1d4ed8; font-size: 24px; font-weight: 800; margin: 4px 0 0 0;">{total_oc_sqft:,.2f}</h3>
            </div>
        """, unsafe_allow_html=True)

    with col3:
        st.markdown(f"""
            <div style="background: #fefce8; border: 1px solid #fef08a; border-radius: 8px; padding: 14px 18px;">
                <p style="color: #854d0e; font-size: 11px; font-weight: 700; margin: 0;">FROSTED GLASS SQFT</p>
                <h3 style="color: #a16207; font-size: 24px; font-weight: 800; margin: 4px 0 0 0;">{total_frosted_sqft:,.2f}</h3>
            </div>
        """, unsafe_allow_html=True)

    with col4:
        st.markdown(f"""
            <div style="background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 8px; padding: 14px 18px;">
                <p style="color: #166534; font-size: 11px; font-weight: 700; margin: 0;">NON-FROSTED SQFT</p>
                <h3 style="color: #15803d; font-size: 24px; font-weight: 800; margin: 4px 0 0 0;">{total_non_frosted_sqft:,.2f}</h3>
            </div>
        """, unsafe_allow_html=True)

    # Detailed Table
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("### 📑 Window-Wise Breakdown Table")
    st.dataframe(res_df, use_container_width=True, height=350)

    # Built-in Chart
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("### 📈 Windows Area Breakdown Chart")
    chart_data = res_df.set_index("Window Code / Name")[["Non-Frosted Sqft", "Frosted Sqft"]]
    st.bar_chart(chart_data)
