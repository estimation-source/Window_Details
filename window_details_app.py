import os
import re
import base64
import pandas as pd
import streamlit as st

# Page Config
st.set_page_config(
    page_title="Window Details Module",
    page_icon="🪟",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Rule for Special Glass
def check_special_glass(spec):
    glass_lower = str(spec).lower()
    if "frosted" in glass_lower and "toughened" not in glass_lower and "tough" not in glass_lower:
        return False
    elif "toughened" in glass_lower or "tough" in glass_lower or "dgu" in glass_lower or "satin" in glass_lower:
        return True
    return False

# ===================================================================
# OPTION 1: MEASUREMENT SHEET READER (HORIZONTAL TABLE FORMAT)
# ===================================================================
def parse_measurement_sheet(file_obj, sheet_name):
    df_raw = pd.read_excel(file_obj, sheet_name=sheet_name, header=None)

    start_row = 0
    for idx, row in df_raw.iterrows():
        val0 = str(row.iloc[0]).strip() if len(row) > 0 else ""
        val1 = str(row.iloc[1]).strip() if len(row) > 1 else ""
        if val0.isdigit() or val1.isdigit():
            start_row = idx
            break

    df_data = df_raw.iloc[start_row:].copy().reset_index(drop=True)
    rows = []

    for _, row in df_data.iterrows():
        win_type = str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else ""
        location = str(row.iloc[3]).strip() if len(row) > 3 and pd.notna(row.iloc[3]) else ""
        
        window_name = f"{win_type} ({location})" if location and location != "nan" else win_type
        if not window_name or window_name == "nan":
            continue

        width = pd.to_numeric(row.iloc[5], errors='coerce') if len(row) > 5 else "-"
        height = pd.to_numeric(row.iloc[6], errors='coerce') if len(row) > 6 else "-"
        sqft = pd.to_numeric(row.iloc[7], errors='coerce') if len(row) > 7 else 0
        thickness = str(row.iloc[10]).strip() if len(row) > 10 and pd.notna(row.iloc[10]) and str(row.iloc[10]).strip() != "nan" else "-"
        glass_spec = str(row.iloc[11]).strip() if len(row) > 11 and pd.notna(row.iloc[11]) and str(row.iloc[11]).strip() != "nan" else ""

        if pd.isna(sqft) or sqft <= 0:
            continue

        rows.append({
            'Window Code / Type': window_name,
            'Width (mm)': width if pd.notna(width) else "-",
            'Height (mm)': height if pd.notna(height) else "-",
            'Thickness': thickness,
            'Glass Specification': glass_spec,
            'SQFT': sqft
        })

    return rows

# ===================================================================
# OPTION 2: QUOTATION SHEET READER (VERTICAL BLOCK FORMAT - WinSquare)
# ===================================================================
def parse_quotation_block_sheet(file_obj, sheet_name):
    df_raw = pd.read_excel(file_obj, sheet_name=sheet_name, header=None)
    rows = []
    num_rows = len(df_raw)

    for r in range(num_rows):
        row_vals = [str(val).strip() for val in df_raw.iloc[r].values if pd.notna(val) and str(val).strip() != "nan"]
        row_str = " ".join(row_vals)

        if "CODE :" in row_str.upper() or "CODE:" in row_str.upper():
            code_val = ""
            name_val = ""
            glass_val = ""
            width_val = "-"
            height_val = "-"
            sqft_val = 0
            thick_val = "-"

            code_match = re.search(r'CODE\s*:\s*([A-Za-z0-9_\-]+)', row_str, re.IGNORECASE)
            if code_match:
                code_val = code_match.group(1).strip()

            for r_offset in range(r, min(r + 25, num_rows)):
                sub_row = df_raw.iloc[r_offset]
                sub_vals = [str(v).strip() for v in sub_row.values if pd.notna(v) and str(v).strip() != "nan"]
                sub_str = " ".join(sub_vals)

                if ("NAME :" in sub_str.upper() or "NAME:" in sub_str.upper()) and not name_val:
                    m = re.search(r'NAME\s*:\s*(.*?)(?=Profile System|Size|Location|$)', sub_str, re.IGNORECASE)
                    if m:
                        name_val = m.group(1).strip()

                if ("GLASS :" in sub_str.upper() or "GLASS:" in sub_str.upper()) and not glass_val:
                    m = re.search(r'GLASS\s*:\s*(.*)', sub_str, re.IGNORECASE)
                    if m:
                        glass_val = m.group(1).strip()
                        tm = re.search(r'(\d+\s*MM(?:\s*DGU)?)', glass_val, re.IGNORECASE)
                        if tm:
                            thick_val = tm.group(1).strip()

                for c in range(len(sub_row)):
                    cell_txt = str(sub_row.iloc[c]).strip().upper()
                    
                    if cell_txt in ["WIDHT", "WIDTH"]:
                        for c_next in range(c + 1, len(sub_row)):
                            num = pd.to_numeric(sub_row.iloc[c_next], errors='coerce')
                            if pd.notna(num) and num > 0:
                                width_val = num
                                break

                    elif cell_txt == "HEIGHT":
                        for c_next in range(c + 1, len(sub_row)):
                            num = pd.to_numeric(sub_row.iloc[c_next], errors='coerce')
                            if pd.notna(num) and num > 0:
                                height_val = num
                                break

                    elif cell_txt == "SQFT":
                        for c_next in range(c + 1, len(sub_row)):
                            num = pd.to_numeric(sub_row.iloc[c_next], errors='coerce')
                            if pd.notna(num) and num > 0:
                                sqft_val = num
                                break

            full_name = f"{code_val} - {name_val}" if code_val and name_val else (code_val or name_val or "Window Block")
            
            if sqft_val > 0:
                rows.append({
                    'Window Code / Type': full_name,
                    'Width (mm)': width_val,
                    'Height (mm)': height_val,
                    'Thickness': thick_val,
                    'Glass Specification': glass_val,
                    'SQFT': sqft_val
                })

    return rows

# ===================================================================
# MAIN PROCESSOR
# ===================================================================
def process_excel_with_mode(file_obj, format_mode):
    excel_file = pd.ExcelFile(file_obj)
    sheet_names = excel_file.sheet_names

    target_sheet = None

    if format_mode == "Option 1: Measurement Table":
        for s in sheet_names:
            if "MEASUREMENT" in str(s).upper():
                target_sheet = s
                break
        if not target_sheet:
            target_sheet = sheet_names[0]
        parsed_rows = parse_measurement_sheet(file_obj, target_sheet)

    elif format_mode == "Option 2: Quotation Block Layout":
        for s in sheet_names:
            if "SHEET2" in str(s).upper() or "QUOTE" in str(s).upper():
                target_sheet = s
                break
        if not target_sheet:
            target_sheet = sheet_names[1] if len(sheet_names) > 1 else sheet_names[0]
        parsed_rows = parse_quotation_block_sheet(file_obj, target_sheet)

    else:  # AUTO-DETECT
        found_block_sheet = None
        for s in sheet_names:
            txt = " ".join([str(v) for v in pd.read_excel(file_obj, sheet_name=s, header=None).fillna('').values.flatten()]).upper()
            if "CODE :" in txt or "CODE:" in txt or "WIDHT" in txt:
                found_block_sheet = s
                break

        if found_block_sheet:
            target_sheet = found_block_sheet
            parsed_rows = parse_quotation_block_sheet(file_obj, target_sheet)
        else:
            for s in sheet_names:
                if "MEASUREMENT" in str(s).upper():
                    target_sheet = s
                    break
            if not target_sheet:
                target_sheet = sheet_names[0]
            parsed_rows = parse_measurement_sheet(file_obj, target_sheet)

    df_clean = pd.DataFrame(parsed_rows)
    if df_clean.empty:
        return pd.DataFrame(), target_sheet

    df_clean['Is_Special'] = df_clean['Glass Specification'].apply(check_special_glass)

    summary = []
    for win_code, group in df_clean.groupby('Window Code / Type'):
        all_sqft = group['SQFT'].sum()
        special_sqft = group[group['Is_Special']]['SQFT'].sum()
        
        sample_w = group['Width (mm)'].iloc[0]
        sample_h = group['Height (mm)'].iloc[0]
        
        thick_vals = [str(t) for t in group['Thickness'].unique() if str(t).strip() not in ["", "-", "nan"]]
        thick_type = ", ".join(thick_vals) if thick_vals else "-"
        
        glass_vals = [str(g) for g in group['Glass Specification'].unique() if str(g).strip() not in ["", "nan"]]
        glass_type = ", ".join(glass_vals) if glass_vals else "Standard Glass"

        summary.append({
            'Window Code / Type': str(win_code),
            'Width (mm)': sample_w,
            'Height (mm)': sample_h,
            'Qty': len(group),
            'Thickness': thick_type,
            'Glass Specification': glass_type,
            'ALL Window SQFT': round(all_sqft, 2),
            'Special glass SQFT': round(special_sqft, 2)
        })

    return pd.DataFrame(summary), target_sheet


# EXACT GLOBAL DASHBOARD CSS
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif !important;
        background-color: #f8fafc !important;
        color: #0f172a !important;
    }

    /* Sidebar Divider Line */
    [data-testid="stSidebar"] hr {
        margin-top: 1rem !important;
        margin-bottom: 1.5rem !important;
        border-color: #e2e8f0 !important;
    }
    
    /* Header Card */
    .header-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 24px 32px;
        margin-bottom: 24px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.02);
    }
    .main-title {
        font-size: 22px !important;
        font-weight: 800 !important;
        color: #0f172a;
        margin-bottom: 4px;
    }
    .main-subtitle {
        font-size: 13px !important;
        color: #64748b;
    }

    .step-header {
        font-size: 15px !important;
        font-weight: 700 !important;
        color: #1e293b;
        margin-bottom: 12px;
    }

    /* SOLID BUTTONS MATCHING ORIGINAL MODULE */
    div.btn-blue button {
        background-color: #2563eb !important;
        color: #ffffff !important;
        border-radius: 8px !important;
        height: 42px !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        border: none !important;
        width: 100% !important;
        box-shadow: 0 1px 2px rgba(0,0,0,0.05) !important;
    }
    div.btn-blue button:hover {
        background-color: #1d4ed8 !important;
    }

    div.btn-red button {
        background-color: #dc2626 !important;
        color: #ffffff !important;
        border-radius: 8px !important;
        height: 42px !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        border: none !important;
        width: 100% !important;
        box-shadow: 0 1px 2px rgba(0,0,0,0.05) !important;
    }
    div.btn-red button:hover {
        background-color: #b91c1c !important;
    }

    /* Force text/icons inside buttons to stay white */
    div.btn-blue button *, div.btn-red button * {
        color: #ffffff !important;
    }
    </style>
""", unsafe_allow_html=True)

# Session State
if 'df_result' not in st.session_state:
    st.session_state['df_result'] = None
if 'sheet_used' not in st.session_state:
    st.session_state['sheet_used'] = None

# Sidebar Matching Module 1 & 2 Layout Exactly
with st.sidebar:
    if os.path.exists("logo.png"):
        st.image("logo.png", width=170)
    
    st.markdown("---")  # Horizontal Line matching original theme
    
    st.markdown("#### ⚙️ Reading Mode Option")
    selected_mode = st.radio(
        "Select Sheet Format Reader:",
        ["Auto-Detect Format", "Option 1: Measurement Table", "Option 2: Quotation Block Layout"],
        help="Choose Option 1 for MEASUREMENT horizontal table sheet, Option 2 for WinSquare Quotation block sheet."
    )

# Header Card
st.markdown("""
    <div class="header-card">
        <div class="main-title">Universal Window Details & Glass SQFT Engine</div>
        <div class="main-subtitle">Supports Measurement Sheets, Quotation Sheets & Block Layouts</div>
    </div>
""", unsafe_allow_html=True)

# Step 1 Section
st.markdown('<div class="step-header">📁 Step 1: Upload BOQ Excel File</div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("", type=["xlsx", "xls"], label_visibility="collapsed")

# Buttons Row - Width & Color Match Original
st.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True)
col_p, col_r, _ = st.columns([2, 2, 4])

with col_p:
    st.markdown('<div class="btn-blue">', unsafe_allow_html=True)
    btn_process = st.button("🔗 Process Sheet", key="btn_process_sheet")
    st.markdown('</div>', unsafe_allow_html=True)

with col_r:
    st.markdown('<div class="btn-red">', unsafe_allow_html=True)
    btn_reset = st.button("🗑️ Reset Data", key="btn_reset_sheet")
    st.markdown('</div>', unsafe_allow_html=True)

# Reset Logic
if btn_reset:
    st.session_state['df_result'] = None
    st.session_state['sheet_used'] = None
    st.rerun()

# Process Logic
if btn_process:
    if uploaded_file is not None:
        try:
            with st.spinner("Processing file..."):
                df_res, used_sheet = process_excel_with_mode(uploaded_file, selected_mode)
                st.session_state['df_result'] = df_res
                st.session_state['sheet_used'] = used_sheet
        except Exception as e:
            st.error(f"Error parsing sheet: {str(e)}")
    else:
        st.warning("Please upload an Excel file first.")

# Results Display
if st.session_state['df_result'] is not None:
    res_df = st.session_state['df_result']
    used_sheet = st.session_state['sheet_used']

    st.markdown("<br>", unsafe_allow_html=True)
    st.success(f"Successfully processed sheet: **'{used_sheet}'**")

    if not res_df.empty:
        st.markdown("### 📑 Window Details Output Table")
        st.dataframe(res_df, use_container_width=True)

        st.markdown("<br>", unsafe_allow_html=True)
        m1, m2, m3 = st.columns(3)
        
        tot_all = res_df["ALL Window SQFT"].sum()
        tot_spec = res_df["Special glass SQFT"].sum()

        with m1:
            st.metric("Total Window Types", len(res_df))
        with m2:
            st.metric("Total ALL Window SQFT", f"{tot_all:,.2f} sqft")
        with m3:
            st.metric("Total Special Glass SQFT", f"{tot_spec:,.2f} sqft")
    else:
        st.warning("No valid window rows found in the sheet. Please check the selected Reading Mode Option in sidebar.")
