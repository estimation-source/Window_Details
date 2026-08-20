import os
import re
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

        # Detect Start of Block (Code :)
        if "CODE :" in row_str.upper() or "CODE:" in row_str.upper():
            code_val = ""
            name_val = ""
            glass_val = ""
            width_val = "-"
            height_val = "-"
            sqft_val = 0
            thick_val = "-"

            # Code
            code_match = re.search(r'CODE\s*:\s*([A-Za-z0-9_\-]+)', row_str, re.IGNORECASE)
            if code_match:
                code_val = code_match.group(1).strip()

            # Scan up to 25 rows below
            for r_offset in range(r, min(r + 25, num_rows)):
                sub_row = df_raw.iloc[r_offset]
                sub_vals = [str(v).strip() for v in sub_row.values if pd.notna(v) and str(v).strip() != "nan"]
                sub_str = " ".join(sub_vals)

                # Name
                if ("NAME :" in sub_str.upper() or "NAME:" in sub_str.upper()) and not name_val:
                    m = re.search(r'NAME\s*:\s*(.*?)(?=Profile System|Size|Location|$)', sub_str, re.IGNORECASE)
                    if m:
                        name_val = m.group(1).strip()

                # Glass
                if ("GLASS :" in sub_str.upper() or "GLASS:" in sub_str.upper()) and not glass_val:
                    m = re.search(r'GLASS\s*:\s*(.*)', sub_str, re.IGNORECASE)
                    if m:
                        glass_val = m.group(1).strip()
                        tm = re.search(r'(\d+\s*MM(?:\s*DGU)?)', glass_val, re.IGNORECASE)
                        if tm:
                            thick_val = tm.group(1).strip()

                # Width, Height, SQFT
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
        # Search all sheets for CODE:
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

# Sidebar Options
with st.sidebar:
    if logo_b64:
        st.markdown(f'<img src="data:image/png;base64,{logo_b64}" style="width:140px;">', unsafe_allow_html=True)
    st.markdown("### 🪟 Window Details Module")
    
    st.markdown("---")
    st.markdown("#### ⚙️ Reading Mode Option")
    selected_mode = st.radio(
        "Select Sheet Format Reader:",
        ["Auto-Detect Format", "Option 1: Measurement Table", "Option 2: Quotation Block Layout"],
        help="Choose Option 1 for MEASUREMENT horizontal table sheet, Option 2 for WinSquare Quotation block sheet."
    )

# Header
st.markdown("""
    <div class="header-container">
        <div class="main-title">Multi-Format Window Details Engine</div>
        <div class="main-subtitle">Select Reader Option for Measurement Table OR Quotation Block Sheets</div>
    </div>
""", unsafe_allow_html=True)

# File Upload
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

if uploaded_file:
    try:
        result_df, sheet_used = process_excel_with_mode(uploaded_file, selected_mode)
        
        st.success(f"Mode Selected: **{selected_mode}** | Processed Sheet: **'{sheet_used}'**")
        
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
            st.warning("No valid window rows found in the sheet. Try changing the Reading Mode option from sidebar.")

    except Exception as e:
        st.error(f"Error parsing sheet: {str(e)}")
