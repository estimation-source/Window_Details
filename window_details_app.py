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

# Smart Block-Layout & Table Parser
def process_excel_file(file_obj):
    excel_file = pd.ExcelFile(file_obj)
    sheet_names = excel_file.sheet_names
    
    # Target Sheet Selection
    target_sheet = None
    for s in sheet_names:
        if "MEASUREMENT" in str(s).upper():
            target_sheet = s
            break
    if not target_sheet:
        target_sheet = sheet_names[1] if len(sheet_names) >= 2 else sheet_names[0]

    df_raw = pd.read_excel(file_obj, sheet_name=target_sheet, header=None)
    full_text = " ".join([str(val) for val in df_raw.fillna('').values.flatten()]).upper()
    
    rows = []

    # -------------------------------------------------------------------
    # FORMAT A: VERTICAL BLOCK FORMAT (Quotation Sheets / Sheet2 Layout)
    # -------------------------------------------------------------------
    if "CODE :" in full_text or "CODE:" in full_text or ("WIDTH" in full_text and "SQFT" in full_text):
        blocks = []
        current_block = {}

        for idx, row in df_raw.iterrows():
            # Get non-empty values
            row_vals = [str(val).strip() for val in row.values if pd.notna(val) and str(val).strip() != "nan" and str(val).strip() != ""]
            row_str = " ".join(row_vals)

            if not row_vals:
                continue

            # Detect New Window Block (CODE :)
            if re.search(r'CODE\s*:', row_str, re.IGNORECASE):
                if current_block.get('SQFT') or current_block.get('Code'):
                    blocks.append(current_block)
                current_block = {'Code': '', 'Name': '', 'Glass': '', 'Width': '-', 'Height': '-', 'Thickness': '-', 'SQFT': 0}

                match = re.search(r'CODE\s*:\s*([A-Za-z0-9_\-]+)', row_str, re.IGNORECASE)
                if match:
                    current_block['Code'] = match.group(1).strip()

            # Detect Name
            if re.search(r'NAME\s*:', row_str, re.IGNORECASE):
                match = re.search(r'NAME\s*:\s*(.*?)(?=Profile System|Size|Location|$)', row_str, re.IGNORECASE)
                if match:
                    current_block['Name'] = match.group(1).strip()

            # Detect Glass
            if re.search(r'GLASS\s*:', row_str, re.IGNORECASE):
                match = re.search(r'GLASS\s*:\s*(.*)', row_str, re.IGNORECASE)
                if match:
                    glass_val = match.group(1).strip()
                    current_block['Glass'] = glass_val
                    # Extract Thickness if specified (e.g. 20mm, 5mm)
                    thick_match = re.search(r'(\d+\s*MM(?:\s*DGU)?)', glass_val, re.IGNORECASE)
                    if thick_match:
                        current_block['Thickness'] = thick_match.group(1).strip()

            # Detect Width, Height, SQFT line-by-line
            for i, val in enumerate(row_vals):
                val_u = val.upper()
                if val_u == "WIDTH" or "WIDTH" in val_u:
                    # Look ahead in same row for number
                    nums = [re.sub(r'[^0-9.]', '', x) for x in row_vals[i+1:] if re.sub(r'[^0-9.]', '', x)]
                    if nums:
                        current_block['Width'] = pd.to_numeric(nums[0], errors='coerce')
                
                elif val_u == "HEIGHT" or "HEIGHT" in val_u:
                    nums = [re.sub(r'[^0-9.]', '', x) for x in row_vals[i+1:] if re.sub(r'[^0-9.]', '', x)]
                    if nums:
                        current_block['Height'] = pd.to_numeric(nums[0], errors='coerce')

                elif val_u == "SQFT" or "SQFT" in val_u:
                    nums = [re.sub(r'[^0-9.]', '', x) for x in row_vals[i+1:] if re.sub(r'[^0-9.]', '', x)]
                    if nums:
                        sq_val = pd.to_numeric(nums[0], errors='coerce')
                        if pd.notna(sq_val):
                            current_block['SQFT'] = sq_val

        if current_block.get('SQFT') or current_block.get('Code'):
            blocks.append(current_block)

        # Map blocks to output table structure
        for b in blocks:
            code_name = b.get('Code', '')
            if b.get('Name'):
                code_name = f"{code_name} - {b['Name']}" if code_name else b['Name']
            
            sqft_val = b.get('SQFT', 0)
            if sqft_val > 0:
                rows.append({
                    'Window Code / Type': code_name if code_name else "Window",
                    'Width (mm)': b.get('Width', '-'),
                    'Height (mm)': b.get('Height', '-'),
                    'Thickness': b.get('Thickness', '-'),
                    'Glass Specification': b.get('Glass', ''),
                    'SQFT': sqft_val
                })

    # -------------------------------------------------------------------
    # FORMAT B: HORIZONTAL TABLE FORMAT (Measurement Sheet Layout)
    # -------------------------------------------------------------------
    if not rows:
        start_row = 0
        for idx, row in df_raw.iterrows():
            val0 = str(row[0]).strip()
            val1 = str(row[1]).strip()
            if val0.isdigit() or val1.isdigit():
                start_row = idx
                break

        df_data = df_raw.iloc[start_row:].copy().reset_index(drop=True)

        for _, row in df_data.iterrows():
            win_type = str(row[2]).strip() if pd.notna(row[2]) else ""
            location = str(row[3]).strip() if pd.notna(row[3]) else ""
            
            window_name = f"{win_type} ({location})" if location and location != "nan" else win_type
            if not window_name or window_name == "nan":
                continue

            width = pd.to_numeric(row[5], errors='coerce')
            height = pd.to_numeric(row[6], errors='coerce')
            sqft = pd.to_numeric(row[7], errors='coerce')
            thickness = str(row[10]).strip() if pd.notna(row[10]) and str(row[10]).strip() != "nan" else "-"
            glass_spec = str(row[11]).strip() if pd.notna(row[11]) and str(row[11]).strip() != "nan" else ""

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

    # Final DataFrame Processing
    df_clean = pd.DataFrame(rows)
    if df_clean.empty:
        return pd.DataFrame(), target_sheet

    # Glass Classification Rule
    def check_special_glass(spec):
        glass_lower = str(spec).lower()
        if "frosted" in glass_lower and "toughened" not in glass_lower and "tough" not in glass_lower:
            return False
        elif "toughened" in glass_lower or "tough" in glass_lower or "dgu" in glass_lower or "satin" in glass_lower:
            return True
        return False

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

# Sidebar
with st.sidebar:
    if logo_b64:
        st.markdown(f'<img src="data:image/png;base64,{logo_b64}" style="width:140px;">', unsafe_allow_html=True)
    st.markdown("### 🪟 Window Details Module")
    st.caption("Auto-supports Horizontal Tables & Vertical Block Quotation Layouts.")

# Header
st.markdown("""
    <div class="header-container">
        <div class="main-title">Universal Window Details & Glass SQFT Engine</div>
        <div class="main-subtitle">Supports Measurement Sheets, Quotation Sheets & Block Layouts</div>
    </div>
""", unsafe_allow_html=True)

# File Upload
uploaded_file = st.file_uploader("Upload Excel BOQ File", type=["xlsx", "xls"])

if uploaded_file:
    try:
        result_df, sheet_used = process_excel_file(uploaded_file)
        
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
