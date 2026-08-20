from __future__ import annotations

import io
import os
import re
import openpyxl
import pandas as pd
from PIL import Image
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ============================================================
# 1. Page Config
# ============================================================
st.set_page_config(
    page_title="Universal Window Details",
    page_icon="🪟",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Helper Function for File Path
def get_image_path(filename: str) -> str:
    return os.path.join(os.path.abspath("."), filename)

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
        glass_spec = str(row.iloc[11]).strip() if len(row) > 11 and pd.notna(row.iloc[11]) and str(row.iloc[11]).strip() != "nan" else "Standard Glass"

        if pd.isna(sqft) or sqft <= 0:
            continue

        rows.append({
            'Window Code / Type': window_name,
            'Width (mm)': width if pd.notna(width) else "-",
            'Height (mm)': height if pd.notna(height) else "-",
            'Thickness': thickness,
            'Glass Specification': glass_spec,
            'SQFT': sqft,
            'SourceFile': getattr(file_obj, 'name', 'BOQ File')
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
            glass_val = "Standard Glass"
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

                if ("GLASS :" in sub_str.upper() or "GLASS:" in sub_str.upper()) and glass_val == "Standard Glass":
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
                    'SQFT': sqft_val,
                    'SourceFile': getattr(file_obj, 'name', 'BOQ File')
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
        return pd.DataFrame(), pd.DataFrame(), target_sheet

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

    return pd.DataFrame(summary), df_clean, target_sheet


# PAGE AND ELEMENT STYLING
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif !important;
        background-color: #f8fafc !important;
        color: #0f172a !important;
    }

    [data-testid="stSidebar"] {
        background-color: #f1f5f9 !important;
        border-right: 1px solid #e2e8f0 !important;
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

    /* KPI CARDS */
    .kpi-card-box {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 16px 20px;
        box-shadow: 0 1px 2px rgba(0,0,0,0.03);
    }

    .kpi-title-lbl {
        font-size: 11px;
        font-weight: 700;
        color: #64748b;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }

    .kpi-val-lbl {
        font-size: 24px;
        font-weight: 800;
        color: #0f172a;
        margin-top: 6px;
    }

    /* FORCE BLUE BUTTON (PRIMARY TYPE) */
    .stButton > button[kind="primary"] {
        background-color: #2563eb !important;
        background: #2563eb !important;
        border: 1px solid #2563eb !important;
        color: #ffffff !important;
        border-radius: 6px !important;
        height: 38px !important;
        padding: 0 16px !important;
        box-shadow: 0 1px 2px rgba(37, 99, 235, 0.2) !important;
        white-space: nowrap !important;
    }
    .stButton > button[kind="primary"]:hover {
        background-color: #1d4ed8 !important;
        background: #1d4ed8 !important;
    }

    /* FORCE RED BUTTON (SECONDARY TYPE OVERRIDE) */
    .stButton > button[kind="secondary"] {
        background-color: #dc2626 !important;
        background: #dc2626 !important;
        border: 1px solid #dc2626 !important;
        color: #ffffff !important;
        border-radius: 6px !important;
        height: 38px !important;
        padding: 0 16px !important;
        box-shadow: 0 1px 2px rgba(220, 38, 38, 0.2) !important;
        white-space: nowrap !important;
    }
    .stButton > button[kind="secondary"]:hover {
        background-color: #b91c1c !important;
        background: #b91c1c !important;
    }

    /* DOWNLOAD GREEN BUTTON */
    div.stDownloadButton > button {
        background-color: #059669 !important;
        background: #059669 !important;
        border: 1px solid #059669 !important;
        color: #ffffff !important;
        border-radius: 6px !important;
        height: 38px !important;
        padding: 0 16px !important;
        box-shadow: 0 1px 2px rgba(5, 150, 105, 0.2) !important;
    }
    div.stDownloadButton > button:hover {
        background-color: #047857 !important;
        background: #047857 !important;
    }

    .stButton > button p, .stButton > button span,
    div.stDownloadButton > button p, div.stDownloadButton > button span {
        color: #ffffff !important;
        font-weight: 500 !important;
        font-size: 14px !important;
    }
    </style>
""", unsafe_allow_html=True)

# Session State Initializations
if 'df_result' not in st.session_state:
    st.session_state['df_result'] = None
if 'df_raw_clean' not in st.session_state:
    st.session_state['df_raw_clean'] = None
if 'sheet_used' not in st.session_state:
    st.session_state['sheet_used'] = None

# SIDEBAR LOGO AND OPTIONS
with st.sidebar:
    logo_file = get_image_path("logo.png")
    if os.path.exists(logo_file):
        col_s1, col_s2, col_s3 = st.columns([1, 2, 1])
        with col_s2:
            st.image(Image.open(logo_file), width=110)
    else:
        st.markdown("<h2 style='text-align: center; color:#1e293b;'><b>win square</b></h2>", unsafe_allow_html=True)
    
    st.markdown("---")
    
    st.markdown("#### ⚙️ Reading Mode Option")
    selected_mode = st.radio(
        "Select Sheet Format Reader:",
        ["Auto-Detect Format", "Option 1: Measurement Table", "Option 2: Quotation Block Layout"],
        help="Choose Option 1 for MEASUREMENT horizontal table sheet, Option 2 for WinSquare Quotation block sheet."
    )

# Header Card
st.markdown("""
    <div class="header-card">
        <div class="main-title">Universal Window Details</div>
        <div class="main-subtitle">Supports Measurement Sheets, Quotation Sheets & Block Layouts</div>
    </div>
""", unsafe_allow_html=True)

# Step 1 Section
st.markdown('<div class="step-header">📁 Step 1: Upload BOQ Excel File</div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("", type=["xlsx", "xls"], label_visibility="collapsed")

st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)

# STRICT COMPACT COLUMN RATIO FOR BUTTONS TO SIT CLOSE TOGETHER
btn_col1, btn_col2, _ = st.columns([1.1, 1.1, 8])

with btn_col1:
    btn_process = st.button("🔗 Process Sheet", type="primary", use_container_width=True)

with btn_col2:
    btn_reset = st.button("🗑️ Reset Data", type="secondary", use_container_width=True)

# Reset Logic
if btn_reset:
    st.session_state['df_result'] = None
    st.session_state['df_raw_clean'] = None
    st.session_state['sheet_used'] = None
    st.rerun()

# Process Logic
if btn_process:
    if uploaded_file is not None:
        try:
            with st.spinner("Processing file..."):
                df_res, df_raw_c, used_sheet = process_excel_with_mode(uploaded_file, selected_mode)
                st.session_state['df_result'] = df_res
                st.session_state['df_raw_clean'] = df_raw_c
                st.session_state['sheet_used'] = used_sheet
        except Exception as e:
            st.error(f"Error parsing sheet: {str(e)}")
    else:
        st.warning("Please upload an Excel file first.")

# Results Display & Requirement Sheet Dashboard
if st.session_state['df_result'] is not None:
    res_df = st.session_state['df_result']
    df_raw_c = st.session_state['df_raw_clean']
    used_sheet = st.session_state['sheet_used']

    st.markdown("<br>", unsafe_allow_html=True)
    st.success(f"Successfully processed sheet: **'{used_sheet}'**")

    if not res_df.empty:
        # Dashboard KPI Cards
        st.markdown('<div class="step-header">📊 Step 2: Requirement Sheet Dashboard Analytics</div>', unsafe_allow_html=True)
        
        tot_types = len(res_df)
        tot_qty = res_df["Qty"].sum()
        tot_all_sqft = res_df["ALL Window SQFT"].sum()
        tot_spec_sqft = res_df["Special glass SQFT"].sum()

        k1, k2, k3, k4 = st.columns(4)
        with k1:
            st.markdown(f"<div class='kpi-card-box'><div class='kpi-title-lbl'>TOTAL WINDOW TYPES</div><div class='kpi-val-lbl'>{tot_types}</div></div>", unsafe_allow_html=True)
        with k2:
            st.markdown(f"<div class='kpi-card-box'><div class='kpi-title-lbl'>TOTAL QUANTITY</div><div class='kpi-val-lbl'>{tot_qty} Pcs</div></div>", unsafe_allow_html=True)
        with k3:
            st.markdown(f"<div class='kpi-card-box'><div class='kpi-title-lbl'>TOTAL ALL WINDOW SQFT</div><div class='kpi-val-lbl'>{tot_all_sqft:,.2f}</div></div>", unsafe_allow_html=True)
        with k4:
            st.markdown(f"<div class='kpi-card-box'><div class='kpi-title-lbl'>SPECIAL GLASS SQFT</div><div class='kpi-val-lbl'>{tot_spec_sqft:,.2f}</div></div>", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # 3 VIEW TABS
        tab1, tab2, tab3 = st.tabs([
            "📄 Window Details Live Preview", 
            "📊 File / OC Summary", 
            "🧩 Glass Specification Breakdown"
        ])

        with tab1:
            st.dataframe(res_df, use_container_width=True, hide_index=True)

        with tab2:
            # File / OC Summary Table
            file_summary = (
                df_raw_c.groupby("SourceFile", as_index=False)
                .agg(
                    Total_Windows=("Window Code / Type", "count"),
                    Total_SQFT=("SQFT", "sum")
                )
            )
            file_summary["Total_SQFT"] = file_summary["Total_SQFT"].round(2)
            file_summary.columns = ["Source File / OC Name", "Total Windows (Pcs)", "Total SQFT"]
            file_summary.insert(0, "Sr. No.", range(1, len(file_summary) + 1))
            st.dataframe(file_summary, use_container_width=True, hide_index=True)

        with tab3:
            # Glass Type Breakdown Table
            glass_summary = (
                df_raw_c.groupby("Glass Specification", as_index=False)
                .agg(
                    Total_Pcs=("Window Code / Type", "count"),
                    Total_SQFT=("SQFT", "sum")
                )
                .sort_values(by="Total_Pcs", ascending=False)
            )
            glass_summary["Total_SQFT"] = glass_summary["Total_SQFT"].round(2)
            glass_summary.columns = ["Glass Specification", "Quantity (Pcs)", "Total SQFT"]
            glass_summary.insert(0, "Sr. No.", range(1, len(glass_summary) + 1))
            st.dataframe(glass_summary, use_container_width=True, hide_index=True)

        # Excel Download Functionality (Calibri / Blue Styling)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "WINDOW DETAILS"
        ws.views.sheetView[0].showGridLines = True

        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        data_font = Font(name="Calibri", size=10)
        total_font = Font(name="Calibri", size=11, bold=True)
        header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        total_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
        thin_border = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9"),
        )

        headers = list(res_df.columns)
        ws.append(headers)

        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_num)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")

        for row_idx, row in res_df.iterrows():
            ws.append(list(row))
            for col_num in range(1, len(headers) + 1):
                cell = ws.cell(row=row_idx + 2, column=col_num)
                cell.font = data_font
                cell.border = thin_border

        # Total Row
        tot_row_num = len(res_df) + 2
        ws.cell(row=tot_row_num, column=1, value="TOTAL").font = total_font
        ws.cell(row=tot_row_num, column=4, value=f"=SUM(D2:D{tot_row_num-1})").font = total_font
        ws.cell(row=tot_row_num, column=7, value=f"=SUM(G2:G{tot_row_num-1})").font = total_font
        ws.cell(row=tot_row_num, column=8, value=f"=SUM(H2:H{tot_row_num-1})").font = total_font

        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=tot_row_num, column=col_num)
            cell.fill = total_fill

        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

        output = io.BytesIO()
        wb.save(output)
        excel_bytes = output.getvalue()

        st.markdown("<br>", unsafe_allow_html=True)
        st.download_button(
            label="📥 DOWNLOAD WINDOW DETAILS SHEET (.XLSX)",
            data=excel_bytes,
            file_name="WINDOW_DETAILS_SUMMARY.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=False
        )
    else:
        st.warning("No valid window rows found in the sheet. Please check the selected Reading Mode Option in sidebar.")
