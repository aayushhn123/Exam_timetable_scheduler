import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from fpdf import FPDF
import os
import re
import io
from PyPDF2 import PdfReader, PdfWriter

# ==========================================
# ⚙️ PAGE CONFIGURATION
# ==========================================
st.set_page_config(
    page_title="Verification to PDF Converter",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Compatibility for Dialogs
if hasattr(st, "dialog"):
    dialog_decorator = st.dialog
else:
    dialog_decorator = st.experimental_dialog

# ==========================================
# 🏫 COLLEGE CONFIGURATION (LATEST)
# ==========================================
COLLEGES = [
    {"name": "Mukesh Patel School of Technology Management & Engineering / School of Technology Management & Engineering", "icon": "🖥️"},
    {"name": "School of Business Management", "icon": "💼"},
    {"name": "Pravin Dalal School of Entrepreneurship & Family Business Management", "icon": "🚀"},
    {"name": "Anil Surendra Modi School of Commerce / School of Commerce", "icon": "📊"},
    {"name": "Kirit P. Mehta School of Law / School of Law", "icon": "⚖️"},
    {"name": "Shobhaben Pratapbhai Patel School of Pharmacy & Technology Management / School of Pharmacy & Technology Management", "icon": "💊"},
    {"name": "Sunandan Divatia School of Science / School of Science", "icon": "🔬"},
    {"name": "Sarla Anil Modi School of Economics", "icon": "📈"},
    {"name": "Balwant Sheth School of Architecture", "icon": "🏛️"},
    {"name": "School of Design", "icon": "🎨"},
    {"name": "Jyoti Dalal School of Liberal Arts", "icon": "📚"},
    {"name": "School of Performing Arts", "icon": "🎭"},
    {"name": "School of Hospitality Management", "icon": "🏨"},
    {"name": "School of Mathematics, Applied Statistics & Analytics", "icon": "📐"},
    {"name": "School of Branding and Advertising", "icon": "📢"},
    {"name": "School of Agricultural Sciences & Technology", "icon": "🌾"},
    {"name": "Centre of Distance and Online Education", "icon": "💻"},
    {"name": "School of Aviation", "icon": "✈️"}
]

BRANCH_FULL_FORM = {
    "B TECH": "BACHELOR OF TECHNOLOGY",
    "B TECH INTG": "BACHELOR OF TECHNOLOGY SIX YEAR INTEGRATED PROGRAM",
    "M TECH": "MASTER OF TECHNOLOGY",
    "MBA TECH": "MASTER OF BUSINESS ADMINISTRATION IN TECHNOLOGY MANAGEMENT",
    "MCA": "MASTER OF COMPUTER APPLICATIONS",
    "DIPLOMA": "DIPLOMA IN ENGINEERING"
}

LOGO_PATH = "logo.png"
wrap_text_cache = {}

# ==========================================
# 🎨 UI & CSS OPTIMIZATION
# ==========================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    * { font-family: 'Inter', sans-serif; }
    
    .main-header {
        background: linear-gradient(135deg, #951C1C 0%, #C73E1D 100%);
        padding: 2.5rem;
        border-radius: 16px;
        margin-bottom: 2rem;
        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.12);
    }
    .main-header h1 {
        color: white; text-align: center; margin: 0; font-size: 2.5rem; font-weight: 700;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.3); letter-spacing: -0.5px;
    }
    .main-header p {
        color: #FFF; text-align: center; margin: 0.5rem 0 0 0; font-size: 1.1rem; opacity: 0.95; font-weight: 500;
    }
    .upload-section {
        background: #f8f9fa; padding: 2.5rem; border-radius: 16px; border: 2px dashed #951C1C;
        margin: 1rem 0; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08); text-align: center;
    }
    
    @media (prefers-color-scheme: dark) {
        .main-header { background: linear-gradient(135deg, #701515 0%, #A23217 100%); }
        .upload-section { background: #2d2d2d; border-color: #A23217; }
    }
    
    .stButton>button {
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        border-radius: 12px; font-weight: 600;
    }
    .stButton>button:hover {
        transform: translateY(-2px); box-shadow: 0 8px 16px rgba(149, 28, 28, 0.3); border-color: #951C1C;
    }
</style>
""", unsafe_allow_html=True)


# ==========================================
# 📥 VERIFICATION DATA PARSING
# ==========================================
def process_verification_file(uploaded_file):
    try:
        # Check if Verification sheet exists
        xls = pd.ExcelFile(uploaded_file)
        if "Verification" in xls.sheet_names:
            df = pd.read_excel(uploaded_file, sheet_name="Verification")
        else:
            df = pd.read_excel(uploaded_file) # Fallback to first sheet

        # Identify required columns flexibly
        required_columns = [
            "Module Abbreviation", "Module Description", "Program", "Stream", 
            "Current Session", "Exam Date", "Exam Time", "Scheduling Status", "Subject Type"
        ]
        
        # Verify essential columns exist
        missing = [col for col in ["Program", "Current Session", "Exam Date"] if col not in df.columns]
        if missing:
            st.error(f"❌ Missing critical columns: {', '.join(missing)}")
            return None, None

        # Filter out unscheduled subjects
        if "Scheduling Status" in df.columns:
            df = df[df["Scheduling Status"].astype(str).str.strip().str.upper() == "SCHEDULED"].copy()
        
        df = df[df['Exam Date'].notna() & (df['Exam Date'].astype(str).str.strip().str.upper() != 'NOT SCHEDULED')].copy()
        
        # Map fields
        df['MainBranch'] = df.get('Program', pd.Series(dtype=str)).fillna("").astype(str).str.strip()
        df['SubBranch'] = df.get('Stream', pd.Series(dtype=str)).fillna("").astype(str).str.strip()
        df['SubBranch'] = df.apply(lambda x: x['MainBranch'] if x['SubBranch'] in ['nan', ''] else x['SubBranch'], axis=1)
        
        df['Subject'] = df.get('Module Description', pd.Series(dtype=str)).fillna("").astype(str).str.strip()
        df['ModuleCode'] = df.get('Module Abbreviation', pd.Series(dtype=str)).fillna("").astype(str).str.strip()
        df['Exam Time'] = df.get('Exam Time', pd.Series(dtype=str)).fillna("").astype(str).str.strip()
        
        # OE detection
        if 'Subject Type' in df.columns:
            df['OE'] = df['Subject Type'].apply(lambda x: 'OE' if str(x).strip().upper() == 'OE' else None)
        else:
            df['OE'] = None

        # Determine numeric semester
        def get_sem_int(val):
            s = str(val).upper().strip()
            m = re.search(r'(\d+)', s)
            if m: return int(m.group(1))
            roman_map = {'I':1, 'II':2, 'III':3, 'IV':4, 'V':5, 'VI':6, 'VII':7, 'VIII':8, 'IX':9, 'X':10}
            for r, i in roman_map.items():
                if r == s or s.endswith(f" {r}") or s.endswith(f"_{r}"): return i
            return 1
            
        df['Semester'] = df.get('Current Session', pd.Series([1]*len(df))).apply(get_sem_int)
        
        timetable = {}
        for sem in sorted(df['Semester'].unique()):
            timetable[sem] = df[df['Semester'] == sem].copy()
            
        return timetable, df

    except Exception as e:
        st.error(f"Error parsing verification file: {e}")
        return None, None

# ==========================================
# 💾 EXCEL ENGINE (FORMATTING FOR PDF GENERATOR)
# ==========================================
def normalize_time(t_str):
    if not isinstance(t_str, str): return ""
    t_str = t_str.strip().upper()
    for i in range(1, 10): t_str = t_str.replace(f"0{i}:", f"{i}:")
    return t_str

def int_to_roman(num):
    val = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]
    res = ""
    for v, r in val:
        while num >= v: res += r; num -= v
    return res

def save_to_excel(semester_wise_timetable):
    time_slots_dict = {1: {"start": "10:00 AM", "end": "1:00 PM"}, 2: {"start": "2:00 PM", "end": "5:00 PM"}}
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        sheets_created = 0
        for sem, df_sem in semester_wise_timetable.items():
            if df_sem.empty: continue
            
            # Predict default slot based on standard grouping
            slot_id = 1 if ((sem + 1) // 2) % 2 == 1 else 2
            p_cfg = time_slots_dict.get(slot_id, time_slots_dict[1])
            header_norm = normalize_time(f"{p_cfg['start']} - {p_cfg['end']}")
            
            for main_branch in df_sem['MainBranch'].unique():
                df_mb = df_sem[df_sem['MainBranch'] == main_branch].copy()
                roman_sem = int_to_roman(sem)
                
                # Sheet names must be <= 31 chars
                sheet_base = f"{main_branch}"[:20] 
                core_sheet = f"{sheet_base}_|_Sem {roman_sem}"[:31]
                elec_sheet = f"{sheet_base}_|_Sem {roman_sem}_Ele"[:31]
                
                df_core = df_mb[df_mb['OE'].isna()].copy()
                df_elec = df_mb[df_mb['OE'].notna()].copy()
                
                # Process Core
                if not df_core.empty:
                    displays = []
                    for _, row in df_core.iterrows():
                        subj = row['Subject']
                        code = row['ModuleCode']
                        actual_time = row.get('Exam Time', '')
                        
                        time_suffix = ""
                        if actual_time and normalize_time(actual_time) != header_norm and actual_time.lower() != 'tbd':
                            time_suffix = f" [{actual_time}]"
                            
                        txt = f"{subj}"
                        if code and str(code).lower() != 'nan': txt += f" - ({code})"
                        txt += time_suffix
                        displays.append(txt)
                    
                    df_core['SubjectDisplay'] = displays
                    df_core["Exam Date"] = pd.to_datetime(df_core["Exam Date"], format="%d-%m-%Y", dayfirst=True, errors='coerce')
                    df_core = df_core.sort_values(by="Exam Date", ascending=True)
                    
                    try:
                        pivot = df_core.groupby(['Exam Date', 'SubBranch']).agg({'SubjectDisplay': lambda x: ", ".join(str(i) for i in x)}).reset_index()
                        pivot = pivot.pivot_table(index='Exam Date', columns='SubBranch', values='SubjectDisplay', aggfunc='first').fillna("---")
                        pivot = pivot.sort_index(ascending=True).reset_index()
                        pivot['Exam Date'] = pivot['Exam Date'].apply(lambda x: x.strftime("%d-%m-%Y") if pd.notna(x) else "")
                        
                        pivot['Program'] = main_branch
                        pivot['Semester'] = roman_sem
                        
                        pivot.to_excel(writer, sheet_name=core_sheet, index=False)
                        sheets_created += 1
                    except Exception as e:
                        pass
                
                # Process Electives
                if not df_elec.empty:
                    e_displays = []
                    for _, row in df_elec.iterrows():
                        subj = row['Subject']
                        code = row['ModuleCode']
                        actual_time = row.get('Exam Time', '')
                        
                        time_suffix = ""
                        if actual_time and normalize_time(actual_time) != header_norm and actual_time.lower() != 'tbd':
                            time_suffix = f" [{actual_time}]"
                            
                        txt = f"{subj}"
                        if code and str(code).lower() != 'nan': txt += f" - ({code})"
                        txt += time_suffix
                        e_displays.append(txt)
                        
                    df_elec['DisplaySubject'] = e_displays
                    
                    try:
                        df_elec["Exam Date"] = pd.to_datetime(df_elec["Exam Date"], format="%d-%m-%Y", dayfirst=True, errors='coerce')
                        df_elec = df_elec.sort_values(by="Exam Date", ascending=True)
                        df_elec['Exam Date'] = df_elec['Exam Date'].apply(lambda x: x.strftime("%d-%m-%Y") if pd.notna(x) else "")
                        
                        ep = df_elec.groupby(['Exam Date', 'OE']).agg({'DisplaySubject': lambda x: ", ".join(sorted(set(x)))}).reset_index()
                        ep.rename(columns={'OE': 'OE Type', 'DisplaySubject': 'Subjects'}, inplace=True)
                        
                        ep['Program'] = main_branch
                        ep['Semester'] = roman_sem
                        
                        ep.to_excel(writer, sheet_name=elec_sheet, index=False)
                        sheets_created += 1
                    except Exception as e:
                        pass
                        
        if sheets_created == 0:
            pd.DataFrame({'Info': ['No valid data']}).to_excel(writer, sheet_name="Empty")
            
    output.seek(0)
    return output

# ==========================================
# 📄 FPDF STRICT PDF ENGINE (MATCHING EXACTLY)
# ==========================================
def wrap_text(pdf, text, col_width):
    cache_key = (text, col_width, pdf.font_style)
    if cache_key in wrap_text_cache:
        return wrap_text_cache[cache_key]
        
    time_pattern = re.compile(r'([\[\(]\s*\d{1,2}:\d{2}\s*[AP]M\s*-\s*\d{1,2}:\d{2}\s*[AP]M\s*[\]\)])', re.IGNORECASE)
    parts = time_pattern.split(str(text))
    tokens = []
    for i, p in enumerate(parts):
        if i % 2 == 1: tokens.append(p)
        else: tokens.extend(p.split())
            
    lines = []
    current_line = ""
    
    old_family = pdf.font_family
    old_style = pdf.font_style
    old_size = pdf.font_size_pt
    
    for token in tokens:
        test_line = token if not current_line else current_line + " " + token
        test_w = 0
        for pt in time_pattern.split(test_line):
            if not pt: continue
            if time_pattern.match(pt):
                pdf.set_font(old_family, 'B', old_size)
                test_w += pdf.get_string_width(pt)
            else:
                pdf.set_font(old_family, old_style, old_size)
                test_w += pdf.get_string_width(pt)
                
        if test_w <= col_width:
            current_line = test_line
        else:
            if current_line: lines.append(current_line)
            current_line = token
            
    if current_line: lines.append(current_line)
    pdf.set_font(old_family, old_style, old_size)
    wrap_text_cache[cache_key] = lines
    return lines

def print_row_custom(pdf, row_data, col_widths, line_height=6, header=False):
    cell_padding = 2
    header_bg_color = (149, 33, 28)
    header_text_color = (255, 255, 255)
    alt_row_color = (255, 255, 255) # NO GREY ROWS: Disabled by setting to white
    
    row_number = getattr(pdf, '_row_counter', 0)
    
    base_font = "Arial"
    if header:
        base_style = 'B'
        base_size = 9
        pdf.set_font(base_font, base_style, base_size)
        pdf.set_text_color(*header_text_color)
        pdf.set_fill_color(*header_bg_color)
    else:
        base_style = ''
        base_size = 8
        pdf.set_font(base_font, base_style, base_size)
        pdf.set_text_color(0, 0, 0)
        pdf.set_fill_color(*alt_row_color)

    wrapped_cells = []
    max_lines = 0
    for i, cell_text in enumerate(row_data):
        text = str(cell_text) if cell_text is not None else ""
        avail_w = col_widths[i] - 2 * cell_padding
        lines = wrap_text(pdf, text, avail_w)
        wrapped_cells.append(lines)
        max_lines = max(max_lines, len(lines))

    row_h = line_height * max_lines
    x0, y0 = pdf.get_x(), pdf.get_y()
    
    pdf.rect(x0, y0, sum(col_widths), row_h, 'F')

    time_pattern = re.compile(r'([\[\(]\s*\d{1,2}:\d{2}\s*[AP]M\s*-\s*\d{1,2}:\d{2}\s*[AP]M\s*[\]\)])', re.IGNORECASE)

    for i, lines in enumerate(wrapped_cells):
        cx = pdf.get_x()
        pad_v = (row_h - len(lines) * line_height) / 2 if len(lines) < max_lines else 0
        for j, ln in enumerate(lines):
            parts = time_pattern.split(ln)
            if len(parts) == 1 or header:
                pdf.set_xy(cx + cell_padding, y0 + j * line_height + pad_v)
                pdf.cell(col_widths[i] - 2 * cell_padding, line_height, ln, border=0, align='C')
            else:
                total_w = 0
                for k, p in enumerate(parts):
                    if not p: continue
                    if k % 2 == 1: pdf.set_font(base_font, 'B', base_size)
                    else: pdf.set_font(base_font, base_style, base_size)
                    total_w += pdf.get_string_width(p)
                
                start_x = cx + max(cell_padding, (col_widths[i] - total_w) / 2)
                current_x = start_x
                
                for k, p in enumerate(parts):
                    if not p: continue
                    if k % 2 == 1: pdf.set_font(base_font, 'B', base_size)
                    else: pdf.set_font(base_font, base_style, base_size)
                    w = pdf.get_string_width(p)
                    
                    # MATHEMATICAL FIX (Subtract margin, add margin * 2)
                    pdf.set_xy(current_x - pdf.c_margin, y0 + j * line_height + pad_v)
                    pdf.cell(w + 2 * pdf.c_margin, line_height, p, border=0, align='L')
                    current_x += w
                
                pdf.set_font(base_font, base_style, base_size)
        
        pdf.rect(cx, y0, col_widths[i], row_h)
        pdf.set_xy(cx + col_widths[i], y0)

    setattr(pdf, '_row_counter', row_number + 1)
    pdf.set_xy(x0, y0 + row_h)


def print_table_custom(pdf, df, columns, col_widths, line_height=6, header_content=None, college_name=None, time_slot=None, declaration_date=None):
    if df.empty: return
    setattr(pdf, '_row_counter', 0)
    
    footer_height = 18   
    
    def render_footer():
        pdf.set_xy(10, pdf.h - footer_height)
        pdf.set_font("Arial", 'B', 9) 
        pdf.cell(0, 5, "Controller of Examinations", 0, 1, 'L')
        pdf.line(10, pdf.h - footer_height + 5, 60, pdf.h - footer_height + 5)
        
        pdf.set_font("Arial", size=9)
        pdf.set_text_color(0, 0, 0)
        page_text = f"{pdf.page_no()} of {{nb}}"
        text_width = pdf.get_string_width(page_text.replace("{nb}", "99"))
        pdf.set_xy(pdf.w - 10 - text_width, pdf.h - footer_height + 5)
        pdf.cell(text_width, 5, page_text, 0, 0, 'R')

    def render_header():
        # Enforcing Exact Strict Headers Rule Set
        pdf.set_y(10)
        
        if os.path.exists(LOGO_PATH):
            pdf.image(LOGO_PATH, x=10, y=8, w=45)
            
        if declaration_date:
            curr_y = pdf.get_y()
            pdf.set_font("Arial", 'B', 9)
            pdf.set_text_color(0, 0, 0)
            decl_str = f"Date: {declaration_date.strftime('%d-%m-%Y')}"
            pdf.set_xy(0, 15)
            pdf.cell(pdf.w - 10, 5, decl_str, 0, 0, 'R')
            pdf.set_y(curr_y)

        pdf.set_xy(10, 15)
        pdf.set_fill_color(255, 255, 255)
        pdf.set_text_color(0, 0, 0)
        
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(0, 6, "SVKM'S NMIMS", 0, 1, 'C')
        
        pdf.set_font("Arial", 'B', 12)
        pdf.cell(0, 5, "Academic year - 2025-2026", 0, 1, 'C')
        
        pdf.cell(0, 5, college_name, 0, 1, 'C')
        
        pdf.set_font("Arial", 'B', 11)
        pdf.cell(0, 5, "Examination Time Table (Tentative)", 0, 1, 'C')
        
        # Branch & Semester combined into the header
        pdf.cell(0, 5, f"{header_content['main_branch_full']} - Semester {header_content['semester_roman']}", 0, 1, 'C')

        # Final Header Y-offset lock updated to 41 to perfectly map exact bounds
        pdf.set_y(41)

    render_footer()
    render_header()
    
    pdf.set_font("Arial", 'B', 9)
    print_row_custom(pdf, columns, col_widths, line_height=line_height, header=True)
    pdf.set_font("Arial", '', 8)
    
    for idx in range(len(df)):
        row = [str(df.iloc[idx][c]) if pd.notna(df.iloc[idx][c]) else "" for c in columns]
        if not any(cell.strip() for cell in row): continue
            
        wrapped_cells = []
        max_lines = 0
        for i, cell_text in enumerate(row):
            text = str(cell_text) if cell_text is not None else ""
            avail_w = col_widths[i] - 2 
            lines = wrap_text(pdf, text, avail_w)
            wrapped_cells.append(lines)
            max_lines = max(max_lines, len(lines))
            
        row_h = line_height * max_lines
        
        if pdf.get_y() + row_h > pdf.h - footer_height - 5:
            pdf.add_page()
            render_footer()
            render_header()
            pdf.set_font("Arial", 'B', 9) 
            print_row_custom(pdf, columns, col_widths, line_height=line_height, header=True)
            pdf.set_font("Arial", '', 8)
        
        print_row_custom(pdf, row, col_widths, line_height=line_height, header=False)


def convert_excel_to_pdf(excel_path, pdf_path, decl_date, clean_college_name):
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=False, margin=15)
    pdf.alias_nb_pages()
    
    t_slots = {1: "10:00 AM - 1:00 PM", 2: "2:00 PM - 5:00 PM"}
    
    try:
        xls = pd.read_excel(excel_path, sheet_name=None)
        
        for sheet, df in xls.items():
            if df.empty or sheet == "Empty": continue
            
            # Reconstruct metadata from Excel sheet names
            parts = sheet.split('_|_')
            prog = parts[0]
            sem_raw = parts[1].replace('Sem ', '').replace('_Ele', '')
            
            roman_map = {'I':1, 'II':2, 'III':3, 'IV':4, 'V':5, 'VI':6, 'VII':7, 'VIII':8, 'IX':9, 'X':10}
            sem_int = roman_map.get(sem_raw, 1)
            
            slot_id = 1 if ((sem_int + 1)//2)%2 == 1 else 2
            time_slot = t_slots[slot_id]
            
            full_prog_name = BRANCH_FULL_FORM.get(prog, prog)
            header_content = {'main_branch_full': full_prog_name, 'semester_roman': sem_raw}
            
            fixed = ['Exam Date']
            others = [c for c in df.columns if c not in fixed and c not in ['Program', 'Semester'] and 'Unnamed' not in str(c)]
            
            if '_Ele' in sheet:
                target_cols = ['Exam Date', 'OE Type', 'Subjects']
                avail_cols = [c for c in target_cols if c in df.columns]
                
                if len(avail_cols) >= 3:
                    pdf.add_page()
                    col_ws = [30, 25, pdf.w - 2 * pdf.l_margin - 55]
                    print_table_custom(pdf, df, avail_cols, col_ws, line_height=6, 
                                       header_content=header_content, college_name=clean_college_name, 
                                       time_slot=time_slot, declaration_date=decl_date)
            else:
                per_page = 4
                for i in range(0, len(others), per_page):
                    chunk = others[i:i+per_page]
                    cols = fixed + chunk
                    
                    sub_df = df[cols].dropna(how='all', subset=chunk).copy()
                    if sub_df.empty: continue
                    
                    try:
                        sub_df["Exam Date"] = pd.to_datetime(sub_df["Exam Date"], format="%d-%m-%Y", errors='coerce').dt.strftime("%A, %d %B, %Y")
                    except: pass
                    
                    w_date = 30
                    w_rem = pdf.w - 2 * pdf.l_margin - w_date
                    w_col = w_rem / max(len(chunk), 1)
                    col_ws = [w_date] + [w_col]*len(chunk)
                    
                    pdf.add_page()
                    print_table_custom(pdf, sub_df, cols, col_ws, line_height=6, 
                                       header_content=header_content, college_name=clean_college_name, 
                                       time_slot=time_slot, declaration_date=decl_date)
                            
    except Exception as e:
        st.error(f"PDF Gen Error: {e}")
        return False
        
    # INSTRUCTIONS PAGE
    try:
        pdf.add_page()
        pdf.set_y(10)
        if os.path.exists(LOGO_PATH):
            pdf.image(LOGO_PATH, x=(pdf.w-45)/2, y=8, w=45)
            
        pdf.set_fill_color(149, 33, 28)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Arial", 'B', 14)
        pdf.rect(10, 25, pdf.w - 20, 8, 'F')
        pdf.set_xy(10, 25)
        pdf.cell(pdf.w - 20, 8, clean_college_name, 0, 1, 'C')
        
        pdf.set_font("Arial", 'B', 12)
        pdf.set_text_color(0, 0, 0)
        pdf.set_xy(10, 40)
        pdf.cell(0, 10, "EXAMINATION GUIDELINES - Semester General", 0, 1, 'C')
        
        pdf.set_y(60)
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(0, 10, "INSTRUCTIONS TO CANDIDATES", 0, 1, 'C')
        pdf.ln(5)
        instrs = [
            "1. Candidates are required to be present at the examination center THIRTY MINUTES before the stipulated time.",
            "2. Candidates must produce their University Identity Card at the time of the examination.",
            "3. Candidates are not permitted to enter the examination hall after stipulated time.",
            "4. Candidates will not be permitted to leave the examination hall during the examination time."
        ]
        pdf.set_font("Arial", '', 10)
        for i in instrs:
            pdf.multi_cell(0, 7, i)
            pdf.ln(2)
            
        # Footer
        pdf.set_xy(10, pdf.h - 20)
        pdf.set_font("Arial", 'B', 9)
        pdf.cell(0, 5, "Controller of Examinations", 0, 1, 'L')
        pdf.line(10, pdf.h - 15, 60, pdf.h - 15)
        pdf.set_font("Arial", size=9)
        pdf.set_xy(pdf.w - 30, pdf.h - 15)
        pdf.cell(20, 5, f"{pdf.page_no()} of {{nb}}", 0, 0, 'R')
        
    except: pass
    
    pdf.output(pdf_path)
    return True

# ==========================================
# 🚀 MAIN APP
# ==========================================
def main():
    st.markdown('<div class="main-header"><h1>📄 Verification File to PDF Converter</h1><p>Generate Strict Standardized Timetables Directly from Verification output</p></div>', unsafe_allow_html=True)
    
    if 'pdf_data' not in st.session_state: st.session_state.pdf_data = None
    if 'raw_df' not in st.session_state: st.session_state.raw_df = None
    if 'processed_tt' not in st.session_state: st.session_state.processed_tt = None

    with st.sidebar:
        st.header("⚙️ Configuration")
        college_options = [c["name"] for c in COLLEGES]
        
        # Determine Clean Name
        selected_display = st.selectbox("Select College", college_options)
        st.session_state.clean_college_name = selected_display
        
        st.markdown("---")
        decl_date = st.date_input("📆 Declaration Date (Optional)", value=None)

    col1, col2 = st.columns([2, 1])
    with col1:
        st.markdown('<div class="upload-section"><h3 style="margin: 0 0 1rem 0; color: #951C1C;">📁 Upload Verification Excel</h3><p style="margin: 0; color: #666; font-size: 1rem;">Drag and drop the Verification file exported from the primary app</p></div>', unsafe_allow_html=True)
        uploaded = st.file_uploader("Upload Excel", type=['xlsx', 'xls'], label_visibility="collapsed")
        
        if uploaded:
            with st.spinner("Parsing format..."):
                tt, df = process_verification_file(uploaded)
                if tt:
                    st.session_state.processed_tt = tt
                    st.session_state.raw_df = df
                    st.success("✅ Verification File Loaded Successfully!")
    
    with col2:
        st.info("ℹ️ **Direct Converter Mode:** This app processes scheduled subjects precisely matching the FPDF boundaries and constraints of the original builder. The formatting adheres strictly to A4 Landscape constraints.")

    if st.session_state.raw_df is not None:
        df = st.session_state.raw_df
        
        # Verify Flags
        flags = df[df.get('Capacity Exceeded Limit', '') == 'YES (MUMBAI LIMIT HIT)']
        if not flags.empty:
            st.warning(f"⚠️ **Note:** Your verification file contains {len(flags)} subjects that exceeded capacity limits.")

        st.markdown("---")
        
        if st.button("🚀 Generate PDF Timetable", type="primary", use_container_width=True):
            with st.spinner("Rendering strict boundaries and converting to PDF..."):
                excel_io = save_to_excel(st.session_state.processed_tt)
                with open("temp_v.xlsx", "wb") as f: f.write(excel_io.getvalue())
                
                if convert_excel_to_pdf("temp_v.xlsx", "Verification_Timetable.pdf", decl_date, st.session_state.clean_college_name):
                    st.balloons()
                    with open("Verification_Timetable.pdf", "rb") as f:
                        st.download_button("📥 Download Finalized PDF", f, f"Verification_Timetable_{datetime.now().strftime('%Y%m%d')}.pdf", "application/pdf", use_container_width=True)
                    
                    try: os.remove("temp_v.xlsx"); os.remove("Verification_Timetable.pdf")
                    except: pass

if __name__ == "__main__":
    main()
