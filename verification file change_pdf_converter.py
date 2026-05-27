import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from fpdf import FPDF
import os
import re
import io
import traceback
import uuid
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
        padding: 2.5rem; border-radius: 16px; margin-bottom: 2rem;
        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.12);
    }
    .main-header h1 { color: white; text-align: center; margin: 0; font-size: 2.5rem; font-weight: 700; text-shadow: 2px 2px 4px rgba(0,0,0,0.3); letter-spacing: -0.5px; }
    .main-header p { color: #FFF; text-align: center; margin: 0.5rem 0 0 0; font-size: 1.1rem; opacity: 0.95; font-weight: 500; }
    .upload-section { background: #f8f9fa; padding: 2.5rem; border-radius: 16px; border: 2px dashed #951C1C; margin: 1rem 0; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08); text-align: center; }
    @media (prefers-color-scheme: dark) {
        .main-header { background: linear-gradient(135deg, #701515 0%, #A23217 100%); }
        .upload-section { background: #2d2d2d; border-color: #A23217; }
    }
    .stButton>button { transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1); border-radius: 12px; font-weight: 600; }
    .stButton>button:hover { transform: translateY(-2px); box-shadow: 0 8px 16px rgba(149, 28, 28, 0.3); border-color: #951C1C; }
</style>
""", unsafe_allow_html=True)


# ==========================================
# 📥 VERIFICATION DATA PARSING
# ==========================================
def process_verification_file(uploaded_file):
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_name = next((s for s in xls.sheet_names if "verification" in s.lower()), xls.sheet_names[0])
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        df.columns = df.columns.str.strip()

        has_date2 = "Exam Date 2" in df.columns
        has_date  = "Exam Date"  in df.columns

        if not has_date2 and not has_date:
            st.error("❌ Missing critical column: need 'Exam Date 2' or 'Exam Date'")
            return None, None

        if "Program" not in df.columns:
            st.error("❌ Missing critical column: Program")
            return None, None

        if "Current Session" not in df.columns:
            st.error("❌ Missing critical column: Current Session")
            return None, None

        # ── SCHEDULING FILTER ────────────────────────────────────────────────
        if has_date2:
            if 'Exam Date' in df.columns:
                df = df.drop(columns=['Exam Date'])
            if 'Exam Time' in df.columns:
                df = df.drop(columns=['Exam Time'])

            df = df[df['Exam Date 2'].notna()].copy()
            df['_date_str'] = df['Exam Date 2'].astype(str).str.strip().str.upper()
            df = df[~df['_date_str'].isin(['NOT SCHEDULED', 'NAN', 'NAT', 'NONE', ''])].drop(columns=['_date_str'])

            parsed_dates = pd.to_datetime(df['Exam Date 2'], errors='coerce')
            df = df[parsed_dates.notna()].copy()
            parsed_dates = parsed_dates[parsed_dates.notna()]

            if 'Exam Time 2' in df.columns:
                df['_time_str'] = df['Exam Time 2'].astype(str).str.strip().str.upper()
                df = df[~df['_time_str'].isin(['NAN', 'NAT', 'NONE', ''])].drop(columns=['_time_str'])

            df['Exam Date'] = parsed_dates[parsed_dates.index.isin(df.index)].dt.strftime('%d-%m-%Y')
            df['Exam Time'] = df['Exam Time 2'].astype(str).str.strip() if 'Exam Time 2' in df.columns else ""
        else:
            df = df[df['Exam Date'].notna()].copy()
            df['_date_str'] = df['Exam Date'].astype(str).str.strip().str.upper()
            df = df[~df['_date_str'].isin(['NOT SCHEDULED', 'NAN', 'NAT', 'NONE', ''])].drop(columns=['_date_str'])
            parsed_dates = pd.to_datetime(df['Exam Date'], errors='coerce')
            df = df[parsed_dates.notna()].copy()
            parsed_dates = parsed_dates[parsed_dates.notna()]
            df['Exam Date'] = parsed_dates.dt.strftime('%d-%m-%Y')
            exam_time_col = 'Exam Time' if 'Exam Time' in df.columns else ('Exam Time 2' if 'Exam Time 2' in df.columns else None)
            df['Exam Time'] = df[exam_time_col].astype(str).str.strip() if exam_time_col else ""

        df['MainBranch'] = df.get('Program', pd.Series(dtype=str)).fillna("").astype(str).str.strip()
        df['SubBranch']  = df.get('Stream',  pd.Series(dtype=str)).fillna("").astype(str).str.strip()
        df['SubBranch']  = df.apply(lambda x: x['MainBranch'] if x['SubBranch'] in ['nan', ''] else x['SubBranch'], axis=1)

        df['Subject']    = df.get('Module Description',  pd.Series(dtype=str)).fillna("").astype(str).str.strip()
        df['ModuleCode'] = df.get('Module Abbreviation', pd.Series(dtype=str)).fillna("").astype(str).str.strip()

        if 'Exam Slot Number' in df.columns:
            df['ExamSlotNumber'] = pd.to_numeric(df['Exam Slot Number'], errors='coerce').fillna(1).astype(int)
        elif 'Exam_Slot_Number' in df.columns:
            df['ExamSlotNumber'] = pd.to_numeric(df['Exam_Slot_Number'], errors='coerce').fillna(1).astype(int)
        elif 'Slot Number' in df.columns:
            df['ExamSlotNumber'] = pd.to_numeric(df['Slot Number'], errors='coerce').fillna(1).astype(int)
        else:
            df['ExamSlotNumber'] = 1

        if 'Subject Type' in df.columns:
            df['OE'] = df['Subject Type'].apply(lambda x: 'OE' if str(x).strip().upper() == 'OE' else None)
        else:
            df['OE'] = None

        def get_sem_int(val):
            s = str(val).upper().strip()
            m = re.search(r'(\d+)', s)
            if m: return int(m.group(1))
            roman_map = {
                'XII': 12, 'XI': 11, 'X': 10, 'IX': 9, 'VIII': 8,
                'VII': 7, 'VI': 6, 'V': 5, 'IV': 4, 'III': 3, 'II': 2, 'I': 1
            }
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
        st.error(traceback.format_exc())
        return None, None


# ==========================================
# 💾 EXCEL ENGINE
# ==========================================
def normalize_time(t_str):
    if not isinstance(t_str, str): return ""
    t_str = t_str.strip().upper()
    for i in range(1, 10): t_str = t_str.replace(f"0{i}:", f"{i}:")
    return t_str

def int_to_roman(num):
    val = [(1000,"M"),(900,"CM"),(500,"D"),(400,"CD"),(100,"C"),(90,"XC"),
           (50,"L"),(40,"XL"),(10,"X"),(9,"IX"),(5,"V"),(4,"IV"),(1,"I")]
    res = ""
    for v, r in val:
        while num >= v: res += r; num -= v
    return res

def _make_program_abbrev(name):
    norm = re.sub(r'\s+', ' ', name.strip())
    words = re.split(r'[\s,&/()+]+', norm)
    abbrev = ''.join(w[0] for w in words if w)
    return abbrev[:12].upper()

def _build_program_key_map(all_program_names):
    key_map     = {}      
    norm_to_key = {}      
    used_keys   = set()

    for name in all_program_names:
        norm = re.sub(r'\s+', ' ', name.strip())
        if norm in norm_to_key:
            key_map[name] = norm_to_key[norm]
            continue

        base = _make_program_abbrev(name)
        key  = base
        n    = 2
        while key in used_keys:
            key = base[:10] + str(n)
            n  += 1

        used_keys.add(key)
        norm_to_key[norm] = key
        key_map[name]     = key

    return key_map


def save_to_excel(semester_wise_timetable):
    time_slots_dict = st.session_state.get('time_slots', {
        1: {"start": "10:00 AM", "end": "1:00 PM"},
        2: {"start": "2:00 PM",  "end": "5:00 PM"}
    })

    current_college_context = st.session_state.get('selected_college', '')
    IS_BUSINESS_SCH = ("School of Business Management" in current_college_context or "Pravin Dalal" in current_college_context)

    if IS_BUSINESS_SCH:
        time_slots_dict = {
            1: {"start": "08:30 AM", "end": "10:30 AM"},
            2: {"start": "11:30 AM", "end": "01:30 PM"},
            3: {"start": "03:00 PM", "end": "05:00 PM"}
        }

    output = io.BytesIO()

    all_programs = []
    for df_s in semester_wise_timetable.values():
        if not df_s.empty:
            all_programs.extend(df_s['MainBranch'].dropna().unique().tolist())
            
    seen_p, deduped = set(), []
    for p in all_programs:
        if p not in seen_p:
            deduped.append(p); seen_p.add(p)
    prog_key_map = _build_program_key_map(deduped)

    used_sheet_names = set()
    def unique_sheet(name):
        base, n = name, 2
        while name in used_sheet_names:
            name = base[:31 - len(str(n))] + str(n)
            n   += 1
        used_sheet_names.add(name)
        return name

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        sheets_created = 0
        for sem, df_sem in semester_wise_timetable.items():
            if df_sem.empty: continue

            slot_id     = 1 if ((sem + 1) // 2) % 2 == 1 else 2
            p_cfg       = time_slots_dict.get(slot_id, time_slots_dict[1])
            header_norm = normalize_time(f"{p_cfg['start']} - {p_cfg['end']}")

            for main_branch in df_sem['MainBranch'].unique():
                df_mb     = df_sem[df_sem['MainBranch'] == main_branch].copy()
                roman_sem = int_to_roman(sem)

                prog_key   = prog_key_map.get(main_branch, main_branch[:12])
                core_sheet = unique_sheet(f"{prog_key}_|_Sem {roman_sem}"[:31])
                elec_sheet = unique_sheet(f"{prog_key}_|_Sem {roman_sem}_Ele"[:31])

                df_core = df_mb[df_mb['OE'].isna()].copy()
                df_elec = df_mb[df_mb['OE'].notna()].copy()

                if not df_core.empty:
                    displays = []
                    for _, row in df_core.iterrows():
                        subj        = row['Subject']
                        code        = row['ModuleCode']
                        actual_time = str(row.get('Exam Time', '')).strip()

                        time_suffix = ""
                        if actual_time and normalize_time(actual_time) != header_norm and actual_time.lower() not in ['tbd', 'nan', '']:
                            time_suffix = f" [{actual_time}]"

                        txt = f"{subj}"
                        if code and str(code).lower() != 'nan': txt += f" - ({code})"
                        txt += time_suffix
                        displays.append(txt)

                    df_core['SubjectDisplay'] = displays
                    df_core["Exam Date"] = pd.to_datetime(df_core["Exam Date"], format="%d-%m-%Y", dayfirst=True, errors='coerce')
                    df_core = df_core.sort_values(by="Exam Date", ascending=True)

                    try:
                        pivot = df_core.groupby(['Exam Date', 'SubBranch']).agg({'SubjectDisplay': lambda x: " <hr> ".join(str(i) for i in x)}).reset_index()
                        pivot = pivot.pivot_table(index='Exam Date', columns='SubBranch', values='SubjectDisplay', aggfunc='first').fillna("---")
                        pivot = pivot.sort_index(ascending=True).reset_index()
                        pivot['Exam Date'] = pivot['Exam Date'].apply(lambda x: x.strftime("%d-%m-%Y") if pd.notna(x) else "")
                        
                        pivot['_prog_'] = main_branch
                        pivot['_sem_']  = roman_sem
                        pivot.to_excel(writer, sheet_name=core_sheet, index=False)
                        sheets_created += 1
                    except Exception:
                        pass

                if not df_elec.empty:
                    e_displays = []
                    for _, row in df_elec.iterrows():
                        subj        = row['Subject']
                        actual_time = str(row.get('Exam Time', '')).strip()

                        time_suffix = ""
                        if actual_time and normalize_time(actual_time) != header_norm and actual_time.lower() not in ['tbd', 'nan', '']:
                            time_suffix = f" [{actual_time}]"

                        txt = f"{subj}"
                        txt += time_suffix
                        e_displays.append(txt)

                    df_elec['DisplaySubject'] = e_displays

                    try:
                        df_elec["Exam Date"] = pd.to_datetime(df_elec["Exam Date"], format="%d-%m-%Y", dayfirst=True, errors='coerce')
                        df_elec = df_elec.sort_values(by="Exam Date", ascending=True)
                        df_elec['Exam Date'] = df_elec['Exam Date'].apply(lambda x: x.strftime("%d-%m-%Y") if pd.notna(x) else "")

                        ep = df_elec.groupby(['Exam Date', 'OE']).agg({'DisplaySubject': lambda x: ", ".join(sorted(set(x)))}).reset_index()
                        ep.rename(columns={'OE': 'OE Type', 'DisplaySubject': 'Open Elective (All Applicable Streams)'}, inplace=True)
                        ep['_prog_'] = main_branch
                        ep['_sem_']  = roman_sem
                        ep.to_excel(writer, sheet_name=elec_sheet, index=False)
                        sheets_created += 1
                    except Exception:
                        pass

        if sheets_created == 0:
            pd.DataFrame({'Info': ['No valid data']}).to_excel(writer, sheet_name="Empty")

    output.seek(0)
    return output


# ==========================================
# 📄 FPDF ENGINE — VERBATIM FROM REFERENCE APP
# ==========================================

def wrap_text(pdf, text, col_width):
    cache_key = (text, col_width, pdf.font_style)
    if cache_key in wrap_text_cache:
        return wrap_text_cache[cache_key]
        
    time_pattern = re.compile(r'([\[\(]\s*\d{1,2}:\d{2}\s*[AP]M\s*-\s*\d{1,2}:\d{2}\s*[AP]M\s*[\]\)])', re.IGNORECASE)
    
    parts = time_pattern.split(str(text))
    tokens = []
    for i, p in enumerate(parts):
        if i % 2 == 1:
            tokens.append(p)
        else:
            p = p.replace("<hr>", " <hr> ")
            tokens.extend(p.split())
            
    lines = []
    current_line = ""
    
    old_family = pdf.font_family
    old_style = pdf.font_style
    old_size = pdf.font_size_pt
    
    for token in tokens:
        if token == "<hr>":
            if current_line:
                lines.append(current_line)
                current_line = ""
            lines.append("<hr>")
            continue

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
            if current_line:
                lines.append(current_line)
            current_line = token
            
    if current_line:
        lines.append(current_line)
        
    pdf.set_font(old_family, old_style, old_size)
    
    wrap_text_cache[cache_key] = lines
    return lines

def print_row_custom(pdf, row_data, col_widths, line_height=4.5, header=False):
    cell_padding = 1
    header_bg_color = (255, 255, 255)
    header_text_color = (0, 0, 0)
    alt_row_color = (255, 255, 255)

    row_number = getattr(pdf, '_row_counter', 0)
    
    base_font = "Times"
    if header:
        base_style = 'B'
        base_size = 9.5
        pdf.set_font(base_font, base_style, base_size)
        pdf.set_text_color(*header_text_color)
        pdf.set_fill_color(*header_bg_color)
    else:
        base_style = ''
        base_size = 8.5
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
            
            if ln == "<hr>":
                line_y = y0 + j * line_height + pad_v + (line_height / 2)
                pdf.line(cx, line_y, cx + col_widths[i], line_y)
                continue
                
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
                    pdf.set_xy(current_x - pdf.c_margin, y0 + j * line_height + pad_v)
                    pdf.cell(w + 2 * pdf.c_margin, line_height, p, border=0, align='L')
                    
                    current_x += w
                
                pdf.set_font(base_font, base_style, base_size)
        
        pdf.rect(cx, y0, col_widths[i], row_h)
        pdf.set_xy(cx + col_widths[i], y0)

    setattr(pdf, '_row_counter', row_number + 1)
    pdf.set_xy(x0, y0 + row_h)

def print_table_custom(pdf, df, columns, col_widths, line_height=4.5, header_content=None, Programs=None, time_slot=None, actual_time_slots=None, declaration_date=None):
    if df.empty: return
    setattr(pdf, '_row_counter', 0)
    
    footer_height = 14   
    header_end_y = 56    
    
    def render_footer():
        pdf.set_xy(10, pdf.h - footer_height)
        pdf.set_font("Times", 'B', 8) 
        pdf.cell(0, 5, "CONTROLLER OF EXAMINATIONS", 0, 1, 'L')
        pdf.line(10, pdf.h - footer_height + 5, 60, pdf.h - footer_height + 5)
        
        pdf.set_font("Times", size=8)
        pdf.set_text_color(0, 0, 0)
        page_text = f"{pdf.page_no()} of {{nb}}"
        text_width = pdf.get_string_width(page_text.replace("{nb}", "99"))
        pdf.set_xy(pdf.w - 10 - text_width, pdf.h - footer_height + 5)
        pdf.cell(text_width, 5, page_text, 0, 0, 'R')

    def render_header():
        pdf.set_y(0)
        if declaration_date:
            day = declaration_date.day
            if 11 <= (day % 100) <= 13:
                suffix = 'TH'
            else:
                suffix = {1: 'ST', 2: 'ND', 3: 'RD'}.get(day % 10, 'TH')
            
            decl_str = f"DATE: {day}{suffix} {declaration_date.strftime('%B, %Y')}".upper()
            
            pdf.set_font("Times", 'B', 12) 
            pdf.set_text_color(0, 0, 0)
            pdf.set_xy(pdf.w - 80, 8)
            pdf.cell(70, 10, decl_str, 0, 0, 'R')

        # Logo
        logo_width = 45
        logo_x = (pdf.w - logo_width) / 2
        if os.path.exists(LOGO_PATH):
            pdf.image(LOGO_PATH, x=logo_x, y=5, w=logo_width)
        
        # College Name
        pdf.set_text_color(0, 0, 0)
        college_name = st.session_state.get('selected_college', "SVKM's NMIMS University").upper()
        pdf.set_font("Times", 'B', 12) 
        
        pdf.set_xy(10, 25)
        pdf.cell(pdf.w - 20, 6, college_name, 0, 1, 'C')
        
        # Final Exam Timetable Header
        pdf.set_font("Times", 'B', 10) 
        pdf.set_text_color(0, 0, 0)
        pdf.set_xy(10, 33)
        pdf.cell(pdf.w - 20, 4, "FINAL EXAMINATION TIMETABLE (ACADEMIC YEAR: 2025-26)", 0, 1, 'C')
        
        current_y = 38
        
        # Program Name
        pdf.set_font("Times", 'B', 10) 
        pdf.set_xy(10, current_y)
        pdf.cell(pdf.w - 20, 4, f"{header_content['main_branch_full']}".upper(), 0, 1, 'C')
        current_y += 4
        
        # Year and Semester Math Calculation
        sem_roman = str(header_content['semester_roman']).upper()
        roman_map = {'XII': 12, 'XI': 11, 'X': 10, 'IX': 9, 'VIII': 8, 'VII': 7, 'VI': 6, 'V': 5, 'IV': 4, 'III': 3, 'II': 2, 'I': 1}
        sem_int = roman_map.get(sem_roman)
        if not sem_int:
            m = re.search(r'(\d+)', sem_roman)
            if m: sem_int = int(m.group(1))
            else: sem_int = 1
        year_int = (sem_int + 1) // 2
        year_roman = int_to_roman(year_int)

        pdf.set_xy(10, current_y)
        pdf.cell(pdf.w - 20, 4, f"YEAR: {year_roman}, SEMESTER: {sem_roman}".upper(), 0, 1, 'C')
        current_y += 4

        if time_slot:
            pdf.set_font("Times", 'B', 9)
            pdf.set_xy(10, current_y)
            pdf.cell(pdf.w - 20, 4, f"EXAM TIME: {time_slot}".upper(), 0, 1, 'C')
            current_y += 4
            
            pdf.set_font("Times", 'BI', 9) 
            pdf.set_xy(10, current_y)
            pdf.cell(pdf.w - 20, 4, "(CHECK THE SUBJECT EXAM TIME)".upper(), 0, 1, 'C')
            current_y += 4

        pdf.set_xy(pdf.l_margin, header_end_y)

    render_footer()
    render_header()
    
    # Uppercase the table columns
    upper_columns = [str(c).upper() for c in columns]
    
    pdf.set_font("Times", 'B', 9.5)
    print_row_custom(pdf, upper_columns, col_widths, line_height=line_height, header=True)
    
    pdf.set_font("Times", '', 8.5) 
    
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
            pdf.set_font("Times", 'B', 9.5) 
            print_row_custom(pdf, upper_columns, col_widths, line_height=line_height, header=True)
            pdf.set_font("Times", '', 8.5)  
        
        print_row_custom(pdf, row, col_widths, line_height=line_height, header=False)

def convert_excel_to_pdf(excel_path, pdf_path=None, sub_branch_cols_per_page=6, declaration_date=None):
    current_college_context = st.session_state.get('selected_college', '')
    IS_BUSINESS_SCH = ("School of Business Management" in current_college_context or "Pravin Dalal" in current_college_context)

    time_slots_dict = st.session_state.get('time_slots', {
        1: {"start": "10:00 AM", "end": "1:00 PM"},
        2: {"start": "2:00 PM", "end": "5:00 PM"}
    })

    if IS_BUSINESS_SCH:
        time_slots_dict = {
            1: {"start": "08:30 AM", "end": "10:30 AM"},
            2: {"start": "11:30 AM", "end": "01:30 PM"},
            3: {"start": "03:00 PM", "end": "05:00 PM"}
        }
    
    try:
        df_dict = pd.read_excel(excel_path, sheet_name=None)
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return {}

    def get_header_time_for_semester(sem_str):
        try:
            s = str(sem_str).strip().upper()
            sem_int = 1
            romans = {'I': 1, 'II': 2, 'III': 3, 'IV': 4, 'V': 5, 'VI': 6, 'VII': 7, 'VIII': 8, 'IX': 9, 'X': 10, 'XI': 11, 'XII': 12}
            found = False
            for r_key, r_val in romans.items():
                if s == r_key or s.endswith(f" {r_key}") or s.endswith(f"_{r_key}"):
                    sem_int = r_val
                    found = True
                    break
            if not found:
                digits = re.findall(r'\d+', s)
                if digits: sem_int = int(digits[0])
            slot_indicator = ((sem_int + 1) // 2) % 2
            slot_num = 1 if slot_indicator == 1 else 2
            slot_cfg = time_slots_dict.get(slot_num, time_slots_dict.get(1))
            return f"{slot_cfg['start']} - {slot_cfg['end']}"
        except:
            return f"{time_slots_dict[1]['start']} - {time_slots_dict[1]['end']}"

    sheets_processed = 0
    pdf_outputs = {}

    if IS_BUSINESS_SCH:
        footer_height    = 14
        header_end_y     = 68   

        INSTRS = [
            ("INSTRUCTIONS TO CANDIDATES", True),   
            ("1. Candidates are required to be present at the examination centre "
             "THIRTY MINUTES before the stipulated time.", False),
            ("2. Candidates must produce their University Identity Card at the "
             "time of the examination.", False),
            ("3. Candidates are not permitted to enter the examination hall after "
             "stipulated time.", False),
            ("4. Candidates will not be permitted to leave the examination hall "
             "during the examination time.", False),
            ("5. Candidates are forbidden from taking any unauthorized material "
             "inside the examination hall. Carrying the same will be treated as "
             "usage of unfair means.", False),
        ]

        def _measure_instructions(pdf_obj, font_size, line_h, usable_w):
            total = 0
            for text, is_heading in INSTRS:
                pdf_obj.set_font("Times", 'BU' if is_heading else '', font_size)
                words = text.split()
                lines, cur = [], ""
                for w in words:
                    test = (cur + " " + w).strip()
                    if pdf_obj.get_string_width(test) <= usable_w:
                        cur = test
                    else:
                        if cur: lines.append(cur)
                        cur = w
                if cur: lines.append(cur)
                total += len(lines) * line_h + (1 if is_heading else 0.5)
            return total

        def render_instructions_sbm(pdf_obj, y_start):
            usable_w  = pdf_obj.w - 2 * pdf_obj.l_margin
            available = pdf_obj.h - footer_height - 2 - y_start
            chosen_size = 6.5
            chosen_lh   = 4.0
            for fs in [8.5, 8.0, 7.5, 7.0, 6.5]:
                lh = fs * 0.55
                if _measure_instructions(pdf_obj, fs, lh, usable_w) <= available:
                    chosen_size = fs
                    chosen_lh   = lh
                    break

            pdf_obj.set_text_color(0, 0, 0)
            cy = y_start + 1
            for text, is_heading in INSTRS:
                if is_heading: pdf_obj.set_font("Times", 'BU', chosen_size)
                else: pdf_obj.set_font("Times", '', chosen_size)

                words = text.split()
                lines, cur = [], ""
                for w in words:
                    test = (cur + " " + w).strip()
                    if pdf_obj.get_string_width(test) <= usable_w: cur = test
                    else:
                        if cur: lines.append(cur)
                        cur = w
                if cur: lines.append(cur)

                for ln in lines:
                    pdf_obj.set_xy(pdf_obj.l_margin, cy)
                    pdf_obj.cell(usable_w, chosen_lh, ln, 0, 0, 'L')
                    cy += chosen_lh
                cy += 0.5 if not is_heading else 1.0

        num_slots   = len(time_slots_dict)
        slot_labels = {}                           
        for sn, scfg in sorted(time_slots_dict.items()):
            slot_labels[sn] = f"{scfg['start']} to {scfg['end']}"

        def render_footer_sbm(pdf_obj):
            pdf_obj.set_xy(10, pdf_obj.h - footer_height)
            pdf_obj.set_font("Times", 'B', 8)
            pdf_obj.cell(0, 5, "CONTROLLER OF EXAMINATIONS", 0, 1, 'L')
            pdf_obj.line(10, pdf_obj.h - footer_height + 5, 70, pdf_obj.h - footer_height + 5)
            pdf_obj.set_font("Times", '', 8)
            page_text = f"{pdf_obj.page_no()} of {{nb}}"
            tw = pdf_obj.get_string_width(page_text.replace("{nb}", "99"))
            pdf_obj.set_xy(pdf_obj.w - 10 - tw, pdf_obj.h - footer_height + 5)
            pdf_obj.cell(tw, 5, page_text, 0, 0, 'R')

        def render_header_sbm(pdf_obj, header_content, declaration_date):
            pdf_obj.set_y(0)
            if declaration_date:
                day = declaration_date.day
                suffix = ('TH' if 11 <= (day % 100) <= 13 else {1:'ST',2:'ND',3:'RD'}.get(day % 10, 'TH'))
                decl_str = f"{day}{suffix} {declaration_date.strftime('%B %Y')}"
                pdf_obj.set_font("Times", 'B', 11)
                pdf_obj.set_text_color(0, 0, 0)
                pdf_obj.set_xy(pdf_obj.w - 80, 8)
                pdf_obj.cell(70, 8, decl_str, 0, 0, 'R')

            logo_w    = 45
            logo_y    = 3
            sbm_logo  = "logo_sbm.png"
            logo_file = sbm_logo if os.path.exists(sbm_logo) else LOGO_PATH
            if os.path.exists(logo_file):
                pdf_obj.image(logo_file, x=(pdf_obj.w - logo_w) / 2, y=logo_y, w=logo_w)

            text_y  = logo_y + 42
            F_COLLEGE, F_TITLE, F_PROG, F_YEAR, LINE_GAP = 10, 8.5, 8.5, 8.5, 1

            college_name = st.session_state.get('selected_college', "SVKM's NMIMS University").upper()
            pdf_obj.set_text_color(0, 0, 0)
            pdf_obj.set_font("Times", 'B', F_COLLEGE)
            cell_h = F_COLLEGE * 0.40
            pdf_obj.set_xy(10, text_y)
            pdf_obj.cell(pdf_obj.w - 20, cell_h, college_name, 0, 1, 'C')
            text_y += cell_h + LINE_GAP

            pdf_obj.set_font("Times", 'B', F_TITLE)
            cell_h = F_TITLE * 0.40
            pdf_obj.set_xy(10, text_y)
            pdf_obj.cell(pdf_obj.w - 20, cell_h, "FINAL EXAMINATION TIMETABLE (ACADEMIC YEAR: 2025-26)", 0, 1, 'C')
            text_y += cell_h + LINE_GAP

            prog_name = str(header_content.get('main_branch_full', '')).upper()
            pdf_obj.set_font("Times", 'B', F_PROG)
            cell_h = F_PROG * 0.40
            pdf_obj.set_xy(10, text_y)
            pdf_obj.cell(pdf_obj.w - 20, cell_h, prog_name, 0, 1, 'C')
            text_y += cell_h + LINE_GAP

            sem_roman = str(header_content.get('semester_roman', '')).upper()
            roman_map = {'XII':12,'XI':11,'X':10,'IX':9,'VIII':8,'VII':7,'VI':6,'V':5,'IV':4,'III':3,'II':2,'I':1}
            sem_int = roman_map.get(sem_roman, None)
            if sem_int is None:
                m = re.search(r'(\d+)', sem_roman)
                sem_int = int(m.group(1)) if m else 1
            year_int = (sem_int + 1) // 2

            def _to_roman(n):
                val = [(1000,"M"),(900,"CM"),(500,"D"),(400,"CD"),(100,"C"),(90,"XC"),(50,"L"),(40,"XL"),(10,"X"),(9,"IX"),(5,"V"),(4,"IV"),(1,"I")]
                r = ""
                for v, s in val:
                    while n >= v: r += s; n -= v
                return r

            pdf_obj.set_font("Times", 'B', F_YEAR)
            cell_h = F_YEAR * 0.40
            pdf_obj.set_xy(10, text_y)
            pdf_obj.cell(pdf_obj.w - 20, cell_h, f"YEAR: {_to_roman(year_int)}, TRIMESTER: {sem_roman}", 0, 1, 'C')
            text_y += cell_h + LINE_GAP
            pdf_obj.set_xy(pdf_obj.l_margin, text_y + 3)

        def _wrap_cell(pdf_obj, text, avail_w, font_style='', font_size=9):
            pdf_obj.set_font("Times", font_style, font_size)
            words, lines, cur = str(text).split(), [], ""
            for w in words:
                test = (cur + " " + w).strip()
                if pdf_obj.get_string_width(test) <= avail_w: cur = test
                else:
                    if cur: lines.append(cur)
                    cur = w
            if cur: lines.append(cur)
            return lines if lines else [""]

        LINE_H, PAD = 5, 1.5

        def _draw_row(pdf_obj, cells, col_widths, font_style='', font_size=9, fill_color=None, text_color=(0,0,0)):
            if fill_color: pdf_obj.set_fill_color(*fill_color)
            pdf_obj.set_text_color(*text_color)
            pdf_obj.set_font("Times", font_style, font_size)

            wrapped = []
            for i, txt in enumerate(cells):
                avail = col_widths[i] - 2 * PAD
                wrapped.append(_wrap_cell(pdf_obj, txt, avail, font_style, font_size))

            max_lines = max(len(lns) for lns in wrapped)
            row_h     = LINE_H * max_lines
            x0, y0    = pdf_obj.get_x(), pdf_obj.get_y()

            if fill_color: pdf_obj.rect(x0, y0, sum(col_widths), row_h, 'F')

            cx = x0
            for i, lines in enumerate(wrapped):
                total_text_h = len(lines) * LINE_H * 0.85
                pad_v = (row_h - total_text_h) / 2
                for j, ln in enumerate(lines):
                    pdf_obj.set_xy(cx + PAD, y0 + pad_v + j * LINE_H * 0.85)
                    pdf_obj.cell(col_widths[i] - 2 * PAD, LINE_H * 0.85, ln, border=0, align='C')
                pdf_obj.rect(cx, y0, col_widths[i], row_h)
                cx += col_widths[i]

            pdf_obj.set_xy(x0, y0 + row_h)
            return row_h

        # Iterate sheets and build individual PDFs
        for sheet_name, sheet_df in df_dict.items():
            try:
                if sheet_df.empty or sheet_name in ["No_Data","Daily_Statistics","Summary","Verification","Empty"]: continue
                if re.search(r'_Ele(\d*)$', sheet_name): continue

                main_branch_full = ""
                for col in ("Program", "MainBranch"):
                    if col in sheet_df.columns and not sheet_df[col].dropna().empty:
                        main_branch_full = str(sheet_df[col].dropna().iloc[0])
                        break
                if not main_branch_full or "Unnamed" in main_branch_full:
                    if '_|_' in sheet_name: main_branch_full = sheet_name.split('_|_')[0]
                    else: main_branch_full = sheet_name

                semester_raw = sheet_name.split('_|_')[1] if '_|_' in sheet_name else sheet_name
                display_sem = semester_raw.strip()
                for prefix in ("trimester", "semester", "sem", "tri"):
                    if display_sem.lower().startswith(prefix):
                        display_sem = display_sem[len(prefix):].strip(); break

                header_content = {'main_branch_full': main_branch_full, 'semester_roman': display_sem}
                
                # Fetch raw original data to process SBM time slots
                timetable_data = st.session_state.get('processed_tt', {})

                rows_for_sheet = []
                disp_check_upper = display_sem.strip().upper()
                roman_map2 = {'XII':12,'XI':11,'X':10,'IX':9,'VIII':8,'VII':7,'VI':6,'V':5,'IV':4,'III':3,'II':2,'I':1}
                dv = roman_map2.get(disp_check_upper)
                if dv is None:
                    m = re.search(r'\d+', disp_check_upper)
                    dv = int(m.group()) if m else 1

                for sem_key, sem_df in timetable_data.items():
                    if sem_df.empty: continue
                    if sem_key != dv: continue
                    
                    matched = sem_df[sem_df['MainBranch'].astype(str) == main_branch_full].copy()
                    if not matched.empty: rows_for_sheet.append(matched)

                slot_pivot = {}
                if rows_for_sheet:
                    combined = pd.concat(rows_for_sheet, ignore_index=True)
                    combined['Exam Date'] = pd.to_datetime(combined['Exam Date'], format="%d-%m-%Y", errors='coerce')
                    combined = combined.dropna(subset=['Exam Date']).sort_values('Exam Date')
                    combined['ExamSlotNumber'] = pd.to_numeric(combined.get('ExamSlotNumber', 1), errors='coerce').fillna(1).astype(int)
                    combined['ExamSlotNumber'] = combined['ExamSlotNumber'].apply(lambda x: x if x in time_slots_dict else 1)

                    for _, row in combined.iterrows():
                        d_str = row['Exam Date'].strftime("%A, %d %B, %Y")
                        sn, subj = int(row.get('ExamSlotNumber', 1)), str(row.get('Subject', '')).strip()
                        if not subj or subj in ('nan', ''): continue
                        oe = str(row.get('OE', '')).strip()
                        if oe and oe not in ('nan', ''): subj = f"{subj} [{oe}]"
                        if d_str not in slot_pivot: slot_pivot[d_str] = {sn2: [] for sn2 in time_slots_dict}
                        slot_pivot[d_str].setdefault(sn, []).append(subj)

                if not slot_pivot: continue

                # INITIALIZE NEW PDF FOR THIS SPECIFIC PROGRAM/SEMESTER
                pdf = FPDF(orientation='P', unit='mm', format='A4')
                pdf.set_auto_page_break(auto=False, margin=15)
                pdf.alias_nb_pages()

                page_w    = pdf.w - 2 * pdf.l_margin   
                date_w    = 38                          
                all_slots_sorted = sorted(time_slots_dict.keys())
                active_slots = [sn for sn in all_slots_sorted if any(slot_pivot[d].get(sn) for d in slot_pivot)]
                if not active_slots: active_slots = all_slots_sorted   

                n_active   = len(active_slots)
                act_slot_w = (page_w - date_w) / n_active
                act_col_widths = [date_w] + [act_slot_w] * n_active
                act_header_cells = ["DAY & DATE"] + ["TIMING & SUBJECT"] * n_active
                act_time_cells   = [""] + [slot_labels[sn] for sn in active_slots]

                pdf.add_page()
                render_footer_sbm(pdf)
                render_header_sbm(pdf, header_content, declaration_date)
                _draw_row(pdf, act_header_cells, act_col_widths, font_style='B', font_size=9.5)
                _draw_row(pdf, act_time_cells, act_col_widths, font_style='B', font_size=9)

                _instr_h = _measure_instructions(pdf, 7.5, 7.5 * 0.55, pdf.w - 2 * pdf.l_margin)
                _table_bottom = pdf.h - footer_height - _instr_h - 4

                pdf.set_font("Times", '', 9.5)
                for d_str, slots in slot_pivot.items():
                    row_cells = [d_str]
                    for sn in active_slots:
                        subj_list = slots.get(sn, [])
                        row_cells.append("\n".join(subj_list) if subj_list else "-------------------")

                    max_lines_est = max(len(_wrap_cell(pdf, txt, act_col_widths[i] - 2 * PAD, '', 9.5)) for i, txt in enumerate(row_cells))
                    if pdf.get_y() + (LINE_H * max_lines_est) > _table_bottom:
                        render_instructions_sbm(pdf, pdf.get_y() + 2)
                        pdf.add_page()
                        render_footer_sbm(pdf)
                        render_header_sbm(pdf, header_content, declaration_date)
                        _draw_row(pdf, act_header_cells, act_col_widths, 'B', 9.5)
                        _draw_row(pdf, act_time_cells,   act_col_widths, 'B', 9)

                    _draw_row(pdf, row_cells, act_col_widths, '', 9.5)

                render_instructions_sbm(pdf, pdf.get_y() + 2)

                # SAVE THIS PDF TO DICTIONARY USING UNIQUE NAME
                clean_branch = re.sub(r'[^A-Za-z0-9_\- ]', '', main_branch_full).strip().replace(" ", "_")
                clean_sem = re.sub(r'[^A-Za-z0-9_\- ]', '', display_sem).strip().replace(" ", "_")
                filename = f"{clean_branch}_Trimester_{clean_sem}.pdf"
                
                base_filename = filename
                counter = 1
                while filename in pdf_outputs:
                    filename = f"{base_filename.replace('.pdf', '')}_{counter}.pdf"
                    counter += 1
                
                temp_path = f"temp_{uuid.uuid4().hex}.pdf"
                pdf.output(temp_path)
                with open(temp_path, "rb") as f:
                    pdf_outputs[filename] = f.read()
                os.remove(temp_path)
                sheets_processed += 1

            except Exception as e:
                st.warning(f"Error processing sheet {sheet_name}: {e}")
                continue

    else:
        pdf = FPDF(orientation='L', unit='mm', format='Legal')
        pdf.set_auto_page_break(auto=False, margin=15)
        pdf.alias_nb_pages()
        
        for sheet_name, sheet_df in df_dict.items():
            try:
                if sheet_df.empty: continue
                if hasattr(sheet_df, 'index') and len(sheet_df.index.names) > 1:
                    sheet_df = sheet_df.reset_index()
                
                main_branch_full = ""
                if "_prog_" in sheet_df.columns and not sheet_df["_prog_"].dropna().empty:
                    main_branch_full = str(sheet_df["_prog_"].dropna().iloc[0])
                elif "Program" in sheet_df.columns and not sheet_df["Program"].dropna().empty:
                    main_branch_full = str(sheet_df["Program"].dropna().iloc[0])
                elif "MainBranch" in sheet_df.columns and not sheet_df["MainBranch"].dropna().empty:
                    main_branch_full = str(sheet_df["MainBranch"].dropna().iloc[0])
                
                if "Unnamed" in main_branch_full or main_branch_full == "":
                    if '_|_' in sheet_name:
                        main_branch_full = sheet_name.split('_|_')[0]

                rename_cols = {}
                for col in sheet_df.columns:
                    if str(col).startswith("Unnamed:"):
                        rename_cols[col] = main_branch_full
                if rename_cols:
                    sheet_df = sheet_df.rename(columns=rename_cols)
                
                sheet_college_name = st.session_state.get('selected_college', "SVKM's NMIMS University")
                
                semester_raw = "General"
                if '_|_' in sheet_name:
                    parts = sheet_name.split('_|_')
                    if not main_branch_full: main_branch_full = parts[0]
                    semester_raw = parts[1]
                else:
                    if sheet_name in ["No_Data", "Daily_Statistics", "Summary", "Verification", "Empty"]: continue
                    if not main_branch_full: main_branch_full = sheet_name

                is_elective = False
                
                if re.search(r'_Ele(\d*)$', semester_raw):
                    semester_raw = re.sub(r'_Ele\d*$', '', semester_raw)
                    is_elective = True

                if "_sem_" in sheet_df.columns and not sheet_df["_sem_"].dropna().empty:
                    display_sem = str(sheet_df["_sem_"].dropna().iloc[0]).strip()
                else:
                    display_sem = semester_raw.strip()
                    if display_sem.lower().startswith("semester"):
                        display_sem = display_sem[8:].strip()
                    elif display_sem.lower().startswith("sem"):
                        display_sem = display_sem[3:].strip()
                
                header_content = {'main_branch_full': main_branch_full, 'semester_roman': display_sem}
                header_exam_time = get_header_time_for_semester(f"Sem {display_sem}")

                if not is_elective:
                    if 'Exam Date' not in sheet_df.columns: continue
                    sheet_df = sheet_df.dropna(how='all').reset_index(drop=True)
                    fixed_cols = ["Exam Date"]
                    _meta_pattern = re.compile(
                        r'^(Program|Semester|MainBranch|Note|Message|_prog_|_sem_)(\.\d+)?$',
                        re.IGNORECASE
                    )
                    sub_branch_cols = [c for c in sheet_df.columns if c not in fixed_cols and not _meta_pattern.match(str(c)) and pd.notna(c) and str(c).strip() != '']
                    if not sub_branch_cols: continue
                    
                    cols_per_page = 6
                    
                    for start in range(0, len(sub_branch_cols), cols_per_page):
                        chunk = sub_branch_cols[start:start + cols_per_page]
                        cols_to_print = fixed_cols + chunk
                        chunk_df = sheet_df[cols_to_print].copy()
                        
                        subset = chunk_df[chunk].astype(str).apply(lambda x: x.str.strip())
                        valid_cells = (subset != "") & (subset != "nan") & (subset != "---")
                        mask = valid_cells.any(axis=1)
                        chunk_df = chunk_df[mask].reset_index(drop=True)
                        if chunk_df.empty: continue

                        try:
                            chunk_df["Exam Date"] = pd.to_datetime(chunk_df["Exam Date"], format="%d-%m-%Y", errors='coerce').dt.strftime("%A, %d %B, %Y")
                        except: pass

                        page_width = pdf.w - 2 * pdf.l_margin 
                        date_col_width = 30
                        remaining_width = page_width - date_col_width
                        
                        num_sub = max(len(chunk), 1)
                        sub_width = remaining_width / num_sub
                        
                        col_widths = [date_col_width] + [sub_width] * len(chunk)
                        
                        pdf.add_page()
                        
                        original_college = st.session_state.get('selected_college')
                        st.session_state['selected_college'] = sheet_college_name
                        
                        print_table_custom(pdf, chunk_df, cols_to_print, col_widths, line_height=5, 
                                         header_content=header_content, Programs=chunk, 
                                         time_slot=header_exam_time, actual_time_slots=None, 
                                         declaration_date=declaration_date)
                        
                        if original_college: st.session_state['selected_college'] = original_college
                        sheets_processed += 1
                else:
                    target_cols = ['Exam Date', 'OE Type', 'Open Elective (All Applicable Streams)']
                    available_cols = [c for c in target_cols if c in sheet_df.columns]
                    
                    if len(available_cols) >= 3:
                        sheet_df = sheet_df.dropna(subset=['Exam Date']).reset_index(drop=True)
                        if sheet_df.empty: continue

                        try:
                            sheet_df["Exam Date"] = pd.to_datetime(sheet_df["Exam Date"], format="%d-%m-%Y", errors='coerce').dt.strftime("%A, %d %B, %Y")
                        except: pass

                        pdf.add_page()
                        
                        col_widths = [30, 25] 
                        remaining_width = pdf.w - 2 * pdf.l_margin - sum(col_widths)
                        col_widths.append(remaining_width)
                        
                        original_college = st.session_state.get('selected_college')
                        st.session_state['selected_college'] = sheet_college_name
                        
                        print_table_custom(pdf, sheet_df, available_cols, col_widths, line_height=5, 
                                         header_content=header_content, Programs=["Electives"], 
                                         time_slot=header_exam_time, actual_time_slots=None, 
                                         declaration_date=declaration_date)
                        
                        if original_college: st.session_state['selected_college'] = original_college
                        sheets_processed += 1
                    
            except Exception as e:
                st.warning(f"Error processing PDF sheet {sheet_name}: {e}")
                continue

        if sheets_processed > 0:
            try:
                pdf.add_page()
                pdf.set_xy(10, pdf.h - 20)
                pdf.set_font("Times", 'B', 8)
                pdf.cell(0, 5, "CONTROLLER OF EXAMINATIONS", 0, 1, 'L')
                pdf.line(10, pdf.h - 15, 60, pdf.h - 15)
                pdf.set_font("Times", size=9)
                pdf.set_xy(pdf.w - 30, pdf.h - 15)
                pdf.cell(20, 5, f"{pdf.page_no()} of {{nb}}", 0, 0, 'R')

                pdf.set_y(0)
                if os.path.exists(LOGO_PATH): pdf.image(LOGO_PATH, x=(pdf.w-45)/2, y=5, w=45)
                pdf.set_text_color(0, 0, 0)
                pdf.set_font("Times", 'B', 12)
                pdf.set_xy(10, 25)
                pdf.cell(pdf.w - 20, 8, st.session_state.get('selected_college', "SVKM's NMIMS University").upper(), 0, 1, 'C')
                
                pdf.set_font("Times", 'B', 10)
                pdf.set_text_color(0, 0, 0)
                pdf.set_xy(10, 40)
                pdf.cell(0, 10, "EXAMINATION GUIDELINES - SEMESTER GENERAL", 0, 1, 'C')
                
                pdf.set_y(60)
                pdf.set_font("Times", 'B', 12)
                pdf.cell(0, 10, "INSTRUCTIONS TO CANDIDATES", 0, 1, 'C')
                pdf.ln(5)
                instrs = [
                    "1. Candidates are required to be present at the examination center THIRTY MINUTES before the stipulated time.",
                    "2. Candidates must produce their University Identity Card at the time of the examination.",
                    "3. Candidates are not permitted to enter the examination hall after stipulated time.",
                    "4. Candidates will not be permitted to leave the examination hall during the examination time."
                ]
                pdf.set_font("Times", size=9)
                for i in instrs:
                    pdf.multi_cell(0, 7, i)
                    pdf.ln(2)
            except Exception as e:
                pass

            temp_path = f"temp_{uuid.uuid4().hex}.pdf"
            pdf.output(temp_path)
            with open(temp_path, "rb") as f:
                pdf_outputs["Verification_Timetable.pdf"] = f.read()
            os.remove(temp_path)
        else:
            st.error("No valid sheets generated in PDF.")
            return {}

    return pdf_outputs


# ==========================================
# 🔄 GENERATE PDF TIMETABLE (ORCHESTRATOR)
# ==========================================
def generate_pdf_timetable(semester_wise_timetable, output_pdf_base, declaration_date=None):
    import zipfile
    temp_dir   = os.path.dirname(output_pdf_base) if os.path.dirname(output_pdf_base) else "."
    temp_excel = os.path.join(temp_dir, "temp_timetable.xlsx")

    excel_data = save_to_excel(semester_wise_timetable)

    if not excel_data:
        st.error("❌ No Excel data generated — cannot create PDF")
        return None

    try:
        with open(temp_excel, "wb") as f:
            f.write(excel_data.getvalue())
    except Exception as e:
        st.error(f"❌ Error saving temporary Excel file: {e}")
        return None

    try:
        pdf_dict = convert_excel_to_pdf(temp_excel, output_pdf_base, declaration_date=declaration_date)
    except Exception as e:
        st.error(f"❌ Error during Excel to PDF conversion: {e}")
        st.error(traceback.format_exc())
        return None
    finally:
        try:
            if os.path.exists(temp_excel): os.remove(temp_excel)
        except Exception:
            pass
            
    if not pdf_dict:
        st.error("❌ No PDFs were generated.")
        return None

    try:
        final_pdfs = {}
        page_number_pattern = re.compile(r'^[\s\n]*(?:Page\s*)?\d+[\s\n]*$')

        for filename, pdf_bytes in pdf_dict.items():
            reader = PdfReader(io.BytesIO(pdf_bytes))
            writer = PdfWriter()
            pages_kept = 0
            for page in reader.pages:
                try: text = page.extract_text() or ""
                except: text = ""
                cleaned = text.strip()
                if cleaned and not page_number_pattern.match(cleaned) and len(cleaned) > 10:
                    writer.add_page(page)
                    pages_kept += 1
            if pages_kept > 0:
                out_buffer = io.BytesIO()
                writer.write(out_buffer)
                final_pdfs[filename] = out_buffer.getvalue()

        if not final_pdfs:
            st.warning("⚠️ All pages were filtered — keeping original PDF")
            return None

        # Return single PDF or ZIP depending on SBM conditions
        if len(final_pdfs) == 1:
            final_path = output_pdf_base
            with open(final_path, 'wb') as f:
                f.write(list(final_pdfs.values())[0])
            st.success(f"✅ PDF generation complete")
            return final_path
        else:
            final_path = output_pdf_base.replace('.pdf', '.zip')
            with zipfile.ZipFile(final_path, "w", zipfile.ZIP_DEFLATED) as zf:
                for fname, fbytes in final_pdfs.items():
                    zf.writestr(fname, fbytes)
            st.success(f"✅ Generated {len(final_pdfs)} individual PDFs bundled into a ZIP file!")
            return final_path

    except Exception as e:
        st.error(f"❌ PDF post-processing error: {e}")
        st.error(traceback.format_exc())
        return None


# ==========================================
# 🚀 MAIN APP
# ==========================================
def main():
    st.markdown(
        '<div class="main-header">'
        '<h1>📄 Verification File to PDF Converter</h1>'
        '<p>Generate Standardized Timetables Directly from the Verification Excel</p>'
        '</div>',
        unsafe_allow_html=True
    )

    for key in ('pdf_data', 'raw_df', 'processed_tt'):
        if key not in st.session_state: st.session_state[key] = None

    with st.sidebar:
        st.header("⚙️ Configuration")
        college_options = [c["name"] for c in COLLEGES]
        selected_display = st.selectbox("Select College", college_options)
        st.session_state.selected_college = selected_display
        st.markdown("---")
        decl_date = st.date_input("📆 Declaration Date (Optional)", value=None)

    col1, col2 = st.columns([2, 1])

    with col1:
        st.markdown(
            '<div class="upload-section">'
            '<h3 style="margin:0 0 1rem 0; color:#951C1C;">📁 Upload Verification Excel</h3>'
            '<p style="margin:0; color:#666; font-size:1rem;">Drag and drop the Verification file exported from the primary app</p>'
            '</div>',
            unsafe_allow_html=True
        )
        uploaded = st.file_uploader("Upload Excel", type=['xlsx', 'xls'], label_visibility="collapsed")

        if uploaded:
            with st.spinner("Parsing verification file…"):
                tt, df = process_verification_file(uploaded)
                if tt:
                    st.session_state.processed_tt = tt
                    st.session_state.raw_df        = df
                    st.success(f"✅ Loaded successfully — {len(df)} scheduled subjects across {len(tt)} semester(s)")

    with col2:
        st.info(
            "ℹ️ **Direct Converter Mode**\n\n"
            "Processes scheduled subjects and replicates the exact A4 Landscape PDF layout "
            "from the primary Exam Timetable Generator — same header, footer, fonts, "
            "bold time-block rendering, and boundary math."
        )

    if st.session_state.raw_df is not None:
        df = st.session_state.raw_df

        if 'Capacity Exceeded Limit' in df.columns:
            flags = df[df['Capacity Exceeded Limit'].astype(str).str.strip().str.upper() == 'YES (MUMBAI LIMIT HIT)']
            if not flags.empty:
                st.warning(f"⚠️ {len(flags)} subject(s) exceeded capacity limits (Mumbai Limit Hit).")

        st.markdown("---")

        if st.button("🚀 Generate PDF Timetable", type="primary", use_container_width=True):
            with st.spinner("Rendering PDF…"):
                pdf_filename_base = "Verification_Timetable.pdf"
                final_file = generate_pdf_timetable(
                    st.session_state.processed_tt,
                    pdf_filename_base,
                    declaration_date=decl_date
                )

                if final_file and os.path.exists(final_file):
                    st.balloons()
                    with open(final_file, "rb") as f:
                        is_zip = final_file.endswith('.zip')
                        dl_name = f"Verification_Timetable_{datetime.now().strftime('%Y%m%d_%H%M')}{'.zip' if is_zip else '.pdf'}"
                        mime_type = "application/zip" if is_zip else "application/pdf"
                        
                        st.download_button(
                            "📥 Download ZIP" if is_zip else "📥 Download PDF",
                            f,
                            dl_name,
                            mime_type,
                            use_container_width=True
                        )
                    try: os.remove(final_file)
                    except: pass


if __name__ == "__main__":
    main()
