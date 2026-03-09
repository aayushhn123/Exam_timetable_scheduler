import streamlit as st
import pandas as pd
from datetime import datetime
from fpdf import FPDF
import os
import re
import io
import traceback
from PyPDF2 import PdfReader, PdfWriter

# ==========================================
# ⚙️ PAGE CONFIGURATION
# ==========================================
st.set_page_config(
    page_title="Re-Exam Data to PDF Converter",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

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
        background: linear-gradient(135deg, #0A2540 0%, #1A4971 100%);
        padding: 2.5rem; border-radius: 16px; margin-bottom: 2rem;
        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.12);
    }
    .main-header h1 { color: white; text-align: center; margin: 0; font-size: 2.5rem; font-weight: 700; text-shadow: 2px 2px 4px rgba(0,0,0,0.3); letter-spacing: -0.5px; }
    .main-header p { color: #FFF; text-align: center; margin: 0.5rem 0 0 0; font-size: 1.1rem; opacity: 0.95; font-weight: 500; }
    .upload-section { background: #f8f9fa; padding: 2.5rem; border-radius: 16px; border: 2px dashed #0A2540; margin: 1rem 0; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08); text-align: center; }
    @media (prefers-color-scheme: dark) {
        .main-header { background: linear-gradient(135deg, #07192A 0%, #133451 100%); }
        .upload-section { background: #2d2d2d; border-color: #133451; }
    }
    .stButton>button { transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1); border-radius: 12px; font-weight: 600; }
    .stButton>button:hover { transform: translateY(-2px); box-shadow: 0 8px 16px rgba(10, 37, 64, 0.3); border-color: #0A2540; }
</style>
""", unsafe_allow_html=True)


# ==========================================
# 📥 RE-EXAM DATA PARSING
# ==========================================
def process_reexam_file(uploaded_file):
    try:
        if uploaded_file.name.lower().endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
            
        df.columns = df.columns.str.strip()

        required_cols = ['Module Abbreviation', 'Module Description', 'Program', 'Current Session', 'Exam Date', 'Exam Time']
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            st.error(f"❌ Missing critical columns in Re-Exam Data: {', '.join(missing)}")
            return None, None

        # Base Data Cleaning
        df = df[df['Exam Date'].notna()].copy()
        parsed_dates = pd.to_datetime(df['Exam Date'], errors='coerce')
        df = df[parsed_dates.notna()].copy()
        df['Exam Date'] = parsed_dates[parsed_dates.notna()].dt.strftime('%d-%m-%Y')

        df['MainBranch'] = df['Program'].fillna("").astype(str).str.strip()
        df['SubBranch']  = df.get('Stream', pd.Series(dtype=str)).fillna("").astype(str).str.strip()
        df['SubBranch']  = df.apply(lambda x: x['MainBranch'] if x['SubBranch'] in ['nan', '', 'None'] else x['SubBranch'], axis=1)

        df['Subject']    = df['Module Description'].fillna("").astype(str).str.strip()
        df['ModuleCode'] = df['Module Abbreviation'].fillna("").astype(str).str.strip()
        df['Exam Time']  = df['Exam Time'].fillna("").astype(str).str.strip()

        if 'Subject Type' in df.columns:
            df['OE'] = df['Subject Type'].apply(lambda x: str(x).strip() if pd.notna(x) and 'OE' in str(x).upper() else None)
        else:
            df['OE'] = None

        def get_sem_int(val):
            s = str(val).upper().strip()
            m = re.search(r'(\d+)', s)
            if m: return int(m.group(1))
            roman_map = {'XII': 12, 'XI': 11, 'X': 10, 'IX': 9, 'VIII': 8, 'VII': 7, 'VI': 6, 'V': 5, 'IV': 4, 'III': 3, 'II': 2, 'I': 1}
            for r, i in roman_map.items():
                if r == s or s.endswith(f" {r}") or s.endswith(f"_{r}"): return i
            return 1

        df['Semester'] = df['Current Session'].apply(get_sem_int)

        timetable = {}
        for sem in sorted(df['Semester'].unique()):
            timetable[sem] = df[df['Semester'] == sem].copy()

        return timetable, df

    except Exception as e:
        st.error(f"Error parsing re-exam file: {e}")
        st.error(traceback.format_exc())
        return None, None


# ==========================================
# 💾 EXCEL PIVOT ENGINE
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

def save_to_excel(semester_wise_timetable):
    time_slots_dict = st.session_state.get('time_slots', {
        1: {"start": "10:00 AM", "end": "1:00 PM"},
        2: {"start": "2:00 PM",  "end": "5:00 PM"}
    })
    output = io.BytesIO()

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

                prog_key   = _make_program_abbrev(main_branch)
                core_sheet = unique_sheet(f"{prog_key}_|_Sem {roman_sem}"[:31])
                elec_sheet = unique_sheet(f"{prog_key}_|_Sem {roman_sem}_Ele"[:31])

                df_core = df_mb[df_mb['OE'].isna()].copy()
                df_elec = df_mb[df_mb['OE'].notna()].copy()

                if not df_core.empty:
                    displays = []
                    sort_times = []
                    for _, row in df_core.iterrows():
                        subj        = row['Subject']
                        code        = row['ModuleCode']
                        actual_time = str(row.get('Exam Time', '')).strip()

                        time_suffix = ""
                        if actual_time and normalize_time(actual_time) != header_norm and actual_time.lower() not in ['tbd', 'nan', '']:
                            time_suffix = f" [{actual_time}]"

                        # Rule: Core subjects show module code
                        txt = f"{subj}"
                        if code and str(code).lower() != 'nan': txt += f" - ({code})"
                        txt += time_suffix
                        displays.append(txt)

                        # Chronological sorting math
                        parse_time_str = actual_time if (actual_time and actual_time.lower() not in ['tbd', 'nan', '']) else header_norm
                        m = re.search(r'(\d{1,2}):(\d{2})\s*([AP]M)', str(parse_time_str).upper())
                        if m:
                            h, mins = int(m.group(1)), int(m.group(2))
                            if 'PM' in m.group(3) and h < 12: h += 12
                            if 'AM' in m.group(3) and h == 12: h = 0
                            sort_times.append(h * 60 + mins)
                        else:
                            sort_times.append(9999)

                    df_core['SubjectDisplay'] = displays
                    df_core['_SortTime'] = sort_times
                    df_core["Exam Date"] = pd.to_datetime(df_core["Exam Date"], format="%d-%m-%Y", dayfirst=True, errors='coerce')
                    
                    df_core = df_core.sort_values(by=["Exam Date", "_SortTime"], ascending=[True, True])

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

                        # Rule: Open Electives DO NOT show module code
                        txt = f"{subj}{time_suffix}"
                        e_displays.append(txt)

                    df_elec['DisplaySubject'] = e_displays

                    try:
                        df_elec["Exam Date"] = pd.to_datetime(df_elec["Exam Date"], format="%d-%m-%Y", dayfirst=True, errors='coerce')
                        df_elec = df_elec.sort_values(by="Exam Date", ascending=True)
                        df_elec['Exam Date'] = df_elec['Exam Date'].apply(lambda x: x.strftime("%d-%m-%Y") if pd.notna(x) else "")

                        # Group by OE Type and comma separate 
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
# 📄 FPDF ENGINE — STRICT REPLICA
# ==========================================
def wrap_text(pdf, text, col_width):
    cache_key = (text, col_width, pdf.font_style)
    if cache_key in wrap_text_cache: return wrap_text_cache[cache_key]
        
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
    old_family, old_style, old_size = pdf.font_family, pdf.font_style, pdf.font_size_pt
    
    for token in tokens:
        if token == "<hr>":
            if current_line: lines.append(current_line); current_line = ""
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
            if current_line: lines.append(current_line)
            current_line = token
            
    if current_line: lines.append(current_line)
    pdf.set_font(old_family, old_style, old_size)
    wrap_text_cache[cache_key] = lines
    return lines

def print_row_custom(pdf, row_data, col_widths, line_height=5, header=False):
    cell_padding = 1
    # STRICT B&W REQUIREMENT
    header_bg_color = (255, 255, 255)
    header_text_color = (0, 0, 0)
    alt_row_color = (255, 255, 255)

    row_number = getattr(pdf, '_row_counter', 0)
    base_font = "Times"
    
    if header:
        base_style, base_size = 'B', 9.5
        pdf.set_font(base_font, base_style, base_size)
        pdf.set_text_color(*header_text_color)
        pdf.set_fill_color(*header_bg_color)
    else:
        base_style, base_size = '', 9.5
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
    text_line_height = line_height * 0.75 
    
    x0, y0 = pdf.get_x(), pdf.get_y()
    pdf.rect(x0, y0, sum(col_widths), row_h, 'F')

    time_pattern = re.compile(r'([\[\(]\s*\d{1,2}:\d{2}\s*[AP]M\s*-\s*\d{1,2}:\d{2}\s*[AP]M\s*[\]\)])', re.IGNORECASE)

    for i, lines in enumerate(wrapped_cells):
        cx = pdf.get_x()
        
        # Split by <hr> for mathematical partition logic
        subjects_lines = []
        current_subject = []
        for ln in lines:
            if ln == "<hr>":
                subjects_lines.append(current_subject)
                current_subject = []
            else:
                current_subject.append(ln)
        subjects_lines.append(current_subject)
        
        num_subjects = len(subjects_lines)
        part_h = row_h / num_subjects if num_subjects > 0 else row_h
        
        for sub_idx, subj_lines in enumerate(subjects_lines):
            total_text_h = len(subj_lines) * text_line_height
            pad_v = (part_h - total_text_h) / 2
            
            for j, ln in enumerate(subj_lines):
                parts = time_pattern.split(ln)
                
                if len(parts) == 1 or header:
                    pdf.set_xy(cx + cell_padding, y0 + (sub_idx * part_h) + pad_v + j * text_line_height)
                    pdf.cell(col_widths[i] - 2 * cell_padding, text_line_height, ln, border=0, align='C')
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
                        pdf.set_xy(current_x - pdf.c_margin, y0 + (sub_idx * part_h) + pad_v + j * text_line_height)
                        pdf.cell(w + 2 * pdf.c_margin, text_line_height, p, border=0, align='L')
                        current_x += w
                    
                    pdf.set_font(base_font, base_style, base_size)
            
            # Draw physical horizontal line exactly on boundary
            if sub_idx < num_subjects - 1:
                line_y = y0 + ((sub_idx + 1) * part_h)
                pdf.line(cx, line_y, cx + col_widths[i], line_y)
        
        pdf.rect(cx, y0, col_widths[i], row_h)
        pdf.set_xy(cx + col_widths[i], y0)

    setattr(pdf, '_row_counter', row_number + 1)
    pdf.set_xy(x0, y0 + row_h)

def print_table_custom(pdf, df, columns, col_widths, line_height=5, header_content=None, Programs=None, time_slot=None, declaration_date=None):
    if df.empty: return
    setattr(pdf, '_row_counter', 0)
    
    footer_height = 14   
    header_end_y = 60    
    
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
            suffix = 'TH' if 11 <= (day % 100) <= 13 else {1: 'ST', 2: 'ND', 3: 'RD'}.get(day % 10, 'TH')
            decl_str = f"DATE: {day}{suffix} {declaration_date.strftime('%B, %Y')}".upper()
            
            pdf.set_font("Times", 'B', 12) 
            pdf.set_text_color(0, 0, 0)
            pdf.set_xy(pdf.w - 80, 8)
            pdf.cell(70, 10, decl_str, 0, 0, 'R')

        logo_width = 45
        if os.path.exists(LOGO_PATH): pdf.image(LOGO_PATH, x=(pdf.w - logo_width) / 2, y=5, w=logo_width)
        
        pdf.set_text_color(0, 0, 0)
        college_name = st.session_state.get('selected_college', "SVKM's NMIMS University").upper()
        pdf.set_font("Times", 'B', 12) 
        pdf.set_xy(10, 25)
        pdf.cell(pdf.w - 20, 6, college_name, 0, 1, 'C')
        
        pdf.set_font("Times", 'B', 10) 
        pdf.set_xy(10, 33)
        pdf.cell(pdf.w - 20, 4, "FINAL EXAMINATION TIMETABLE (ACADEMIC YEAR: 2025-26)", 0, 1, 'C')
        
        current_y = 38
        pdf.set_font("Times", 'B', 10) 
        pdf.set_xy(10, current_y)
        pdf.cell(pdf.w - 20, 4, f"{header_content['main_branch_full']}".upper(), 0, 1, 'C')
        current_y += 4
        
        sem_roman = str(header_content['semester_roman']).upper()
        roman_map = {'XII': 12, 'XI': 11, 'X': 10, 'IX': 9, 'VIII': 8, 'VII': 7, 'VI': 6, 'V': 5, 'IV': 4, 'III': 3, 'II': 2, 'I': 1}
        sem_int = roman_map.get(sem_roman)
        if not sem_int:
            m = re.search(r'(\d+)', sem_roman)
            sem_int = int(m.group(1)) if m else 1
        year_roman = int_to_roman((sem_int + 1) // 2)

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

        pdf.set_xy(pdf.l_margin, header_end_y)

    render_footer()
    render_header()
    
    # UPPERCASE TABLE COLUMNS
    upper_columns = [str(c).upper() for c in columns]
    pdf.set_font("Times", 'B', 9.5)
    print_row_custom(pdf, upper_columns, col_widths, line_height=line_height, header=True)
    pdf.set_font("Times", '', 9.5) 
    
    for idx in range(len(df)):
        row = [str(df.iloc[idx][c]) if pd.notna(df.iloc[idx][c]) else "" for c in columns]
        if not any(cell.strip() for cell in row): continue
            
        wrapped_cells = []
        max_lines = 0
        for i, cell_text in enumerate(row):
            avail_w = col_widths[i] - 2 
            lines = wrap_text(pdf, str(cell_text), avail_w)
            wrapped_cells.append(lines)
            max_lines = max(max_lines, len(lines))
        row_h = line_height * max_lines
        
        if pdf.get_y() + row_h > pdf.h - footer_height - 5:
            pdf.add_page()
            render_footer()
            render_header()
            pdf.set_font("Times", 'B', 9.5) 
            print_row_custom(pdf, upper_columns, col_widths, line_height=line_height, header=True)
            pdf.set_font("Times", '', 9.5)  
        
        print_row_custom(pdf, row, col_widths, line_height=line_height, header=False)

def convert_excel_to_pdf(excel_path, pdf_path, declaration_date=None):
    pdf = FPDF(orientation='L', unit='mm', format='Legal')
    pdf.set_auto_page_break(auto=False, margin=15)
    pdf.alias_nb_pages()
    
    time_slots_dict = st.session_state.get('time_slots', {
        1: {"start": "10:00 AM", "end": "1:00 PM"},
        2: {"start": "2:00 PM", "end": "5:00 PM"}
    })
    
    try: df_dict = pd.read_excel(excel_path, sheet_name=None)
    except Exception as e: st.error(f"Error reading Excel file: {e}"); return

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
        except: return f"{time_slots_dict[1]['start']} - {time_slots_dict[1]['end']}"

    sheets_processed = 0
    for sheet_name, sheet_df in df_dict.items():
        try:
            if sheet_df.empty or sheet_name in ["Empty"]: continue
            
            main_branch_full = ""
            if "_prog_" in sheet_df.columns: main_branch_full = str(sheet_df["_prog_"].dropna().iloc[0])
            
            rename_cols = {}
            for col in sheet_df.columns:
                if str(col).startswith("Unnamed:"): rename_cols[col] = main_branch_full
            if rename_cols: sheet_df = sheet_df.rename(columns=rename_cols)
            
            semester_raw = sheet_name.split('_|_')[1] if '_|_' in sheet_name else sheet_name
            is_elective = bool(re.search(r'_Ele(\d*)$', semester_raw))
            semester_raw = re.sub(r'_Ele\d*$', '', semester_raw)

            display_sem = str(sheet_df["_sem_"].dropna().iloc[0]).strip() if "_sem_" in sheet_df.columns else semester_raw.replace("Sem ", "").strip()
            
            header_content = {'main_branch_full': main_branch_full, 'semester_roman': display_sem}
            header_exam_time = get_header_time_for_semester(f"Sem {display_sem}")

            if not is_elective:
                if 'Exam Date' not in sheet_df.columns: continue
                sheet_df = sheet_df.dropna(how='all').reset_index(drop=True)
                fixed_cols = ["Exam Date"]
                _meta_pattern = re.compile(r'^(Program|Semester|MainBranch|Note|Message|_prog_|_sem_)(\.\d+)?$', re.IGNORECASE)
                sub_branch_cols = [c for c in sheet_df.columns if c not in fixed_cols and not _meta_pattern.match(str(c)) and pd.notna(c) and str(c).strip() != '']
                if not sub_branch_cols: continue
                
                cols_per_page = 6
                for start in range(0, len(sub_branch_cols), cols_per_page):
                    chunk = sub_branch_cols[start:start + cols_per_page]
                    cols_to_print = fixed_cols + chunk
                    chunk_df = sheet_df[cols_to_print].copy()
                    
                    subset = chunk_df[chunk].astype(str).apply(lambda x: x.str.strip())
                    chunk_df = chunk_df[(subset != "") & (subset != "nan") & (subset != "---")].dropna(how='all').reset_index(drop=True)
                    if chunk_df.empty: continue

                    try: chunk_df["Exam Date"] = pd.to_datetime(chunk_df["Exam Date"], format="%d-%m-%Y", errors='coerce').dt.strftime("%A, %d %B, %Y")
                    except: pass

                    page_width = pdf.w - 2 * pdf.l_margin 
                    date_col_width = 30
                    sub_width = (page_width - date_col_width) / max(len(chunk), 1)
                    col_widths = [date_col_width] + [sub_width] * len(chunk)
                    
                    pdf.add_page()
                    print_table_custom(pdf, chunk_df, cols_to_print, col_widths, line_height=5, 
                                     header_content=header_content, time_slot=header_exam_time, 
                                     declaration_date=declaration_date)
                    sheets_processed += 1
            else:
                target_cols = ['Exam Date', 'OE Type', 'OPEN ELECTIVE (ALL APPLICABLE STREAMS)']
                
                # Correct internal column mapping from excel output mapping
                if 'Open Elective (All Applicable Streams)' in sheet_df.columns:
                    sheet_df = sheet_df.rename(columns={'Open Elective (All Applicable Streams)': 'OPEN ELECTIVE (ALL APPLICABLE STREAMS)'})

                available_cols = [c for c in target_cols if c in sheet_df.columns]
                
                if len(available_cols) >= 3:
                    sheet_df = sheet_df.dropna(subset=['Exam Date']).reset_index(drop=True)
                    if sheet_df.empty: continue
                    try: sheet_df["Exam Date"] = pd.to_datetime(sheet_df["Exam Date"], format="%d-%m-%Y", errors='coerce').dt.strftime("%A, %d %B, %Y")
                    except: pass

                    pdf.add_page()
                    col_widths = [30, 25, pdf.w - 2 * pdf.l_margin - 55] 
                    
                    print_table_custom(pdf, sheet_df, available_cols, col_widths, line_height=5, 
                                     header_content=header_content, time_slot=header_exam_time, 
                                     declaration_date=declaration_date)
                    sheets_processed += 1
        except Exception as e:
            st.warning(f"Error processing PDF sheet {sheet_name}: {e}")
            continue

    if sheets_processed == 0: st.error("No valid sheets generated in PDF."); return

    # Guidelines Page (12 Bold / 12 Regular strict layout)
    try:
        pdf.add_page()
        pdf.set_xy(10, pdf.h - 20)
        pdf.set_font("Times", 'B', 8)
        pdf.cell(0, 5, "CONTROLLER OF EXAMINATIONS", 0, 1, 'L')
        pdf.line(10, pdf.h - 15, 60, pdf.h - 15)
        pdf.set_font("Times", size=8)
        pdf.set_xy(pdf.w - 30, pdf.h - 15)
        pdf.cell(20, 5, f"{pdf.page_no()} of {{nb}}", 0, 0, 'R')

        pdf.set_y(0)
        if os.path.exists(LOGO_PATH): pdf.image(LOGO_PATH, x=(pdf.w-45)/2, y=5, w=45)
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Times", 'B', 12)
        pdf.set_xy(10, 25)
        pdf.cell(pdf.w - 20, 8, st.session_state.get('selected_college', "SVKM's NMIMS University").upper(), 0, 1, 'C')
        
        pdf.set_font("Times", 'B', 12)
        pdf.set_xy(10, 40)
        pdf.cell(0, 10, "EXAMINATION GUIDELINES - SEMESTER GENERAL", 0, 1, 'C')
        
        pdf.set_y(60)
        pdf.cell(0, 10, "INSTRUCTIONS TO CANDIDATES", 0, 1, 'C')
        pdf.ln(5)
        instrs = [
            "1. Candidates are required to be present at the examination center THIRTY MINUTES before the stipulated time.",
            "2. Candidates must produce their University Identity Card at the time of the examination.",
            "3. Candidates are not permitted to enter the examination hall after stipulated time.",
            "4. Candidates will not be permitted to leave the examination hall during the examination time."
        ]
        pdf.set_font("Times", size=12)
        for i in instrs:
            pdf.multi_cell(0, 7, i)
            pdf.ln(2)
    except Exception: pass

    try: pdf.output(pdf_path)
    except Exception as e: st.error(f"Save PDF failed: {e}")

def generate_pdf_timetable(semester_wise_timetable, output_pdf, declaration_date=None):
    temp_dir   = os.path.dirname(output_pdf) if os.path.dirname(output_pdf) else "."
    temp_excel = os.path.join(temp_dir, "temp_reexam.xlsx")

    excel_data = save_to_excel(semester_wise_timetable)
    if not excel_data: st.error("❌ No Excel data generated — cannot create PDF"); return

    try:
        with open(temp_excel, "wb") as f: f.write(excel_data.getvalue())
    except Exception as e: st.error(f"❌ Error saving temporary Excel file: {e}"); return

    try: convert_excel_to_pdf(temp_excel, output_pdf, declaration_date=declaration_date)
    except Exception as e: st.error(f"❌ Error during conversion: {e}"); return
    finally:
        try:
            if os.path.exists(temp_excel): os.remove(temp_excel)
        except Exception: pass

    if os.path.exists(output_pdf): st.success("✅ PDF generation complete")
    else: st.error(f"❌ PDF file was not created at {output_pdf}")

# ==========================================
# 🚀 MAIN APP
# ==========================================
def main():
    st.markdown(
        '<div class="main-header">'
        '<h1>📄 Re-Exam Data to PDF Converter</h1>'
        '<p>Generate Standardized Timetables Directly from Re-Exam CSV/Excel Data</p>'
        '</div>',
        unsafe_allow_html=True
    )

    for key in ('raw_df', 'processed_tt'):
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
            '<h3 style="margin:0 0 1rem 0; color:#0A2540;">📁 Upload Re-Exam Data</h3>'
            '<p style="margin:0; color:#666; font-size:1rem;">Upload the CSV or Excel file containing Re-Exam schedules</p>'
            '</div>',
            unsafe_allow_html=True
        )
        uploaded = st.file_uploader("Upload Data", type=['csv', 'xlsx', 'xls'], label_visibility="collapsed")

        if uploaded:
            with st.spinner("Parsing Re-Exam Data…"):
                tt, df = process_reexam_file(uploaded)
                if tt:
                    st.session_state.processed_tt = tt
                    st.session_state.raw_df       = df
                    st.success(f"✅ Loaded successfully — {len(df)} scheduled subjects parsed.")

    with col2:
        st.info(
            "ℹ️ **Strict Converter Mode**\n\n"
            "This replicates the master FPDF engine: strict Times font, Legal Landscape, B&W layout, "
            "automatic chronologic grouping (`<hr>` partition borders), and specific Elective parsing rules."
        )

    if st.session_state.raw_df is not None:
        st.markdown("---")
        if st.button("🚀 Generate PDF Timetable", type="primary", use_container_width=True):
            with st.spinner("Rendering PDF via Math Boundaries Engine…"):
                pdf_filename = "ReExam_Timetable.pdf"
                generate_pdf_timetable(
                    st.session_state.processed_tt,
                    pdf_filename,
                    declaration_date=decl_date
                )

                if os.path.exists(pdf_filename):
                    st.balloons()
                    with open(pdf_filename, "rb") as f:
                        st.download_button(
                            "📥 Download PDF",
                            f,
                            f"ReExam_Timetable_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                            "application/pdf",
                            use_container_width=True
                        )
                    try: os.remove(pdf_filename)
                    except: pass


if __name__ == "__main__":
    main()
