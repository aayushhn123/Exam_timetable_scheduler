import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
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
    page_title="Re-Exam Auto-Scheduler & PDF Generator",
    page_icon="📅",
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
# 📥 AUTO-SCHEDULER & DATA PARSER
# ==========================================
def schedule_and_process_reexam(uploaded_file, start_date, num_days):
    try:
        # 1. Read File
        if uploaded_file.name.lower().endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
            
        df.columns = df.columns.str.strip()

        required_cols = ['Module Abbreviation', 'Module Description', 'Program', 'Current Session']
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            st.error(f"❌ Missing critical columns in Re-Exam Data: {', '.join(missing)}")
            return None, None

        # 2. Extract unique subjects for exact mapping (ignores Sem/Program differences)
        df['Group_Subject'] = df['Module Description'].fillna("").astype(str).str.strip().str.upper()
        unique_subjects = df['Group_Subject'].unique()

        # 3. Generate Valid Dates (Skipping Sundays)
        valid_dates = []
        current_dt = start_date
        while len(valid_dates) < num_days:
            if current_dt.weekday() != 6:  # 6 is Sunday
                valid_dates.append(current_dt)
            current_dt += timedelta(days=1)

        # 4. Generate Slots (2 slots per valid day)
        slots = []
        for d in valid_dates:
            d_str = d.strftime('%d-%m-%Y')
            slots.append((d_str, "10:00 AM - 01:00 PM"))
            slots.append((d_str, "2:00 PM - 05:00 PM"))

        # 5. Distribute Subjects Across Slots evenly
        subject_schedule_map = {}
        for i, subj in enumerate(unique_subjects):
            assigned_slot = slots[i % len(slots)]
            subject_schedule_map[subj] = assigned_slot

        # 6. Apply Schedule back to DataFrame
        df['Exam Date'] = df['Group_Subject'].map(lambda x: subject_schedule_map[x][0])
        df['Exam Time'] = df['Group_Subject'].map(lambda x: subject_schedule_map[x][1])

        # 7. Map to the Master Engine Structure
        df['MainBranch'] = df['Program'].fillna("").astype(str).str.strip()
        df['SubBranch']  = df.get('Stream', pd.Series(dtype=str)).fillna("").astype(str).str.strip()
        df['SubBranch']  = df.apply(lambda x: x['MainBranch'] if x['SubBranch'] in ['nan', '', 'None'] else x['SubBranch'], axis=1)

        df['Subject']    = df['Module Description'].fillna("").astype(str).str.strip()
        df['ModuleCode'] = df['Module Abbreviation'].fillna("").astype(str).str.strip()

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

        # 8. Build Timetable Dictionary by Semester
        timetable = {}
        for sem in sorted(df['Semester'].unique()):
            timetable[sem] = df[df['Semester'] == sem].copy()

        return timetable, df

    except Exception as e:
        st.error(f"Error parsing and scheduling file: {e}")
        st.error(traceback.format_exc())
        return None, None


# ==========================================
# 💾 EXCEL PIVOT ENGINE (MASTER)
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

                        txt = f"{subj}"
                        # USER UPDATE: Render Module Code without Hyphen
                        if code and str(code).lower() != 'nan': txt += f" ({code})"
                        txt += time_suffix
                        displays.append(txt)

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

                        txt = f"{subj}{time_suffix}"
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

        if sheets_created == 0: pd.DataFrame({'Info': ['No valid data']}).to_excel(writer, sheet_name="Empty")

    output.seek(0)
    return output


# ==========================================
# 📄 FPDF ENGINE — STRICT MASTER REPLICA
# ==========================================
def wrap_text(pdf, text, col_width):
    cache_key = (text, col_width, pdf.font_style)
    if cache_key in wrap_text_cache: return wrap_text_cache[cache_key]
        
    time_pattern = re.compile(r'([\[\(]\s*\d{1,2}:\d{2}\s*[AP]M\s*-\s*\d{1,2}:\d{2}\s*[AP]M\s*[\]\)])', re.IGNORECASE)
    parts = time_pattern.split(str(text))
    tokens = []
    
    for i, p in enumerate(parts):
        if i % 2 == 1: tokens.append(p)
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
        test_w = sum(pdf.get_string_width(pt) for pt in time_pattern.split(test_line) if pt)
        
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
    row_number = getattr(pdf, '_row_counter', 0)
    base_font = "Times"
    
    if header:
        pdf.set_font(base_font, 'B', 9.5)
        pdf.set_text_color(0, 0, 0)
        pdf.set_fill_color(255, 255, 255)
    else:
        pdf.set_font(base_font, '', 9.5)
        pdf.set_text_color(0, 0, 0)
        pdf.set_fill_color(255, 255, 255)

    wrapped_cells = []
    max_lines = 0
    for i, cell_text in enumerate(row_data):
        avail_w = col_widths[i] - 2 * cell_padding
        lines = wrap_text(pdf, str(cell_text) if cell_text is not None else "", avail_w)
        wrapped_cells.append(lines)
        max_lines = max(max_lines, len(lines))

    row_h = line_height * max_lines
    text_line_height = line_height * 0.75 
    x0, y0 = pdf.get_x(), pdf.get_y()
    pdf.rect(x0, y0, sum(col_widths), row_h, 'F')

    time_pattern = re.compile(r'([\[\(]\s*\d{1,2}:\d{2}\s*[AP]M\s*-\s*\d{1,2}:\d{2}\s*[AP]M\s*[\]\)])', re.IGNORECASE)

    for i, lines in enumerate(wrapped_cells):
        cx = pdf.get_x()
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
            pad_v = (part_h - (len(subj_lines) * text_line_height)) / 2
            
            for j, ln in enumerate(subj_lines):
                parts = time_pattern.split(ln)
                if len(parts) == 1 or header:
                    pdf.set_xy(cx + cell_padding, y0 + (sub_idx * part_h) + pad_v + j * text_line_height)
                    pdf.cell(col_widths[i] - 2 * cell_padding, text_line_height, ln, border=0, align='C')
                else:
                    total_w = sum(pdf.get_string_width(p) for p in parts if p)
                    current_x = cx + max(cell_padding, (col_widths[i] - total_w) / 2)
                    for k, p in enumerate(parts):
                        if not p: continue
                        pdf.set_font(base_font, 'B' if k % 2 == 1 else ('' if not header else 'B'), 9.5)
                        w = pdf.get_string_width(p)
                        pdf.set_xy(current_x - pdf.c_margin, y0 + (sub_idx * part_h) + pad_v + j * text_line_height)
                        pdf.cell(w + 2 * pdf.c_margin, text_line_height, p, border=0, align='L')
                        current_x += w
                    pdf.set_font(base_font, '' if not header else 'B', 9.5)
            
            if sub_idx < num_subjects - 1:
                line_y = y0 + ((sub_idx + 1) * part_h)
                pdf.line(cx, line_y, cx + col_widths[i], line_y)
        
        pdf.rect(cx, y0, col_widths[i], row_h)
        pdf.set_xy(cx + col_widths[i], y0)

    setattr(pdf, '_row_counter', row_number + 1)
    pdf.set_xy(x0, y0 + row_h)

def print_table_custom(pdf, df, columns, col_widths, line_height=5, header_content=None, time_slot=None, declaration_date=None):
    if df.empty: return
    setattr(pdf, '_row_counter', 0)
    footer_height, header_end_y = 14, 60    
    
    def render_footer():
        pdf.set_xy(10, pdf.h - footer_height)
        pdf.set_font("Times", 'B', 8) 
        pdf.cell(0, 5, "CONTROLLER OF EXAMINATIONS", 0, 1, 'L')
        pdf.line(10, pdf.h - footer_height + 5, 60, pdf.h - footer_height + 5)
        pdf.set_font("Times", size=8)
        page_text = f"{pdf.page_no()} of {{nb}}"
        text_width = pdf.get_string_width(page_text.replace("{nb}", "99"))
        pdf.set_xy(pdf.w - 10 - text_width, pdf.h - footer_height + 5)
        pdf.cell(text_width, 5, page_text, 0, 0, 'R')

    def render_header():
        pdf.set_y(0)
        if declaration_date:
            day = declaration_date.day
            suffix = 'TH' if 11 <= (day % 100) <= 13 else {1: 'ST', 2: 'ND', 3: 'RD'}.get(day % 10, 'TH')
            pdf.set_font("Times", 'B', 12) 
            pdf.set_xy(pdf.w - 80, 8)
            pdf.cell(70, 10, f"DATE: {day}{suffix} {declaration_date.strftime('%B, %Y')}".upper(), 0, 0, 'R')

        if os.path.exists(LOGO_PATH): pdf.image(LOGO_PATH, x=(pdf.w - 45) / 2, y=5, w=45)
        
        pdf.set_font("Times", 'B', 12) 
        pdf.set_xy(10, 25)
        pdf.cell(pdf.w - 20, 6, st.session_state.get('selected_college', "SVKM's NMIMS University").upper(), 0, 1, 'C')
        
        pdf.set_font("Times", 'B', 10) 
        pdf.set_xy(10, 33)
        pdf.cell(pdf.w - 20, 4, "FINAL EXAMINATION TIMETABLE (ACADEMIC YEAR: 2025-26)", 0, 1, 'C')
        
        current_y = 38
        pdf.set_xy(10, current_y)
        pdf.cell(pdf.w - 20, 4, f"{header_content['main_branch_full']}".upper(), 0, 1, 'C')
        current_y += 4
        
        sem_roman = str(header_content['semester_roman']).upper()
        sem_int = {'XII': 12, 'XI': 11, 'X': 10, 'IX': 9, 'VIII': 8, 'VII': 7, 'VI': 6, 'V': 5, 'IV': 4, 'III': 3, 'II': 2, 'I': 1}.get(sem_roman, 1)
        pdf.set_xy(10, current_y)
        pdf.cell(pdf.w - 20, 4, f"YEAR: {int_to_roman((sem_int + 1) // 2)}, SEMESTER: {sem_roman}".upper(), 0, 1, 'C')
        current_y += 4

        if time_slot:
            pdf.set_font("Times", 'B', 9)
            pdf.set_xy(10, current_y)
            pdf.cell(pdf.w - 20, 4, f"EXAM TIME: {time_slot}".upper(), 0, 1, 'C')
            pdf.set_font("Times", 'BI', 9) 
            pdf.set_xy(10, current_y + 4)
            pdf.cell(pdf.w - 20, 4, "(CHECK THE SUBJECT EXAM TIME)".upper(), 0, 1, 'C')

        pdf.set_xy(pdf.l_margin, header_end_y)

    render_footer()
    render_header()
    
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
            lines = wrap_text(pdf, str(cell_text), col_widths[i] - 2)
            wrapped_cells.append(lines)
            max_lines = max(max_lines, len(lines))
        
        if pdf.get_y() + (line_height * max_lines) > pdf.h - footer_height - 5:
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
    
    try: df_dict = pd.read_excel(excel_path, sheet_name=None)
    except Exception as e: st.error(f"Error reading Excel file: {e}"); return

    def get_header_time_for_semester(sem_str):
        slot_cfg = st.session_state.get('time_slots', {1: {"start": "10:00 AM", "end": "1:00 PM"}, 2: {"start": "2:00 PM", "end": "5:00 PM"}})
        try:
            s = str(sem_str).strip().upper()
            romans = {'XII': 12, 'XI': 11, 'X': 10, 'IX': 9, 'VIII': 8, 'VII': 7, 'VI': 6, 'V': 5, 'IV': 4, 'III': 3, 'II': 2, 'I': 1}
            sem_int = next((v for k, v in romans.items() if s.endswith(k) or s == k), 1)
            cfg = slot_cfg[1 if ((sem_int + 1) // 2) % 2 == 1 else 2]
            return f"{cfg['start']} - {cfg['end']}"
        except: return f"{slot_cfg[1]['start']} - {slot_cfg[1]['end']}"

    sheets_processed = 0
    for sheet_name, sheet_df in df_dict.items():
        try:
            if sheet_df.empty or sheet_name in ["Empty"]: continue
            
            main_branch_full = str(sheet_df["_prog_"].dropna().iloc[0]) if "_prog_" in sheet_df.columns else ""
            sheet_df = sheet_df.rename(columns={c: main_branch_full for c in sheet_df.columns if str(c).startswith("Unnamed:")})
            
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
                
                for start in range(0, len(sub_branch_cols), 6):
                    chunk = sub_branch_cols[start:start + 6]
                    cols_to_print = fixed_cols + chunk
                    chunk_df = sheet_df[cols_to_print].copy()
                    
                    subset = chunk_df[chunk].astype(str).apply(lambda x: x.str.strip())
                    chunk_df = chunk_df[(subset != "") & (subset != "nan") & (subset != "---")].dropna(how='all').reset_index(drop=True)
                    if chunk_df.empty: continue

                    try: chunk_df["Exam Date"] = pd.to_datetime(chunk_df["Exam Date"], format="%d-%m-%Y", errors='coerce').dt.strftime("%A, %d %B, %Y")
                    except: pass

                    col_widths = [30] + [(pdf.w - 2 * pdf.l_margin - 30) / max(len(chunk), 1)] * len(chunk)
                    pdf.add_page()
                    print_table_custom(pdf, chunk_df, cols_to_print, col_widths, line_height=5, header_content=header_content, time_slot=header_exam_time, declaration_date=declaration_date)
                    sheets_processed += 1
            else:
                target_cols = ['Exam Date', 'OE Type', 'OPEN ELECTIVE (ALL APPLICABLE STREAMS)']
                if 'Open Elective (All Applicable Streams)' in sheet_df.columns: sheet_df = sheet_df.rename(columns={'Open Elective (All Applicable Streams)': 'OPEN ELECTIVE (ALL APPLICABLE STREAMS)'})
                available_cols = [c for c in target_cols if c in sheet_df.columns]
                
                if len(available_cols) >= 3:
                    sheet_df = sheet_df.dropna(subset=['Exam Date']).reset_index(drop=True)
                    if sheet_df.empty: continue
                    try: sheet_df["Exam Date"] = pd.to_datetime(sheet_df["Exam Date"], format="%d-%m-%Y", errors='coerce').dt.strftime("%A, %d %B, %Y")
                    except: pass

                    pdf.add_page()
                    print_table_custom(pdf, sheet_df, available_cols, [30, 25, pdf.w - 2 * pdf.l_margin - 55], line_height=5, header_content=header_content, time_slot=header_exam_time, declaration_date=declaration_date)
                    sheets_processed += 1
        except Exception as e:
            continue

    if sheets_processed == 0: st.error("No valid sheets generated in PDF."); return

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
        pdf.set_font("Times", 'B', 12)
        pdf.set_xy(10, 25)
        pdf.cell(pdf.w - 20, 8, st.session_state.get('selected_college', "SVKM's NMIMS University").upper(), 0, 1, 'C')
        pdf.set_xy(10, 40)
        pdf.cell(0, 10, "EXAMINATION GUIDELINES - SEMESTER GENERAL", 0, 1, 'C')
        
        pdf.set_y(60)
        pdf.cell(0, 10, "INSTRUCTIONS TO CANDIDATES", 0, 1, 'C')
        pdf.ln(5)
        pdf.set_font("Times", size=12)
        for i in [
            "1. Candidates are required to be present at the examination center THIRTY MINUTES before the stipulated time.",
            "2. Candidates must produce their University Identity Card at the time of the examination.",
            "3. Candidates are not permitted to enter the examination hall after stipulated time.",
            "4. Candidates will not be permitted to leave the examination hall during the examination time."
        ]:
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

    if os.path.exists(output_pdf): st.success("✅ Auto-Scheduling & PDF generation complete")
    else: st.error(f"❌ PDF file was not created at {output_pdf}")


# ==========================================
# 🚀 MAIN APP
# ==========================================
def main():
    st.markdown(
        '<div class="main-header">'
        '<h1>📅 Re-Exam Auto-Scheduler & PDF Generator</h1>'
        '<p>Upload raw CSV/Excel, automatically schedule dates without gaps, and export strictly to PDF</p>'
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
        
        st.subheader("Scheduling Rules")
        start_date = st.date_input("🗓️ Exam Start Date", value=datetime.now().date())
        num_days = st.number_input("⏱️ Total Consecutive Days", min_value=1, max_value=30, value=11, help="Skips Sundays automatically")
        
        st.markdown("---")
        decl_date = st.date_input("📆 Declaration Date (Optional)", value=None)

    col1, col2 = st.columns([2, 1])

    with col1:
        st.markdown(
            '<div class="upload-section">'
            '<h3 style="margin:0 0 1rem 0; color:#0A2540;">📁 Upload Re-Exam Template Data</h3>'
            '<p style="margin:0; color:#666; font-size:1rem;">Upload CSV or Excel containing unscheduled Re-Exam subjects</p>'
            '</div>',
            unsafe_allow_html=True
        )
        uploaded = st.file_uploader("Upload Data", type=['csv', 'xlsx', 'xls'], label_visibility="collapsed")

        if uploaded:
            with st.spinner("Auto-Scheduling subjects across dates..."):
                tt, df = schedule_and_process_reexam(uploaded, start_date, num_days)
                if tt:
                    st.session_state.processed_tt = tt
                    st.session_state.raw_df       = df
                    st.success(f"✅ Auto-Scheduled Successfully: {len(df['Module Description'].unique())} unique subjects mapped across {num_days} days.")

    with col2:
        st.info(
            "ℹ️ **Scheduler Constraints Active**\n\n"
            "- Subjects sharing the exact same name are scheduled on the same Date and Time slot.\n"
            "- Fills slots sequentially across valid dates.\n"
            "- Sundays are automatically skipped.\n"
            "- Module Codes render without hyphens.\n"
            "- Strict FPDF Layout preserved exactly as Master."
        )

    if st.session_state.raw_df is not None:
        st.markdown("---")
        if st.button("🚀 Generate Scheduled PDF Timetable", type="primary", use_container_width=True):
            with st.spinner("Rendering Master PDF Layout…"):
                pdf_filename = "AutoScheduled_ReExam_Timetable.pdf"
                generate_pdf_timetable(
                    st.session_state.processed_tt,
                    pdf_filename,
                    declaration_date=decl_date
                )

                if os.path.exists(pdf_filename):
                    st.balloons()
                    with open(pdf_filename, "rb") as f:
                        st.download_button(
                            "📥 Download Timetable PDF",
                            f,
                            f"Scheduled_ReExam_Timetable_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                            "application/pdf",
                            use_container_width=True
                        )
                    try: os.remove(pdf_filename)
                    except: pass

if __name__ == "__main__":
    main()
