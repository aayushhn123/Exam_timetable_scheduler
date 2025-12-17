import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from fpdf import FPDF
import os
import io
import re
from PyPDF2 import PdfReader, PdfWriter

# ==========================================
# ⚙️ PAGE CONFIGURATION
# ==========================================
st.set_page_config(
    page_title="Timetable Generator (Clone)",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Check for dialog support
if hasattr(st, "dialog"):
    dialog_decorator = st.dialog
else:
    dialog_decorator = st.experimental_dialog

# ==========================================
# 🎨 UI & CSS OPTIMIZATION
# ==========================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    
    /* Main Header */
    .main-header {
        background: linear-gradient(135deg, #951C1C, #D92828);
        padding: 2rem;
        border-radius: 12px;
        color: white;
        text-align: center;
        margin-bottom: 2rem;
        box-shadow: 0 4px 20px rgba(149, 28, 28, 0.2);
    }
    .main-header h1 { color: white; margin: 0; font-weight: 700; font-size: 2rem; }
    .main-header p { margin-top: 0.5rem; opacity: 0.9; font-size: 1rem; }

    /* File Uploader Container */
    .upload-container {
        border: 2px dashed #e0e0e0;
        border-radius: 12px;
        padding: 2rem;
        text-align: center;
        background-color: #f8f9fa;
        transition: border-color 0.3s;
    }
    .upload-container:hover { border-color: #951C1C; }

    /* Statistics Buttons (Tiles) */
    div.row-widget.stButton > button {
        width: 100%;
        min-height: 100px;
        padding: 1rem;
        border-radius: 12px;
        background: white;
        color: #333;
        border: 1px solid #e0e0e0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        transition: all 0.2s;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        line-height: 1.4;
    }
    div.row-widget.stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 16px rgba(0,0,0,0.1);
        border-color: #951C1C;
        color: #951C1C;
    }
    div.row-widget.stButton > button p {
        font-size: 1.1rem !important;
        font-weight: 600 !important;
        margin: 0 !important;
    }
    
    /* Footer */
    .footer {
        text-align: center;
        padding: 2rem;
        color: #888;
        font-size: 0.9rem;
        border-top: 1px solid #eee;
        margin-top: 3rem;
    }
    
    /* Hide Streamlit Default Elements */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# ==========================================
# 🏫 COLLEGE CONFIGURATION
# ==========================================
# Dictionary mapping: "Display Name (with Icon)" -> "Clean Name (for PDF)"
COLLEGE_MAP = {
    "SVKM's NMIMS University": "SVKM's NMIMS University",
    "Mukesh Patel School of Technology Management & Engineering 🖥️": "Mukesh Patel School of Technology Management & Engineering",
    "School of Business Management 💼": "School of Business Management",
    "Pravin Dalal School of Entrepreneurship & Family Business Management 🚀": "Pravin Dalal School of Entrepreneurship & Family Business Management",
    "Anil Surendra Modi School of Commerce 📊": "Anil Surendra Modi School of Commerce",
    "School of Commerce 💰": "School of Commerce",
    "Kirit P. Mehta School of Law ⚖️": "Kirit P. Mehta School of Law",
    "School of Law 📜": "School of Law",
    "Shobhaben Pratapbhai Patel School of Pharmacy & Technology Management 💊": "Shobhaben Pratapbhai Patel School of Pharmacy & Technology Management",
    "School of Pharmacy & Technology Management 🧪": "School of Pharmacy & Technology Management",
    "Sunandan Divatia School of Science 🔬": "Sunandan Divatia School of Science",
    "School of Science 🧬": "School of Science",
    "Sarla Anil Modi School of Economics 📈": "Sarla Anil Modi School of Economics",
    "Balwant Sheth School of Architecture 🏛️": "Balwant Sheth School of Architecture",
    "School of Design 🎨": "School of Design",
    "Jyoti Dalal School of Liberal Arts 📚": "Jyoti Dalal School of Liberal Arts",
    "School of Performing Arts 🎭": "School of Performing Arts",
    "School of Hospitality Management 🏨": "School of Hospitality Management",
    "School of Mathematics, Applied Statistics & Analytics 📐": "School of Mathematics, Applied Statistics & Analytics",
    "School of Branding and Advertising 📢": "School of Branding and Advertising",
    "School of Agricultural Sciences & Technology 🌾": "School of Agricultural Sciences & Technology",
    "Centre of Distance and Online Education 💻": "Centre of Distance and Online Education",
    "School of Aviation ✈️": "School of Aviation"
}

BRANCH_FULL_FORM = {
    "B TECH": "BACHELOR OF TECHNOLOGY",
    "B TECH INTG": "BACHELOR OF TECHNOLOGY SIX YEAR INTEGRATED PROGRAM",
    "M TECH": "MASTER OF TECHNOLOGY",
    "MBA TECH": "MASTER OF BUSINESS ADMINISTRATION IN TECHNOLOGY MANAGEMENT",
    "MCA": "MASTER OF COMPUTER APPLICATIONS",
    "DIPLOMA": "DIPLOMA IN ENGINEERING"
}

LOGO_PATH = "logo.png"

# ==========================================
# 📊 DIALOGS (STATISTICS)
# ==========================================
@dialog_decorator("📚 Total Exams Breakdown")
def show_exams_breakdown(df):
    st.markdown(f"### Total Scheduled Exams: **{len(df)}**")
    tab1, tab2 = st.tabs(["📊 By Category", "🌿 By Branch"])
    with tab1:
        if 'OE' in df.columns:
            df['Category'] = df['OE'].apply(lambda x: 'Elective (OE)' if pd.notna(x) and str(x).strip() not in ['', 'nan', 'None'] else 'Core')
            cat_counts = df['Category'].value_counts().reset_index()
            cat_counts.columns = ['Category', 'Count']
            st.dataframe(cat_counts, use_container_width=True, hide_index=True)
    with tab2:
        if 'MainBranch' in df.columns:
            branch_counts = df['MainBranch'].value_counts().reset_index()
            branch_counts.columns = ['Branch', 'Exam Count']
            st.dataframe(branch_counts, use_container_width=True, hide_index=True)

@dialog_decorator("🎓 Semesters Breakdown")
def show_semesters_breakdown(df):
    unique_sems = sorted(df['Semester'].unique())
    st.write(f"**Active Semesters:** {', '.join(map(str, unique_sems))}")
    sem_counts = df['Semester'].value_counts().sort_index().reset_index()
    sem_counts.columns = ['Semester', 'Subject Count']
    st.dataframe(sem_counts, use_container_width=True, hide_index=True)

@dialog_decorator("🏫 Programs & Streams Breakdown")
def show_programs_streams_breakdown(df):
    programs = sorted(df['MainBranch'].dropna().unique())
    streams = sorted(df['SubBranch'].dropna().unique())
    c1, c2 = st.columns(2)
    c1.metric("Total Programs", len(programs))
    c2.metric("Total Streams", len(streams))
    st.markdown("---")
    tab1, tab2 = st.tabs(["📂 By Program", "💧 All Streams"])
    with tab1:
        for prog in programs:
            s_count = len(df[df['MainBranch'] == prog]['SubBranch'].dropna().unique())
            with st.expander(f"**{prog}** ({s_count} Streams)"):
                st.write(f"Streams in {prog}")
    with tab2:
        st.dataframe(pd.DataFrame(streams, columns=["Stream Name"]), use_container_width=True)

@dialog_decorator("📅 Schedule Span")
def show_span_breakdown(df):
    dates = pd.to_datetime(df['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
    if not dates.empty:
        start, end = min(dates), max(dates)
        st.metric("Total Span", f"{(end-start).days+1} Days")
        st.caption(f"{start.strftime('%d %b')} to {end.strftime('%d %b')}")
        daily = df['Exam Date'].value_counts().reset_index()
        daily.columns = ['Date', 'Count']
        st.bar_chart(daily.set_index('Date'))

# ==========================================
# 🛠️ HELPER FUNCTIONS
# ==========================================
def normalize_time(t_str):
    if not isinstance(t_str, str): return ""
    t_str = t_str.strip().upper()
    for i in range(1, 10):
        t_str = t_str.replace(f"0{i}:", f"{i}:")
    return t_str

def calculate_end_time(start_time, duration_hours):
    try:
        start_time = str(start_time).strip()
        if "AM" in start_time.upper() or "PM" in start_time.upper():
            start = datetime.strptime(start_time, "%I:%M %p")
        else:
            start = datetime.strptime(start_time, "%H:%M")
        end = start + timedelta(hours=float(duration_hours))
        return end.strftime("%I:%M %p").replace("AM", "AM").replace("PM", "PM")
    except:
        return f"{start_time}"

def int_to_roman(num):
    roman_values = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]
    result = ""
    for value, numeral in roman_values:
        while num >= value:
            result += numeral
            num -= value
    return result

# ==========================================
# 📥 PROCESSING LOGIC
# ==========================================
def process_verification_file(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file)
        
        # Check columns
        req_cols = ['Program', 'Stream', 'Current Session', 'Module Description', 'Exam Date', 'Configured Slot']
        if not all(col in df.columns for col in req_cols):
            st.error("❌ Invalid File Format. Missing required columns.")
            return None, None

        # Clean
        df = df[df['Exam Date'].notna() & (df['Exam Date'].astype(str) != 'Not Scheduled')].copy()
        
        # Map Columns
        df['MainBranch'] = df['Program'].astype(str).str.strip()
        df['SubBranch'] = df['Stream'].astype(str).str.strip()
        # Handle empty stream case
        df['SubBranch'] = df.apply(lambda x: x['MainBranch'] if x['SubBranch'] in ['nan', ''] else x['SubBranch'], axis=1)
        
        df['Subject'] = df['Module Description'].astype(str).str.strip()
        df['Configured Slot'] = df['Configured Slot'].fillna("").astype(str).str.strip()
        df['ModuleCode'] = df['Module Abbreviation'].astype(str).str.strip() if 'Module Abbreviation' in df.columns else ''
        
        # Duration
        if 'Exam Duration' in df.columns:
            df['Exam Duration'] = pd.to_numeric(df['Exam Duration'], errors='coerce').fillna(3.0)
        else:
            df['Exam Duration'] = 3.0
            
        # OE Check
        if 'Subject Type' in df.columns:
            df['OE'] = df['Subject Type'].apply(lambda x: 'OE' if str(x).strip().upper() == 'OE' else None)
        else:
            df['OE'] = None

        # Semester Parse
        def get_sem_int(val):
            s = str(val).upper()
            m = re.search(r'(\d+)', s)
            if m: return int(m.group(1))
            roman_map = {'I':1, 'II':2, 'III':3, 'IV':4, 'V':5, 'VI':6, 'VII':7, 'VIII':8, 'IX':9, 'X':10}
            for r, i in roman_map.items():
                if r in s: return i
            return 1
        
        df['Semester'] = df['Current Session'].apply(get_sem_int)
        
        # Date Format
        df['Exam Date'] = pd.to_datetime(df['Exam Date'], dayfirst=True, errors='coerce').dt.strftime('%d-%m-%Y')
        
        # Create Dictionary
        timetable = {}
        for sem in sorted(df['Semester'].unique()):
            timetable[sem] = df[df['Semester'] == sem].copy()
            
        return timetable, df

    except Exception as e:
        st.error(f"Error processing file: {e}")
        return None, None

# ==========================================
# 💾 EXCEL ENGINE (LOGIC & FORMATTING)
# ==========================================
def save_to_excel(semester_wise_timetable):
    time_slots_dict = {1: {"start": "10:00 AM", "end": "1:00 PM"}, 2: {"start": "2:00 PM", "end": "5:00 PM"}}
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        has_data = False
        for sem, df_sem in semester_wise_timetable.items():
            if df_sem.empty: continue
            
            # Header Time Logic
            slot_id = 1 if ((sem + 1) // 2) % 2 == 1 else 2
            p_cfg = time_slots_dict[slot_id]
            header_norm = normalize_time(f"{p_cfg['start']} - {p_cfg['end']}")
            
            for main_branch in df_sem['MainBranch'].unique():
                df_mb = df_sem[df_sem['MainBranch'] == main_branch].copy()
                roman_sem = int_to_roman(sem)
                sheet_name = f"{main_branch}_Sem_{roman_sem}"[:31]
                
                # Core vs Elec
                df_core = df_mb[df_mb['OE'].isna()].copy()
                df_elec = df_mb[df_mb['OE'].notna()].copy()
                
                # Process Core
                if not df_core.empty:
                    displays = []
                    for _, row in df_core.iterrows():
                        subj = row['Subject']
                        code = row['ModuleCode']
                        conf_slot = row['Configured Slot']
                        dur = row['Exam Duration']
                        
                        # Calc actual time
                        actual_time = conf_slot
                        if conf_slot and " - " in conf_slot:
                            try:
                                s = conf_slot.split(" - ")[0].strip()
                                e = calculate_end_time(s, dur)
                                actual_time = f"{s} - {e}"
                            except: pass
                        
                        # Compare
                        time_suffix = ""
                        if normalize_time(actual_time) != header_norm and actual_time:
                            time_suffix = f" [{actual_time}]"
                            
                        txt = f"{subj}"
                        if code and code != 'nan': txt += f" - ({code})"
                        txt += time_suffix
                        displays.append(txt)
                    
                    df_core['Display'] = displays
                    try:
                        pivot = df_core.pivot_table(index='Exam Date', columns='SubBranch', values='Display', aggfunc=lambda x: "\n".join(x)).fillna("---")
                        pivot.reset_index(inplace=True)
                        # Sort
                        pivot['DateObj'] = pd.to_datetime(pivot['Exam Date'], dayfirst=True, errors='coerce')
                        pivot = pivot.sort_values('DateObj').drop(columns=['DateObj'])
                        pivot.to_excel(writer, sheet_name=sheet_name, index=False)
                        has_data = True
                    except: pass
                
                # Process Electives
                if not df_elec.empty:
                    e_displays = []
                    for _, row in df_elec.iterrows():
                        subj = row['Subject']
                        code = row['ModuleCode']
                        conf_slot = row['Configured Slot']
                        dur = row['Exam Duration']
                        
                        actual_time = conf_slot
                        if conf_slot and " - " in conf_slot:
                            try:
                                s = conf_slot.split(" - ")[0].strip()
                                e = calculate_end_time(s, dur)
                                actual_time = f"{s} - {e}"
                            except: pass
                        
                        time_suffix = ""
                        if normalize_time(actual_time) != header_norm and actual_time:
                            time_suffix = f" [{actual_time}]"
                            
                        txt = f"{subj}"
                        if code and code != 'nan': txt += f" - ({code})"
                        txt += time_suffix
                        e_displays.append(txt)
                        
                    df_elec['Display'] = e_displays
                    try:
                        ep = df_elec.groupby(['OE', 'Exam Date']).agg({'Display': lambda x: "\n".join(set(x))}).reset_index()
                        ep.rename(columns={'OE': 'OE Type', 'Display': 'Subjects'}, inplace=True)
                        ep['DateObj'] = pd.to_datetime(ep['Exam Date'], dayfirst=True, errors='coerce')
                        ep = ep.sort_values('DateObj').drop(columns=['DateObj'])
                        ep.to_excel(writer, sheet_name=f"{main_branch}_Sem_{roman_sem}_Electives"[:31], index=False)
                        has_data = True
                    except: pass
                    
        if not has_data:
            pd.DataFrame({'Info': ['No valid data']}).to_excel(writer, sheet_name="Empty")
            
    output.seek(0)
    return output

# ==========================================
# 📄 PDF ENGINE (EXACT FORMAT)
# ==========================================
def wrap_text_pdf(pdf, text, col_width):
    lines = []
    current_line = ""
    words = re.split(r'(\s+)', str(text))
    for word in words:
        test_line = current_line + word
        if pdf.get_string_width(test_line) <= col_width:
            current_line = test_line
        else:
            if current_line: lines.append(current_line.strip())
            current_line = word
    if current_line: lines.append(current_line.strip())
    return lines

def print_row(pdf, row_data, col_widths, line_height=5, header=False):
    row_num = getattr(pdf, '_row_counter', 0)
    if header:
        pdf.set_font("Arial", 'B', 10)
        pdf.set_fill_color(149, 33, 28)
        pdf.set_text_color(255, 255, 255)
    else:
        pdf.set_font("Arial", '', 10)
        pdf.set_text_color(0, 0, 0)
        pdf.set_fill_color(240, 240, 240) if row_num % 2 == 1 else pdf.set_fill_color(255, 255, 255)
    
    wrapped = [wrap_text_pdf(pdf, txt, w-4) for txt, w in zip(row_data, col_widths)]
    max_lines = max(max((len(l) for l in wrapped)), 1)
    row_h = line_height * max_lines
    
    if pdf.get_y() + row_h > pdf.h - 25: return False # Page break needed
    
    x, y = pdf.get_x(), pdf.get_y()
    if header or row_num % 2 == 1: pdf.rect(x, y, sum(col_widths), row_h, 'F')
    
    for i, lines in enumerate(wrapped):
        cx = pdf.get_x()
        pad = (row_h - len(lines)*line_height)/2
        for j, ln in enumerate(lines):
            pdf.set_xy(cx+2, y+j*line_height+pad)
            pdf.cell(col_widths[i]-4, line_height, ln, 0, 0, 'L' if i==0 else 'C')
        pdf.set_xy(cx+col_widths[i], y)
        pdf.rect(cx, y, col_widths[i], row_h)
        
    pdf.set_xy(x, y+row_h)
    setattr(pdf, '_row_counter', row_num + 1)
    return True

def add_header(pdf, content, branches, time_slot, decl_date, clean_college_name):
    pdf.set_y(0)
    if decl_date:
        pdf.set_font("Arial", 'B', 10)
        pdf.set_text_color(0,0,0)
        pdf.set_xy(pdf.w-60, 10)
        pdf.cell(50, 10, f"Declaration Date: {decl_date.strftime('%d-%m-%Y')}", 0, 0, 'R')
        
    if os.path.exists(LOGO_PATH):
        pdf.image(LOGO_PATH, x=(pdf.w-45)/2, y=10, w=45)
        
    pdf.set_fill_color(149, 33, 28)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", 'B', 16 if len(clean_college_name)<60 else 12)
    pdf.rect(10, 30, pdf.w-20, 14, 'F')
    pdf.set_xy(10, 30)
    pdf.cell(pdf.w-20, 14, clean_college_name, 0, 1, 'C')
    
    pdf.set_text_color(0,0,0)
    pdf.set_font("Arial", 'B', 15)
    pdf.set_xy(10, 51)
    pdf.cell(pdf.w-20, 8, f"{content['main_branch_full']} - Semester {content['semester_roman']}", 0, 1, 'C')
    
    if time_slot:
        pdf.set_font("Arial", 'B', 13)
        pdf.set_xy(10, 59)
        pdf.cell(pdf.w-20, 6, f"Exam Time: {time_slot}", 0, 1, 'C')
        pdf.set_font("Arial", 'I', 10)
        pdf.set_xy(10, 65)
        pdf.cell(pdf.w-20, 6, "(Check the subject exam time)", 0, 1, 'C')
        
    pdf.set_font("Arial", '', 12)
    pdf.set_xy(10, 71)
    pdf.cell(pdf.w-20, 6, f"Branches: {', '.join(branches)}", 0, 1, 'C')
    pdf.set_y(85)

def add_footer(pdf):
    pdf.set_xy(10, pdf.h-25)
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 5, "Controller of Examinations", 0, 1, 'L')
    pdf.line(10, pdf.h-20, 60, pdf.h-20)
    pdf.set_font("Arial", '', 14)
    pg = f"{pdf.page_no()} of {{nb}}"
    pdf.set_xy(pdf.w-10-pdf.get_string_width(pg), pdf.h-13)
    pdf.cell(0, 5, pg, 0, 0)

def convert_excel_to_pdf(excel_path, pdf_path, decl_date, clean_college_name):
    pdf = FPDF('L', 'mm', (210, 500))
    pdf.set_auto_page_break(False)
    pdf.alias_nb_pages()
    
    t_slots = {1: "10:00 AM - 1:00 PM", 2: "2:00 PM - 5:00 PM"}
    
    try:
        xls = pd.read_excel(excel_path, sheet_name=None)
        
        for sheet, df in xls.items():
            if df.empty: continue
            
            # Header Config
            parts = sheet.split('_Sem_')
            prog = parts[0]
            sem_raw = parts[1].replace('_Electives', '')
            
            roman_map = {'I':1, 'II':2, 'III':3, 'IV':4, 'V':5, 'VI':6, 'VII':7, 'VIII':8, 'IX':9, 'X':10}
            sem_int = roman_map.get(sem_raw, 1)
            
            # Use logic: Sem 1,2 -> Slot 1
            slot_id = 1 if ((sem_int + 1)//2)%2 == 1 else 2
            time_slot = t_slots[slot_id]
            
            full_prog_name = BRANCH_FULL_FORM.get(prog, prog)
            header_content = {'main_branch_full': full_prog_name, 'semester_roman': sem_raw}
            
            # Columns
            fixed = ['Exam Date']
            others = [c for c in df.columns if c not in fixed and 'Unnamed' not in str(c)]
            
            # PDF Loop
            if '_Electives' in sheet:
                # Elective Layout
                cols = ['Exam Date', 'OE Type', 'Subjects']
                col_ws = [60, 40, pdf.w-120]
                
                pdf.add_page()
                add_footer(pdf)
                add_header(pdf, header_content, ['All Streams'], time_slot, decl_date, clean_college_name)
                print_row(pdf, cols, col_ws, 10, True)
                
                for i in range(len(df)):
                    row = [str(df.iloc[i].get(c, '')) for c in cols]
                    if not print_row(pdf, row, col_ws, 10):
                        pdf.add_page()
                        add_footer(pdf)
                        add_header(pdf, header_content, ['All Streams'], time_slot, decl_date, clean_college_name)
                        print_row(pdf, cols, col_ws, 10, True)
                        print_row(pdf, row, col_ws, 10)
            else:
                # Core Layout (Paginate Columns)
                per_page = 4
                for i in range(0, len(others), per_page):
                    chunk = others[i:i+per_page]
                    cols = fixed + chunk
                    
                    # Widths
                    w_date = 60
                    w_rem = pdf.w - 20 - w_date
                    w_col = w_rem / max(len(chunk), 1)
                    col_ws = [w_date] + [w_col]*len(chunk)
                    
                    pdf.add_page()
                    add_footer(pdf)
                    add_header(pdf, header_content, chunk, time_slot, decl_date, clean_college_name)
                    print_row(pdf, cols, col_ws, 10, True)
                    
                    sub_df = df[cols].dropna(how='all', subset=chunk)
                    for r_idx in range(len(sub_df)):
                        row_raw = []
                        # Format Date
                        d_val = str(sub_df.iloc[r_idx]['Exam Date'])
                        try: 
                            row_raw.append(pd.to_datetime(d_val, dayfirst=True).strftime("%A, %d %B, %Y"))
                        except: 
                            row_raw.append(d_val)
                            
                        for c in chunk:
                            row_raw.append(str(sub_df.iloc[r_idx][c]))
                            
                        if not print_row(pdf, row_raw, col_ws, 10):
                            pdf.add_page()
                            add_footer(pdf)
                            add_header(pdf, header_content, chunk, time_slot, decl_date, clean_college_name)
                            print_row(pdf, cols, col_ws, 10, True)
                            print_row(pdf, row_raw, col_ws, 10)
                            
    except Exception as e:
        st.error(f"PDF Gen Error: {e}")
        return False
        
    # INSTRUCTIONS PAGE
    try:
        pdf.add_page()
        add_footer(pdf)
        i_head = {'main_branch_full': 'EXAMINATION GUIDELINES', 'semester_roman': 'General'}
        add_header(pdf, i_head, ['All Candidates'], None, decl_date, clean_college_name)
        pdf.set_y(95)
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(0, 10, "INSTRUCTIONS TO CANDIDATES", 0, 1, 'C')
        pdf.ln(5)
        instrs = [
            "1. Candidates are required to be present at the examination center THIRTY MINUTES before the stipulated time.",
            "2. Candidates must produce their University Identity Card at the time of the examination.",
            "3. Candidates are not permitted to enter the examination hall after stipulated time.",
            "4. Candidates will not be permitted to leave the examination hall during the examination time.",
            "5. Candidates are forbidden from taking any unauthorized material inside the examination hall. Carrying the same will be treated as usage of unfair means."
        ]
        pdf.set_font("Arial", '', 12)
        for i in instrs:
            pdf.multi_cell(0, 8, i)
            pdf.ln(2)
    except: pass
    
    pdf.output(pdf_path)
    return True

# ==========================================
# 🚀 MAIN APP
# ==========================================
def main():
    st.markdown('<div class="main-header"><h1>📅 Timetable PDF Generator</h1><p>From Verification File (Clone Mode)</p></div>', unsafe_allow_html=True)
    
    # State Init
    if 'pdf_data' not in st.session_state: st.session_state.pdf_data = None
    if 'raw_df' not in st.session_state: st.session_state.raw_df = None
    if 'processed_tt' not in st.session_state: st.session_state.processed_tt = None

    # Sidebar
    with st.sidebar:
        st.header("⚙️ Configuration")
        # Selector shows Icons, Value maps to Clean Name
        selected_display = st.selectbox("College Name", list(COLLEGE_MAP.keys()))
        
        # Determine Clean Name
        if selected_display == "Custom...":
            clean_name = st.text_input("Enter Custom Name")
        else:
            clean_name = COLLEGE_MAP[selected_display]
            
        st.session_state.clean_college_name = clean_name
        
        st.markdown("---")
        decl_date = st.date_input("📆 Declaration Date", value=None)

    # Layout
    col1, col2 = st.columns([2, 1])
    with col1:
        st.markdown('<div class="upload-container"><h3>📁 Upload Verification File</h3><p>Drag and drop Excel file here</p></div>', unsafe_allow_html=True)
        uploaded = st.file_uploader("Upload Excel", type=['xlsx', 'xls'], label_visibility="collapsed")
        
        if uploaded:
            tt, df = process_verification_file(uploaded)
            if tt:
                st.session_state.processed_tt = tt
                st.session_state.raw_df = df
                st.markdown('<div style="background:#d4edda;padding:1rem;border-radius:8px;color:#155724;margin-top:1rem;text-align:center;">✅ File Loaded Successfully!</div>', unsafe_allow_html=True)
    
    with col2:
        st.info("ℹ️ **Clone Mode:** This tool skips scheduling algorithms. It uses the exact dates/slots from your Verification File but applies the standard PDF styling.")

    # Processed UI
    if st.session_state.raw_df is not None:
        df = st.session_state.raw_df
        
        st.markdown("### 📊 Statistics")
        
        # Stats Grid
        s1, s2, s3, s4 = st.columns(4)
        with s1: 
            if st.button(f"{len(df)}\nTOTAL EXAMS", key="b1"): show_exams_breakdown(df)
        with s2:
            if st.button(f"{df['Semester'].nunique()}\nSEMESTERS", key="b2"): show_semesters_breakdown(df)
        with s3:
            if st.button(f"{df['MainBranch'].nunique()} / {df['SubBranch'].nunique()}\nPROGS / STREAMS", key="b3"): show_programs_streams_breakdown(df)
        with s4:
            if st.button("📅\nSPAN INFO", key="b4"): show_span_breakdown(df)

        st.markdown("---")
        
        # Generate Action
        if st.button("🚀 Generate PDF Timetable", type="primary", use_container_width=True):
            with st.spinner("Generating PDF..."):
                # 1. Excel
                excel_io = save_to_excel(st.session_state.processed_tt)
                with open("temp_v.xlsx", "wb") as f: f.write(excel_io.getvalue())
                
                # 2. PDF
                if convert_excel_to_pdf("temp_v.xlsx", "Verification_Timetable.pdf", decl_date, st.session_state.clean_college_name):
                    st.balloons()
                    with open("Verification_Timetable.pdf", "rb") as f:
                        st.download_button("📥 Download PDF", f, "Timetable.pdf", "application/pdf", use_container_width=True)
                    
                    # Cleanup
                    try: os.remove("temp_v.xlsx"); os.remove("Verification_Timetable.pdf")
                    except: pass

    # Footer
    st.markdown('<div class="footer">Timetable Generator • Clone Mode • Standardized PDF Output</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
