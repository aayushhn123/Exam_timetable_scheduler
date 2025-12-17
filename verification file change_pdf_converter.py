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
    page_title="Timetable PDF Generator (Clone)",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Check for dialog support (Streamlit version compatibility)
if hasattr(st, "dialog"):
    dialog_decorator = st.dialog
else:
    dialog_decorator = st.experimental_dialog

# ==========================================
# 🎨 UI & CSS (EXACT MATCH)
# ==========================================
st.markdown("""
<style>
    /* Import Google Fonts */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    * { font-family: 'Inter', sans-serif; }
    
    .main-header {
        padding: 2rem;
        border-radius: 12px;
        margin-bottom: 2rem;
        background: linear-gradient(90deg, #951C1C, #C73E1D);
        color: white;
        text-align: center;
        box-shadow: 0 4px 12px rgba(149, 28, 28, 0.2);
    }
    .main-header h1 { color: white; margin: 0; font-size: 2.2rem; font-weight: 700; }
    .main-header p { color: rgba(255,255,255,0.9); margin-top: 0.5rem; font-size: 1.1rem; }

    /* Upload Section */
    .upload-section {
        background-color: #f8f9fa;
        border: 2px dashed #e9ecef;
        border-radius: 12px;
        padding: 2rem;
        text-align: center;
        margin-bottom: 1.5rem;
    }
    
    /* Stats Button Styling - The "Cards" */
    div.row-widget.stButton > button {
        width: 100%;
        padding: 1rem;
        height: auto;
        min-height: 120px;
        border-radius: 12px;
        border: 1px solid rgba(255,255,255,0.1);
        background: linear-gradient(135deg, #4a5db0 0%, #764ba2 100%);
        color: white;
        text-align: center;
        transition: transform 0.2s, box-shadow 0.2s;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
    }
    div.row-widget.stButton > button:hover {
        transform: translateY(-4px);
        box-shadow: 0 8px 16px rgba(118, 75, 162, 0.3);
        border-color: rgba(255,255,255,0.3);
        color: white !important;
    }
    /* Force text styling inside buttons */
    div.row-widget.stButton > button p {
        font-size: 1.8rem !important;
        font-weight: 700 !important;
        line-height: 1.2 !important;
        margin: 0 !important;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 📊 STATISTICS DIALOGS (EXACT REPLICA)
# ==========================================

@dialog_decorator("📚 Total Exams Breakdown")
def show_exams_breakdown(df):
    st.markdown(f"### Total Scheduled Exams: **{len(df)}**")
    tab1, tab2 = st.tabs(["📊 By Category", "🌿 By Branch"])
    with tab1:
        if 'OE' in df.columns:
            # Helper to categorize based on OE column presence/content
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
    st.dataframe(
        sem_counts,
        column_config={
            "Subject Count": st.column_config.ProgressColumn(
                "Subject Count", format="%d", min_value=0, max_value=int(sem_counts['Subject Count'].max()),
            ),
        },
        use_container_width=True,
        hide_index=True,
    )

@dialog_decorator("🏫 Programs & Streams Breakdown")
def show_programs_streams_breakdown(df):
    programs = sorted(df['MainBranch'].dropna().unique()) if 'MainBranch' in df.columns else []
    streams = sorted(df['SubBranch'].dropna().unique()) if 'SubBranch' in df.columns else []
    col1, col2 = st.columns(2)
    col1.metric("Total Programs", len(programs))
    col2.metric("Total Streams", len(streams))
    st.markdown("---")
    tab1, tab2 = st.tabs(["📂 Grouped by Program", "💧 All Streams List"])
    with tab1:
        if 'MainBranch' in df.columns and 'SubBranch' in df.columns:
            for prog in programs:
                prog_streams = df[df['MainBranch'] == prog]['SubBranch'].dropna().unique()
                prog_streams = [s for s in prog_streams if str(s).strip() != '']
                with st.expander(f"**{prog}** ({len(prog_streams)} Streams)"):
                    if len(prog_streams) > 0:
                        for s in sorted(prog_streams): st.caption(f"• {s}")
                    else: st.caption("No specific streams defined.")
    with tab2:
        search = st.text_input("🔍 Search Streams", placeholder="Type to search...")
        filtered = [s for s in streams if search.lower() in str(s).lower()]
        if filtered:
            sc1, sc2 = st.columns(2)
            for i, s in enumerate(filtered):
                if i % 2 == 0: sc1.caption(f"• {s}")
                else: sc2.caption(f"• {s}")

@dialog_decorator("📅 Schedule Span Details")
def show_span_breakdown(df):
    dates = pd.to_datetime(df['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
    if dates.empty:
        st.warning("No valid dates found.")
        return
    start, end = min(dates), max(dates)
    span = (end - start).days + 1
    col1, col2 = st.columns(2)
    col1.metric("Start Date", start.strftime('%d %b %Y'))
    col2.metric("End Date", end.strftime('%d %b %Y'))
    col3, col4 = st.columns(2)
    col3.metric("Total Days", span)
    col4.metric("Active Exam Days", len(dates.unique()))
    st.markdown("### 📈 Exams per Day")
    daily = df['Exam Date'].value_counts().reset_index()
    daily.columns = ['Date', 'Count']
    daily['DateObj'] = pd.to_datetime(daily['Date'], format="%d-%m-%Y")
    daily = daily.sort_values('DateObj')
    st.bar_chart(data=daily, x='Date', y='Count')

# ==========================================
# 🛠️ UTILITIES & CONFIG
# ==========================================

LOGO_PATH = "logo.png"

BRANCH_FULL_FORM = {
    "B TECH": "BACHELOR OF TECHNOLOGY",
    "B TECH INTG": "BACHELOR OF TECHNOLOGY SIX YEAR INTEGRATED PROGRAM",
    "M TECH": "MASTER OF TECHNOLOGY",
    "MBA TECH": "MASTER OF BUSINESS ADMINISTRATION IN TECHNOLOGY MANAGEMENT",
    "MCA": "MASTER OF COMPUTER APPLICATIONS",
    "DIPLOMA": "DIPLOMA IN ENGINEERING"
}

def normalize_time(t_str):
    """Standardizes time strings (e.g., '01:00 PM' -> '1:00 PM')"""
    if not isinstance(t_str, str): return ""
    t_str = t_str.strip().upper()
    # Remove leading zeros from hours (01: -> 1:, 09: -> 9:)
    for i in range(1, 10):
        t_str = t_str.replace(f"0{i}:", f"{i}:")
    return t_str

def calculate_end_time(start_time, duration_hours):
    """Calculate the end time given a start time string and duration float."""
    try:
        start_time = str(start_time).strip()
        # Parse start time
        if "AM" in start_time.upper() or "PM" in start_time.upper():
            start = datetime.strptime(start_time, "%I:%M %p")
        else:
            start = datetime.strptime(start_time, "%H:%M")
            
        duration = timedelta(hours=float(duration_hours))
        end = start + duration
        # Return in 12-hour format with AM/PM
        return end.strftime("%I:%M %p").replace("AM", "AM").replace("PM", "PM")
    except:
        # Fallback if parsing fails
        return f"{start_time} + {duration_hours}h"

# ==========================================
# 📥 PROCESS VERIFICATION FILE (INPUT)
# ==========================================
def process_verification_file(uploaded_file):
    """
    Reads the Verification Excel and standardizes it into a DataFrame 
    compatible with the PDF generation logic.
    """
    try:
        df = pd.read_excel(uploaded_file)
        
        # 1. Check Required Columns
        # Note: We look for the standard names used in Verification files
        req_cols = ['Program', 'Stream', 'Current Session', 'Module Description', 'Exam Date', 'Configured Slot']
        missing = [c for c in req_cols if c not in df.columns]
        if missing:
            st.error(f"❌ Missing columns in Verification File: {', '.join(missing)}")
            return None, None

        # 2. Basic Cleaning
        df = df[df['Exam Date'].notna() & (df['Exam Date'].astype(str) != 'Not Scheduled')]
        
        # 3. Map to Internal Standard Names
        df['MainBranch'] = df['Program'].astype(str).str.strip()
        df['SubBranch'] = df['Stream'].astype(str).str.strip()
        df['SubBranch'] = df.apply(lambda x: x['MainBranch'] if x['SubBranch'] in ['nan', ''] else x['SubBranch'], axis=1)
        
        df['Subject'] = df['Module Description'].astype(str).str.strip()
        df['ModuleCode'] = df['Module Abbreviation'].astype(str).str.strip() if 'Module Abbreviation' in df.columns else ''
        
        # 4. Handle Duration (Critical for formatting)
        if 'Exam Duration' in df.columns:
            df['Exam Duration'] = pd.to_numeric(df['Exam Duration'], errors='coerce').fillna(3.0)
        else:
            df['Exam Duration'] = 3.0
            
        # 5. Handle OE
        if 'Subject Type' in df.columns:
            df['OE'] = df['Subject Type'].apply(lambda x: 'OE' if str(x).strip().upper() == 'OE' else None)
        else:
            df['OE'] = None
            
        # 6. Parse Semester
        def get_sem_int(val):
            s = str(val).upper()
            import re
            m = re.search(r'(\d+)', s)
            if m: return int(m.group(1))
            roman_map = {'I':1, 'II':2, 'III':3, 'IV':4, 'V':5, 'VI':6, 'VII':7, 'VIII':8, 'IX':9, 'X':10}
            for r, i in roman_map.items():
                if r in s: return i
            return 1
        
        df['Semester'] = df['Current Session'].apply(get_sem_int)
        
        # 7. Format Date (Standardize to DD-MM-YYYY string for logic)
        df['Exam Date'] = pd.to_datetime(df['Exam Date'], dayfirst=True, errors='coerce').dt.strftime('%d-%m-%Y')
        
        # 8. Create Semester-wise Dictionary
        timetable = {}
        for sem in sorted(df['Semester'].unique()):
            timetable[sem] = df[df['Semester'] == sem].copy()
            
        return timetable, df

    except Exception as e:
        st.error(f"Error processing file: {e}")
        return None, None

# ==========================================
# 💾 EXCEL GENERATION (THE FORMATTING ENGINE)
# ==========================================
def save_to_excel(semester_wise_timetable):
    """
    Core Logic: Formats the specific subject strings, handles the comparison 
    between Actual Duration Time vs Header Time, and pivots the data.
    """
    time_slots_dict = {
        1: {"start": "10:00 AM", "end": "1:00 PM"},
        2: {"start": "2:00 PM", "end": "5:00 PM"}
    }
    
    def int_to_roman(num):
        roman_values = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]
        result = ""
        for value, numeral in roman_values:
            while num >= value:
                result += numeral
                num -= value
        return result

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        sheets_written = False
        
        for sem, df_sem in semester_wise_timetable.items():
            if df_sem.empty: continue
            
            # --- HEADER TIME LOGIC ---
            # Sem 1,2 -> Slot 1 | Sem 3,4 -> Slot 2
            slot_indicator = ((sem + 1) // 2) % 2
            primary_slot_num = 1 if slot_indicator == 1 else 2
            
            p_cfg = time_slots_dict.get(primary_slot_num, time_slots_dict[1])
            header_std_time = f"{p_cfg['start']} - {p_cfg['end']}"
            header_norm = normalize_time(header_std_time)
            
            for main_branch in df_sem['MainBranch'].unique():
                df_mb = df_sem[df_sem['MainBranch'] == main_branch].copy()
                if df_mb.empty: continue
                
                roman_sem = int_to_roman(sem)
                sheet_name = f"{main_branch}_Sem_{roman_sem}"[:31]
                
                # Split Core vs Elective
                df_core = df_mb[df_mb['OE'].isna()].copy()
                df_elec = df_mb[df_mb['OE'].notna()].copy()
                
                # --- PROCESS CORE ---
                if not df_core.empty:
                    displays = []
                    for _, row in df_core.iterrows():
                        subj = str(row['Subject'])
                        code = str(row['ModuleCode'])
                        
                        # 1. Get Assigned Slot (Start)
                        assigned_slot = str(row.get('Configured Slot', '')).strip()
                        
                        # 2. Get Duration
                        duration = float(row.get('Exam Duration', 3.0))
                        
                        # 3. Calculate Actual Time (Start - End)
                        actual_time_str = assigned_slot
                        if assigned_slot and " - " in assigned_slot:
                            try:
                                start_t = assigned_slot.split(" - ")[0].strip()
                                end_t = calculate_end_time(start_t, duration)
                                actual_time_str = f"{start_t} - {end_t}"
                            except: pass
                        
                        # 4. Compare vs Header
                        # If mismatch, we show the time. If match, we show nothing.
                        time_suffix = ""
                        if normalize_time(actual_time_str) != header_norm and actual_time_str:
                            time_suffix = f" [{actual_time_str}]"
                            
                        # 5. Build String: "Subject - (Code) [Time]"
                        full_str = subj
                        if code and code.lower() != 'nan':
                            full_str += f" - ({code})"
                        full_str += time_suffix
                        displays.append(full_str)
                    
                    df_core['Display'] = displays
                    
                    # Pivot: Date x SubBranch
                    try:
                        pivot = df_core.pivot_table(
                            index='Exam Date', 
                            columns='SubBranch', 
                            values='Display', 
                            aggfunc=lambda x: "\n".join(x)
                        ).fillna("---")
                        
                        # Sort by date
                        pivot.reset_index(inplace=True)
                        pivot['DateObj'] = pd.to_datetime(pivot['Exam Date'], dayfirst=True, errors='coerce')
                        pivot = pivot.sort_values('DateObj').drop(columns=['DateObj'])
                        
                        pivot.to_excel(writer, sheet_name=sheet_name, index=False)
                        sheets_written = True
                    except Exception as e:
                        print(f"Pivot error: {e}")

                # --- PROCESS ELECTIVES ---
                if not df_elec.empty:
                    # Same logic for electives
                    e_displays = []
                    for _, row in df_elec.iterrows():
                        subj = str(row['Subject'])
                        code = str(row['ModuleCode'])
                        assigned_slot = str(row.get('Configured Slot', '')).strip()
                        duration = float(row.get('Exam Duration', 3.0))
                        
                        actual_time_str = assigned_slot
                        if assigned_slot and " - " in assigned_slot:
                            try:
                                start_t = assigned_slot.split(" - ")[0].strip()
                                end_t = calculate_end_time(start_t, duration)
                                actual_time_str = f"{start_t} - {end_t}"
                            except: pass
                            
                        time_suffix = ""
                        if normalize_time(actual_time_str) != header_norm and actual_time_str:
                            time_suffix = f" [{actual_time_str}]"
                            
                        full_str = subj
                        if code and code.lower() != 'nan':
                            full_str += f" - ({code})"
                        full_str += time_suffix
                        e_displays.append(full_str)
                        
                    df_elec['Display'] = e_displays
                    
                    try:
                        # Group electives by date & Type
                        e_pivot = df_elec.groupby(['OE', 'Exam Date']).agg({
                            'Display': lambda x: "\n".join(sorted(set(x)))
                        }).reset_index()
                        
                        e_pivot.rename(columns={'OE': 'OE Type', 'Display': 'Subjects'}, inplace=True)
                        
                        # Sort
                        e_pivot['DateObj'] = pd.to_datetime(e_pivot['Exam Date'], dayfirst=True, errors='coerce')
                        e_pivot = e_pivot.sort_values('DateObj').drop(columns=['DateObj'])
                        
                        es_name = f"{main_branch}_Sem_{roman_sem}_Electives"[:31]
                        e_pivot.to_excel(writer, sheet_name=es_name, index=False)
                        sheets_written = True
                    except: pass
        
        if not sheets_written:
            pd.DataFrame({'Info': ['No data found']}).to_excel(writer, sheet_name="Empty")
            
    writer.close()
    output.seek(0)
    return output

# ==========================================
# 📄 PDF GENERATION (EXACT STYLE)
# ==========================================
# Helper for PDF
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
    
    # Calculate height
    wrapped_cells = []
    max_lines = 0
    for i, txt in enumerate(row_data):
        lines = wrap_text_pdf(pdf, txt, col_widths[i] - 4)
        wrapped_cells.append(lines)
        max_lines = max(max_lines, len(lines))
    
    row_h = line_height * max_lines
    x, y = pdf.get_x(), pdf.get_y()
    
    # Check page break
    if y + row_h > pdf.h - 25: return False
    
    if header or row_num % 2 == 1:
        pdf.rect(x, y, sum(col_widths), row_h, 'F')
        
    for i, lines in enumerate(wrapped_cells):
        cx = pdf.get_x()
        pad = (row_h - len(lines)*line_height)/2 if len(lines) < max_lines else 0
        for j, ln in enumerate(lines):
            pdf.set_xy(cx + 2, y + j*line_height + pad)
            pdf.cell(col_widths[i]-4, line_height, ln, 0, 0, 'L' if i==0 else 'C')
        pdf.rect(cx, y, col_widths[i], row_h)
        pdf.set_xy(cx + col_widths[i], y)
        
    pdf.set_xy(x, y + row_h)
    setattr(pdf, '_row_counter', row_num + 1)
    return True

def add_header(pdf, content, branches, time_slot, declaration_date):
    pdf.set_y(0)
    # Decl Date
    if declaration_date:
        pdf.set_font("Arial", 'B', 10)
        pdf.set_text_color(0,0,0)
        pdf.set_xy(pdf.w - 60, 10)
        pdf.cell(50, 10, f"Declaration Date: {declaration_date.strftime('%d-%m-%Y')}", 0, 0, 'R')
        
    # Logo
    if os.path.exists(LOGO_PATH):
        pdf.image(LOGO_PATH, x=(pdf.w-45)/2, y=10, w=45)
        
    # College Name
    pdf.set_fill_color(149, 33, 28)
    pdf.set_text_color(255, 255, 255)
    c_name = st.session_state.get('selected_college', "SVKM's NMIMS University")
    pdf.set_font("Arial", 'B', 16 if len(c_name)<60 else 12)
    pdf.rect(10, 30, pdf.w-20, 14, 'F')
    pdf.set_xy(10, 30)
    pdf.cell(pdf.w-20, 14, c_name, 0, 1, 'C')
    
    # Semester Info
    pdf.set_text_color(0,0,0)
    pdf.set_font("Arial", 'B', 15)
    pdf.set_xy(10, 51)
    pdf.cell(pdf.w-20, 8, f"{content['main_branch_full']} - Semester {content['semester_roman']}", 0, 1, 'C')
    
    # Time & Branch
    if time_slot:
        pdf.set_font("Arial", 'B', 13)
        pdf.cell(0, 6, f"Exam Time: {time_slot}", 0, 1, 'C')
        pdf.set_font("Arial", 'I', 10)
        pdf.cell(0, 6, "(Check the subject exam time)", 0, 1, 'C')
    
    pdf.set_font("Arial", '', 12)
    pdf.set_xy(10, 71)
    pdf.cell(pdf.w-20, 6, f"Branches: {', '.join(branches)}", 0, 1, 'C')
    pdf.set_y(85)

def add_footer(pdf):
    pdf.set_xy(10, pdf.h - 25)
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 5, "Controller of Examinations", 0, 1, 'L')
    pdf.line(10, pdf.h - 20, 60, pdf.h - 20)
    
    pdf.set_font("Arial", '', 14)
    pg = f"{pdf.page_no()} of {{nb}}"
    pdf.set_xy(pdf.w - 10 - pdf.get_string_width(pg), pdf.h - 13)
    pdf.cell(0, 5, pg, 0, 0)

def convert_excel_to_pdf(excel_path, pdf_path, declaration_date):
    pdf = FPDF('L', 'mm', (210, 500))
    pdf.set_auto_page_break(False)
    pdf.alias_nb_pages()
    
    # Header Time Config
    t_slots = {1: "10:00 AM - 1:00 PM", 2: "2:00 PM - 5:00 PM"}
    
    try:
        xls = pd.read_excel(excel_path, sheet_name=None)
        
        for sheet, df in xls.items():
            if df.empty: continue
            
            # Parse Header Info
            try:
                parts = sheet.split('_Sem_')
                prog = parts[0]
                sem_raw = parts[1].replace('_Electives', '')
                
                # Roman Logic
                roman_map = {'I':1, 'II':2, 'III':3, 'IV':4, 'V':5, 'VI':6, 'VII':7, 'VIII':8, 'IX':9, 'X':10}
                sem_int = roman_map.get(sem_raw, 1)
                
                # Header Time
                slot_id = 1 if ((sem_int + 1)//2)%2 == 1 else 2
                time_slot = t_slots[slot_id]
                
                header_data = {'main_branch_full': BRANCH_FULL_FORM.get(prog, prog), 'semester_roman': sem_raw}
                
                # Print Columns
                fixed = ['Exam Date']
                others = [c for c in df.columns if c not in fixed and 'Unnamed' not in str(c) and c != 'OE Type' and c != 'Subjects']
                
                if '_Electives' in sheet:
                    # Elective Sheet Layout
                    cols = ['Exam Date', 'OE Type', 'Subjects']
                    col_ws = [60, 40, pdf.w - 120]
                    pdf.add_page()
                    add_footer(pdf)
                    add_header(pdf, header_data, ['All Streams'], time_slot, declaration_date)
                    print_row(pdf, cols, col_ws, 10, True)
                    
                    for i in range(len(df)):
                        row = [str(df.iloc[i].get(c, '')) for c in cols]
                        if not print_row(pdf, row, col_ws, 10):
                            pdf.add_page()
                            add_footer(pdf)
                            add_header(pdf, header_data, ['All Streams'], time_slot, declaration_date)
                            print_row(pdf, cols, col_ws, 10, True)
                            print_row(pdf, row, col_ws, 10)
                            
                else:
                    # Core Sheet Layout (Paginate Columns)
                    per_page = 4
                    for i in range(0, len(others), per_page):
                        chunk = others[i:i+per_page]
                        cols = fixed + chunk
                        
                        # Dynamic Widths
                        w_date = 60
                        w_rem = pdf.w - 20 - w_date
                        w_col = w_rem / len(chunk)
                        col_ws = [w_date] + [w_col]*len(chunk)
                        
                        pdf.add_page()
                        add_footer(pdf)
                        add_header(pdf, header_data, chunk, time_slot, declaration_date)
                        print_row(pdf, cols, col_ws, 10, True)
                        
                        # Rows
                        sub_df = df[cols].dropna(how='all', subset=chunk)
                        for r_idx in range(len(sub_df)):
                            row_vals = []
                            # Format Date specifically for PDF
                            d_val = str(sub_df.iloc[r_idx]['Exam Date'])
                            try:
                                d_obj = pd.to_datetime(d_val, dayfirst=True)
                                row_vals.append(d_obj.strftime("%A, %d %B, %Y"))
                            except:
                                row_vals.append(d_val)
                                
                            for c in chunk:
                                row_vals.append(str(sub_df.iloc[r_idx][c]))
                                
                            if not print_row(pdf, row_vals, col_ws, 10):
                                pdf.add_page()
                                add_footer(pdf)
                                add_header(pdf, header_data, chunk, time_slot, declaration_date)
                                print_row(pdf, cols, col_ws, 10, True)
                                print_row(pdf, row_vals, col_ws, 10)

            except Exception as e:
                print(f"Sheet error: {e}")
                continue
                
        # INSTRUCTIONS PAGE
        pdf.add_page()
        add_footer(pdf)
        i_head = {'main_branch_full': 'EXAMINATION GUIDELINES', 'semester_roman': 'General'}
        add_header(pdf, i_head, ['All Candidates'], None, declaration_date)
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

        pdf.output(pdf_path)
        return True
    except Exception as e:
        st.error(f"PDF Error: {e}")
        return False

# ==========================================
# 🚀 MAIN APP LOGIC
# ==========================================
# ==========================================
# 🚀 MAIN APP LOGIC
# ==========================================
def main():
    st.markdown('<div class="main-header"><h1>📅 Timetable PDF Generator</h1><p>From Verification File (Clone Mode)</p></div>', unsafe_allow_html=True)
    
    if 'pdf_data' not in st.session_state: st.session_state.pdf_data = None
    if 'raw_df' not in st.session_state: st.session_state.raw_df = None
    if 'processed_tt' not in st.session_state: st.session_state.processed_tt = None
    if 'selected_college' not in st.session_state: st.session_state.selected_college = "SVKM's NMIMS University"

    with st.sidebar:
        st.header("⚙️ Settings")
        
        # --- NEW: UPDATED COLLEGE LIST ---
        college_options = [
            "SVKM's NMIMS University",
            "Mukesh Patel School of Technology Management & Engineering",
            "School of Business Management",
            "Pravin Dalal School of Entrepreneurship & Family Business Management",
            "Anil Surendra Modi School of Commerce",
            "School of Commerce",
            "Kirit P. Mehta School of Law",
            "School of Law",
            "Shobhaben Pratapbhai Patel School of Pharmacy & Technology Management",
            "School of Pharmacy & Technology Management",
            "Sunandan Divatia School of Science",
            "School of Science",
            "Sarla Anil Modi School of Economics",
            "Balwant Sheth School of Architecture",
            "School of Design",
            "Jyoti Dalal School of Liberal Arts",
            "School of Performing Arts",
            "School of Hospitality Management",
            "School of Mathematics, Applied Statistics & Analytics",
            "School of Branding and Advertising",
            "School of Agricultural Sciences & Technology",
            "Centre of Distance and Online Education",
            "School of Aviation",
            "Custom..."
        ]
        
        clg = st.selectbox("College Name", college_options)
        
        if clg == "Custom...":
            st.session_state.selected_college = st.text_input("Enter Custom Name")
        else:
            st.session_state.selected_college = clg
            
        st.markdown("---")
        decl_date = st.date_input("📆 Declaration Date", value=None)

    col1, col2 = st.columns([2, 1])
    with col1:
        st.markdown('<div class="upload-section"><h3>📁 Upload Verification File</h3></div>', unsafe_allow_html=True)
        uploaded = st.file_uploader("Upload Excel", type=['xlsx', 'xls'])
        
        if uploaded:
            tt, df = process_verification_file(uploaded)
            if tt:
                st.session_state.processed_tt = tt
                st.session_state.raw_df = df
                st.success("✅ File Loaded Successfully!")
    
    with col2:
        st.info("ℹ️ This is a Clone of the main generator. It skips scheduling and uses the dates in your file, but applies the exact same PDF formatting.")

    if st.session_state.raw_df is not None:
        df = st.session_state.raw_df
        # Stats
        st.markdown("### 📊 Statistics")
        b1, b2, b3, b4 = st.columns(4)
        with b1: 
            if st.button(f"{len(df)}\nTOTAL EXAMS", key="s1"): show_exams_breakdown(df)
        with b2:
            if st.button(f"{df['Semester'].nunique()}\nSEMESTERS", key="s2"): show_semesters_breakdown(df)
        with b3:
            if st.button(f"{df['MainBranch'].nunique()} / {df['SubBranch'].nunique()}\nPROGS / STREAMS", key="s3"): show_programs_streams_breakdown(df)
        with b4:
            if st.button("📅\nSPAN INFO", key="s4"): show_span_breakdown(df)

        st.markdown("---")
        if st.button("🚀 Generate PDF Timetable", type="primary", use_container_width=True):
            with st.spinner("Generating PDF..."):
                # 1. Format Data to Excel
                excel_io = save_to_excel(st.session_state.processed_tt)
                with open("temp_v.xlsx", "wb") as f: f.write(excel_io.getvalue())
                
                # 2. Convert to PDF
                if convert_excel_to_pdf("temp_v.xlsx", "Verification_Timetable.pdf", decl_date):
                    st.success("🎉 PDF Generated Successfully!")
                    with open("Verification_Timetable.pdf", "rb") as f:
                        st.download_button("📥 Download PDF", f, "Timetable.pdf", "application/pdf", use_container_width=True)
                    try: os.remove("temp_v.xlsx"); os.remove("Verification_Timetable.pdf")
                    except: pass

if __name__ == "__main__":
    main()
