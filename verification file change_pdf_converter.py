import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from fpdf import FPDF
import os
import io
import re
from PyPDF2 import PdfReader, PdfWriter

# Set page configuration
st.set_page_config(
    page_title="Timetable PDF Generator (From Verification)",
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
# 📊 STATISTICS BREAKDOWN DIALOGS
# ==========================================

@dialog_decorator("📚 Total Exams Breakdown")
def show_exams_breakdown(df):
    st.markdown(f"### Total Scheduled Exams: **{len(df)}**")
    
    tab1, tab2 = st.tabs(["📊 By Category", "🌿 By Branch"])
    
    with tab1:
        if 'OE' in df.columns:
            # Create a category column for display
            df['Category'] = df['OE'].apply(lambda x: 'Elective (OE)' if pd.notna(x) and str(x).strip() != '' else 'Core')
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
                "Subject Count",
                format="%d",
                min_value=0,
                max_value=int(sem_counts['Subject Count'].max()),
            ),
        },
        use_container_width=True,
        hide_index=True,
    )

@dialog_decorator("🏫 Programs & Streams Breakdown")
def show_programs_streams_breakdown(df):
    # Get lists of programs and streams
    programs = sorted(df['MainBranch'].dropna().unique()) if 'MainBranch' in df.columns else []
    streams = sorted(df['SubBranch'].dropna().unique()) if 'SubBranch' in df.columns else []
    
    # Summary Metrics
    col1, col2 = st.columns(2)
    col1.metric("Total Programs", len(programs))
    col2.metric("Total Streams", len(streams))
    
    st.markdown("---")
    
    # Tabs for detailed view
    tab1, tab2 = st.tabs(["📂 Grouped by Program", "💧 All Streams List"])
    
    with tab1:
        if 'MainBranch' in df.columns and 'SubBranch' in df.columns:
            for prog in programs:
                # Find streams belonging to this program
                prog_streams = df[df['MainBranch'] == prog]['SubBranch'].dropna().unique()
                prog_streams = [s for s in prog_streams if str(s).strip() != '']
                
                with st.expander(f"**{prog}** ({len(prog_streams)} Streams)"):
                    if len(prog_streams) > 0:
                        for s in sorted(prog_streams):
                            st.caption(f"• {s}")
                    else:
                        st.caption("No specific streams defined (General).")
    
    with tab2:
        search = st.text_input("🔍 Search Streams", placeholder="Type stream name...")
        filtered = [s for s in streams if search.lower() in str(s).lower()]
        
        if filtered:
            sc1, sc2 = st.columns(2)
            for i, s in enumerate(filtered):
                if i % 2 == 0:
                    sc1.caption(f"• {s}")
                else:
                    sc2.caption(f"• {s}")
        else:
            st.warning("No streams match your search.")

@dialog_decorator("📅 Schedule Span Details")
def show_span_breakdown(df):
    dates = pd.to_datetime(df['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
    if dates.empty:
        st.warning("No valid dates found in schedule.")
        return

    start = min(dates)
    end = max(dates)
    span = (end - start).days + 1
    active_days = len(dates.unique())
    
    col1, col2 = st.columns(2)
    col1.metric("Start Date", start.strftime('%d %b %Y'))
    col2.metric("End Date", end.strftime('%d %b %Y'))
    
    col3, col4 = st.columns(2)
    col3.metric("Total Days Span", span)
    col4.metric("Active Exam Days", active_days)
    
    st.markdown("### 📈 Exams per Day")
    daily = df['Exam Date'].value_counts().reset_index()
    daily.columns = ['Date', 'Count']
    daily['DateObj'] = pd.to_datetime(daily['Date'], format="%d-%m-%Y")
    daily = daily.sort_values('DateObj')
    
    st.bar_chart(data=daily, x='Date', y='Count')

# ==========================================
# CONSTANTS & CONFIG
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

# ==========================================
# 🛠️ HELPER FUNCTIONS
# ==========================================

def calculate_end_time(start_time, duration_hours):
    """Calculate the end time given a start time and duration in hours."""
    try:
        start_time = str(start_time).strip()
        if "AM" in start_time.upper() or "PM" in start_time.upper():
            start = datetime.strptime(start_time, "%I:%M %p")
        else:
            start = datetime.strptime(start_time, "%H:%M")
        
        duration = timedelta(hours=float(duration_hours))
        end = start + duration
        return end.strftime("%I:%M %p").replace("AM", "AM").replace("PM", "PM")
    except Exception:
        return f"{start_time} + {duration_hours}h"

def wrap_text(pdf, text, col_width):
    # Simple wrapper used in PDF generation
    lines = []
    current_line = ""
    words = re.split(r'(\s+)', str(text))
    for word in words:
        test_line = current_line + word
        if pdf.get_string_width(test_line) <= col_width:
            current_line = test_line
        else:
            if current_line:
                lines.append(current_line.strip())
            current_line = word
    if current_line:
        lines.append(current_line.strip())
    return lines

def print_row_custom(pdf, row_data, col_widths, line_height=5, header=False):
    cell_padding = 2
    header_bg_color = (149, 33, 28)
    header_text_color = (255, 255, 255)
    alt_row_color = (240, 240, 240)

    row_number = getattr(pdf, '_row_counter', 0)
    if header:
        pdf.set_font("Arial", 'B', 10)
        pdf.set_text_color(*header_text_color)
        pdf.set_fill_color(*header_bg_color)
    else:
        pdf.set_font("Arial", size=10)
        pdf.set_text_color(0, 0, 0)
        pdf.set_fill_color(*alt_row_color if row_number % 2 == 1 else (255, 255, 255))

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
    
    # Check for page break before printing row
    if y0 + row_h > pdf.h - 25: # 25 is footer height
        return False # Signal to add new page

    if header or row_number % 2 == 1:
        pdf.rect(x0, y0, sum(col_widths), row_h, 'F')

    for i, lines in enumerate(wrapped_cells):
        cx = pdf.get_x()
        pad_v = (row_h - len(lines) * line_height) / 2 if len(lines) < max_lines else 0
        for j, ln in enumerate(lines):
            pdf.set_xy(cx + cell_padding, y0 + j * line_height + pad_v)
            pdf.cell(col_widths[i] - 2 * cell_padding, line_height, ln, border=0, align='L' if i == 0 else 'C')
        pdf.rect(cx, y0, col_widths[i], row_h)
        pdf.set_xy(cx + col_widths[i], y0)

    setattr(pdf, '_row_counter', row_number + 1)
    pdf.set_xy(x0, y0 + row_h)
    return True

def add_footer_with_page_number(pdf, footer_height):
    pdf.set_xy(10, pdf.h - footer_height)
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 5, "Controller of Examinations", 0, 1, 'L')
    pdf.line(10, pdf.h - footer_height + 5, 60, pdf.h - footer_height + 5)
    
    pdf.set_font("Arial", size=14)
    pdf.set_text_color(0, 0, 0)
    page_text = f"{pdf.page_no()} of {{nb}}"
    text_width = pdf.get_string_width(page_text.replace("{nb}", "99"))
    pdf.set_xy(pdf.w - 10 - text_width, pdf.h - footer_height + 12)
    pdf.cell(text_width, 5, page_text, 0, 0, 'R')

def add_header_to_page(pdf, logo_x, logo_width, header_content, Programs, time_slot=None, actual_time_slots=None, declaration_date=None):
    pdf.set_y(0)
    
    # NEW: Print Declaration Date at Top Right
    if declaration_date:
        pdf.set_font("Arial", 'B', 10)
        pdf.set_text_color(0, 0, 0)
        decl_str = f"Declaration Date: {declaration_date.strftime('%d-%m-%Y')}"
        pdf.set_xy(pdf.w - 60, 10)
        pdf.cell(50, 10, decl_str, 0, 0, 'R')

    if os.path.exists(LOGO_PATH):
        pdf.image(LOGO_PATH, x=logo_x, y=10, w=logo_width)
    
    pdf.set_fill_color(149, 33, 28)
    pdf.set_text_color(255, 255, 255)
    
    college_name = st.session_state.get('selected_college', 'SVKM\'s NMIMS University')
    
    if len(college_name) > 80:
        pdf.set_font("Arial", 'B', 12)
    elif len(college_name) > 60:
        pdf.set_font("Arial", 'B', 14)
    else:
        pdf.set_font("Arial", 'B', 16)
    
    pdf.rect(10, 30, pdf.w - 20, 14, 'F')
    pdf.set_xy(10, 30)
    pdf.cell(pdf.w - 20, 14, college_name, 0, 1, 'C')
    
    pdf.set_font("Arial", 'B', 15)
    pdf.set_text_color(0, 0, 0)
    pdf.set_xy(10, 51)
    pdf.cell(pdf.w - 20, 8, f"{header_content['main_branch_full']} - Semester {header_content['semester_roman']}", 0, 1, 'C')
    
    # SIMPLE TIME SLOT DISPLAY (MATCHING YOUR REQUIREMENT)
    if time_slot:
        pdf.set_font("Arial", 'B', 13)
        pdf.set_xy(10, 59)
        pdf.cell(pdf.w - 20, 6, f"Exam Time: {time_slot}", 0, 1, 'C')
        
        pdf.set_font("Arial", 'I', 10)
        pdf.set_xy(10, 65)
        pdf.cell(pdf.w - 20, 6, "(Check the subject exam time)", 0, 1, 'C')
        
        pdf.set_font("Arial", '', 12)
        pdf.set_xy(10, 71)
        pdf.cell(pdf.w - 20, 6, f"Programs: {', '.join(Programs)}", 0, 1, 'C')
        pdf.set_y(85)
    else:
        pdf.set_font("Arial", '', 12)
        pdf.set_xy(10, 65)
        pdf.cell(pdf.w - 20, 6, f"Programs: {', '.join(Programs)}", 0, 1, 'C')
        pdf.set_y(75)

def print_table_custom(pdf, df, columns, col_widths, line_height=5, header_content=None, Programs=None, time_slot=None, actual_time_slots=None, declaration_date=None):
    if df.empty: return
    setattr(pdf, '_row_counter', 0)
    
    footer_height = 25
    add_footer_with_page_number(pdf, footer_height)
    
    logo_width = 45
    logo_x = (pdf.w - logo_width) / 2
    add_header_to_page(pdf, logo_x, logo_width, header_content, Programs, time_slot, actual_time_slots, declaration_date)
    
    # Print header row
    print_row_custom(pdf, columns, col_widths, line_height=line_height, header=True)
    
    for idx in range(len(df)):
        row = [str(df.iloc[idx][c]) if pd.notna(df.iloc[idx][c]) else "" for c in columns]
        if not any(cell.strip() for cell in row): continue
        
        # Try printing row. If returns False (page full), add new page.
        if not print_row_custom(pdf, row, col_widths, line_height=line_height, header=False):
            pdf.add_page()
            add_footer_with_page_number(pdf, footer_height)
            add_header_to_page(pdf, logo_x, logo_width, header_content, Programs, time_slot, actual_time_slots, declaration_date)
            print_row_custom(pdf, columns, col_widths, line_height=line_height, header=True)
            print_row_custom(pdf, row, col_widths, line_height=line_height, header=False)

# ==========================================
# 📄 PDF GENERATION LOGIC
# ==========================================

def convert_excel_to_pdf(excel_path, pdf_path, sub_branch_cols_per_page=4, declaration_date=None):
    """
    Consumes the TEMP Excel file created from the verification data and generates the Final PDF.
    Contains the exact formatting, Declaration Date logic, and Instructions Page.
    """
    pdf = FPDF(orientation='L', unit='mm', format=(210, 500)) # Wide format to support many columns
    pdf.set_auto_page_break(auto=False, margin=15)
    pdf.alias_nb_pages()
    
    # Time slots configuration (for Header Logic)
    time_slots_dict = {
        1: {"start": "10:00 AM", "end": "1:00 PM"},
        2: {"start": "2:00 PM", "end": "5:00 PM"}
    }
    
    try:
        df_dict = pd.read_excel(excel_path, sheet_name=None)
    except Exception as e:
        st.error(f"Error reading Excel for PDF: {e}")
        return

    def int_to_roman(num):
        roman_values = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]
        result = ""
        for value, numeral in roman_values:
            while num >= value:
                result += numeral
                num -= value
        return result

    def get_header_time_for_semester(semester_num):
        try:
            sem_int = int(semester_num)
            slot_indicator = ((sem_int + 1) // 2) % 2
            slot_num = 1 if slot_indicator == 1 else 2
            cfg = time_slots_dict.get(slot_num, time_slots_dict[1])
            return f"{cfg['start']} - {cfg['end']}"
        except:
            return f"{time_slots_dict[1]['start']} - {time_slots_dict[1]['end']}"

    sheets_processed = 0
    
    for sheet_name, sheet_df in df_dict.items():
        if sheet_df.empty: continue
        
        # Parse Sheet Name
        parts = sheet_name.split('_Sem_')
        if len(parts) < 2: continue
        main_branch_full = parts[0] # Assuming temp excel saves Full Name or we map it back
        
        semester_raw = parts[1]
        if semester_raw.endswith('_Electives'):
            semester_raw = semester_raw.replace('_Electives', '')
        
        # Convert Roman/String sem to int for logic
        roman_map = {'I':1, 'II':2, 'III':3, 'IV':4, 'V':5, 'VI':6, 'VII':7, 'VIII':8, 'IX':9, 'X':10}
        try:
            sem_int = roman_map.get(semester_raw, int(semester_raw) if semester_raw.isdigit() else 1)
        except:
            sem_int = 1
            
        header_content = {'main_branch_full': main_branch_full, 'semester_roman': semester_raw if not semester_raw.isdigit() else int_to_roman(int(semester_raw))}
        header_exam_time = get_header_time_for_semester(sem_int)

        if not sheet_name.endswith('_Electives'):
            fixed_cols = ["Exam Date"]
            sub_branch_cols = [c for c in sheet_df.columns if c not in fixed_cols and c != 'Note' and "Unnamed" not in str(c)]
            
            if not sub_branch_cols: continue
            
            exam_date_width = 60
            
            # Pagination for columns
            for start in range(0, len(sub_branch_cols), sub_branch_cols_per_page):
                chunk = sub_branch_cols[start:start + sub_branch_cols_per_page]
                cols_to_print = fixed_cols + chunk
                chunk_df = sheet_df[cols_to_print].copy()
                
                # Filter empty rows
                chunk_df = chunk_df.dropna(how='all', subset=chunk)
                if chunk_df.empty: continue

                # Calculate widths
                page_width = pdf.w - 2 * pdf.l_margin
                sub_width = (page_width - exam_date_width) / max(len(chunk), 1)
                col_widths = [exam_date_width] + [sub_width] * len(chunk)
                
                pdf.add_page()
                print_table_custom(pdf, chunk_df, cols_to_print, col_widths, 10, header_content, chunk, header_exam_time, None, declaration_date)
                sheets_processed += 1
                
        else: # Electives
            cols = [c for c in ['Exam Date', 'OE Type', 'Subjects'] if c in sheet_df.columns]
            if len(cols) < 2: continue
            
            elec_df = sheet_df[cols].dropna(how='all')
            if elec_df.empty: continue
            
            exam_date_width = 60
            oe_width = 30
            subject_width = pdf.w - 2 * pdf.l_margin - exam_date_width - oe_width
            col_widths = [exam_date_width, subject_width + oe_width] if len(cols) == 2 else [exam_date_width, oe_width, subject_width]
            
            pdf.add_page()
            print_table_custom(pdf, elec_df, cols, col_widths, 10, header_content, ['All Streams'], header_exam_time, None, declaration_date)
            sheets_processed += 1

    # INSTRUCTIONS PAGE
    try:
        pdf.add_page()
        footer_height = 25
        add_footer_with_page_number(pdf, footer_height)
        instr_header = {'main_branch_full': 'EXAMINATION GUIDELINES', 'semester_roman': 'General'}
        logo_w = 45
        add_header_to_page(pdf, (pdf.w-logo_w)/2, logo_w, instr_header, ["All Candidates"], None, None, declaration_date)
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
        pdf.set_font("Arial", size=12)
        for i in instrs:
            pdf.multi_cell(0, 8, i)
            pdf.ln(2)
    except Exception as e:
        st.warning(f"Instructions error: {e}")

    try:
        pdf.output(pdf_path)
        return True
    except Exception as e:
        st.error(f"Error saving PDF: {e}")
        return False

# ==========================================
# 💾 EXCEL INTERMEDIARY (The "Formatting" Engine)
# ==========================================

def save_to_excel(semester_wise_timetable):
    """
    Converts the processed Dictionary data into the pivoted Excel structure needed for the PDF generator.
    Handles the Logic: Compare Subject Duration Time vs Header Time.
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

    def normalize_time(t_str):
        if not isinstance(t_str, str): return ""
        t_str = t_str.strip().upper()
        for i in range(1, 10):
            t_str = t_str.replace(f"0{i}:", f"{i}:")
        return t_str

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sem, df_sem in semester_wise_timetable.items():
            if df_sem.empty: continue
            
            # HEADER TIME LOGIC
            slot_indicator = ((sem + 1) // 2) % 2
            primary_slot_num = 1 if slot_indicator == 1 else 2
            primary_slot_config = time_slots_dict.get(primary_slot_num, time_slots_dict[1])
            primary_slot_str = f"{primary_slot_config['start']} - {primary_slot_config['end']}"
            primary_slot_norm = normalize_time(primary_slot_str)
            
            for main_branch in df_sem["MainBranch"].unique():
                df_mb = df_sem[df_sem["MainBranch"] == main_branch].copy()
                if df_mb.empty: continue
                
                df_non_elec = df_mb[df_mb['OE'].isna() | (df_mb['OE'].str.strip() == "")].copy()
                df_elec = df_mb[df_mb['OE'].notna() & (df_mb['OE'].str.strip() != "")].copy()
                
                roman_sem = int_to_roman(sem)
                sheet_name = f"{main_branch}_Sem_{roman_sem}"[:31]

                # --- CORE SUBJECTS ---
                if not df_non_elec.empty:
                    displays = []
                    for idx, row in df_non_elec.iterrows():
                        base_subject = str(row.get('Subject', ''))
                        module_code = str(row.get('ModuleCode', ''))
                        
                        # Calculate Actual Time
                        assigned_slot = str(row.get('Time Slot', '')).strip()
                        duration = float(row.get('Exam Duration', 3.0))
                        
                        actual_time = assigned_slot
                        if assigned_slot and " - " in assigned_slot:
                            try:
                                start = assigned_slot.split(" - ")[0].strip()
                                end = calculate_end_time(start, duration)
                                actual_time = f"{start} - {end}"
                            except: pass
                        
                        # Comparison
                        if normalize_time(actual_time) != primary_slot_norm and actual_time:
                            time_suffix = f" [{actual_time}]"
                        else:
                            time_suffix = ""
                        
                        full_str = f"{base_subject}"
                        if module_code and module_code.lower() != 'nan':
                            full_str += f" - ({module_code})"
                        full_str += time_suffix
                        displays.append(full_str)
                    
                    df_non_elec['SubjectDisplay'] = displays
                    
                    # Pivot
                    try:
                        pivot = df_non_elec.pivot_table(index='Exam Date', columns='SubBranch', values='SubjectDisplay', aggfunc=lambda x: "\n".join(x)).fillna("---")
                        pivot.reset_index(inplace=True)
                        pivot.to_excel(writer, sheet_name=sheet_name, index=False)
                    except: pass

                # --- ELECTIVES ---
                if not df_elec.empty:
                    displays = []
                    for idx, row in df_elec.iterrows():
                        base_subject = str(row.get('Subject', ''))
                        module_code = str(row.get('ModuleCode', ''))
                        assigned_slot = str(row.get('Time Slot', '')).strip()
                        duration = float(row.get('Exam Duration', 3.0))
                        
                        actual_time = assigned_slot
                        if assigned_slot and " - " in assigned_slot:
                            try:
                                start = assigned_slot.split(" - ")[0].strip()
                                end = calculate_end_time(start, duration)
                                actual_time = f"{start} - {end}"
                            except: pass
                            
                        if normalize_time(actual_time) != primary_slot_norm and actual_time:
                            time_suffix = f" [{actual_time}]"
                        else:
                            time_suffix = ""

                        full_str = f"{base_subject}"
                        if module_code and module_code.lower() != 'nan':
                            full_str += f" - ({module_code})"
                        full_str += time_suffix
                        displays.append(full_str)
                        
                    df_elec['SubjectDisplay'] = displays
                    try:
                        # Electives usually aggregated
                        elec_pivot = df_elec.groupby(['OE', 'Exam Date']).agg({'SubjectDisplay': lambda x: "\n".join(set(x))}).reset_index()
                        elec_pivot.rename(columns={'OE': 'OE Type', 'SubjectDisplay': 'Subjects'}, inplace=True)
                        elec_sheet = f"{main_branch}_Sem_{roman_sem}_Electives"[:31]
                        elec_pivot.to_excel(writer, sheet_name=elec_sheet, index=False)
                    except: pass
    
    output.seek(0)
    return output

# ==========================================
# 📥 INPUT PROCESSOR (Verification File -> Logic Dict)
# ==========================================

def process_verification_file(uploaded_file):
    """
    Reads the Verification Excel and converts it into the 'semester_wise_timetable'
    dictionary structure used by the PDF engine.
    """
    try:
        df = pd.read_excel(uploaded_file)
        
        # 1. Column Mapping (Verification -> Internal Logic)
        # Required: Program, Stream, Current Session, Module Description, Exam Date, Configured Slot
        required = ['Program', 'Stream', 'Current Session', 'Module Description', 'Exam Date', 'Configured Slot']
        if not all(col in df.columns for col in required):
            st.error(f"Missing columns. Required: {required}")
            return None, None

        # Clean Data
        df = df[df['Exam Date'].notna() & (df['Exam Date'] != "Not Scheduled")].copy()
        
        # 2. Parse Semesters
        def parse_sem(val):
            s = str(val).upper()
            import re
            m = re.search(r'(\d+)', s)
            if m: return int(m.group(1))
            roman_to_int = {'I':1, 'II':2, 'III':3, 'IV':4, 'V':5, 'VI':6, 'VII':7, 'VIII':8, 'IX':9, 'X':10}
            for r, i in roman_to_int.items():
                if r in s: return i
            return 1

        df['Semester_Int'] = df['Current Session'].apply(parse_sem)
        
        # 3. Standardize Columns for "Save To Excel" function
        df['MainBranch'] = df['Program'].astype(str).str.strip()
        df['SubBranch'] = df['Stream'].astype(str).str.strip()
        # Fill empty streams with MainBranch
        df.loc[df['SubBranch'].isin(['nan', '']), 'SubBranch'] = df.loc[df['SubBranch'].isin(['nan', '']), 'MainBranch']
        
        df['Subject'] = df['Module Description']
        df['Time Slot'] = df['Configured Slot']
        df['ModuleCode'] = df['Module Abbreviation'] if 'Module Abbreviation' in df.columns else ''
        
        # Handle OE
        if 'Subject Type' in df.columns:
            df['OE'] = df['Subject Type'].apply(lambda x: 'OE' if str(x).strip().upper() == 'OE' else None)
        else:
            df['OE'] = None

        # Handle Duration
        if 'Exam Duration' in df.columns:
            df['Exam Duration'] = pd.to_numeric(df['Exam Duration'], errors='coerce').fillna(3.0)
        else:
            df['Exam Duration'] = 3.0
            
        # Format Date
        # Ensure dates are strings DD-MM-YYYY for consistency
        df['Exam Date'] = pd.to_datetime(df['Exam Date'], dayfirst=True, errors='coerce').dt.strftime('%d-%m-%Y')

        # 4. Group into Dictionary
        timetable = {}
        for sem in sorted(df['Semester_Int'].unique()):
            timetable[sem] = df[df['Semester_Int'] == sem].copy()
            
        return timetable, df # Return df for Stats
        
    except Exception as e:
        st.error(f"Error processing file: {e}")
        return None, None

# ==========================================
# 🖥️ MAIN APP
# ==========================================

def main():
    # CSS Styling
    st.markdown("""
    <style>
        .main-header { padding: 2rem; border-radius: 10px; margin-bottom: 2rem; background: linear-gradient(90deg, #951C1C, #C73E1D); color: white; text-align: center; }
        .main-header h1 { color: white; margin: 0; font-size: 2.5rem; text-shadow: 2px 2px 4px rgba(0,0,0,0.3); }
        .upload-section { background: #f8f9fa; padding: 2rem; border-radius: 10px; border: 2px dashed #951C1C; margin: 1rem 0; }
        div.row-widget.stButton > button { width: 100%; height: auto; min-height: 100px; border-radius: 12px; background: linear-gradient(135deg, #4a5db0 0%, #764ba2 100%); color: white; border: none; }
        div.row-widget.stButton > button p { font-size: 1.5rem !important; font-weight: 700 !important; }
        div.row-widget.stButton > button:hover { transform: translateY(-4px); box-shadow: 0 6px 12px rgba(0,0,0,0.2); color: white !important; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="main-header"><h1>📅 Timetable PDF Generator</h1><p>From Verification File</p></div>', unsafe_allow_html=True)

    # Session State
    if 'processed_data' not in st.session_state: st.session_state.processed_data = None
    if 'raw_df' not in st.session_state: st.session_state.raw_df = None
    if 'selected_college' not in st.session_state: st.session_state.selected_college = "SVKM's NMIMS University"

    # Sidebar
    with st.sidebar:
        st.header("⚙️ Configuration")
        colleges = ["SVKM's NMIMS University", "MUKESH PATEL SCHOOL OF TECHNOLOGY", "Custom..."]
        sel = st.selectbox("College Name", colleges)
        if sel == "Custom...":
            st.session_state.selected_college = st.text_input("Enter Name")
        else:
            st.session_state.selected_college = sel
            
        st.markdown("---")
        decl_date = st.date_input("📆 Declaration Date", value=None)

    # Layout
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown('<div class="upload-section"><h3>📁 Upload Verification File</h3></div>', unsafe_allow_html=True)
        uploaded = st.file_uploader("Upload Excel", type=['xlsx', 'xls'])
        
        if uploaded:
            timetable, raw_df = process_verification_file(uploaded)
            if timetable:
                st.session_state.processed_data = timetable
                st.session_state.raw_df = raw_df
                st.success("✅ File Processed Successfully!")

    with col2:
        st.info("ℹ️ This tool converts your finalized Verification Excel file into the official PDF format. No rescheduling is performed.")

    # Stats & Generate
    if st.session_state.raw_df is not None:
        df = st.session_state.raw_df
        
        # Calculate Stats
        s_exams = len(df)
        s_sems = df['Semester_Int'].nunique()
        s_progs = df['MainBranch'].nunique()
        s_streams = df['SubBranch'].nunique()
        
        # Stats Buttons
        st.markdown("### 📊 Overview")
        b1, b2, b3, b4 = st.columns(4)
        with b1:
            if st.button(f"{s_exams}\nTOTAL EXAMS", key="s1"): show_exams_breakdown(df)
        with b2:
            if st.button(f"{s_sems}\nSEMESTERS", key="s2"): show_semesters_breakdown(df)
        with b3:
            if st.button(f"{s_progs} / {s_streams}\nPROGS / STREAMS", key="s3"): show_programs_streams_breakdown(df)
        with b4:
            if st.button(f"📅\nSPAN INFO", key="s4"): show_span_breakdown(df)

        st.markdown("---")
        
        # Generate Button
        if st.button("🚀 Generate PDF Timetable", type="primary", use_container_width=True):
            with st.spinner("Generating PDF..."):
                # 1. Create Temp Excel (Formatted)
                excel_bytes = save_to_excel(st.session_state.processed_data)
                
                # 2. Save Temp file
                with open("temp_pdf_input.xlsx", "wb") as f:
                    f.write(excel_bytes.getvalue())
                
                # 3. Convert to PDF
                if convert_excel_to_pdf("temp_pdf_input.xlsx", "Final_Timetable.pdf", declaration_date=decl_date):
                    st.success("🎉 PDF Generated!")
                    with open("Final_Timetable.pdf", "rb") as f:
                        st.download_button("📥 Download PDF", f, "Timetable.pdf", "application/pdf", use_container_width=True)
                    
                    # Cleanup
                    try: os.remove("temp_pdf_input.xlsx")
                    except: pass
                    try: os.remove("Final_Timetable.pdf")
                    except: pass

if __name__ == "__main__":
    main()
