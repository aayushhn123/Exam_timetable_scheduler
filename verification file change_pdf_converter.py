import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from fpdf import FPDF
import os
import io
from PyPDF2 import PdfReader, PdfWriter
import re

# Set page configuration
st.set_page_config(
    page_title="Excel to PDF Timetable Converter",
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
    
    tab1, tab2 = st.tabs(["📊 By Subject Type", "🌿 By Program"])
    
    with tab1:
        # Breakdown by OE vs Core
        if 'Subject Type' in df.columns:
            type_counts = df['Subject Type'].replace('', 'Core').value_counts().reset_index()
            type_counts.columns = ['Type', 'Count']
            st.dataframe(type_counts, use_container_width=True, hide_index=True)
            
    with tab2:
        if 'Program' in df.columns:
            prog_counts = df['Program'].value_counts().reset_index()
            prog_counts.columns = ['Program', 'Exam Count']
            st.dataframe(prog_counts, use_container_width=True, hide_index=True)

@dialog_decorator("🎓 Semesters Breakdown")
def show_semesters_breakdown(df):
    if 'Semester' in df.columns:
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
    else:
        st.warning("Semester information not available yet.")

@dialog_decorator("🏫 Programs & Streams Breakdown")
def show_programs_streams_breakdown(df):
    # Get lists of programs and streams
    programs = sorted(df['Program'].dropna().unique()) if 'Program' in df.columns else []
    streams = sorted(df['Stream'].dropna().unique()) if 'Stream' in df.columns else []
    
    # Summary Metrics
    col1, col2 = st.columns(2)
    col1.metric("Total Programs", len(programs))
    col2.metric("Total Streams", len(streams))
    
    st.markdown("---")
    
    # Tabs for detailed view
    tab1, tab2 = st.tabs(["📂 Grouped by Program", "💧 All Streams List"])
    
    with tab1:
        if 'Program' in df.columns and 'Stream' in df.columns:
            for prog in programs:
                # Find streams belonging to this program
                prog_streams = df[df['Program'] == prog]['Stream'].dropna().unique()
                prog_streams = [s for s in prog_streams if str(s).strip() != '']
                
                with st.expander(f"**{prog}** ({len(prog_streams)} Streams)"):
                    if len(prog_streams) > 0:
                        for s in sorted(prog_streams):
                            st.caption(f"• {s}")
                    else:
                        st.caption("No specific streams defined.")
    
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
# END DIALOGS
# ==========================================

# Custom CSS for consistent dark and light mode styling
st.markdown("""
<style>
    /* Base styles */
    .main-header {
        padding: 2rem;
        border-radius: 10px;
        margin-bottom: 2rem;
    }

    .main-header h1 {
        color: white;
        text-align: center;
        margin: 0;
        font-size: 2.5rem;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    }

    .main-header p {
        color: #FFF;
        text-align: center;
        margin: 0.5rem 0 0 0;
        font-size: 1.2rem;
        opacity: 0.9;
    }

    /* Stats Button Styling */
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
        transition: transform 0.2s;
    }
    div.row-widget.stButton > button:hover {
        transform: translateY(-4px);
        box-shadow: 0 6px 12px rgba(0,0,0,0.2);
        border-color: rgba(255,255,255,0.5);
        color: white !important;
    }
    div.row-widget.stButton > button p {
        font-size: 1.8rem !important;
        font-weight: 700 !important;
        margin-bottom: 0.2rem !important;
        line-height: 1.2 !important;
    }
    div.row-widget.stButton > button div p:last-child {
        font-size: 0.9rem !important;
        opacity: 0.9;
        font-weight: 500 !important;
        text-transform: uppercase;
    }

    /* Light mode styles */
    @media (prefers-color-scheme: light) {
        .main-header {
            background: linear-gradient(90deg, #951C1C, #C73E1D);
        }

        .upload-section {
            background: #f8f9fa;
            padding: 2rem;
            border-radius: 10px;
            border: 2px dashed #951C1C;
            margin: 1rem 0;
        }

        .results-section {
            background: #ffffff;
            padding: 2rem;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            margin: 1rem 0;
        }

        .status-success {
            background: #d4edda;
            color: #155724;
            padding: 1rem;
            border-radius: 5px;
            border-left: 4px solid #28a745;
        }

        .feature-card {
            background: white;
            padding: 1.5rem;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            margin: 1rem 0;
            border-left: 4px solid #951C1C;
        }
    }

    /* Dark mode styles */
    @media (prefers-color-scheme: dark) {
        .main-header {
            background: linear-gradient(90deg, #701515, #A23217);
        }

        .upload-section {
            background: #333;
            padding: 2rem;
            border-radius: 10px;
            border: 2px dashed #A23217;
            margin: 1rem 0;
        }

        .results-section {
            background: #222;
            padding: 2rem;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
            margin: 1rem 0;
        }

        .status-success {
            background: #2d4b2d;
            color: #e6f4ea;
            padding: 1rem;
            border-radius: 5px;
            border-left: 4px solid #4caf50;
        }

        .feature-card {
            background: #333;
            padding: 1.5rem;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.3);
            margin: 1rem 0;
            border-left: 4px solid #A23217;
        }
    }
</style>
""", unsafe_allow_html=True)

# Define the mapping of main branch abbreviations to full forms
BRANCH_FULL_FORM = {
    "B TECH": "BACHELOR OF TECHNOLOGY",
    "B TECH INTG": "BACHELOR OF TECHNOLOGY SIX YEAR INTEGRATED PROGRAM",
    "M TECH": "MASTER OF TECHNOLOGY",
    "MBA TECH": "MASTER OF BUSINESS ADMINISTRATION IN TECHNOLOGY MANAGEMENT",
    "MCA": "MASTER OF COMPUTER APPLICATIONS",
    "DIPLOMA": "DIPLOMA IN ENGINEERING"
}

# Define logo path
LOGO_PATH = "logo.png"

# Cache for text wrapping results
wrap_text_cache = {}

def wrap_text(pdf, text, col_width):
    cache_key = (text, col_width)
    if cache_key in wrap_text_cache:
        return wrap_text_cache[cache_key]
    
    lines = []
    current_line = ""
    words = re.split(r'(\s+)', text)
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
    wrap_text_cache[cache_key] = lines
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

def print_table_custom(pdf, df, columns, col_widths, line_height=5, header_content=None, branches=None, time_slot=None, actual_time_slots=None, declaration_date=None):
    if df.empty:
        return
    setattr(pdf, '_row_counter', 0)
    
    # Add footer first
    footer_height = 25
    pdf.set_xy(10, pdf.h - footer_height)
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 5, "Controller of Examinations", 0, 1, 'L')
    pdf.line(10, pdf.h - footer_height + 5, 60, pdf.h - footer_height + 5)
    
    # Add page numbers in bottom right
    pdf.set_font("Arial", size=14)
    pdf.set_text_color(0, 0, 0)
    page_text = f"{pdf.page_no()} of {{nb}}"
    text_width = pdf.get_string_width(page_text.replace("{nb}", "99"))
    pdf.set_xy(pdf.w - 10 - text_width, pdf.h - footer_height + 12)
    pdf.cell(text_width, 5, page_text, 0, 0, 'R')
    
    # Add header
    header_height = 95
    pdf.set_y(0)
    
    # NEW: Declaration Date at Top Right
    if declaration_date:
        pdf.set_font("Arial", 'B', 10)
        pdf.set_text_color(0, 0, 0)
        decl_str = f"Declaration Date: {declaration_date.strftime('%d-%m-%Y')}"
        pdf.set_xy(pdf.w - 60, 10)
        pdf.cell(50, 10, decl_str, 0, 0, 'R')

    logo_width = 45
    logo_x = (pdf.w - logo_width) / 2
    if os.path.exists(LOGO_PATH):
        pdf.image(LOGO_PATH, x=logo_x, y=10, w=logo_width)
    
    pdf.set_fill_color(149, 33, 28)
    pdf.set_text_color(255, 255, 255)
    
    # Get selected college name from session state
    college_name = st.session_state.get('selected_college', 'SVKM\'s NMIMS University')
    
    # Adjust font size based on college name length
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
    
    # Display actual time slots used
    if actual_time_slots and len(actual_time_slots) > 0:
        pdf.set_font("Arial", 'B', 13)
        pdf.set_xy(10, 59)
        
        if len(actual_time_slots) == 1:
            slot_text = f"Exam Time: {list(actual_time_slots)[0]}"
        else:
            sorted_slots = sorted(actual_time_slots)
            slot_text = f"Exam Times: {' | '.join(sorted_slots)}"
        
        pdf.cell(pdf.w - 20, 6, slot_text, 0, 1, 'C')
        pdf.set_font("Arial", 'I', 10)
        pdf.set_xy(10, 65)
        pdf.cell(pdf.w - 20, 6, "(Check individual subject for specific exam time if multiple slots shown)", 0, 1, 'C')
        pdf.set_font("Arial", '', 12)
        pdf.set_xy(10, 71)
        pdf.cell(pdf.w - 20, 6, f"Branches: {', '.join(branches)}", 0, 1, 'C')
        pdf.set_y(85)
    elif time_slot:
        pdf.set_font("Arial", 'B', 14)
        pdf.set_xy(10, 59)
        pdf.cell(pdf.w - 20, 6, f"Exam Time: {time_slot}", 0, 1, 'C')
        pdf.set_font("Arial", 'I', 10)
        pdf.set_xy(10, 65)
        pdf.cell(pdf.w - 20, 6, "(Check the subject exam time)", 0, 1, 'C')
        pdf.set_font("Arial", '', 12)
        pdf.set_xy(10, 71)
        pdf.cell(pdf.w - 20, 6, f"Branches: {', '.join(branches)}", 0, 1, 'C')
        pdf.set_y(85)
    else:
        pdf.set_font("Arial", 'I', 10)
        pdf.set_xy(10, 59)
        pdf.cell(pdf.w - 20, 6, "(Check the subject exam time)", 0, 1, 'C')
        pdf.set_font("Arial", '', 12)
        pdf.set_xy(10, 65)
        pdf.cell(pdf.w - 20, 6, f"Branches: {', '.join(branches)}", 0, 1, 'C')
        pdf.set_y(71)
    
    # Calculate available space
    available_height = pdf.h - pdf.t_margin - footer_height - header_height
    pdf.set_font("Arial", size=12)
    
    # Print header row
    print_row_custom(pdf, columns, col_widths, line_height=line_height, header=True)
    
    # Print data rows
    for idx in range(len(df)):
        row = [str(df.iloc[idx][c]) if pd.notna(df.iloc[idx][c]) else "" for c in columns]
        if not any(cell.strip() for cell in row):
            continue
            
        # Estimate row height
        wrapped_cells = []
        max_lines = 0
        for i, cell_text in enumerate(row):
            text = str(cell_text) if cell_text is not None else ""
            avail_w = col_widths[i] - 4
            lines = wrap_text(pdf, text, avail_w)
            wrapped_cells.append(lines)
            max_lines = max(max_lines, len(lines))
        row_h = line_height * max_lines
        
        # Check if row fits
        if pdf.get_y() + row_h > pdf.h - footer_height:
            # Add footer to current page
            add_footer_with_page_number(pdf, footer_height)
            
            # Start new page
            pdf.add_page()
            add_footer_with_page_number(pdf, footer_height)
            
            # Add header to new page
            add_header_to_page(pdf, logo_x, logo_width, header_content, branches, time_slot, actual_time_slots, declaration_date)
            
            # Reprint header row
            pdf.set_font("Arial", size=12)
            print_row_custom(pdf, columns, col_widths, line_height=line_height, header=True)
        
        print_row_custom(pdf, row, col_widths, line_height=line_height, header=False)

def add_footer_with_page_number(pdf, footer_height):
    """Add footer with signature and page number"""
    pdf.set_xy(10, pdf.h - footer_height)
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 5, "Controller of Examinations", 0, 1, 'L')
    pdf.line(10, pdf.h - footer_height + 5, 60, pdf.h - footer_height + 5)
    
    # Add page numbers in bottom right
    pdf.set_font("Arial", size=14)
    pdf.set_text_color(0, 0, 0)
    page_text = f"{pdf.page_no()} of {{nb}}"
    text_width = pdf.get_string_width(page_text.replace("{nb}", "99"))
    pdf.set_xy(pdf.w - 10 - text_width, pdf.h - footer_height + 12)
    pdf.cell(text_width, 5, page_text, 0, 0, 'R')

def add_header_to_page(pdf, logo_x, logo_width, header_content, branches, time_slot=None, actual_time_slots=None, declaration_date=None):
    """Add header to a new page"""
    pdf.set_y(0)
    
    # NEW: Declaration Date
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
    
    # Get selected college name
    college_name = st.session_state.get('selected_college', 'SVKM\'s NMIMS University')
    
    # Adjust font size
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
    
    # Display time slots
    if actual_time_slots and len(actual_time_slots) > 0:
        pdf.set_font("Arial", 'B', 13)
        pdf.set_xy(10, 59)
        
        if len(actual_time_slots) == 1:
            slot_text = f"Exam Time: {list(actual_time_slots)[0]}"
        else:
            sorted_slots = sorted(actual_time_slots)
            slot_text = f"Exam Times: {' | '.join(sorted_slots)}"
        
        pdf.cell(pdf.w - 20, 6, slot_text, 0, 1, 'C')
        pdf.set_font("Arial", 'I', 10)
        pdf.set_xy(10, 65)
        pdf.cell(pdf.w - 20, 6, "(Check individual subject for specific exam time if multiple slots shown)", 0, 1, 'C')
        pdf.set_font("Arial", '', 12)
        pdf.set_xy(10, 71)
        pdf.cell(pdf.w - 20, 6, f"Branches: {', '.join(branches)}", 0, 1, 'C')
        pdf.set_y(85)
    elif time_slot:
        pdf.set_font("Arial", 'B', 14)
        pdf.set_xy(10, 59)
        pdf.cell(pdf.w - 20, 6, f"Exam Time: {time_slot}", 0, 1, 'C')
        pdf.set_font("Arial", 'I', 10)
        pdf.set_xy(10, 65)
        pdf.cell(pdf.w - 20, 6, "(Check the subject exam time)", 0, 1, 'C')
        pdf.set_font("Arial", '', 12)
        pdf.set_xy(10, 71)
        pdf.cell(pdf.w - 20, 6, f"Branches: {', '.join(branches)}", 0, 1, 'C')
        pdf.set_y(85)
    else:
        pdf.set_font("Arial", 'I', 10)
        pdf.set_xy(10, 59)
        pdf.cell(pdf.w - 20, 6, "(Check the subject exam time)", 0, 1, 'C')
        pdf.set_font("Arial", '', 12)
        pdf.set_xy(10, 65)
        pdf.cell(pdf.w - 20, 6, f"Branches: {', '.join(branches)}", 0, 1, 'C')
        pdf.set_y(71)

def int_to_roman(num):
    """Convert integer to Roman numeral"""
    roman_values = [
        (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
        (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
        (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")
    ]
    result = ""
    for value, numeral in roman_values:
        while num >= value:
            result += numeral
            num -= value
    return result

def read_verification_excel(uploaded_file):
    """Read the NEW verification Excel file format"""
    try:
        # Try to read all sheets
        excel_file = pd.ExcelFile(uploaded_file)
        
        # Look for the Verification sheet
        if 'Verification' in excel_file.sheet_names:
            df = pd.read_excel(uploaded_file, sheet_name='Verification', engine='openpyxl')
        else:
            # Use the first sheet if Verification not found
            df = pd.read_excel(uploaded_file, sheet_name=0, engine='openpyxl')
        
        # NEW column names - Use "Configured Slot" instead of "Exam Time"
        required_columns = ['Program', 'Stream', 'Current Session', 'Module Description', 'Exam Date', 'Configured Slot']
        
        # Check for required columns
        missing_required = [col for col in required_columns if col not in df.columns]
        if missing_required:
            st.error(f"❌ Missing required columns: {missing_required}")
            st.info("💡 Required columns: " + ", ".join(required_columns))
            return None
        
        # Filter out rows with no exam date or "Not Scheduled"
        df = df[
            (df['Exam Date'].notna()) & 
            (df['Exam Date'] != "") & 
            (df['Exam Date'] != 'Not Scheduled') &
            (df['Exam Date'].astype(str).str.strip() != "")
        ].copy()
        
        if df.empty:
            st.error("❌ No valid scheduled exams found in the verification file")
            return None
        
        # Clean data
        df['Program'] = df['Program'].astype(str).str.strip().str.upper()
        df['Stream'] = df['Stream'].astype(str).str.strip()
        df['Current Session'] = df['Current Session'].astype(str).str.strip()
        df['Module Description'] = df['Module Description'].astype(str).str.strip()
        df['Exam Date'] = df['Exam Date'].astype(str).str.strip()
        df['Configured Slot'] = df['Configured Slot'].astype(str).str.strip()
        
        # Handle optional columns
        if 'Module Abbreviation' in df.columns:
            df['Module Abbreviation'] = df['Module Abbreviation'].astype(str).str.strip()
        
        if 'CM Group' in df.columns:
            df['CM Group'] = df['CM Group'].fillna("").astype(str).str.strip()
        
        if 'Exam Slot Number' in df.columns:
            df['Exam Slot Number'] = pd.to_numeric(df['Exam Slot Number'], errors='coerce').fillna(0).astype(int)
        
        # NEW: Handle Subject Type column to identify OE subjects
        if 'Subject Type' in df.columns:
            df['Subject Type'] = df['Subject Type'].fillna("").astype(str).str.strip()
        else:
            df['Subject Type'] = ""
        
        return df
        
    except Exception as e:
        st.error(f"❌ Error reading Excel file: {str(e)}")
        import traceback
        st.error(f"Full error details: {traceback.format_exc()}")
        return None

def create_excel_sheets_for_pdf(df):
    """Convert verification data to Excel sheet format for PDF generation"""
    
    # Parse Program-Stream combination
    def parse_program_stream(row):
        program = str(row['Program']).strip().upper()
        stream = str(row['Stream']).strip()
        
        # Create combined identifier
        if stream and stream != 'nan' and stream != program:
            return f"{program}-{stream}"
        else:
            return program
    
    df['Branch'] = df.apply(parse_program_stream, axis=1)
    
    # Parse semester to get number
    def parse_semester(session_str):
        if pd.isna(session_str):
            return 1
        
        session_str = str(session_str).strip()
        
        # Extract number from various formats
        import re
        num_match = re.search(r'(\d+)', session_str)
        if num_match:
            return int(num_match.group(1))
        
        # Try Roman numerals
        roman_to_num = {
            'I': 1, 'II': 2, 'III': 3, 'IV': 4, 'V': 5, 'VI': 6,
            'VII': 7, 'VIII': 8, 'IX': 9, 'X': 10, 'XI': 11, 'XII': 12
        }
        for roman, num in roman_to_num.items():
            if roman in session_str.upper():
                return num
        
        return 1
    
    df['Semester'] = df['Current Session'].apply(parse_semester)
    
    # Group by main program and semester
    def get_main_program(branch):
        # Extract main program part before dash
        if '-' in branch:
            return branch.split('-')[0].strip()
        return branch
    
    df['MainBranch'] = df['Branch'].apply(get_main_program)
    
    # Extract SubBranch (stream)
    def get_sub_branch(branch):
        if '-' in branch:
            parts = branch.split('-', 1)
            return parts[1].strip() if len(parts) > 1 else ""
        return ""
    
    df['SubBranch'] = df['Branch'].apply(get_sub_branch)
    
    # If no subbranch, use main branch
    df.loc[df['SubBranch'] == "", 'SubBranch'] = df.loc[df['SubBranch'] == "", 'MainBranch']
    
    # NEW: Identify OE subjects
    df['IsOE'] = df['Subject Type'].str.upper() == 'OE'
    
    excel_data = {}
    
    # Group by MainBranch and Semester
    for (main_branch, semester), group_df in df.groupby(['MainBranch', 'Semester']):
        
        # Separate OE and non-OE subjects
        non_oe_df = group_df[~group_df['IsOE']].copy()
        oe_df = group_df[group_df['IsOE']].copy()
        
        # Process non-OE subjects (CORE SUBJECTS)
        if not non_oe_df.empty:
            # Get all unique sub-branches (streams)
            all_sub_branches = sorted(non_oe_df['SubBranch'].unique())
            
            # Split into groups of 4 branches per page
            branches_per_page = 4
            
            for page_num, i in enumerate(range(0, len(all_sub_branches), branches_per_page), start=1):
                sub_branches = all_sub_branches[i:i + branches_per_page]
                
                # Create sheet name
                roman_sem = int_to_roman(semester)
                if len(all_sub_branches) > branches_per_page:
                    sheet_name = f"{main_branch}_Sem_{roman_sem}_Part{page_num}"
                else:
                    sheet_name = f"{main_branch}_Sem_{roman_sem}"
                
                if len(sheet_name) > 31:
                    sheet_name = sheet_name[:31]
                
                # Get all unique exam dates for this group
                all_dates = sorted(non_oe_df['Exam Date'].unique(), key=lambda x: pd.to_datetime(x, format='%d-%m-%Y', errors='coerce'))
                
                processed_data = []
                
                for exam_date in all_dates:
                    # Format date
                    try:
                        parsed_date = pd.to_datetime(exam_date, format='%d-%m-%Y', errors='coerce')
                        if pd.notna(parsed_date):
                            formatted_date = parsed_date.strftime("%A, %d %B, %Y")
                        else:
                            formatted_date = str(exam_date)
                    except:
                        formatted_date = str(exam_date)
                    
                    row_data = {'Exam Date': formatted_date}
                    
                    # For each sub-branch in this page group
                    for sub_branch in sub_branches:
                        subjects_on_date = non_oe_df[
                            (non_oe_df['Exam Date'] == exam_date) & 
                            (non_oe_df['SubBranch'] == sub_branch)
                        ]
                        
                        if not subjects_on_date.empty:
                            subjects = []
                            for _, row in subjects_on_date.iterrows():
                                subject_name = str(row['Module Description'])
                                module_code = str(row.get('Module Abbreviation', ''))
                                exam_time = str(row.get('Configured Slot', ''))
                                cm_group = str(row.get('CM Group', '')).strip()
                                exam_slot = row.get('Exam Slot Number', 0)
                                
                                # Build subject display
                                subject_display = subject_name
                                
                                # Add module code if present
                                if module_code and module_code != 'nan':
                                    subject_display = f"{subject_display} - ({module_code})"
                                
                                # Add CM Group prefix if present
                                if cm_group and cm_group != 'nan' and cm_group != '':
                                    try:
                                        cm_num = int(float(cm_group))
                                        subject_display = f"[CM:{cm_num}] {subject_display}"
                                    except:
                                        subject_display = f"[CM:{cm_group}] {subject_display}"
                                
                                # Add exam time from Configured Slot
                                if exam_time and exam_time != 'nan' and exam_time.strip():
                                    subject_display = f"{subject_display} ({exam_time})"
                                
                                # Add slot number if present
                                if exam_slot and exam_slot != 0:
                                    subject_display = f"{subject_display} [Slot {exam_slot}]"
                                
                                subjects.append(subject_display)
                            
                            # Join multiple subjects with line breaks
                            row_data[sub_branch] = "\n".join(subjects) if len(subjects) > 1 else subjects[0]
                        else:
                            # No subjects for this stream on this date
                            row_data[sub_branch] = "---"
                    
                    processed_data.append(row_data)
                
                # Convert to DataFrame
                if processed_data:
                    sheet_df = pd.DataFrame(processed_data)
                    
                    # Reorder columns to have Exam Date first, then the streams in order
                    column_order = ['Exam Date'] + sub_branches
                    sheet_df = sheet_df[column_order]
                    
                    # Fill any missing cells with "---"
                    sheet_df = sheet_df.fillna("---")
                    
                    excel_data[sheet_name] = sheet_df
        
        # Process OE subjects (OPEN ELECTIVE)
        if not oe_df.empty:
            roman_sem = int_to_roman(semester)
            oe_sheet_name = f"{main_branch}_Sem_{roman_sem}_OE"
            if len(oe_sheet_name) > 31: oe_sheet_name = oe_sheet_name[:31]
            
            oe_dates = sorted(oe_df['Exam Date'].unique(), key=lambda x: pd.to_datetime(x, format='%d-%m-%Y', errors='coerce'))
            oe_processed_data = []
            
            for exam_date in oe_dates:
                try:
                    parsed_date = pd.to_datetime(exam_date, format='%d-%m-%Y', errors='coerce')
                    formatted_date = parsed_date.strftime("%A, %d %B, %Y") if pd.notna(parsed_date) else str(exam_date)
                except:
                    formatted_date = str(exam_date)
                
                oe_subjects_on_date = oe_df[oe_df['Exam Date'] == exam_date]
                if not oe_subjects_on_date.empty:
                    seen_subjects = set()
                    subjects = []
                    
                    for _, row in oe_subjects_on_date.iterrows():
                        subject_name = str(row['Module Description'])
                        module_code = str(row.get('Module Abbreviation', ''))
                        exam_time = str(row.get('Configured Slot', ''))
                        exam_slot = row.get('Exam Slot Number', 0)
                        
                        subject_key = f"{subject_name}|{module_code}"
                        if subject_key in seen_subjects: continue
                        seen_subjects.add(subject_key)
                        
                        subject_display = subject_name
                        if module_code and module_code != 'nan':
                            subject_display = f"{subject_display} - ({module_code})"
                        if exam_time and exam_time != 'nan' and exam_time.strip():
                            subject_display = f"{subject_display} ({exam_time})"
                        if exam_slot and exam_slot != 0:
                            subject_display = f"{subject_display} [Slot {exam_slot}]"
                        
                        subjects.append(subject_display)
                    
                    if subjects:
                        row_data = {'Exam Date': formatted_date, 'Open Elective Subjects': "\n".join(subjects)}
                        oe_processed_data.append(row_data)
            
            if oe_processed_data:
                oe_sheet_df = pd.DataFrame(oe_processed_data)
                excel_data[oe_sheet_name] = oe_sheet_df
    
    return excel_data
    
def generate_pdf_from_excel_data(excel_data, output_pdf, declaration_date=None):
    """Generate PDF from Excel data dictionary"""
    pdf = FPDF(orientation='L', unit='mm', format='A3')
    pdf.set_auto_page_break(auto=False, margin=15)
    pdf.alias_nb_pages()
    
    sheets_processed = 0
    
    for sheet_name, sheet_df in excel_data.items():
        if sheet_df.empty:
            continue
        
        # Check if this is an OE sheet
        is_oe_sheet = '_OE' in sheet_name
        
        # Parse sheet name
        try:
            name_parts = sheet_name.split('_')
            sem_index = name_parts.index('Sem')
            program_parts = name_parts[:sem_index]
            program = '_'.join(program_parts) if program_parts else ''
            semester_roman = name_parts[sem_index + 1]
            
            program_norm = re.sub(r'[.\s]+', ' ', program).strip().upper()
            main_branch_full = BRANCH_FULL_FORM.get(program_norm, program_norm)
            
            if is_oe_sheet:
                header_content = {'main_branch_full': f"{main_branch_full} - OPEN ELECTIVES", 'semester_roman': semester_roman}
            else:
                header_content = {'main_branch_full': main_branch_full, 'semester_roman': semester_roman}
            
            fixed_cols = ["Exam Date"]
            stream_cols = [c for c in sheet_df.columns if c not in fixed_cols]
            if not stream_cols: continue
            
            if is_oe_sheet:
                cols_to_print = fixed_cols + stream_cols
                exam_date_width = 60
                oe_width = pdf.w - 20 - exam_date_width
                col_widths = [exam_date_width, oe_width]
                
                actual_time_slots = set()
                for cell_value in sheet_df['Open Elective Subjects']:
                    if pd.notna(cell_value) and str(cell_value) != "---":
                        time_match = re.findall(r'\(([^)]+)\)', str(cell_value))
                        for match in time_match:
                            if any(time_str in match for time_str in ['AM', 'PM', 'am', 'pm', ':']):
                                actual_time_slots.add(match)
                
                pdf.add_page()
                print_table_custom(
                    pdf, sheet_df, cols_to_print, col_widths, 
                    line_height=10, header_content=header_content, 
                    branches=['All Branches (Open Elective)'], 
                    actual_time_slots=actual_time_slots,
                    declaration_date=declaration_date
                )
            else:
                actual_stream_count = len(stream_cols)
                cols_to_print = fixed_cols + stream_cols
                
                actual_time_slots = set()
                for col in stream_cols:
                    for cell_value in sheet_df[col]:
                        if pd.notna(cell_value) and str(cell_value) != "---":
                            time_match = re.findall(r'\(([^)]+)\)', str(cell_value))
                            for match in time_match:
                                if any(time_str in match for time_str in ['AM', 'PM', 'am', 'pm', ':']):
                                    actual_time_slots.add(match)
                
                exam_date_width = 60
                remaining_width = pdf.w - 20 - exam_date_width
                stream_width = remaining_width / actual_stream_count if actual_stream_count > 0 else remaining_width
                col_widths = [exam_date_width] + [stream_width] * actual_stream_count
                
                pdf.add_page()
                print_table_custom(
                    pdf, sheet_df, cols_to_print, col_widths, 
                    line_height=10, header_content=header_content, 
                    branches=stream_cols, actual_time_slots=actual_time_slots,
                    declaration_date=declaration_date
                )
            
            sheets_processed += 1
            
        except Exception as e:
            st.warning(f"Error processing sheet {sheet_name}: {e}")
            continue
    
    if sheets_processed == 0:
        st.error("No sheets were processed for PDF generation!")
        return False
        
    # ====================================================================
    # --- NEW: Add Instructions to Candidates Page (Last Page) ---
    # ====================================================================
    try:
        # 1. Add new page
        pdf.add_page()

        # 2. Add Footer (Same as other pages)
        footer_height = 25
        add_footer_with_page_number(pdf, footer_height)

        # 3. Add Header
        instr_header_content = {
            'main_branch_full': 'EXAMINATION GUIDELINES', 
            'semester_roman': 'General'
        }
        
        logo_width = 45
        logo_x = (pdf.w - logo_width) / 2
        
        add_header_to_page(
            pdf, 
            logo_x=logo_x, 
            logo_width=logo_width,
            header_content=instr_header_content, 
            branches=["All Candidates"],
            time_slot=None, 
            actual_time_slots=None, 
            declaration_date=declaration_date
        )

        # 4. Set cursor position for content (Header ends around y=85-90)
        pdf.set_y(95)

        # 5. Print Title
        pdf.set_font("Arial", 'B', 14)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(0, 10, "INSTRUCTIONS TO CANDIDATES", 0, 1, 'C')
        pdf.ln(5)

        # 6. Print Instructions List
        instructions = [
            "1. Candidates are required to be present at the examination center THIRTY MINUTES before the stipulated time.",
            "2. Candidates must produce their University Identity Card at the time of the examination.",
            "3. Candidates are not permitted to enter the examination hall after stipulated time.",
            "4. Candidates will not be permitted to leave the examination hall during the examination time.",
            "5. Candidates are forbidden from taking any unauthorized material inside the examination hall. Carrying the same will be treated as usage of unfair means."
        ]

        pdf.set_font("Arial", size=12)
        
        # We use a loop with multi_cell to ensure text wrapping works properly
        for instr in instructions:
            pdf.multi_cell(0, 8, instr)
            pdf.ln(2)

    except Exception as e:
        st.warning(f"Could not add instructions page: {e}")
    
    try:
        pdf.output(output_pdf)
        st.success(f"PDF generated successfully with {sheets_processed} pages")
        return True
    except Exception as e:
        st.error(f"Error saving PDF: {e}")
        return False

def main():
    st.markdown("""
    <div class="main-header">
        <h1>📅 Excel to PDF Timetable Converter</h1>
        <p>Convert Excel verification files to formatted PDF timetables</p>
    </div>
    """, unsafe_allow_html=True)

    # Initialize session state
    if 'pdf_data' not in st.session_state:
        st.session_state.pdf_data = None
    if 'processing_complete' not in st.session_state:
        st.session_state.processing_complete = False
    if 'selected_college' not in st.session_state:
        st.session_state.selected_college = 'SVKM\'s NMIMS University'

    # Sidebar for college selection
    with st.sidebar:
        st.header("⚙️ Settings")
        
        college_options = [
            'SVKM\'s NMIMS University',
            'MUKESH PATEL SCHOOL OF TECHNOLOGY MANAGEMENT & ENGINEERING',
            'SCHOOL OF TECHNOLOGY MANAGEMENT & ENGINEERING',
            'Custom College Name'
        ]
        
        selected_college = st.selectbox(
            "Select College Name",
            college_options,
            index=0
        )
        
        if selected_college == 'Custom College Name':
            custom_college = st.text_input("Enter Custom College Name", "")
            if custom_college:
                st.session_state.selected_college = custom_college
            else:
                st.session_state.selected_college = 'SVKM\'s NMIMS University'
        else:
            st.session_state.selected_college = selected_college

        # NEW: Declaration Date Selector
        st.markdown("---")
        st.markdown("#### 📆 PDF Configuration")
        declaration_date = st.date_input(
            "Declaration Date (Optional)",
            value=None,
            help="Select a date to appear on the top right of the PDF. Leave empty to hide."
        )

    col1, col2 = st.columns([2, 1])

    with col1:
        st.markdown("""
        <div class="upload-section">
            <h3>📁 Upload Excel Verification File</h3>
            <p>Upload your verification Excel file with Exam Date and Exam Time columns</p>
        </div>
        """, unsafe_allow_html=True)

        uploaded_file = st.file_uploader(
            "Choose an Excel file",
            type=['xlsx', 'xls'],
            help="Upload the Excel verification file containing exam dates and times"
        )

        if uploaded_file is not None:
            st.markdown('<div class="status-success">✅ File uploaded successfully!</div>', unsafe_allow_html=True)

            file_details = {
                "Filename": uploaded_file.name,
                "File size": f"{uploaded_file.size / 1024:.2f} KB",
                "File type": uploaded_file.type
            }

            st.markdown("#### File Details:")
            for key, value in file_details.items():
                st.write(f"**{key}:** {value}")

    with col2:
        st.markdown("""
        <div class="feature-card">
            <h4>🚀 Features</h4>
            <ul>
                <li>📊 Direct Excel to PDF conversion</li>
                <li>📅 Preserves exam dates and times</li>
                <li>🎯 Program and stream grouping</li>
                <li>📝 Professional PDF formatting</li>
                <li>🏫 Customizable college header</li>
                <li>📋 Automatic page management</li>
                <li>⚡ Supports CM Groups & Slots</li>
                <li>🎓 Separate OE subject handling</li>
                <li>📱 Mobile-friendly interface</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

    if uploaded_file is not None:
        # Pre-read DF for stats
        df_for_stats = read_verification_excel(uploaded_file)
        
        if df_for_stats is not None:
            # Add mapping columns for stats compatibility
            if 'Current Session' in df_for_stats.columns:
                # Helper to parse semester number
                def parse_sem(val):
                    import re
                    s = str(val).upper()
                    if 'SEM' in s:
                        # Extract Roman or Number
                        pass
                    # Simple unique values
                    return str(val)
                df_for_stats['Semester'] = df_for_stats['Current Session']
                
            # ==========================================
            # 🔄 RE-CALCULATE VARIABLES FOR BREAKDOWN
            # ==========================================
            
            # --- 1. Non-Elective Stats ---
            non_elective_data = df_for_stats[df_for_stats['Subject Type'] != 'OE']
            ne_count = len(non_elective_data)
            
            if not non_elective_data.empty:
                ne_dates = pd.to_datetime(non_elective_data['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
                if not ne_dates.empty:
                    ne_start = min(ne_dates).strftime("%-d %B")
                    ne_end = max(ne_dates).strftime("%-d %B")
                    non_elec_display = f"{ne_start}" if ne_start == ne_end else f"{ne_start} to {ne_end}"
                else:
                    non_elec_display = "No dates"
            else:
                non_elec_display = "No subjects"

            # --- 2. OE Stats ---
            oe_data = df_for_stats[df_for_stats['Subject Type'] == 'OE']
            oe_count = len(oe_data)
            
            if not oe_data.empty:
                oe_dates = pd.to_datetime(oe_data['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
                if not oe_dates.empty:
                    sorted_oe_dt = sorted(oe_dates.unique())
                    u_oe_dates = [pd.to_datetime(d).strftime("%-d %B") for d in sorted_oe_dt]
                    
                    if len(u_oe_dates) == 1:
                        oe_display = u_oe_dates[0]
                    elif len(u_oe_dates) == 2:
                        oe_display = f"{u_oe_dates[0]}, {u_oe_dates[1]}"
                    else:
                        oe_display = f"{u_oe_dates[0]} to {u_oe_dates[-1]}"
                else:
                    oe_display = "No dates"
            else:
                oe_display = "No OE subjects"
            
            # --- Stats Metrics ---
            s_exams = len(df_for_stats)
            s_sems = df_for_stats['Current Session'].nunique()
            s_progs = df_for_stats['Program'].nunique()
            s_streams = df_for_stats['Stream'].nunique()
            
            d_dates = pd.to_datetime(df_for_stats['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
            s_span = (max(d_dates) - min(d_dates)).days + 1 if not d_dates.empty else 0

            # 2. Styling (Maintains your aesthetic)
            st.markdown("""
            <style>
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
                    transition: transform 0.2s;
                }
                div.row-widget.stButton > button:hover {
                    transform: translateY(-4px);
                    box-shadow: 0 6px 12px rgba(0,0,0,0.2);
                    border-color: rgba(255,255,255,0.5);
                    color: white !important;
                }
                div.row-widget.stButton > button p {
                    font-size: 1.8rem !important;
                    font-weight: 700 !important;
                    margin-bottom: 0.2rem !important;
                    line-height: 1.2 !important;
                }
                div.row-widget.stButton > button div p:last-child {
                    font-size: 0.9rem !important;
                    opacity: 0.9;
                    font-weight: 500 !important;
                    text-transform: uppercase;
                }
            </style>
            """, unsafe_allow_html=True)

            st.markdown("### 📈 Statistics Overview")
            
            # 3. Display Buttons
            stat_c1, stat_c2, stat_c3, stat_c4 = st.columns(4)
            
            with stat_c1:
                if st.button(f"{s_exams}\nTOTAL EXAMS", key="stat_btn_exams", use_container_width=True):
                    show_exams_breakdown(df_for_stats)

            with stat_c2:
                if st.button(f"{s_sems}\nSEMESTERS", key="stat_btn_sems", use_container_width=True):
                    show_semesters_breakdown(df_for_stats)

            with stat_c3:
                if st.button(f"{s_progs} / {s_streams}\nPROGS / STREAMS", key="stat_btn_progs", use_container_width=True):
                    show_programs_streams_breakdown(df_for_stats)

            with stat_c4:
                if st.button(f"{s_span}\nOVERALL SPAN", key="stat_btn_span", use_container_width=True):
                    show_span_breakdown(df_for_stats)

            # ==========================================
            # 📅 EXAMINATION SCHEDULE BREAKDOWN (DISPLAY)
            # ==========================================
            st.markdown("### 📅 Examination Schedule Breakdown")

            col1, col2 = st.columns(2)
            with col1:
                st.markdown(f"""
                <div class="metric-card" style="text-align: left; padding: 1rem;">
                    <h4 style="margin: 0; color: white;">📖 Non-Elective Exams</h4>
                    <p style="margin: 0.5rem 0 0 0; font-size: 1rem; opacity: 0.9;">
                        {non_elec_display} <span style="opacity: 0.8; font-size: 0.9em;">({ne_count} Exams)</span>
                    </p>
                </div>
                """, unsafe_allow_html=True)

            with col2:
                st.markdown(f"""
                <div class="metric-card" style="text-align: left; padding: 1rem;">
                    <h4 style="margin: 0; color: white;">🎓 Open Elective (OE) Exams</h4>
                    <p style="margin: 0.5rem 0 0 0; font-size: 1rem; opacity: 0.9;">
                        {oe_display} <span style="opacity: 0.8; font-size: 0.9em;">({oe_count} Exams)</span>
                    </p>
                </div>
                """, unsafe_allow_html=True)
            
            st.markdown("---")

        if st.button("🔄 Convert to PDF", type="primary", use_container_width=True):
            with st.spinner("Converting your Excel file to PDF... Please wait..."):
                try:
                    # Read the Excel file
                    df = read_verification_excel(uploaded_file)
                    
                    if df is not None:
                        # Create Excel sheets format for PDF
                        excel_data = create_excel_sheets_for_pdf(df)
                        
                        if excel_data:
                            # Generate PDF with Declaration Date
                            temp_pdf_path = "temp_timetable_conversion.pdf"
                            
                            if generate_pdf_from_excel_data(excel_data, temp_pdf_path, declaration_date=declaration_date):
                                # Read the generated PDF
                                if os.path.exists(temp_pdf_path):
                                    with open(temp_pdf_path, "rb") as f:
                                        st.session_state.pdf_data = f.read()
                                    os.remove(temp_pdf_path)
                                    
                                    st.session_state.processing_complete = True
                                    
                                    st.markdown('<div class="status-success">🎉 PDF conversion completed successfully!</div>',
                                              unsafe_allow_html=True)
                                else:
                                    st.error("PDF file was not created successfully")
                            else:
                                st.error("Failed to generate PDF")
                        else:
                            st.error("No valid data found for PDF generation")
                    else:
                        st.error("Failed to read Excel file. Please check the format and required columns.")
                        
                except Exception as e:
                    st.error(f"An error occurred during conversion: {str(e)}")
                    import traceback
                    st.error(f"Full error details: {traceback.format_exc()}")

    # Display results and download option
    if st.session_state.processing_complete and st.session_state.pdf_data:
        st.markdown("---")
        
        st.markdown("### 📥 Download PDF")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.download_button(
                label="📄 Download PDF Timetable",
                data=st.session_state.pdf_data,
                file_name=f"timetable_converted_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                mime="application/pdf",
                use_container_width=True
            )
        
        with col2:
            if st.button("🔄 Convert Another File", use_container_width=True):
                st.session_state.processing_complete = False
                st.session_state.pdf_data = None
                st.rerun()
        
        with col3:
            if st.session_state.pdf_data:
                pdf_size = len(st.session_state.pdf_data) / 1024
                st.metric("PDF Size", f"{pdf_size:.1f} KB")
        
    
    # Instructions and help
    st.markdown("---")
    st.markdown("### 📋 Instructions")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        #### 📌 Required Excel Columns:
        - **Program**: B TECH, M TECH, etc.
        - **Stream**: IT, COMPUTER, etc. 
        - **Current Session**: Sem 1, Sem 2, etc.
        - **Module Description**: Subject name
        - **Exam Date**: Date format (DD-MM-YYYY)
        - **Configured Slot**: Exam time slot
        """)
    
    with col2:
        st.markdown("""
        #### ✨ Optional Columns:
        - **Module Abbreviation**: Subject code
        - **CM Group**: Common module group number
        - **Exam Slot Number**: Slot identifier
        - **Student count**: Number of students
        - **Campus**: Campus name
        - **Subject Type**: OE for Open Electives
        """)

    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; padding: 2rem; color: #666;">
        <p><strong>📄 Excel to PDF Timetable Converter</strong></p>
        <p>Convert verification Excel files to professionally formatted PDF timetables</p>
        <p style="font-size: 0.9em;">Direct conversion • Professional formatting • Custom branding • Automatic organization • OE support</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
