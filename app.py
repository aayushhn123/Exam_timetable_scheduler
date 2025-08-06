import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, date
from fpdf import FPDF
import os
import re
import random
import io
from PyPDF2 import PdfReader, PdfWriter
from collections import deque, defaultdict

# Set page configuration
st.set_page_config(
    page_title="Exam Timetable Generator",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded"
)

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

    .stats-section {
        padding: 1.5rem;
        border-radius: 10px;
        margin: 1rem 0;
    }

    /* Updated metric card with icons */
    .metric-card {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        padding: 1.5rem;
        border-radius: 8px;
        color: white;
        text-align: center;
        margin: 0.5rem;
        transition: transform 0.2s;
    }

    .metric-card:hover {
        transform: scale(1.05);
    }

    .metric-card h3 {
        margin: 0;
        font-size: 1.8rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }

    .metric-card p {
        margin: 0.3rem 0 0 0;
        font-size: 1rem;
        opacity: 0.9;
    }

    /* Add gap between difficulty selector and holiday collapsible menu */
    .stCheckbox + .stExpander {
        margin-top: 2rem;
    }

    /* Button hover animations for regular buttons */
    .stButton>button {
        transition: all 0.3s ease;
        border-radius: 5px;
        border: 1px solid transparent;
    }

    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2);
        border: 1px solid #951C1C;
        background-color: #C73E1D;
        color: white;
    }

    /* Download button hover effects (aligned with regular buttons) */
    .stDownloadButton>button {
        transition: all 0.3s ease;
        border-radius: 5px;
        border: 1px solid transparent;
    }

    .stDownloadButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2);
        border: 1px solid #951C1C;
        background-color: #C73E1D;
        color: white;
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

        .stats-section {
            background: #f8f9fa;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        }

        .status-success {
            background: #d4edda;
            color: #155724;
            padding: 1rem;
            border-radius: 5px;
            border-left: 4px solid #28a745;
        }

        .status-error {
            background: #f8d7da;
            color: #721c24;
            padding: 1rem;
            border-radius: 5px;
            border-left: 4px solid #dc3545;
        }

        .status-info {
            background: #d1ecf1;
            color: #0c5460;
            padding: 1rem;
            border-radius: 5px;
            border-left: 4px solid #17a2b8;
        }

        .feature-card {
            background: white;
            padding: 1.5rem;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            margin: 1rem 0;
            border-left: 4px solid #951C1C;
        }

        .metric-card {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        }

        .footer {
            text-align: center;
            color: #666;
            padding: 2rem;
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

        .stats-section {
            background: #333;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
        }

        .status-success {
            background: #2d4b2d;
            color: #e6f4ea;
            padding: 1rem;
            border-radius: 5px;
            border-left: 4px solid #4caf50;
        }

        .status-error {
            background: #4b2d2d;
            color: #f8d7da;
            padding: 1rem;
            border-radius: 5px;
            border-left: 4px solid #f44336;
        }

        .status-info {
            background: #2d4b4b;
            color: #d1ecf1;
            padding: 1rem;
            border-radius: 5px;
            border-left: 4px solid #00bcd4;
        }

        .feature-card {
            background: #333;
            padding: 1.5rem;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.3);
            margin: 1rem 0;
            border-left: 4px solid #A23217;
        }

        .metric-card {
            background: linear-gradient(135deg, #4a5db0 0%, #5a3e8a 100%);
        }

        .footer {
            text-align: center;
            color: #ccc;
            padding: 2rem;
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
    "MCA": "MASTER OF COMPUTER APPLICATIONS"
}

# Define logo path (ensure 'logo.png' is in the same directory as the script)
LOGO_PATH = "logo.png"

# Cache for text wrapping results
wrap_text_cache = {}

class RealTimeOptimizer:
    """Handles real-time optimization during scheduling"""
    
    def __init__(self, branches, holidays, time_slots=None):
        self.branches = branches
        self.holidays = holidays
        self.time_slots = time_slots or ["10:00 AM - 1:00 PM", "2:00 PM - 5:00 PM"]
        self.schedule_grid = {}  # date -> time_slot -> branch -> subject/None
        self.optimization_log = []
        self.moves_made = 0
        
    def add_exam_to_grid(self, date_str, time_slot, branch, subject):
        """Add an exam to the schedule grid"""
        if date_str not in self.schedule_grid:
            self.schedule_grid[date_str] = {}
        if time_slot not in self.schedule_grid[date_str]:
            self.schedule_grid[date_str][time_slot] = {}
        self.schedule_grid[date_str][time_slot][branch] = subject
    
    def find_earliest_empty_slot(self, branch, start_date, preferred_time_slot=None):
        """Find the earliest empty slot for a branch - ensuring only one exam per day per branch"""
        sorted_dates = sorted(self.schedule_grid.keys(), 
                            key=lambda x: datetime.strptime(x, "%d-%m-%Y"))
        
        for date_str in sorted_dates:
            date_obj = datetime.strptime(date_str, "%d-%m-%Y")
            
            if date_obj < start_date:
                continue
            
            if date_obj.weekday() == 6 or date_obj.date() in self.holidays:
                continue
            
            branch_has_exam_today = False
            if date_str in self.schedule_grid:
                for time_slot in self.time_slots:
                    if (time_slot in self.schedule_grid[date_str] and
                        branch in self.schedule_grid[date_str][time_slot] and
                        self.schedule_grid[date_str][time_slot][branch] is not None):
                        branch_has_exam_today = True
                        break
            
            if branch_has_exam_today:
                continue
            
            if preferred_time_slot:
                if (date_str in self.schedule_grid and 
                    preferred_time_slot in self.schedule_grid[date_str]):
                    if self.schedule_grid[date_str][preferred_time_slot].get(branch) is None:
                        return date_str, preferred_time_slot
            
            for time_slot in self.time_slots:
                if date_str not in self.schedule_grid:
                    return date_str, time_slot
                    
                if time_slot not in self.schedule_grid[date_str]:
                    return date_str, time_slot
                    
                if self.schedule_grid[date_str][time_slot].get(branch) is None:
                    return date_str, time_slot
        
        return None, None
    
    def initialize_grid_with_empty_days(self, start_date, num_days=40):
        """Pre-populate grid with empty days"""
        current_date = start_date
        for _ in range(num_days):
            if current_date.weekday() != 6 and current_date.date() not in self.holidays:
                date_str = current_date.strftime("%d-%m-%Y")
                if date_str not in self.schedule_grid:
                    self.schedule_grid[date_str] = {}
                for time_slot in self.time_slots:
                    if time_slot not in self.schedule_grid[date_str]:
                        self.schedule_grid[date_str][time_slot] = {branch: None for branch in self.branches}
            current_date += timedelta(days=1)
    
    def get_schedule_summary(self):
        """Get a summary of the current schedule"""
        total_slots = 0
        filled_slots = 0
        
        for date_str, time_slots in self.schedule_grid.items():
            for time_slot, branches in time_slots.items():
                for branch, subject in branches.items():
                    total_slots += 1
                    if subject is not None:
                        filled_slots += 1
        
        return {
            'total_slots': total_slots,
            'filled_slots': filled_slots,
            'empty_slots': total_slots - filled_slots,
            'utilization': (filled_slots / total_slots * 100) if total_slots > 0 else 0
        }

def wrap_text(pdf, text, col_width):
    """Caches text wrapping results to improve performance."""
    cache_key = (text, col_width)
    if cache_key in wrap_text_cache:
        return wrap_text_cache[cache_key]
    words = text.split()
    lines = []
    current_line = ""
    for word in words:
        test_line = word if not current_line else current_line + " " + word
        if pdf.get_string_width(test_line) <= col_width:
            current_line = test_line
        else:
            lines.append(current_line)
            current_line = word
    if current_line:
        lines.append(current_line)
    wrap_text_cache[cache_key] = lines
    return lines

def print_row_custom(pdf, row_data, col_widths, line_height=5, header=False):
    """Prints a single row in a PDF table with custom styling."""
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
            pdf.cell(col_widths[i] - 2 * cell_padding, line_height, ln, border=0, align='C')
        pdf.rect(cx, y0, col_widths[i], row_h)
        pdf.set_xy(cx + col_widths[i], y0)

    setattr(pdf, '_row_counter', row_number + 1)
    pdf.set_xy(x0, y0 + row_h)

def print_table_custom(pdf, df, columns, col_widths, line_height=5, header_content=None, branches=None, time_slot=None):
    """Prints a complete table to the PDF, handling page breaks."""
    if df.empty:
        return
    setattr(pdf, '_row_counter', 0)
    
    # Add footer first
    footer_height = 25
    pdf.set_xy(10, pdf.h - footer_height)
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 5, "Controller of Examinations", 0, 1, 'L')
    pdf.line(10, pdf.h - footer_height + 5, 60, pdf.h - footer_height + 5)
    pdf.set_font("Arial", size=13)
    pdf.set_xy(10, pdf.h - footer_height + 7)
    pdf.cell(0, 5, "Signature", 0, 1, 'L')
    
    # Add page numbers in bottom right
    pdf.set_font("Arial", size=14)
    pdf.set_text_color(0, 0, 0)
    page_text = f"{pdf.page_no()} of {{nb}}"
    text_width = pdf.get_string_width(page_text.replace("{nb}", "99"))  # Estimate width
    pdf.set_xy(pdf.w - 10 - text_width, pdf.h - footer_height + 12)
    pdf.cell(text_width, 5, page_text, 0, 0, 'R')
    
    # Add header
    header_height = 85
    pdf.set_y(0)
    current_date = datetime.now().strftime("%A, %B %d, %Y, %I:%M %p IST")
    pdf.set_font("Arial", size=14)
    text_width = pdf.get_string_width(current_date)
    x = pdf.w - 10 - text_width
    pdf.set_xy(x, 5)
    pdf.cell(text_width, 10, f"Generated on: {current_date}", 0, 0, 'R')
    if os.path.exists(LOGO_PATH):
        logo_width = 45
        logo_x = (pdf.w - logo_width) / 2
        pdf.image(LOGO_PATH, x=logo_x, y=10, w=logo_width)
    pdf.set_fill_color(149, 33, 28)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", 'B', 16)
    pdf.rect(10, 30, pdf.w - 20, 14, 'F')
    pdf.set_xy(10, 30)
    pdf.cell(pdf.w - 20, 14,
             "MUKESH PATEL SCHOOL OF TECHNOLOGY MANAGEMENT & ENGINEERING / SCHOOL OF TECHNOLOGY MANAGEMENT & ENGINEERING",
             0, 1, 'C')
    pdf.set_font("Arial", 'B', 15)
    pdf.set_text_color(0, 0, 0)
    pdf.set_xy(10, 51)
    pdf.cell(pdf.w - 20, 8, f"{header_content['main_branch_full']} - Semester {header_content['semester_roman']}", 0, 1, 'C')
    
    if time_slot:
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
    
    available_height = pdf.h - pdf.t_margin - footer_height - header_height
    pdf.set_font("Arial", size=12)
    
    print_row_custom(pdf, columns, col_widths, line_height=line_height, header=True)
    
    for idx in range(len(df)):
        row = [str(df.iloc[idx][c]) if pd.notna(df.iloc[idx][c]) else "" for c in columns]
        if not any(cell.strip() for cell in row):
            continue
            
        wrapped_cells = []
        max_lines = 0
        for i, cell_text in enumerate(row):
            text = str(cell_text) if cell_text is not None else ""
            avail_w = col_widths[i] - 4 
            lines = wrap_text(pdf, text, avail_w)
            wrapped_cells.append(lines)
            max_lines = max(max_lines, len(lines))
        row_h = line_height * max_lines
        
        if pdf.get_y() + row_h > pdf.h - footer_height:
            add_footer_with_page_number(pdf, footer_height)
            pdf.add_page()
            add_footer_with_page_number(pdf, footer_height)
            add_header_to_page(pdf, current_date, logo_x, logo_width, header_content, branches, time_slot)
            pdf.set_font("Arial", size=12)
            print_row_custom(pdf, columns, col_widths, line_height=line_height, header=True)
        
        print_row_custom(pdf, row, col_widths, line_height=line_height, header=False)

def add_footer_with_page_number(pdf, footer_height):
    """Add footer with signature and page number"""
    pdf.set_xy(10, pdf.h - footer_height)
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 5, "Controller of Examinations", 0, 1, 'L')
    pdf.line(10, pdf.h - footer_height + 5, 60, pdf.h - footer_height + 5)
    pdf.set_font("Arial", size=13)
    pdf.set_xy(10, pdf.h - footer_height + 7)
    pdf.cell(0, 5, "Signature", 0, 1, 'L')
    
    pdf.set_font("Arial", size=14)
    pdf.set_text_color(0, 0, 0)
    page_text = f"{pdf.page_no()} of {{nb}}"
    text_width = pdf.get_string_width(page_text.replace("{nb}", "99"))
    pdf.set_xy(pdf.w - 10 - text_width, pdf.h - footer_height + 12)
    pdf.cell(text_width, 5, page_text, 0, 0, 'R')

def add_header_to_page(pdf, current_date, logo_x, logo_width, header_content, branches, time_slot=None):
    """Add header to a new page"""
    pdf.set_y(0)
    pdf.set_font("Arial", size=14)
    text_width = pdf.get_string_width(current_date)
    x = pdf.w - 10 - text_width
    pdf.set_xy(x, 5)
    pdf.cell(text_width, 10, f"Generated on: {current_date}", 0, 0, 'R')
    if os.path.exists(LOGO_PATH):
        pdf.image(LOGO_PATH, x=logo_x, y=10, w=logo_width)
    pdf.set_fill_color(149, 33, 28)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", 'B', 16)
    pdf.rect(10, 30, pdf.w - 20, 14, 'F')
    pdf.set_xy(10, 30)
    pdf.cell(pdf.w - 20, 14,
             "MUKESH PATEL SCHOOL OF TECHNOLOGY MANAGEMENT & ENGINEERING / SCHOOL OF TECHNOLOGY MANAGEMENT & ENGINEERING",
             0, 1, 'C')
    pdf.set_font("Arial", 'B', 15)
    pdf.set_text_color(0, 0, 0)
    pdf.set_xy(10, 51)
    pdf.cell(pdf.w - 20, 8, f"{header_content['main_branch_full']} - Semester {header_content['semester_roman']}", 0, 1, 'C')
    
    if time_slot:
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

def calculate_end_time(start_time, duration_hours):
    """Calculate the end time given a start time and duration in hours."""
    start = datetime.strptime(start_time, "%I:%M %p")
    duration = timedelta(hours=duration_hours)
    end = start + duration
    return end.strftime("%I:%M %p").replace("AM", "am").replace("PM", "pm")

def convert_excel_to_pdf(excel_path, pdf_path, sub_branch_cols_per_page=4):
    """Converts a formatted Excel file to a styled PDF."""
    pdf = FPDF(orientation='L', unit='mm', format=(210, 500))
    pdf.set_auto_page_break(auto=False, margin=15)
    pdf.alias_nb_pages()
    
    df_dict = pd.read_excel(excel_path, sheet_name=None, index_col=[0, 1])

    def int_to_roman(num):
        val = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
        syb = ["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"]
        roman_num = ''
        i = 0
        while num > 0:
            for _ in range(num // val[i]):
                roman_num += syb[i]
                num -= val[i]
            i += 1
        return roman_num

    def extract_duration(subject_str):
        match = re.search(r'\[Duration: (\d+\.?\d*) hrs\]', subject_str)
        return float(match.group(1)) if match else 3.0

    def get_semester_default_time_slot(semester, branch):
        return "10:00 AM - 1:00 PM"

    for sheet_name, pivot_df in df_dict.items():
        if pivot_df.empty:
            continue
        parts = sheet_name.split('_Sem_')
        main_branch = parts[0]
        main_branch_full = BRANCH_FULL_FORM.get(main_branch, main_branch)
        semester = parts[1] if len(parts) > 1 else ""
        semester_roman = semester if not semester.isdigit() else int_to_roman(int(semester))
        header_content = {'main_branch_full': main_branch_full, 'semester_roman': semester_roman}

        if not sheet_name.endswith('_Electives'):
            pivot_df = pivot_df.reset_index().dropna(how='all', axis=0).reset_index(drop=True)
            fixed_cols = ["Exam Date", "Time Slot"]
            sub_branch_cols = [c for c in pivot_df.columns if c not in fixed_cols]
            exam_date_width = 60
            line_height = 10

            for start in range(0, len(sub_branch_cols), sub_branch_cols_per_page):
                chunk = sub_branch_cols[start:start + sub_branch_cols_per_page]
                cols_to_print = fixed_cols[:1] + chunk
                chunk_df = pivot_df[fixed_cols + chunk].copy()
                mask = chunk_df[chunk].apply(lambda row: row.astype(str).str.strip() != "").any(axis=1)
                chunk_df = chunk_df[mask].reset_index(drop=True)
                if chunk_df.empty:
                    continue

                time_slot = pivot_df['Time Slot'].iloc[0] if 'Time Slot' in pivot_df.columns and not pivot_df['Time Slot'].empty else None
                chunk_df["Exam Date"] = pd.to_datetime(chunk_df["Exam Date"], format="%d-%m-%Y", errors='coerce').dt.strftime("%A, %d %B, %Y")
                default_time_slot = get_semester_default_time_slot(semester, main_branch)

                for sub_branch in chunk:
                    for idx in chunk_df.index:
                        cell_value = chunk_df.at[idx, sub_branch]
                        if pd.isna(cell_value) or cell_value.strip() == "---":
                            continue
                        subjects = cell_value.split(", ")
                        modified_subjects = []
                        for subject in subjects:
                            duration = extract_duration(subject)
                            base_subject = re.sub(r' \[Duration: \d+\.?\d* hrs\]', '', subject)
                            subject_time_slot = chunk_df.at[idx, "Time Slot"] if pd.notna(chunk_df.at[idx, "Time Slot"]) else None
                            
                            if duration != 3 and subject_time_slot and subject_time_slot.strip():
                                start_time = subject_time_slot.split(" - ")[0]
                                end_time = calculate_end_time(start_time, duration)
                                modified_subjects.append(f"{base_subject} ({start_time} to {end_time})")
                            else:
                                if subject_time_slot and default_time_slot and subject_time_slot.strip() and default_time_slot.strip():
                                    if subject_time_slot != default_time_slot:
                                        modified_subjects.append(f"{base_subject} ({subject_time_slot})")
                                    else:
                                        modified_subjects.append(base_subject)
                                else:
                                    modified_subjects.append(base_subject)
                        
                        chunk_df.at[idx, sub_branch] = ", ".join(modified_subjects)

                page_width = pdf.w - 2 * pdf.l_margin
                remaining = page_width - exam_date_width
                sub_width = remaining / max(len(chunk), 1)
                col_widths = [exam_date_width] + [sub_width] * len(chunk)
                
                pdf.add_page()
                footer_height = 25
                add_footer_with_page_number(pdf, footer_height)
                print_table_custom(pdf, chunk_df[cols_to_print], cols_to_print, col_widths, line_height=line_height, 
                                 header_content=header_content, branches=chunk, time_slot=time_slot)

        if sheet_name.endswith('_Electives'):
            pivot_df = pivot_df.reset_index().dropna(how='all', axis=0).reset_index(drop=True)
            elective_data = pivot_df.groupby(['OE', 'Exam Date', 'Time Slot']).agg({
                'SubjectDisplay': lambda x: ", ".join(x)
            }).reset_index()

            elective_data["Exam Date"] = pd.to_datetime(elective_data["Exam Date"], format="%d-%m-%Y", errors='coerce').dt.strftime("%A, %d %B, %Y")
            elective_data['SubjectDisplay'] = elective_data.apply(
                lambda row: ", ".join([s.replace(f" [{row['OE']}]", "") for s in row['SubjectDisplay'].split(", ")]),
                axis=1
            )
            default_time_slot = get_semester_default_time_slot(semester, main_branch)
            time_slot = pivot_df['Time Slot'].iloc[0] if 'Time Slot' in pivot_df.columns and not pivot_df['Time Slot'].empty else None

            for idx in elective_data.index:
                cell_value = elective_data.at[idx, 'SubjectDisplay']
                if pd.isna(cell_value) or cell_value.strip() == "---":
                    continue
                subjects = cell_value.split(", ")
                modified_subjects = []
                for subject in subjects:
                    duration = extract_duration(subject)
                    base_subject = re.sub(r' \[Duration: \d+\.?\d* hrs\]', '', subject)
                    subject_time_slot = elective_data.at[idx, "Time Slot"] if pd.notna(elective_data.at[idx, "Time Slot"]) else None
                    
                    if duration != 3 and subject_time_slot and subject_time_slot.strip():
                        start_time = subject_time_slot.split(" - ")[0]
                        end_time = calculate_end_time(start_time, duration)
                        modified_subjects.append(f"{base_subject} ({start_time} to {end_time})")
                    else:
                        if subject_time_slot and default_time_slot and subject_time_slot.strip() and default_time_slot.strip():
                            if subject_time_slot != default_time_slot:
                                modified_subjects.append(f"{base_subject} ({subject_time_slot})")
                            else:
                                modified_subjects.append(base_subject)
                        else:
                            modified_subjects.append(base_subject)
                
                elective_data.at[idx, 'SubjectDisplay'] = ", ".join(modified_subjects)

            elective_data = elective_data.rename(columns={'OE': 'OE Type', 'SubjectDisplay': 'Subjects'})
            exam_date_width = 60
            oe_width = 30
            subject_width = pdf.w - 2 * pdf.l_margin - exam_date_width - oe_width
            col_widths = [exam_date_width, oe_width, subject_width]
            cols_to_print = ['Exam Date', 'OE Type', 'Subjects']
            
            pdf.add_page()
            footer_height = 25
            add_footer_with_page_number(pdf, footer_height)
            print_table_custom(pdf, elective_data, cols_to_print, col_widths, line_height=10, 
                             header_content=header_content, branches=['All Streams'], time_slot=time_slot)

    pdf.output(pdf_path)
    
def generate_pdf_timetable(semester_wise_timetable, output_pdf):
    """Orchestrates the creation of the PDF file from timetable data."""
    temp_excel = os.path.join(os.path.dirname(output_pdf), "temp_timetable.xlsx")
    excel_data = save_to_excel(semester_wise_timetable)
    if excel_data:
        with open(temp_excel, "wb") as f:
            f.write(excel_data.getvalue())
        convert_excel_to_pdf(temp_excel, output_pdf)
        if os.path.exists(temp_excel):
            os.remove(temp_excel)
    else:
        st.error("No data to save to Excel.")
        return
    try:
        reader = PdfReader(output_pdf)
        writer = PdfWriter()
        page_number_pattern = re.compile(r'^[\s\n]*(?:Page\s*)?\d+[\s\n]*$')
        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]
            try:
                text = page.extract_text() if page else ""
            except:
                text = ""
            cleaned_text = text.strip() if text else ""
            is_blank_or_page_number = (
                    not cleaned_text or
                    page_number_pattern.match(cleaned_text) or
                    len(cleaned_text) <= 10
            )
            if not is_blank_or_page_number:
                writer.add_page(page)
        if len(writer.pages) > 0:
            with open(output_pdf, 'wb') as output_file:
                writer.write(output_file)
        else:
            st.warning("Warning: All pages were filtered out - keeping original PDF")
    except Exception as e:
        st.error(f"Error during PDF post-processing: {str(e)}")

def read_timetable(uploaded_file):
    """Reads and preprocesses the uploaded Excel file."""
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        df = df.rename(columns={
            "Program": "Program", "Stream": "Stream", "Current Session": "Semester",
            "Module Description": "SubjectName", "Module Abbreviation": "ModuleCode",
            "Campus Name": "Campus", "Difficulty Score": "Difficulty",
            "Exam Duration": "Exam Duration", "Student count": "StudentCount",
            "Is Common": "IsCommon"
        })
        
        def convert_sem(sem):
            if pd.isna(sem): return 0
            m = { "Sem I": 1, "Sem II": 2, "Sem III": 3, "Sem IV": 4, "Sem V": 5, 
                  "Sem VI": 6, "Sem VII": 7, "Sem VIII": 8, "Sem IX": 9, "Sem X": 10, "Sem XI": 11 }
            return m.get(sem.strip(), 0)
        
        df["Semester"] = df["Semester"].apply(convert_sem).astype(int)
        df["Branch"] = df["Program"].astype(str).str.strip() + "-" + df["Stream"].astype(str).str.strip()
        df["Subject"] = df["SubjectName"].astype(str) + " - (" + df["ModuleCode"].astype(str) + ")"
        
        df["Exam Date"] = ""
        df["Time Slot"] = ""
        df["Exam Duration"] = df["Exam Duration"].fillna(3).astype(float)
        df["StudentCount"] = df["StudentCount"].fillna(0).astype(int)
        df["IsCommon"] = df["IsCommon"].fillna("NO").str.strip().str.upper()
        
        df_non = df[df["Category"] != "INTD"].copy()
        df_ele = df[df["Category"] == "INTD"].copy()
        
        def split_br(b):
            p = b.split("-", 1)
            return pd.Series([p[0].strip(), p[1].strip() if len(p) > 1 else ""])
        
        for d in (df_non, df_ele):
            d[["MainBranch", "SubBranch"]] = d["Branch"].apply(split_br)
        
        cols = ["MainBranch", "SubBranch", "Branch", "Semester", "Subject", "Category", "OE", "Exam Date", "Time Slot",
                "Difficulty", "Exam Duration", "StudentCount", "IsCommon", "ModuleCode"]
        
        return df_non[cols], df_ele[cols], df
        
    except Exception as e:
        st.error(f"Error reading the Excel file: {str(e)}")
        return None, None, None

def parse_date_safely(date_input, input_format="%d-%m-%Y"):
    """Safely parse date input ensuring DD-MM-YYYY format interpretation."""
    if pd.isna(date_input): return None
    if isinstance(date_input, pd.Timestamp): return date_input
    if isinstance(date_input, str):
        try:
            return pd.to_datetime(date_input, format=input_format, errors='raise')
        except:
            try:
                return pd.to_datetime(date_input, dayfirst=True, errors='raise')
            except:
                return None
    return pd.to_datetime(date_input, errors='coerce')

def schedule_semester_non_electives_with_optimization(df_sem, holidays, base_date, exam_days, optimizer, schedule_by_difficulty=False):
    """Schedules non-elective subjects for a semester, ensuring no branch has two exams on the same day."""
    sem = df_sem["Semester"].iloc[0]
    preferred_slot = "10:00 AM - 1:00 PM" if ((sem + 1) // 2) % 2 == 1 else "2:00 PM - 5:00 PM"

    subjects_to_schedule = df_sem[(df_sem['IsCommon'] == 'NO') & (df_sem['Exam Date'] == "")]

    for idx, row in subjects_to_schedule.iterrows():
        branch = row['Branch']
        subject = row['Subject']
        current_date = base_date
        scheduled = False
        max_attempts = 100
        
        for _ in range(max_attempts):
            date_str = current_date.strftime("%d-%m-%Y")
            
            if current_date.weekday() == 6 or current_date.date() in holidays:
                current_date += timedelta(days=1)
                continue
            
            branch_has_exam_today = current_date.date() in exam_days.get(branch, set())
            
            grid_conflict = False
            if date_str in optimizer.schedule_grid and preferred_slot in optimizer.schedule_grid[date_str]:
                if optimizer.schedule_grid[date_str][preferred_slot].get(branch) is not None:
                    grid_conflict = True

            df_conflict = ((df_sem['Branch'] == branch) & (df_sem['Exam Date'] == date_str)).any()

            if not branch_has_exam_today and not grid_conflict and not df_conflict:
                df_sem.at[idx, 'Exam Date'] = date_str
                df_sem.at[idx, 'Time Slot'] = preferred_slot
                optimizer.add_exam_to_grid(date_str, preferred_slot, branch, subject)
                exam_days[branch].add(current_date.date())
                scheduled = True
                optimizer.optimization_log.append(f"✅ Scheduled {row['Category']} {subject} for {branch} on {date_str}")
                optimizer.moves_made += 1
                break
            else:
                current_date += timedelta(days=1)
        
        if not scheduled:
            st.error(f"❌ Could not schedule {row['Category']} subject {subject} for {branch} after {max_attempts} attempts")
    
    df_sem.loc[df_sem['Time Slot'] == "", 'Time Slot'] = preferred_slot
    return df_sem

def process_constraints_with_real_time_optimization(df, holidays, base_date, schedule_by_difficulty=False):
    """Main function to process all scheduling constraints and run optimizations."""
    all_branches = df['Branch'].unique()
    exam_days = {branch: set() for branch in all_branches}
    optimizer = RealTimeOptimizer(all_branches, holidays)
    optimizer.initialize_grid_with_empty_days(base_date, num_days=50)

    def find_earliest_available_slot_with_one_exam_per_day(start_day, for_branches, subject):
        current_date = start_day
        while True:
            if current_date.weekday() == 6 or current_date.date() in holidays:
                current_date += timedelta(days=1)
                continue
            if all(current_date.date() not in exam_days.get(branch, set()) for branch in for_branches):
                return current_date
            current_date += timedelta(days=1)

    common_subjects = df[df['IsCommon'] == 'YES']
    for module_code, group in common_subjects.groupby('ModuleCode'):
        branches = group['Branch'].unique()
        subject = group['Subject'].iloc[0]
        exam_day = find_earliest_available_slot_with_one_exam_per_day(base_date, branches, subject)
        min_sem = group['Semester'].min()
        slot_str = "10:00 AM - 1:00 PM" if ((min_sem + 1) // 2) % 2 == 1 else "2:00 PM - 5:00 PM"
        date_str = exam_day.strftime("%d-%m-%Y")
        
        df.loc[group.index, 'Exam Date'] = date_str
        df.loc[group.index, 'Time Slot'] = slot_str
        
        for branch in branches:
            exam_days[branch].add(exam_day.date())
            optimizer.add_exam_to_grid(date_str, slot_str, branch, subject)

    final_list = []
    for sem in sorted(df["Semester"].unique()):
        if sem == 0: continue
        df_sem = df[df["Semester"] == sem].copy()
        if df_sem.empty: continue
        
        scheduled_sem = schedule_semester_non_electives_with_optimization(
            df_sem, holidays, base_date, exam_days, optimizer, schedule_by_difficulty
        )
        final_list.append(scheduled_sem)

    if not final_list: return {}

    df_combined = pd.concat(final_list, ignore_index=True)
    df_combined_clean = df_combined.drop_duplicates(subset=['Branch', 'Exam Date', 'Subject', 'ModuleCode', 'Semester'])
    
    # Validation
    validation_check = df_combined_clean.groupby(['Branch', 'Exam Date', 'Semester']).size()
    multiple_exams_same_day = validation_check[validation_check > 1]
    if not multiple_exams_same_day.empty:
        st.error("❌ VALIDATION FAILED: Found cases where branches have multiple exams on the same day within the same semester!")
    else:
        st.success("✅ VALIDATION PASSED: No branch has multiple exams on the same day within any semester!")

    if df_combined_clean[df_combined_clean['Exam Date'] == ""].any().any():
        st.error("❌ Some subjects remain unscheduled!")

    sem_dict = {sem: df_combined_clean[df_combined_clean["Semester"] == sem].copy() for sem in sorted(df_combined_clean["Semester"].unique())}
    return sem_dict

def find_next_valid_day_for_electives(start_day, holidays):
    """Find the next valid day for scheduling electives (skip weekends and holidays)"""
    day = start_day
    while True:
        if day.weekday() == 6 or day.date() in holidays:
            day += timedelta(days=1)
            continue
        return day

def optimize_oe_subjects_after_scheduling(sem_dict, holidays):
    """After main scheduling, check if OE subjects can be moved to earlier empty slots."""
    if not sem_dict: return sem_dict
    
    all_data = pd.concat(sem_dict.values(), ignore_index=True)
    
    def normalize_date_to_ddmmyyyy(date_val):
        if pd.isna(date_val) or date_val == "": return ""
        try:
            return pd.to_datetime(date_val, dayfirst=True).strftime("%d-%m-%Y")
        except:
            return str(date_val)
    
    all_data['Exam Date'] = all_data['Exam Date'].apply(normalize_date_to_ddmmyyyy)
    
    oe_data = all_data[all_data['OE'].notna() & (all_data['OE'].str.strip() != "")]
    if oe_data.empty: return sem_dict
    
    # This is a placeholder for a more complex optimization logic.
    # For now, it just ensures the dates are correctly formatted and returns the dictionary.
    st.info("ℹ️ OE subject placement optimization is a complex step. The current version ensures they are scheduled after core subjects.")

    return sem_dict
    
def save_to_excel(semester_wise_timetable):
    """Saves the final timetable to a multi-sheet Excel file."""
    if not semester_wise_timetable: return None

    def int_to_roman(num):
        val = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
        syb = ["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"]
        roman_num = ''
        i = 0
        while num > 0:
            for _ in range(num // val[i]):
                roman_num += syb[i]
                num -= val[i]
            i += 1
        return roman_num

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sem, df_sem in semester_wise_timetable.items():
            for main_branch in df_sem["MainBranch"].unique():
                df_mb = df_sem[df_sem["MainBranch"] == main_branch].copy()
                df_non_elec = df_mb[df_mb['OE'].isna() | (df_mb['OE'].str.strip() == "")].copy()
                df_elec = df_mb[df_mb['OE'].notna() & (df_mb['OE'].str.strip() != "")].copy()

                if not df_non_elec.empty:
                    duration_suffix = df_non_elec.apply(lambda r: f" [Duration: {r['Exam Duration']} hrs]" if r['Exam Duration'] != 3 else '', axis=1)
                    df_non_elec["SubjectDisplay"] = df_non_elec["Subject"] + duration_suffix
                    df_non_elec["Exam Date"] = pd.to_datetime(df_non_elec["Exam Date"], format="%d-%m-%Y", errors='coerce')
                    pivot_df = df_non_elec.pivot_table(index=["Exam Date", "Time Slot"], columns="SubBranch", values="SubjectDisplay", aggfunc=lambda x: ", ".join(map(str, x))).fillna("---")
                    pivot_df = pivot_df.sort_index(level="Exam Date")
                    pivot_df.index = pivot_df.index.set_levels(pivot_df.index.levels[0].strftime("%d-%m-%Y"), level=0)
                    sheet_name = f"{main_branch}_Sem_{int_to_roman(sem)}"[:31]
                    pivot_df.to_excel(writer, sheet_name=sheet_name)

                if not df_elec.empty:
                    duration_suffix = df_elec.apply(lambda r: f" [Duration: {r['Exam Duration']} hrs]" if r['Exam Duration'] != 3 else '', axis=1)
                    df_elec["SubjectDisplay"] = df_elec["Subject"] + " [" + df_elec["OE"] + "]" + duration_suffix
                    elec_pivot = df_elec.groupby(['OE', 'Exam Date', 'Time Slot'])['SubjectDisplay'].apply(lambda x: ", ".join(sorted(set(x)))).reset_index()
                    elec_pivot['Exam Date'] = pd.to_datetime(elec_pivot['Exam Date'], format="%d-%m-%Y", errors='coerce').dt.strftime("%d-%m-%Y")
                    sheet_name = f"{main_branch}_Sem_{int_to_roman(sem)}_Electives"[:31]
                    elec_pivot.to_excel(writer, sheet_name=sheet_name, index=False)

    output.seek(0)
    return output
    
def save_verification_excel(original_df, semester_wise_timetable):
    """Saves a verification file mapping scheduled dates back to the original data."""
    if not semester_wise_timetable: return None

    columns_to_retain = [col for col in ["School Name", "Campus", "Program", "Stream", "Current Academic Year",
        "Semester", "ModuleCode", "SubjectName", "Difficulty", "Category", "OE", "Exam mode", "Exam Duration"] if col in original_df.columns]
    verification_df = original_df[columns_to_retain].copy()
    verification_df["Exam Date"] = ""
    verification_df["Exam Time"] = ""

    scheduled_data = pd.concat(semester_wise_timetable.values(), ignore_index=True)
    scheduled_data["ModuleCode"] = scheduled_data["Subject"].str.extract(r'\((.*?)\)', expand=False)

    for idx, row in verification_df.iterrows():
        match = scheduled_data[(scheduled_data["ModuleCode"] == row["ModuleCode"]) & (scheduled_data["Semester"] == row["Semester"])]
        if not match.empty:
            exam_date = match.iloc[0]["Exam Date"]
            time_slot = match.iloc[0]["Time Slot"]
            duration = row["Exam Duration"]
            start_time = time_slot.split(" - ")[0]
            end_time = calculate_end_time(start_time, duration)
            verification_df.at[idx, "Exam Date"] = exam_date
            verification_df.at[idx, "Exam Time"] = f"{start_time} to {end_time}"

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        verification_df.to_excel(writer, sheet_name="Verification", index=False)
    output.seek(0)
    return output

def main():
    """Main function to run the Streamlit application."""
    st.markdown("""
    <div class="main-header">
        <h1>📅 Exam Timetable Generator</h1>
        <p>MUKESH PATEL SCHOOL OF TECHNOLOGY MANAGEMENT & ENGINEERING</p>
    </div>
    """, unsafe_allow_html=True)

    # Initialize session state
    for key in ['timetable_data', 'excel_data', 'pdf_data', 'verification_data', 'original_df']:
        if key not in st.session_state:
            st.session_state[key] = None
    if 'processing_complete' not in st.session_state:
        st.session_state.processing_complete = False

    with st.sidebar:
        st.markdown("### ⚙️ Configuration")
        base_date = st.date_input("Start date for exams", value=date(2025, 4, 1))
        base_date = datetime.combine(base_date, datetime.min.time())

        with st.expander("Holiday Configuration", expanded=True):
            predefined_holidays = {
                "April 14, 2025": date(2025, 4, 14),
                "May 1, 2025": date(2025, 5, 1),
                "August 15, 2025": date(2025, 8, 15)
            }
            holidays_set = set()
            for name, hdate in predefined_holidays.items():
                if st.checkbox(name, value=True):
                    holidays_set.add(hdate)
            
            # Simplified custom holidays
            custom_holidays_str = st.text_area("Add Custom Holidays (DD-MM-YYYY, one per line)", "")
            for h_str in custom_holidays_str.split('\n'):
                if h_str.strip():
                    try:
                        holidays_set.add(datetime.strptime(h_str.strip(), "%d-%m-%Y").date())
                    except ValueError:
                        st.warning(f"Ignoring invalid date format: {h_str}")

    st.markdown("""
    <div class="upload-section">
        <h3>📁 Upload Excel File</h3>
        <p>Upload your timetable data file (.xlsx format)</p>
    </div>
    """, unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])

    if uploaded_file:
        if st.button("🔄 Generate Timetable", type="primary", use_container_width=True):
            with st.spinner("Processing your timetable..."):
                df_non_elec, df_ele, original_df = read_timetable(uploaded_file)
                if df_non_elec is not None:
                    st.session_state.original_df = original_df
                    
                    # Schedule non-electives
                    non_elec_sched = process_constraints_with_real_time_optimization(df_non_elec, holidays_set, base_date)
                    
                    # Schedule electives after non-electives
                    final_df_list = list(non_elec_sched.values())
                    if df_ele is not None and not df_ele.empty:
                        all_non_elec_dates = pd.to_datetime(pd.concat(final_df_list)['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
                        last_day = max(all_non_elec_dates) if not all_non_elec_dates.empty else base_date
                        
                        elective_day1 = find_next_valid_day_for_electives(last_day + timedelta(days=1), holidays_set)
                        elective_day2 = find_next_valid_day_for_electives(elective_day1 + timedelta(days=1), holidays_set)
                        
                        df_ele.loc[df_ele['OE'].isin(['OE1', 'OE5']), 'Exam Date'] = elective_day1.strftime("%d-%m-%Y")
                        df_ele.loc[df_ele['OE'].isin(['OE1', 'OE5']), 'Time Slot'] = "10:00 AM - 1:00 PM"
                        df_ele.loc[df_ele['OE'] == 'OE2', 'Exam Date'] = elective_day2.strftime("%d-%m-%Y")
                        df_ele.loc[df_ele['OE'] == 'OE2', 'Time Slot'] = "2:00 PM - 5:00 PM"
                        final_df_list.append(df_ele)

                    final_df = pd.concat(final_df_list, ignore_index=True)
                    sem_dict = {s: final_df[final_df["Semester"] == s].copy() for s in sorted(final_df["Semester"].unique())}

                    st.session_state.timetable_data = optimize_oe_subjects_after_scheduling(sem_dict, holidays_set)
                    st.session_state.excel_data = save_to_excel(st.session_state.timetable_data)
                    st.session_state.verification_data = save_verification_excel(original_df, st.session_state.timetable_data)
                    
                    temp_pdf_path = "timetable_temp.pdf"
                    generate_pdf_timetable(st.session_state.timetable_data, temp_pdf_path)
                    with open(temp_pdf_path, "rb") as f:
                        st.session_state.pdf_data = f.read()
                    if os.path.exists(temp_pdf_path):
                        os.remove(temp_pdf_path)

                    st.session_state.processing_complete = True
                    st.success("🎉 Timetable generated successfully!")

    if st.session_state.processing_complete:
        st.markdown("---")
        st.markdown("### 📥 Download Options")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.download_button("📊 Download Excel", st.session_state.excel_data.getvalue(), "timetable.xlsx", use_container_width=True)
        with col2:
            st.download_button("📄 Download PDF", st.session_state.pdf_data, "timetable.pdf", use_container_width=True)
        with col3:
            st.download_button("📋 Download Verification", st.session_state.verification_data.getvalue(), "verification.xlsx", use_container_width=True)

        st.markdown("---")
        st.markdown("## 📊 Timetable Results")
        for sem, df_sem in st.session_state.timetable_data.items():
            st.markdown(f"### 📚 Semester {sem}")
            for main_branch in df_sem["MainBranch"].unique():
                st.markdown(f"#### {BRANCH_FULL_FORM.get(main_branch, main_branch)}")
                df_mb = df_sem[df_sem["MainBranch"] == main_branch].copy()
                df_display = df_mb[["SubBranch", "Subject", "Exam Date", "Time Slot", "OE"]].sort_values("Exam Date").reset_index(drop=True)
                st.dataframe(df_display, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("""
    <div class="footer">
        <p>🎓 <strong>Exam Timetable Generator</strong></p>
        <p>Developed for MUKESH PATEL SCHOOL OF TECHNOLOGY MANAGEMENT & ENGINEERING</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
