import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
import io
import os
import re
from fpdf import FPDF
from PyPDF2 import PdfReader, PdfWriter
from collections import defaultdict

# Set page configuration for Streamlit
st.set_page_config(page_title="Optimized Exam Timetable Generator", layout="wide")

# Constants from the original script
BRANCH_FULL_FORM = {
    "B TECH": "BACHELOR OF TECHNOLOGY",
    "B TECH INTG": "BACHELOR OF TECHNOLOGY SIX YEAR INTEGRATED PROGRAM",
    "M TECH": "MASTER OF TECHNOLOGY",
    "MBA TECH": "MASTER OF BUSINESS ADMINISTRATION IN TECHNOLOGY MANAGEMENT",
    "MCA": "MASTER OF COMPUTER APPLICATIONS"
}

LOGO_PATH = "logo.png"  # Adjust as needed

def int_to_roman(num):
    """Convert integer to Roman numeral."""
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

def calculate_end_time(start_time, duration_hours):
    """Calculate end time given start time and duration."""
    start = datetime.strptime(start_time, "%I:%M %p")
    duration = timedelta(hours=duration_hours)
    end = start + duration
    return end.strftime("%I:%M %p").replace("AM", "am").replace("PM", "pm")

def find_next_valid_day(start_day, holidays, exam_days, for_branches):
    """Find the next valid day that doesn't conflict with existing exams or holidays."""
    day = start_day
    while True:
        day_date = day.date()
        if day.weekday() == 6 or day_date in holidays:
            day += timedelta(days=1)
            continue
        if all(day_date not in exam_days[branch] for branch in for_branches):
            return day
        day += timedelta(days=1)

def read_timetable(excel_path):
    """Read the timetable Excel file and return DataFrames for non-electives and electives."""
    try:
        df_dict = pd.read_excel(excel_path, sheet_name=None, index_col=[0, 1])
        non_elec_dfs = {}
        elec_dfs = {}

        for sheet_name, df in df_dict.items():
            if df.empty:
                continue
            parts = sheet_name.split('_Sem_')
            main_branch = parts[0]
            semester = parts[1] if len(parts) > 1 else ""
            if semester.isdigit():
                semester = int(semester)
            else:
                semester = {"I": 1, "III": 3, "V": 5, "VII": 7, "IX": 9, "XI": 11}.get(semester, 0)

            if sheet_name.endswith('_Electives'):
                df = df.reset_index().dropna(how='all', axis=0).reset_index(drop=True)
                df['MainBranch'] = main_branch
                df['Semester'] = semester
                elec_dfs[(main_branch, semester)] = df
            else:
                df = df.reset_index().dropna(how='all', axis=0).reset_index(drop=True)
                df['MainBranch'] = main_branch
                df['Semester'] = semester
                non_elec_dfs[(main_branch, semester)] = df

        return non_elec_dfs, elec_dfs
    except Exception as e:
        st.error(f"Error reading Excel file: {str(e)}")
        return None, None

def optimize_timetable(non_elec_dfs, elec_dfs, holidays, base_date):
    """Optimize the timetable by filling empty slots and moving OEs earlier."""
    exam_days = defaultdict(set)
    optimized_dfs = {}
    all_dates = []

    # Process non-elective schedules
    for (main_branch, semester), df in non_elec_dfs.items():
        # Convert Exam Date to datetime
        df['Exam Date'] = pd.to_datetime(df['Exam Date'], format="%d-%m-%Y", errors='coerce')
        sub_branches = [col for col in df.columns if col not in ['Exam Date', 'Time Slot']]

        # Collect all subjects and their current slots
        subjects = []
        for idx, row in df.iterrows():
            exam_date = row['Exam Date']
            time_slot = row['Time Slot']
            for sub_branch in sub_branches:
                subject = row[sub_branch]
                if pd.notna(subject) and subject != "---":
                    subjects.append({
                        'SubBranch': sub_branch,
                        'Subject': subject,
                        'Exam Date': exam_date,
                        'Time Slot': time_slot,
                        'Original Index': idx
                    })
                if pd.notna(exam_date):
                    exam_days[sub_branch].add(exam_date.date())
                    all_dates.append(exam_date)

        # Sort subjects by date
        subjects = sorted(subjects, key=lambda x: x['Exam Date'] if pd.notna(x['Exam Date']) else pd.Timestamp.max)

        # Create a new DataFrame for the optimized schedule
        new_df = df.copy()
        new_df[sub_branches] = '---'  # Reset sub-branch columns
        current_date = base_date

        # Reschedule subjects to fill gaps
        for subject in subjects:
            sub_branch = subject['SubBranch']
            current_date = find_next_valid_day(current_date, holidays, exam_days, [sub_branch])
            current_date_dt = pd.Timestamp(current_date)

            # Determine time slot based on semester
            slot_str = "10:00 AM - 1:00 PM" if (semester % 2 == 1 and (semester + 1) // 2 % 2 == 1) else "2:00 PM - 5:00 PM"

            # Find an empty slot on the current date
            found_slot = False
            for idx in new_df.index:
                if new_df.at[idx, 'Exam Date'].date() == current_date.date() and new_df.at[idx, sub_branch] == '---':
                    new_df.at[idx, sub_branch] = subject['Subject']
                    found_slot = True
                    break

            if not found_slot:
                new_row = {'Exam Date': current_date_dt, 'Time Slot': slot_str}
                for col in sub_branches:
                    new_row[col] = subject['Subject'] if col == sub_branch else '---'
                new_df = pd.concat([new_df, pd.DataFrame([new_row])], ignore_index=True)

            exam_days[sub_branch].add(current_date.date())
            all_dates.append(current_date_dt)
            current_date += timedelta(days=1)

        optimized_dfs[(main_branch, semester)] = new_df.sort_values('Exam Date')

    # Schedule electives after non-electives
    max_non_elec_date = max(all_dates).date() if all_dates else base_date.date()
    elective_day1 = find_next_valid_day(datetime.combine(max_non_elec_date, datetime.min.time()) + timedelta(days=1), holidays, exam_days, ['All Streams'])
    elective_day2 = find_next_valid_day(elective_day1 + timedelta(days=1), holidays, exam_days, ['All Streams'])

    for (main_branch, semester), df in elec_dfs.items():
        df['Exam Date'] = pd.to_datetime(df['Exam Date'], format="%d-%m-%Y", errors='coerce')
        new_df = df.copy()
        new_df.loc[new_df['OE'].isin(['OE1', 'OE5']), 'Exam Date'] = elective_day1.strftime("%d-%m-%Y")
        new_df.loc[new_df['OE'].isin(['OE1', 'OE5']), 'Time Slot'] = "10:00 AM - 1:00 PM"
        new_df.loc[new_df['OE'] == 'OE2', 'Exam Date'] = elective_day2.strftime("%d-%m-%Y")
        new_df.loc[new_df['OE'] == 'OE2', 'Time Slot'] = "2:00 PM - 5:00 PM"
        optimized_dfs[(main_branch, semester)] = new_df

    # Calculate total span
    all_dates = pd.to_datetime([d for d in all_dates if pd.notna(d)] + [elective_day1, elective_day2])
    total_span = (max(all_dates) - min(all_dates)).days + 1 if all_dates.size > 0 else 0
    if total_span <= 20:
        st.success(f"✅ Optimized timetable spans {total_span} days (within 20-day target)")
    else:
        st.warning(f"⚠️ Optimized timetable spans {total_span} days, exceeding 20-day target")

    return optimized_dfs, total_span

def save_to_excel(optimized_dfs):
    """Save the optimized timetable to an Excel file."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for (main_branch, semester), df in optimized_dfs.items():
            if df.empty:
                continue
            if '_Electives' in str(df):
                sheet_name = f"{main_branch}_Sem_{int_to_roman(semester)}_Electives"
                df = df.rename(columns={'OE': 'OE Type', 'SubjectDisplay': 'Subjects'})
                cols_to_print = ['Exam Date', 'Time Slot', 'OE Type', 'Subjects']
                df[cols_to_print].to_excel(writer, sheet_name=sheet_name[:31], index=False)
            else:
                sheet_name = f"{main_branch}_Sem_{int_to_roman(semester)}"
                df.to_excel(writer, sheet_name=sheet_name[:31])
    output.seek(0)
    return output

def convert_excel_to_pdf(excel_path, pdf_path, sub_branch_cols_per_page=4):
    """Convert the Excel timetable to a PDF, reusing the original script's logic."""
    pdf = FPDF(orientation='L', unit='mm', format=(210, 500))
    pdf.set_auto_page_break(auto=False, margin=15)
    pdf.alias_nb_pages()
    df_dict = pd.read_excel(excel_path, sheet_name=None, index_col=[0, 1])

    for sheet_name, df in df_dict.items():
        if df.empty:
            continue
        parts = sheet_name.split('_Sem_')
        main_branch = parts[0]
        main_branch_full = BRANCH_FULL_FORM.get(main_branch, main_branch)
        semester = parts[1] if len(parts) > 1 else ""
        semester_roman = semester if not semester.isdigit() else int_to_roman(int(semester))
        header_content = {'main_branch_full': main_branch_full, 'semester_roman': semester_roman}

        if not sheet_name.endswith('_Electives'):
            fixed_cols = ["Exam Date", "Time Slot"]
            sub_branch_cols = [c for c in df.columns if c not in fixed_cols]
            exam_date_width = 60
            page_width = pdf.w - 2 * pdf.l_margin
            remaining = page_width - exam_date_width
            sub_width = remaining / max(len(sub_branch_cols), 1)
            col_widths = [exam_date_width] + [sub_width] * len(sub_branch_cols)

            pdf.add_page()
            footer_height = 25
            current_date = datetime.now().strftime("%A, %B %d, %Y, %I:%M %p IST")
            logo_width = 45
            logo_x = (pdf.w - logo_width) / 2
            add_footer_with_page_number(pdf, footer_height)
            add_header_to_page(pdf, current_date, logo_x, logo_width, header_content, sub_branch_cols, df['Time Slot'].iloc[0] if 'Time Slot' in df.columns else None)
            print_table_custom(pdf, df, df.columns, col_widths, line_height=10, header_content=header_content, branches=sub_branch_cols)

        else:
            exam_date_width = 60
            oe_width = 30
            subject_width = pdf.w - 2 * pdf.l_margin - exam_date_width - oe_width
            col_widths = [exam_date_width, oe_width, subject_width]
            cols_to_print = ['Exam Date', 'Time Slot', 'OE Type', 'Subjects']
            pdf.add_page()
            footer_height = 25
            current_date = datetime.now().strftime("%A, %B %d, %Y, %I:%M %p IST")
            logo_width = 45
            logo_x = (pdf.w - logo_width) / 2
            add_footer_with_page_number(pdf, footer_height)
            add_header_to_page(pdf, current_date, logo_x, logo_width, header_content, ['All Streams'], df['Time Slot'].iloc[0] if 'Time Slot' in df.columns else None)
            print_table_custom(pdf, df, cols_to_print, col_widths, line_height=10, header_content=header_content, branches=['All Streams'])

    pdf.output(pdf_path)

# Reuse PDF functions from the original script
def print_table_custom(pdf, df, columns, col_widths, line_height=5, header_content=None, branches=None, time_slot=None):
    # Implementation copied from working_with_feedback.py
    if df.empty:
        return
    setattr(pdf, '_row_counter', 0)
    
    footer_height = 25
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
    
    header_height = 85
    pdf.set_y(0)
    current_date = datetime.now().strftime("%A, %B %d, %Y, %I:%M %p IST")
    pdf.set_font("Arial", size=14)
    text_width = pdf.get_string_width(current_date)
    x = pdf.w - 10 - text_width
    pdf.set_xy(x, 5)
    pdf.cell(text_width, 10, f"Generated on: {current_date}", 0, 0, 'R')
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
            avail_w = col_widths[i] - 2 * 2
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

def wrap_text(pdf, text, col_width):
    """Wrap text to fit within a column width."""
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
    return lines

def print_row_custom(pdf, row_data, col_widths, line_height=5, header=False):
    """Print a row in the PDF with custom formatting."""
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

def add_footer_with_page_number(pdf, footer_height):
    """Add footer with signature and page number."""
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
    """Add header to a new page."""
    pdf.set_y(0)
    pdf.set_font("Arial", size=14)
    text_width = pdf.get_string_width(current_date)
    x = pdf.w - 10 - text_width
    pdf.set_xy(x, 5)
    pdf.cell(text_width, 10, f"Generated on: {current_date}", 0, 0, 'R')
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

def main():
    st.markdown("""
    <div style="background: linear-gradient(90deg, #951C1C, #C73E1D); padding: 2rem; border-radius: 10px;">
        <h1>📅 Optimized Exam Timetable Generator</h1>
        <p>MUKESH PATEL SCHOOL OF TECHNOLOGY MANAGEMENT & ENGINEERING</p>
    </div>
    """, unsafe_allow_html=True)

    # Sidebar for configuration
    with st.sidebar:
        st.markdown("### ⚙️ Configuration")
        base_date = st.date_input("Start date for exams", value=datetime(2025, 4, 1))
        base_date = datetime.combine(base_date, datetime.min.time())

        st.markdown("#### 📅 Select Holidays")
        holiday_dates = []
        if st.checkbox("April 14, 2025", value=True):
            holiday_dates.append(datetime(2025, 4, 14))
        if st.checkbox("May 1, 2025", value=True):
            holiday_dates.append(datetime(2025, 5, 1))

        st.markdown("#### 📅 Add Custom Holidays")
        num_holidays = st.number_input("Number of custom holidays", min_value=1, value=1, step=1)
        custom_holidays = []
        for i in range(num_holidays):
            holiday = st.date_input(f"Custom Holiday {i+1}")
            if holiday:
                custom_holidays.append(datetime.combine(holiday, datetime.min.time()))
        holiday_dates.extend(custom_holidays)

    # File uploader
    st.markdown("### 📁 Upload Timetable Excel File")
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])

    if uploaded_file:
        st.success("✅ File uploaded successfully!")
        if st.button("🔄 Optimize Timetable"):
            with st.spinner("Optimizing timetable..."):
                holidays_set = set(holiday_dates)
                non_elec_dfs, elec_dfs = read_timetable(uploaded_file)
                if non_elec_dfs and elec_dfs:
                    optimized_dfs, total_span = optimize_timetable(non_elec_dfs, elec_dfs, holidays_set, base_date)
                    
                    # Save to Excel
                    excel_data = save_to_excel(optimized_dfs)
                    temp_excel = "temp_optimized_timetable.xlsx"
                    with open(temp_excel, "wb") as f:
                        f.write(excel_data.getvalue())

                    # Convert to PDF
                    pdf_path = f"optimized_timetable_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
                    convert_excel_to_pdf(temp_excel, pdf_path)

                    # Clean up
                    if os.path.exists(temp_excel):
                        os.remove(temp_excel)

                    # Provide download options
                    st.markdown("### 📥 Download Optimized Timetable")
                    with open(pdf_path, "rb") as f:
                        st.download_button(
                            label="📄 Download PDF",
                            data=f,
                            file_name=pdf_path,
                            mime="application/pdf"
                        )
                    st.download_button(
                        label="📊 Download Excel",
                        data=excel_data.getvalue(),
                        file_name=f"optimized_timetable_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    # Display results
                    st.markdown("### 📊 Optimized Timetable Results")
                    for (main_branch, semester), df in optimized_dfs.items():
                        main_branch_full = BRANCH_FULL_FORM.get(main_branch, main_branch)
                        st.markdown(f"#### {main_branch_full} - Semester {int_to_roman(semester)}")
                        st.dataframe(df)

if __name__ == "__main__":
    main()
