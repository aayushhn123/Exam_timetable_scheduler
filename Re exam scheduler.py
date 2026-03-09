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
    page_title="Re-Exam Data to PDF Converter",
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
# 🏫 COLLEGE CONFIGURATION
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
# 🎨 UI & CSS
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
# 📥 RE-EXAM DATA PARSING
# ==========================================
def process_reexam_file(uploaded_file):
    """
    Parses the Re-Exam Excel file.
    Expected columns: Module Abbreviation, Module Description, Program, Stream,
                      Academic Year, Current Session, Exam Date, Exam Time, Subject Type
    Subject Type: NaN = core subject, OE1/OE2 = Open Elective
    """
    try:
        xls = pd.ExcelFile(uploaded_file)
        # Try to find the right sheet
        sheet_name = xls.sheet_names[0]
        for s in xls.sheet_names:
            if any(kw in s.lower() for kw in ['re', 'exam', 'sheet']):
                sheet_name = s
                break

        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        df.columns = df.columns.str.strip()

        # Validate critical columns
        required = ['Module Description', 'Program', 'Current Session', 'Exam Date']
        missing = [c for c in required if c not in df.columns]
        if missing:
            st.error(f"❌ Missing critical column(s): {', '.join(missing)}")
            return None, None

        # Filter out rows without an exam date
        df = df[df['Exam Date'].notna()].copy()
        df['_date_str'] = df['Exam Date'].astype(str).str.strip().str.upper()
        df = df[~df['_date_str'].isin(['NOT SCHEDULED', 'NAN', 'NAT', 'NONE', ''])].drop(columns=['_date_str'])

        parsed_dates = pd.to_datetime(df['Exam Date'], errors='coerce')
        df = df[parsed_dates.notna()].copy()
        parsed_dates = parsed_dates[parsed_dates.notna()]
        df['Exam Date'] = parsed_dates[parsed_dates.index.isin(df.index)].dt.strftime('%d-%m-%Y')

        # Normalize Exam Time
        if 'Exam Time' in df.columns:
            df['Exam Time'] = df['Exam Time'].astype(str).str.strip()
            df['Exam Time'] = df['Exam Time'].replace({'nan': '', 'NaN': '', 'None': ''})
        else:
            df['Exam Time'] = ''

        # Map columns to internal names
        df['MainBranch'] = df.get('Program', pd.Series(dtype=str)).fillna('').astype(str).str.strip()
        df['SubBranch']  = df.get('Stream',  pd.Series(dtype=str)).fillna('').astype(str).str.strip()
        # If stream is empty/nan, fall back to program name
        df['SubBranch'] = df.apply(
            lambda x: x['MainBranch'] if x['SubBranch'] in ['nan', ''] else x['SubBranch'],
            axis=1
        )

        df['Subject']    = df.get('Module Description', pd.Series(dtype=str)).fillna('').astype(str).str.strip()
        df['ModuleCode'] = ''   # Not used for re-exam

        # Subject Type: OE1, OE2 → treat as OE; NaN → core
        if 'Subject Type' in df.columns:
            df['OE'] = df['Subject Type'].apply(
                lambda x: 'OE' if (str(x).strip().upper() not in ['NAN', 'NONE', ''] and
                                   re.match(r'^OE', str(x).strip(), re.IGNORECASE)) else None
            )
        else:
            df['OE'] = None

        def get_sem_int(val):
            s = str(val).upper().strip()
            # Try arabic digits first (e.g. "Semester 2")
            m = re.search(r'\b(\d+)\b', s)
            if m: return int(m.group(1))
            # Try roman numerals — longest first (e.g. "Semester VI", "Semester XII")
            roman_map = {
                'XII': 12, 'XI': 11, 'X': 10, 'IX': 9, 'VIII': 8,
                'VII': 7, 'VI': 6, 'V': 5, 'IV': 4, 'III': 3, 'II': 2, 'I': 1
            }
            for r, i in roman_map.items():
                if re.search(r'\b' + r + r'\b', s):
                    return i
            return 1

        df['Semester'] = df.get('Current Session', pd.Series([1]*len(df))).apply(get_sem_int)

        timetable = {}
        for sem in sorted(df['Semester'].unique()):
            timetable[sem] = df[df['Semester'] == sem].copy()

        return timetable, df

    except Exception as e:
        st.error(f"Error parsing re-exam file: {e}")
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
    """
    Build a structured Excel (intermediate) from parsed re-exam data.
    Each (Program, Semester) → one core sheet + one elective sheet.
    Multi-subject same-date cells are joined with ' <hr> ' (chronologically sorted).
    OE subjects are joined with ', ' (comma-separated, no partitioning).
    """
    time_slots_dict = st.session_state.get('time_slots', {
        1: {"start": "10:00 AM", "end": "1:00 PM"},
        2: {"start": "2:00 PM",  "end": "5:00 PM"}
    })
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

            roman_sem   = int_to_roman(sem)
            slot_id     = 1 if ((sem + 1) // 2) % 2 == 1 else 2
            p_cfg       = time_slots_dict.get(slot_id, time_slots_dict[1])
            header_norm = normalize_time(f"{p_cfg['start']} - {p_cfg['end']}")

            for main_branch in df_sem['MainBranch'].unique():
                df_mb = df_sem[df_sem['MainBranch'] == main_branch].copy()

                prog_key   = prog_key_map.get(main_branch, main_branch[:12])
                core_sheet = unique_sheet(f"{prog_key}_|_Sem {roman_sem}"[:31])
                elec_sheet = unique_sheet(f"{prog_key}_|_Sem {roman_sem}_Ele"[:31])

                df_core = df_mb[df_mb['OE'].isna()].copy()
                df_elec = df_mb[df_mb['OE'].notna()].copy()

                # ── CORE SUBJECTS ──────────────────────────────────────────
                if not df_core.empty:

                    def shorten_year(y):
                        # " 2022-2023" -> "22-23", "2024-2025" -> "24-25"
                        y = str(y).strip()
                        m = re.findall(r'\d{4}', y)
                        if len(m) >= 2:
                            return f"{m[0][2:]}-{m[1][2:]}"
                        elif len(m) == 1:
                            return m[0][2:]
                        return y

                    # Deduplicate: same Subject + SubBranch + Exam Date + Exam Time
                    # across different Academic Years → merge years into one display row
                    dedup_rows = []
                    group_keys = ['Subject', 'SubBranch', 'Exam Date', 'Exam Time']
                    ay_col = 'Academic Year' if 'Academic Year' in df_core.columns else None

                    for gkey, grp in df_core.groupby(group_keys, sort=False):
                        subj        = gkey[0]
                        actual_time = str(gkey[3]).strip()

                        # Collect and deduplicate academic years
                        if ay_col:
                            raw_years = grp[ay_col].dropna().astype(str).str.strip().unique().tolist()
                            raw_years = [y for y in raw_years if y.lower() not in ['nan', 'none', '']]
                            short_years = sorted(set(shorten_year(y) for y in raw_years))
                        else:
                            short_years = []

                        time_suffix = ""
                        if actual_time and normalize_time(actual_time) != header_norm and actual_time.lower() not in ['tbd', 'nan', '']:
                            time_suffix = f" [{actual_time}]"

                        txt = subj
                        if short_years:
                            txt += f" ({', '.join(short_years)})"
                        txt += time_suffix

                        # Parse sort time for chronological ordering
                        m = re.search(r'(\d{1,2}):(\d{2})\s*([AP]M)', actual_time.upper())
                        if m:
                            h, mins = int(m.group(1)), int(m.group(2))
                            if 'PM' in m.group(3) and h < 12: h += 12
                            if 'AM' in m.group(3) and h == 12: h = 0
                            sort_time = h * 60 + mins
                        else:
                            sort_time = 9999

                        dedup_rows.append({
                            'SubBranch':      gkey[1],
                            'Exam Date':      gkey[2],
                            'Exam Time':      actual_time,
                            'SubjectDisplay': txt,
                            '_SortTime':      sort_time,
                        })

                    df_core = pd.DataFrame(dedup_rows)
                    df_core['Exam Date']      = pd.to_datetime(
                        df_core['Exam Date'], format='%d-%m-%Y', dayfirst=True, errors='coerce'
                    )
                    # Sort by date then by exam time within date
                    df_core = df_core.sort_values(
                        by=['Exam Date', '_SortTime'], ascending=[True, True]
                    )

                    try:
                        pivot = df_core.groupby(['Exam Date', 'SubBranch']).agg(
                            {'SubjectDisplay': lambda x: " <hr> ".join(str(i) for i in x)}
                        ).reset_index()
                        pivot = pivot.pivot_table(
                            index='Exam Date', columns='SubBranch',
                            values='SubjectDisplay', aggfunc='first'
                        ).fillna("---")
                        pivot = pivot.sort_index(ascending=True).reset_index()
                        pivot['Exam Date'] = pivot['Exam Date'].apply(
                            lambda x: x.strftime('%d-%m-%Y') if pd.notna(x) else ""
                        )
                        pivot['_prog_'] = main_branch
                        pivot['_sem_']  = roman_sem
                        pivot.to_excel(writer, sheet_name=core_sheet, index=False)
                        sheets_created += 1
                    except Exception:
                        pass

                # ── OPEN ELECTIVES ─────────────────────────────────────────
                if not df_elec.empty:
                    e_displays = []
                    for _, row in df_elec.iterrows():
                        subj        = row['Subject']
                        actual_time = str(row.get('Exam Time', '')).strip()

                        # OE: NO module code; show time only if different from header norm
                        time_suffix = ""
                        if actual_time and normalize_time(actual_time) != header_norm and actual_time.lower() not in ['tbd', 'nan', '']:
                            time_suffix = f" [{actual_time}]"

                        # Academic year dedup for OE — same subject same day → merge years
                        ay_col_oe = 'Academic Year' if 'Academic Year' in row.index else None
                        # (OE dedup handled at groupby level below; individual display is plain)
                        txt = f"{subj}{time_suffix}"
                        e_displays.append(txt)

                    df_elec['DisplaySubject'] = e_displays

                    try:
                        df_elec['Exam Date'] = pd.to_datetime(
                            df_elec['Exam Date'], format='%d-%m-%Y', dayfirst=True, errors='coerce'
                        )
                        df_elec = df_elec.sort_values(by='Exam Date', ascending=True)
                        df_elec['Exam Date'] = df_elec['Exam Date'].apply(
                            lambda x: x.strftime('%d-%m-%Y') if pd.notna(x) else ""
                        )
                        # Group OE by date → comma-separated, deduplicated, sorted
                        ep = df_elec.groupby(['Exam Date', 'OE']).agg(
                            {'DisplaySubject': lambda x: ", ".join(sorted(set(x)))}
                        ).reset_index()
                        ep.rename(
                            columns={'OE': 'OE Type',
                                     'DisplaySubject': 'Open Elective (All Applicable Streams)'},
                            inplace=True
                        )
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

    time_pattern = re.compile(
        r'([\[\(]\s*\d{1,2}:\d{2}\s*[AP]M\s*-\s*\d{1,2}:\d{2}\s*[AP]M\s*[\]\)])',
        re.IGNORECASE
    )

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
    old_style  = pdf.font_style
    old_size   = pdf.font_size_pt

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


def print_row_custom(pdf, row_data, col_widths, line_height=5, header=False):
    cell_padding    = 1
    header_bg_color = (255, 255, 255)
    header_text_color = (0, 0, 0)
    alt_row_color   = (255, 255, 255)

    row_number = getattr(pdf, '_row_counter', 0)

    base_font = "Times"
    if header:
        base_style = 'B'
        base_size  = 9.5
        pdf.set_font(base_font, base_style, base_size)
        pdf.set_text_color(*header_text_color)
        pdf.set_fill_color(*header_bg_color)
    else:
        base_style = ''
        base_size  = 9.5
        pdf.set_font(base_font, base_style, base_size)
        pdf.set_text_color(0, 0, 0)
        pdf.set_fill_color(*alt_row_color)

    wrapped_cells = []
    max_lines = 0
    for i, cell_text in enumerate(row_data):
        text   = str(cell_text) if cell_text is not None else ""
        avail_w = col_widths[i] - 2 * cell_padding
        lines  = wrap_text(pdf, text, avail_w)
        wrapped_cells.append(lines)
        max_lines = max(max_lines, len(lines))

    # Outer row height — strictly line_height * max_lines
    row_h = line_height * max_lines

    # Inner text line height
    text_line_height = line_height * 0.75

    x0, y0 = pdf.get_x(), pdf.get_y()

    pdf.rect(x0, y0, sum(col_widths), row_h, 'F')

    time_pattern = re.compile(
        r'([\[\(]\s*\d{1,2}:\d{2}\s*[AP]M\s*-\s*\d{1,2}:\d{2}\s*[AP]M\s*[\]\)])',
        re.IGNORECASE
    )

    for i, lines in enumerate(wrapped_cells):
        cx = pdf.get_x()

        # Split by <hr> to get distinct subject partitions
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
            # Vertically center each subject within its partition
            total_text_h = len(subj_lines) * text_line_height
            pad_v = (part_h - total_text_h) / 2

            for j, ln in enumerate(subj_lines):
                parts = time_pattern.split(ln)

                if len(parts) == 1 or header:
                    pdf.set_xy(cx + cell_padding,
                               y0 + (sub_idx * part_h) + pad_v + j * text_line_height)
                    pdf.cell(col_widths[i] - 2 * cell_padding, text_line_height, ln, border=0, align='C')
                else:
                    total_w = 0
                    for k, p in enumerate(parts):
                        if not p: continue
                        if k % 2 == 1: pdf.set_font(base_font, 'B', base_size)
                        else:          pdf.set_font(base_font, base_style, base_size)
                        total_w += pdf.get_string_width(p)

                    start_x   = cx + max(cell_padding, (col_widths[i] - total_w) / 2)
                    current_x = start_x

                    for k, p in enumerate(parts):
                        if not p: continue
                        if k % 2 == 1: pdf.set_font(base_font, 'B', base_size)
                        else:          pdf.set_font(base_font, base_style, base_size)

                        w = pdf.get_string_width(p)
                        pdf.set_xy(current_x - pdf.c_margin,
                                   y0 + (sub_idx * part_h) + pad_v + j * text_line_height)
                        pdf.cell(w + 2 * pdf.c_margin, text_line_height, p, border=0, align='L')
                        current_x += w

                    pdf.set_font(base_font, base_style, base_size)

            # Draw horizontal partition border between subjects
            if sub_idx < num_subjects - 1:
                line_y = y0 + ((sub_idx + 1) * part_h)
                pdf.line(cx, line_y, cx + col_widths[i], line_y)

        pdf.rect(cx, y0, col_widths[i], row_h)
        pdf.set_xy(cx + col_widths[i], y0)

    setattr(pdf, '_row_counter', row_number + 1)
    pdf.set_xy(x0, y0 + row_h)


def print_table_custom(pdf, df, columns, col_widths, line_height=5,
                        header_content=None, Programs=None, time_slot=None,
                        actual_time_slots=None, declaration_date=None):
    if df.empty: return
    setattr(pdf, '_row_counter', 0)

    footer_height = 14
    header_end_y  = 60   # Header block must end exactly at y=60

    def render_footer():
        pdf.set_xy(10, pdf.h - footer_height)
        pdf.set_font("Times", 'B', 8)
        pdf.cell(0, 5, "CONTROLLER OF EXAMINATIONS", 0, 1, 'L')
        pdf.line(10, pdf.h - footer_height + 5, 60, pdf.h - footer_height + 5)

        pdf.set_font("Times", size=8)
        pdf.set_text_color(0, 0, 0)
        page_text  = f"{pdf.page_no()} of {{nb}}"
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
        logo_x     = (pdf.w - logo_width) / 2
        if os.path.exists(LOGO_PATH):
            pdf.image(LOGO_PATH, x=logo_x, y=5, w=logo_width)

        # College Name — size 12, bold
        pdf.set_text_color(0, 0, 0)
        college_name = st.session_state.get('selected_college', "SVKM's NMIMS University").upper()
        pdf.set_font("Times", 'B', 12)
        pdf.set_xy(10, 25)
        pdf.cell(pdf.w - 20, 6, college_name, 0, 1, 'C')

        # Main Title — size 10, bold
        pdf.set_font("Times", 'B', 10)
        pdf.set_text_color(0, 0, 0)
        pdf.set_xy(10, 33)
        pdf.cell(pdf.w - 20, 4, "RE-EXAMINATION TIMETABLE (ACADEMIC YEAR: 2025-26)", 0, 1, 'C')

        current_y = 38

        # Program Name — size 10, bold
        pdf.set_font("Times", 'B', 10)
        pdf.set_xy(10, current_y)
        pdf.cell(pdf.w - 20, 4, f"{header_content['main_branch_full']}".upper(), 0, 1, 'C')
        current_y += 4

        # Year and Semester
        sem_roman = str(header_content['semester_roman']).upper()
        roman_map = {
            'XII': 12, 'XI': 11, 'X': 10, 'IX': 9, 'VIII': 8,
            'VII': 7, 'VI': 6, 'V': 5, 'IV': 4, 'III': 3, 'II': 2, 'I': 1
        }
        sem_int = roman_map.get(sem_roman)
        if not sem_int:
            m = re.search(r'(\d+)', sem_roman)
            if m: sem_int = int(m.group(1))
            else: sem_int = 1
        year_int   = (sem_int + 1) // 2
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

        # Lock position to header_end_y
        pdf.set_xy(pdf.l_margin, header_end_y)

    render_footer()
    render_header()

    # Uppercase headers
    upper_columns = [str(c).upper() for c in columns]

    pdf.set_font("Times", 'B', 9.5)
    print_row_custom(pdf, upper_columns, col_widths, line_height=line_height, header=True)

    pdf.set_font("Times", '', 9.5)

    for idx in range(len(df)):
        row = [str(df.iloc[idx][c]) if pd.notna(df.iloc[idx][c]) else "" for c in columns]
        if not any(cell.strip() for cell in row): continue

        wrapped_cells = []
        max_lines     = 0
        for i, cell_text in enumerate(row):
            text    = str(cell_text) if cell_text is not None else ""
            avail_w = col_widths[i] - 2
            lines   = wrap_text(pdf, text, avail_w)
            wrapped_cells.append(lines)
            max_lines = max(max_lines, len(lines))
        row_h = line_height * max_lines

        # Page break check
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
        2: {"start": "2:00 PM",  "end": "5:00 PM"}
    })

    def get_header_time_for_semester(sem_str):
        try:
            s = str(sem_str).strip().upper()
            sem_int = 1
            romans = {'I': 1, 'II': 2, 'III': 3, 'IV': 4, 'V': 5, 'VI': 6,
                      'VII': 7, 'VIII': 8, 'IX': 9, 'X': 10, 'XI': 11, 'XII': 12}
            found = False
            for r_key, r_val in romans.items():
                if s == r_key or s.endswith(f" {r_key}") or s.endswith(f"_{r_key}"):
                    sem_int = r_val
                    found   = True
                    break
            if not found:
                digits = re.findall(r'\d+', s)
                if digits: sem_int = int(digits[0])
            slot_indicator = ((sem_int + 1) // 2) % 2
            slot_num  = 1 if slot_indicator == 1 else 2
            slot_cfg  = time_slots_dict.get(slot_num, time_slots_dict.get(1))
            return f"{slot_cfg['start']} - {slot_cfg['end']}"
        except:
            return f"{time_slots_dict[1]['start']} - {time_slots_dict[1]['end']}"

    try:
        df_dict = pd.read_excel(excel_path, sheet_name=None)
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return

    sheets_processed = 0

    for sheet_name, sheet_df in df_dict.items():
        try:
            if sheet_df.empty: continue
            if hasattr(sheet_df, 'index') and len(sheet_df.index.names) > 1:
                sheet_df = sheet_df.reset_index()

            # Resolve main_branch_full
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

            # Parse semester from sheet name
            semester_raw = "General"
            if '_|_' in sheet_name:
                parts = sheet_name.split('_|_')
                if not main_branch_full: main_branch_full = parts[0]
                semester_raw = parts[1]
            else:
                if sheet_name in ["No_Data", "Daily_Statistics", "Summary",
                                  "Verification", "Empty"]:
                    continue
                if not main_branch_full: main_branch_full = sheet_name

            is_elective = False
            if re.search(r'_Ele(\d*)$', semester_raw):
                semester_raw = re.sub(r'_Ele\d*$', '', semester_raw)
                is_elective  = True

            if "_sem_" in sheet_df.columns and not sheet_df["_sem_"].dropna().empty:
                display_sem = str(sheet_df["_sem_"].dropna().iloc[0]).strip()
            else:
                display_sem = semester_raw.strip()
                if display_sem.lower().startswith("semester"):
                    display_sem = display_sem[8:].strip()
                elif display_sem.lower().startswith("sem"):
                    display_sem = display_sem[3:].strip()

            header_content = {
                'main_branch_full': main_branch_full,
                'semester_roman':   display_sem
            }

            header_exam_time = get_header_time_for_semester(f"Sem {display_sem}")

            if not is_elective:
                if 'Exam Date' not in sheet_df.columns: continue
                sheet_df = sheet_df.dropna(how='all').reset_index(drop=True)

                fixed_cols = ["Exam Date"]
                _meta_pattern = re.compile(
                    r'^(Program|Semester|MainBranch|Note|Message|_prog_|_sem_)(\.\d+)?$',
                    re.IGNORECASE
                )
                sub_branch_cols = [
                    c for c in sheet_df.columns
                    if c not in fixed_cols
                    and not _meta_pattern.match(str(c))
                    and pd.notna(c)
                    and str(c).strip() != ''
                ]
                if not sub_branch_cols: continue

                cols_per_page = 6

                for start in range(0, len(sub_branch_cols), cols_per_page):
                    chunk       = sub_branch_cols[start:start + cols_per_page]
                    cols_to_print = fixed_cols + chunk
                    chunk_df    = sheet_df[cols_to_print].copy()

                    subset      = chunk_df[chunk].astype(str).apply(lambda x: x.str.strip())
                    valid_cells = (subset != "") & (subset != "nan") & (subset != "---")
                    mask        = valid_cells.any(axis=1)
                    chunk_df    = chunk_df[mask].reset_index(drop=True)
                    if chunk_df.empty: continue

                    try:
                        chunk_df["Exam Date"] = pd.to_datetime(
                            chunk_df["Exam Date"], format="%d-%m-%Y", errors='coerce'
                        ).dt.strftime("%A, %d %B, %Y")
                    except:
                        pass

                    page_width   = pdf.w - 2 * pdf.l_margin
                    date_col_width  = 30
                    remaining_width = page_width - date_col_width
                    num_sub      = max(len(chunk), 1)
                    sub_width    = remaining_width / num_sub
                    col_widths   = [date_col_width] + [sub_width] * len(chunk)

                    pdf.add_page()

                    original_college = st.session_state.get('selected_college')
                    st.session_state['selected_college'] = sheet_college_name

                    print_table_custom(
                        pdf, chunk_df, cols_to_print, col_widths, line_height=5,
                        header_content=header_content, Programs=chunk,
                        time_slot=header_exam_time, actual_time_slots=None,
                        declaration_date=declaration_date
                    )

                    if original_college:
                        st.session_state['selected_college'] = original_college
                    sheets_processed += 1

            else:
                # Elective sheet
                target_cols   = ['Exam Date', 'OE Type', 'Open Elective (All Applicable Streams)']
                available_cols = [c for c in target_cols if c in sheet_df.columns]

                if len(available_cols) >= 3:
                    sheet_df = sheet_df.dropna(subset=['Exam Date']).reset_index(drop=True)
                    if sheet_df.empty: continue

                    try:
                        sheet_df["Exam Date"] = pd.to_datetime(
                            sheet_df["Exam Date"], format="%d-%m-%Y", errors='coerce'
                        ).dt.strftime("%A, %d %B, %Y")
                    except:
                        pass

                    pdf.add_page()

                    col_widths      = [30, 25]
                    remaining_width = pdf.w - 2 * pdf.l_margin - sum(col_widths)
                    col_widths.append(remaining_width)

                    original_college = st.session_state.get('selected_college')
                    st.session_state['selected_college'] = sheet_college_name

                    print_table_custom(
                        pdf, sheet_df, available_cols, col_widths, line_height=5,
                        header_content=header_content, Programs=["Electives"],
                        time_slot=header_exam_time, actual_time_slots=None,
                        declaration_date=declaration_date
                    )

                    if original_college:
                        st.session_state['selected_college'] = original_college
                    sheets_processed += 1

        except Exception as e:
            st.warning(f"Error processing PDF sheet {sheet_name}: {e}")
            continue

    if sheets_processed == 0:
        st.error("No valid sheets generated in PDF.")
        return

    # Instructions / last page
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
        if os.path.exists(LOGO_PATH):
            pdf.image(LOGO_PATH, x=(pdf.w - 45) / 2, y=5, w=45)
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Times", 'B', 12)
        pdf.set_xy(10, 25)
        pdf.cell(
            pdf.w - 20, 8,
            st.session_state.get('selected_college', "SVKM's NMIMS University").upper(),
            0, 1, 'C'
        )

        pdf.set_font("Times", 'B', 12)
        pdf.set_text_color(0, 0, 0)
        pdf.set_xy(10, 40)
        pdf.cell(0, 10, "RE-EXAMINATION GUIDELINES", 0, 1, 'C')

        pdf.set_y(60)
        pdf.set_font("Times", 'B', 12)
        pdf.cell(0, 10, "INSTRUCTIONS TO CANDIDATES", 0, 1, 'C')
        pdf.ln(5)
        instrs = [
            "1. Candidates are required to be present at the examination center THIRTY MINUTES before the stipulated time.",
            "2. Candidates must produce their University Identity Card at the time of the examination.",
            "3. Candidates are not permitted to enter the examination hall after stipulated time.",
            "4. Candidates will not be permitted to leave the examination hall during the examination time.",
            "5. Please verify the individual subject exam time as indicated in the timetable. Different subjects may have different time slots.",
        ]
        pdf.set_font("Times", size=12)
        for instr in instrs:
            pdf.multi_cell(0, 7, instr)
            pdf.ln(2)
    except Exception:
        pass

    try:
        pdf.output(pdf_path)
    except Exception as e:
        st.error(f"Save PDF failed: {e}")


# ==========================================
# 🔄 GENERATE PDF TIMETABLE (ORCHESTRATOR)
# ==========================================
def generate_pdf_timetable(semester_wise_timetable, output_pdf, declaration_date=None):
    temp_dir   = os.path.dirname(output_pdf) if os.path.dirname(output_pdf) else "."
    temp_excel = os.path.join(temp_dir, "temp_reexam_timetable.xlsx")

    excel_data = save_to_excel(semester_wise_timetable)

    if not excel_data:
        st.error("❌ No Excel data generated — cannot create PDF")
        return

    try:
        with open(temp_excel, "wb") as f:
            f.write(excel_data.getvalue())
    except Exception as e:
        st.error(f"❌ Error saving temporary Excel file: {e}")
        return

    try:
        convert_excel_to_pdf(temp_excel, output_pdf, declaration_date=declaration_date)
    except Exception as e:
        st.error(f"❌ Error during Excel to PDF conversion: {e}")
        st.error(traceback.format_exc())
        return
    finally:
        try:
            if os.path.exists(temp_excel):
                os.remove(temp_excel)
        except Exception:
            pass

    try:
        if not os.path.exists(output_pdf):
            st.error(f"❌ PDF file was not created at {output_pdf}")
            return

        reader              = PdfReader(output_pdf)
        writer              = PdfWriter()
        page_number_pattern = re.compile(r'^[\s\n]*(?:Page\s*)?\d+[\s\n]*$')

        for page in reader.pages:
            try:    text = page.extract_text() or ""
            except: text = ""
            cleaned = text.strip()
            if cleaned and not page_number_pattern.match(cleaned) and len(cleaned) > 10:
                writer.add_page(page)

        if writer.pages:
            with open(output_pdf, 'wb') as f:
                writer.write(f)
            st.success(f"✅ PDF generation complete — {len(writer.pages)} page(s)")
        else:
            st.warning("⚠️ All pages were filtered — keeping original PDF")

    except Exception as e:
        st.error(f"❌ PDF post-processing error: {e}")
        st.error(traceback.format_exc())


# ==========================================
# 🚀 MAIN APP
# ==========================================
def main():
    st.markdown(
        '<div class="main-header">'
        '<h1>📄 Re-Exam Data to PDF Converter</h1>'
        '<p>Generate Standardized Re-Examination Timetables Directly from Re-Exam Excel Data</p>'
        '</div>',
        unsafe_allow_html=True
    )

    for key in ('pdf_data', 'raw_df', 'processed_tt'):
        if key not in st.session_state:
            st.session_state[key] = None

    with st.sidebar:
        st.header("⚙️ Configuration")
        college_options  = [c["name"] for c in COLLEGES]
        selected_display = st.selectbox("Select College", college_options)
        st.session_state.selected_college = selected_display
        st.markdown("---")
        decl_date = st.date_input("📆 Declaration Date (Optional)", value=None)
        st.markdown("---")
        st.subheader("⏰ Exam Time Slots")
        st.caption("Odd semesters (I, III, V…) → Slot 1 | Even semesters (II, IV, VI…) → Slot 2")
        slot1_start = st.text_input("Slot 1 Start", value="10:00 AM", key="s1s")
        slot1_end   = st.text_input("Slot 1 End",   value="1:00 PM",  key="s1e")
        slot2_start = st.text_input("Slot 2 Start", value="2:00 PM",  key="s2s")
        slot2_end   = st.text_input("Slot 2 End",   value="5:00 PM",  key="s2e")
        st.session_state['time_slots'] = {
            1: {"start": slot1_start, "end": slot1_end},
            2: {"start": slot2_start, "end": slot2_end},
        }
        st.markdown("---")
        st.info(
            "**Re-Exam Columns Expected:**\n\n"
            "- Module Abbreviation\n"
            "- Module Description\n"
            "- Program\n"
            "- Stream\n"
            "- Academic Year\n"
            "- Current Session\n"
            "- Exam Date\n"
            "- Exam Time\n"
            "- Subject Type (OE1/OE2 = Open Elective)"
        )

    col1, col2 = st.columns([2, 1])

    with col1:
        st.markdown(
            '<div class="upload-section">'
            '<h3 style="margin:0 0 1rem 0; color:#951C1C;">📁 Upload Re-Exam Excel File</h3>'
            '<p style="margin:0; color:#666; font-size:1rem;">Drag and drop the Re-Exam scheduling Excel file</p>'
            '</div>',
            unsafe_allow_html=True
        )
        uploaded = st.file_uploader(
            "Upload Excel", type=['xlsx', 'xls'], label_visibility="collapsed"
        )

        if uploaded:
            with st.spinner("Parsing re-exam file…"):
                tt, df = process_reexam_file(uploaded)
                if tt:
                    st.session_state.processed_tt = tt
                    st.session_state.raw_df        = df
                    st.success(
                        f"✅ Loaded successfully — {len(df)} scheduled subjects "
                        f"across {len(tt)} semester(s)"
                    )

    with col2:
        st.info(
            "ℹ️ **Re-Exam Converter Mode**\n\n"
            "Processes re-exam scheduled subjects and generates a Legal Landscape PDF "
            "with the exact same header, footer, fonts, bold time-block rendering, "
            "and <hr> boundary-partitioning math as the master timetable app.\n\n"
            "Each subject's individual exam time is always shown in **bold** brackets."
        )

    if st.session_state.raw_df is not None:
        df = st.session_state.raw_df

        # Preview summary
        st.markdown("---")
        st.subheader("📊 Data Preview")
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            st.metric("Total Subjects", len(df))
        with col_b:
            st.metric("Programs", df['MainBranch'].nunique())
        with col_c:
            st.metric("Semesters", df['Semester'].nunique())

        with st.expander("🔍 View Raw Data (first 50 rows)"):
            st.dataframe(df[['Subject', 'ModuleCode', 'MainBranch', 'SubBranch',
                             'Semester', 'Exam Date', 'Exam Time', 'OE']].head(50))

        st.markdown("---")

        if st.button("🚀 Generate Re-Exam PDF Timetable", type="primary", use_container_width=True):
            with st.spinner("Rendering PDF…"):
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
                            "📥 Download Re-Exam PDF",
                            f,
                            f"ReExam_Timetable_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                            "application/pdf",
                            use_container_width=True
                        )
                    try:
                        os.remove(pdf_filename)
                    except:
                        pass


if __name__ == "__main__":
    main()
