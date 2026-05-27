import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from fpdf import FPDF
import os
import re
import io
import zipfile
import uuid
import traceback
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

        # Rule 2 Tagging Setup: Capture the actual OE string dynamically if it contains "OE"
        if 'Subject Type' in df.columns:
            def extract_oe(x):
                if pd.isna(x): return None
                val = str(x).strip()
                if 'OE' in val.upper(): return val
                return None
            df['OE'] = df['Subject Type'].apply(extract_oe)
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

        # ── Derive ExamSlotNumber from Exam Time for SBM/Pravin Dalal ──────
        current_college = st.session_state.get('selected_college', '')
        IS_BUSINESS_SCH = ("School of Business Management" in current_college
                           or "Pravin Dalal" in current_college)
        if IS_BUSINESS_SCH:
            time_slots_dict = st.session_state.get('time_slots', {
                1: {"start": "10:00 AM", "end": "1:00 PM"},
                2: {"start": "2:00 PM",  "end": "5:00 PM"}
            })
            # Build normalised time → slot lookup
            def _norm_t(s):
                s = str(s).strip().upper()
                for i in range(1, 10): s = s.replace(f"0{i}:", f"{i}:")
                return s
            _t2slot = {}
            for _sn, _scfg in time_slots_dict.items():
                _t2slot[_norm_t(f"{_scfg['start']} - {_scfg['end']}")] = _sn
                _t2slot[_norm_t(_scfg['start'])] = _sn  # match on start time alone

            def _derive_slot(exam_time):
                t = _norm_t(str(exam_time))
                # try full range match first, then start-time prefix
                for key, sn in _t2slot.items():
                    if key in t or t.startswith(key.split(' - ')[0]):
                        return sn
                return 1  # default to slot 1

            df['ExamSlotNumber'] = df['Exam Time'].apply(_derive_slot)
        else:
            df['ExamSlotNumber'] = 1

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
    IS_LAW_SCHOOL = "LAW" in current_college_context.upper()

    # SOL Fix 2: Override default time slots for School of Law
    if IS_LAW_SCHOOL:
        time_slots_dict = {
            1: {"start": "11:00 AM", "end": "1:00 PM"},
            2: {"start": "2:30 PM",  "end": "4:30 PM"}
        }

    output = io.BytesIO()

    # Rule 1: B.A. / B.B.A. Side-by-Side Merging Pre-Processing
    if IS_LAW_SCHOOL:

        # SOL Fix 1: Canonical program name normaliser — enforces exact spacing & capitalisation
        # Converts any variant (e.g. "B.A., LL.B. (Hons.)", "BA LLB Hons", "b.a. llb (hons)")
        # to the exact canonical forms: "B.A., LL.B.(Hons.)" or "B.B.A., LL.B.(Hons.)"
        def _sol_normalise_program_name(raw):
            s = str(raw).strip()
            su = s.upper().replace(" ", "")
            # Detect B.B.A. variant first (must come before B.A. check)
            if re.search(r'B\.?B\.?A\.?', su):
                return "B.B.A., LL.B.(Hons.)"
            # Detect plain B.A. variant (not B.B.A.)
            if re.search(r'^B\.?A\.?', su):
                return "B.A., LL.B.(Hons.)"
            return s  # not a BA/BBA LLB — return unchanged

        for sem in semester_wise_timetable:
            df_sem = semester_wise_timetable[sem]
            if df_sem.empty: continue

            def is_ba_bba_llb(p):
                p_str = str(p).strip().upper()
                if re.search(r'(LL\.M|MASTER\s+OF\s+LAW|LLM)', p_str):
                    return False
                return bool(re.search(r'^(B\.A\.|B\.B\.A\.)[,\s].*LL\.B', p_str))

            mask = df_sem['MainBranch'].apply(is_ba_bba_llb)
            if mask.any():
                # Normalise the raw MainBranch to canonical name BEFORE using it in SubBranch prefix
                normalised_prefix = df_sem.loc[mask, 'MainBranch'].apply(_sol_normalise_program_name)
                df_sem.loc[mask, 'SubBranch'] = normalised_prefix + " - " + df_sem.loc[mask, 'SubBranch']
                df_sem.loc[mask, 'MainBranch'] = "B.A., LL.B.(Hons.) / B.B.A., LL.B.(Hons.)"

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

                # Rule 2: Inline Open Electives execution - Do not split OE for SOL BA/BBA
                if IS_LAW_SCHOOL and main_branch == "B.A., LL.B.(Hons.) / B.B.A., LL.B.(Hons.)":
                    df_core = df_mb.copy()
                    df_elec = pd.DataFrame()
                else:
                    df_core = df_mb[df_mb['OE'].isna()].copy()
                    df_elec = df_mb[df_mb['OE'].notna()].copy()

                if not df_core.empty:
                    displays = []
                    sort_times = []
                    for _, row in df_core.iterrows():
                        subj        = row['Subject']
                        code        = row['ModuleCode']
                        actual_time = str(row.get('Exam Time', '')).strip()
                        oe_type     = row.get('OE', None)

                        time_suffix = ""
                        if actual_time and actual_time.lower() not in ['tbd', 'nan', '']:
                            if IS_LAW_SCHOOL:
                                # SOL Fix: always embed time so PDF stage can compare against majority
                                time_suffix = f" [{actual_time}]"
                            elif normalize_time(actual_time) != header_norm:
                                time_suffix = f" [{actual_time}]"

                        # Rule 2: Tagging OE subjects visually inside the core dataframe cell
                        prefix = ""
                        if IS_LAW_SCHOOL and main_branch == "B.A., LL.B.(Hons.) / B.B.A., LL.B.(Hons.)" and pd.notna(oe_type) and str(oe_type).strip() != '':
                            prefix = f"[OE: {oe_type}] "

                        # Rule 4: Subject String Formatting (No hyphen before module code)
                        txt = f"{prefix}{subj}"
                        if code and str(code).lower() != 'nan': txt += f" ({code})"
                        txt += time_suffix
                        displays.append(txt)

                        # Parse time for chronological sorting within partitioned cells
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
                    # Sort chronologically (Exam Date, then internal Subject Time)
                    df_core = df_core.sort_values(by=["Exam Date", "_SortTime"], ascending=[True, True])

                    try:
                        pivot = df_core.groupby(['Exam Date', 'SubBranch']).agg({'SubjectDisplay': lambda x: " <hr> ".join(str(i) for i in x)}).reset_index()
                        pivot = pivot.pivot_table(index='Exam Date', columns='SubBranch', values='SubjectDisplay', aggfunc='first').fillna("---")
                        pivot = pivot.sort_index(ascending=True).reset_index()
                        pivot['Exam Date'] = pivot['Exam Date'].apply(lambda x: x.strftime("%d-%m-%Y") if pd.notna(x) else "")

                        # SOL Fix: For B.A./B.B.A. LL.B. Year 3+ (Sem 5+), subjects/streams
                        # are common across both programs — collapse duplicate column pairs
                        # into a single column showing just the stream name.
                        # Also handles Sem X where stream columns appear twice (once per program).
                        if IS_LAW_SCHOOL and main_branch == "B.A., LL.B.(Hons.) / B.B.A., LL.B.(Hons.)" and sem >= 5:
                            _fixed = ['Exam Date']
                            _data_cols = [c for c in pivot.columns if c not in _fixed + ['_prog_', '_sem_']]
                            # Build a map: stream_name -> first column that owns it
                            # Strip the "PROG - " prefix to get the bare stream name
                            _stream_map = {}   # stream_name -> canonical col name to keep
                            _drop_cols  = []   # duplicate cols to drop
                            for _col in _data_cols:
                                _col_str = str(_col)
                                _stream = _col_str.rsplit(' - ', 1)[-1].strip() if ' - ' in _col_str else _col_str
                                if _stream not in _stream_map:
                                    _stream_map[_stream] = _col
                                else:
                                    # Merge: combine non-'---' content from this col into the kept col
                                    _kept = _stream_map[_stream]
                                    for _idx in pivot.index:
                                        _kept_val = str(pivot.at[_idx, _kept]).strip()
                                        _this_val = str(pivot.at[_idx, _col]).strip()
                                        if (_kept_val == '---' or _kept_val == '') and _this_val not in ('---', ''):
                                            pivot.at[_idx, _kept] = _this_val
                                    _drop_cols.append(_col)
                            # Drop duplicate columns
                            if _drop_cols:
                                pivot = pivot.drop(columns=_drop_cols)
                            # Rename kept columns to bare stream name (strip program prefix)
                            _rename = {}
                            for _stream, _col in _stream_map.items():
                                if _col in pivot.columns and str(_col) != _stream:
                                    _rename[_col] = _stream
                            if _rename:
                                pivot = pivot.rename(columns=_rename)

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

def print_row_custom(pdf, row_data, col_widths, line_height=5, header=False):
    cell_padding = 1
    header_bg_color = (255, 255, 255)
    header_text_color = (0, 0, 0)
    alt_row_color = (255, 255, 255)

    # SOL Fix 3: Use light grey fill and increased font size for SOL header rows when flagged
    if header and hasattr(pdf, '_sol_header_fill'):
        header_bg_color = pdf._sol_header_fill

    row_number = getattr(pdf, '_row_counter', 0)
    
    base_font = "Times"
    if header:
        base_style = 'B'
        # SOL Fix: use increased font size for SOL header if flagged, else standard 9.5
        base_size = getattr(pdf, '_sol_header_font_size', 9.5)
        pdf.set_font(base_font, base_style, base_size)
        pdf.set_text_color(*header_text_color)
        pdf.set_fill_color(*header_bg_color)
    else:
        base_style = ''
        base_size = 9.5
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

    # Keep outer row height standard based on max total lines
    row_h = line_height * max_lines
    
    # Tighter line spacing internally for text
    text_line_height = line_height * 0.75 
    
    x0, y0 = pdf.get_x(), pdf.get_y()
    
    pdf.rect(x0, y0, sum(col_widths), row_h, 'F')

    time_pattern = re.compile(r'([\[\(]\s*\d{1,2}:\d{2}\s*[AP]M\s*-\s*\d{1,2}:\d{2}\s*[AP]M\s*[\]\)])', re.IGNORECASE)

    for i, lines in enumerate(wrapped_cells):
        cx = pdf.get_x()
        
        # Split lines by <hr> into distinct subjects to partition the cell
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
            # Vertically center each subject cleanly inside its designated horizontal partition
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
            
            # Draw the horizontal partition border exactly on the boundary between subjects
            if sub_idx < num_subjects - 1:
                line_y = y0 + ((sub_idx + 1) * part_h)
                pdf.line(cx, line_y, cx + col_widths[i], line_y)
        
        pdf.rect(cx, y0, col_widths[i], row_h)
        pdf.set_xy(cx + col_widths[i], y0)

    setattr(pdf, '_row_counter', row_number + 1)
    pdf.set_xy(x0, y0 + row_h)

def print_table_custom(pdf, df, columns, col_widths, line_height=5, header_content=None, Programs=None, time_slot=None, actual_time_slots=None, declaration_date=None):
    if df.empty: return
    setattr(pdf, '_row_counter', 0)
    
    footer_height = 14   
    header_end_y = 60    
    
    def render_footer():
        pdf.set_xy(10, pdf.h - footer_height)
        pdf.set_font("Times", 'B', 11) 
        pdf.cell(0, 5, "CONTROLLER OF EXAMINATIONS", 0, 1, 'L')
        pdf.line(10, pdf.h - footer_height + 5, 80, pdf.h - footer_height + 5)
        
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
        
        # College Name (Rule 3 handled inside convert_excel_to_pdf which passes it to st.session_state)
        pdf.set_text_color(0, 0, 0)
        college_name = st.session_state.get('selected_college', "SVKM's NMIMS University").upper()
        pdf.set_font("Times", 'B', 12) 
        
        pdf.set_xy(10, 25)
        pdf.cell(pdf.w - 20, 6, college_name, 0, 1, 'C')
        
        # Final Exam Timetable Header
        # SOL Fix: increase header section font sizes by 1.5
        _hdr_is_law = "LAW" in st.session_state.get('selected_college', '').upper()
        pdf.set_font("Times", 'B', 11.5 if _hdr_is_law else 10)
        pdf.set_text_color(0, 0, 0)
        pdf.set_xy(10, 33)
        pdf.cell(pdf.w - 20, 4, "FINAL EXAMINATION TIMETABLE (ACADEMIC YEAR: 2025-26)", 0, 1, 'C')
        
        current_y = 38
        
        # Program Name
        pdf.set_font("Times", 'B', 11.5 if _hdr_is_law else 10)
        pdf.set_xy(10, current_y)
        # SOL Fix: preserve exact capitalisation of program name (e.g. B.A., LL.B.(Hons.))
        # All other colleges continue to use uppercase
        _prog_name = header_content['main_branch_full']
        _is_law = "LAW" in st.session_state.get('selected_college', '').upper()
        _prog_display = _prog_name if _is_law else _prog_name.upper()
        pdf.cell(pdf.w - 20, 4, _prog_display, 0, 1, 'C')
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

        pdf.set_font("Times", 'B', 11.5 if _hdr_is_law else 10)  # SOL Fix: year/sem line
        pdf.set_xy(10, current_y)
        pdf.cell(pdf.w - 20, 4, f"YEAR: {year_roman}, SEMESTER: {sem_roman}".upper(), 0, 1, 'C')
        current_y += 4

        if time_slot:
            pdf.set_font("Times", 'B', 10.5 if _hdr_is_law else 9)  # SOL Fix
            pdf.set_xy(10, current_y)
            pdf.cell(pdf.w - 20, 4, f"EXAM TIME: {time_slot}".upper(), 0, 1, 'C')
            current_y += 4
            
            pdf.set_font("Times", 'BI', 10.5 if _hdr_is_law else 9)  # SOL Fix
            pdf.set_xy(10, current_y)
            pdf.cell(pdf.w - 20, 4, "(CHECK THE SUBJECT EXAM TIME)".upper(), 0, 1, 'C')
            current_y += 4

        pdf.set_xy(pdf.l_margin, header_end_y)

    render_footer()
    render_header()
    
    # SOL Fix 3 & 4: For School of Law, strip program prefix from sub-branch column headers
    # (show stream name only, highlighted in light grey). Also apply Proper Case for Sem X & LL.M Sem II.
    current_college_context_for_header = st.session_state.get('selected_college', '')
    IS_LAW_SCHOOL_HEADER = "LAW" in current_college_context_for_header.upper()

    def sol_clean_header_label(col_label, sem_roman_str):
        """For SOL: strip the program prefix leaving only the stream name.
        For Sem X and LL.M Sem II: return in Proper Case (title case).
        Also normalises any residual program name to canonical capitalisation/spacing."""
        raw = str(col_label)
        # Strip program prefix: everything before and including the LAST ' - '
        # SubBranch was built as "CANONICAL_PROG - STREAM", so split once from right
        if " - " in raw:
            raw = raw.rsplit(" - ", 1)[-1].strip()
        # SOL Fix: preserve exact casing of stream/program names — never force uppercase
        # For Sem X and LL.M Sem II: use Proper Case (title case) for stream names
        sem_upper = str(sem_roman_str).strip().upper()
        if sem_upper in ("X", "10") or "LL.M" in sem_upper or "LLM" in sem_upper:
            return raw.title()
        # All other SOL semesters: return as-is (canonical casing already applied by normaliser)
        return raw

    if IS_LAW_SCHOOL_HEADER:
        sem_roman_for_header = str(header_content.get('semester_roman', '')).strip().upper()
        display_columns = []
        for c in columns:
            if c == "Exam Date":
                display_columns.append("EXAM DATE")
            else:
                display_columns.append(sol_clean_header_label(c, sem_roman_for_header))
        # Use light grey fill for SOL header row
        sol_header_fill = (255, 255, 255)  # SOL Fix: plain white header
        pdf.set_font("Times", 'B', 9.5)
        # Temporarily patch header_bg_color via a custom attribute on the pdf object
        setattr(pdf, '_sol_header_fill', sol_header_fill)
        setattr(pdf, '_sol_header_font_size', 10.5)  # SOL Fix: header font +1 (9.5 -> 10.5)
        print_row_custom(pdf, display_columns, col_widths, line_height=line_height, header=True)
        delattr(pdf, '_sol_header_fill')
        delattr(pdf, '_sol_header_font_size')
    else:
        # Uppercase the table columns (standard non-SOL path — unchanged)
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
            text = str(cell_text) if cell_text is not None else ""
            avail_w = col_widths[i] - 2 
            lines = wrap_text(pdf, text, avail_w)
            wrapped_cells.append(lines)
            max_lines = max(max_lines, len(lines))
        row_h = line_height * max_lines
        
        # Check Page Break
        if pdf.get_y() + row_h > pdf.h - footer_height - 5:
            pdf.add_page()
            render_footer()
            render_header()
            if IS_LAW_SCHOOL_HEADER:
                pdf.set_font("Times", 'B', 9.5)
                setattr(pdf, '_sol_header_fill', sol_header_fill)
                setattr(pdf, '_sol_header_font_size', 10.5)  # SOL Fix: header font +1 (9.5 -> 10.5)
                print_row_custom(pdf, display_columns, col_widths, line_height=line_height, header=True)
                delattr(pdf, '_sol_header_fill')
                delattr(pdf, '_sol_header_font_size')
            else:
                upper_columns = [str(c).upper() for c in columns]
                pdf.set_font("Times", 'B', 9.5) 
                print_row_custom(pdf, upper_columns, col_widths, line_height=line_height, header=True)
            pdf.set_font("Times", '', 9.5)  
        
        print_row_custom(pdf, row, col_widths, line_height=line_height, header=False)

def convert_excel_to_pdf(excel_path, pdf_path, sub_branch_cols_per_page=6, declaration_date=None):
    import uuid
    current_college_context = st.session_state.get('selected_college', '')
    IS_LAW_SCHOOL   = "LAW" in current_college_context.upper()
    IS_BUSINESS_SCH = ("School of Business Management" in current_college_context
                       or "Pravin Dalal" in current_college_context)

    time_slots_dict = st.session_state.get('time_slots', {
        1: {"start": "10:00 AM", "end": "1:00 PM"},
        2: {"start": "2:00 PM",  "end": "5:00 PM"}
    })

    # SOL Fix 2: Override default time slots for School of Law
    if IS_LAW_SCHOOL:
        time_slots_dict = {
            1: {"start": "11:00 AM", "end": "1:00 PM"},
            2: {"start": "2:30 PM",  "end": "4:30 PM"}
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
            romans = {'I':1,'II':2,'III':3,'IV':4,'V':5,'VI':6,
                      'VII':7,'VIII':8,'IX':9,'X':10,'XI':11,'XII':12}
            found = False
            for r_key, r_val in romans.items():
                if s == r_key or s.endswith(f" {r_key}") or s.endswith(f"_{r_key}"):
                    sem_int = r_val; found = True; break
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
    pdf_outputs = {}  # dictionary: filename -> bytes

    # ══════════════════════════════════════════════════════════════════════════
    #  BRANCH A — School of Business Management / Pravin Dalal
    #  Portrait A4, one PDF per program/semester, inline instructions
    # ══════════════════════════════════════════════════════════════════════════
    if IS_BUSINESS_SCH:
        footer_height = 14

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
                    if pdf_obj.get_string_width(test) <= usable_w: cur = test
                    else:
                        if cur: lines.append(cur)
                        cur = w
                if cur: lines.append(cur)
                total += len(lines) * line_h + (1 if is_heading else 0.5)
            return total

        def render_instructions_sbm(pdf_obj, y_start):
            usable_w  = pdf_obj.w - 2 * pdf_obj.l_margin
            available = pdf_obj.h - footer_height - 2 - y_start
            chosen_size, chosen_lh = 6.5, 4.0
            for fs in [8.5, 8.0, 7.5, 7.0, 6.5]:
                lh = fs * 0.55
                if _measure_instructions(pdf_obj, fs, lh, usable_w) <= available:
                    chosen_size = fs; chosen_lh = lh; break
            pdf_obj.set_text_color(0, 0, 0)
            cy = y_start + 1
            for text, is_heading in INSTRS:
                pdf_obj.set_font("Times", 'BU' if is_heading else '', chosen_size)
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

        num_slots  = len(time_slots_dict)
        slot_labels = {sn: f"{scfg['start']} to {scfg['end']}"
                       for sn, scfg in sorted(time_slots_dict.items())}

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
                day    = declaration_date.day
                suffix = ('TH' if 11 <= (day % 100) <= 13
                          else {1:'ST',2:'ND',3:'RD'}.get(day % 10, 'TH'))
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

            text_y = logo_y + 42
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
            pdf_obj.cell(pdf_obj.w - 20, cell_h,
                         "FINAL EXAMINATION TIMETABLE (ACADEMIC YEAR: 2025-26)", 0, 1, 'C')
            text_y += cell_h + LINE_GAP

            prog_name = str(header_content.get('main_branch_full', '')).upper()
            pdf_obj.set_font("Times", 'B', F_PROG)
            cell_h = F_PROG * 0.40
            pdf_obj.set_xy(10, text_y)
            pdf_obj.cell(pdf_obj.w - 20, cell_h, prog_name, 0, 1, 'C')
            text_y += cell_h + LINE_GAP

            sem_roman = str(header_content.get('semester_roman', '')).upper()
            roman_map = {'XII':12,'XI':11,'X':10,'IX':9,'VIII':8,'VII':7,
                         'VI':6,'V':5,'IV':4,'III':3,'II':2,'I':1}
            sem_int = roman_map.get(sem_roman) or 1
            m = re.search(r'(\d+)', sem_roman)
            if m and sem_roman not in roman_map: sem_int = int(m.group(1))
            year_int = (sem_int + 1) // 2

            def _to_roman(n):
                val = [(1000,"M"),(900,"CM"),(500,"D"),(400,"CD"),(100,"C"),
                       (90,"XC"),(50,"L"),(40,"XL"),(10,"X"),(9,"IX"),
                       (5,"V"),(4,"IV"),(1,"I")]
                r = ""
                for v, s in val:
                    while n >= v: r += s; n -= v
                return r

            pdf_obj.set_font("Times", 'B', F_YEAR)
            cell_h = F_YEAR * 0.40
            pdf_obj.set_xy(10, text_y)
            pdf_obj.cell(pdf_obj.w - 20, cell_h,
                         f"YEAR: {_to_roman(year_int)}, TRIMESTER: {sem_roman}", 0, 1, 'C')
            text_y += cell_h + LINE_GAP
            pdf_obj.set_xy(pdf_obj.l_margin, text_y + 3)

        LINE_H, PAD = 5, 1.5

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

        def _draw_row(pdf_obj, cells, col_widths, font_style='', font_size=9,
                      fill_color=None, text_color=(0,0,0)):
            if fill_color: pdf_obj.set_fill_color(*fill_color)
            pdf_obj.set_text_color(*text_color)
            pdf_obj.set_font("Times", font_style, font_size)
            wrapped = [_wrap_cell(pdf_obj, txt, col_widths[i] - 2*PAD, font_style, font_size)
                       for i, txt in enumerate(cells)]
            max_lines = max(len(lns) for lns in wrapped)
            row_h = LINE_H * max_lines
            x0, y0 = pdf_obj.get_x(), pdf_obj.get_y()
            if fill_color: pdf_obj.rect(x0, y0, sum(col_widths), row_h, 'F')
            cx = x0
            for i, lines in enumerate(wrapped):
                total_text_h = len(lines) * LINE_H * 0.85
                pad_v = (row_h - total_text_h) / 2
                for j, ln in enumerate(lines):
                    pdf_obj.set_xy(cx + PAD, y0 + pad_v + j * LINE_H * 0.85)
                    pdf_obj.cell(col_widths[i] - 2*PAD, LINE_H * 0.85, ln, border=0, align='C')
                pdf_obj.rect(cx, y0, col_widths[i], row_h)
                cx += col_widths[i]
            pdf_obj.set_xy(x0, y0 + row_h)
            return row_h

        # ── Process each sheet and build one PDF per program/semester ─────────
        for sheet_name, sheet_df in df_dict.items():
            try:
                if sheet_df.empty or sheet_name in ["No_Data","Daily_Statistics",
                                                    "Summary","Verification","Empty"]: continue
                if re.search(r'_Ele(\d*)$', sheet_name): continue

                main_branch_full = ""
                for col in ("_prog_", "Program", "MainBranch"):
                    if col in sheet_df.columns and not sheet_df[col].dropna().empty:
                        main_branch_full = str(sheet_df[col].dropna().iloc[0]); break
                if not main_branch_full or "Unnamed" in main_branch_full:
                    main_branch_full = sheet_name.split('_|_')[0] if '_|_' in sheet_name else sheet_name

                semester_raw = sheet_name.split('_|_')[1] if '_|_' in sheet_name else sheet_name
                display_sem  = semester_raw.strip()
                for pfx in ("trimester", "semester", "sem", "tri"):
                    if display_sem.lower().startswith(pfx):
                        display_sem = display_sem[len(pfx):].strip(); break

                header_content = {'main_branch_full': main_branch_full, 'semester_roman': display_sem}

                # ── Build slot_pivot from session-state timetable_data ────────
                # In the verification app, timetable_data is the processed_tt
                # dict stored under 'processed_tt' (integers → DataFrames with
                # ExamSlotNumber already derived in process_verification_file).
                timetable_data = st.session_state.get('processed_tt', {})

                def _strip_sem_prefix(s):
                    s = str(s).strip()
                    for pfx in ("trimester", "semester", "sem", "tri"):
                        if s.lower().startswith(pfx): s = s[len(pfx):].strip(); break
                    return s.upper()

                roman_map2 = {'XII':12,'XI':11,'X':10,'IX':9,'VIII':8,'VII':7,
                              'VI':6,'V':5,'IV':4,'III':3,'II':2,'I':1}
                disp_check = display_sem.strip().upper()

                rows_for_sheet = []
                for sem_key, sem_df in timetable_data.items():
                    if isinstance(sem_df, pd.DataFrame) and sem_df.empty: continue
                    sem_str_check = _strip_sem_prefix(sem_key)
                    if sem_str_check != disp_check:
                        sv = roman_map2.get(sem_str_check) or (
                            int(re.search(r'\d+', sem_str_check).group())
                            if re.search(r'\d+', sem_str_check) else None)
                        dv = roman_map2.get(disp_check) or (
                            int(re.search(r'\d+', disp_check).group())
                            if re.search(r'\d+', disp_check) else None)
                        if sv != dv: continue
                    if isinstance(sem_df, pd.DataFrame):
                        matched = sem_df[sem_df['MainBranch'].astype(str) == main_branch_full].copy()
                        if not matched.empty: rows_for_sheet.append(matched)

                slot_pivot = {}

                if rows_for_sheet:
                    combined = pd.concat(rows_for_sheet, ignore_index=True)
                    combined['Exam Date'] = pd.to_datetime(
                        combined['Exam Date'], format="%d-%m-%Y", errors='coerce')
                    combined = combined.dropna(subset=['Exam Date']).sort_values('Exam Date')
                    if 'ExamSlotNumber' not in combined.columns:
                        combined['ExamSlotNumber'] = 1
                    combined['ExamSlotNumber'] = (
                        pd.to_numeric(combined['ExamSlotNumber'], errors='coerce')
                        .fillna(1).astype(int))
                    combined['ExamSlotNumber'] = combined['ExamSlotNumber'].apply(
                        lambda x: x if x in time_slots_dict else 1)

                    for _, row in combined.iterrows():
                        d_str = row['Exam Date'].strftime("%A, %d %B, %Y")
                        sn    = int(row.get('ExamSlotNumber', 1))
                        subj  = str(row.get('Subject', '')).strip()
                        if not subj or subj in ('nan', ''): continue
                        code  = str(row.get('ModuleCode', '')).strip()
                        if code and code.lower() != 'nan': subj = f"{subj} ({code})"
                        oe = str(row.get('OE', '')).strip()
                        if oe and oe not in ('nan', ''): subj = f"{subj} [{oe}]"
                        if d_str not in slot_pivot:
                            slot_pivot[d_str] = {sn2: [] for sn2 in time_slots_dict}
                        slot_pivot[d_str].setdefault(sn, []).append(subj)

                else:
                    # Fallback: parse time suffixes from Excel cells
                    if 'Exam Date' not in sheet_df.columns: continue
                    _meta = re.compile(
                        r'^(Program|Semester|MainBranch|Note|Message|_prog_|_sem_)(\d+)?$',
                        re.IGNORECASE)
                    _sub_cols = [c for c in sheet_df.columns
                                 if c != 'Exam Date' and not _meta.match(str(c))
                                 and pd.notna(c) and str(c).strip() != '']
                    if not _sub_cols: continue

                    _time_to_slot = {}
                    for _sn, _scfg in time_slots_dict.items():
                        _raw = f"{_scfg['start']} - {_scfg['end']}"
                        _time_to_slot[_raw.upper().replace(' ', '')] = _sn

                    def _slot_from_subject(text):
                        m = re.search(
                            r'\[\s*(\d{1,2}:\d{2}\s*[AP]M\s*-\s*\d{1,2}:\d{2}\s*[AP]M)\s*\]',
                            text, re.IGNORECASE)
                        if not m: return None
                        return _time_to_slot.get(m.group(1).upper().replace(' ', ''))

                    for _, row in sheet_df.iterrows():
                        d = str(row.get('Exam Date', '')).strip()
                        if not d or d in ('nan', '', '---'): continue
                        try:
                            d = pd.to_datetime(d, format="%d-%m-%Y",
                                               errors='coerce').strftime("%A, %d %B, %Y")
                        except: pass
                        if d not in slot_pivot:
                            slot_pivot[d] = {sn: [] for sn in time_slots_dict}
                        for sc in _sub_cols:
                            val = str(row.get(sc, '')).strip()
                            for raw_subj in val.split('<hr>'):
                                raw_subj = raw_subj.strip()
                                if not raw_subj or raw_subj in ('nan', '---', ''): continue
                                detected_slot = _slot_from_subject(raw_subj)
                                clean_subj = re.sub(
                                    r'\s*\[\s*\d{1,2}:\d{2}\s*[AP]M\s*-\s*\d{1,2}:\d{2}\s*[AP]M\s*\]',
                                    '', raw_subj, flags=re.IGNORECASE).strip()
                                sn = detected_slot if detected_slot else 1
                                slot_pivot[d].setdefault(sn, []).append(clean_subj)

                if not slot_pivot: continue

                # ── Detect active slots (only cols with data) ─────────────────
                all_slots_sorted = sorted(time_slots_dict.keys())
                active_slots = [sn for sn in all_slots_sorted
                                if any(slot_pivot[d].get(sn) for d in slot_pivot)]
                if not active_slots: active_slots = all_slots_sorted

                n_active       = len(active_slots)
                pdf            = FPDF(orientation='P', unit='mm', format='A4')
                pdf.set_auto_page_break(auto=False, margin=15)
                pdf.alias_nb_pages()

                page_w         = pdf.w - 2 * pdf.l_margin
                date_w         = 38
                act_slot_w     = (page_w - date_w) / n_active
                act_col_widths = [date_w] + [act_slot_w] * n_active
                act_header_cells = ["DAY & DATE"] + ["TIMING & SUBJECT"] * n_active
                act_time_cells   = [""] + [slot_labels[sn] for sn in active_slots]

                pdf.add_page()
                render_footer_sbm(pdf)
                render_header_sbm(pdf, header_content, declaration_date)
                _draw_row(pdf, act_header_cells, act_col_widths, font_style='B', font_size=9.5)
                _draw_row(pdf, act_time_cells,   act_col_widths, font_style='B', font_size=9)

                _instr_h      = _measure_instructions(pdf, 7.5, 7.5*0.55, page_w)
                _table_bottom = pdf.h - footer_height - _instr_h - 4

                pdf.set_font("Times", '', 9.5)
                for d_str, slots in slot_pivot.items():
                    row_cells = [d_str] + [
                        "\n".join(slots.get(sn, [])) if slots.get(sn) else "-------------------"
                        for sn in active_slots
                    ]
                    max_lines_est = max(
                        len(_wrap_cell(pdf, txt, act_col_widths[i] - 2*PAD, '', 9.5))
                        for i, txt in enumerate(row_cells))
                    if pdf.get_y() + (LINE_H * max_lines_est) > _table_bottom:
                        render_instructions_sbm(pdf, pdf.get_y() + 2)
                        pdf.add_page()
                        render_footer_sbm(pdf)
                        render_header_sbm(pdf, header_content, declaration_date)
                        _draw_row(pdf, act_header_cells, act_col_widths, 'B', 9.5)
                        _draw_row(pdf, act_time_cells,   act_col_widths, 'B', 9)
                    _draw_row(pdf, row_cells, act_col_widths, '', 9.5)

                render_instructions_sbm(pdf, pdf.get_y() + 2)

                # Save this program's PDF into the dict
                clean_branch = re.sub(r'[^A-Za-z0-9_\- ]', '', main_branch_full).strip().replace(" ", "_")
                clean_sem    = re.sub(r'[^A-Za-z0-9_\- ]', '', display_sem).strip().replace(" ", "_")
                filename     = f"{clean_branch}_Trimester_{clean_sem}.pdf"
                base_filename = filename
                counter = 1
                while filename in pdf_outputs:
                    filename = f"{base_filename.replace('.pdf', '')}_{counter}.pdf"
                    counter += 1

                temp_path = f"temp_{uuid.uuid4().hex}.pdf"
                pdf.output(temp_path)
                with open(temp_path, "rb") as fh:
                    pdf_outputs[filename] = fh.read()
                os.remove(temp_path)
                sheets_processed += 1

            except Exception as e:
                st.warning(f"Error processing SBM sheet {sheet_name}: {e}")
                continue

    # ══════════════════════════════════════════════════════════════════════════
    #  BRANCH B — All other colleges (original Landscape Legal, single PDF)
    # ══════════════════════════════════════════════════════════════════════════
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
                if IS_LAW_SCHOOL and main_branch_full:
                    prog_upper = main_branch_full.upper()
                    if "LL.M" in prog_upper or "MASTER OF LAW" in prog_upper or "LLM" in prog_upper:
                        sheet_college_name = "Kirit P. Mehta School of Law"
                    elif "B.A." in prog_upper or "B.B.A." in prog_upper or "LL.B" in prog_upper:
                        sheet_college_name = "Kirit P. Mehta School of Law / School of Law"
                    else:
                        sheet_college_name = "Kirit P. Mehta School of Law / School of Law"

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

                if IS_LAW_SCHOOL and is_elective: continue

                if "_sem_" in sheet_df.columns and not sheet_df["_sem_"].dropna().empty:
                    display_sem = str(sheet_df["_sem_"].dropna().iloc[0]).strip()
                else:
                    display_sem = semester_raw.strip()
                    for pfx in ("trimester", "semester", "sem", "tri"):
                        if display_sem.lower().startswith(pfx):
                            display_sem = display_sem[len(pfx):].strip(); break

                header_content   = {'main_branch_full': main_branch_full, 'semester_roman': display_sem}
                header_exam_time = get_header_time_for_semester(f"Sem {display_sem}")

                if not is_elective:
                    if 'Exam Date' not in sheet_df.columns: continue
                    sheet_df = sheet_df.dropna(how='all').reset_index(drop=True)
                    fixed_cols = ["Exam Date"]
                    _meta_pattern = re.compile(
                        r'^(Program|Semester|MainBranch|Note|Message|_prog_|_sem_)(\d+)?$',
                        re.IGNORECASE)
                    sub_branch_cols = [c for c in sheet_df.columns if c not in fixed_cols
                                       and not _meta_pattern.match(str(c))
                                       and pd.notna(c) and str(c).strip() != '']
                    if not sub_branch_cols: continue

                    cols_per_page = 6
                    for start in range(0, len(sub_branch_cols), cols_per_page):
                        chunk         = sub_branch_cols[start:start + cols_per_page]
                        cols_to_print = fixed_cols + chunk
                        chunk_df      = sheet_df[cols_to_print].copy()

                        subset     = chunk_df[chunk].astype(str).apply(lambda x: x.str.strip())
                        valid_cells = (subset != "") & (subset != "nan") & (subset != "---")
                        mask       = valid_cells.any(axis=1)
                        chunk_df   = chunk_df[mask].reset_index(drop=True)
                        if chunk_df.empty: continue

                        try:
                            chunk_df["Exam Date"] = pd.to_datetime(
                                chunk_df["Exam Date"], format="%d-%m-%Y",
                                errors='coerce').dt.strftime("%A, %d %B, %Y")
                        except: pass

                        page_width      = pdf.w - 2 * pdf.l_margin
                        date_col_width  = 30
                        remaining_width = page_width - date_col_width
                        num_sub         = max(len(chunk), 1)
                        sub_width       = remaining_width / num_sub
                        col_widths      = [date_col_width] + [sub_width] * len(chunk)

                        pdf.add_page()
                        original_college = st.session_state.get('selected_college')
                        st.session_state['selected_college'] = sheet_college_name

                        page_time_slot = header_exam_time
                        if IS_LAW_SCHOOL:
                            _time_pat = re.compile(
                                r'\[\s*(\d{1,2}:\d{2}\s*[AP]M\s*-\s*\d{1,2}:\d{2}\s*[AP]M)\s*\]',
                                re.IGNORECASE)
                            _time_counts = {}
                            for _col in chunk:
                                if _col not in chunk_df.columns: continue
                                for _cell in chunk_df[_col].astype(str):
                                    for _t in _time_pat.findall(_cell):
                                        _t_norm = re.sub(r'\s+', ' ', _t.strip().upper())
                                        _time_counts[_t_norm] = _time_counts.get(_t_norm, 0) + 1
                            if _time_counts:
                                page_time_slot = max(_time_counts, key=_time_counts.get)

                            _strip_pat   = re.compile(r'\s*\[\s*\d{1,2}:\d{2}\s*[AP]M\s*-\s*\d{1,2}:\d{2}\s*[AP]M\s*\]', re.IGNORECASE)
                            _extract_pat = re.compile(r'\[\s*(\d{1,2}:\d{2}\s*[AP]M\s*-\s*\d{1,2}:\d{2}\s*[AP]M)\s*\]', re.IGNORECASE)
                            _page_norm   = re.sub(r'\s+', ' ', page_time_slot.strip().upper())

                            def _retag_cell(cell_text):
                                parts = str(cell_text).split(' <hr> ')
                                result = []
                                for part in parts:
                                    m = _extract_pat.search(part)
                                    subj_time = re.sub(r'\s+', ' ', m.group(1).strip().upper()) if m else None
                                    clean = _strip_pat.sub('', part).rstrip()
                                    if subj_time and subj_time != _page_norm:
                                        clean = clean + f' [{m.group(1).strip()}]'
                                    result.append(clean)
                                return ' <hr> '.join(result)

                            for _col in chunk:
                                if _col in chunk_df.columns:
                                    chunk_df[_col] = chunk_df[_col].astype(str).apply(
                                        lambda v: _retag_cell(v) if v not in ('---', 'nan', '') else v)

                        print_table_custom(pdf, chunk_df, cols_to_print, col_widths, line_height=5,
                                           header_content=header_content, Programs=chunk,
                                           time_slot=page_time_slot, actual_time_slots=None,
                                           declaration_date=declaration_date)

                        if original_college: st.session_state['selected_college'] = original_college
                        sheets_processed += 1

                else:
                    target_cols    = ['Exam Date', 'OE Type', 'Open Elective (All Applicable Streams)']
                    available_cols = [c for c in target_cols if c in sheet_df.columns]

                    if len(available_cols) >= 3:
                        sheet_df = sheet_df.dropna(subset=['Exam Date']).reset_index(drop=True)
                        if sheet_df.empty: continue
                        try:
                            sheet_df["Exam Date"] = pd.to_datetime(
                                sheet_df["Exam Date"], format="%d-%m-%Y",
                                errors='coerce').dt.strftime("%A, %d %B, %Y")
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
                pdf.set_font("Times", 'B', 11)
                pdf.cell(0, 5, "CONTROLLER OF EXAMINATIONS", 0, 1, 'L')
                pdf.line(10, pdf.h - 15, 80, pdf.h - 15)
                pdf.set_font("Times", size=9)
                pdf.set_xy(pdf.w - 30, pdf.h - 15)
                pdf.cell(20, 5, f"{pdf.page_no()} of {{nb}}", 0, 0, 'R')
                pdf.set_y(0)
                if os.path.exists(LOGO_PATH): pdf.image(LOGO_PATH, x=(pdf.w-45)/2, y=5, w=45)
                pdf.set_text_color(0, 0, 0)
                pdf.set_font("Times", 'B', 12)
                pdf.set_xy(10, 25)
                pdf.cell(pdf.w - 20, 8,
                         st.session_state.get('selected_college', "SVKM's NMIMS University").upper(),
                         0, 1, 'C')
                pdf.set_font("Times", 'B', 12)
                pdf.set_xy(10, 40)
                pdf.cell(0, 10, "EXAMINATION GUIDELINES", 0, 1, 'C')
                pdf.set_y(60)
                pdf.set_font("Times", 'B', 12)
                pdf.cell(0, 10, "INSTRUCTIONS TO CANDIDATES", 0, 1, 'C')
                pdf.ln(5)
                instrs_final = [
                    "1. Candidates are required to be present at the examination center THIRTY MINUTES before the stipulated time.",
                    "2. Candidates must produce their University Identity Card at the time of the examination.",
                    "3. Candidates are not permitted to enter the examination hall after stipulated time.",
                    "4. Candidates will not be permitted to leave the examination hall during the examination time.",
                    "5. Candidates are forbidden from taking any unauthorized material inside the examination hall. Carrying the same will be treated as usage of unfair means."
                ]
                pdf.set_font("Times", size=12)
                for i in instrs_final:
                    pdf.multi_cell(0, 7, i); pdf.ln(2)
            except Exception: pass

            temp_path = f"temp_{uuid.uuid4().hex}.pdf"
            pdf.output(temp_path)
            with open(temp_path, "rb") as fh:
                pdf_outputs["Timetable.pdf"] = fh.read()
            os.remove(temp_path)

    if sheets_processed == 0:
        st.error("No valid sheets generated in PDF.")
        return {}

    return pdf_outputs


# ==========================================
# 🔄 GENERATE PDF TIMETABLE (ORCHESTRATOR)
# ==========================================
def generate_pdf_timetable(semester_wise_timetable, output_pdf, declaration_date=None):
    import zipfile
    temp_dir   = os.path.dirname(output_pdf) if os.path.dirname(output_pdf) else "."
    temp_excel = os.path.join(temp_dir, "temp_timetable.xlsx")

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
        pdf_dict = convert_excel_to_pdf(temp_excel, output_pdf, declaration_date=declaration_date)
    except Exception as e:
        st.error(f"❌ Error during Excel to PDF conversion: {e}")
        st.error(traceback.format_exc())
        return
    finally:
        try:
            if os.path.exists(temp_excel): os.remove(temp_excel)
        except Exception:
            pass

    if not pdf_dict:
        st.error("❌ No PDFs were generated.")
        return

    current_college = st.session_state.get('selected_college', '')
    IS_BUSINESS_SCH = ("School of Business Management" in current_college
                       or "Pravin Dalal" in current_college)

    try:
        page_number_pattern = re.compile(r'^[\s\n]*(?:Page\s*)?\d+[\s\n]*$')

        # Post-process every PDF — remove blank/page-number-only pages
        final_pdfs = {}
        for filename, pdf_bytes in pdf_dict.items():
            try:
                reader = PdfReader(io.BytesIO(pdf_bytes))
                writer = PdfWriter()
                for page in reader.pages:
                    try:    text = page.extract_text() or ""
                    except: text = ""
                    cleaned = text.strip()
                    if cleaned and not page_number_pattern.match(cleaned) and len(cleaned) > 10:
                        writer.add_page(page)
                if writer.pages:
                    buf = io.BytesIO()
                    writer.write(buf)
                    final_pdfs[filename] = buf.getvalue()
                else:
                    final_pdfs[filename] = pdf_bytes  # keep original if all filtered
            except Exception:
                final_pdfs[filename] = pdf_bytes

        if IS_BUSINESS_SCH and len(final_pdfs) > 1:
            # Multiple PDFs → serve as a ZIP
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                for fname, fbytes in final_pdfs.items():
                    zf.writestr(fname, fbytes)
            zip_buf.seek(0)
            zip_path = output_pdf.replace(".pdf", ".zip")
            with open(zip_path, "wb") as f:
                f.write(zip_buf.getvalue())
            st.success(f"✅ PDF generation complete — {len(final_pdfs)} PDF(s) in ZIP")
            with open(zip_path, "rb") as f:
                st.download_button(
                    "📥 Download All PDFs (ZIP)",
                    f,
                    f"SBM_Timetables_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                    "application/zip",
                    use_container_width=True
                )
            try: os.remove(zip_path)
            except: pass

        else:
            # Single PDF (non-SBM or only one SBM program)
            single_filename = list(final_pdfs.keys())[0]
            pdf_bytes = final_pdfs[single_filename]
            with open(output_pdf, "wb") as f:
                f.write(pdf_bytes)
            st.success(f"✅ PDF generation complete — {len(PdfReader(io.BytesIO(pdf_bytes)).pages)} page(s)")

    except Exception as e:
        st.error(f"❌ PDF post-processing error: {e}")
        st.error(traceback.format_exc())


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
                pdf_filename = "Verification_Timetable.pdf"
                generate_pdf_timetable(
                    st.session_state.processed_tt,
                    pdf_filename,
                    declaration_date=decl_date
                )

                # For SBM/Pravin Dalal, download is handled inside generate_pdf_timetable (ZIP).
                # For all other colleges, a single PDF is written to pdf_filename.
                _current_college = st.session_state.get('selected_college', '')
                _is_sbm = ("School of Business Management" in _current_college
                           or "Pravin Dalal" in _current_college)
                if not _is_sbm and os.path.exists(pdf_filename):
                    st.balloons()
                    with open(pdf_filename, "rb") as f:
                        st.download_button(
                            "📥 Download PDF",
                            f,
                            f"Verification_Timetable_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                            "application/pdf",
                            use_container_width=True
                        )
                    try: os.remove(pdf_filename)
                    except: pass
                elif _is_sbm:
                    st.balloons()


if __name__ == "__main__":
    main()
