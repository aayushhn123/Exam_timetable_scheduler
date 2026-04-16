"""
Re-Exam Auto-Scheduler
======================
• Reads the same input template as the existing Re-Exam Data-to-PDF converter
  (columns: Program, Stream, Academic Year, Current Session,
           Module Abbreviation, Module Description, Subject Type, CM group Number)
• Schedules all re-exams across 10/11 working days (no Sundays, no bank-holidays)
  CONSTRAINT: subjects with the same Module Description are always assigned the
              SAME date AND SAME time-slot, regardless of programme / semester.
• Generates BOTH:
    – Excel output  (same intermediate format as re_exam_to_pdf.py)
    – PDF timetable (pixel-perfect copy of the existing re-exam PDF format)

FIX: Excel pivot now groups all SubBranch columns for a programme on a single sheet
     (up to cols_per_page=6), matching re_exam_to_pdf.py exactly. Academic Year
     column name is normalised so year-suffix deduplication works correctly.
"""

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, date as _date
from fpdf import FPDF
import os
import re
import io
import traceback
from PyPDF2 import PdfReader, PdfWriter

# ── Streamlit compat ──────────────────────────────────────────────────────────
if hasattr(st, "dialog"):
    dialog_decorator = st.dialog
else:
    dialog_decorator = st.experimental_dialog

st.set_page_config(
    page_title="Re-Exam Auto-Scheduler",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Colleges ──────────────────────────────────────────────────────────────────
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
    {"name": "School of Aviation", "icon": "✈️"},
]

LOGO_PATH = "logo.png"
wrap_text_cache = {}

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    * { font-family: 'Inter', sans-serif; }
    .main-header {
        background: linear-gradient(135deg, #951C1C 0%, #C73E1D 100%);
        padding: 2.5rem; border-radius: 16px; margin-bottom: 2rem;
        box-shadow: 0 8px 24px rgba(0,0,0,0.12);
    }
    .main-header h1 { color:white; text-align:center; margin:0; font-size:2.4rem; font-weight:700; text-shadow:2px 2px 4px rgba(0,0,0,0.3); }
    .main-header p  { color:#FFF; text-align:center; margin:0.5rem 0 0; font-size:1.05rem; opacity:.95; font-weight:500; }
    .upload-section { background:#f8f9fa; padding:2.5rem; border-radius:16px; border:2px dashed #951C1C; margin:1rem 0; text-align:center; }
    .stButton>button { transition:all .3s; border-radius:12px; font-weight:600; }
    .stButton>button:hover { transform:translateY(-2px); box-shadow:0 8px 16px rgba(149,28,28,.3); }
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — INPUT PARSING
# ═══════════════════════════════════════════════════════════════════════════════

def get_sem_int(val):
    s = str(val).upper().strip()
    m = re.search(r'\b(\d+)\b', s)
    if m: return int(m.group(1))
    roman_map = {'XII':12,'XI':11,'X':10,'IX':9,'VIII':8,'VII':7,'VI':6,'V':5,'IV':4,'III':3,'II':2,'I':1}
    for r, i in roman_map.items():
        if re.search(r'\b' + r + r'\b', s):
            return i
    return 1


def parse_input_file(uploaded_file):
    """
    Accepts the same Excel template used by the existing re-exam data-to-PDF converter.
    Returns (df_all, error_msg).
    """
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet = xls.sheet_names[0]
        df = pd.read_excel(uploaded_file, sheet_name=sheet)
        df.columns = df.columns.str.strip()

        # Flexible column mapping — keep 'Academic Year' with its space (matches re_exam_to_pdf.py)
        col_map = {
            'Programme':          'Program',
            'Module Description': 'Subject',
            'Module Abbreviation':'ModuleCode',
            'Current Session':    'CurrentSession',
            'Subject Type':       'SubjectType',
            'CM group Number':    'CMGroup',
        }
        df = df.rename(columns={k: v for k, v in col_map.items() if k in df.columns})
        # Keep 'Academic Year' as-is (with space) — same as re_exam_to_pdf.py

        # Collapse duplicate columns that result from renaming
        if df.columns.duplicated().any():
            seen = {}
            new_cols = []
            for i, col in enumerate(df.columns):
                if col not in seen:
                    seen[col] = i
                    new_cols.append(col)
                else:
                    first_idx = seen[col]
                    first_col = df.iloc[:, first_idx]
                    dup_col   = df.iloc[:, i]
                    df.iloc[:, first_idx] = first_col.where(
                        first_col.notna() & (first_col.astype(str).str.strip() != ''),
                        other=dup_col
                    )
                    new_cols.append(f'_DROP_{i}')
            df.columns = new_cols
            df = df.loc[:, ~df.columns.str.startswith('_DROP_')]

        for required in ['Program', 'Subject', 'CurrentSession']:
            if required not in df.columns:
                return None, f"Missing required column: {required}"

        # Fill down merged cells
        for c in ['Program', 'Stream', 'CurrentSession', 'Academic Year']:
            if c in df.columns:
                df[c] = df[c].ffill()

        for c in df.columns:
            df[c] = df[c].fillna('').astype(str).str.strip()

        df['Subject'] = df['Subject'].replace({'': None, 'nan': None})
        df = df[df['Subject'].notna()].copy()

        df['Semester']   = df['CurrentSession'].apply(get_sem_int)
        df['MainBranch'] = df['Program']
        df['SubBranch']  = df['Stream'] if 'Stream' in df.columns else df['Program']
        df['SubBranch']  = df.apply(
            lambda r: r['MainBranch'] if r['SubBranch'] in ['nan', ''] else r['SubBranch'], axis=1
        )

        # Normalise OE flag
        df['OE'] = df['SubjectType'].apply(
            lambda x: 'OE' if re.match(r'^OE', str(x).strip(), re.IGNORECASE) else None
        ) if 'SubjectType' in df.columns else None

        # CMGroup
        if 'CMGroup' in df.columns:
            df['CMGroup'] = df['CMGroup'].apply(
                lambda x: '' if str(x).strip() in ['0', '0.0', 'nan', 'None', ''] else str(x).split('.')[0].strip()
            )
        else:
            df['CMGroup'] = ''

        # ModuleCode for SOL
        _selected_college = st.session_state.get('selected_college', '')
        if 'LAW' in _selected_college.upper() and 'ModuleCode' in df.columns:
            pass  # already mapped
        else:
            df['ModuleCode'] = df.get('ModuleCode', pd.Series([''] * len(df), index=df.index))

        return df, None

    except Exception as e:
        return None, str(e)


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — SCHEDULING ENGINE
# ═══════════════════════════════════════════════════════════════════════════════

def get_valid_dates(start: datetime, num_days: int, holidays: set) -> list:
    """Return num_days working dates (Mon–Sat, excl. holidays) from start."""
    valid, cur = [], start
    while len(valid) < num_days:
        if cur.weekday() != 6 and cur.date() not in holidays:   # 6 = Sunday
            valid.append(cur)
        cur += timedelta(days=1)
    return valid


def schedule_reexams(df: pd.DataFrame, start_date: datetime,
                     num_days: int, holidays: set,
                     time_slots_dict: dict) -> pd.DataFrame:
    """
    Core scheduling logic.

    KEY CONSTRAINT
    ──────────────
    Any two rows that share the SAME normalised Subject name must get the SAME
    (date, time_slot) assignment regardless of programme / semester / academic year.

    Algorithm
    ─────────
    1. Separate OE subjects from core subjects.
    2. Build a list of unique *subject names* for core (de-duplicated).
    3. Get available dates.  Reserve last 2 days for OE if ≥4 days available.
    4. Assign dates round-robin across available days.
    5. Map each row's date back via the subject-name key.
    6. Assign time-slot per semester (odd sem → slot 1, even sem → slot 2).
       But if the same subject spans odd+even sems, use slot 1 (morning).
    """
    df = df.copy()
    df['Exam Date'] = ''
    df['Exam Time'] = ''

    IS_LAW = "LAW" in st.session_state.get('selected_college', '').upper()
    if IS_LAW:
        time_slots_dict = {
            1: {"start": "11:00 AM", "end": "1:00 PM"},
            2: {"start": "2:30 PM",  "end": "4:30 PM"},
        }

    def slot_for_sem(sem_int):
        slot_id = 1 if ((sem_int + 1) // 2) % 2 == 1 else 2
        cfg = time_slots_dict.get(slot_id, time_slots_dict[1])
        return f"{cfg['start']} - {cfg['end']}"

    valid_dates = get_valid_dates(start_date, num_days, holidays)

    # Split OE / core
    oe_mask   = df['OE'].notna() & (df['OE'].astype(str).str.strip() != '')
    core_mask = ~oe_mask

    # ── CORE scheduling ──────────────────────────────────────────────────────
    df_core = df[core_mask].copy()

    if not df_core.empty:
        if len(valid_dates) >= 4 and df[oe_mask].shape[0] > 0:
            core_dates = valid_dates[:-2]
            oe_dates   = valid_dates[-2:]
        else:
            core_dates = valid_dates
            oe_dates   = valid_dates[-1:] if valid_dates else []

        norm_subject = df_core['Subject'].str.strip().str.upper()
        unique_subjects = list(dict.fromkeys(norm_subject.tolist()))

        if not core_dates:
            df.loc[core_mask, 'Exam Date'] = 'Not Scheduled'
            df.loc[core_mask, 'Exam Time'] = ''
        else:
            subject_date_map = {}
            for idx, subj in enumerate(unique_subjects):
                assigned_date = core_dates[idx % len(core_dates)]
                subject_date_map[subj] = assigned_date.strftime('%d-%m-%Y')

            subject_sem_map = {}
            for _, row in df_core.iterrows():
                key = row['Subject'].strip().upper()
                sem = row['Semester']
                if key not in subject_sem_map:
                    subject_sem_map[key] = sem
                else:
                    existing_slot = ((subject_sem_map[key] + 1) // 2) % 2
                    new_slot      = ((sem + 1) // 2) % 2
                    if new_slot == 1:
                        subject_sem_map[key] = sem

            subject_time_map = {}
            for subj, sem in subject_sem_map.items():
                subject_time_map[subj] = slot_for_sem(sem)

            for i, row in df_core.iterrows():
                key = row['Subject'].strip().upper()
                df.at[i, 'Exam Date'] = subject_date_map.get(key, 'Not Scheduled')
                df.at[i, 'Exam Time'] = subject_time_map.get(key, slot_for_sem(row['Semester']))

    # ── OE scheduling ────────────────────────────────────────────────────────
    df_oe = df[oe_mask].copy()
    if not df_oe.empty:
        oe_dates_avail = oe_dates if oe_dates else valid_dates[-1:]
        unique_oe_subjects = list(dict.fromkeys(
            df_oe['Subject'].str.strip().str.upper().tolist()
        ))
        oe_date_map = {}
        for idx, subj in enumerate(unique_oe_subjects):
            oe_date_map[subj] = oe_dates_avail[idx % len(oe_dates_avail)].strftime('%d-%m-%Y')

        oe_time_map = {}
        for _, row in df_oe.iterrows():
            key = row['Subject'].strip().upper()
            if key not in oe_time_map:
                oe_time_map[key] = slot_for_sem(row['Semester'])

        for i, row in df_oe.iterrows():
            key = row['Subject'].strip().upper()
            df.at[i, 'Exam Date'] = oe_date_map.get(key, 'Not Scheduled')
            df.at[i, 'Exam Time'] = oe_time_map.get(key, slot_for_sem(row['Semester']))

    return df


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — BUILD SEMESTER-WISE TIMETABLE
# ═══════════════════════════════════════════════════════════════════════════════

def build_semester_timetable(scheduled_df: pd.DataFrame) -> dict:
    timetable = {}
    for sem in sorted(scheduled_df['Semester'].unique()):
        timetable[sem] = scheduled_df[scheduled_df['Semester'] == sem].copy()
    return timetable


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — EXCEL ENGINE
# Verbatim logic from re_exam_to_pdf.py — key fix: use 'Academic Year' (with space)
# ═══════════════════════════════════════════════════════════════════════════════

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
    key_map, norm_to_key, used_keys = {}, {}, set()
    for name in all_program_names:
        norm = re.sub(r'\s+', ' ', name.strip())
        if norm in norm_to_key:
            key_map[name] = norm_to_key[norm]
            continue
        base = _make_program_abbrev(name)
        key  = base; n = 2
        while key in used_keys:
            key = base[:10] + str(n); n += 1
        used_keys.add(key); norm_to_key[norm] = key; key_map[name] = key
    return key_map


def save_to_excel(semester_wise_timetable):
    """
    Build structured Excel from scheduled re-exam data.
    MATCHES re_exam_to_pdf.py exactly:
      - 'Academic Year' column (with space) for year deduplication
      - One sheet per (programme, semester); up to cols_per_page=6 SubBranch columns
      - Multi-subject same-date cells joined with ' <hr> '
      - OE subjects joined with ', '
    """
    time_slots_dict = st.session_state.get('time_slots', {
        1: {"start": "10:00 AM", "end": "1:00 PM"},
        2: {"start": "2:00 PM",  "end": "5:00 PM"},
    })

    current_college_context = st.session_state.get('selected_college', '')
    IS_LAW_SCHOOL = "LAW" in current_college_context.upper()

    if IS_LAW_SCHOOL:
        time_slots_dict = {
            1: {"start": "11:00 AM", "end": "1:00 PM"},
            2: {"start": "2:30 PM",  "end": "4:30 PM"},
        }

    output = io.BytesIO()

    # SOL merge (verbatim from re_exam_to_pdf.py)
    if IS_LAW_SCHOOL:
        def _sol_normalise(raw):
            s = str(raw).strip(); su = s.upper().replace(' ', '')
            if re.search(r'B\.?B\.?A\.?', su): return 'B.B.A., LL.B.(Hons.)'
            if re.search(r'^B\.?A\.?', su):    return 'B.A., LL.B.(Hons.)'
            return s

        def _is_ba_bba_llb(p):
            p_str = str(p).strip().upper()
            if re.search(r'(LL\.M|MASTER\s+OF\s+LAW|LLM)', p_str): return False
            return bool(re.search(r'^(B\.A\.|B\.B\.A\.)[,\s].*LL\.B', p_str))

        for sem in semester_wise_timetable:
            df_sem = semester_wise_timetable[sem]
            if df_sem.empty: continue
            mask = df_sem['MainBranch'].apply(_is_ba_bba_llb)
            if mask.any():
                norm_prefix = df_sem.loc[mask, 'MainBranch'].apply(_sol_normalise)
                df_sem.loc[mask, 'SubBranch']  = norm_prefix + ' - ' + df_sem.loc[mask, 'SubBranch']
                df_sem.loc[mask, 'MainBranch'] = 'B.A., LL.B.(Hons.) / B.B.A., LL.B.(Hons.)'

    all_programs = []
    for df_s in semester_wise_timetable.values():
        if not df_s.empty:
            all_programs.extend(df_s['MainBranch'].dropna().unique().tolist())
    seen_p, deduped = set(), []
    for p in all_programs:
        if p not in seen_p: deduped.append(p); seen_p.add(p)
    prog_key_map = _build_program_key_map(deduped)

    used_sheet_names = set()
    def unique_sheet(name):
        base, n = name, 2
        while name in used_sheet_names:
            name = base[:31 - len(str(n))] + str(n); n += 1
        used_sheet_names.add(name); return name

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        sheets_created = 0
        for sem, df_sem in semester_wise_timetable.items():
            if df_sem.empty: continue

            roman_sem   = int_to_roman(sem)
            slot_id     = 1 if ((sem + 1) // 2) % 2 == 1 else 2
            p_cfg       = time_slots_dict.get(slot_id, time_slots_dict[1])
            header_norm = normalize_time(f"{p_cfg['start']} - {p_cfg['end']}")

            # ── GROUP ALL PROGRAMMES FOR THIS SEMESTER ONTO ONE SHEET ────────
            # Streams (SubBranch) across all MainBranch values become columns.
            # _prog_ is set to the first MainBranch value (or a shared prefix).

            def shorten_year(y):
                y = str(y).strip()
                m = re.findall(r'\d{4}', y)
                if len(m) >= 2: return f"{m[0][2:]}-{m[1][2:]}"
                elif len(m) == 1: return m[0][2:]
                return y

            all_main_branches = df_sem['MainBranch'].unique().tolist()
            # Derive a combined programme name: common prefix across all MainBranch values,
            # stripped of trailing separators (dash, comma, space).
            def _common_prefix(names):
                if not names: return "Programme"
                if len(names) == 1: return names[0]
                prefix = names[0]
                for n in names[1:]:
                    new_p = ""
                    for a, b in zip(prefix, n):
                        if a == b: new_p += a
                        else: break
                    prefix = new_p
                return re.sub(r'[\s,\-\u2013/]+$', '', prefix).strip() or names[0]
            combined_prog = _common_prefix(all_main_branches)
            prog_key   = prog_key_map.get(combined_prog, combined_prog[:12])
            core_sheet = unique_sheet(f"{prog_key}_|_Sem {roman_sem}"[:31])
            elec_sheet = unique_sheet(f"{prog_key}_|_Sem {roman_sem}_Ele"[:31])

            df_core_all = df_sem[df_sem['OE'].isna()].copy()
            df_elec_all = df_sem[df_sem['OE'].notna()].copy()

            # Build a display column label: if SubBranch == MainBranch, use MainBranch;
            # otherwise use SubBranch (the stream name).
            def col_label(row):
                mb = str(row['MainBranch']).strip()
                sb = str(row['SubBranch']).strip()
                return sb if sb and sb not in ['nan', '', mb] else mb

            if not df_core_all.empty:
                df_core_all['_ColLabel'] = df_core_all.apply(col_label, axis=1)

                ay_col = 'Academic Year' if 'Academic Year' in df_core_all.columns else None

                dedup_rows = []
                group_keys = ['Subject', '_ColLabel', 'Exam Date', 'Exam Time']

                for gkey, grp in df_core_all.groupby(group_keys, sort=False):
                    subj        = gkey[0]
                    actual_time = str(gkey[3]).strip()

                    if ay_col:
                        raw_years   = grp[ay_col].dropna().astype(str).str.strip().unique().tolist()
                        raw_years   = [y for y in raw_years if y.lower() not in ['nan','none','']]
                        short_years = sorted(set(shorten_year(y) for y in raw_years))
                    else:
                        short_years = []

                    time_suffix = ""
                    if actual_time and actual_time.lower() not in ['tbd','nan','']:
                        if IS_LAW_SCHOOL:
                            time_suffix = f" [{actual_time}]"
                        elif normalize_time(actual_time) != header_norm:
                            time_suffix = f" [{actual_time}]"

                    code = ''
                    if IS_LAW_SCHOOL and 'ModuleCode' in grp.columns:
                        _codes = grp['ModuleCode'].dropna().astype(str).str.strip()
                        _codes = [c for c in _codes if c and c.lower() not in ['nan','']]
                        if _codes: code = _codes[0]

                    txt = subj
                    if IS_LAW_SCHOOL and code: txt += f" ({code})"
                    if short_years: txt += f" ({', '.join(short_years)})"
                    txt += time_suffix

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

                df_core2 = pd.DataFrame(dedup_rows)
                df_core2['Exam Date'] = pd.to_datetime(
                    df_core2['Exam Date'], format='%d-%m-%Y', dayfirst=True, errors='coerce'
                )
                df_core2 = df_core2.sort_values(by=['Exam Date','_SortTime'])

                try:
                    pivot = df_core2.groupby(['Exam Date','SubBranch']).agg(
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
                    pivot['_prog_'] = combined_prog
                    pivot['_sem_']  = roman_sem
                    pivot.to_excel(writer, sheet_name=core_sheet, index=False)
                    sheets_created += 1
                except Exception:
                    pass

            # ── OPEN ELECTIVES (combined across all programmes) ──────────────
            if not df_elec_all.empty:
                e_displays = []
                for _, row in df_elec_all.iterrows():
                    subj        = row['Subject']
                    actual_time = str(row.get('Exam Time', '')).strip()
                    time_suffix = ""
                    if actual_time and normalize_time(actual_time) != header_norm \
                            and actual_time.lower() not in ['tbd','nan','']:
                        time_suffix = f" [{actual_time}]"
                    e_displays.append(f"{subj}{time_suffix}")

                df_elec_all = df_elec_all.copy()
                df_elec_all['DisplaySubject'] = e_displays

                try:
                    df_elec_all['Exam Date'] = pd.to_datetime(
                        df_elec_all['Exam Date'], format='%d-%m-%Y', dayfirst=True, errors='coerce'
                    )
                    df_elec_all = df_elec_all.sort_values(by='Exam Date')
                    df_elec_all['Exam Date'] = df_elec_all['Exam Date'].apply(
                        lambda x: x.strftime('%d-%m-%Y') if pd.notna(x) else ""
                    )
                    ep = df_elec_all.groupby(['Exam Date','OE']).agg(
                        {'DisplaySubject': lambda x: ", ".join(sorted(set(x)))}
                    ).reset_index()
                    ep.rename(columns={
                        'OE': 'OE Type',
                        'DisplaySubject': 'Open Elective (All Applicable Streams)'
                    }, inplace=True)
                    ep['_prog_'] = combined_prog
                    ep['_sem_']  = roman_sem
                    ep.to_excel(writer, sheet_name=elec_sheet, index=False)
                    sheets_created += 1
                except Exception:
                    pass

        if sheets_created == 0:
            pd.DataFrame({'Info': ['No valid data']}).to_excel(writer, sheet_name="Empty")

    output.seek(0)
    return output


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 5 — FPDF ENGINE  (verbatim copy from re_exam_to_pdf.py — NO changes)
# ═══════════════════════════════════════════════════════════════════════════════

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
            if current_line: lines.append(current_line); current_line = ""
            lines.append("<hr>"); continue

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
    cell_padding      = 1
    header_bg_color   = (255, 255, 255)
    header_text_color = (0, 0, 0)
    alt_row_color     = (255, 255, 255)

    if header and hasattr(pdf, '_sol_header_fill'):
        header_bg_color = pdf._sol_header_fill

    row_number = getattr(pdf, '_row_counter', 0)
    base_font  = "Times"

    if header:
        base_style = 'B'
        base_size  = getattr(pdf, '_sol_header_font_size', 9.5)
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

    row_h          = line_height * max_lines
    text_line_height = line_height * 0.75
    x0, y0        = pdf.get_x(), pdf.get_y()

    pdf.rect(x0, y0, sum(col_widths), row_h, 'F')

    time_pattern = re.compile(
        r'([\[\(]\s*\d{1,2}:\d{2}\s*[AP]M\s*-\s*\d{1,2}:\d{2}\s*[AP]M\s*[\]\)])',
        re.IGNORECASE
    )

    for i, lines in enumerate(wrapped_cells):
        cx = pdf.get_x()

        subjects_lines = []
        current_subject = []
        for ln in lines:
            if ln == "<hr>":
                subjects_lines.append(current_subject); current_subject = []
            else:
                current_subject.append(ln)
        subjects_lines.append(current_subject)

        num_subjects = len(subjects_lines)
        part_h = row_h / num_subjects if num_subjects > 0 else row_h

        for sub_idx, subj_lines in enumerate(subjects_lines):
            total_text_h = len(subj_lines) * text_line_height
            pad_v        = (part_h - total_text_h) / 2

            for j, ln in enumerate(subj_lines):
                parts = time_pattern.split(ln)
                if len(parts) == 1 or header:
                    pdf.set_xy(cx + cell_padding,
                               y0 + (sub_idx * part_h) + pad_v + j * text_line_height)
                    pdf.cell(col_widths[i] - 2*cell_padding, text_line_height, ln, border=0, align='C')
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
                        pdf.cell(w + 2*pdf.c_margin, text_line_height, p, border=0, align='L')
                        current_x += w

                    pdf.set_font(base_font, base_style, base_size)

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
    header_end_y  = 60

    def render_footer():
        pdf.set_xy(10, pdf.h - footer_height)
        pdf.set_font("Times", 'B', 11)
        pdf.cell(0, 5, "CONTROLLER OF EXAMINATIONS", 0, 1, 'L')
        pdf.line(10, pdf.h - footer_height + 5, 80, pdf.h - footer_height + 5)
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
            suffix = 'TH' if 11 <= (day % 100) <= 13 else {1:'ST',2:'ND',3:'RD'}.get(day % 10, 'TH')
            decl_str = f"DATE: {day}{suffix} {declaration_date.strftime('%B, %Y')}".upper()
            pdf.set_font("Times", 'B', 12)
            pdf.set_text_color(0, 0, 0)
            pdf.set_xy(pdf.w - 80, 8)
            pdf.cell(70, 10, decl_str, 0, 0, 'R')

        logo_width = 45
        logo_x = (pdf.w - logo_width) / 2
        if os.path.exists(LOGO_PATH):
            pdf.image(LOGO_PATH, x=logo_x, y=5, w=logo_width)

        pdf.set_text_color(0, 0, 0)
        college_name = st.session_state.get('selected_college', "SVKM's NMIMS University").upper()
        pdf.set_font("Times", 'B', 12)
        pdf.set_xy(10, 25)
        pdf.cell(pdf.w - 20, 6, college_name, 0, 1, 'C')

        _hdr_is_law = "LAW" in st.session_state.get('selected_college', '').upper()

        pdf.set_font("Times", 'B', 11.5 if _hdr_is_law else 10)
        pdf.set_text_color(0, 0, 0)
        pdf.set_xy(10, 33)
        pdf.cell(pdf.w - 20, 4, "RE-EXAMINATION TIMETABLE (ACADEMIC YEAR: 2025-26)", 0, 1, 'C')

        current_y = 38
        pdf.set_font("Times", 'B', 11.5 if _hdr_is_law else 10)
        pdf.set_xy(10, current_y)
        _prog_name    = header_content['main_branch_full']
        _prog_display = _prog_name if _hdr_is_law else _prog_name.upper()
        pdf.cell(pdf.w - 20, 4, _prog_display, 0, 1, 'C')
        current_y += 4

        sem_roman = str(header_content['semester_roman']).upper()
        roman_map = {'XII':12,'XI':11,'X':10,'IX':9,'VIII':8,'VII':7,'VI':6,'V':5,'IV':4,'III':3,'II':2,'I':1}
        sem_int   = roman_map.get(sem_roman)
        if not sem_int:
            m = re.search(r'(\d+)', sem_roman)
            sem_int = int(m.group(1)) if m else 1
        year_roman = int_to_roman((sem_int + 1) // 2)

        pdf.set_font("Times", 'B', 11.5 if _hdr_is_law else 10)
        pdf.set_xy(10, current_y)
        pdf.cell(pdf.w - 20, 4, f"YEAR: {year_roman}, SEMESTER: {sem_roman}".upper(), 0, 1, 'C')
        current_y += 4

        if time_slot:
            pdf.set_font("Times", 'B', 10.5 if _hdr_is_law else 9)
            pdf.set_xy(10, current_y)
            pdf.cell(pdf.w - 20, 4, f"EXAM TIME: {time_slot}".upper(), 0, 1, 'C')
            current_y += 4
            pdf.set_font("Times", 'BI', 10.5 if _hdr_is_law else 9)
            pdf.set_xy(10, current_y)
            pdf.cell(pdf.w - 20, 4, "(CHECK THE SUBJECT EXAM TIME)".upper(), 0, 1, 'C')
            current_y += 4

        pdf.set_xy(pdf.l_margin, header_end_y)

    render_footer()
    render_header()

    IS_LAW_SCHOOL_HEADER = "LAW" in st.session_state.get('selected_college', '').upper()

    def sol_clean_header_label(col_label, sem_roman_str):
        raw = str(col_label)
        if " - " in raw: raw = raw.rsplit(" - ", 1)[-1].strip()
        sem_upper = str(sem_roman_str).strip().upper()
        if sem_upper in ("X", "10") or "LL.M" in sem_upper or "LLM" in sem_upper:
            return raw.title()
        return raw

    if IS_LAW_SCHOOL_HEADER:
        sem_roman_for_header = str(header_content.get('semester_roman', '')).strip().upper()
        display_columns = []
        for c in columns:
            if c == "Exam Date": display_columns.append("EXAM DATE")
            else: display_columns.append(sol_clean_header_label(c, sem_roman_for_header))
        sol_header_fill = (255, 255, 255)
        pdf.set_font("Times", 'B', 9.5)
        setattr(pdf, '_sol_header_fill', sol_header_fill)
        setattr(pdf, '_sol_header_font_size', 10.5)
        print_row_custom(pdf, display_columns, col_widths, line_height=line_height, header=True)
        delattr(pdf, '_sol_header_fill')
        delattr(pdf, '_sol_header_font_size')
    else:
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
            text   = str(cell_text) if cell_text is not None else ""
            lines  = wrap_text(pdf, text, col_widths[i] - 2)
            wrapped_cells.append(lines)
            max_lines = max(max_lines, len(lines))
        row_h = line_height * max_lines

        if pdf.get_y() + row_h > pdf.h - footer_height - 5:
            pdf.add_page()
            render_footer()
            render_header()
            if IS_LAW_SCHOOL_HEADER:
                pdf.set_font("Times", 'B', 9.5)
                setattr(pdf, '_sol_header_fill', sol_header_fill)
                setattr(pdf, '_sol_header_font_size', 10.5)
                print_row_custom(pdf, display_columns, col_widths, line_height=line_height, header=True)
                delattr(pdf, '_sol_header_fill')
                delattr(pdf, '_sol_header_font_size')
            else:
                upper_columns = [str(c).upper() for c in columns]
                pdf.set_font("Times", 'B', 9.5)
                print_row_custom(pdf, upper_columns, col_widths, line_height=line_height, header=True)
            pdf.set_font("Times", '', 9.5)

        print_row_custom(pdf, row, col_widths, line_height=line_height, header=False)


def _ordinal_suffix(day):
    if 11 <= (day % 100) <= 13: return 'th'
    return {1:'st',2:'nd',3:'rd'}.get(day % 10, 'th')


def convert_excel_to_pdf(excel_path, pdf_path, declaration_date=None,
                         portal_dates=None, all_semesters=None):
    """
    Verbatim from re_exam_to_pdf.py — no structural changes.
    cols_per_page=6: up to 6 SubBranch columns fit on one landscape Legal page.
    Multiple streams of the same programme share one page (or span to next only
    if >6 streams).
    """
    pdf = FPDF(orientation='L', unit='mm', format='Legal')
    pdf.set_auto_page_break(auto=False, margin=15)
    pdf.alias_nb_pages()

    time_slots_dict = st.session_state.get('time_slots', {
        1: {"start": "10:00 AM", "end": "1:00 PM"},
        2: {"start": "2:00 PM",  "end": "5:00 PM"},
    })
    IS_LAW_SCHOOL = "LAW" in st.session_state.get('selected_college', '').upper()
    if IS_LAW_SCHOOL:
        time_slots_dict = {
            1: {"start": "11:00 AM", "end": "1:00 PM"},
            2: {"start": "2:30 PM",  "end": "4:30 PM"},
        }

    def get_header_time_for_semester(sem_str):
        try:
            s = str(sem_str).strip().upper(); sem_int = 1
            romans = {'I':1,'II':2,'III':3,'IV':4,'V':5,'VI':6,'VII':7,'VIII':8,'IX':9,'X':10,'XI':11,'XII':12}
            found = False
            for r_key, r_val in romans.items():
                if s == r_key or s.endswith(f" {r_key}") or s.endswith(f"_{r_key}"):
                    sem_int = r_val; found = True; break
            if not found:
                digits = re.findall(r'\d+', s)
                if digits: sem_int = int(digits[0])
            slot_num = 1 if ((sem_int + 1) // 2) % 2 == 1 else 2
            slot_cfg = time_slots_dict.get(slot_num, time_slots_dict.get(1))
            return f"{slot_cfg['start']} - {slot_cfg['end']}"
        except:
            return f"{time_slots_dict[1]['start']} - {time_slots_dict[1]['end']}"

    # ── Instructions page ────────────────────────────────────────────────────
    try:
        pdf.add_page()
        footer_height = 14

        pdf.set_xy(10, pdf.h - footer_height)
        pdf.set_font("Times", 'B', 11)
        pdf.cell(0, 5, "CONTROLLER OF EXAMINATIONS", 0, 1, 'L')
        pdf.line(10, pdf.h - footer_height + 5, 80, pdf.h - footer_height + 5)
        page_text  = f"{pdf.page_no()} of {{nb}}"
        text_width = pdf.get_string_width(page_text.replace("{nb}", "99"))
        pdf.set_font("Times", size=8)
        pdf.set_text_color(0, 0, 0)
        pdf.set_xy(pdf.w - 10 - text_width, pdf.h - footer_height + 5)
        pdf.cell(text_width, 5, page_text, 0, 0, 'R')

        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Times", 'B', 12)
        pdf.set_xy(10, 25)
        pdf.cell(pdf.w - 20, 6,
                 st.session_state.get('selected_college', "SVKM's NMIMS University").upper(),
                 0, 1, 'C')

        pdf.set_font("Times", 'B', 12)
        pdf.set_xy(10, 33)
        pdf.cell(pdf.w - 20, 4, "RE-EXAMINATION TIMETABLE (ACADEMIC YEAR: 2025-26)", 0, 1, 'C')

        if all_semesters:
            _roman_map = {1:'I',2:'II',3:'III',4:'IV',5:'V',6:'VI',7:'VII',8:'VIII',9:'IX',10:'X',11:'XI',12:'XII'}
            _sem_romans = '/'.join(_roman_map.get(s, str(s)) for s in sorted(all_semesters))
            pdf.ln(2)
            pdf.set_font("Times", 'B', 12)
            pdf.cell(pdf.w - 20, 4, f"SEMESTER - {_sem_romans}", 0, 1, 'C')

        pdf.ln(2)
        pdf.set_font("Times", 'BU', 13)
        pdf.set_text_color(255, 0, 0)
        pdf.cell(pdf.w - 20, 7, "IMPORTANT INSTRUCTIONS TO CANDIDATES", 0, 1, 'C')
        pdf.ln(4)

        margin_l = pdf.l_margin
        text_w   = pdf.w - margin_l - pdf.r_margin

        pdf.set_font("Times", '', 11)
        pdf.set_text_color(0, 0, 0)
        pdf.set_x(margin_l)
        pdf.write(6, "1.  All the eligible students are hereby informed to apply for the respective re-examination/s ")
        pdf.set_font("Times", 'B', 11)
        pdf.write(6, "only through SAP Student portal")
        pdf.set_font("Times", '', 11)
        pdf.write(6, " by using the available online payment facility.")
        pdf.ln(6); pdf.ln(1)

        pdf.set_x(margin_l + 8)
        pdf.set_font("Times", 'U', 11)
        pdf.set_text_color(0, 0, 255)
        pdf.cell(0, 6, "SAP portal link: https://sdcsppscs.svkm.ac.in:44300/irj/portal", ln=1)
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Times", '', 11)
        pdf.set_x(margin_l + 8)
        pdf.cell(0, 6, "User ID: Student Registration Number", ln=1)
        pdf.set_x(margin_l + 8)
        pdf.cell(0, 6, 'Password: Init@123 ("I" is a Capital letter) (Initial password). (In case already changed, please ignore)', ln=1)
        pdf.ln(2)

        pdf.set_x(margin_l)
        pdf.multi_cell(text_w, 6,
            "2.  Re-examination application link on the portal will be active during the below mentioned period:")
        pdf.ln(2)

        p_start_d = portal_dates.get('start_date') if portal_dates else None
        p_start_t = portal_dates.get('start_time', '') if portal_dates else ''
        p_end_d   = portal_dates.get('end_date')   if portal_dates else None
        p_end_t   = portal_dates.get('end_time',   '') if portal_dates else ''

        if p_start_d and p_end_d:
            col_w = (text_w - 10) / 2
            tbl_x = margin_l + 5
            pdf.set_x(tbl_x)
            pdf.set_font("Times", 'B', 11)
            pdf.cell(col_w, 8, "Start Date", border=1, align='C')
            pdf.cell(col_w, 8, "End Date",   border=1, align='C', ln=1)

            s_day = p_start_d.day; e_day = p_end_d.day
            start_str = f"{s_day}{_ordinal_suffix(s_day)} {p_start_d.strftime('%B, %Y')}"
            end_str   = f"{e_day}{_ordinal_suffix(e_day)} {p_end_d.strftime('%B, %Y')}"
            end_str2  = f"(Closing time {p_end_t.strip()})" if p_end_t and str(p_end_t).strip() else ""
            row_h = 8 if not end_str2 else 14

            pdf.set_x(tbl_x)
            cy = pdf.get_y()
            pdf.rect(tbl_x, cy, col_w, row_h)
            pdf.set_text_color(255, 0, 0)
            pdf.set_xy(tbl_x, cy + (row_h - 6) / 2)
            pdf.cell(col_w, 6, start_str, border=0, align='C')

            pdf.set_xy(tbl_x + col_w, cy)
            pdf.rect(tbl_x + col_w, cy, col_w, row_h)
            if end_str2:
                pdf.set_xy(tbl_x + col_w, cy + 1);   pdf.cell(col_w, 6, end_str,  border=0, align='C')
                pdf.set_xy(tbl_x + col_w, cy + 7);   pdf.cell(col_w, 6, end_str2, border=0, align='C')
            else:
                pdf.set_xy(tbl_x + col_w, cy + (row_h - 6) / 2)
                pdf.cell(col_w, 6, end_str, border=0, align='C')

            pdf.set_text_color(0, 0, 0)
            pdf.set_xy(margin_l, cy + row_h + 3)

        pdf.ln(2)
        pdf.set_font("Times", '', 11)

        instrs = [
            ("3.  Only those students, who have applied for the re-examination by paying the prescribed fees within the given time limit, will be allowed to appear at the respective re-examination/s. The acknowledgment receipt generated after applying online should be produced during the examination on all the days.", True),
        ]
        if p_end_d and p_end_t:
            e_day = p_end_d.day
            instrs.append((f"4.  Applications will not be accepted after {p_end_t.strip()} on {e_day}{_ordinal_suffix(e_day)} {p_end_d.strftime('%B %Y')}.", True))
        else:
            instrs.append(("4.  Applications will not be accepted after the closing date and time.", True))

        instrs += [
            ("5.  Candidates are required to be present at the examination centre THIRTY MINUTES before the stipulated time.", False),
            ("6.  Candidates must produce their University Identity Card at the time of the examination.", False),
            ("7.  Candidates are not permitted to enter the examination hall after stipulated time.", False),
            ("8.  Candidates will not be permitted to leave the examination hall during the examination time.", False),
            ("9.  Candidates are forbidden from taking any unauthorized material inside the examination hall. Carrying the same will be treated as usage of unfair means.", False),
        ]
        for text, bold in instrs:
            pdf.set_x(margin_l)
            pdf.set_font("Times", 'B' if bold else '', 11)
            pdf.multi_cell(text_w, 6, text)
            pdf.ln(3)
    except Exception:
        pass

    # ── Timetable pages ───────────────────────────────────────────────────────
    try:
        df_dict = pd.read_excel(excel_path, sheet_name=None)
    except Exception as e:
        st.error(f"Error reading Excel file: {e}"); return

    sheets_processed = 0

    for sheet_name, sheet_df in df_dict.items():
        try:
            if sheet_df.empty: continue

            main_branch_full = ""
            if "_prog_" in sheet_df.columns and not sheet_df["_prog_"].dropna().empty:
                main_branch_full = str(sheet_df["_prog_"].dropna().iloc[0])
            elif "Program" in sheet_df.columns and not sheet_df["Program"].dropna().empty:
                main_branch_full = str(sheet_df["Program"].dropna().iloc[0])

            if "Unnamed" in main_branch_full or main_branch_full == "":
                if '_|_' in sheet_name:
                    main_branch_full = sheet_name.split('_|_')[0]

            rename_cols = {}
            for col in sheet_df.columns:
                if str(col).startswith("Unnamed:"): rename_cols[col] = main_branch_full
            if rename_cols: sheet_df = sheet_df.rename(columns=rename_cols)

            sheet_college_name = st.session_state.get('selected_college', "SVKM's NMIMS University")
            if IS_LAW_SCHOOL and main_branch_full:
                prog_upper = main_branch_full.upper()
                if "LL.M" in prog_upper or "MASTER OF LAW" in prog_upper or "LLM" in prog_upper:
                    sheet_college_name = "Kirit P. Mehta School of Law"
                else:
                    sheet_college_name = "Kirit P. Mehta School of Law / School of Law"

            semester_raw = "General"
            if '_|_' in sheet_name:
                parts = sheet_name.split('_|_')
                if not main_branch_full: main_branch_full = parts[0]
                semester_raw = parts[1]
            else:
                if sheet_name in ["No_Data","Daily_Statistics","Summary","Verification","Empty"]:
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
                if display_sem.lower().startswith("semester"): display_sem = display_sem[8:].strip()
                elif display_sem.lower().startswith("sem"):    display_sem = display_sem[3:].strip()

            header_content    = {'main_branch_full': main_branch_full, 'semester_roman': display_sem}
            header_exam_time  = get_header_time_for_semester(f"Sem {display_sem}")

            if not is_elective:
                if 'Exam Date' not in sheet_df.columns: continue
                sheet_df = sheet_df.dropna(how='all').reset_index(drop=True)

                fixed_cols = ["Exam Date"]
                _meta_pattern = re.compile(
                    r'^(Program|Semester|MainBranch|Note|Message|_prog_|_sem_)(\.\d+)?$', re.IGNORECASE
                )
                sub_branch_cols = [
                    c for c in sheet_df.columns
                    if c not in fixed_cols
                    and not _meta_pattern.match(str(c))
                    and pd.notna(c) and str(c).strip() != ''
                ]
                if not sub_branch_cols: continue

                # KEY: cols_per_page=6 — same as re_exam_to_pdf.py
                # Up to 6 SubBranch (stream) columns share one landscape Legal page.
                cols_per_page = 6

                for start in range(0, len(sub_branch_cols), cols_per_page):
                    chunk        = sub_branch_cols[start:start + cols_per_page]
                    cols_to_print = fixed_cols + chunk
                    chunk_df     = sheet_df[cols_to_print].copy()

                    subset      = chunk_df[chunk].astype(str).apply(lambda x: x.str.strip())
                    valid_cells = (subset != "") & (subset != "nan") & (subset != "---")
                    mask        = valid_cells.any(axis=1)
                    chunk_df    = chunk_df[mask].reset_index(drop=True)
                    if chunk_df.empty: continue

                    try:
                        chunk_df["Exam Date"] = pd.to_datetime(
                            chunk_df["Exam Date"], format="%d-%m-%Y", errors='coerce'
                        ).dt.strftime("%A, %d %B, %Y")
                    except: pass

                    page_width      = pdf.w - 2 * pdf.l_margin
                    date_col_width  = 30
                    remaining_width = page_width - date_col_width
                    sub_width       = remaining_width / max(len(chunk), 1)
                    col_widths      = [date_col_width] + [sub_width] * len(chunk)

                    pdf.add_page()
                    orig = st.session_state.get('selected_college')
                    st.session_state['selected_college'] = sheet_college_name

                    page_time_slot = header_exam_time
                    if IS_LAW_SCHOOL:
                        _time_pat = re.compile(
                            r'\[\s*(\d{1,2}:\d{2}\s*[AP]M\s*-\s*\d{1,2}:\d{2}\s*[AP]M)\s*\]', re.IGNORECASE
                        )
                        _time_counts = {}
                        for _col in chunk:
                            if _col not in chunk_df.columns: continue
                            for _cell in chunk_df[_col].astype(str):
                                for _t in _time_pat.findall(_cell):
                                    _t_norm = re.sub(r'\s+', ' ', _t.strip().upper())
                                    _time_counts[_t_norm] = _time_counts.get(_t_norm, 0) + 1
                        if _time_counts:
                            page_time_slot = max(_time_counts, key=_time_counts.get)

                    print_table_custom(
                        pdf, chunk_df, cols_to_print, col_widths, line_height=5,
                        header_content=header_content, Programs=chunk,
                        time_slot=page_time_slot, actual_time_slots=None,
                        declaration_date=declaration_date
                    )
                    if orig: st.session_state['selected_college'] = orig
                    sheets_processed += 1

            else:
                target_cols    = ['Exam Date', 'OE Type', 'Open Elective (All Applicable Streams)']
                available_cols = [c for c in target_cols if c in sheet_df.columns]
                if len(available_cols) < 3: continue

                sheet_df = sheet_df.dropna(subset=['Exam Date']).reset_index(drop=True)
                if sheet_df.empty: continue

                try:
                    sheet_df["Exam Date"] = pd.to_datetime(
                        sheet_df["Exam Date"], format="%d-%m-%Y", errors='coerce'
                    ).dt.strftime("%A, %d %B, %Y")
                except: pass

                pdf.add_page()
                col_widths      = [30, 25]
                remaining_width = pdf.w - 2 * pdf.l_margin - sum(col_widths)
                col_widths.append(remaining_width)

                orig = st.session_state.get('selected_college')
                st.session_state['selected_college'] = sheet_college_name
                print_table_custom(
                    pdf, sheet_df, available_cols, col_widths, line_height=5,
                    header_content=header_content, Programs=["Electives"],
                    time_slot=header_exam_time, actual_time_slots=None,
                    declaration_date=declaration_date
                )
                if orig: st.session_state['selected_college'] = orig
                sheets_processed += 1

        except Exception as e:
            st.warning(f"Error processing PDF sheet {sheet_name}: {e}")
            continue

    if sheets_processed == 0:
        st.error("No valid sheets generated in PDF."); return

    try:
        pdf.output(pdf_path)
    except Exception as e:
        st.error(f"Save PDF failed: {e}")


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 6 — ORCHESTRATOR
# ═══════════════════════════════════════════════════════════════════════════════

def generate_outputs(semester_wise_timetable, output_pdf,
                     declaration_date=None, portal_dates=None):
    temp_excel = "temp_auto_reexam.xlsx"
    excel_data = save_to_excel(semester_wise_timetable)

    if not excel_data:
        st.error("❌ No Excel data generated."); return None

    try:
        with open(temp_excel, "wb") as f:
            f.write(excel_data.getvalue())
    except Exception as e:
        st.error(f"❌ Error saving temp Excel: {e}"); return None

    try:
        _all_sems = sorted(semester_wise_timetable.keys())
        convert_excel_to_pdf(temp_excel, output_pdf,
                             declaration_date=declaration_date,
                             portal_dates=portal_dates,
                             all_semesters=_all_sems)
    except Exception as e:
        st.error(f"❌ PDF error: {e}\n{traceback.format_exc()}"); return excel_data
    finally:
        try: os.remove(temp_excel)
        except: pass

    # Post-process: remove blank pages
    try:
        if os.path.exists(output_pdf):
            reader  = PdfReader(output_pdf)
            writer  = PdfWriter()
            pat     = re.compile(r'^[\s\n]*(?:Page\s*)?\d+[\s\n]*$')
            for page in reader.pages:
                try:    text = page.extract_text() or ""
                except: text = ""
                cleaned = text.strip()
                if cleaned and not pat.match(cleaned) and len(cleaned) > 10:
                    writer.add_page(page)
            if writer.pages:
                with open(output_pdf, 'wb') as f: writer.write(f)
    except: pass

    return excel_data


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 7 — MAIN STREAMLIT APP
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    st.markdown(
        '<div class="main-header">'
        '<h1>📅 Re-Exam Auto-Scheduler</h1>'
        '<p>Upload the re-exam subject list → get a fully-scheduled PDF timetable</p>'
        '</div>',
        unsafe_allow_html=True
    )

    for k in ('parsed_df', 'scheduled_df', 'timetable', 'pdf_data', 'excel_data'):
        if k not in st.session_state: st.session_state[k] = None

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.header("⚙️ Configuration")

        college_options  = [c["name"] for c in COLLEGES]
        selected_display = st.selectbox("Select College", college_options)
        st.session_state.selected_college = selected_display

        st.markdown("---")
        st.subheader("📅 Exam Schedule Window")

        start_date_input = st.date_input("Exam Start Date", value=_date.today())
        num_days_input   = st.number_input(
            "Number of Exam Days (10 or 11)", min_value=5, max_value=15, value=10, step=1
        )

        st.markdown("---")
        st.subheader("🗓️ Bank Holidays")
        st.caption("Add holiday dates to exclude from scheduling")

        if 'holiday_list' not in st.session_state:
            st.session_state.holiday_list = []

        new_holiday = st.date_input("Add Holiday", value=None, key="holiday_add")
        if st.button("➕ Add Holiday") and new_holiday:
            if new_holiday not in st.session_state.holiday_list:
                st.session_state.holiday_list.append(new_holiday)

        if st.session_state.holiday_list:
            st.write("**Excluded dates:**")
            to_remove = []
            for h in st.session_state.holiday_list:
                c1, c2 = st.columns([3, 1])
                c1.write(h.strftime("%d %b %Y"))
                if c2.button("✖", key=f"rm_{h}"):
                    to_remove.append(h)
            for h in to_remove:
                st.session_state.holiday_list.remove(h)

        st.markdown("---")
        decl_date = st.date_input("📆 Declaration Date (Optional)", value=None)
        st.markdown("---")
        st.subheader("🌐 Portal Application Window")
        portal_start_date = st.date_input("Portal Start Date", value=_date.today(), key="psd")
        portal_start_time = st.text_input("Portal Start Time", value="10:00 am", key="pst")
        portal_end_date   = st.date_input("Portal End Date",   value=_date.today(), key="ped")
        portal_end_time   = st.text_input("Portal End Time",   value="4:00 pm",  key="pet")

        st.markdown("---")
        st.subheader("⏰ Time Slots")
        s1_start = st.text_input("Slot 1 Start", value="10:00 AM")
        s1_end   = st.text_input("Slot 1 End",   value="1:00 PM")
        s2_start = st.text_input("Slot 2 Start", value="2:00 PM")
        s2_end   = st.text_input("Slot 2 End",   value="5:00 PM")
        st.session_state['time_slots'] = {
            1: {"start": s1_start, "end": s1_end},
            2: {"start": s2_start, "end": s2_end},
        }

        st.markdown("---")
        st.info(
            "**Input Excel columns expected:**\n\n"
            "- Program / Programme\n"
            "- Stream\n"
            "- Academic Year\n"
            "- Current Session\n"
            "- Module Abbreviation\n"
            "- Module Description\n"
            "- Subject Type (OE1/OE2 = Open Elective)\n"
            "- CM group Number"
        )

    # ── Upload ────────────────────────────────────────────────────────────────
    col1, col2 = st.columns([2, 1])

    with col1:
        st.markdown(
            '<div class="upload-section">'
            '<h3 style="margin:0 0 1rem 0;color:#951C1C;">📁 Upload Re-Exam Subject List (Excel)</h3>'
            '<p style="margin:0;color:#666;">Same template used by the existing Re-Exam Data-to-PDF app</p>'
            '</div>',
            unsafe_allow_html=True
        )
        uploaded = st.file_uploader("Upload Excel", type=['xlsx','xls'], label_visibility="collapsed")

        if uploaded:
            with st.spinner("Parsing input file…"):
                df_parsed, err = parse_input_file(uploaded)
                if err:
                    st.error(f"❌ {err}")
                else:
                    st.session_state.parsed_df = df_parsed
                    st.success(f"✅ Loaded {len(df_parsed)} subjects across "
                               f"{df_parsed['Semester'].nunique()} semester(s), "
                               f"{df_parsed['MainBranch'].nunique()} programme(s)")

    with col2:
        st.info(
            "ℹ️ **Scheduling Rules**\n\n"
            "• Same subject name → same date & time across all programmes/sems\n\n"
            "• 10–11 working days (Mon–Sat), no Sundays, no holidays\n\n"
            "• OE subjects placed on last 2 reserved days\n\n"
            "• Up to 6 streams per page — identical to re_exam_to_pdf.py layout"
        )

    # ── Data preview + scheduling ─────────────────────────────────────────────
    if st.session_state.parsed_df is not None:
        df = st.session_state.parsed_df

        st.markdown("---")
        st.subheader("📊 Input Data Summary")
        ca, cb, cc, cd = st.columns(4)
        ca.metric("Total Subjects", len(df))
        cb.metric("Unique Subject Names", df['Subject'].nunique())
        cc.metric("Programmes", df['MainBranch'].nunique())
        cd.metric("Semesters", df['Semester'].nunique())

        with st.expander("🔍 Raw Data Preview (first 50 rows)"):
            st.dataframe(
                df[['MainBranch','SubBranch','Semester','Subject','OE','CMGroup']].head(50)
            )

        st.markdown("---")

        if st.button("🗓️ Auto-Schedule Re-Exams", type="primary", use_container_width=True):
            with st.spinner("Scheduling…"):
                holidays_set = set(st.session_state.get('holiday_list', []))
                start_dt     = datetime.combine(start_date_input, datetime.min.time())

                scheduled = schedule_reexams(
                    df=df,
                    start_date=start_dt,
                    num_days=int(num_days_input),
                    holidays=holidays_set,
                    time_slots_dict=st.session_state['time_slots'],
                )
                st.session_state.scheduled_df = scheduled
                st.session_state.timetable    = build_semester_timetable(scheduled)

                # Verify same-name → same-date constraint
                grouped = scheduled.groupby(
                    scheduled['Subject'].str.strip().str.upper()
                )['Exam Date'].nunique()
                conflicts = grouped[grouped > 1].index.tolist()
                if conflicts:
                    st.warning(f"⚠️ {len(conflicts)} subject(s) ended up on multiple dates "
                               f"(should not happen — review data): {conflicts[:5]}")
                else:
                    st.success("✅ Scheduling complete — same subject name → same date everywhere ✔")

    # ── Scheduled preview ─────────────────────────────────────────────────────
    if st.session_state.scheduled_df is not None:
        sdf = st.session_state.scheduled_df

        st.markdown("---")
        st.subheader("📋 Scheduled Timetable Preview")

        summary = (
            sdf.drop_duplicates(subset=['Subject'])
               [['Subject','Exam Date','Exam Time','OE']]
               .sort_values('Exam Date')
               .reset_index(drop=True)
        )
        st.dataframe(summary, use_container_width=True)

        st.markdown("---")

        if st.button("🚀 Generate PDF & Excel Timetable", type="primary", use_container_width=True):
            with st.spinner("Rendering PDF…"):
                pdf_filename = "AutoScheduled_ReExam_Timetable.pdf"
                _portal_dates = {
                    'start_date': portal_start_date,
                    'start_time': portal_start_time,
                    'end_date':   portal_end_date,
                    'end_time':   portal_end_time,
                }
                excel_data = generate_outputs(
                    st.session_state.timetable,
                    pdf_filename,
                    declaration_date=decl_date,
                    portal_dates=_portal_dates,
                )
                st.session_state.excel_data = excel_data

            if os.path.exists(pdf_filename):
                st.success("✅ PDF generated successfully!")
                st.balloons()
                c1, c2 = st.columns(2)
                with c1:
                    with open(pdf_filename, "rb") as f:
                        st.download_button(
                            "📥 Download PDF Timetable",
                            f,
                            f"ReExam_AutoScheduled_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                            "application/pdf",
                            use_container_width=True
                        )
                with c2:
                    if st.session_state.excel_data:
                        st.download_button(
                            "📥 Download Excel Timetable",
                            st.session_state.excel_data.getvalue(),
                            f"ReExam_AutoScheduled_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                try: os.remove(pdf_filename)
                except: pass
            else:
                st.error("❌ PDF file was not created. Check errors above.")
                if st.session_state.excel_data:
                    st.download_button(
                        "📥 Download Excel Only",
                        st.session_state.excel_data.getvalue(),
                        "ReExam_AutoScheduled.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )


if __name__ == "__main__":
    main()
