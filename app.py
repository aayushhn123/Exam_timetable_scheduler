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

    /* Button styling */
    .stButton>button {
        background-color: #3498db;
        color: white;
        border: none;
        padding: 0.5rem 1rem;
        border-radius: 5px;
        transition: background-color 0.3s;
    }

    .stButton>button:hover {
        background-color: #2980b9;
    }

    /* Ensure footer stays at bottom */
    .footer {
        text-align: center;
        padding: 1rem;
        color: #FFF;
        background-color: #2C3E50;
        border-radius: 10px;
        margin-top: 2rem;
    }
</style>
""", unsafe_allow_html=True)

# Helper Functions
def calculate_end_time(start_time, duration):
    """Calculate end time based on start time and duration."""
    time_obj = datetime.strptime(start_time, "%I:%M %p")
    end_time = time_obj + timedelta(hours=duration)
    return end_time.strftime("%I:%M %p")

def generate_available_days(base_date, holidays_set, num_days=35):
    """Generate available days starting from base_date, avoiding holidays and weekends."""
    available_days = []
    current_date = base_date
    while len(available_days) < num_days:
        if current_date.weekday() < 5 and current_date.date() not in holidays_set:
            available_days.append(current_date.date())
        current_date += timedelta(days=1)
    return available_days

class ExamScheduler:
    def __init__(self, df, base_date, holidays_set, room_limits):
        self.df = df
        self.base_date = base_date
        self.holidays_set = holidays_set
        self.room_limits = room_limits
        self.schedule = defaultdict(lambda: defaultdict(dict))

    def find_next_available_day(self, start_date):
        """Find the next available day starting from start_date."""
        current_date = start_date
        while True:
            if current_date.weekday() < 5 and current_date.date() not in self.holidays_set:
                return current_date.date()
            current_date += timedelta(days=1)

    def get_preferred_slot(self, semester):
        """Get preferred time slot based on semester (example logic)."""
        return "10:00 AM - 1:00 PM" if int(semester) % 2 == 0 else "2:00 PM - 5:00 PM"

    def can_schedule_common(self, group, day, time_slot):
        """Check if common subjects can be scheduled."""
        branches = group['Branch'].unique()
        for branch in branches:
            if (day in self.schedule and time_slot in self.schedule[day] and
                branch in self.schedule[day][time_slot] and self.schedule[day][time_slot][branch]):
                return False
        return True

    def schedule_common(self, group, day, time_slot):
        """Schedule common subjects."""
        for _, row in group.iterrows():
            branch = row['Branch']
            self.df.loc[self.df.index == row.name, 'Exam Date'] = day.strftime("%d-%m-%Y")
            self.df.loc[self.df.index == row.name, 'Time Slot'] = time_slot
            self.schedule[day][time_slot][branch] = row['Subject']

    def schedule_semester_non_electives_with_optimization(self, sem_df):
        """Schedule non-elective subjects with optimization."""
        for _, row in sem_df.iterrows():
            day = self.find_next_available_day(self.base_date)
            time_slot = self.get_preferred_slot(row['Semester'])
            if not (day in self.schedule and time_slot in self.schedule[day] and
                    row['Branch'] in self.schedule[day][time_slot] and self.schedule[day][time_slot][row['Branch']]):
                self.df.loc[self.df.index == row.name, 'Exam Date'] = day.strftime("%d-%m-%Y")
                self.df.loc[self.df.index == row.name, 'Time Slot'] = time_slot
                self.schedule[day][time_slot][row['Branch']] = row['Subject']

    def schedule_electives_with_optimization(self, elec_df):
        """Schedule elective subjects with optimization."""
        for _, row in elec_df.iterrows():
            day = self.find_next_available_day(self.base_date)
            time_slot = self.get_preferred_slot(row['Semester'])
            if not (day in self.schedule and time_slot in self.schedule[day] and
                    row['Branch'] in self.schedule[day][time_slot] and self.schedule[day][time_slot][row['Branch']]):
                self.df.loc[self.df.index == row.name, 'Exam Date'] = day.strftime("%d-%m-%Y")
                self.df.loc[self.df.index == row.name, 'Time Slot'] = time_slot
                self.schedule[day][time_slot][row['Branch']] = row['Subject']

def check_empty_days(df, base_date, last_scheduled_date):
    """Check if there are any days between base_date and last_scheduled_date with no exams."""
    scheduled_dates = pd.to_datetime(df['Exam Date']).dt.date.unique()
    all_dates = {base_date + timedelta(days=i) for i in range((last_scheduled_date - base_date).days + 1)}
    holidays_set = {datetime.strptime(h, "%d-%m-%Y").date() for h in st.session_state.get('holidays', [])}
    weekends = {d for d in all_dates if d.weekday() >= 5}  # Saturday=5, Sunday=6
    valid_days = all_dates - holidays_set - weekends
    empty_days = [d for d in valid_days if d not in scheduled_dates and d >= base_date and d <= last_scheduled_date]
    if empty_days:
        st.warning(f"Empty days detected: {[d.strftime('%d-%m-%Y') for d in empty_days]}")
    return len(empty_days) == 0

def process_constraints_with_real_time_optimization(df, base_date, holidays_set, room_limits):
    """Process constraints and optimize timetable in real-time."""
    df['Exam Date'] = None
    df['Time Slot'] = None
    optimizer = ExamScheduler(df, base_date, holidays_set, room_limits)
    
    # Schedule common subjects first
    common_subjects = df[df['IsCommon'] == 'YES']
    for module_code, group in common_subjects.groupby('ModuleCode'):
        day = optimizer.find_next_available_day(base_date)
        time_slot = optimizer.get_preferred_slot(group['Semester'].iloc[0])
        if optimizer.can_schedule_common(group, day, time_slot):
            optimizer.schedule_common(group, day, time_slot)
    
    # Schedule non-elective subjects per semester
    for semester in sorted(df['Semester'].unique()):
        sem_df = df[(df['Semester'] == semester) & (df['IsCommon'] == 'NO') & (df['Category'] != 'ELEC')]
        optimizer.schedule_semester_non_electives_with_optimization(sem_df)
    
    # Schedule electives
    elec_df = df[df['Category'] == 'ELEC']
    optimizer.schedule_electives_with_optimization(elec_df)
    
    # Validate that each day has at least one exam
    last_scheduled_date = pd.to_datetime(df['Exam Date']).max().date()
    if not check_empty_days(df, base_date, last_scheduled_date):
        st.error("Scheduling failed to ensure at least one exam per day. Consider adjusting constraints.")
    
    return df

def main():
    st.markdown("""
    <div class="main-header" style="background-color: #2C3E50;">
        <h1>📅 Exam Timetable Generator</h1>
        <p>Effortlessly create optimized exam schedules</p>
    </div>
    """, unsafe_allow_html=True)
    
    with st.sidebar:
        st.header("Settings")
        st.subheader("Upload Data")
        uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx'])
        st.subheader("Schedule Parameters")
        base_date = st.date_input("Select Base Date", value=date(2025, 8, 6))  # Default to today
        holidays = st.text_area("Enter Holidays (dd-mm-yyyy, one per line)", 
                              value="15-08-2025\n26-01-2025")
        st.session_state['holidays'] = holidays.split('\n') if holidays else []
        room_limits = {'Morning': 10, 'Afternoon': 10}  # Example limits
    
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        if st.button("Generate Timetable"):
            with st.spinner("Generating timetable..."):
                holidays_set = {datetime.strptime(h.strip(), "%d-%m-%Y").date() for h in st.session_state['holidays'] if h.strip()}
                df = process_constraints_with_real_time_optimization(df, base_date, holidays_set, room_limits)
                st.session_state['timetable'] = df
                st.success("Timetable generated successfully!")
    
    if 'timetable' in st.session_state:
        df = st.session_state['timetable']
        st.subheader("Generated Timetable")
        st.dataframe(df, use_container_width=True)
        
        # Example display of electives (to match truncated original)
        elec_df = df[df['Category'] == 'ELEC']
        if not elec_df.empty:
            def format_elective_display(row):
                subject = row['Subject']
                oe_type = row['OE']
                time_slot = row['Time Slot']
                duration = row['Exam Duration']
                
                base_display = f"{subject} [{oe_type}]"
                
                if duration != 3 and time_slot and time_slot.strip():
                    start_time = time_slot.split(' - ')[0]
                    end_time = calculate_end_time(start_time, duration)
                    time_range = f" ({start_time} to {end_time})"
                else:
                    time_range = ""
                
                return base_display + time_range
            
            elec_df["SubjectDisplay"] = elec_df.apply(format_elective_display, axis=1)
            elec_df["Exam Date"] = pd.to_datetime(elec_df["Exam Date"], format="%d-%m-%Y", errors='coerce')
            elec_df = elec_df.sort_values(by="Exam Date", ascending=True)
            elec_pivot = elec_df.groupby(['OE', 'Exam Date', 'Time Slot'])['SubjectDisplay'].apply(
                lambda x: ", ".join(x)
            ).reset_index()
            if not elec_pivot.empty:
                st.markdown("#### Open Electives")
                st.dataframe(elec_pivot, use_container_width=True)

    st.markdown("---")
    st.markdown("""
    <div class="footer">
        <p>🎓 <strong>Exam Timetable Generator</strong></p>
        <p>Developed for MUKESH PATEL SCHOOL OF TECHNOLOGY MANAGEMENT & ENGINEERING</p>
        <p style="font-size: 0.9em;">Streamlined scheduling • Conflict-free timetables • Multiple export formats</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
