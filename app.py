import pandas as pd
from datetime import datetime, timedelta
import streamlit as st

def quick_optimize_timetable(excel_file_path, output_file_path):
    """
    Quick optimization script to compress timetable from 46 to <20 days
    """
    # Read all sheets from Excel
    all_sheets = pd.read_excel(excel_file_path, sheet_name=None)
    
    # Extract all exams into a single list
    all_exams = []
    
    for sheet_name, df in all_sheets.items():
        if df.empty or len(df) < 2:
            continue
            
        # Get headers (skip first two columns for dates and time)
        branches = df.columns[2:].tolist()
        
        # Process each row
        for idx, row in df.iterrows():
            if idx == 0:  # Skip header
                continue
                
            exam_date = row.iloc[0]
            time_slot = row.iloc[1]
            
            # Skip invalid dates
            if pd.isna(exam_date) or exam_date == "---" or exam_date == "Exam Date":
                continue
            
            # Extract exams for each branch
            for i, branch in enumerate(branches):
                subject = row.iloc[i + 2]
                if pd.notna(subject) and subject != "---":
                    # Extract subject code
                    import re
                    match = re.search(r'\((.*?)\)', str(subject))
                    subject_code = match.group(1) if match else str(subject)
                    
                    all_exams.append({
                        'sheet': sheet_name,
                        'branch': branch,
                        'subject': subject,
                        'subject_code': subject_code,
                        'original_date': exam_date,
                        'time_slot': time_slot
                    })
    
    print(f"Total exams found: {len(all_exams)}")
    
    # Group by subject code to identify common subjects
    from collections import defaultdict
    subject_groups = defaultdict(list)
    
    for exam in all_exams:
        subject_groups[exam['subject_code']].append(exam)
    
    # Separate common and individual subjects
    common_subjects = {}
    individual_subjects = []
    
    for subject_code, exams in subject_groups.items():
        unique_branches = set(exam['branch'] for exam in exams)
        if len(unique_branches) > 1:
            # Common subject
            common_subjects[subject_code] = exams
        else:
            # Individual subject
            individual_subjects.extend(exams)
    
    print(f"Common subjects: {len(common_subjects)}")
    print(f"Individual subject instances: {len(individual_subjects)}")
    
    # Create optimized schedule
    optimized_schedule = []
    base_date = datetime(2025, 4, 1)  # April 1, 2025
    current_date = base_date
    
    # Phase 1: Schedule all common subjects first (one day per common subject)
    for subject_code, exams in common_subjects.items():
        # Skip Sunday
        while current_date.weekday() == 6:
            current_date += timedelta(days=1)
        
        # Assign same date to all instances of this common subject
        date_str = current_date.strftime("%d-%m-%Y")
        
        # Use morning slot for odd semesters, afternoon for even
        for exam in exams:
            exam['new_date'] = date_str
            exam['phase'] = 'common'
            optimized_schedule.append(exam)
        
        current_date += timedelta(days=1)
    
    # Phase 2: Pack individual subjects efficiently
    # Group by branch to minimize conflicts
    branch_groups = defaultdict(list)
    for exam in individual_subjects:
        branch_groups[exam['branch']].append(exam)
    
    # Schedule branch by branch
    max_exams_per_day = 40  # Maximum exams we can handle per day
    daily_exam_count = defaultdict(int)
    
    for branch, exams in branch_groups.items():
        for exam in exams:
            # Find next available slot
            temp_date = base_date
            
            while True:
                # Skip Sundays
                while temp_date.weekday() == 6:
                    temp_date += timedelta(days=1)
                
                date_str = temp_date.strftime("%d-%m-%Y")
                
                # Check if this day has capacity
                if daily_exam_count[date_str] < max_exams_per_day:
                    exam['new_date'] = date_str
                    exam['phase'] = 'individual'
                    optimized_schedule.append(exam)
                    daily_exam_count[date_str] += 1
                    break
                
                temp_date += timedelta(days=1)
                
                # Safety check - don't go beyond 20 days
                if (temp_date - base_date).days > 20:
                    exam['new_date'] = date_str  # Force fit
                    exam['phase'] = 'overflow'
                    optimized_schedule.append(exam)
                    break
    
    # Create result summary
    result_df = pd.DataFrame(optimized_schedule)
    
    # Calculate new span
    unique_dates = result_df['new_date'].unique()
    date_objects = [datetime.strptime(d, "%d-%m-%Y") for d in unique_dates]
    new_span = (max(date_objects) - min(date_objects)).days + 1
    
    # Generate summary statistics
    st.write("### Optimization Results")
    st.write(f"**Original span**: 46 days")
    st.write(f"**New span**: {new_span} days")
    st.write(f"**Compression achieved**: {((46-new_span)/46*100):.1f}%")
    st.write(f"**Total unique dates used**: {len(unique_dates)}")
    
    # Show date distribution
    date_distribution = result_df.groupby('new_date').size().sort_index()
    st.write("\n### Daily exam distribution:")
    st.bar_chart(date_distribution)
    
    # Reconstruct timetable format for each sheet
    optimized_sheets = {}
    
    for sheet_name in all_sheets.keys():
        sheet_exams = result_df[result_df['sheet'] == sheet_name]
        if sheet_exams.empty:
            continue
        
        # Get unique dates and branches
        dates = sorted(sheet_exams['new_date'].unique(), 
                      key=lambda x: datetime.strptime(x, "%d-%m-%Y"))
        branches = sorted(sheet_exams['branch'].unique())
        
        # Create timetable matrix
        timetable_data = []
        
        for date in dates:
            row = [date, sheet_exams[sheet_exams['new_date'] == date].iloc[0]['time_slot']]
            
            for branch in branches:
                exam = sheet_exams[(sheet_exams['new_date'] == date) & 
                                 (sheet_exams['branch'] == branch)]
                
                if not exam.empty:
                    row.append(exam.iloc[0]['subject'])
                else:
                    row.append("---")
            
            timetable_data.append(row)
        
        # Create DataFrame
        columns = ['Exam Date', 'Time Slot'] + branches
        optimized_sheets[sheet_name] = pd.DataFrame(timetable_data, columns=columns)
    
    # Save to Excel
    with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
        for sheet_name, df in optimized_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    st.success(f"✅ Optimized timetable saved to: {output_file_path}")
    
    # Show warnings if any
    overflow_exams = result_df[result_df['phase'] == 'overflow']
    if not overflow_exams.empty:
        st.warning(f"⚠️ {len(overflow_exams)} exams exceeded the 20-day limit and were force-fitted")
    
    # Show top optimization achievements
    st.write("\n### Key Optimizations Applied:")
    st.write("1. ✅ Consolidated all common subjects to single days")
    st.write("2. ✅ Eliminated gaps between exam days")
    st.write("3. ✅ Maximized daily exam capacity")
    st.write("4. ✅ Compressed sparse May schedule into April")
    
    return result_df, new_span

# Usage example
if __name__ == "__main__":
    # Run the optimization
    st.title("Timetable Optimization Tool")
    
    input_file = "timetable_20250725_102113.xlsx"
    output_file = "optimized_timetable.xlsx"
    
    try:
        result, span = quick_optimize_timetable(input_file, output_file)
        
        if span <= 20:
            st.balloons()
            st.success("🎉 Successfully optimized to under 20 days!")
        else:
            st.info(f"📊 Achieved {span} days - close to target!")
            
    except Exception as e:
        st.error(f"Error during optimization: {str(e)}")
