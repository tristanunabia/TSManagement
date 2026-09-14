import streamlit as st
import pandas as pd
import datetime
import calendar
from date_engine import generate_cutoff_dates, get_previous_month
from database import (
    init_db, get_instructors, add_instructor, update_instructor, delete_instructor,
    get_class_loads, add_class_load, delete_class_load,
    get_timesheet_overrides, save_timesheet_overrides, clear_timesheet_overrides
)
from excel_generator import generate_excel_timesheet
from pdf_generator import generate_pdf_timesheet

# Set page config
st.set_page_config(
    page_title="Academic Timesheet & Schedule System",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Load database schema
init_db()

# Premium Fonts and CSS Styling
st.markdown("""
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700&family=Plus+Jakarta+Sans:wght@300;400;500;600;700&display=swap" rel="stylesheet">
<style>
    /* Global Styles */
    * {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }
    h1, h2, h3, .metric-val {
        font-family: 'Outfit', sans-serif !important;
        font-weight: 700 !important;
    }
    
    /* Premium UI Card */
    .metric-card {
        background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
        margin-bottom: 20px;
        text-align: center;
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.05), 0 4px 6px -2px rgba(0, 0, 0, 0.05);
    }
    .metric-title {
        color: #64748b;
        font-size: 0.85rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        margin-bottom: 8px;
    }
    .metric-val {
        font-size: 2.2rem;
        color: #1e293b;
        margin-bottom: 4px;
    }
    .metric-status-pass {
        color: #10b981;
        font-weight: 700;
        font-size: 1.1rem;
    }
    .metric-status-warn {
        color: #ef4444;
        font-weight: 700;
        font-size: 1.1rem;
    }
    
    /* Timetable styles */
    .timetable-container {
        max-height: 600px;
        overflow-y: auto;
        border: 1px solid #cbd5e1;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        margin-bottom: 20px;
    }
    .timetable {
        width: 100%;
        border-collapse: separate;
        border-spacing: 0;
        text-align: center;
    }
    .timetable th {
        background-color: #1e3a8a;
        color: white;
        font-weight: 600;
        font-size: 0.85rem;
        padding: 12px 8px;
        border-bottom: 2px solid #0f172a;
        position: sticky;
        top: 0;
        z-index: 10;
    }
    .timetable th:first-child {
        border-top-left-radius: 12px;
        position: sticky;
        left: 0;
        z-index: 11;
    }
    .timetable th:last-child {
        border-top-right-radius: 12px;
    }
    .timetable td {
        border-bottom: 1px solid #e2e8f0;
        border-right: 1px solid #e2e8f0;
        padding: 4px;
        font-size: 0.75rem;
        height: 36px;
        vertical-align: middle;
    }
    .timetable td:first-child {
        font-weight: 600;
        background-color: #f8fafc;
        position: sticky;
        left: 0;
        z-index: 5;
        border-right: 2px solid #cbd5e1;
    }
    .class-block-container {
        padding: 6px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.75rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.04);
        text-align: center;
        display: flex;
        flex-direction: column;
        justify-content: center;
        height: 100%;
        line-height: 1.3;
        transition: transform 0.15s ease;
    }
    .class-block-container:hover {
        transform: scale(1.02);
    }
    .class-lec-theme {
        background-color: #e2efda !important;
        color: #274e13 !important;
        border-left: 4px solid #375623;
    }
    .class-lab-theme {
        background-color: #ddebf7 !important;
        color: #1f4e78 !important;
        border-left: 4px solid #1f4e78;
    }
    .class-consult-theme {
        background-color: #fff2cc !important;
        color: #7f6000 !important;
        border-left: 4px solid #c65911;
    }
    .weekend-cell-theme {
        background-color: #f1f5f9 !important;
    }
</style>
""", unsafe_allow_html=True)

# ----------------- SIDEBAR CONFIGURATION -----------------
st.sidebar.title("Configuration Panel")

# Instructor Selection
instructors_df = get_instructors()

if instructors_df.empty:
    st.sidebar.warning("No instructors registered. Please create one.")
    # Fallback default input if no instructors exist
    with st.sidebar.expander("➕ Add First Instructor", expanded=True):
        new_name = st.text_input("Instructor Name", key="init_name")
        new_status = st.selectbox("Employment Status", ["SC-Based", "Part-Time", "Pro-Rated"], key="init_status")
        new_dept = st.text_input("Department", "Computer Science", key="init_dept")
        new_email = st.text_input("Email", key="init_email")
        new_phone = st.text_input("Phone", key="init_phone")
        if st.button("Register Instructor", key="init_btn"):
            if new_name:
                add_instructor(new_name, new_email, new_phone, new_dept, new_status)
                st.rerun()
    st.stop()

# Populate selectors
instructor_options = {row['name']: row['id'] for _, row in instructors_df.iterrows()}
selected_inst_name = st.sidebar.selectbox("Select Instructor", list(instructor_options.keys()))
selected_inst_id = instructor_options[selected_inst_name]
selected_inst_row = instructors_df[instructors_df['id'] == selected_inst_id].iloc[0]

# Month and Year selection
current_year = datetime.datetime.now().year
years = [current_year - 1, current_year, current_year + 1]
selected_year = st.sidebar.selectbox("Select Year", years, index=1)

months_map = {calendar.month_name[i]: i for i in range(1, 13)}
selected_month_name = st.sidebar.selectbox("Select Month", list(months_map.keys()), index=datetime.datetime.now().month - 1)
selected_month = months_map[selected_month_name]

selected_cutoff = st.sidebar.radio("Select Cutoff Period", ["11-25", "26-10"], index=0)

# Sidebar Profile Manager
with st.sidebar.expander("👤 Instructor Profiles Manager", expanded=False):
    st.subheader("Add New Instructor")
    add_name = st.text_input("Name", key="add_name")
    add_status = st.selectbox("Status", ["SC-Based", "Part-Time", "Pro-Rated"], key="add_status")
    add_dept = st.text_input("Department", key="add_dept")
    add_email = st.text_input("Email", key="add_email")
    add_phone = st.text_input("Phone", key="add_phone")
    if st.button("Save New Profile"):
        if add_name:
            add_instructor(add_name, add_email, add_phone, add_dept, add_status)
            st.success(f"Added {add_name}")
            st.rerun()
            
    st.markdown("---")
    st.subheader("Edit Current Instructor")
    edit_name = st.text_input("Edit Name", value=selected_inst_row["name"])
    edit_status = st.selectbox("Edit Status", ["SC-Based", "Part-Time", "Pro-Rated"], index=["SC-Based", "Part-Time", "Pro-Rated"].index(selected_inst_row["employment_status"]))
    edit_dept = st.text_input("Edit Department", value=selected_inst_row.get("department", ""))
    edit_email = st.text_input("Edit Email", value=selected_inst_row.get("email", ""))
    edit_phone = st.text_input("Edit Phone", value=selected_inst_row.get("phone", ""))
    col_edit_1, col_edit_2 = st.columns(2)
    with col_edit_1:
        if st.button("Update Profile"):
            update_instructor(selected_inst_id, edit_name, edit_email, edit_phone, edit_dept, edit_status)
            st.success("Profile Updated!")
            st.rerun()
    with col_edit_2:
        if st.button("Delete Profile", type="primary"):
            if len(instructors_df) <= 1:
                st.error("Cannot delete the only registered instructor.")
            else:
                delete_instructor(selected_inst_id)
                st.success("Deleted Profile!")
                st.rerun()

# ----------------- MAIN VIEW -----------------

st.title("Academic Instructor Timesheet & Schedule Management System")
st.markdown(f"**Current Instructor:** `{selected_inst_row['name']}` | **Status:** `{selected_inst_row['employment_status']}` | **Department:** `{selected_inst_row.get('department', 'N/A')}`")

# Create main tabs
tab_timesheet, tab_class_load, tab_schedule = st.tabs([
    "📝 Timesheet View", 
    "📁 Class Load Directory", 
    "📅 Master Timetable"
])

# Fetch data for this instructor
class_loads_df = get_class_loads(selected_inst_id)

# ----------------- TAB 1: TIMESHEET VIEW -----------------
with tab_timesheet:
    st.header(f"Timesheet for {selected_month_name} ({selected_cutoff}), {selected_year}")
    
    # Generate Cutoff Dates
    cutoff_dates = generate_cutoff_dates(selected_year, selected_month, selected_cutoff)
    date_columns = [f"{d['day_num']} ({d['day_abbr']})" for d in cutoff_dates]
    
    # Create the standard rows for the timesheet
    teaching_classes = class_loads_df[class_loads_df["load_category"].isin(["Regular Load", "Excess/Overload"])]
    activity_rows = []
    
    # Section A
    for _, cl in teaching_classes.iterrows():
        key_name = f"{cl['type']}: {cl['subject_name']} ({cl['section_code']})"
        activity_rows.append(("Teaching Activities", key_name))
        
    # Section B & C
    activity_rows.append(("Non-Teaching SC", "Career Orientation Seminars (COS)"))
    activity_rows.append(("Non-Teaching SC", "Guidance/Counseling"))
    activity_rows.append(("Non-Teaching HQ", "Consultation"))
    activity_rows.append(("Non-Teaching HQ", "Administrative Hours"))
    
    # Load overrides from SQLite
    db_overrides = get_timesheet_overrides(selected_inst_id, selected_year, selected_month, selected_cutoff)
    
    # Build initial timesheet grid data
    grid_data = {}
    for cat, key_name in activity_rows:
        row_values = []
        for d in cutoff_dates:
            date_str = d['date'].strftime("%Y-%m-%d")
            
            # 1. Check if there is a manual override saved
            val = db_overrides.get((key_name, date_str))
            
            if val is None:
                # 2. No override, calculate scheduled default
                if cat == "Teaching Activities":
                    val = 0.0
                    for _, cl in teaching_classes.iterrows():
                        cl_key = f"{cl['type']}: {cl['subject_name']} ({cl['section_code']})"
                        if cl_key == key_name and cl['day_of_week'] == d['day_abbr']:
                            val = cl['hours']
                            break
                elif key_name == "Consultation":
                    val = 0.0
                    consultations = class_loads_df[class_loads_df["load_category"] == "Consultation"]
                    for _, cl in consultations.iterrows():
                        if cl['day_of_week'] == d['day_abbr']:
                            val = cl['hours']
                            break
                else:
                    val = 0.0
            
            row_values.append(val)
        grid_data[key_name] = row_values
        
    # Create DataFrame for st.data_editor
    timesheet_df = pd.DataFrame.from_dict(grid_data, orient='index', columns=date_columns)
    
    st.markdown("⚠️ *Adjust hours inside the cell grid. Changes will compute in real-time below. Make sure to click **Save Timesheet Changes** to persist overrides.*")
    
    # Render Data Editor
    edited_df = st.data_editor(
        timesheet_df,
        use_container_width=True,
        column_config={
            col: st.column_config.NumberColumn(
                min_value=0.0,
                max_value=24.0,
                step=0.5,
                format="%.1f"
            ) for col in date_columns
        }
    )
    
    # Compute totals in real-time
    row_totals = edited_df.sum(axis=1)
    daily_totals = edited_df.sum(axis=0)
    grand_total = edited_df.sum().sum()
    
    # Save/Reset controls
    col_ctrl_1, col_ctrl_2, col_ctrl_3, col_ctrl_4 = st.columns(4)
    with col_ctrl_1:
        if st.button("💾 Save Timesheet Changes", use_container_width=True):
            # Parse edits back to SQLite format
            overrides_list = []
            for act_name in edited_df.index:
                for idx, d_info in enumerate(cutoff_dates):
                    col_name = date_columns[idx]
                    hours_val = float(edited_df.loc[act_name, col_name])
                    
                    # Store as overrides
                    overrides_list.append({
                        'row_name': act_name,
                        'date_str': d_info['date'].strftime("%Y-%m-%d"),
                        'hours': hours_val
                    })
            save_timesheet_overrides(selected_inst_id, selected_year, selected_month, selected_cutoff, overrides_list)
            st.success("Timesheet overrides successfully saved to SQLite!")
            st.rerun()
            
    with col_ctrl_2:
        if st.button("🔄 Reset to Default Schedule", use_container_width=True):
            clear_timesheet_overrides(selected_inst_id, selected_year, selected_month, selected_cutoff)
            st.success("Reset successfully. Loading standard scheduled hours...")
            st.rerun()
            
    with col_ctrl_3:
        # Generate Excel Download
        excel_bytes = generate_excel_timesheet(selected_inst_row, class_loads_df, selected_year, selected_month, selected_cutoff, db_overrides)
        filename_excel = f"Timesheet_{selected_inst_name.replace(' ', '_')}_{selected_year}_{selected_month}_{selected_cutoff}.xlsx"
        st.download_button(
            label="📊 Download Excel Report",
            data=excel_bytes,
            file_name=filename_excel,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
    with col_ctrl_4:
        # Generate PDF Download
        pdf_bytes = generate_pdf_timesheet(selected_inst_row, class_loads_df, selected_year, selected_month, selected_cutoff, db_overrides)
        filename_pdf = f"Timesheet_{selected_inst_name.replace(' ', '_')}_{selected_year}_{selected_month}_{selected_cutoff}.pdf"
        st.download_button(
            label="📄 Download PDF Report",
            data=pdf_bytes,
            file_name=filename_pdf,
            mime="application/pdf",
            use_container_width=True
        )
        
    st.markdown("---")
    
    # Real-Time Computation Dashboard
    st.subheader("Real-Time Computation Dashboard")
    
    col_dash_1, col_dash_2, col_dash_3 = st.columns(3)
    
    with col_dash_1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">Grand Total Hours</div>
            <div class="metric-val">{grand_total:.2f} hrs</div>
            <p style='margin:0;font-size:0.8rem;color:#64748b;'>Sum of all hours in selected period</p>
        </div>
        """, unsafe_allow_html=True)
        
    with col_dash_2:
        # standard standard is 80 hours per period
        target = 80.0
        pct = min(100.0, (grand_total / target) * 100)
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">Target Load Progress</div>
            <div class="metric-val">{pct:.1f}%</div>
            <p style='margin:0;font-size:0.8rem;color:#64748b;'>Standard target period: 80.00 hrs</p>
        </div>
        """, unsafe_allow_html=True)
        
    with col_dash_3:
        status_pass = grand_total >= 80.0
        status_text = "PASSED (80h+ Standard)" if status_pass else "UNDER-HOURS"
        status_class = "metric-status-pass" if status_pass else "metric-status-warn"
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">Validation Status</div>
            <div class="metric-val {status_class}">{status_text}</div>
            <p style='margin:0;font-size:0.8rem;color:#64748b;'>Target standard verification</p>
        </div>
        """, unsafe_allow_html=True)
        
    # Render summaries side-by-side
    col_tbl_1, col_tbl_2 = st.columns([1, 2])
    
    with col_tbl_1:
        st.subheader("Summary per Activity")
        summary_rows = []
        for act in row_totals.index:
            summary_rows.append({"Activity": act, "Total Hours": f"{row_totals[act]:.1f} hrs"})
        st.table(pd.DataFrame(summary_rows))
        
    with col_tbl_2:
        st.subheader("Summary per Date Column")
        # Format daily totals into a neat grid
        daily_summary_df = pd.DataFrame(daily_totals).T
        daily_summary_df.index = ["Daily Total"]
        st.dataframe(daily_summary_df, use_container_width=True)

# ----------------- TAB 2: CLASS LOAD DIRECTORY -----------------
with tab_class_load:
    st.header(f"Class Load Directory (1st24) for {selected_inst_name}")
    
    if class_loads_df.empty:
        st.info("No class load schedules currently registered for this instructor.")
    else:
        # Calculate summary metrics
        reg_units = class_loads_df[class_loads_df["load_category"] == "Regular Load"]["units"].sum()
        overload_units = class_loads_df[class_loads_df["load_category"] == "Excess/Overload"]["units"].sum()
        comb_hours = class_loads_df["hours"].sum()
        
        # Display Metrics Panel
        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1:
            st.metric("Total Regular Units", f"{reg_units:.1f} Units")
        with col_m2:
            st.metric("Total Overload Units", f"{overload_units:.1f} Units")
        with col_m3:
            st.metric("Combined Weekly Hours", f"{comb_hours:.2f} hrs")
            
        # Display table with delete buttons
        st.subheader("Scheduled Course Lists")
        
        for _, row in class_loads_df.iterrows():
            with st.container():
                col_d1, col_d2, col_d3, col_d4, col_d5 = st.columns([2, 2, 2, 2, 1])
                with col_d1:
                    st.markdown(f"**{row['subject_name']}** ({row['type']})")
                    st.markdown(f"Code: `{row['section_code']}` | Room: `{row['room']}`")
                with col_d2:
                    st.markdown(f"📅 **Day:** `{row['day_of_week']}`")
                    st.markdown(f"🕒 **Time:** `{row['start_time']} - {row['end_time']}`")
                with col_d3:
                    st.markdown(f"Units: `{row['units']:.1f}`")
                    st.markdown(f"Duration: `{row['hours']:.1f} hrs`")
                with col_d4:
                    st.markdown(f"Category: **{row['load_category']}**")
                    st.markdown(f"Enrollment: `{row['enrollment']}`")
                with col_d5:
                    if st.button("❌ Remove", key=f"del_{row['id']}"):
                        delete_class_load(row['id'])
                        # clear overrides to force recalculation of dates
                        clear_timesheet_overrides(selected_inst_id, selected_year, selected_month, selected_cutoff)
                        st.success("Entry removed!")
                        st.rerun()
                st.markdown("---")
                
    # Add new load meeting schedule
    with st.expander("➕ Add Class Load Schedule Meeting"):
        add_form_name = st.text_input("Subject Name", placeholder="e.g. Computer Programming 5")
        add_form_code = st.text_input("Section Code", placeholder="e.g. ICTE201P")
        add_form_room = st.text_input("Room Location", placeholder="e.g. Lab 3 / Room 402")
        
        col_form_1, col_form_2 = st.columns(2)
        with col_form_1:
            add_form_day = st.selectbox("Meeting Weekday", ["M", "T", "W", "TH", "F", "S", "SU"])
            # Format times as time objects
            add_form_start = st.time_input("Start Time", datetime.time(8, 0))
            add_form_end = st.time_input("End Time", datetime.time(9, 30))
            add_form_type = st.selectbox("Meeting Type", ["LEC", "LAB"])
            
        with col_form_2:
            add_form_units = st.number_input("Units Value", min_value=0.0, max_value=6.0, value=3.0, step=0.5)
            # Pre-calculate hour duration but allow override
            duration_dt = datetime.datetime.combine(datetime.date.today(), add_form_end) - datetime.datetime.combine(datetime.date.today(), add_form_start)
            duration_hours = max(0.5, duration_dt.seconds / 3600.0)
            add_form_hours = st.number_input("Actual Meeting Hours", min_value=0.5, max_value=8.0, value=duration_hours, step=0.5)
            
            add_form_cat = st.selectbox("Load Category", ["Regular Load", "Excess/Overload", "Consultation"])
            add_form_enrol = st.number_input("Course Enrollment Count", min_value=0, max_value=200, value=30)
            
        if st.button("Submit Schedule Entry"):
            if add_form_name and add_form_code and add_form_room:
                # format start and end
                st_str = add_form_start.strftime("%H:%M")
                en_str = add_form_end.strftime("%H:%M")
                
                # Check for start/end time validity
                if add_form_end <= add_form_start:
                    st.error("End Time must be after Start Time.")
                else:
                    add_class_load(
                        selected_inst_id, add_form_name, add_form_code, add_form_room,
                        add_form_day, st_str, en_str, add_form_type,
                        add_form_units, add_form_hours, add_form_cat, add_form_enrol
                    )
                    # Clear overrides
                    clear_timesheet_overrides(selected_inst_id, selected_year, selected_month, selected_cutoff)
                    st.success("New Class Load added successfully!")
                    st.rerun()
            else:
                st.error("Please fill out all fields.")

# ----------------- TAB 3: MASTER TIMETABLE -----------------
with tab_schedule:
    st.header(f"Weekly Master Timetable grid ('Sched') - {selected_inst_name}")
    
    col_grid_1, col_grid_2 = st.columns([3, 1])
    
    with col_grid_2:
        # Summary panel
        st.subheader("Timetable Details")
        st.markdown(f"**Faculty Name:** {selected_inst_row['name']}")
        st.markdown(f"**Employment Status:** {selected_inst_row['employment_status']}")
        st.markdown(f"**Department:** {selected_inst_row.get('department', 'N/A')}")
        st.markdown(f"**Email:** {selected_inst_row.get('email', 'N/A')}")
        st.markdown(f"**Phone:** {selected_inst_row.get('phone', 'N/A')}")
        
        st.markdown("---")
        total_sched_units = class_loads_df["units"].sum()
        total_enrol_students = class_loads_df["enrollment"].sum()
        
        st.metric("Total Load Units", f"{total_sched_units:.1f} Units")
        st.metric("Total Student Enrollment", f"{total_enrol_students} Students")
        
    with col_grid_1:
        # Timetable generation
        # Slots list
        slots = []
        t_start = datetime.time(6, 30)
        t_end = datetime.time(21, 0)
        curr = datetime.datetime.combine(datetime.date.today(), t_start)
        end_dt = datetime.datetime.combine(datetime.date.today(), t_end)
        while curr <= end_dt:
            slots.append(curr.strftime("%H:%M"))
            curr += datetime.timedelta(minutes=15)
            
        # Map weekday index
        day_cols = ["M", "T", "W", "TH", "F", "S", "SU"]
        day_names = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        
        # Build 2D grid matrix
        # Cell formats: None or {'span': int, 'class': dict}
        grid_matrix = [[None for _ in range(7)] for _ in range(len(slots))]
        
        for _, row in class_loads_df.iterrows():
            day = row["day_of_week"]
            if day not in day_cols:
                continue
            day_idx = day_cols.index(day)
            
            try:
                s_str = row["start_time"]
                e_str = row["end_time"]
                
                # Find slot indices
                s_time = datetime.datetime.strptime(s_str, "%H:%M").time()
                e_time = datetime.datetime.strptime(e_str, "%H:%M").time()
                
                start_slot_idx = -1
                end_slot_idx = -1
                
                for idx, slot_str in enumerate(slots):
                    slot_t = datetime.datetime.strptime(slot_str, "%H:%M").time()
                    if slot_t == s_time:
                        start_slot_idx = idx
                    if slot_t == e_time:
                        end_slot_idx = idx - 1
                        
                if start_slot_idx != -1 and end_slot_idx != -1 and start_slot_idx <= end_slot_idx:
                    span = (end_slot_idx - start_slot_idx) + 1
                    
                    # Put info in top cell
                    grid_matrix[start_slot_idx][day_idx] = {
                        'span': span,
                        'class': row
                    }
                    
                    # Fill the rest with placeholder span=0 to skip rendering
                    for r in range(start_slot_idx + 1, end_slot_idx + 1):
                        grid_matrix[r][day_idx] = {
                            'span': 0,
                            'class': None
                        }
            except Exception as ex:
                continue
                
        # Generate HTML Table
        html = '<div class="timetable-container">'
        html += '<table class="timetable">'
        html += '<thead><tr><th>Time</th>'
        for dn in day_names:
            html += f'<th>{dn}</th>'
        html += '</tr></thead><tbody>'
        
        for r_idx, slot in enumerate(slots):
            html += f'<tr><td class="time-col">{slot}</td>'
            for c_idx in range(7):
                cell = grid_matrix[r_idx][c_idx]
                is_weekend = c_idx in (5, 6) # Sat, Sun
                cell_class = "weekend-cell-theme" if is_weekend else ""
                
                if cell is None:
                    html += f'<td class="{cell_class}"></td>'
                elif cell['span'] > 0:
                    cl = cell['class']
                    # Theme based on load category and type
                    cat = cl["load_category"]
                    t_type = cl["type"]
                    theme = "class-lec-theme"
                    if cat == "Consultation":
                        theme = "class-consult-theme"
                    elif t_type == "LAB":
                        theme = "class-lab-theme"
                        
                    content = f"""
                    <div class="class-block-container {theme}">
                        <div style='font-size:0.75rem;font-weight:700;'>{cl['subject_name']}</div>
                        <div style='font-size:0.65rem;'>Section: {cl['section_code']}</div>
                        <div style='font-size:0.65rem;'>Room: {cl['room']}</div>
                        <div style='font-size:0.65rem;'>({cl['type']})</div>
                    </div>
                    """
                    html += f'<td rowspan="{cell["span"]}" class="{cell_class}" style="padding: 2px;">{content}</td>'
                elif cell['span'] == 0:
                    # Skipped because of rowspan
                    pass
            html += '</tr>'
            
        html += '</tbody></table></div>'
        
        # Display schedule HTML
        st.markdown(html, unsafe_allow_html=True)
