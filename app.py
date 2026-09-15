import streamlit as st
import pandas as pd
import datetime
import calendar
from date_engine import generate_cutoff_dates, get_previous_month
from database import (
    init_db, get_instructors, add_instructor, update_instructor, delete_instructor,
    set_instructor_academic_level,
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
    
    /* Timesheet Section Banners */
    .section-banner {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 10px 16px;
        border-radius: 8px;
        margin-top: 14px;
        margin-bottom: 8px;
        font-weight: 700;
        font-size: 0.95rem;
    }
    .section-banner-reg {
        background: linear-gradient(90deg, rgba(30, 58, 138, 0.15) 0%, rgba(30, 58, 138, 0.05) 100%);
        border-left: 5px solid #1e3a8a;
        color: #1e3a8a;
    }
    .section-banner-exc {
        background: linear-gradient(90deg, rgba(217, 119, 6, 0.15) 0%, rgba(217, 119, 6, 0.05) 100%);
        border-left: 5px solid #d97706;
        color: #b45309;
    }
    .section-banner-non {
        background: linear-gradient(90deg, rgba(79, 70, 229, 0.15) 0%, rgba(79, 70, 229, 0.05) 100%);
        border-left: 5px solid #4f46e5;
        color: #4338ca;
    }
    .section-tag {
        font-size: 0.75rem;
        padding: 3px 8px;
        border-radius: 12px;
        font-weight: 600;
        background: rgba(255, 255, 255, 0.85);
        box-shadow: 0 1px 2px rgba(0,0,0,0.06);
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
        new_level = st.selectbox("Academic Level", ["Tertiary", "SHS"], key="init_level")
        new_dept = st.text_input("Department", "Computer Science", key="init_dept")
        new_email = st.text_input("Email", key="init_email")
        new_phone = st.text_input("Phone", key="init_phone")
        if st.button("Register Instructor", key="init_btn"):
            if new_name:
                add_instructor(new_name, new_email, new_phone, new_dept, new_status, new_level)
                st.rerun()
    st.stop()

# Populate selectors
instructor_options = {row['name']: row['id'] for _, row in instructors_df.iterrows()}
selected_inst_name = st.sidebar.selectbox("Select Instructor", list(instructor_options.keys()))
selected_inst_id = instructor_options[selected_inst_name]
selected_inst_row = instructors_df[instructors_df['id'] == selected_inst_id].iloc[0]

# Academic Level Selection (Tertiary: 24 units max, SHS: 30 units max)
current_level = selected_inst_row.get("academic_level", "Tertiary")
if pd.isna(current_level) or not current_level:
    current_level = "Tertiary"

selected_level = st.sidebar.radio(
    "Academic Teaching Level",
    ["Tertiary", "SHS"],
    index=0 if current_level == "Tertiary" else 1,
    horizontal=True,
    help="Select teaching level. Tertiary max regular load: 24 units. SHS max regular load: 30 units."
)

if selected_level != current_level:
    set_instructor_academic_level(selected_inst_id, selected_level)
    st.rerun()

max_reg_units = 30.0 if selected_level == "SHS" else 24.0

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
    add_level = st.selectbox("Academic Level", ["Tertiary", "SHS"], key="add_level")
    add_dept = st.text_input("Department", key="add_dept")
    add_email = st.text_input("Email", key="add_email")
    add_phone = st.text_input("Phone", key="add_phone")
    if st.button("Save New Profile"):
        if add_name:
            add_instructor(add_name, add_email, add_phone, add_dept, add_status, add_level)
            st.success(f"Added {add_name}")
            st.rerun()
            
    st.markdown("---")
    st.subheader("Edit Current Instructor")
    edit_name = st.text_input("Edit Name", value=selected_inst_row["name"])
    edit_status = st.selectbox("Edit Status", ["SC-Based", "Part-Time", "Pro-Rated"], index=["SC-Based", "Part-Time", "Pro-Rated"].index(selected_inst_row["employment_status"]))
    edit_level = st.selectbox("Edit Academic Level", ["Tertiary", "SHS"], index=0 if current_level == "Tertiary" else 1)
    edit_dept = st.text_input("Edit Department", value=selected_inst_row.get("department", ""))
    edit_email = st.text_input("Edit Email", value=selected_inst_row.get("email", ""))
    edit_phone = st.text_input("Edit Phone", value=selected_inst_row.get("phone", ""))
    col_edit_1, col_edit_2 = st.columns(2)
    with col_edit_1:
        if st.button("Update Profile"):
            update_instructor(selected_inst_id, edit_name, edit_email, edit_phone, edit_dept, edit_status, edit_level)
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
st.markdown(
    f"**Current Instructor:** `{selected_inst_row['name']}` | "
    f"**Status:** `{selected_inst_row['employment_status']}` | "
    f"**Level:** `{selected_level}` (Max Regular: **{max_reg_units:.0f} units**) | "
    f"**Department:** `{selected_inst_row.get('department', 'N/A')}`"
)

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
    
    # Classes groups
    regular_classes = class_loads_df[class_loads_df["load_category"] == "Regular Load"]
    excess_classes = class_loads_df[class_loads_df["load_category"].isin(["Excess Load", "Excess/Overload"])]
    consultations = class_loads_df[class_loads_df["load_category"] == "Consultation"]
    
    # Load overrides from SQLite
    db_overrides = get_timesheet_overrides(selected_inst_id, selected_year, selected_month, selected_cutoff)
    
    # Column configuration for editable grids
    col_config = {
        col: st.column_config.NumberColumn(
            min_value=0.0,
            max_value=24.0,
            step=0.5,
            format="%.1f"
        ) for col in date_columns
    }
    
    st.markdown("⚠️ *Adjust hours inside the cell grid below. Subtotals and validation compute in real-time. Make sure to click **Save Timesheet Changes** to persist overrides.*")
    
    # ---------------- 1. REGULAR LOAD TABLE ----------------
    reg_grid_data = {}
    for _, cl in regular_classes.iterrows():
        key_name = f"{cl['type']}: {cl['subject_name']} ({cl['section_code']})"
        row_values = []
        for d in cutoff_dates:
            date_str = d['date'].strftime("%Y-%m-%d")
            val = db_overrides.get((key_name, date_str))
            if val is None:
                val = cl['hours'] if cl['day_of_week'] == d['day_abbr'] else 0.0
            row_values.append(val)
        reg_grid_data[key_name] = row_values
        
    reg_timesheet_df = pd.DataFrame.from_dict(reg_grid_data, orient='index', columns=date_columns) if reg_grid_data else pd.DataFrame(columns=date_columns)
    reg_timesheet_df.index.name = "Course / Subject"
    
    st.markdown(f"""
    <div class="section-banner section-banner-reg">
        <span>📘 1. Regular Load Timesheet</span>
        <span class="section-tag" style="color: #1e3a8a;">Cap: {max_reg_units:.0f} Units ({selected_level}) • {len(regular_classes)} Course(s)</span>
    </div>
    """, unsafe_allow_html=True)
    
    if not reg_timesheet_df.empty:
        edited_reg_df = st.data_editor(
            reg_timesheet_df,
            use_container_width=True,
            column_config=col_config,
            key=f"editor_reg_{selected_inst_id}_{selected_year}_{selected_month}_{selected_cutoff}"
        )
    else:
        st.info("ℹ️ No Regular Load courses registered for this instructor.")
        edited_reg_df = pd.DataFrame(columns=date_columns)
        
    # ---------------- 2. EXCESS LOAD TABLE ----------------
    exc_grid_data = {}
    for _, cl in excess_classes.iterrows():
        key_name = f"{cl['type']}: {cl['subject_name']} ({cl['section_code']})"
        row_values = []
        for d in cutoff_dates:
            date_str = d['date'].strftime("%Y-%m-%d")
            val = db_overrides.get((key_name, date_str))
            if val is None:
                val = cl['hours'] if cl['day_of_week'] == d['day_abbr'] else 0.0
            row_values.append(val)
        exc_grid_data[key_name] = row_values
        
    exc_timesheet_df = pd.DataFrame.from_dict(exc_grid_data, orient='index', columns=date_columns) if exc_grid_data else pd.DataFrame(columns=date_columns)
    exc_timesheet_df.index.name = "Course / Subject"
    
    st.markdown(f"""
    <div class="section-banner section-banner-exc">
        <span>⚡ 2. Excess Load Timesheet</span>
        <span class="section-tag" style="color: #b45309;">Overload / Additional Teaching • {len(excess_classes)} Course(s)</span>
    </div>
    """, unsafe_allow_html=True)
    
    if not exc_timesheet_df.empty:
        edited_exc_df = st.data_editor(
            exc_timesheet_df,
            use_container_width=True,
            column_config=col_config,
            key=f"editor_exc_{selected_inst_id}_{selected_year}_{selected_month}_{selected_cutoff}"
        )
    else:
        st.info("ℹ️ No Excess Load courses registered for this instructor.")
        edited_exc_df = pd.DataFrame(columns=date_columns)
        
    # ---------------- 3. NON-TEACHING SC / HQ TABLE ----------------
    non_teaching_rows = [
        ("Non-Teaching SC", "Career Orientation Seminars (COS)"),
        ("Non-Teaching SC", "Guidance/Counseling"),
        ("Non-Teaching HQ", "Consultation"),
        ("Non-Teaching HQ", "Administrative Hours")
    ]
    non_grid_data = {}
    for cat, key_name in non_teaching_rows:
        row_values = []
        for d in cutoff_dates:
            date_str = d['date'].strftime("%Y-%m-%d")
            val = db_overrides.get((key_name, date_str))
            if val is None:
                if key_name == "Consultation":
                    val = 0.0
                    for _, cl in consultations.iterrows():
                        if cl['day_of_week'] == d['day_abbr']:
                            val = cl['hours']
                            break
                else:
                    val = 0.0
            row_values.append(val)
        non_grid_data[key_name] = row_values
        
    non_timesheet_df = pd.DataFrame.from_dict(non_grid_data, orient='index', columns=date_columns)
    non_timesheet_df.index.name = "Activity Description"
    
    st.markdown("""
    <div class="section-banner section-banner-non">
        <span>🏢 3. Non-Teaching Activities (SC & HQ)</span>
        <span class="section-tag" style="color: #4338ca;">Student Consultation, Guidance, Seminars & Admin</span>
    </div>
    """, unsafe_allow_html=True)
    
    edited_non_df = st.data_editor(
        non_timesheet_df,
        use_container_width=True,
        column_config=col_config,
        key=f"editor_non_{selected_inst_id}_{selected_year}_{selected_month}_{selected_cutoff}"
    )
    
    # Compute subtotals and grand totals in real-time
    reg_hours = float(edited_reg_df.sum().sum()) if not edited_reg_df.empty else 0.0
    exc_hours = float(edited_exc_df.sum().sum()) if not edited_exc_df.empty else 0.0
    non_hours = float(edited_non_df.sum().sum()) if not edited_non_df.empty else 0.0
    grand_total = reg_hours + exc_hours + non_hours
    
    # Save/Reset controls
    st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
    col_ctrl_1, col_ctrl_2, col_ctrl_3, col_ctrl_4 = st.columns(4)
    with col_ctrl_1:
        if st.button("💾 Save Timesheet Changes", use_container_width=True):
            # Parse edits from all 3 tables back to SQLite format
            overrides_list = []
            for df_to_save in [edited_reg_df, edited_exc_df, edited_non_df]:
                if not df_to_save.empty:
                    for act_name in df_to_save.index:
                        for idx, d_info in enumerate(cutoff_dates):
                            col_name = date_columns[idx]
                            hours_val = float(df_to_save.loc[act_name, col_name])
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
    
    col_dash_1, col_dash_2, col_dash_3, col_dash_4 = st.columns(4)
    
    with col_dash_1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">Grand Total Hours</div>
            <div class="metric-val">{grand_total:.2f} hrs</div>
            <p style='margin:0;font-size:0.8rem;color:#64748b;'>All hours in cutoff period</p>
        </div>
        """, unsafe_allow_html=True)
        
    with col_dash_2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">Regular Load Hours</div>
            <div class="metric-val" style="color: #15803d;">{reg_hours:.2f} hrs</div>
            <p style='margin:0;font-size:0.8rem;color:#64748b;'>Regular teaching courses</p>
        </div>
        """, unsafe_allow_html=True)
        
    with col_dash_3:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">Excess Load Hours</div>
            <div class="metric-val" style="color: #b45309;">{exc_hours:.2f} hrs</div>
            <p style='margin:0;font-size:0.8rem;color:#64748b;'>Overload teaching courses</p>
        </div>
        """, unsafe_allow_html=True)
        
    with col_dash_4:
        status_pass = grand_total >= 80.0
        status_text = "PASSED (80h+)" if status_pass else "UNDER-HOURS"
        status_class = "metric-status-pass" if status_pass else "metric-status-warn"
        pct = min(100.0, (grand_total / 80.0) * 100)
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">Non-Teaching Hours: {non_hours:.1f}h</div>
            <div class="metric-val {status_class}">{status_text}</div>
            <p style='margin:0;font-size:0.8rem;color:#64748b;'>Progress: {pct:.0f}% of 80h standard</p>
        </div>
        """, unsafe_allow_html=True)
        
    # Render summaries side-by-side
    col_tbl_1, col_tbl_2 = st.columns([1, 2])
    
    with col_tbl_1:
        st.subheader("Summary per Activity")
        summary_rows = []
        if not edited_reg_df.empty:
            for act in edited_reg_df.index:
                summary_rows.append({"Category": "Regular Load", "Activity": act, "Total Hours": f"{edited_reg_df.loc[act].sum():.1f} hrs"})
        if not edited_exc_df.empty:
            for act in edited_exc_df.index:
                summary_rows.append({"Category": "Excess Load", "Activity": act, "Total Hours": f"{edited_exc_df.loc[act].sum():.1f} hrs"})
        if not edited_non_df.empty:
            for act in edited_non_df.index:
                cat_tag = "Non-Teaching HQ" if act in ("Consultation", "Administrative Hours") else "Non-Teaching SC"
                summary_rows.append({"Category": cat_tag, "Activity": act, "Total Hours": f"{edited_non_df.loc[act].sum():.1f} hrs"})
        st.table(pd.DataFrame(summary_rows) if summary_rows else pd.DataFrame(columns=["Category", "Activity", "Total Hours"]))
        
    with col_tbl_2:
        st.subheader("Summary per Date Column")
        # Format daily totals into a neat grid
        daily_totals = pd.Series(0.0, index=date_columns)
        for df_item in [edited_reg_df, edited_exc_df, edited_non_df]:
            if not df_item.empty:
                daily_totals = daily_totals.add(df_item.sum(axis=0), fill_value=0.0)
        daily_summary_df = pd.DataFrame(daily_totals).T
        daily_summary_df.index = ["Daily Total"]
        st.dataframe(daily_summary_df, use_container_width=True)

# ----------------- TAB 2: CLASS LOAD DIRECTORY -----------------
with tab_class_load:
    st.header(f"Class Load Directory for {selected_inst_name} ({selected_level})")
    
    # Calculate summary metrics
    reg_df = class_loads_df[class_loads_df["load_category"] == "Regular Load"]
    excess_df = class_loads_df[class_loads_df["load_category"].isin(["Excess Load", "Excess/Overload"])]
    consult_df = class_loads_df[class_loads_df["load_category"] == "Consultation"]
    
    reg_units = reg_df["units"].sum() if not reg_df.empty else 0.0
    excess_units = excess_df["units"].sum() if not excess_df.empty else 0.0
    total_units = reg_units + excess_units
    comb_hours = class_loads_df["hours"].sum() if not class_loads_df.empty else 0.0
    
    # Display Metrics Panel
    col_m1, col_m2, col_m3, col_m4 = st.columns(4)
    with col_m1:
        st.metric(
            "Regular Load Units",
            f"{reg_units:.1f} / {max_reg_units:.0f} u",
            delta=f"{selected_level} Limit: {max_reg_units:.0f} max",
            delta_color="normal"
        )
    with col_m2:
        st.metric("Excess Load Units", f"{excess_units:.1f} Units")
    with col_m3:
        st.metric("Total Load Units", f"{total_units:.1f} Units")
    with col_m4:
        st.metric("Combined Weekly Hours", f"{comb_hours:.2f} hrs")
        
    # Cap validation alert
    if reg_units > max_reg_units:
        st.error(
            f"⚠️ **Regular Load Limit Exceeded!** Regular load is currently **{reg_units:.1f} units**, "
            f"exceeding the maximum allowed **{max_reg_units:.0f} units** for **{selected_level}** instructors. "
            f"Please change excess course(s) to **Excess Load**."
        )
    elif reg_units == max_reg_units:
        st.info(f"ℹ️ Regular load is exactly at the **{max_reg_units:.0f} units** maximum limit for **{selected_level}**.")
    else:
        st.success(
            f"✅ Regular load is within limits: **{reg_units:.1f} units** used of **{max_reg_units:.0f} max** for **{selected_level}** "
            f"({max_reg_units - reg_units:.1f} units available for Regular Load)."
        )
        
    st.markdown("---")
    
    if class_loads_df.empty:
        st.info("No class load schedules currently registered for this instructor.")
    else:
        # Separate Section 1: Regular Load
        st.subheader(f"📘 Regular Load Courses ({len(reg_df)} classes • {reg_units:.1f} Units)")
        if reg_df.empty:
            st.caption("No Regular Load courses assigned.")
        else:
            for _, row in reg_df.iterrows():
                with st.container():
                    col_d1, col_d2, col_d3, col_d4 = st.columns([3, 2, 2, 1])
                    with col_d1:
                        st.markdown(f"**{row['subject_name']}** (`{row['type']}`)")
                        st.markdown(f"Code: `{row['section_code']}` | Room: `{row['room']}`")
                    with col_d2:
                        st.markdown(f"📅 **Day:** `{row['day_of_week']}`")
                        st.markdown(f"🕒 **Time:** `{row['start_time']} - {row['end_time']}`")
                    with col_d3:
                        st.markdown(f"Units: **{row['units']:.1f} u**")
                        st.markdown(f"Duration: `{row['hours']:.1f} hrs`")
                    with col_d4:
                        if st.button("❌ Remove", key=f"del_reg_{row['id']}"):
                            delete_class_load(row['id'])
                            clear_timesheet_overrides(selected_inst_id, selected_year, selected_month, selected_cutoff)
                            st.success("Entry removed!")
                            st.rerun()
                    st.markdown("<hr style='margin: 8px 0; border: none; border-top: 1px dashed #e2e8f0;'>", unsafe_allow_html=True)
                    
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Separate Section 2: Excess Load
        st.subheader(f"⚡ Excess Load Courses ({len(excess_df)} classes • {excess_units:.1f} Units)")
        if excess_df.empty:
            st.caption("No Excess Load courses assigned.")
        else:
            for _, row in excess_df.iterrows():
                with st.container():
                    col_d1, col_d2, col_d3, col_d4 = st.columns([3, 2, 2, 1])
                    with col_d1:
                        st.markdown(f"**{row['subject_name']}** (`{row['type']}`)")
                        st.markdown(f"Code: `{row['section_code']}` | Room: `{row['room']}`")
                    with col_d2:
                        st.markdown(f"📅 **Day:** `{row['day_of_week']}`")
                        st.markdown(f"🕒 **Time:** `{row['start_time']} - {row['end_time']}`")
                    with col_d3:
                        st.markdown(f"Units: **{row['units']:.1f} u**")
                        st.markdown(f"Duration: `{row['hours']:.1f} hrs`")
                    with col_d4:
                        if st.button("❌ Remove", key=f"del_exc_{row['id']}"):
                            delete_class_load(row['id'])
                            clear_timesheet_overrides(selected_inst_id, selected_year, selected_month, selected_cutoff)
                            st.success("Entry removed!")
                            st.rerun()
                    st.markdown("<hr style='margin: 8px 0; border: none; border-top: 1px dashed #e2e8f0;'>", unsafe_allow_html=True)

        if not consult_df.empty:
            st.markdown("<br>", unsafe_allow_html=True)
            st.subheader(f"🕒 Consultation Hours ({len(consult_df)} sessions)")
            for _, row in consult_df.iterrows():
                with st.container():
                    col_d1, col_d2, col_d3, col_d4 = st.columns([3, 2, 2, 1])
                    with col_d1:
                        st.markdown(f"**{row['subject_name']}**")
                        st.markdown(f"Room: `{row['room']}`")
                    with col_d2:
                        st.markdown(f"📅 **Day:** `{row['day_of_week']}`")
                        st.markdown(f"🕒 **Time:** `{row['start_time']} - {row['end_time']}`")
                    with col_d3:
                        st.markdown(f"Duration: `{row['hours']:.1f} hrs`")
                    with col_d4:
                        if st.button("❌ Remove", key=f"del_con_{row['id']}"):
                            delete_class_load(row['id'])
                            clear_timesheet_overrides(selected_inst_id, selected_year, selected_month, selected_cutoff)
                            st.success("Entry removed!")
                            st.rerun()
                    st.markdown("<hr style='margin: 8px 0; border: none; border-top: 1px dashed #e2e8f0;'>", unsafe_allow_html=True)

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
            add_form_cat = st.selectbox("Load Category", ["Regular Load", "Excess Load", "Consultation"])
            
        if add_form_cat == "Regular Load" and (reg_units + add_form_units) > max_reg_units:
            st.warning(
                f"⚠️ **Limit Notice:** Adding this {add_form_units:.1f}-unit course under Regular Load will exceed the "
                f"**{max_reg_units:.0f} units max limit** for **{selected_level}** (Total would be {reg_units + add_form_units:.1f} units). "
                f"Please consider designating this entry as **Excess Load**."
            )
            
        if st.button("Submit Schedule Entry"):
            if add_form_name and add_form_code and add_form_room:
                st_str = add_form_start.strftime("%H:%M")
                en_str = add_form_end.strftime("%H:%M")
                
                if add_form_end <= add_form_start:
                    st.error("End Time must be after Start Time.")
                else:
                    add_class_load(
                        selected_inst_id, add_form_name, add_form_code, add_form_room,
                        add_form_day, st_str, en_str, add_form_type,
                        add_form_units, add_form_hours, add_form_cat, 0
                    )
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
        st.markdown(f"**Academic Level:** `{selected_level}` (Max Reg: **{max_reg_units:.0f}u**)")
        st.markdown(f"**Department:** {selected_inst_row.get('department', 'N/A')}")
        st.markdown(f"**Email:** {selected_inst_row.get('email', 'N/A')}")
        st.markdown(f"**Phone:** {selected_inst_row.get('phone', 'N/A')}")
        
        st.markdown("---")
        total_sched_units = class_loads_df["units"].sum() if not class_loads_df.empty else 0.0
        reg_sched_units = class_loads_df[class_loads_df["load_category"] == "Regular Load"]["units"].sum() if not class_loads_df.empty else 0.0
        excess_sched_units = class_loads_df[class_loads_df["load_category"].isin(["Excess Load", "Excess/Overload"])]["units"].sum() if not class_loads_df.empty else 0.0
        
        st.metric("Total Load Units", f"{total_sched_units:.1f} Units")
        st.metric("Regular Load Units", f"{reg_sched_units:.1f} / {max_reg_units:.0f} u")
        st.metric("Excess Load Units", f"{excess_sched_units:.1f} Units")
        
    with col_grid_1:
        # Timetable generation
        slots = []
        t_start = datetime.time(6, 30)
        t_end = datetime.time(21, 0)
        curr = datetime.datetime.combine(datetime.date.today(), t_start)
        end_dt = datetime.datetime.combine(datetime.date.today(), t_end)
        while curr <= end_dt:
            slots.append(curr.strftime("%H:%M"))
            curr += datetime.timedelta(minutes=15)
            
        day_cols = ["M", "T", "W", "TH", "F", "S", "SU"]
        day_names = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        
        # Build 2D grid matrix
        grid_matrix = [[None for _ in range(7)] for _ in range(len(slots))]
        
        for _, row in class_loads_df.iterrows():
            day = row["day_of_week"]
            if day not in day_cols:
                continue
            day_idx = day_cols.index(day)
            
            try:
                s_str = row["start_time"]
                e_str = row["end_time"]
                
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
                    grid_matrix[start_slot_idx][day_idx] = {
                        'span': span,
                        'class': row
                    }
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
                    cat = cl["load_category"]
                    t_type = cl["type"]
                    theme = "class-lec-theme"
                    if cat == "Consultation":
                        theme = "class-consult-theme"
                    elif t_type == "LAB":
                        theme = "class-lab-theme"
                        
                    badge_bg = "#dcfce7" if cat == "Regular Load" else ("#fef3c7" if cat in ("Excess Load", "Excess/Overload") else "#f3f4f6")
                    badge_color = "#166534" if cat == "Regular Load" else ("#92400e" if cat in ("Excess Load", "Excess/Overload") else "#374151")
                    badge_text = "REGULAR" if cat == "Regular Load" else ("EXCESS" if cat in ("Excess Load", "Excess/Overload") else "CONSULT")
                    
                    content = f"""
                    <div class="class-block-container {theme}">
                        <div style='font-size:0.75rem;font-weight:700;'>{cl['subject_name']}</div>
                        <div style='font-size:0.65rem;'>Section: {cl['section_code']}</div>
                        <div style='font-size:0.65rem;'>Room: {cl['room']}</div>
                        <div style='font-size:0.65rem;'>({cl['type']}) <span style='background:{badge_bg}; color:{badge_color}; font-size:0.55rem; padding:1px 3px; border-radius:3px; font-weight:700;'>{badge_text}</span></div>
                    </div>
                    """
                    html += f'<td rowspan="{cell["span"]}" class="{cell_class}" style="padding: 2px;">{content}</td>'
                elif cell['span'] == 0:
                    pass
            html += '</tr>'
            
        html += '</tbody></table></div>'
        st.markdown(html, unsafe_allow_html=True)
