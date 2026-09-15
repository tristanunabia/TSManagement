import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import datetime
import calendar
import pandas as pd
from date_engine import generate_cutoff_dates, get_previous_month

# Color constants
PRIMARY_COLOR = "1F4E78"        # Navy
SECONDARY_COLOR = "D9E1F2"      # Ice Blue
WEEKEND_COLOR = "F2F2F2"        # Light Grey
TEXT_LIGHT = "FFFFFF"           # White
BORDER_COLOR = "D9D9D9"         # Thin Grey
ACCENT_GREEN = "E2EFDA"         # Soft Green (LEC)
ACCENT_BLUE = "DDEBF7"          # Soft Blue (LAB)
ACCENT_ORANGE = "FFF2CC"        # Soft Orange (Consultation)
EXCESS_HEADER_COLOR = "C65911"  # Amber/Orange (Excess Load)
EXCESS_SUB_COLOR = "FCE4D6"     # Peach (Excess Subtotal)
NON_TEACH_COLOR = "2E4053"      # Slate Navy (Non-Teaching)
GRAND_TOTAL_COLOR = "C6EFCE"    # Soft Green (Grand Total)

def get_thin_border():
    thin = Side(border_style="thin", color=BORDER_COLOR)
    return Border(left=thin, right=thin, top=thin, bottom=thin)

def generate_excel_timesheet(instructor_row, class_loads_df, year, month, selected_cutoff, timesheet_data):
    """
    Generates an Excel workbook byte stream.
    instructor_row: dict or sqlite3.Row with instructor details
    class_loads_df: Pandas DataFrame of instructor's class loads
    year, month, selected_cutoff: parameters of the selected timesheet
    timesheet_data: dict containing edited cell values {(row_name, date_str): hours}
    """
    wb = openpyxl.Workbook()
    
    # Remove the default sheet
    default_sheet = wb.active
    wb.remove(default_sheet)
    
    # 1. Create Sched sheet
    create_sched_sheet(wb, instructor_row, class_loads_df)
    
    # 2. Create Class Load sheet
    create_class_load_sheet(wb, class_loads_df)
    
    # 3. Create Timesheet sheets
    # We create BOTH sheet structures, but we hide the inactive one
    ws_11_25 = create_timesheet_sheet(wb, "Timesheet 11-25", instructor_row, class_loads_df, year, month, "11-25", timesheet_data)
    ws_26_10 = create_timesheet_sheet(wb, "Timesheet 26-10", instructor_row, class_loads_df, year, month, "26-10", timesheet_data)
    
    # Control sheet visibility
    if selected_cutoff == "11-25":
        ws_26_10.sheet_state = 'hidden'
        ws_11_25.sheet_state = 'visible'
        # Set active sheet to 11-25 (which is index 2, Sched=0, Class Load=1, 11-25=2)
        wb.active = wb.sheetnames.index("Timesheet 11-25")
    else:
        ws_11_25.sheet_state = 'hidden'
        ws_26_10.sheet_state = 'visible'
        wb.active = wb.sheetnames.index("Timesheet 26-10")
        
    # Write workbook to in-memory bytes
    fp = io.BytesIO()
    wb.save(fp)
    fp.seek(0)
    return fp.getvalue()

def create_sched_sheet(wb, instructor, class_loads):
    ws = wb.create_sheet(title="Sched")
    ws.views.sheetView[0].showGridLines = True
    
    # Title Block
    ws.merge_cells("A1:H1")
    ws["A1"] = "WEEKLY MASTER SCHEDULE"
    ws["A1"].font = Font(name="Calibri", size=16, bold=True, color=TEXT_LIGHT)
    ws["A1"].fill = PatternFill(start_color=PRIMARY_COLOR, end_color=PRIMARY_COLOR, fill_type="solid")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 40
    
    # Faculty Details
    details = [
        ("Faculty Name:", instructor["name"], "Dept/College:", instructor["department"]),
        ("Status:", instructor["employment_status"], "Email:", instructor.get("email", "N/A"))
    ]
    for r_idx, row_data in enumerate(details, start=2):
        ws.row_dimensions[r_idx].height = 20
        ws.cell(row=r_idx, column=1, value=row_data[0]).font = Font(name="Calibri", bold=True)
        ws.cell(row=r_idx, column=2, value=row_data[1])
        ws.cell(row=r_idx, column=4, value=row_data[2]).font = Font(name="Calibri", bold=True)
        ws.cell(row=r_idx, column=5, value=row_data[3])
    
    # Schedule Grid Headers
    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    ws.cell(row=5, column=1, value="Time").font = Font(bold=True, color=TEXT_LIGHT)
    ws.cell(row=5, column=1).fill = PatternFill(start_color=PRIMARY_COLOR, end_color=PRIMARY_COLOR, fill_type="solid")
    ws.cell(row=5, column=1).alignment = Alignment(horizontal="center")
    
    for c_idx, day in enumerate(days, start=2):
        cell = ws.cell(row=5, column=c_idx, value=day)
        cell.font = Font(bold=True, color=TEXT_LIGHT)
        cell.fill = PatternFill(start_color=PRIMARY_COLOR, end_color=PRIMARY_COLOR, fill_type="solid")
        cell.alignment = Alignment(horizontal="center")
        
    ws.row_dimensions[5].height = 25
    
    # Generate Time Slots (06:30 to 21:00, 15 min steps)
    time_slots = []
    start_t = datetime.time(6, 30)
    end_t = datetime.time(21, 0)
    curr_t = datetime.datetime.combine(datetime.date.today(), start_t)
    end_dt = datetime.datetime.combine(datetime.date.today(), end_t)
    
    while curr_t <= end_dt:
        time_slots.append(curr_t.time())
        curr_t += datetime.timedelta(minutes=15)
        
    # Map weekdays to columns
    day_col_map = {"M": 2, "T": 3, "W": 4, "TH": 5, "F": 6, "S": 7, "SU": 8}
    
    # Initialize empty cells with thin borders
    for r_idx, t in enumerate(time_slots, start=6):
        ws.row_dimensions[r_idx].height = 18
        # Format time display
        t_str = t.strftime("%H:%M")
        cell = ws.cell(row=r_idx, column=1, value=t_str)
        cell.font = Font(size=9)
        cell.alignment = Alignment(horizontal="center")
        cell.border = get_thin_border()
        
        for col_idx in range(2, 9):
            grid_cell = ws.cell(row=r_idx, column=col_idx)
            grid_cell.border = get_thin_border()
            # Highlight Saturday (col 7) and Sunday (col 8)
            if col_idx in (7, 8):
                grid_cell.fill = PatternFill(start_color=WEEKEND_COLOR, end_color=WEEKEND_COLOR, fill_type="solid")
                
    # Place classes and merge slots
    # Standard time intervals are 15 mins. E.g. start at 08:30, end at 10:00 is 6 slots.
    # We find the start row and end row, then merge.
    for _, row in class_loads.iterrows():
        day = row["day_of_week"]
        col_idx = day_col_map.get(day)
        if not col_idx:
            continue
            
        try:
            start_str = row["start_time"]
            end_str = row["end_time"]
            
            s_time = datetime.datetime.strptime(start_str, "%H:%M").time()
            e_time = datetime.datetime.strptime(end_str, "%H:%M").time()
            
            # Find row index matching start_time and end_time
            start_r_idx = -1
            end_r_idx = -1
            
            for r_idx, t in enumerate(time_slots, start=6):
                if t == s_time:
                    start_r_idx = r_idx
                if t == e_time:
                    end_r_idx = r_idx - 1 # class finishes before this final slot starts
                    
            if start_r_idx != -1 and end_r_idx != -1 and start_r_idx <= end_r_idx:
                # Merge cells in this column from start_r_idx to end_r_idx
                ws.merge_cells(start_row=start_r_idx, start_column=col_idx, end_row=end_r_idx, end_column=col_idx)
                
                # Style top cell
                top_cell = ws.cell(row=start_r_idx, column=col_idx)
                top_cell.value = f"{row['subject_name']}\n{row['section_code']}\n{row['room']}\n({row['type']})"
                top_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                top_cell.font = Font(size=9, bold=True)
                
                # Determine soft color based on type & category
                cat = row["load_category"]
                ctype = row["type"]
                if cat == "Consultation":
                    fill_c = ACCENT_ORANGE
                elif ctype == "LEC":
                    fill_c = ACCENT_GREEN
                else:
                    fill_c = ACCENT_BLUE
                    
                p_fill = PatternFill(start_color=fill_c, end_color=fill_c, fill_type="solid")
                for r in range(start_r_idx, end_r_idx + 1):
                    ws.cell(row=r, column=col_idx).fill = p_fill
        except Exception as e:
            # Skip invalid times or formats
            continue
            
    # Set column widths
    ws.column_dimensions["A"].width = 10
    for col in ["B", "C", "D", "E", "F", "G", "H"]:
        ws.column_dimensions[col].width = 18

def create_class_load_sheet(wb, class_loads):
    ws = wb.create_sheet(title="Class Load")
    ws.views.sheetView[0].showGridLines = True
    thin_border = get_thin_border()
    
    # Title Banner
    ws.merge_cells("A1:J1")
    ws["A1"] = "FACULTY CLASS LOAD & TEACHING DIRECTORY"
    ws["A1"].font = Font(name="Calibri", size=15, bold=True, color=TEXT_LIGHT)
    ws["A1"].fill = PatternFill(start_color=PRIMARY_COLOR, end_color=PRIMARY_COLOR, fill_type="solid")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 35
    
    headers = [
        "Subject Name", "Section Code", "Room", "Day", "Start Time", 
        "End Time", "Type", "Units", "Hours", "Load Category"
    ]
    
    regular_classes = class_loads[class_loads["load_category"] == "Regular Load"] if not class_loads.empty else pd.DataFrame()
    excess_classes = class_loads[class_loads["load_category"].isin(["Excess Load", "Excess/Overload"])] if not class_loads.empty else pd.DataFrame()
    consultations = class_loads[class_loads["load_category"] == "Consultation"] if not class_loads.empty else pd.DataFrame()
    
    current_row = 3
    
    def write_load_section(title, df, banner_color, subtotal_label):
        nonlocal current_row
        # Section Header Banner
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=10)
        b_cell = ws.cell(row=current_row, column=1, value=title)
        b_cell.font = Font(name="Calibri", size=11, bold=True, color=TEXT_LIGHT)
        b_cell.fill = PatternFill(start_color=banner_color, end_color=banner_color, fill_type="solid")
        b_cell.alignment = Alignment(horizontal="left", vertical="center")
        for col in range(1, 11):
            ws.cell(row=current_row, column=col).border = thin_border
            ws.cell(row=current_row, column=col).fill = PatternFill(start_color=banner_color, end_color=banner_color, fill_type="solid")
        ws.row_dimensions[current_row].height = 24
        current_row += 1
        
        # Table Column Headers
        ws.row_dimensions[current_row].height = 22
        for c_idx, h in enumerate(headers, start=1):
            c_cell = ws.cell(row=current_row, column=c_idx, value=h)
            c_cell.font = Font(bold=True, size=10, color=TEXT_LIGHT)
            c_cell.fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
            c_cell.alignment = Alignment(horizontal="center", vertical="center")
            c_cell.border = thin_border
        current_row += 1
        
        start_data_row = current_row
        if df.empty:
            ws.row_dimensions[current_row].height = 20
            ws.cell(row=current_row, column=1, value="None scheduled").border = thin_border
            for col in range(2, 8):
                ws.cell(row=current_row, column=col, value="").border = thin_border
            ws.cell(row=current_row, column=8, value=0.0).border = thin_border
            ws.cell(row=current_row, column=9, value=0.0).border = thin_border
            ws.cell(row=current_row, column=10, value="").border = thin_border
            current_row += 1
            end_data_row = start_data_row
        else:
            for _, row in df.iterrows():
                ws.row_dimensions[current_row].height = 20
                ws.cell(row=current_row, column=1, value=row["subject_name"]).border = thin_border
                ws.cell(row=current_row, column=2, value=row["section_code"]).border = thin_border
                ws.cell(row=current_row, column=3, value=row["room"]).border = thin_border
                
                day_cell = ws.cell(row=current_row, column=4, value=row["day_of_week"])
                day_cell.alignment = Alignment(horizontal="center")
                day_cell.border = thin_border
                
                start_cell = ws.cell(row=current_row, column=5, value=row["start_time"])
                start_cell.alignment = Alignment(horizontal="center")
                start_cell.border = thin_border
                
                end_cell = ws.cell(row=current_row, column=6, value=row["end_time"])
                end_cell.alignment = Alignment(horizontal="center")
                end_cell.border = thin_border
                
                type_cell = ws.cell(row=current_row, column=7, value=row["type"])
                type_cell.alignment = Alignment(horizontal="center")
                type_cell.border = thin_border
                
                units_cell = ws.cell(row=current_row, column=8, value=row["units"])
                units_cell.alignment = Alignment(horizontal="right")
                units_cell.border = thin_border
                
                hours_cell = ws.cell(row=current_row, column=9, value=row["hours"])
                hours_cell.alignment = Alignment(horizontal="right")
                hours_cell.border = thin_border
                
                cat_cell = ws.cell(row=current_row, column=10, value=row["load_category"])
                cat_cell.border = thin_border
                current_row += 1
            end_data_row = current_row - 1
            
        # Subtotal Row
        subtotal_row = current_row
        ws.row_dimensions[subtotal_row].height = 22
        ws.merge_cells(start_row=subtotal_row, start_column=1, end_row=subtotal_row, end_column=7)
        tot_label = ws.cell(row=subtotal_row, column=1, value=subtotal_label)
        tot_label.font = Font(bold=True)
        tot_label.alignment = Alignment(horizontal="right", vertical="center")
        for col in range(1, 8):
            ws.cell(row=subtotal_row, column=col).border = thin_border
            ws.cell(row=subtotal_row, column=col).fill = PatternFill(start_color=SECONDARY_COLOR, end_color=SECONDARY_COLOR, fill_type="solid")
            
        cell_u = ws.cell(row=subtotal_row, column=8, value=f"=SUM(H{start_data_row}:H{end_data_row})")
        cell_u.font = Font(bold=True)
        cell_u.border = thin_border
        cell_u.alignment = Alignment(horizontal="right")
        cell_u.fill = PatternFill(start_color=SECONDARY_COLOR, end_color=SECONDARY_COLOR, fill_type="solid")
        
        cell_h = ws.cell(row=subtotal_row, column=9, value=f"=SUM(I{start_data_row}:I{end_data_row})")
        cell_h.font = Font(bold=True)
        cell_h.border = thin_border
        cell_h.alignment = Alignment(horizontal="right")
        cell_h.fill = PatternFill(start_color=SECONDARY_COLOR, end_color=SECONDARY_COLOR, fill_type="solid")
        
        ws.cell(row=subtotal_row, column=10).border = thin_border
        ws.cell(row=subtotal_row, column=10).fill = PatternFill(start_color=SECONDARY_COLOR, end_color=SECONDARY_COLOR, fill_type="solid")
        
        current_row += 2 # Leave a spacer row
        return subtotal_row
        
    sub_r1 = write_load_section("📘 1. REGULAR LOAD SCHEDULE", regular_classes, PRIMARY_COLOR, "Regular Load Subtotal")
    sub_r2 = write_load_section("⚡ 2. EXCESS LOAD SCHEDULE", excess_classes, EXCESS_HEADER_COLOR, "Excess Load Subtotal")
    sub_r3 = write_load_section("🏢 3. CONSULTATION / NON-TEACHING SCHEDULE", consultations, NON_TEACH_COLOR, "Consultation Subtotal")
    
    # Combined Totals Row
    ws.row_dimensions[current_row].height = 24
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
    tot_label = ws.cell(row=current_row, column=1, value="TOTAL LOAD (ALL CATEGORIES)")
    tot_label.font = Font(bold=True, size=11)
    tot_label.alignment = Alignment(horizontal="right", vertical="center")
    for col in range(1, 8):
        ws.cell(row=current_row, column=col).border = thin_border
        ws.cell(row=current_row, column=col).fill = PatternFill(start_color=GRAND_TOTAL_COLOR, end_color=GRAND_TOTAL_COLOR, fill_type="solid")
        
    tot_u = ws.cell(row=current_row, column=8, value=f"=H{sub_r1}+H{sub_r2}+H{sub_r3}")
    tot_u.font = Font(bold=True, size=11)
    tot_u.border = thin_border
    tot_u.alignment = Alignment(horizontal="right")
    tot_u.fill = PatternFill(start_color=GRAND_TOTAL_COLOR, end_color=GRAND_TOTAL_COLOR, fill_type="solid")
    
    tot_h = ws.cell(row=current_row, column=9, value=f"=I{sub_r1}+I{sub_r2}+I{sub_r3}")
    tot_h.font = Font(bold=True, size=11)
    tot_h.border = thin_border
    tot_h.alignment = Alignment(horizontal="right")
    tot_h.fill = PatternFill(start_color=GRAND_TOTAL_COLOR, end_color=GRAND_TOTAL_COLOR, fill_type="solid")
    
    ws.cell(row=current_row, column=10).border = thin_border
    ws.cell(row=current_row, column=10).fill = PatternFill(start_color=GRAND_TOTAL_COLOR, end_color=GRAND_TOTAL_COLOR, fill_type="solid")
    
    # Column width formatting
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 8
    ws.column_dimensions["E"].width = 12
    ws.column_dimensions["F"].width = 12
    ws.column_dimensions["G"].width = 8
    ws.column_dimensions["H"].width = 10
    ws.column_dimensions["I"].width = 10
    ws.column_dimensions["J"].width = 18

def create_timesheet_sheet(wb, sheet_title, instructor, class_loads, year, month, cutoff_type, timesheet_data):
    ws = wb.create_sheet(title=sheet_title)
    ws.views.sheetView[0].showGridLines = True
    thin_border = get_thin_border()
    
    # Generate dates for this cutoff
    dates = generate_cutoff_dates(year, month, cutoff_type)
    num_dates = len(dates)
    last_col_idx = 2 + num_dates
    last_col_letter = get_column_letter(last_col_idx)
    total_col_idx = last_col_idx + 1
    total_col_letter = get_column_letter(total_col_idx)
    
    # Title Block
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_col_idx)
    ws["A1"] = f"ACADEMIC INSTRUCTOR TIMESHEET ({cutoff_type})"
    ws["A1"].font = Font(name="Calibri", size=16, bold=True, color=TEXT_LIGHT)
    ws["A1"].fill = PatternFill(start_color=PRIMARY_COLOR, end_color=PRIMARY_COLOR, fill_type="solid")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36
    
    # Header metadata
    month_name = calendar.month_name[month]
    acad_level = instructor.get("academic_level", "Tertiary")
    max_reg_units = 30.0 if acad_level == "SHS" else 24.0
    headers_meta = [
        ("Instructor Name:", instructor["name"], "Employment Status:", f"{instructor['employment_status']} ({acad_level})"),
        ("Cutoff Period:", f"{month_name} {cutoff_type}, {year}", "Department:", instructor.get("department", "N/A"))
    ]
    for r_idx, row_data in enumerate(headers_meta, start=2):
        ws.row_dimensions[r_idx].height = 20
        ws.cell(row=r_idx, column=1, value=row_data[0]).font = Font(bold=True)
        ws.cell(row=r_idx, column=2, value=row_data[1])
        ws.cell(row=r_idx, column=5, value=row_data[2]).font = Font(bold=True)
        ws.cell(row=r_idx, column=6, value=row_data[3])
        
    regular_classes = class_loads[class_loads["load_category"] == "Regular Load"] if not class_loads.empty else pd.DataFrame()
    excess_classes = class_loads[class_loads["load_category"].isin(["Excess Load", "Excess/Overload"])] if not class_loads.empty else pd.DataFrame()
    consultations = class_loads[class_loads["load_category"] == "Consultation"] if not class_loads.empty else pd.DataFrame()
    
    current_row = 5
    
    def write_timesheet_section(banner_title, banner_color, subtotal_title, subtotal_color, activity_items, default_cat):
        nonlocal current_row
        
        # 1. Section Banner
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=total_col_idx)
        b_cell = ws.cell(row=current_row, column=1, value=banner_title)
        b_cell.font = Font(name="Calibri", size=11, bold=True, color=TEXT_LIGHT)
        b_cell.alignment = Alignment(horizontal="left", vertical="center")
        for col in range(1, total_col_idx + 1):
            ws.cell(row=current_row, column=col).border = thin_border
            ws.cell(row=current_row, column=col).fill = PatternFill(start_color=banner_color, end_color=banner_color, fill_type="solid")
        ws.row_dimensions[current_row].height = 24
        current_row += 1
        
        # 2. Table Column Headers
        h1_row = current_row
        h2_row = current_row + 1
        ws.row_dimensions[h1_row].height = 18
        ws.row_dimensions[h2_row].height = 18
        
        # Col 1: Category
        ws.merge_cells(start_row=h1_row, start_column=1, end_row=h2_row, end_column=1)
        ws.cell(row=h1_row, column=1, value="Category").font = Font(bold=True, size=9, color=TEXT_LIGHT)
        ws.cell(row=h1_row, column=1).alignment = Alignment(horizontal="center", vertical="center")
        
        # Col 2: Details
        ws.merge_cells(start_row=h1_row, start_column=2, end_row=h2_row, end_column=2)
        ws.cell(row=h1_row, column=2, value="Activity / Course Details").font = Font(bold=True, size=9, color=TEXT_LIGHT)
        ws.cell(row=h1_row, column=2).alignment = Alignment(horizontal="center", vertical="center")
        
        for c in (1, 2):
            for r in (h1_row, h2_row):
                ws.cell(row=r, column=c).border = thin_border
                ws.cell(row=r, column=c).fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
                
        # Date Columns
        for d_idx, d_info in enumerate(dates):
            c_idx = 3 + d_idx
            
            # Row H1: Day number
            d_cell = ws.cell(row=h1_row, column=c_idx, value=d_info['day_num'])
            d_cell.font = Font(bold=True, size=9, color=TEXT_LIGHT)
            d_cell.alignment = Alignment(horizontal="center", vertical="center")
            d_cell.border = thin_border
            d_cell.fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
            
            # Row H2: Day Abbr
            abbr_cell = ws.cell(row=h2_row, column=c_idx, value=d_info['day_abbr'])
            abbr_cell.font = Font(bold=True, size=8.5, color=TEXT_LIGHT)
            abbr_cell.alignment = Alignment(horizontal="center", vertical="center")
            abbr_cell.border = thin_border
            abbr_cell.fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
            
            # Weekend header highlight
            if d_info['day_abbr'] in ("S", "SU"):
                d_cell.fill = PatternFill(start_color="475569", end_color="475569", fill_type="solid")
                abbr_cell.fill = PatternFill(start_color="475569", end_color="475569", fill_type="solid")
                
        # Total Column Header
        ws.merge_cells(start_row=h1_row, start_column=total_col_idx, end_row=h2_row, end_column=total_col_idx)
        tot_header = ws.cell(row=h1_row, column=total_col_idx, value="Total Hours")
        tot_header.font = Font(bold=True, size=9, color=TEXT_LIGHT)
        tot_header.alignment = Alignment(horizontal="center", vertical="center")
        for r in (h1_row, h2_row):
            ws.cell(row=r, column=total_col_idx).border = thin_border
            ws.cell(row=r, column=total_col_idx).fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
            
        current_row += 2
        
        # 3. Data Rows
        start_data_row = current_row
        if not activity_items:
            ws.row_dimensions[current_row].height = 20
            c_cell = ws.cell(row=current_row, column=1, value=default_cat)
            c_cell.border = thin_border
            d_cell = ws.cell(row=current_row, column=2, value=f"No {default_cat} activities scheduled")
            d_cell.border = thin_border
            d_cell.font = Font(italic=True, color="64748B")
            
            for d_idx, d_info in enumerate(dates):
                col_idx = 3 + d_idx
                h_cell = ws.cell(row=current_row, column=col_idx, value=0.0)
                h_cell.alignment = Alignment(horizontal="right")
                h_cell.border = thin_border
                if d_info['day_abbr'] in ("S", "SU"):
                    h_cell.fill = PatternFill(start_color=WEEKEND_COLOR, end_color=WEEKEND_COLOR, fill_type="solid")
                    
            t_cell = ws.cell(row=current_row, column=total_col_idx, value=f"=SUM(C{current_row}:{last_col_letter}{current_row})")
            t_cell.font = Font(bold=True)
            t_cell.alignment = Alignment(horizontal="right")
            t_cell.border = thin_border
            t_cell.fill = PatternFill(start_color=SECONDARY_COLOR, end_color=SECONDARY_COLOR, fill_type="solid")
            current_row += 1
            end_data_row = start_data_row
        else:
            for cat, detail, key_name in activity_items:
                r_num = current_row
                ws.row_dimensions[r_num].height = 20
                
                c_cell = ws.cell(row=r_num, column=1, value=cat)
                c_cell.border = thin_border
                
                d_cell = ws.cell(row=r_num, column=2, value=detail)
                d_cell.border = thin_border
                
                # Write hours for each date
                for d_idx, d_info in enumerate(dates):
                    col_idx = 3 + d_idx
                    date_str = d_info['date'].strftime("%Y-%m-%d")
                    
                    val = timesheet_data.get((key_name, date_str))
                    if val is None:
                        if cat == "Regular Load":
                            val = 0.0
                            for _, cl in regular_classes.iterrows():
                                cl_key = f"{cl['type']}: {cl['subject_name']} ({cl['section_code']})"
                                if cl_key == key_name and cl['day_of_week'] == d_info['day_abbr']:
                                    val = cl['hours']
                                    break
                        elif cat == "Excess Load":
                            val = 0.0
                            for _, cl in excess_classes.iterrows():
                                cl_key = f"{cl['type']}: {cl['subject_name']} ({cl['section_code']})"
                                if cl_key == key_name and cl['day_of_week'] == d_info['day_abbr']:
                                    val = cl['hours']
                                    break
                        elif key_name == "Consultation":
                            val = 0.0
                            for _, cl in consultations.iterrows():
                                if cl['day_of_week'] == d_info['day_abbr']:
                                    val = cl['hours']
                                    break
                        else:
                            val = 0.0
                            
                    h_cell = ws.cell(row=r_num, column=col_idx, value=val)
                    h_cell.alignment = Alignment(horizontal="right")
                    h_cell.border = thin_border
                    if d_info['day_abbr'] in ("S", "SU"):
                        h_cell.fill = PatternFill(start_color=WEEKEND_COLOR, end_color=WEEKEND_COLOR, fill_type="solid")
                        
                # Total Formula for the row
                tot_formula = f"=SUM(C{r_num}:{last_col_letter}{r_num})"
                t_cell = ws.cell(row=r_num, column=total_col_idx, value=tot_formula)
                t_cell.font = Font(bold=True)
                t_cell.alignment = Alignment(horizontal="right")
                t_cell.border = thin_border
                t_cell.fill = PatternFill(start_color=SECONDARY_COLOR, end_color=SECONDARY_COLOR, fill_type="solid")
                
                current_row += 1
            end_data_row = current_row - 1
            
        # 4. Subtotal Row
        subtotal_row = current_row
        ws.row_dimensions[subtotal_row].height = 22
        ws.merge_cells(start_row=subtotal_row, start_column=1, end_row=subtotal_row, end_column=2)
        sub_label = ws.cell(row=subtotal_row, column=1, value=subtotal_title)
        sub_label.font = Font(bold=True)
        sub_label.alignment = Alignment(horizontal="right", vertical="center")
        
        for col in range(1, 3):
            ws.cell(row=subtotal_row, column=col).border = thin_border
            ws.cell(row=subtotal_row, column=col).fill = PatternFill(start_color=subtotal_color, end_color=subtotal_color, fill_type="solid")
            
        # Daily subtotal formulas
        for d_idx, d_info in enumerate(dates):
            col_idx = 3 + d_idx
            col_let = get_column_letter(col_idx)
            sub_formula = f"=SUM({col_let}{start_data_row}:{col_let}{end_data_row})"
            
            s_cell = ws.cell(row=subtotal_row, column=col_idx, value=sub_formula)
            s_cell.font = Font(bold=True)
            s_cell.alignment = Alignment(horizontal="right")
            s_cell.border = thin_border
            s_cell.fill = PatternFill(start_color=subtotal_color, end_color=subtotal_color, fill_type="solid")
            
        # Subtotal total formula
        sub_tot_formula = f"=SUM({total_col_letter}{start_data_row}:{total_col_letter}{end_data_row})"
        st_cell = ws.cell(row=subtotal_row, column=total_col_idx, value=sub_tot_formula)
        st_cell.font = Font(bold=True)
        st_cell.alignment = Alignment(horizontal="right")
        st_cell.border = thin_border
        st_cell.fill = PatternFill(start_color=subtotal_color, end_color=subtotal_color, fill_type="solid")
        
        current_row += 2 # leave blank row
        return subtotal_row
        
    # Build list of items for each section
    reg_items = []
    for _, cl in regular_classes.iterrows():
        key_name = f"{cl['type']}: {cl['subject_name']} ({cl['section_code']})"
        reg_items.append(("Regular Load", key_name, key_name))
        
    exc_items = []
    for _, cl in excess_classes.iterrows():
        key_name = f"{cl['type']}: {cl['subject_name']} ({cl['section_code']})"
        exc_items.append(("Excess Load", key_name, key_name))
        
    non_items = [
        ("Non-Teaching SC", "Career Orientation Seminars (COS)", "Career Orientation Seminars (COS)"),
        ("Non-Teaching SC", "Guidance/Counseling", "Guidance/Counseling"),
        ("Non-Teaching HQ", "Consultation", "Consultation"),
        ("Non-Teaching HQ", "Administrative Hours", "Administrative Hours")
    ]
    
    # 1. Regular Load Section
    sub_r1 = write_timesheet_section(
        f"📘 1. REGULAR LOAD TIMESHEET (Regular Cap: {max_reg_units:.0f} Units - {acad_level})",
        PRIMARY_COLOR,
        "Regular Load Subtotal",
        SECONDARY_COLOR,
        reg_items,
        "Regular Load"
    )
    
    # 2. Excess Load Section
    sub_r2 = write_timesheet_section(
        "⚡ 2. EXCESS LOAD TIMESHEET (Overload / Additional Teaching)",
        EXCESS_HEADER_COLOR,
        "Excess Load Subtotal",
        EXCESS_SUB_COLOR,
        exc_items,
        "Excess Load"
    )
    
    # 3. Non-Teaching Section
    sub_r3 = write_timesheet_section(
        "🏢 3. NON-TEACHING ACTIVITIES (SC & HQ)",
        NON_TEACH_COLOR,
        "Non-Teaching Subtotal",
        SECONDARY_COLOR,
        non_items,
        "Non-Teaching"
    )
    
    # 4. Consolidated Daily & Grand Totals Section
    cons_banner_row = current_row
    ws.merge_cells(start_row=cons_banner_row, start_column=1, end_row=cons_banner_row, end_column=total_col_idx)
    cb_cell = ws.cell(row=cons_banner_row, column=1, value="CONSOLIDATED DAILY & GRAND TOTALS (ALL ACTIVITIES)")
    cb_cell.font = Font(name="Calibri", size=11, bold=True, color=TEXT_LIGHT)
    cb_cell.alignment = Alignment(horizontal="left", vertical="center")
    for col in range(1, total_col_idx + 1):
        ws.cell(row=cons_banner_row, column=col).border = thin_border
        ws.cell(row=cons_banner_row, column=col).fill = PatternFill(start_color=PRIMARY_COLOR, end_color=PRIMARY_COLOR, fill_type="solid")
    ws.row_dimensions[cons_banner_row].height = 24
    current_row += 1
    
    tot_row_idx = current_row
    ws.row_dimensions[tot_row_idx].height = 24
    
    ws.merge_cells(start_row=tot_row_idx, start_column=1, end_row=tot_row_idx, end_column=2)
    tot_label = ws.cell(row=tot_row_idx, column=1, value="DAILY TOTALS (ALL ACTIVITIES)")
    tot_label.font = Font(bold=True, size=10)
    tot_label.alignment = Alignment(horizontal="right", vertical="center")
    for col in range(1, 3):
        ws.cell(row=tot_row_idx, column=col).border = thin_border
        ws.cell(row=tot_row_idx, column=col).fill = PatternFill(start_color=SECONDARY_COLOR, end_color=SECONDARY_COLOR, fill_type="solid")
        
    for d_idx, d_info in enumerate(dates):
        c_idx = 3 + d_idx
        col_let = get_column_letter(c_idx)
        daily_formula = f"={col_let}{sub_r1}+{col_let}{sub_r2}+{col_let}{sub_r3}"
        
        d_tot_cell = ws.cell(row=tot_row_idx, column=c_idx, value=daily_formula)
        d_tot_cell.font = Font(bold=True)
        d_tot_cell.alignment = Alignment(horizontal="right")
        d_tot_cell.border = thin_border
        d_tot_cell.fill = PatternFill(start_color=SECONDARY_COLOR, end_color=SECONDARY_COLOR, fill_type="solid")
        
    grand_formula = f"={total_col_letter}{sub_r1}+{total_col_letter}{sub_r2}+{total_col_letter}{sub_r3}"
    g_cell = ws.cell(row=tot_row_idx, column=total_col_idx, value=grand_formula)
    g_cell.font = Font(bold=True, size=11, color="000000")
    g_cell.alignment = Alignment(horizontal="right")
    g_cell.border = thin_border
    g_cell.fill = PatternFill(start_color=GRAND_TOTAL_COLOR, end_color=GRAND_TOTAL_COLOR, fill_type="solid")
    
    # Column width formatting
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 30
    for col in range(3, last_col_idx + 1):
        ws.column_dimensions[get_column_letter(col)].width = 6
    ws.column_dimensions[total_col_letter].width = 12
    
    return ws
