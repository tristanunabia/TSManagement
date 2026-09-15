from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import calendar
import datetime
from date_engine import generate_cutoff_dates

def generate_pdf_timesheet(instructor_row, class_loads_df, year, month, selected_cutoff, timesheet_data):
    """
    Generates a PDF timesheet using ReportLab and returns a byte stream.
    """
    import io
    pdf_buffer = io.BytesIO()
    
    # Page setup
    # Letter size: 8.5 x 11 inches. Landscape: 11 x 8.5 inches = 792 x 612 points.
    # Set margins to 0.25 inches (18 points) to maximize printable space.
    doc = SimpleDocTemplate(
        pdf_buffer,
        pagesize=landscape(letter),
        leftMargin=18,
        rightMargin=18,
        topMargin=25,
        bottomMargin=18
    )
    
    story = []
    styles = getSampleStyleSheet()
    
    # Custom styles
    title_style = ParagraphStyle(
        'DocTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=16,
        leading=20,
        textColor=colors.HexColor("#1F4E78"),
        alignment=1, # Center
        spaceAfter=15
    )
    
    meta_label_style = ParagraphStyle(
        'MetaLabel',
        fontName='Helvetica-Bold',
        fontSize=9,
        leading=11,
        textColor=colors.HexColor("#333333")
    )
    
    meta_value_style = ParagraphStyle(
        'MetaVal',
        fontName='Helvetica',
        fontSize=9,
        leading=11,
        textColor=colors.HexColor("#000000")
    )
    
    # 1. Document Title
    story.append(Paragraph(f"ACADEMIC INSTRUCTOR TIMESHEET ({selected_cutoff})", title_style))
    
    # 2. Metadata Block (Instructor details)
    month_name = calendar.month_name[month]
    acad_level = instructor_row["academic_level"] if "academic_level" in instructor_row.keys() else "Tertiary"
    meta_data = [
        [
            Paragraph("Instructor Name:", meta_label_style), Paragraph(instructor_row["name"], meta_value_style),
            Paragraph("Employment Status:", meta_label_style), Paragraph(f"{instructor_row['employment_status']} ({acad_level})", meta_value_style)
        ],
        [
            Paragraph("Cutoff Period:", meta_label_style), Paragraph(f"{month_name} {selected_cutoff}, {year}", meta_value_style),
            Paragraph("Department:", meta_label_style), Paragraph(instructor_row.get("department", "N/A"), meta_value_style)
        ]
    ]
    
    # Page width: 792 - 36 = 756 points
    meta_table = Table(meta_data, colWidths=[110, 240, 120, 286])
    meta_table.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('TOPPADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(meta_table)
    story.append(Spacer(1, 15))
    
    # 3. Build the Grid Data
    dates = generate_cutoff_dates(year, month, selected_cutoff)
    num_dates = len(dates)
    
    # Headers
    # Row 0: Date numbers
    # Row 1: Weekday abbreviations
    row_dates = ["Activity Category", "Activity Details"] + [str(d['day_num']) for d in dates] + ["Total"]
    row_weekdays = ["", ""] + [d['day_abbr'] for d in dates] + [""]
    
    table_data = [row_dates, row_weekdays]
    
    # Prepare rows: Separate Regular Load and Excess Load
    regular_classes = class_loads_df[class_loads_df["load_category"] == "Regular Load"]
    excess_classes = class_loads_df[class_loads_df["load_category"].isin(["Excess Load", "Excess/Overload"])]
    activity_rows = []
    
    # Section A: Regular Load
    for _, cl in regular_classes.iterrows():
        key_name = f"{cl['type']}: {cl['subject_name']} ({cl['section_code']})"
        activity_rows.append(("Regular Load", key_name, key_name))
        
    # Section B: Excess Load
    for _, cl in excess_classes.iterrows():
        key_name = f"{cl['type']}: {cl['subject_name']} ({cl['section_code']})"
        activity_rows.append(("Excess Load", key_name, key_name))
        
    # Section C
    activity_rows.append(("Non-Teaching SC", "Career Orientation Seminars (COS)", "Career Orientation Seminars (COS)"))
    activity_rows.append(("Non-Teaching SC", "Guidance/Counseling", "Guidance/Counseling"))
    
    # Section D
    activity_rows.append(("Non-Teaching HQ", "Consultation", "Consultation"))
    activity_rows.append(("Non-Teaching HQ", "Administrative Hours", "Administrative Hours"))
    
    # Cell formatting styles
    cell_hdr_style = ParagraphStyle('HdrCell', fontName='Helvetica-Bold', fontSize=8, leading=10, textColor=colors.white, alignment=1)
    cell_hdr_lbl_style = ParagraphStyle('HdrLblCell', fontName='Helvetica-Bold', fontSize=8, leading=10, textColor=colors.white, alignment=0)
    cell_lbl_cat_style = ParagraphStyle('LblCat', fontName='Helvetica-Bold', fontSize=8, leading=9, textColor=colors.HexColor("#333333"))
    cell_lbl_det_style = ParagraphStyle('LblDet', fontName='Helvetica', fontSize=7.5, leading=8.5, textColor=colors.HexColor("#111111"))
    cell_num_style = ParagraphStyle('NumVal', fontName='Helvetica', fontSize=8, leading=9, alignment=1)
    cell_num_tot_style = ParagraphStyle('NumTotVal', fontName='Helvetica-Bold', fontSize=8, leading=9, alignment=1)
    
    # Convert header rows to Paragraph flowables (or standard text if styling in TableStyle)
    # Actually, ReportLab Table works well with Paragraphs to wrap text. Let's wrap headers in Paragraphs.
    table_rows = []
    table_rows.append([
        Paragraph(row_dates[0], cell_hdr_lbl_style), Paragraph(row_dates[1], cell_hdr_lbl_style)
    ] + [Paragraph(x, cell_hdr_style) for x in row_dates[2:-1]] + [Paragraph(row_dates[-1], cell_hdr_style)])
    
    table_rows.append([
        Paragraph("", cell_hdr_lbl_style), Paragraph("", cell_hdr_lbl_style)
    ] + [Paragraph(x, cell_hdr_style) for x in row_weekdays[2:-1]] + [Paragraph("", cell_hdr_style)])
    
    # Compute widths: total width = 756 points
    # Category = 90pt, Details = 140pt, Total = 56pt
    # Remaining = 756 - (90 + 140 + 56) = 470pt
    # Per date column = 470 / num_dates (approx 29 - 31pt)
    cat_width = 85
    det_width = 135
    tot_width = 46
    date_width = (756.0 - (cat_width + det_width + tot_width)) / num_dates
    col_widths = [cat_width, det_width] + [date_width] * num_dates + [tot_width]
    
    # Track daily column totals
    daily_totals = [0.0] * num_dates
    
    # Add data rows
    table_style_cmds = [
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#D9D9D9")),
        ('SPAN', (0,0), (0,1)), # Merge category header
        ('SPAN', (1,0), (1,1)), # Merge details header
        ('BACKGROUND', (0,0), (-1,1), colors.HexColor("#1F4E78")), # Header Background
        ('BOTTOMPADDING', (0,0), (-1,1), 4),
        ('TOPPADDING', (0,0), (-1,1), 4),
    ]
    
    for idx, (cat, detail, key_name) in enumerate(activity_rows):
        r_num = 2 + idx  # 2 header rows
        row_cells = [
            Paragraph(cat, cell_lbl_cat_style),
            Paragraph(detail, cell_lbl_det_style)
        ]
        
        row_sum = 0.0
        # For each date
        for d_idx, d_info in enumerate(dates):
            date_str = d_info['date'].strftime("%Y-%m-%d")
            
            # Get hours
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
                elif cat == "Teaching Activities":
                    val = 0.0
                    for _, cl in pd.concat([regular_classes, excess_classes]).iterrows():
                        cl_key = f"{cl['type']}: {cl['subject_name']} ({cl['section_code']})"
                        if cl_key == key_name and cl['day_of_week'] == d_info['day_abbr']:
                            val = cl['hours']
                            break
                elif key_name == "Consultation":
                    val = 0.0
                    consultations = class_loads_df[class_loads_df["load_category"] == "Consultation"]
                    for _, cl in consultations.iterrows():
                        if cl['day_of_week'] == d_info['day_abbr']:
                            val = cl['hours']
                            break
                else:
                    val = 0.0
            
            row_sum += val
            daily_totals[d_idx] += val
            
            val_str = f"{val:.1f}" if val > 0 else "-"
            row_cells.append(Paragraph(val_str, cell_num_style))
            
            # Weekend Column highlight
            if d_info['day_abbr'] in ("S", "SU"):
                table_style_cmds.append(('BACKGROUND', (2 + d_idx, r_num), (2 + d_idx, r_num), colors.HexColor("#F9F9F9")))
                
        # Total
        row_cells.append(Paragraph(f"{row_sum:.1f}", cell_num_tot_style))
        table_rows.append(row_cells)
        
        # Row styling
        # Highlight total cell of the row
        table_style_cmds.append(('BACKGROUND', (total_col_idx := (2 + num_dates), r_num), (total_col_idx, r_num), colors.HexColor("#D9E1F2")))
        # Alternating background colors
        if idx % 2 == 1:
            table_style_cmds.append(('BACKGROUND', (0, r_num), (total_col_idx - 1, r_num), colors.HexColor("#F5F7FA")))
            
    # Add Daily Totals Row
    r_total = 2 + len(activity_rows)
    tot_row_cells = [
        Paragraph("Daily Totals", cell_lbl_cat_style),
        Paragraph("", cell_lbl_det_style)
    ]
    for tot_val in daily_totals:
        tot_str = f"{tot_val:.1f}" if tot_val > 0 else "-"
        tot_row_cells.append(Paragraph(tot_str, cell_num_tot_style))
    
    grand_total = sum(daily_totals)
    tot_row_cells.append(Paragraph(f"{grand_total:.1f}", cell_num_tot_style))
    table_rows.append(tot_row_cells)
    
    # Merge label cell for daily totals
    table_style_cmds.extend([
        ('SPAN', (0, r_total), (1, r_total)),
        ('BACKGROUND', (0, r_total), (-1, r_total), colors.HexColor("#D9E1F2")),
        # Green shade for grand total cell
        ('BACKGROUND', (-1, r_total), (-1, r_total), colors.HexColor("#C6EFCE")),
        ('BOTTOMPADDING', (0, 2), (-1, -1), 3),
        ('TOPPADDING', (0, 2), (-1, -1), 3),
    ])
    
    # Style weekends in the Daily Totals row too
    for d_idx, d_info in enumerate(dates):
        if d_info['day_abbr'] in ("S", "SU"):
            table_style_cmds.append(('BACKGROUND', (2 + d_idx, r_total), (2 + d_idx, r_total), colors.HexColor("#ECECEC")))
            
    # Add table to PDF story
    t = Table(table_rows, colWidths=col_widths)
    t.setStyle(TableStyle(table_style_cmds))
    story.append(t)
    
    story.append(Spacer(1, 20))
    
    # 4. Verification Check and Signatures Section
    sig_label_style = ParagraphStyle('SigLbl', fontName='Helvetica-Bold', fontSize=9, alignment=0)
    sig_line_style = ParagraphStyle('SigLine', fontName='Helvetica', fontSize=8, alignment=0)
    
    # Progress gauge against 80 hour standard target
    target_status = "PASSED" if grand_total >= 80.0 else "UNDER-HOURS"
    status_color = "#2E7D32" if grand_total >= 80.0 else "#C62828"
    
    verif_text = f"Grand Total Hours: <b>{grand_total:.2f} hrs</b> / Target Period Standard: <b>80.0 hrs</b> | Validation Status: <font color='{status_color}'><b>{target_status}</b></font>"
    story.append(Paragraph(verif_text, ParagraphStyle('Verif', fontName='Helvetica', fontSize=9, leading=12)))
    story.append(Spacer(1, 25))
    
    # Signatures layout
    sig_data = [
        [
            Paragraph("________________________________________", sig_label_style),
            Paragraph("", sig_label_style),
            Paragraph("________________________________________", sig_label_style)
        ],
        [
            Paragraph("Faculty Instructor's Signature", sig_line_style),
            Paragraph("", sig_line_style),
            Paragraph("Approved by: Dean / Academic Supervisor", sig_line_style)
        ],
        [
            Paragraph("Date signed: ___________________________", sig_line_style),
            Paragraph("", sig_line_style),
            Paragraph("Date approved: _________________________", sig_line_style)
        ]
    ]
    
    sig_table = Table(sig_data, colWidths=[320, 116, 320])
    sig_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
    ]))
    story.append(sig_table)
    
    # Build Document
    doc.build(story)
    pdf_buffer.seek(0)
    return pdf_buffer.getvalue()
