from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import calendar
import datetime
import pandas as pd
from date_engine import generate_cutoff_dates

def generate_pdf_timesheet(instructor_row, class_loads_df, year, month, selected_cutoff, timesheet_data):
    """
    Generates a PDF timesheet using ReportLab and returns a byte stream.
    Separates tables for Regular Load, Excess Load, Non-teaching SC/HQ, and Consolidated Totals.
    """
    import io
    pdf_buffer = io.BytesIO()
    
    # Page setup
    # Letter size landscape: 11 x 8.5 inches = 792 x 612 points.
    # Margins: 18pt left/right/bottom, 20pt top. Printable width: 756 points.
    doc = SimpleDocTemplate(
        pdf_buffer,
        pagesize=landscape(letter),
        leftMargin=18,
        rightMargin=18,
        topMargin=20,
        bottomMargin=18
    )
    
    story = []
    styles = getSampleStyleSheet()
    
    # Typography Styles
    title_style = ParagraphStyle(
        'DocTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=14,
        leading=17,
        textColor=colors.HexColor("#1F4E78"),
        alignment=1, # Center
        spaceAfter=8
    )
    
    meta_label_style = ParagraphStyle(
        'MetaLabel',
        fontName='Helvetica-Bold',
        fontSize=8.5,
        leading=10.5,
        textColor=colors.HexColor("#333333")
    )
    
    meta_value_style = ParagraphStyle(
        'MetaVal',
        fontName='Helvetica',
        fontSize=8.5,
        leading=10.5,
        textColor=colors.HexColor("#000000")
    )
    
    # 1. Document Title
    story.append(Paragraph(f"ACADEMIC INSTRUCTOR TIMESHEET ({selected_cutoff})", title_style))
    
    # 2. Metadata Block (Instructor details)
    month_name = calendar.month_name[month]
    acad_level = instructor_row["academic_level"] if "academic_level" in instructor_row.keys() else "Tertiary"
    max_reg_units = 30.0 if acad_level == "SHS" else 24.0
    
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
    
    meta_table = Table(meta_data, colWidths=[110, 240, 120, 286])
    meta_table.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
        ('TOPPADDING', (0,0), (-1,-1), 2),
    ]))
    story.append(meta_table)
    story.append(Spacer(1, 8))
    
    # 3. Column Widths and Common Styles
    dates = generate_cutoff_dates(year, month, selected_cutoff)
    num_dates = len(dates)
    
    cat_width = 85
    det_width = 135
    tot_width = 46
    date_width = (756.0 - (cat_width + det_width + tot_width)) / num_dates
    col_widths = [cat_width, det_width] + [date_width] * num_dates + [tot_width]
    
    cell_banner_style = ParagraphStyle('BannerCell', fontName='Helvetica-Bold', fontSize=8.5, leading=10, textColor=colors.white, alignment=0)
    cell_hdr_style = ParagraphStyle('HdrCell', fontName='Helvetica-Bold', fontSize=7.5, leading=9, textColor=colors.white, alignment=1)
    cell_hdr_lbl_style = ParagraphStyle('HdrLblCell', fontName='Helvetica-Bold', fontSize=7.5, leading=9, textColor=colors.white, alignment=1)
    cell_lbl_cat_style = ParagraphStyle('LblCat', fontName='Helvetica-Bold', fontSize=7.5, leading=8.5, textColor=colors.HexColor("#333333"))
    cell_lbl_det_style = ParagraphStyle('LblDet', fontName='Helvetica', fontSize=7.0, leading=8.0, textColor=colors.HexColor("#111111"))
    cell_num_style = ParagraphStyle('NumVal', fontName='Helvetica', fontSize=7.5, leading=8.5, alignment=1)
    cell_num_tot_style = ParagraphStyle('NumTotVal', fontName='Helvetica-Bold', fontSize=7.5, leading=8.5, alignment=1)
    cell_sub_lbl_style = ParagraphStyle('SubLbl', fontName='Helvetica-Bold', fontSize=7.5, leading=8.5, textColor=colors.HexColor("#111111"), alignment=2)
    cell_grand_lbl_style = ParagraphStyle('GrandLbl', fontName='Helvetica-Bold', fontSize=8.0, leading=9.0, textColor=colors.HexColor("#111111"), alignment=2)
    cell_grand_num_style = ParagraphStyle('GrandNum', fontName='Helvetica-Bold', fontSize=8.5, leading=9.5, textColor=colors.HexColor("#006100"), alignment=1)
    
    regular_classes = class_loads_df[class_loads_df["load_category"] == "Regular Load"] if not class_loads_df.empty else pd.DataFrame()
    excess_classes = class_loads_df[class_loads_df["load_category"].isin(["Excess Load", "Excess/Overload"])] if not class_loads_df.empty else pd.DataFrame()
    consultations = class_loads_df[class_loads_df["load_category"] == "Consultation"] if not class_loads_df.empty else pd.DataFrame()
    
    def build_section_table(banner_text, banner_bg_hex, subtotal_lbl, sub_bg_hex, activity_items, default_cat):
        t_rows = []
        t_styles = [
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#D9D9D9")),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2),
            ('TOPPADDING', (0,0), (-1,-1), 2),
            ('LEFTPADDING', (0,0), (-1,-1), 2),
            ('RIGHTPADDING', (0,0), (-1,-1), 2),
        ]
        
        # Row 0: Banner row
        banner_cells = [Paragraph(banner_text, cell_banner_style)] + [Paragraph("", cell_banner_style)] * (num_dates + 2)
        t_rows.append(banner_cells)
        t_styles.append(('SPAN', (0, 0), (-1, 0)))
        t_styles.append(('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(banner_bg_hex)))
        t_styles.append(('BOTTOMPADDING', (0, 0), (-1, 0), 3))
        t_styles.append(('TOPPADDING', (0, 0), (-1, 0), 3))
        
        # Row 1: Date numbers
        # Row 2: Weekday abbreviations
        hdr_row1 = [
            Paragraph("Category", cell_hdr_lbl_style),
            Paragraph("Course / Activity Details", cell_hdr_lbl_style)
        ] + [Paragraph(str(d['day_num']), cell_hdr_style) for d in dates] + [Paragraph("Total", cell_hdr_style)]
        
        hdr_row2 = [
            Paragraph("", cell_hdr_lbl_style),
            Paragraph("", cell_hdr_lbl_style)
        ] + [Paragraph(d['day_abbr'], cell_hdr_style) for d in dates] + [Paragraph("", cell_hdr_style)]
        
        t_rows.append(hdr_row1)
        t_rows.append(hdr_row2)
        
        t_styles.extend([
            ('SPAN', (0, 1), (0, 2)),
            ('SPAN', (1, 1), (1, 2)),
            ('SPAN', (-1, 1), (-1, 2)),
            ('BACKGROUND', (0, 1), (-1, 2), colors.HexColor("#334155")),
        ])
        
        # Highlight weekends in headers
        for d_idx, d_info in enumerate(dates):
            if d_info['day_abbr'] in ("S", "SU"):
                t_styles.append(('BACKGROUND', (2 + d_idx, 1), (2 + d_idx, 2), colors.HexColor("#475569")))
                
        # Data Rows
        section_daily_totals = [0.0] * num_dates
        data_start_r = 3
        
        if not activity_items:
            empty_row = [
                Paragraph(default_cat, cell_lbl_cat_style),
                Paragraph(f"No {default_cat} activities scheduled", cell_lbl_det_style)
            ] + [Paragraph("-", cell_num_style)] * num_dates + [Paragraph("0.0", cell_num_tot_style)]
            t_rows.append(empty_row)
            t_styles.append(('BACKGROUND', (-1, data_start_r), (-1, data_start_r), colors.HexColor("#D9E1F2")))
            r_count = 1
        else:
            for idx, (cat, detail, key_name) in enumerate(activity_items):
                r_num = data_start_r + idx
                row_cells = [
                    Paragraph(cat, cell_lbl_cat_style),
                    Paragraph(detail, cell_lbl_det_style)
                ]
                row_sum = 0.0
                for d_idx, d_info in enumerate(dates):
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
                    row_sum += val
                    section_daily_totals[d_idx] += val
                    val_str = f"{val:.1f}" if val > 0 else "-"
                    row_cells.append(Paragraph(val_str, cell_num_style))
                    
                    if d_info['day_abbr'] in ("S", "SU"):
                        t_styles.append(('BACKGROUND', (2 + d_idx, r_num), (2 + d_idx, r_num), colors.HexColor("#F9F9F9")))
                        
                row_cells.append(Paragraph(f"{row_sum:.1f}", cell_num_tot_style))
                t_rows.append(row_cells)
                t_styles.append(('BACKGROUND', (-1, r_num), (-1, r_num), colors.HexColor("#D9E1F2")))
                if idx % 2 == 1:
                    t_styles.append(('BACKGROUND', (0, r_num), (-2, r_num), colors.HexColor("#F8FAFC")))
            r_count = len(activity_items)
            
        # Subtotal Row
        sub_r_num = data_start_r + r_count
        sub_row_cells = [
            Paragraph(subtotal_lbl, cell_sub_lbl_style),
            Paragraph("", cell_sub_lbl_style)
        ]
        for tot_val in section_daily_totals:
            tot_str = f"{tot_val:.1f}" if tot_val > 0 else "-"
            sub_row_cells.append(Paragraph(tot_str, cell_num_tot_style))
        section_sum = sum(section_daily_totals)
        sub_row_cells.append(Paragraph(f"{section_sum:.1f}", cell_num_tot_style))
        t_rows.append(sub_row_cells)
        
        t_styles.extend([
            ('SPAN', (0, sub_r_num), (1, sub_r_num)),
            ('BACKGROUND', (0, sub_r_num), (-1, sub_r_num), colors.HexColor(sub_bg_hex)),
        ])
        
        # Style weekend cells in subtotal row
        for d_idx, d_info in enumerate(dates):
            if d_info['day_abbr'] in ("S", "SU"):
                t_styles.append(('BACKGROUND', (2 + d_idx, sub_r_num), (2 + d_idx, sub_r_num), colors.HexColor("#ECECEC")))
                
        table = Table(t_rows, colWidths=col_widths)
        table.setStyle(TableStyle(t_styles))
        return table, section_daily_totals, section_sum
        
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
    
    # 1. Regular Load Table
    t_reg, reg_daily, reg_sum = build_section_table(
        f"📘 1. REGULAR LOAD TIMESHEET (Max Cap: {max_reg_units:.0f} Units - {acad_level})",
        "#1F4E78",
        "Regular Load Subtotal",
        "#D9E1F2",
        reg_items,
        "Regular Load"
    )
    story.append(t_reg)
    story.append(Spacer(1, 8))
    
    # 2. Excess Load Table
    t_exc, exc_daily, exc_sum = build_section_table(
        "⚡ 2. EXCESS LOAD TIMESHEET (Overload / Additional Teaching)",
        "#C65911",
        "Excess Load Subtotal",
        "#FCE4D6",
        exc_items,
        "Excess Load"
    )
    story.append(t_exc)
    story.append(Spacer(1, 8))
    
    # 3. Non-Teaching Table
    t_non, non_daily, non_sum = build_section_table(
        "🏢 3. NON-TEACHING ACTIVITIES (SC & HQ)",
        "#2E4053",
        "Non-Teaching Subtotal",
        "#D9E1F2",
        non_items,
        "Non-Teaching"
    )
    story.append(t_non)
    story.append(Spacer(1, 8))
    
    # 4. Consolidated Daily & Grand Totals Table
    grand_daily = [reg_daily[i] + exc_daily[i] + non_daily[i] for i in range(num_dates)]
    grand_total = sum(grand_daily)
    
    cons_rows = []
    cons_styles = [
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#D9D9D9")),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
        ('TOPPADDING', (0,0), (-1,-1), 2),
        ('LEFTPADDING', (0,0), (-1,-1), 2),
        ('RIGHTPADDING', (0,0), (-1,-1), 2),
    ]
    
    # Banner
    cons_banner = [Paragraph("CONSOLIDATED DAILY & GRAND TOTALS (ALL ACTIVITIES)", cell_banner_style)] + [Paragraph("", cell_banner_style)] * (num_dates + 2)
    cons_rows.append(cons_banner)
    cons_styles.extend([
        ('SPAN', (0, 0), (-1, 0)),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 3),
        ('TOPPADDING', (0, 0), (-1, 0), 3),
    ])
    
    # Totals row
    cons_tot_cells = [
        Paragraph("DAILY TOTALS (ALL ACTIVITIES)", cell_grand_lbl_style),
        Paragraph("", cell_grand_lbl_style)
    ]
    for tot_val in grand_daily:
        tot_str = f"{tot_val:.1f}" if tot_val > 0 else "-"
        cons_tot_cells.append(Paragraph(tot_str, cell_num_tot_style))
        
    cons_tot_cells.append(Paragraph(f"{grand_total:.1f}", cell_grand_num_style))
    cons_rows.append(cons_tot_cells)
    
    cons_styles.extend([
        ('SPAN', (0, 1), (1, 1)),
        ('BACKGROUND', (0, 1), (-2, 1), colors.HexColor("#D9E1F2")),
        ('BACKGROUND', (-1, 1), (-1, 1), colors.HexColor("#C6EFCE")),
    ])
    
    for d_idx, d_info in enumerate(dates):
        if d_info['day_abbr'] in ("S", "SU"):
            cons_styles.append(('BACKGROUND', (2 + d_idx, 1), (2 + d_idx, 1), colors.HexColor("#ECECEC")))
            
    t_cons = Table(cons_rows, colWidths=col_widths)
    t_cons.setStyle(TableStyle(cons_styles))
    story.append(t_cons)
    story.append(Spacer(1, 10))
    
    # 5. Verification Check and Signatures Section
    sig_label_style = ParagraphStyle('SigLbl', fontName='Helvetica-Bold', fontSize=8.5, alignment=0)
    sig_line_style = ParagraphStyle('SigLine', fontName='Helvetica', fontSize=8, alignment=0)
    
    target_status = "PASSED" if grand_total >= 80.0 else "UNDER-HOURS"
    status_color = "#2E7D32" if grand_total >= 80.0 else "#C62828"
    
    verif_text = (
        f"<b>Summary Breakdown:</b> Regular Load: <b>{reg_sum:.2f} hrs</b> | "
        f"Excess Load: <b>{exc_sum:.2f} hrs</b> | "
        f"Non-Teaching: <b>{non_sum:.2f} hrs</b> | "
        f"Grand Total: <b>{grand_total:.2f} hrs</b> / 80.0 hrs target &nbsp;&nbsp;|&nbsp;&nbsp; "
        f"Status: <font color='{status_color}'><b>{target_status}</b></font>"
    )
    story.append(Paragraph(verif_text, ParagraphStyle('Verif', fontName='Helvetica', fontSize=8.5, leading=11)))
    story.append(Spacer(1, 15))
    
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
        ('BOTTOMPADDING', (0,0), (-1,-1), 1),
    ]))
    story.append(sig_table)
    
    # Build Document
    doc.build(story)
    pdf_buffer.seek(0)
    return pdf_buffer.getvalue()
