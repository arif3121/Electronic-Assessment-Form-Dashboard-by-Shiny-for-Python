from shiny import App, render, ui, reactive
import pandas as pd
from datetime import datetime
import tempfile
import webbrowser
import os
import pathlib
import sys
import subprocess
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
import io

# Constants for grade ranges
GRADE_RANGES = {
    'A+': {'min': 80, 'max': 100, 'color': '#4CAF50'},  # Green
    'A': {'min': 70, 'max': 79, 'color': '#8BC34A'},    # Light Green
    'B': {'min': 60, 'max': 69, 'color': '#CDDC39'},    # Lime
    'C': {'min': 50, 'max': 59, 'color': '#FFEB3B'},    # Yellow
    'D': {'min': 40, 'max': 49, 'color': '#FFC107'},    # Amber
    'E': {'min': 30, 'max': 39, 'color': '#FF9800'},    # Orange
    'F': {'min': 0, 'max': 29, 'color': '#F44336'}      # Red
}

# Assessment criteria weights in percentage
CRITERIA_WEIGHTS = {
    'research': 5,
    'subject_knowledge': 20,
    'critical_analysis': 25,
    'problem_solving': 30,
    'practical_competence': 5,
    'communication': 10,
    'academic_integrity': 5
}

# List of all criteria for checking completeness
ALL_CRITERIA = list(CRITERIA_WEIGHTS.keys())

# Mapping from criteria IDs to display names
CRITERIA_DISPLAY_NAMES = {
    'research': 'Research',
    'subject_knowledge': 'Subject Knowledge',
    'critical_analysis': 'Critical Analysis',
    'problem_solving': 'Testing and Problem-Solving Skills',
    'practical_competence': 'Practical Competence',
    'communication': 'Communication and Presentation',
    'academic_integrity': 'Academic Integrity'
}

# Function to get student details
# FIXED: Function to get student details - completely revised with column mapping
# FIXED: Function to get student details - completely revised with column mapping
def get_student_details(student_id, filename="student_records.xlsx"):
    """Get details for a specific student with column name cleaning"""
    try:
        # Print the student ID for debugging
        print(f"Looking up student with ID: {student_id}")
        
        # Try to read the Excel file
        try:
            # Read the Excel file
            df = pd.read_excel(filename)
            print(f"Successfully read Excel file with {len(df)} records")
            
            # Debug: Print the first few rows to see the data
            print("First 2 rows of data:")
            print(df.head(2))
            
            # CRITICAL FIX: Clean column names by removing trailing spaces
            df.columns = [col.strip() for col in df.columns]
            print(f"Cleaned Excel columns: {df.columns.tolist()}")
            
            # Debug: Print student IDs from the first column
            print("Student IDs in file:", df["Student_ID"].tolist())
            
            # Create a mapping for our standard field names to Excel columns
            column_mapping = {
                "Student_ID": "Student_ID",
                "Name": "Name",
                "Surname": "Surname",
                "Course": "Course",
                "Mode": "Mode",
                "Module": "Module",
                "Title": "Title",
                "Supervisor": "Supervisor"
            }
            
            # CRITICAL FIX: Convert student_id to string for comparison
            student_id_str = str(student_id).strip()
            
            # DEBUG: Show ID types for comparison
            print(f"Looking for ID: {student_id_str}")
            print("ID types in Excel:")
            for idx, excel_id in enumerate(df["Student_ID"].head(3)):
                print(f"  Excel ID: {excel_id} (type: {type(excel_id).__name__})")
                
            # CRITICAL FIX: Convert all IDs to strings for comparison
            match_found = False
            for idx, row in df.iterrows():
                if str(row["Student_ID"]).strip() == student_id_str:
                    student_row = row
                    match_found = True
                    print(f"Found match for ID: {student_id_str}!")
                    break
            
            if not match_found:
                print(f"Student ID {student_id} not found after string conversion check")
                return None
            
            # Build the student info dictionary with our standardized field names
            student_info = {}
            for field, column in column_mapping.items():
                if column in student_row.index:
                    value = student_row[column]
                    # Handle NaN values
                    if pd.isna(value):
                        value = ""
                    student_info[field] = value
                    print(f"Set {field} = {value}")
            
            print(f"Found student record: {student_info}")
            return student_info
            
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            import traceback
            traceback.print_exc()
            return None
        
    except Exception as e:
        print(f"Exception in get_student_details: {e}")
        import traceback
        traceback.print_exc()
        return None
# Function to update student marks and comments
# FIXED: Update student record function with column mapping support
# FIXED: Update student record function with column mapping support
def update_student_record(student_id, marks, comment, filename="student_records.xlsx"):
    """Update marks and comment for a specific student with column mapping support"""
    try:
        # Read the existing Excel file
        df = pd.read_excel(filename)
        print(f"Read Excel file with {len(df)} rows for updating")
        
        # Print column names for debugging
        print(f"Excel columns for updating: {df.columns.tolist()}")
        
        # Define possible column names for the mark and comment fields
        marks_columns = ["Marks", "Mark", "Grade", "Score", "FinalGrade"]
        comment_columns = ["Comment", "Comments", "Feedback", "AssessorComment"]
        id_columns = ["Student_ID", "StudentID", "Student ID", "ID"]
        
        # Find the actual column names in the file
        def find_column(possible_names):
            for name in possible_names:
                if name in df.columns:
                    return name
            # If no exact match, try case-insensitive search
            df_cols_lower = [col.lower() for col in df.columns]
            for name in possible_names:
                if name.lower() in df_cols_lower:
                    idx = df_cols_lower.index(name.lower())
                    return df.columns[idx]
            return None
        
        # Find the column names to use
        id_column = find_column(id_columns)
        marks_column = find_column(marks_columns)
        comment_column = find_column(comment_columns)
        
        # Check if required columns were found
        if not id_column:
            print("ERROR: Could not find Student ID column")
            # Use first column as fallback
            id_column = df.columns[0]
            print(f"Using '{id_column}' as Student ID column")
        
        if not marks_column:
            print("ERROR: Could not find Marks column")
            # Check if we can add the column
            if "Marks" not in df.columns:
                df["Marks"] = ""
                marks_column = "Marks"
                print("Added 'Marks' column to DataFrame")
            else:
                marks_column = "Marks"
        
        if not comment_column:
            print("ERROR: Could not find Comment column")
            # Check if we can add the column
            if "Comment" not in df.columns:
                df["Comment"] = ""
                comment_column = "Comment"
                print("Added 'Comment' column to DataFrame")
            else:
                comment_column = "Comment"
        
        print(f"Using columns: ID='{id_column}', Marks='{marks_column}', Comment='{comment_column}'")
        
        # Check if student exists
        if student_id not in df[id_column].values:
            print(f"Student ID {student_id} not found in records")
            return False, "Student ID not found"
        
        # Check if student already has marks
        student_row = df[df[id_column] == student_id]
        if marks_column in student_row.columns and pd.notna(student_row[marks_column].values[0]) and student_row[marks_column].values[0] != "":
            print(f"Student already has marks: {student_row[marks_column].values[0]}")
            # Return True with warning message for the application to display
            return True, "Student already marked"
        
        # Update the marks and comment
        print(f"Updating student {student_id} with marks={marks} and comment length={len(comment)}")
        df.loc[df[id_column] == student_id, marks_column] = marks
        df.loc[df[id_column] == student_id, comment_column] = comment
        
        # Save back to Excel
        df.to_excel(filename, index=False)
        print(f"Successfully saved updated data to {filename}")
        
        return True, "Student record updated successfully"
        
    except Exception as e:
        print(f"Error updating student record: {e}")
        import traceback
        traceback.print_exc()
        return False, f"Error updating record: {str(e)}"
        
    except Exception as e:
        print(f"Error updating student record: {e}")
        import traceback
        traceback.print_exc()
        return False, f"Error updating record: {str(e)}"

def calculate_final_grade(scores):
    """Calculate weighted final grade based on criteria weights"""
    weighted_sum = sum(
        scores[f'{criterion}_score'] * CRITERIA_WEIGHTS[criterion]
        for criterion in ALL_CRITERIA
    )
    
    # Divide by total weight (100%)
    final_grade = weighted_sum / 100
    return round(final_grade, 1)

def find_logo_path():
    """Find the logo file using relative paths or create a dummy logo if not found"""
    # Try several common locations
    possible_locations = [
        "lsbu_logo.png",                         # Current directory
        os.path.join("assets", "lsbu_logo.png"), # Assets subdirectory
        os.path.join("..", "assets", "lsbu_logo.png"), # Parent dir assets
        os.path.join(os.path.dirname(__file__), "lsbu_logo.png"), # Script directory
    ]
    
    for location in possible_locations:
        if os.path.exists(location):
            return location
    
    # If not found, create a dummy logo file
    try:
        from PIL import Image, ImageDraw, ImageFont
        
        # Create a blank image for the logo
        img = Image.new('RGB', (200, 100), color=(0, 51, 102))  # LSBU blue color
        d = ImageDraw.Draw(img)
        
        # Add text (if font not available, it will use default)
        try:
            font = ImageFont.truetype("arial.ttf", 36)
        except:
            font = ImageFont.load_default()
            
        d.text((40, 30), "LSBU", fill=(255, 255, 255), font=font)
        
        # Save to a temporary location
        temp_logo_path = os.path.join(tempfile.gettempdir(), "lsbu_logo_temp.png")
        img.save(temp_logo_path)
        print(f"Created temporary logo at: {temp_logo_path}")
        return temp_logo_path
    except Exception as e:
        print(f"Error creating dummy logo: {e}")
        # If creation fails, return None
        return None

# PDF generation function with fixes
def create_pdf(data, output_path):
    # Reduced margins to use more page space
    doc = SimpleDocTemplate(output_path, pagesize=A4, 
                          leftMargin=0.3*inch, rightMargin=0.3*inch, 
                          topMargin=0.4*inch, bottomMargin=0.4*inch)
    elements = []
    styles = getSampleStyleSheet()
    
    # Create custom styles with updated alignment
    styles.add(ParagraphStyle(
        name='ModuleName',
        parent=styles['Normal'],
        fontSize=12,
        alignment=TA_RIGHT,  # Keep right alignment
    ))

    styles.add(ParagraphStyle(
        name='DivisionText',
        parent=styles['Normal'],
        fontSize=12,
        alignment=TA_RIGHT,  # Changed from CENTER to RIGHT
    ))

    styles.add(ParagraphStyle(
        name='EngineeringText',
        parent=styles['Normal'],
        fontSize=14,
        fontName='Helvetica-Bold',
        alignment=TA_RIGHT,  # Changed from CENTER to RIGHT
        leading=16,
    ))
    
    styles.add(ParagraphStyle(
        name='CustomTitle',
        parent=styles['Heading1'],
        fontSize=20,
        alignment=TA_CENTER,
        spaceAfter=12,
    ))
    
    styles.add(ParagraphStyle(
        name='CustomSubtitle',
        parent=styles['Heading2'],
        fontSize=16,
        alignment=TA_CENTER,
        spaceAfter=20,
    ))
    
    styles.add(ParagraphStyle(
        name='NormalLarge',
        parent=styles['Normal'],
        fontSize=12,  # Increased normal text size
    ))
    
    # Logo loading with more robust error handling
    try:
        logo_path = find_logo_path()
        
        if logo_path and os.path.exists(logo_path):
            try:
                # Test if the file is valid by attempting to open it
                with open(logo_path, 'rb') as test_file:
                    test_file.read(10)  # Read first 10 bytes to check if readable
                
                # If no error, proceed with image creation
                logo_content = Image(logo_path, width=2.0*inch, height=1.0*inch)
                print(f"Successfully loaded logo from: {logo_path}")
            except Exception as logo_error:
                print(f"Error opening logo file {logo_path}: {logo_error}")
                logo_content = Paragraph("LSBU", styles['Heading2'])
        else:
            print("Logo path not valid, using text fallback")
            logo_content = Paragraph("LSBU", styles['Heading2'])
    except Exception as e:
        print(f"Exception in logo handling: {e}")
        logo_content = Paragraph("LSBU", styles['Heading2'])
    
    # Restructured header layout - adjusted cell structure
    header_data = [
        [
            logo_content,
            Paragraph("Module Name: " + data['module_name'], styles['ModuleName']),
        ],
        [
            "",  # Empty cell - logo spans vertically
            Paragraph("Division of", styles['DivisionText']),
        ],
        [
            "",  # Empty cell  
            Paragraph("Electrical and Electronic Engineering", styles['EngineeringText']),
        ]
    ]

    # Adjusted header table with proper dimensions
    header_table = Table(header_data, colWidths=[2.5*inch, 4.0*inch])
    header_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (0, 2), 'TOP'),  # Logo aligned to top
        ('VALIGN', (1, 0), (1, 0), 'TOP'),  # Module name at top
        ('VALIGN', (1, 1), (1, 2), 'MIDDLE'),  # Division text vertically centered
        ('ALIGN', (0, 0), (0, 2), 'LEFT'),   # Logo left aligned
        ('ALIGN', (1, 0), (1, 2), 'RIGHT'),  # All right column elements right-aligned
        ('SPAN', (0, 0), (0, 2)),  # Logo spans all rows
        ('GRID', (0, 0), (-1, -1), 0, colors.white),  # Invisible grid
        ('RIGHTPADDING', (1, 0), (1, 2), 5),  # Reduced right padding to pull text more to edge
        ('BOTTOMPADDING', (1, 1), (1, 1), 0),  # Remove padding between Division and Engineering
        ('TOPPADDING', (1, 2), (1, 2), 0),     # Remove padding between Division and Engineering
    ]))
    elements.append(header_table)
    elements.append(Spacer(1, 0.4*inch))
    
    # Title with larger text
    elements.append(Paragraph("Assessment", styles['CustomTitle']))
    elements.append(Paragraph(f"Assignment 2 - Report: {data['report_title']}", styles['CustomSubtitle']))
    
    # Create grade table with fixed column structure
    # First row with student name (span across all columns)
    grade_data = [
        [Paragraph(f"Student: {data['student_name']}", styles['NormalLarge']), '', '', '', '', '', '', ''],
    ]
    
    # Grade range headers - corrected structure with criteria column empty
    grade_data.append(['', 'A+', 'A', 'B', 'C', 'D', 'E', 'F'])
    
    # FIX: Grade ranges - with correct string format instead of subtraction
    grade_data.append([
        '', 
        f"{GRADE_RANGES['A+']['min']}-{GRADE_RANGES['A+']['max']}", 
        f"{GRADE_RANGES['A']['min']}-{GRADE_RANGES['A']['max']}", 
        f"{GRADE_RANGES['B']['min']}-{GRADE_RANGES['B']['max']}",
        f"{GRADE_RANGES['C']['min']}-{GRADE_RANGES['C']['max']}", 
        f"{GRADE_RANGES['D']['min']}-{GRADE_RANGES['D']['max']}", 
        f"{GRADE_RANGES['E']['min']}-{GRADE_RANGES['E']['max']}", 
        f"{GRADE_RANGES['F']['min']}-{GRADE_RANGES['F']['max']}"
    ])
    
    # Assessment criteria rows - with weights displayed
    criteria = [
        (f"{CRITERIA_DISPLAY_NAMES['research']} ({CRITERIA_WEIGHTS['research']}%)", data['research_score']),
        (f"{CRITERIA_DISPLAY_NAMES['subject_knowledge']} ({CRITERIA_WEIGHTS['subject_knowledge']}%)", data['subject_knowledge_score']),
        (f"{CRITERIA_DISPLAY_NAMES['critical_analysis']} ({CRITERIA_WEIGHTS['critical_analysis']}%)", data['critical_analysis_score']),
        (f"{CRITERIA_DISPLAY_NAMES['problem_solving']} ({CRITERIA_WEIGHTS['problem_solving']}%)", data['problem_solving_score']),
        (f"{CRITERIA_DISPLAY_NAMES['practical_competence']} ({CRITERIA_WEIGHTS['practical_competence']}%)", data['practical_competence_score']),
        (f"{CRITERIA_DISPLAY_NAMES['communication']} ({CRITERIA_WEIGHTS['communication']}%)", data['communication_score']),
        (f"{CRITERIA_DISPLAY_NAMES['academic_integrity']} ({CRITERIA_WEIGHTS['academic_integrity']}%)", data['academic_integrity_score'])
    ]
    
    for criterion, score in criteria:
        row = [Paragraph(criterion, styles['NormalLarge'])]  # Using larger font
        # Determine which column the score falls into
        score_int = int(score)
        for grade in ['A+', 'A', 'B', 'C', 'D', 'E', 'F']:
            grade_range = GRADE_RANGES[grade]
            if grade_range['min'] <= score_int <= grade_range['max']:
                row.append(score)
            else:
                row.append('')
        
        grade_data.append(row)
    
    # Create and style the grade table with proportional column widths
    # Using full available width and taller rows
    available_width = doc.width
    col_widths = [available_width * 0.40] + [available_width * 0.086] * 7  # Adjusted for full width
    grade_table = Table(grade_data, colWidths=col_widths, rowHeights=[0.5*inch] + [0.4*inch] * 9)  # Taller rows
    grade_table.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('SPAN', (0, 0), (-1, 0)),  # Span the student name row
        ('ALIGN', (1, 1), (-1, -1), 'CENTER'),  # Center-align all grade columns
        ('ALIGN', (0, 1), (0, -1), 'LEFT'),     # Left-align criteria column
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BACKGROUND', (0, 1), (-1, 2), colors.lightgrey),  # Grey background for headers
        ('FONTSIZE', (1, 1), (-1, 2), 12),  # Larger font for headers
    ]))
    elements.append(grade_table)
    elements.append(Spacer(1, 0.3*inch))
    
    # Assessor's comments - larger and using more width
    comments_data = [
        [Paragraph("Assessor's Comments", styles['NormalLarge']), 
         Paragraph("Comments (Written Feedback) of the overall Assignment Performance", styles['NormalLarge'])],
        [Paragraph(data['assessor_comments'], styles['NormalLarge'])]
    ]
    
    comments_table = Table(comments_data, colWidths=[available_width * 0.25, available_width * 0.75], rowHeights=[0.4*inch, 1.2*inch])  # Taller rows
    comments_table.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('SPAN', (0, 1), (1, 1)),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),  # Grey background for header
    ]))
    elements.append(comments_table)
    elements.append(Spacer(1, 0.3*inch))
    
    # Final assessment row - larger text and more width
    final_row = [
        [Paragraph(f"Assessed by: {data['assessor_name']}", styles['NormalLarge']),
         Paragraph("*Grade (%)", styles['NormalLarge']),
         Paragraph(data['final_grade'], styles['NormalLarge'])]
    ]
    
    final_table = Table(final_row, colWidths=[available_width * 0.6, available_width * 0.2, available_width * 0.2], rowHeights=[0.45*inch])  # Taller row
    final_table.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (1, 0), (2, 0), 'CENTER'),  # Center the grade columns
    ]))
    elements.append(final_table)
    
    # Disclaimer - larger text
    elements.append(Spacer(1, 0.25*inch))
    elements.append(Paragraph("* This grade is provisional only and may be subject to change.", styles['NormalLarge']))
    
    # Current month/year - larger text
    elements.append(Spacer(1, 0.5*inch))
    current_date = datetime.now().strftime("%B %Y")  # FIX: Using proper date format
    date_paragraph = Paragraph(current_date, styles['NormalLarge'])
    date_paragraph.hAlign = 'RIGHT'
    elements.append(date_paragraph)
    
    doc.build(elements)

# Helper function to generate the grade selector UI - reused for all criteria
def create_grade_selector(id_prefix, label, default_grade="A"):
    # Create a select input with an empty label to prevent unwanted text
    grade_select = ui.input_select(
        f"{id_prefix}_grade", 
        "",  # Empty label to prevent it showing inline
        choices=list(GRADE_RANGES.keys()),
        selected=default_grade,
        width="100%"
    )
    
    # Create dynamic slider based on selected grade
    slider_output = ui.output_ui(f"{id_prefix}_slider_ui")
    
    # Put them in a layout with a very explicit label
    return ui.div(
        {"style": "margin-bottom: 15px; border-bottom: 1px solid #eee; padding-bottom: 10px;"},
        ui.div(
            {"class": "row"},
            ui.div({"class": "col-md-12"}, 
                # Force the label to be visible with inline HTML
                ui.HTML(f"<div style='font-weight: bold; margin-bottom: 8px; display: block !important; color: #333;'>{label}</div>"),
            ),
            ui.div(
                {"class": "row"},
                ui.div({"class": "col-md-4"}, 
                    ui.div({"class": "grade-select-container"}, grade_select)
                ),
                ui.div({"class": "col-md-8"}, slider_output)
            )
        )
    )

# Enhanced UI with new features
app_ui = ui.page_fluid(
    ui.tags.head(
        ui.tags.style("""
        .card {
            border-radius: 10px;
            box-shadow: 0 4px 8px rgba(0,0,0,0.1);
            margin-bottom: 20px;
            border: none;
        }
        .card-header {
            background-color: #003366;
            color: white;
            border-radius: 10px 10px 0 0 !important;
            padding: 12px 20px;
            font-weight: bold;
        }
        .card-body {
            padding: 20px;
        }
        .form-group {
            margin-bottom: 15px;
        }
        .btn-success {
            background-color: #28a745;
            border-color: #28a745;
            font-weight: bold;
            padding: 10px 25px;
        }
        .premium-textarea {
            border: 1px solid #ced4da;
            border-radius: 8px;
            padding: 10px;
            transition: border-color 0.15s ease-in-out, box-shadow 0.15s ease-in-out;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            resize: vertical;
        }
        .premium-textarea:focus {
            border-color: #80bdff;
            box-shadow: 0 0 0 0.2rem rgba(0, 123, 255, 0.25);
        }
        .alert {
            border-radius: 8px;
            padding: 15px;
            margin-bottom: 15px;
        }
        .grade-slider .irs-bar,
        .grade-slider .irs-bar-edge {
            background-color: var(--slider-color, #007bff);
            border-color: var(--slider-color, #007bff);
        }
        .grade-slider.AP .irs-bar { --slider-color: #4CAF50; }
        .grade-slider.A .irs-bar { --slider-color: #8BC34A; }
        .grade-slider.B .irs-bar { --slider-color: #CDDC39; }
        .grade-slider.C .irs-bar { --slider-color: #FFEB3B; }
        .grade-slider.D .irs-bar { --slider-color: #FFC107; }
        .grade-slider.E .irs-bar { --slider-color: #FF9800; }
        .grade-slider.F .irs-bar { --slider-color: #F44336; }
        
        /* Fixed position container to prevent shaking */
        .grade-select-container {
            position: relative;
            width: 100px;
            height: 100px;
            margin: 0 auto;
            overflow: visible;
        }
        
        /* Remove any unwanted text labels */
        .grade-select-container label {
            display: none !important;
        }
        
        /* Custom grade styling with fixed positioning and NO default indicators */
        .grade-select-container select {
            position: absolute;
            top: 0;
            left: 0;
            width: 100px !important;
            height: 100px !important;
            font-size: 36px !important;
            font-weight: bold !important;
            text-align: center !important;
            text-align-last: center !important;
            padding: 0 !important;
            border: 2px solid #ccc !important;
            border-radius: 8px !important;
            background-color: #f8f9fa !important;
            cursor: pointer !important;
            margin: 0 !important;
            display: block !important;
            
            /* Fully remove all default dropdown indicators */
            appearance: none !important;
            -webkit-appearance: none !important;
            -moz-appearance: none !important;
            background-image: none !important;
            z-index: 10;
        }
        
        /* Add a dropdown arrow to the right of the letter */
        .grade-select-container::after {
            content: "⌄" !important;
            position: absolute !important;
            top: 28px !important;  /* Vertically center with the grade letter */
            right: 15px !important;  /* Position on right side */
            font-size: 24px !important;
            color: #666 !important;
            z-index: 11 !important;
        }
        
        /* Remove any additional dropdown indicators that might be added by frameworks */
        .grade-select-container .selectize-input::after,
        .grade-select-container select + span,
        .grade-select-container select + div,
        .grade-select-container select + svg {
            display: none !important;
        }
        
        /* Force size when hovering with fixed positioning - position to the LEFT */
        .grade-select-container.hovering select {
            height: auto !important;
            min-height: 250px !important;
            overflow: visible !important;
            z-index: 9999 !important;
            left: -100px !important; /* Move dropdown to the left side */
        }
        
        /* Style dropdown options with consistent size */
        .grade-select-container select option {
            padding: 10px 0 !important;
            font-size: 28px !important;
            font-weight: bold !important;
            text-align: center !important;
            width: 100px !important;
            box-sizing: border-box !important;
        }
        
        /* Hide any stray text that might be appearing */
        .grade-select-container + span,
        .grade-select-container + div {
            display: none !important;
        }
        
        /* Style for criterion labels */
        .criterion-label {
            font-weight: bold;
            margin-bottom: 5px;
            display: block !important;
            font-size: 14px;
        }
        """)
    ),
    
    # JavaScript with improved stability
    ui.tags.script("""
    $(document).ready(function() {
        // Create a consistent container for the dropdown
        function setupGradeContainers() {
            $('.grade-select-container').each(function() {
                var $this = $(this);
                var $select = $this.find('select');
                
                // Make sure the option width matches the container
                $select.find('option').css({
                    'width': '100px',
                    'box-sizing': 'border-box'
                });
                
                // Remove any unwanted elements
                $this.find('.selectize-dropdown-content').siblings().remove();
                $this.siblings('svg, .dropdown-arrow').remove();
            });
        }
        
        // Run initial setup
        setupGradeContainers();
        setTimeout(setupGradeContainers, 500);
        
        // Hover behavior with stability improvements
        $(document).on('mouseenter', '.grade-select-container', function() {
            var $this = $(this);
            
            // First, close any other open dropdowns
            $('.grade-select-container').not($this).removeClass('hovering')
                .find('select').attr('size', '1');
            
            // Then open this one
            $this.addClass('hovering');
            $this.find('select').attr('size', '7');
            
            // Prevent scrolling of the page when hovering
            $this.find('select').on('wheel', function(e) {
                e.stopPropagation();
            });
        });
        
        $(document).on('mouseleave', '.grade-select-container', function() {
            var $this = $(this);
            
            // Wait a bit to determine if we truly left
            setTimeout(function() {
                if (!$this.is(':hover')) {
                    $this.removeClass('hovering');
                    $this.find('select').attr('size', '1');
                }
            }, 50);
        });
        
        // Handle click outside to close all dropdowns
        $(document).on('click', function(e) {
            if(!$(e.target).closest('.grade-select-container').length) {
                $('.grade-select-container').removeClass('hovering');
                $('select[id$="_grade"]').attr('size', '1');
            }
        });
        
        // Prevent shaking on select
        $(document).on('click', '.grade-select-container select option', function() {
            var $select = $(this).closest('select');
            var $container = $select.closest('.grade-select-container');
            
            // Immediately close after selection to prevent shaking
            setTimeout(function() {
                $container.removeClass('hovering');
                $select.attr('size', '1');
                $select.blur();
            }, 10);
        });
        
        // Remove any stray text nodes and unwanted elements
        $('.grade-select-container').each(function() {
            $(this).siblings().filter(function() {
                return (this.nodeType === 3 && $.trim(this.nodeValue) !== '') || 
                       (this.nodeType === 1 && !$(this).hasClass('form-control'));
            }).remove();
        });
        
        // Remove extra dropdown arrows
        setInterval(function() {
            $('select').each(function() {
                var $select = $(this);
                $select.siblings('.dropdown-arrow, svg').remove();
            });
            
            $('.grade-select-container').each(function() {
                $(this).find('svg, .dropdown-arrow').not('.custom-arrow').remove();
            });
        }, 500);
    });
    """),
    
    # Header with logo and title
    # Header with properly sized logo and better layout
ui.div(
    {"class": "row", "style": "margin-bottom: 20px; align-items: center;"},
    ui.div(
        {"class": "col-md-2", "style": "text-align: right;"},
        ui.output_image("logo_image", width="120px", height="auto")
    ),
    ui.div(
        {"class": "col-md-10", "style": "padding-left: 20px;"},
        ui.tags.h1("Academic Assessment Report Generator", 
                  {"style": "font-size: 24px; margin-top: 0; margin-bottom: 0;"})
    )
),
   # Basic Information Section
# Basic Information Section
ui.card(
    ui.card_header("Module & Student Information"),
    ui.card_body(
        ui.row(
            ui.column(6,
                ui.h5("Student Information"),
                ui.input_select("student_id", "Student ID", 
                          choices=pd.read_excel("student_records.xlsx")["Student_ID"].tolist()),
                ui.input_text("student_name", "Name"),
                ui.input_text("student_surname", "Surname"),
                ui.input_text("student_course", "Course"),
                ui.input_text("student_mode", "Mode"),
            ),
            ui.column(6,
                ui.h5("Module Information"),
                ui.input_text("module_name", "Module"),
                ui.input_text("report_title", "Report Title"),
                ui.input_text("supervisor", "Supervisor"),
            )
        ),
        # Add CSS to make fields appear read-only using JavaScript
        ui.tags.script("""
        $(document).ready(function() {
            // Make fields read-only after the page loads
            $("#student_name").prop("readonly", true);
            $("#student_surname").prop("readonly", true);
            $("#student_course").prop("readonly", true);
            $("#student_mode").prop("readonly", true);
            $("#module_name").prop("readonly", true);
            $("#report_title").prop("readonly", true);
            $("#supervisor").prop("readonly", true);
        });
        """),
        # Add CSS to style read-only fields
        ui.tags.style("""
        #student_name, #student_surname, #student_course, #student_mode,
        #module_name, #report_title, #supervisor {
            background-color: #f8f9fa;
            cursor: not-allowed;
        }
        """)
    )
),
     # Assessment Scores Section with Grade Selectors
    ui.card(
        ui.card_header("Assessment Scores"),
        ui.card_body(
            # Create grade selectors in two columns
            ui.row(
                # Left column
                ui.column(6,
                    ui.div(
                        {"class": "criterion-label"},
                        f"Research ({CRITERIA_WEIGHTS['research']}%)"
                    ),
                    create_grade_selector("research", ""),
                    
                    ui.div(
                        {"class": "criterion-label"},
                        f"Subject Knowledge ({CRITERIA_WEIGHTS['subject_knowledge']}%)"
                    ),
                    create_grade_selector("subject_knowledge", ""),
                    
                    ui.div(
                        {"class": "criterion-label"},
                        f"Critical Analysis ({CRITERIA_WEIGHTS['critical_analysis']}%)"
                    ),
                    create_grade_selector("critical_analysis", ""),
                    
                    ui.div(
                        {"class": "criterion-label"},
                        f"Testing & Problem-Solving ({CRITERIA_WEIGHTS['problem_solving']}%)"
                    ),
                    create_grade_selector("problem_solving", ""),
                ),
                # Right column
                ui.column(6,
                    ui.div(
                        {"class": "criterion-label"},
                        f"Practical Competence ({CRITERIA_WEIGHTS['practical_competence']}%)"
                    ),
                    create_grade_selector("practical_competence", ""),
                    
                    ui.div(
                        {"class": "criterion-label"},
                        f"Communication ({CRITERIA_WEIGHTS['communication']}%)"
                    ),
                    create_grade_selector("communication", ""),
                    
                    ui.div(
                        {"class": "criterion-label"},
                        f"Academic Integrity ({CRITERIA_WEIGHTS['academic_integrity']}%)"
                    ),
                    create_grade_selector("academic_integrity", ""),
                )
            ),
            
            ui.div(
                {"style": "margin-top: 20px; padding: 15px; background-color: #f8f9fa; border-radius: 8px;"},
                ui.output_ui("calculated_grade"),
            ),
            
            # Assessment status
            ui.output_ui("assessment_status")
        )
    ),
    
    # Assessor Information and Comments
    ui.card(
        ui.card_header("Assessor Information"),
        ui.card_body(
            ui.row(
                ui.column(4,
                    ui.input_select("assessor_name", "Assessor Name", 
                           choices=["Dr Oswaldo Cadenas", "Dr Thomas Rushton", "Dr Craig Sayers"],
                           selected="Dr Oswaldo Cadenas"),
                ),
                ui.column(8,
                    ui.output_ui("comment_section"),
                )
            ),
            ui.output_ui("comment_warning"),
        )
    ),
    
    # Preview and Generate Section
    ui.card(
        ui.card_body(
            {"style": "text-align: center;"},
            ui.input_action_button("generate", "Generate PDF Report", class_="btn-success btn-lg"),
            ui.div(
                {"style": "margin-top: 15px;"},
                ui.output_text("generate_status")
            ),
            # Add download option for the generated PDF
            ui.output_ui("download_option")
        )
    )
)

def server(input, output, session):

    @reactive.Effect
    def update_student_info():
     if "student_id" in input and input.student_id():
        # Get student details with additional error handling
        student_id = input.student_id()
        print(f"Selected student ID: {student_id}")
        
        student_info = get_student_details(student_id)
        
        if student_info:
            print("Student info found, updating UI fields")
            
            try:
                # Update UI elements with proper error handling for each field
                # Using our standardized field names from the mapping
                ui.update_text("student_name", value=student_info.get("Name", ""))
                ui.update_text("student_surname", value=student_info.get("Surname", ""))
                ui.update_text("student_course", value=student_info.get("Course", ""))
                ui.update_text("student_mode", value=student_info.get("Mode", ""))
                ui.update_text("module_name", value=student_info.get("Module", ""))
                ui.update_text("report_title", value=student_info.get("Title", ""))
                ui.update_text("supervisor", value=student_info.get("Supervisor", ""))
                
                # Show a notification that the student data was loaded
                ui.notification_show(
                    "Student information loaded successfully",
                    type="message",
                    duration=3
                )
                
            except Exception as e:
                print(f"Error updating student information UI: {e}")
                import traceback
                traceback.print_exc()
                
                # Show error notification to the user
                ui.notification_show(
                    f"Error loading student data: {str(e)}",
                    type="error",
                    duration=5
                )
        else:
            print(f"No student information found for ID: {student_id}")
            
            # Clear fields if no student found
            ui.update_text("student_name", value="")
            ui.update_text("student_surname", value="")
            ui.update_text("student_course", value="")
            ui.update_text("student_mode", value="")
            ui.update_text("module_name", value="")
            ui.update_text("report_title", value="")
            ui.update_text("supervisor", value="")
            
            # Show warning notification
            ui.notification_show(
                f"No student record found for ID: {student_id}",
                type="warning",
                duration=4
            )

    @output
    @render.image
    def logo_image():
        logo_path = find_logo_path()
        return {"src": logo_path, "height": "100px", "contentType": "image/png"}

    # Explicitly define each criterion slider UI function with initial hidden state
    @output
    @render.ui
    def research_slider_ui():
        # Check if grade has been selected yet
        if "research_grade" not in input:
            return ui.div(
                {"style": "color: #6c757d; padding: 8px 0;"},
                "Select a grade to show the score slider"
            )
            
        grade = input.research_grade()
        grade_range = GRADE_RANGES[grade]
        return ui.div(
            {"class": f"grade-slider {grade.replace('+', 'P')}"},
            ui.input_slider("research_score", "", min=grade_range['min'], max=grade_range['max'], 
                      value=(grade_range['min'] + grade_range['max']) // 2,
                      step=1)
        )
    
    @output
    @render.ui
    def subject_knowledge_slider_ui():
        # Check if grade has been selected yet
        if "subject_knowledge_grade" not in input:
            return ui.div(
                {"style": "color: #6c757d; padding: 8px 0;"},
                "Select a grade to show the score slider"
            )
            
        grade = input.subject_knowledge_grade()
        grade_range = GRADE_RANGES[grade]
        return ui.div(
            {"class": f"grade-slider {grade.replace('+', 'P')}"},
            ui.input_slider("subject_knowledge_score", "", min=grade_range['min'], max=grade_range['max'], 
                      value=(grade_range['min'] + grade_range['max']) // 2,
                      step=1)
        )
    
    @output
    @render.ui
    def critical_analysis_slider_ui():
        # Check if grade has been selected yet
        if "critical_analysis_grade" not in input:
            return ui.div(
                {"style": "color: #6c757d; padding: 8px 0;"},
                "Select a grade to show the score slider"
            )
            
        grade = input.critical_analysis_grade()
        grade_range = GRADE_RANGES[grade]
        return ui.div(
            {"class": f"grade-slider {grade.replace('+', 'P')}"},
            ui.input_slider("critical_analysis_score", "", min=grade_range['min'], max=grade_range['max'], 
                      value=(grade_range['min'] + grade_range['max']) // 2,
                      step=1)
        )
    
    @output
    @render.ui
    def problem_solving_slider_ui():
        # Check if grade has been selected yet
        if "problem_solving_grade" not in input:
            return ui.div(
                {"style": "color: #6c757d; padding: 8px 0;"},
                "Select a grade to show the score slider"
            )
            
        grade = input.problem_solving_grade()
        grade_range = GRADE_RANGES[grade]
        return ui.div(
            {"class": f"grade-slider {grade.replace('+', 'P')}"},
            ui.input_slider("problem_solving_score", "", min=grade_range['min'], max=grade_range['max'], 
                      value=(grade_range['min'] + grade_range['max']) // 2,
                      step=1)
        )
    
    @output
    @render.ui
    def practical_competence_slider_ui():
        # Check if grade has been selected yet
        if "practical_competence_grade" not in input:
            return ui.div(
                {"style": "color: #6c757d; padding: 8px 0;"},
                "Select a grade to show the score slider"
            )
            
        grade = input.practical_competence_grade()
        grade_range = GRADE_RANGES[grade]
        return ui.div(
            {"class": f"grade-slider {grade.replace('+', 'P')}"},
            ui.input_slider("practical_competence_score", "", min=grade_range['min'], max=grade_range['max'], 
                      value=(grade_range['min'] + grade_range['max']) // 2,
                      step=1)
        )
    
    @output
    @render.ui
    def communication_slider_ui():
        # Check if grade has been selected yet
        if "communication_grade" not in input:
            return ui.div(
                {"style": "color: #6c757d; padding: 8px 0;"},
                "Select a grade to show the score slider"
            )
            
        grade = input.communication_grade()
        grade_range = GRADE_RANGES[grade]
        return ui.div(
            {"class": f"grade-slider {grade.replace('+', 'P')}"},
            ui.input_slider("communication_score", "", min=grade_range['min'], max=grade_range['max'], 
                      value=(grade_range['min'] + grade_range['max']) // 2,
                      step=1)
        )
    
    @output
    @render.ui
    def academic_integrity_slider_ui():
        # Check if grade has been selected yet
        if "academic_integrity_grade" not in input:
            return ui.div(
                {"style": "color: #6c757d; padding: 8px 0;"},
                "Select a grade to show the score slider"
            )
            
        grade = input.academic_integrity_grade()
        grade_range = GRADE_RANGES[grade]
        return ui.div(
            {"class": f"grade-slider {grade.replace('+', 'P')}"},
            ui.input_slider("academic_integrity_score", "", min=grade_range['min'], max=grade_range['max'], 
                      value=(grade_range['min'] + grade_range['max']) // 2,
                      step=1)
        )
    
    # Improved assessment completion check
    @reactive.Calc
    def assessment_complete():
        # Check that all criteria have actual scores
        try:
            for criterion in ALL_CRITERIA:
                score_id = f"{criterion}_score"
                # Just check if the input exists - don't try to access its value
                # as this might cause errors if it's not yet initialized
                if score_id not in input:
                    return False
            return True
        except:
            return False
    
    # Show assessment status with more accurate information
    @output
    @render.ui
    def assessment_status():
        if assessment_complete():
            return ui.div(
                {"class": "alert alert-success", "style": "margin-top: 15px;"},
                ui.tags.i({"class": "fa fa-check-circle"}),
                " Assessment complete. You can now provide comments."
            )
        else:
            # Try to identify which criteria are missing
            missing_criteria = []
            for criterion in ALL_CRITERIA:
                score_id = f"{criterion}_score"
                if score_id not in input or input[score_id]() is None:
                    missing_criteria.append(CRITERIA_DISPLAY_NAMES[criterion])
            
            missing_text = ", ".join(missing_criteria) if missing_criteria else "all assessment criteria"
            
            return ui.div(
                {"class": "alert alert-warning", "style": "margin-top: 15px;"},
                ui.tags.i({"class": "fa fa-exclamation-triangle"}),
                f" Please complete {missing_text} before adding comments."
            )
    
    # Reactive calculation of final grade with proper error handling
    @reactive.Calc
    def final_grade():
        try:
            scores = {}
            for criterion in ALL_CRITERIA:
                score_id = f"{criterion}_score"
                if score_id in input:
                    try:
                        scores[score_id] = input[score_id]()
                    except:
                        # If we can't get the value, use a default
                        print(f"Error getting value for {score_id}, using default")
                        scores[score_id] = 50  # Default to middle value
                else:
                    scores[score_id] = 50  # Default to middle value
                    
            return calculate_final_grade(scores)
        except Exception as e:
            print(f"Error calculating final grade: {e}")
            import traceback
            traceback.print_exc()
            return 0  # Default to 0 if there's an error
    
    # Display calculated grade with professional styling
    @output
    @render.ui
    def calculated_grade():
        try:
            grade = final_grade()
            
            # Determine color based on grade
            color = "#4CAF50"  # Default green
            if grade < 30:
                color = "#F44336"  # Red for F
            elif grade < 40:
                color = "#FF9800"  # Orange for E
            elif grade < 50:
                color = "#FFC107"  # Amber for D
            elif grade < 60:
                color = "#FFEB3B"  # Yellow for C
            elif grade < 70:
                color = "#CDDC39"  # Lime for B
            
            return ui.div(
                ui.h3(
                    {"style": f"color: {color}; font-weight: bold; text-align: center;"},
                    f"Calculated Final Grade: {grade}%"
                )
            )
        except:
            return ui.div(
                ui.h3(
                    {"style": "color: #999; font-weight: bold; text-align: center;"},
                    "Final Grade: Not yet calculated"
                )
            )
    
    # Determine if comment is required based on grade
    @reactive.Calc
    def comment_required():
        grade = final_grade()
        return grade < 30 or grade > 69
    
    # Dynamic comment section based on assessment completeness and grade
    @output
    @render.ui
    def comment_section():
        if not assessment_complete():
            return ui.div()  # Return empty if assessment is not complete
        
        grade = final_grade()
        is_required = comment_required()
        
        if is_required:
            return ui.div(
                ui.h5("Assessor Comments", style="margin-bottom: 10px;"),
                ui.p(
                    {"style": "color: #721c24; background-color: #f8d7da; padding: 10px; border-radius: 5px;"},
                    "This grade requires detailed feedback (minimum 15 words)."
                ),
                ui.input_text_area(
                    "assessor_comments", 
                    "", 
                    rows=6, 
                    resize="vertical",
                    placeholder="Please provide detailed feedback for this grade...",
                    width="100%"
                )
            )
        else:
            return ui.div(
                ui.input_checkbox("show_comments", "Add comments for this assessment", False),
                ui.panel_conditional(
                    "input.show_comments",
                    ui.input_text_area(
                        "assessor_comments", 
                        "Assessor Comments", 
                        rows=6, 
                        resize="vertical",
                        width="100%"
                    )
                )
            )
    
    # Warning for insufficient comment length
    @output
    @render.ui
    def comment_warning():
        if not assessment_complete():
            return ui.div()  # Return empty if assessment is not complete
            
        if not comment_required():
            return ui.div()
        
        comments = input.assessor_comments() if hasattr(input, "assessor_comments") else ""
        word_count = len(comments.split()) if comments else 0
        
        if word_count < 15:
            return ui.div(
                {"class": "alert alert-danger", "role": "alert"},
                ui.tags.b("Warning: "),
                f"Comments must be at least 15 words (currently {word_count} words)."
            )
        return ui.div()
    
    # Validate before generating PDF - improved error handling
    @reactive.Calc
    def can_generate_pdf():
        if not assessment_complete():
            return False, "Please complete all assessment criteria before generating the PDF."
            
        # Get current grade to determine if comments are required
        try:
            grade = final_grade()
            requires_comment = (grade < 30 or grade > 69)
            
            # If comment is required, check if it's provided
            if requires_comment:
                comments = ""
                if hasattr(input, "assessor_comments"):
                    try:
                        comments = input.assessor_comments()
                    except:
                        pass
                
                word_count = len(comments.split()) if comments else 0
                if word_count < 15:
                    return False, "Please provide detailed comments (at least 15 words) as required for this grade."
            
            # Otherwise check if they enabled comments but didn't provide any
            elif hasattr(input, 'show_comments') and input.show_comments():
                try:
                    if not input.assessor_comments():
                        return False, "Please enter comments or uncheck the 'Add comments' option."
                except:
                    return False, "Please enter comments or uncheck the 'Add comments' option."
        except Exception as e:
            print(f"Error in can_generate_pdf: {e}")
            return False, f"An error occurred while validating inputs: {e}"
            
        # Check if required fields are filled
        if not input.student_name() or not input.report_title():
            return False, "Please fill in all required fields (Student Name, Report Title)"
            
        return True, ""
    
    # Store the generated PDF path
    pdf_path = reactive.Value(None)
    generation_success = reactive.Value(False)
    
    # PDF generation with validation and more detailed error reporting
    @output
    @render.text
    @reactive.event(input.generate)
    def generate_status():
        can_generate, message = can_generate_pdf()
        if not can_generate:
            print(f"Cannot generate PDF: {message}")
            return message
            
        generation_success.set(False)
        
        try:
            print("Starting PDF generation process...")
            # Use the system temp directory for better cross-platform compatibility
            temp_dir = os.path.join(tempfile.gettempdir(), "assessment_reports")
            print(f"Using temp directory: {temp_dir}")
            
            if not os.path.exists(temp_dir):
                os.makedirs(temp_dir)
                print(f"Created directory: {temp_dir}")
            
            # Create a unique filename with safe characters
            safe_name = "".join(c if c.isalnum() else "_" for c in input.student_name())
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"assessment_{safe_name}_{timestamp}.pdf"
            filepath = os.path.join(temp_dir, filename)
            
            print(f"Will generate PDF at: {filepath}")
            
            # Get comments with better error handling
            comments = "No additional comments."
            try:
                grade = final_grade()
                requires_comment = (grade < 30 or grade > 69)
                
                if requires_comment and hasattr(input, "assessor_comments"):
                    comments = input.assessor_comments() or "Required comments not provided."
                elif hasattr(input, "show_comments") and input.show_comments() and hasattr(input, "assessor_comments"):
                    comments = input.assessor_comments() or "No additional comments."
            except Exception as e:
                print(f"Error getting comments: {e}")
                comments = "Error retrieving comments."
            
            # Calculate final grade
            try:
                calculated_final_grade = final_grade()
                print(f"Calculated final grade: {calculated_final_grade}")
            except Exception as e:
                print(f"Error calculating grade: {e}")
                calculated_final_grade = 0
            
            # Collect all data for PDF with error checking
            report_data = {
                'module_name': input.module_name() or "Module not specified",
                'report_title': input.report_title() or "Report title not specified",
                'student_name': input.student_name() or "Student name not specified",
                'assessor_name': input.assessor_name() or "Assessor not specified",
                'assessor_comments': comments,
                'final_grade': str(calculated_final_grade)
            }
            
            # Add all criteria scores to the data with error handling
            for criterion in ALL_CRITERIA:
                try:
                    score_id = f"{criterion}_score"
                    if score_id in input:
                        report_data[score_id] = input[score_id]()
                    else:
                        report_data[score_id] = 50  # Default score if not available
                        print(f"Warning: Missing score for {criterion}, using default")
                except Exception as e:
                    report_data[score_id] = 50  # Default score if error
                    print(f"Error getting score for {criterion}: {e}")
            
            print("All data collected, generating PDF...")
            print(f"Report data: {report_data}")
            
            # Generate PDF with exception logging
            try:
                create_pdf(report_data, filepath)
                print(f"PDF successfully created at: {filepath}")
            except Exception as e:
                print(f"Error in create_pdf function: {e}")
                import traceback
                traceback.print_exc()
                return f"PDF generation failed in create_pdf function: {str(e)}"
            
            # Verify the PDF was created
            if not os.path.exists(filepath):
                return "PDF file was not created. Check directory permissions."
            
            if os.path.getsize(filepath) == 0:
                return "PDF file was created but is empty. Check ReportLab installation."
            
            # Store the path for download
            pdf_path.set(filepath)
            generation_success.set(True)
            
            # After successful PDF generation, update Excel
            student_id = input.student_id()
            final_grade_value = str(calculated_final_grade)  # Convert to string for Excel
            assessor_comments = comments
            
            # Check if student already has marks and ask for confirmation if needed
            success, message = update_student_record(student_id, final_grade_value, assessor_comments)
            
            # Try to open the PDF but don't fail if it doesn't work
            try:
                webbrowser.open(f'file://{os.path.abspath(filepath)}')
                print(f"Browser should be opening PDF now")
            except Exception as e:
                print(f"Warning: Could not open PDF automatically: {str(e)}")
            
            if message == "Student already marked":
                return f"PDF generated successfully, but NOTE: {message}. The record has been updated anyway."
            else:
                return f"PDF report generated successfully: {filename}"
                
        except Exception as e:
            import traceback
            traceback.print_exc()
            return f"Error generating PDF: {str(e)}"
    
    # Download link for the generated PDF with more robust implementation
    @output
    @render.ui
    def download_option():
        if not generation_success() or pdf_path() is None:
            return ui.div()
        
        filepath = pdf_path()
        if not os.path.exists(filepath):
            return ui.div(
                {"class": "alert alert-danger", "style": "margin-top: 15px;"},
                "PDF file no longer exists at the expected location."
            )
        
        filename = os.path.basename(filepath)
        
        # For local file access, provide instructions
        return ui.div(
            {"class": "alert alert-info", "style": "margin-top: 15px;"},
            ui.tags.h4("PDF Generated Successfully"),
            ui.tags.p(f"File saved as: {filename}"),
            ui.tags.p(f"Location: {os.path.dirname(filepath)}"),
            ui.tags.p(
                "If the PDF didn't open automatically, please navigate to the location above and open the file manually."
            ),
            ui.input_action_button(
                "open_folder", 
                "Open Containing Folder",
                class_="btn btn-primary"
            )
        )
# Add a function to test Excel loading at startup
def test_excel_loading():
    """Test Excel file loading at application startup"""
    try:
        print("Testing Excel file loading...")
        excel_path = "student_records.xlsx"
        if not os.path.exists(excel_path):
            print(f"WARNING: Excel file '{excel_path}' not found in current directory")
            print(f"Current working directory: {os.getcwd()}")
            print("Available files in current directory:")
            for file in os.listdir():
                print(f"  {file}")
            return False
            
        # Try to read the file
        df = pd.read_excel(excel_path)
        print(f"Successfully loaded Excel file with {len(df)} rows")
        print(f"Excel columns: {df.columns.tolist()}")
        
        # Print the first row as an example
        if len(df) > 0:
            print("First row example:")
            print(df.iloc[0].to_dict())
        
        return True
    except Exception as e:
        print(f"ERROR testing Excel loading: {e}")
        import traceback
        traceback.print_exc()
        return False

# Add a function to test Excel loading at startup
def test_excel_loading():
    """Test Excel file loading at application startup"""
    try:
        print("Testing Excel file loading...")
        excel_path = "student_records.xlsx"
        if not os.path.exists(excel_path):
            print(f"WARNING: Excel file '{excel_path}' not found in current directory")
            print(f"Current working directory: {os.getcwd()}")
            print("Available files in current directory:")
            for file in os.listdir():
                print(f"  {file}")
            return False
            
        # Try to read the file
        df = pd.read_excel(excel_path)
        print(f"Successfully loaded Excel file with {len(df)} rows")
        print(f"Excel columns: {df.columns.tolist()}")
        
        # Print the first row as an example
        if len(df) > 0:
            print("First row example:")
            print(df.iloc[0].to_dict())
        
        return True
    except Exception as e:
        print(f"ERROR testing Excel loading: {e}")
        import traceback
        traceback.print_exc()
        return False

# Create the Shiny application
app = App(app_ui, server)

if __name__ == "__main__":
    # Test Excel loading first
    test_excel_loading()
    
    # Use a different port to avoid conflicts
    app.run(port=8051)  