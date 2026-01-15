import os
import re
import pandas as pd
import logging
from flask import Flask, request, jsonify, render_template, send_file
from werkzeug.utils import secure_filename
from pdfminer.high_level import extract_text
from flask import send_from_directory

# Configure logging
logging.basicConfig(level=logging.INFO)
logging.getLogger("pdfminer").setLevel(logging.ERROR)

app = Flask(__name__, static_folder="static", static_url_path="/static")

# Define directories
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER


# Function to extract text from a PDF file
def extract_text_from_pdf(file_path):
    text = extract_text(file_path)
    with open(
        os.path.join(UPLOAD_FOLDER, "pdf.txt"), "w", encoding="utf-8"
    ) as output_file:
        output_file.write(text)
    return text


# Helper functions
def safe_search(pattern, text, default="N/A"):
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(1).strip() if match else default


def extract_by_line_offset(text, keyword, offset, default="N/A"):
    lines = text.split("\n")
    for i, line in enumerate(lines):
        if keyword in line:
            target_line = i + offset
            if 0 <= target_line < len(lines):
                match = re.search(r"(\d+)", lines[target_line])
                return match.group(1) if match else default
    return default


# # Function to extract subjects
# def extract_subjects(text):
#     """Extracts subject names from the marksheet text and removes unwanted entries."""

#     # Initialize subjects list to store results from both conditions
#     subjects = []

#     # First condition - Extract using "TITLE OF COURSES" pattern
#     match1 = re.search(r"TITLE OF COURSES\s*TH/\s*PR\s*COURSE\s*HEAD\s*(.*?)\s*ESE", text, re.DOTALL)
#     if match1:
#         subjects_raw = match1.group(1).strip()
#         subjects += [s.strip() for s in re.split(r"\n+", subjects_raw) if s.strip() and not re.match(r"^(TH|PR|TW|ESE|PA|OR|PJ|IT|LSP|AG|@|PLY|AP|\*|WFLS|%|WFLY|CON|FT|A.T.K.T)$", s)]

#     # Second condition - Extract using "TITLE OF SUBJECTS" pattern
#     match2 = re.search(r"TITLE OF SUBJECTS\s*(.*?)\s*(?=TOTAL MARKS|RESULT DECLARED|INSTRUCTIONS)", text, re.DOTALL)
#     if match2:
#         subjects_raw = match2.group(1).strip()
#         subjects += [s.strip() for s in re.split(r"\n+", subjects_raw) if s.strip()]

#         # Remove unwanted terms from the subjects
#         unwanted_terms = [
#             "THEORY", "PRACTICALS", "TOTAL", "CREDITS", "MAX", "OBT", "SLA",
#             "FA", "SA", "TH", "PR", "TW", "ESE", "PA", "OR", "PJ", "IT",
#             "LSP", "AG", "@", "PLY", "*", "WFLS", "%", "WFLY", "CON", "A.T.K.T",
#             "OBT MAX OBT", "DATE :", "This Marksheet is Downloaded from Internet",
#             "MAHARASHTRA STATE BOARD OF TECHNICAL EDUCATION", "TOTAL MAX."
#         ]

#         # Remove any unwanted term or numeric entries
#         subjects = [
#             subject for subject in subjects if subject not in unwanted_terms and not subject.isdigit() and len(subject.split()) > 1
#         ]

#     return subjects


# Function to extract subject-wise marks from marksheet
def extract_subject_table(pdf_text):
    """
    Extracts subject-wise marks from MSBTE marksheet text.
    
    Returns:
        List[Dict]: List of subject dictionaries with marks details
    """
    subjects = []
    
    try:
        # Find the subject table section
        start_pattern = r"TITLE OF SUBJECTS"
        end_pattern = r"DATE\s*:\s*[\d/]+"
        
        start_match = re.search(start_pattern, pdf_text)
        end_match = re.search(end_pattern, pdf_text)
        
        if not start_match or not end_match:
            logging.warning("Could not find subject table boundaries")
            return subjects
        
        # Extract the subject table content
        start_pos = start_match.end()
        end_pos = end_match.start()
        subject_table_text = pdf_text[start_pos:end_pos]
        
        # Split into lines and clean up
        lines = [line.strip() for line in subject_table_text.split('\n') if line.strip()]
        
        # Parse based on the exact MSBTE format
        subjects = parse_msbte_format(lines)
            
    except Exception as e:
        logging.error(f"Error extracting subject table: {str(e)}")
    
    return subjects


def parse_msbte_format(lines):
    """
    Parse MSBTE marksheet format based on the actual interleaved structure.
    Each subject is followed by its own marks in a row.
    """
    subjects = []
    
    try:
        # The actual structure is:
        # Headers -> Subject1 -> Marks1 -> Subject2 -> Marks2 -> etc.
        
        if len(lines) < 25:
            return subjects
        
        # Find all subjects and their marks
        current_subject = None
        subject_marks = []
        all_subjects_data = []
        
        for i, line in enumerate(lines):
            # Check if this is a subject name
            if (line.isupper() and 
                len(line.split()) > 1 and 
                not re.match(r'^[\d\-\s]+$', line) and
                line not in ['MAX', 'OBT', 'TOTAL', 'THEORY', 'PRACTICALS', 'CREDITS', 'SLA', 'FA-TH', 'SA-TH', 'FA-PR', 'SA-PR', 'OBT MAX OBT']):
                
                # Save previous subject if exists
                if current_subject and subject_marks:
                    all_subjects_data.append({
                        'name': current_subject,
                        'marks': subject_marks.copy()
                    })
                
                # Start new subject
                current_subject = line
                subject_marks = []
            
            # Check if this is a mark
            elif re.match(r'^[\d\-\s]+$', line) and current_subject:
                clean_mark = parse_numeric(line.strip())
                subject_marks.append(clean_mark)
        
        # Save the last subject
        if current_subject and subject_marks:
            all_subjects_data.append({
                'name': current_subject,
                'marks': subject_marks.copy()
            })
        
        print(f"DEBUG: Found {len(all_subjects_data)} subjects with interleaved marks:")
        for i, subj_data in enumerate(all_subjects_data):
            print(f"  {i}: {subj_data['name']} - {len(subj_data['marks'])} marks")
        
        # Now process each subject's marks
        for subj_data in all_subjects_data:
            subject_name = subj_data['name']
            marks = subj_data['marks']
            
            subject_marks = {
                "subject_name": subject_name.strip(),
                "fa_th_max": None,
                "fa_th_obt": None,
                "sa_th_max": None,
                "sa_th_obt": None,
                "th_total_max": None,
                "th_total_obt": None,
                "fa_pr_max": None,
                "fa_pr_obt": None,
                "sa_pr_max": None,
                "sa_pr_obt": None,
                "sla_max": None,
                "sla_obt": None,
                "credits": None
            }
            
            # Map marks based on position (MSBTE standard order)
            # Expected order: FA-TH Max, FA-TH Obt, SA-TH Max, SA-TH Obt, 
            # TH Total Max, TH Total Obt, FA-PR Max, FA-PR Obt, 
            # SA-PR Max, SA-PR Obt, SLA Max, SLA Obt, Credits
            
            if len(marks) >= 2:
                subject_marks["fa_th_max"] = marks[0] if len(marks) > 0 else None
                subject_marks["fa_th_obt"] = marks[1] if len(marks) > 1 else None
            
            if len(marks) >= 4:
                subject_marks["sa_th_max"] = marks[2] if len(marks) > 2 else None
                subject_marks["sa_th_obt"] = marks[3] if len(marks) > 3 else None
            
            if len(marks) >= 6:
                subject_marks["th_total_max"] = marks[4] if len(marks) > 4 else None
                subject_marks["th_total_obt"] = marks[5] if len(marks) > 5 else None
            
            if len(marks) >= 8:
                subject_marks["fa_pr_max"] = marks[6] if len(marks) > 6 else None
                subject_marks["fa_pr_obt"] = marks[7] if len(marks) > 7 else None
            
            if len(marks) >= 10:
                subject_marks["sa_pr_max"] = marks[8] if len(marks) > 8 else None
                subject_marks["sa_pr_obt"] = marks[9] if len(marks) > 9 else None
            
            if len(marks) >= 12:
                subject_marks["sla_max"] = marks[10] if len(marks) > 10 else None
                subject_marks["sla_obt"] = marks[11] if len(marks) > 11 else None
            
            if len(marks) >= 13:
                subject_marks["credits"] = marks[12] if len(marks) > 12 else None
            
            subjects.append(subject_marks)
            
    except Exception as e:
        logging.error(f"Error parsing MSBTE format: {str(e)}")
    
    return subjects


def parse_column_format(subject_names, marks_lines):
    """
    Parse marks data in column format where marks are arranged vertically.
    """
    subjects = []
    
    try:
        # The MSBTE format has marks in columns
        # Each subject has marks for different components in separate columns
        
        # Group marks by subject (based on the actual PDF structure)
        # From the PDF, we can see the pattern:
        # FA-TH: 30, 30, 30 (max marks for 3 subjects)
        # FA-TH OBT: 024, 015, 026 (obtained marks for 3 subjects)
        # etc.
        
        num_subjects = len(subject_names)
        if num_subjects == 0:
            return subjects
        
        # Parse marks into a 2D array [component][subject]
        marks_matrix = []
        current_component = []
        
        for mark_line in marks_lines:
            if mark_line == '-':
                current_component.append(None)
            else:
                clean_mark = parse_numeric(mark_line)
                current_component.append(clean_mark)
            
            # When we have marks for all subjects, start new component
            if len(current_component) == num_subjects:
                marks_matrix.append(current_component)
                current_component = []
        
        # Add any remaining marks
        if current_component:
            marks_matrix.append(current_component)
        
        # Map marks to subjects based on MSBTE format
        # Expected order: FA-TH Max, FA-TH Obt, SA-TH Max, SA-TH Obt, 
        # TH Total Max, TH Total Obt, FA-PR Max, FA-PR Obt, SA-PR Max, SA-PR Obt, SLA Max, SLA Obt, Credits
        
        for i, subject_name in enumerate(subject_names):
            subject_marks = {
                "subject_name": subject_name,
                "fa_th_max": None,
                "fa_th_obt": None,
                "sa_th_max": None,
                "sa_th_obt": None,
                "th_total_max": None,
                "th_total_obt": None,
                "fa_pr_max": None,
                "fa_pr_obt": None,
                "sa_pr_max": None,
                "sa_pr_obt": None,
                "sla_max": None,
                "sla_obt": None,
                "credits": None
            }
            
            # Map marks from matrix to subject
            component_idx = 0
            if component_idx < len(marks_matrix) and i < len(marks_matrix[component_idx]):
                subject_marks["fa_th_max"] = marks_matrix[component_idx][i]
            component_idx += 1
            
            if component_idx < len(marks_matrix) and i < len(marks_matrix[component_idx]):
                subject_marks["fa_th_obt"] = marks_matrix[component_idx][i]
            component_idx += 1
            
            if component_idx < len(marks_matrix) and i < len(marks_matrix[component_idx]):
                subject_marks["sa_th_max"] = marks_matrix[component_idx][i]
            component_idx += 1
            
            if component_idx < len(marks_matrix) and i < len(marks_matrix[component_idx]):
                subject_marks["sa_th_obt"] = marks_matrix[component_idx][i]
            component_idx += 1
            
            if component_idx < len(marks_matrix) and i < len(marks_matrix[component_idx]):
                subject_marks["th_total_max"] = marks_matrix[component_idx][i]
            component_idx += 1
            
            if component_idx < len(marks_matrix) and i < len(marks_matrix[component_idx]):
                subject_marks["th_total_obt"] = marks_matrix[component_idx][i]
            component_idx += 1
            
            if component_idx < len(marks_matrix) and i < len(marks_matrix[component_idx]):
                subject_marks["fa_pr_max"] = marks_matrix[component_idx][i]
            component_idx += 1
            
            if component_idx < len(marks_matrix) and i < len(marks_matrix[component_idx]):
                subject_marks["fa_pr_obt"] = marks_matrix[component_idx][i]
            component_idx += 1
            
            if component_idx < len(marks_matrix) and i < len(marks_matrix[component_idx]):
                subject_marks["sa_pr_max"] = marks_matrix[component_idx][i]
            component_idx += 1
            
            if component_idx < len(marks_matrix) and i < len(marks_matrix[component_idx]):
                subject_marks["sa_pr_obt"] = marks_matrix[component_idx][i]
            component_idx += 1
            
            if component_idx < len(marks_matrix) and i < len(marks_matrix[component_idx]):
                subject_marks["sla_max"] = marks_matrix[component_idx][i]
            component_idx += 1
            
            if component_idx < len(marks_matrix) and i < len(marks_matrix[component_idx]):
                subject_marks["sla_obt"] = marks_matrix[component_idx][i]
            component_idx += 1
            
            if component_idx < len(marks_matrix) and i < len(marks_matrix[component_idx]):
                subject_marks["credits"] = marks_matrix[component_idx][i]
            
            subjects.append(subject_marks)
            
    except Exception as e:
        logging.error(f"Error parsing column format: {str(e)}")
    
    return subjects


def process_marks_data_improved(subject_marks, marks_data):
    """
    Improved processing of marks data based on actual MSBTE format.
    """
    try:
        # Filter out empty strings and clean the data
        clean_marks = [mark for mark in marks_data if mark and mark.strip() != '']
        
        if len(clean_marks) < 2:
            return
        
        # The MSBTE format typically has marks in columns
        # We need to map them correctly based on the actual structure
        
        # For subjects with all components: FA-TH, SA-TH, TH Total, FA-PR, SA-PR, SLA, Credits
        # Each component has MAX and OBT values
        
        idx = 0
        
        # FA-TH (Max, Obt)
        if idx < len(clean_marks):
            subject_marks["fa_th_max"] = parse_numeric(clean_marks[idx])
            idx += 1
        if idx < len(clean_marks):
            subject_marks["fa_th_obt"] = parse_numeric(clean_marks[idx])
            idx += 1
        
        # SA-TH (Max, Obt)
        if idx < len(clean_marks):
            subject_marks["sa_th_max"] = parse_numeric(clean_marks[idx])
            idx += 1
        if idx < len(clean_marks):
            subject_marks["sa_th_obt"] = parse_numeric(clean_marks[idx])
            idx += 1
        
        # TH Total (Max, Obt)
        if idx < len(clean_marks):
            subject_marks["th_total_max"] = parse_numeric(clean_marks[idx])
            idx += 1
        if idx < len(clean_marks):
            subject_marks["th_total_obt"] = parse_numeric(clean_marks[idx])
            idx += 1
        
        # FA-PR (Max, Obt)
        if idx < len(clean_marks):
            subject_marks["fa_pr_max"] = parse_numeric(clean_marks[idx])
            idx += 1
        if idx < len(clean_marks):
            subject_marks["fa_pr_obt"] = parse_numeric(clean_marks[idx])
            idx += 1
        
        # SA-PR (Max, Obt)
        if idx < len(clean_marks):
            subject_marks["sa_pr_max"] = parse_numeric(clean_marks[idx])
            idx += 1
        if idx < len(clean_marks):
            subject_marks["sa_pr_obt"] = parse_numeric(clean_marks[idx])
            idx += 1
        
        # SLA (Max, Obt)
        if idx < len(clean_marks):
            subject_marks["sla_max"] = parse_numeric(clean_marks[idx])
            idx += 1
        if idx < len(clean_marks):
            subject_marks["sla_obt"] = parse_numeric(clean_marks[idx])
            idx += 1
        
        # Credits (usually the last value)
        if idx < len(clean_marks):
            subject_marks["credits"] = parse_numeric(clean_marks[idx])
                
    except Exception as e:
        logging.error(f"Error processing marks data: {str(e)}")


def process_marks_data(subject_marks, marks_data):
    """
    Process the collected marks data and populate subject_marks dictionary.
    """
    try:
        # Expected pattern for MSBTE marks:
        # FA-TH MAX, FA-TH OBT, SA-TH MAX, SA-TH OBT, TH TOTAL MAX, TH TOTAL OBT,
        # FA-PR MAX, FA-PR OBT, SA-PR MAX, SA-PR OBT, SLA MAX, SLA OBT, Credits
        
        if len(marks_data) >= 12:
            idx = 0
            
            # FA-TH marks
            if idx < len(marks_data):
                subject_marks["fa_th_max"] = parse_numeric(marks_data[idx])
                idx += 1
            if idx < len(marks_data):
                subject_marks["fa_th_obt"] = parse_numeric(marks_data[idx])
                idx += 1
            
            # SA-TH marks
            if idx < len(marks_data):
                subject_marks["sa_th_max"] = parse_numeric(marks_data[idx])
                idx += 1
            if idx < len(marks_data):
                subject_marks["sa_th_obt"] = parse_numeric(marks_data[idx])
                idx += 1
            
            # TH Total marks
            if idx < len(marks_data):
                subject_marks["th_total_max"] = parse_numeric(marks_data[idx])
                idx += 1
            if idx < len(marks_data):
                subject_marks["th_total_obt"] = parse_numeric(marks_data[idx])
                idx += 1
            
            # FA-PR marks
            if idx < len(marks_data):
                subject_marks["fa_pr_max"] = parse_numeric(marks_data[idx])
                idx += 1
            if idx < len(marks_data):
                subject_marks["fa_pr_obt"] = parse_numeric(marks_data[idx])
                idx += 1
            
            # SA-PR marks
            if idx < len(marks_data):
                subject_marks["sa_pr_max"] = parse_numeric(marks_data[idx])
                idx += 1
            if idx < len(marks_data):
                subject_marks["sa_pr_obt"] = parse_numeric(marks_data[idx])
                idx += 1
            
            # SLA marks
            if idx < len(marks_data):
                subject_marks["sla_max"] = parse_numeric(marks_data[idx])
                idx += 1
            if idx < len(marks_data):
                subject_marks["sla_obt"] = parse_numeric(marks_data[idx])
                idx += 1
            
            # Credits (usually last)
            if idx < len(marks_data):
                subject_marks["credits"] = parse_numeric(marks_data[idx])
                
    except Exception as e:
        logging.error(f"Error processing marks data: {str(e)}")


def parse_numeric(value):
    """
    Convert string value to integer or return None.
    Handles '-' and other non-numeric values.
    """
    try:
        if value == '-' or value == '' or value is None:
            return None
        
        # Remove any non-digit characters
        clean_value = re.sub(r'[^\d]', '', str(value))
        
        if clean_value == '':
            return None
            
        return int(clean_value)
    except:
        return None


# Function to parse marksheet data
def parse_marksheet(text):
    result_data = {
        "Student Name": safe_search(
            r"MR\. / MS\.\s*([\w\s]+)",
            text.replace("\n\n   ENROLLMENT NO", "").replace(
                "\n\nSTATEMENT OF MARKS", ""
            ),
        ),
        "Enrollment No": safe_search(r"ENROLLMENT NO\.\s*(\d+)", text),
        "Examination": safe_search(r"EXAMINATION\s*([A-Z]+\s+\d+)", text),
        "Seat No": safe_search(r"SEAT NO\.\s*(\d+)", text),
        "Semester": safe_search(
            r"(\bFIRST\b|\bSECOND\b|\bTHIRD\b|\bFOURTH\b|\bFIFTH\b|\bSIXTH\b)\s+SEMESTER",
            text,
        ),
        "Percentage": extract_by_line_offset(text, "PERCENTAGE", 9),
        "Gain Marks": extract_by_line_offset(text, "PERCENTAGE", 7),
        "Total Marks": extract_by_line_offset(text, "PERCENTAGE", 5),
        "Total Credits": extract_by_line_offset(text, "TOTAL CREDIT", 8),
        # "Subjects": extract_subjects(text), # Function to Extract subjects
    }

    # Map Semester to Year
    semester_mapping = {
        "FIRST": "First Year",
        "SECOND": "First Year",
        "THIRD": "Second Year",
        "FOURTH": "Second Year",
        "FIFTH": "Third Year",
        "SIXTH": "Third Year",
    }

    result_data["Student Year"] = semester_mapping.get(
        result_data["Semester"], "Unknown"
    )

    # Default Result
    result_data["Result"] = "FAIL"

    # Check if percentage is valid
    if result_data["Percentage"]:
        percentage_str = result_data["Percentage"].replace("%", "").strip()

        if percentage_str.isdigit() or (
            "." in percentage_str and percentage_str.replace(".", "").isdigit()
        ):
            percentage = float(percentage_str)

            # Assign Result based on MSBTE scale
            if percentage < 40:
                result_data["Result"] = "FAIL"
                result_data["Percentage"] = None  # Hide percentage for failed students
                result_data["Total Credits"] = extract_by_line_offset(
                    text, "TOTAL CREDIT", 6
                )
            elif 40 <= percentage < 45:
                result_data["Result"] = "PASS"
            elif 45 <= percentage < 60:
                result_data["Result"] = "SECOND CLASS"
            elif 60 <= percentage < 75:
                result_data["Result"] = "FIRST CLASS"
            else:
                result_data["Result"] = "FIRST CLASS DIST"

    return result_data


# Function to save data to Excel with two sheets
def save_to_excel(data_list, output_file):
    """
    Save student data to Excel with two sheets:
    Sheet 1: Student Summary (existing format)
    Sheet 2: Subject Marks (new format)
    """
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Sheet 1: Student Summary (unchanged format, but without subjects field)
        summary_data = []
        for student_data in data_list:
            # Create a copy without the subjects field for the summary sheet
            summary_row = student_data.copy()
            if 'subjects' in summary_row:
                del summary_row['subjects']
            summary_data.append(summary_row)
        
        df_summary = pd.DataFrame(summary_data)
        
        # Ensure sorting columns exist before applying sort
        sort_columns = [col for col in ["Percentage", "Gain Marks"] if col in df_summary.columns]
        
        if sort_columns:
            df_summary = df_summary.sort_values(by=sort_columns, ascending=False)
        
        df_summary.to_excel(writer, sheet_name="Student Summary", index=False)
        
        # Sheet 2: Subject Marks
        subject_rows = []
        
        for student_data in data_list:
            enrollment_no = student_data.get("Enrollment No", "")
            semester = student_data.get("Semester", "")
            
            # Get subject data if available
            subjects = student_data.get("subjects", [])
            
            if subjects:
                for subject in subjects:
                    subject_row = {
                        "Enrollment No": enrollment_no,
                        "Semester": semester,
                        "Subject Name": subject.get("subject_name", ""),
                        "FA-TH Max": subject.get("fa_th_max"),
                        "FA-TH Obt": subject.get("fa_th_obt"),
                        "SA-TH Max": subject.get("sa_th_max"),
                        "SA-TH Obt": subject.get("sa_th_obt"),
                        "TH Total Max": subject.get("th_total_max"),
                        "TH Total Obt": subject.get("th_total_obt"),
                        "FA-PR Max": subject.get("fa_pr_max"),
                        "FA-PR Obt": subject.get("fa_pr_obt"),
                        "SA-PR Max": subject.get("sa_pr_max"),
                        "SA-PR Obt": subject.get("sa_pr_obt"),
                        "SLA Max": subject.get("sla_max"),
                        "SLA Obt": subject.get("sla_obt"),
                        "Credits": subject.get("credits")
                    }
                    subject_rows.append(subject_row)
            else:
                # If no subjects extracted, still create a placeholder row
                subject_row = {
                    "Enrollment No": enrollment_no,
                    "Semester": semester,
                    "Subject Name": "No subject data available",
                    "FA-TH Max": None,
                    "FA-TH Obt": None,
                    "SA-TH Max": None,
                    "SA-TH Obt": None,
                    "TH Total Max": None,
                    "TH Total Obt": None,
                    "FA-PR Max": None,
                    "FA-PR Obt": None,
                    "SA-PR Max": None,
                    "SA-PR Obt": None,
                    "SLA Max": None,
                    "SLA Obt": None,
                    "Credits": None
                }
                subject_rows.append(subject_row)
        
        if subject_rows:
            df_subjects = pd.DataFrame(subject_rows)
            df_subjects.to_excel(writer, sheet_name="Subject Marks", index=False)
        else:
            # Create empty sheet with headers if no subject data
            empty_df = pd.DataFrame(columns=[
                "Enrollment No", "Semester", "Subject Name",
                "FA-TH Max", "FA-TH Obt", "SA-TH Max", "SA-TH Obt",
                "TH Total Max", "TH Total Obt", "FA-PR Max", "FA-PR Obt",
                "SA-PR Max", "SA-PR Obt", "SLA Max", "SLA Obt", "Credits"
            ])
            empty_df.to_excel(writer, sheet_name="Subject Marks", index=False)


# Allow only PDF files
def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() == "pdf"


# Route for the home page
@app.route("/")
def home():
    return render_template("home.html")


# Upload route to handle multiple PDFs
@app.route("/upload", methods=["GET", "POST"])
def upload():
    if request.method == "POST":
        if "files" not in request.files:
            return render_template("upload.html", message="No file part")

        files = request.files.getlist("files")

        if not files or all(file.filename == "" for file in files):
            return render_template("upload.html", message="No selected files")

        extracted_data = []

        for file in files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                filepath = os.path.join(UPLOAD_FOLDER, filename)
                file.save(filepath)
                pdf_text = extract_text_from_pdf(filepath)
                parsed_data = parse_marksheet(pdf_text)
                
                # Extract subject-wise marks
                try:
                    subjects = extract_subject_table(pdf_text)
                    parsed_data["subjects"] = subjects
                    logging.info(f"Extracted {len(subjects)} subjects for enrollment {parsed_data.get('Enrollment No', 'Unknown')}")
                except Exception as e:
                    logging.warning(f"Failed to extract subjects for enrollment {parsed_data.get('Enrollment No', 'Unknown')}: {str(e)}")
                    parsed_data["subjects"] = []
                
                extracted_data.append(parsed_data)

                # Delete uploaded PDF after processing
                if os.path.exists(filepath):
                    os.remove(filepath)

        excel_output_file = os.path.join(UPLOAD_FOLDER, "output.xlsx")
        save_to_excel(extracted_data, excel_output_file)

        return render_template("download.html", excel_file="output.xlsx")

    return render_template("upload.html")


# Download route
@app.route("/download/<filename>")
def download_file(filename):
    return send_file(
        os.path.join(app.config["UPLOAD_FOLDER"], filename), as_attachment=True
    )


# Test route to process a single predefined PDF
@app.route("/test")
def test():
    file_path = os.path.join(UPLOAD_FOLDER, "result.pdf")

    if not os.path.exists(file_path):
        return jsonify({"error": f"File '{file_path}' not found"}), 404

    pdf_text = extract_text_from_pdf(file_path)
    parsed_data = parse_marksheet(pdf_text)
    
    # Extract subject-wise marks
    try:
        subjects = extract_subject_table(pdf_text)
        parsed_data["subjects"] = subjects
        logging.info(f"Extracted {len(subjects)} subjects for enrollment {parsed_data.get('Enrollment No', 'Unknown')}")
    except Exception as e:
        logging.warning(f"Failed to extract subjects for enrollment {parsed_data.get('Enrollment No', 'Unknown')}: {str(e)}")
        parsed_data["subjects"] = []
    
    excel_output_file = os.path.join(UPLOAD_FOLDER, "output.xlsx")
    save_to_excel([parsed_data], excel_output_file)

    return jsonify(parsed_data)


if __name__ == "__main__":
   app.run(host="0.0.0.0", port=5000, debug=True)
