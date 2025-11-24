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


# Function to save data to Excel with sorting
def save_to_excel(data_list, output_file):
    df = pd.DataFrame(data_list)

    # Ensure sorting columns exist before applying sort
    sort_columns = [col for col in ["Percentage", "Gain Marks"] if col in df.columns]

    if sort_columns:
        df = df.sort_values(by=sort_columns, ascending=False)

    df.to_excel(output_file, index=False)


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
    excel_output_file = os.path.join(UPLOAD_FOLDER, "output.xlsx")
    save_to_excel([parsed_data], excel_output_file)

    return jsonify(parsed_data)


if __name__ == "__main__":
    app.run(debug=True)
    app.run(host="0.0.0.0", port=5000, debug=True)
