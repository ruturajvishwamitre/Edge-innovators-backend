import os
import re
import pandas as pd
from flask import Flask, request, jsonify
from pdfminer.high_level import extract_text

app = Flask(__name__)

# Add current folder path
UPLOAD_FOLDER = "./MSBTE-Marksheet-Extractor/Uploads"

def extract_text_from_pdf(file_path):
    return extract_text(file_path)

def safe_search(pattern, text, default="N/A"):
    """Helper function to safely extract regex matches"""
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(1).strip() if match else default

def extract_by_line_offset(text, keyword, offset, default="N/A"):
    """Extract value based on a keyword and line offset"""
    lines = text.split("\n")
    for i, line in enumerate(lines):
        if keyword in line:
            target_line = i + offset
            if 0 <= target_line < len(lines):
                match = re.search(r"(\d+)", lines[target_line])  # Extracts first numeric value
                return match.group(1) if match else default
    return default

def extract_subjects(text):
    """Extracts subject names from the marksheet text and removes unwanted entries."""
    
    # Initialize subjects list to store results from both conditions
    subjects = []

    # First condition - Extract using "TITLE OF COURSES" pattern
    match1 = re.search(r"TITLE OF COURSES\s*TH/\s*PR\s*COURSE\s*HEAD\s*(.*?)\s*ESE", text, re.DOTALL)
    if match1:
        subjects_raw = match1.group(1).strip()
        subjects += [s.strip() for s in re.split(r"\n+", subjects_raw) if s.strip() and not re.match(r"^(TH|PR|TW|ESE|PA|OR|PJ|IT|LSP|AG|@|PLY|AP|\*|WFLS|%|WFLY|CON|FT|A.T.K.T)$", s)]
    
    # Second condition - Extract using "TITLE OF SUBJECTS" pattern
    match2 = re.search(r"TITLE OF SUBJECTS\s*(.*?)\s*(?=TOTAL MARKS|RESULT DECLARED|INSTRUCTIONS)", text, re.DOTALL)
    if match2:
        subjects_raw = match2.group(1).strip()
        subjects += [s.strip() for s in re.split(r"\n+", subjects_raw) if s.strip()]

        # Remove unwanted terms from the subjects
        unwanted_terms = [
            "THEORY", "PRACTICALS", "TOTAL", "CREDITS", "MAX", "OBT", "SLA", 
            "FA", "SA", "TH", "PR", "TW", "ESE", "PA", "OR", "PJ", "IT", 
            "LSP", "AG", "@", "PLY", "*", "WFLS", "%", "WFLY", "CON", "A.T.K.T",
            "OBT MAX OBT", "DATE :", "This Marksheet is Downloaded from Internet", 
            "MAHARASHTRA STATE BOARD OF TECHNICAL EDUCATION", "TOTAL MAX."
        ]
        
        # Remove any unwanted term or numeric entries
        subjects = [
            subject for subject in subjects if subject not in unwanted_terms and not subject.isdigit() and len(subject.split()) > 1
        ]
    
    return subjects


def parse_marksheet(text):
    """Parses the given marksheet text and extracts structured data."""
    
    result_data = {
        "Student Name": safe_search(r"MR\. / MS\.\s*([\w\s]+)", text.replace('\n\n   ENROLLMENT NO','').replace('\n\nSTATEMENT OF MARKS','')),  
        "Enrollment No": safe_search(r"ENROLLMENT NO\.\s*(\d+)", text),
        "Examination": safe_search(r"EXAMINATION\s*([A-Z]+\s+\d+)", text),
        # "Year": safe_search(r"EXAMINATION\s*[A-Z]+\s+(\d{4})", text),  # Extracts the year (e.g., 2023)
        "Seat No": safe_search(r"SEAT NO\.\s*(\d+)", text),
        "Semester": safe_search(r"(\bFIRST\b|\bSECOND\b|\bTHIRD\b|\bFOURTH\b|\bFIFTH\b|\bSIXTH\b)\s+SEMESTER", text),
        "Percentage": extract_by_line_offset(text, "PERCENTAGE", 9),
        "Gain Marks": extract_by_line_offset(text, "PERCENTAGE", 7),
        "Total Marks": extract_by_line_offset(text, "PERCENTAGE", 5),
        "Total Credits": extract_by_line_offset(text, "TOTAL CREDIT", 8),
        "Subjects":extract_subjects(text),
    }
            
    # Map Semester to Year
    semester_mapping = {
        "FIRST": "First Year", "SECOND": "First Year",
        "THIRD": "Second Year", "FOURTH": "Second Year",
        "FIFTH": "Third Year", "SIXTH": "Third Year"
    }
    
    result_data["Student Year"] = semester_mapping.get(result_data["Semester"], "Unknown")
    
    # Default Result
    result_data["Result"] = "FAIL"

    # Check if percentage is valid
    if result_data["Percentage"]:
        percentage_str = result_data["Percentage"].replace('%', '').strip()

        if percentage_str.isdigit() or ('.' in percentage_str and percentage_str.replace('.', '').isdigit()):
            percentage = float(percentage_str)

            # Assign Result based on MSBTE scale
            if percentage < 40:
                result_data["Result"] = "FAIL"

                # New condition: Set "Percentage" to None if below 40%
                result_data["Percentage"] = "None"  # or result_data["Percentage"] = ""

                # Update "Total Credits" at line 10 if failed
                result_data["Total Credits"] = extract_by_line_offset(text, "TOTAL CREDIT", 6)
            
            elif 40 <= percentage < 45:
                result_data["Result"] = "PASS"
            elif 45 <= percentage < 60:
                result_data["Result"] = "SECOND CLASS"
            elif 60 <= percentage < 75:
                result_data["Result"] = "FIRST CLASS"
            else:
                result_data["Result"] = "FIRST CLASS DIST"
    
    return result_data


def save_to_excel(data, output_file):
    df = pd.DataFrame([data])
    df.to_excel(output_file, index=False)
    
@app.route('/test')
def test():
    file_path = os.path.join(UPLOAD_FOLDER, "copy_1.pdf")

    if not os.path.exists(file_path):
        return jsonify({"error": f"File '{file_path}' not found"}), 404

    pdf_text = extract_text_from_pdf(file_path)
    
    # Save the extracted text o pdf.txt in the root directory
    with open('pdf.txt', 'w', encoding='utf-8') as output_file:
        output_file.write(pdf_text)
        
    parsed_data = parse_marksheet(pdf_text)

    excel_output_file = os.path.join(UPLOAD_FOLDER, "output.xlsx")
    save_to_excel(parsed_data, excel_output_file)

    return jsonify(parsed_data)

if __name__ == '__main__':
    app.run(debug=True)

