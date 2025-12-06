# MSBTE Marksheet Analyzer

![Python](https://img.shields.io/badge/Python-3.8%2B-blue)
![Flask](https://img.shields.io/badge/Flask-2.0%2B-green)
![Pandas](https://img.shields.io/badge/Pandas-1.3%2B-orange)
![PDFMiner](https://img.shields.io/badge/PDFMiner-20201018-lightgrey)

**MSBTE Marksheet Analyzer** is a powerful and user-friendly tool designed to automate the extraction and analysis of student marksheet data from PDF files provided by the Maharashtra State Board of Technical Education (MSBTE). This application is built using Python, Flask, and Pandas, making it easy to upload, process, and export marksheet data into a structured Excel format.

Whether you're an educational institution, a teacher, or a student, this tool simplifies the process of extracting key details such as student names, enrollment numbers, semester results, percentages, and more. It also provides insights into student performance by categorizing results into classes (e.g., FIRST CLASS, SECOND CLASS) and identifying pass/fail statuses.

---

## Key Features

- **PDF Text Extraction**: Extract text from MSBTE marksheet PDFs using **PDFMiner**.
- **Data Parsing**: Automatically parse and organize key details such as:

  - Student Name
  - Enrollment Number
  - Semester
  - Percentage
  - Total Marks
  - Result Status (PASS/FAIL, Class, etc.)

- **Excel Export**: Export parsed data into a structured Excel file for further analysis or record-keeping.
- **Web Interface**: Simple and intuitive web interface for uploading PDFs and downloading results.
- **Error Handling**: Gracefully handles invalid or unsupported files.
- **Batch Processing**: Process multiple marksheets simultaneously.

---

## Why Use MSBTE Marksheet Analyzer?

- **Save Time**: Automate the tedious task of manually extracting data from PDF marksheets.
- **Accuracy**: Reduce human errors in data extraction and processing.
- **Insights**: Quickly analyze student performance and generate reports.
- **User-Friendly**: No technical expertise required — just upload, process, and download.

---

## Screenshots

### Home Page

![Home Page](/static/images/home.png)

_The home page provides an overview of the tool and instructions for use._

### Upload Page

![Upload Page](/static/images/upload.png)

_Upload your MSBTE marksheet PDFs here._

### Download Page

![Download Page](/static/images/download.png)

_Download the parsed data in Excel format._

---

## Installation

### Prerequisites

- Python 3.8 or higher
- Pip (Python package manager)

### Steps

1. Clone the repository:

   ```bash
   git clone https://github.com/krushnakaale/MSBTE-Marksheet-Analyzer.git
   cd MSBTE-Marksheet-Analyzer
   ```

````


2. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

3. Run the application:

   ```bash
   python app.py
   ```

4. Open your browser and navigate to:

   ```plaintext
   http://127.0.0.1:5000
   ```

---

## Usage

1. **Upload PDFs**:

   - Go to the home page and click "Upload PDFs".
   - Select one or more MSBTE marksheet PDF files and upload them.

2. **Process Data**:

   - The application will extract and parse the data from the PDFs.

3. **Download Excel**:

   - After processing, click the download link to get the Excel file containing the parsed data.

---

## API Endpoints

| Endpoint           | Method | Description                             |
| ------------------ | ------ | --------------------------------------- |
| `/`                | GET    | Home page with upload instructions.     |
| `/upload`          | POST   | Upload PDF files for processing.        |
| `/download/<file>` | GET    | Download the generated Excel file.      |
| `/test`            | GET    | Test route for processing a sample PDF. |

---

## Technologies Used

- **Python**: Core programming language.
- **Flask**: Web framework for building the application.
- **PDFMiner**: Library for extracting text from PDFs.
- **Pandas**: Data manipulation and Excel file generation.
- **HTML/CSS**: Frontend templates for the web interface.
- **Bootstrap**: Styling for the web pages.

---
````
