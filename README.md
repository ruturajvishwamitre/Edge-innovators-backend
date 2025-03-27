# MSBTE Marksheet Extractor

![Python](https://img.shields.io/badge/Python-3.8%2B-blue)
![Flask](https://img.shields.io/badge/Flask-2.0%2B-green)
![Pandas](https://img.shields.io/badge/Pandas-1.3%2B-orange)
![PDFMiner](https://img.shields.io/badge/PDFMiner-20201018-lightgrey)

**MSBTE Marksheet Extractor** is a powerful and user-friendly tool designed to automate the extraction and analysis of student marksheet data from PDF files provided by the **Maharashtra State Board of Technical Education (MSBTE)**. This application is built using Python, Flask, and Pandas, making it easy to upload, process, and export marksheet data into a structured Excel format.

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
- **Scalable**: Process multiple marksheets simultaneously.

---

## Why Use MSBTE Marksheet Extractor?

- **Save Time**: Automate the tedious task of manually extracting data from PDF marksheets.
- **Accuracy**: Reduce human errors in data extraction and processing.
- **Insights**: Quickly analyze student performance and generate reports.
- **User-Friendly**: No technical expertise required—just upload, process, and download.

---

## Screenshots

### Home Page

![Home Page](/Static/Images/home.png)
_The home page provides an overview of the tool and instructions for use._

### Upload Page

![Upload Page](/Static/Images/upload.png)
_Upload your MSBTE marksheet PDFs here._

### Download Page

![Download Page](/Static/Images/download.png)
_Download the parsed data in Excel format._

---

## Installation

### Prerequisites

- Python 3.8 or higher
- Pip (Python package manager)

### Steps

1. Clone the repository:

   ```bash
   git clone https://github.com/your-username/MSBTE-Marksheet-Extractor.git
   cd MSBTE-Marksheet-Extractor
   ```

2. Create a virtual environment (optional but recommended):

   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

4. Run the application:

   ```bash
   python app.py
   ```

5. Open your browser and navigate to:
   ```
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

## Contributing

Contributions are welcome! If you'd like to contribute, please follow these steps:

1. Fork the repository.
2. Create a new branch (`git checkout -b feature/YourFeatureName`).
3. Commit your changes (`git commit -m 'Add some feature'`).
4. Push to the branch (`git push origin feature/YourFeatureName`).
5. Open a pull request.

---

## License

This project is licensed under the **MIT License**. See the [LICENSE](LICENSE) file for details.

---

## Acknowledgments

- Thanks to the developers of **PDFMiner**, **Flask**, and **Pandas** for their amazing libraries.
- Inspired by the need to automate marksheet processing in educational institutions.

---

## Contact

For questions or feedback, feel free to reach out:

- **Krushna Kale**: [krushnakaale@gmail.com](mailto:your.email@example.com)
- **GitHub**: [https://github.com/krushnakaale](https://github.com/your-username)

---
