# Transcript PDF to Excel Converter

A desktop application for converting text-based PDF university transcripts into a structured Excel workbook.

## Excel sheets created
- **Student Info** — registration number, name, DOB, program, transcript date, credit-hour totals, and CGPA.
- **Courses** — course code, title, credits, grade, section (including deficiency courses), and semester.
- **Semester Summary** — semester credits/GPA and cumulative credits/GPA where present in the transcript.
- **Conversion Log** — conversion status and warnings for each PDF.

## Requirements
- Python 3.10 or newer
- Windows, Linux, or macOS
- Text-based PDFs (PDFs with selectable text)

## Install and run

### Windows
1. Install Python from https://www.python.org/downloads/ and enable **Add Python to PATH**.
2. Extract this ZIP file.
3. Open Command Prompt in the extracted folder and run:
   ```bat
   py -m venv .venv
   .venv\Scripts\activate
   pip install -r requirements.txt
   python transcript_to_excel.py
   ```

### Linux
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python3 transcript_to_excel.py
```

On Linux Mint, if Tkinter is missing:
```bash
sudo apt install python3-tk
```

## Use
1. Click **Select PDF files** and choose one or more transcripts.
2. Choose the output `.xlsx` file.
3. Click **Convert to Excel**.
4. Open the workbook and verify the extracted values before using them for official records.

## Important limitations
- Scanned/image-only PDFs are not supported by this version; OCR can be added as a next step.
- Transcript layouts vary. Check the Conversion Log and verify student details, grades, course credits, and GPA figures.
- The app does not calculate or alter grades/GPA; it extracts the values printed in the transcript.
