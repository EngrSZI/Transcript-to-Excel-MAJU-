from pathlib import Path
import re
import json
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog

import fitz
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

APP_TITLE = "MAJU Transcript PDF to Excel"

def extract_pdf_text(path):
    with fitz.open(path) as pdf:
        return "\n".join(page.get_text("text") for page in pdf)

def nonempty_lines(text):
    return [re.sub(r"\s+", " ", x).strip() for x in text.splitlines() if x.strip()]

def extract_student_info(text, filename):
    """
    Handles MAJU transcript layout where the five column headers are followed
    by five separate values:
    Registration No / Name / DOB / Program / Current Date
    then registration number / student name / DOB / program / transcript date.
    """
    lines = nonempty_lines(text)
    result = {
        "Source File": filename,
        "Registration No": "",
        "Student Name": "",
        "Date of Birth": "",
        "Program": "",
        "Transcript Date": "",
        "Credit Hours Attempted": "",
        "Credit Hours Earned": "",
        "CGPA": "",
        "Credit Hours Accepted Towards Degree": "",
        "Passed Courses Count": 0,
        "Passed Credit Hours Toward Degree": 0,
        "Core Passed Courses Count": 0,
        "Core Passed Credit Hours": 0,
        "Elective Passed Courses Count": 0,
        "Elective Passed Credit Hours": 0,
        "Thesis Passed Courses Count": 0,
        "Thesis Passed Credit Hours": 0,
        "Deficiency Passed Courses Count": 0,
        "Deficiency Passed Credit Hours": 0,
        "Thesis I Passed": "No",
        "Thesis II Passed": "No",
        "Software Engineering Domain Electives Passed": 0,
        "AI & Data Science Domain Electives Passed": 0,
        "Network and Information Security Domain Electives Passed": 0,
        "MSCS Specialization Awarded": "Not evaluated",
        "MSCS Specialization Eligibility Details": "",
    }

    # Most reliable: match the ordered header group, then read the next five lines.
    expected = ["registration no", "name", "dob", "program", "current date"]
    for i in range(len(lines) - 4):
        if [v.casefold().rstrip(":") for v in lines[i:i+5]] == expected:
            values = lines[i+5:i+10]
            if len(values) == 5:
                result["Registration No"] = values[0]
                result["Student Name"] = values[1]
                result["Date of Birth"] = values[2]
                result["Program"] = values[3]
                result["Transcript Date"] = values[4]
                break

    # Fallback: find registration number and use the nearby sequence of values.
    reg_pattern = re.compile(r"^[A-Z]{2}\d{2}-[A-Z]+-\d{4}$", re.I)
    if not result["Registration No"]:
        for i, line in enumerate(lines):
            if reg_pattern.fullmatch(line):
                result["Registration No"] = line
                # In the source layout the name follows the registration number.
                if i + 1 < len(lines) and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", lines[i+1]):
                    result["Student Name"] = lines[i+1]
                break

    # Also handle PDFs that flatten the header and values onto one line.
    if not result["Registration No"] or not result["Student Name"]:
        flat = re.sub(r"\s+", " ", text)
        m = re.search(
            r"(FA|SP|SU|WI)?\d{2}-[A-Z]+-\d{4}\s+"
            r"([A-Za-z][A-Za-z .'-]+?)\s+"
            r"(\d{4}-\d{2}-\d{2})\s+"
            r"(.*?)\s+(\d{4}-\d{2}-\d{2})",
            flat, re.I
        )
        if m:
            result["Registration No"] = result["Registration No"] or m.group(0).split()[0]
            result["Student Name"] = result["Student Name"] or m.group(2).strip()
            result["Date of Birth"] = result["Date of Birth"] or m.group(3)
            result["Program"] = result["Program"] or m.group(4).strip()
            result["Transcript Date"] = result["Transcript Date"] or m.group(5)

    # Summary fields appear as label then value on the next line in this PDF.
    labels = {
        "credit hours attempted": "Credit Hours Attempted",
        "credit hours earned": "Credit Hours Earned",
        "cumulative grade point average": "CGPA",
        "credit hours accepted towards degree completion": "Credit Hours Accepted Towards Degree",
    }
    for i, line in enumerate(lines[:-1]):
        key = labels.get(line.casefold().rstrip(":"))
        if key:
            result[key] = lines[i+1]

    return result


CATEGORIES = ["Core", "Elective", "Thesis", "Deficiency"]
def get_memory_file():
    # Store user decisions outside the application folder so they survive app updates
    # and running the application from a different extracted directory.
    appdata = os.environ.get("APPDATA")
    if appdata:
        memory_dir = Path(appdata) / "MAJUTranscriptToExcel"
    else:
        memory_dir = Path.home() / ".maju_transcript_to_excel"
    memory_dir.mkdir(parents=True, exist_ok=True)
    return memory_dir / "course_category_memory.json"

MEMORY_FILE = get_memory_file()

def normalize_course_title(value):
    value = (value or "").replace("\u00a0", " ").replace("\u202f", " ")
    value = value.casefold().replace("&", " and ")
    value = re.sub(r"[^a-z0-9]+", " ", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value

def load_category_memory():
    try:
        data = json.loads(MEMORY_FILE.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {}

def save_category_memory(memory):
    MEMORY_FILE.write_text(json.dumps(memory, indent=2, ensure_ascii=False), encoding="utf-8")

class CategoryDialog(tk.Toplevel):
    def __init__(self, parent, program, course_code, course_title, source_section):
        super().__init__(parent)
        self.title("Categorize Unmatched Course")
        self.resizable(False, False)
        self.result = None
        self.transient(parent)
        self.grab_set()
        frame = ttk.Frame(self, padding=16)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="This course is not in the core-course rules.", font=("TkDefaultFont", 11, "bold")).pack(anchor="w")
        details = f"Program: {program or 'Unknown'}\nCourse: {course_code} — {course_title}\nTranscript section: {source_section}"
        ttk.Label(frame, text=details, wraplength=460, justify="left").pack(anchor="w", pady=(10, 14))
        ttk.Label(frame, text="Choose a category. Your choice will be remembered for future conversions.").pack(anchor="w", pady=(0, 8))
        self.choice = tk.StringVar(value="Elective")
        for category in CATEGORIES:
            ttk.Radiobutton(frame, text=category, value=category, variable=self.choice).pack(anchor="w", pady=2)
        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=(14, 0))
        ttk.Button(buttons, text="Save category", command=self.confirm).pack(side="right")
        ttk.Button(buttons, text="Cancel conversion", command=self.cancel).pack(side="right", padx=(0, 8))
        self.protocol("WM_DELETE_WINDOW", self.cancel)
        self.wait_window(self)

    def confirm(self):
        self.result = self.choice.get()
        self.destroy()

    def cancel(self):
        self.result = None
        self.destroy()

def parse_transcript(path):
    text = extract_pdf_text(path)
    lines = nonempty_lines(text)
    student = extract_student_info(text, Path(path).name)
    courses, semesters = [], []
    current_semester = ""
    section = ""
    course_code = re.compile(r"^[A-Z]{2,5}\d{3,5}$", re.I)
    grade = re.compile(r"^(A[+-]?|B[+-]?|C[+-]?|D[+-]?|F|W|\(F\)|\(W\)|I|P|NP|IP)$", re.I)
    semester_re = re.compile(r"^Semester\s+\d+\s*\[\s*(.*?)\s*\]$", re.I)

    i = 0
    while i < len(lines):
        line = lines[i]
        if line.casefold() == "deficiency courses":
            current_semester, section = "Deficiency Courses", "Deficiency Courses"
            i += 1
            continue
        sm = semester_re.match(line)
        if sm:
            current_semester, section = sm.group(1).strip(), "Semester Courses"
            i += 1
            continue

        if course_code.fullmatch(line) and i + 3 < len(lines):
            title, credits, grd = lines[i+1:i+4]
            if re.fullmatch(r"\d+(?:\.\d+)?", credits) and grade.fullmatch(grd):
                # Exclude courses with failing or withdrawn grades, including parenthesized
                # and lowercase variants: F, (F), W, and (W).
                normalized_grade = grd.strip().upper().strip("()")
                if normalized_grade not in {"F", "W"}:
                    courses.append({
                        "Source File": Path(path).name,
                        "Registration No": student["Registration No"],
                        "Student Name": student["Student Name"],
                        "Course Section": section or "Unclassified",
                        "Semester": current_semester or "Unspecified",
                        "Course Code": line,
                        "Course Title": title,
                        "Credit Hours": float(credits),
                        "Grade": grd.upper(),
                    })
                i += 4
                continue

        if line.casefold().startswith("total ") and i + 2 < len(lines):
            if line.casefold() == "total beginning to date":
                if semesters and re.fullmatch(r"\d+(?:\.\d+)?", lines[i+1]) and re.fullmatch(r"\d+(?:\.\d+)?", lines[i+2]):
                    semesters[-1]["Cumulative Credit Hours"] = float(lines[i+1])
                    semesters[-1]["Cumulative GPA"] = float(lines[i+2])
                i += 3
                continue
            if current_semester and current_semester != "Deficiency Courses":
                if re.fullmatch(r"\d+(?:\.\d+)?", lines[i+1]) and re.fullmatch(r"\d+(?:\.\d+)?", lines[i+2]):
                    semesters.append({
                        "Source File": Path(path).name,
                        "Registration No": student["Registration No"],
                        "Student Name": student["Student Name"],
                        "Semester": current_semester,
                        "Semester Credit Hours": float(lines[i+1]),
                        "Semester GPA": float(lines[i+2]),
                        "Cumulative Credit Hours": "",
                        "Cumulative GPA": "",
                    })
                    i += 3
                    continue
        i += 1
    return student, courses, semesters

def is_thesis_1(course):
    """Recognize Thesis I / Thesis 1 course titles."""
    title = normalize_course_title(course.get("Course Title", ""))
    return bool(re.search(r"\bthesis\b.*\b(?:i|1|one)\b", title))

def is_thesis_2(course):
    """Recognize a completed Thesis II / Thesis 2 course."""
    title = normalize_course_title(course.get("Course Title", ""))
    grade = (course.get("Grade", "") or "").strip().upper().strip("()")
    return (
        grade not in {"F", "W"}
        and bool(re.search(r"\bthesis\b.*\b(?:ii|2|two)\b", title))
    )

def apply_thesis_completion_rule(courses):
    """Drop Thesis I unless a completed (non-F/non-W) Thesis II is present."""
    eligible_courses = [
        course for course in courses
        if (course.get("Grade", "") or "").strip().upper().strip("()") not in {"F", "W"}
    ]
    has_thesis_2 = any(
        is_thesis_2(course) and is_passing_grade(course)
        for course in eligible_courses
    )
    if has_thesis_2:
        return eligible_courses, 0
    retained = [course for course in eligible_courses if not is_thesis_1(course)]
    return retained, len(eligible_courses) - len(retained)

def categorize_courses(parent, student, courses, memory, selected_program):
    program = selected_program
    for course in courses:
        title = course.get("Course Title", "")
        # The selected program is part of the key, so the same course can be Core in one
        # program and Elective in another.
        key = f"{program}|{normalize_course_title(title)}"
        course["Program"] = program

        # A course in the transcript's explicit deficiency section is a deficiency course.
        if course.get("Course Section") == "Deficiency Courses":
            course["Category"] = "Deficiency"
            continue

        # No hardcoded Core list is used. Each course is categorized per program by the user.
        if key in memory and memory[key] in CATEGORIES:
            course["Category"] = memory[key]
            continue

        dialog = CategoryDialog(parent, program, course.get("Course Code", ""),
                                title, course.get("Course Section", ""))
        if dialog.result is None:
            return False
        memory[key] = dialog.result
        course["Category"] = dialog.result
        save_category_memory(memory)

    # Deficiency courses must contribute zero credit hours in the exported course list,
    # regardless of the credit hours printed on the transcript.
    for course in courses:
        if course.get("Category") == "Deficiency":
            course["Credit Hours"] = 0

    # Ensure exported student program reflects the program selected for this analysis.
    student["Program"] = program
    return True

def is_passing_grade(course):
    """Return True for completed passing grades; incomplete/in-progress/non-pass do not count."""
    grade = (course.get("Grade", "") or "").strip().upper().strip("()")
    return bool(re.fullmatch(r"(?:A|B|C|D)[+-]?|P", grade))

def calculate_passed_course_summary(courses):
    """Return overall and category-wise passed-course counts and passed credit hours."""
    passed = [course for course in courses if is_passing_grade(course)]
    summary = {
        "Passed Courses Count": len(passed),
        "Passed Credit Hours Toward Degree": round(
            sum(float(course.get("Credit Hours", 0) or 0) for course in passed), 2
        ),
    }
    for category in CATEGORIES:
        category_passed = [
            course for course in passed
            if course.get("Category", "").casefold() == category.casefold()
        ]
        summary[f"{category} Passed Courses Count"] = len(category_passed)
        summary[f"{category} Passed Credit Hours"] = round(
            sum(float(course.get("Credit Hours", 0) or 0) for course in category_passed), 2
        )
    return summary

MSCS_SPECIALIZATION_COURSES = {
    "Software Engineering": [
        "Agile Software Project Management",
        "Business Process Re-Engineering",
        "Information Processing Techniques",
        "Advanced Topics in Software Quality Assurance",
        "Software Risk Management",
    ],
    "AI & Data Science": [
        "Data Science for Mathematics",
        "Deep Learning",
        "Machine Learning",
        "Information Extraction",
        "Big Data",
        "Computer Vision",
        "Data Science",
        "Generative AI Large Language Model",
        "Information Retrieval Techniques",
    ],
    "Network and Information Security": [
        "Bioinformatics",
        "Wireless Sensor and Mesh Networking",
        "Topics in Network Security",
        "Enterprise Information Security",
        "Information Security Management",
        "Cyber Security",
        
    ],
}


def calculate_mscs_specialization(student, courses):
    """Evaluate MSCS specialization and report per-domain passed elective count and credit hours."""
    result = {
        "Thesis I Passed": "No",
        "Thesis II Passed": "No",
        "Software Engineering Domain Electives Passed": 0,
        "Software Engineering Specialization Course Count": 0,
        "Software Engineering Specialization Credit Hours": 0,
        "AI & Data Science Domain Electives Passed": 0,
        "AI & Data Science Specialization Course Count": 0,
        "AI & Data Science Specialization Credit Hours": 0,
        "Network and Information Security Domain Electives Passed": 0,
        "Network and Information Security Specialization Course Count": 0,
        "Network and Information Security Specialization Credit Hours": 0,
        "Software Engineering Specialization Status": "Not eligible",
        "AI & Data Science Specialization Status": "Not eligible",
        "Network and Information Security Specialization Status": "Not eligible",
        "MSCS Specialization Awarded": "Not eligible",
        "MSCS Specialization Eligibility Details": "",
    }

    if (student.get("Program", "") or "").strip().upper() != "MSCS":
        result["MSCS Specialization Awarded"] = "Not applicable — selected program is not MSCS"
        result["MSCS Specialization Eligibility Details"] = "Specialization rules are configured for MSCS only."
        for area in MSCS_SPECIALIZATION_COURSES:
            result[f"{area} Specialization Status"] = "Not applicable"
        return result

    thesis_1_passed = any(is_thesis_1(c) and is_passing_grade(c) for c in courses)
    thesis_2_passed = any(is_thesis_2(c) and is_passing_grade(c) for c in courses)
    result["Thesis I Passed"] = "Yes" if thesis_1_passed else "No"
    result["Thesis II Passed"] = "Yes" if thesis_2_passed else "No"

    domain_count_fields = {
        "Software Engineering": "Software Engineering Domain Electives Passed",
        "AI & Data Science": "AI & Data Science Domain Electives Passed",
        "Network and Information Security": "Network and Information Security Domain Electives Passed",
    }
    eligible_areas = []
    for area, titles in MSCS_SPECIALIZATION_COURSES.items():
        normalized_titles = {normalize_course_title(title) for title in titles}
        # Count distinct passed elective courses within this domain only.
        # Deduplicate by normalized course title so repeated attempts are not double-counted.
        domain_courses_by_title = {}
        for course in courses:
            normalized_title = normalize_course_title(course.get("Course Title", ""))
            if (normalized_title in normalized_titles
                    and course.get("Category", "").casefold() == "elective"
                    and is_passing_grade(course)):
                domain_courses_by_title.setdefault(normalized_title, course)

        domain_courses = list(domain_courses_by_title.values())
        count = len(domain_courses)
        credit_hours = round(sum(float(c.get("Credit Hours", 0) or 0) for c in domain_courses), 2)
        result[domain_count_fields[area]] = count
        result[f"{area} Specialization Course Count"] = count
        result[f"{area} Specialization Credit Hours"] = credit_hours

        awarded = thesis_1_passed and thesis_2_passed and count >= 3
        result[f"{area} Specialization Status"] = "Awarded" if awarded else "Not eligible"
        if awarded:
            eligible_areas.append(area)

    if eligible_areas:
        result["MSCS Specialization Awarded"] = "; ".join(eligible_areas)
        result["MSCS Specialization Eligibility Details"] = (
            "Eligible: both Thesis I and Thesis II passed, and at least 3 listed domain electives passed."
        )
    else:
        missing = []
        if not thesis_1_passed:
            missing.append("Thesis I not passed")
        if not thesis_2_passed:
            missing.append("Thesis II not passed")
        for area, field in domain_count_fields.items():
            count = result[field]
            if count < 3:
                missing.append(f"{area}: {count}/3 passed domain electives")
        result["MSCS Specialization Eligibility Details"] = "; ".join(missing) if missing else "No specialization met all conditions."

    return result


def style_sheet(ws):
    fill = PatternFill("solid", fgColor="17365D")
    for cell in ws[1]:
        cell.fill = fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = Border(bottom=Side(style="thin", color="D9E2F3"))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for col in ws.columns:
        width = min(max(max((len(str(c.value)) for c in col if c.value is not None), default=10) + 2, 12), 48)
        ws.column_dimensions[get_column_letter(col[0].column)].width = width
    ws.row_dimensions[1].height = 30

def create_excel(output, students, courses, semesters, logs):
    wb = Workbook()
    ws = wb.active
    ws.title = "Student Info"

    # Print-friendly vertical layout: each student is a Field/Value block.
    fields = [
        "Source File", "Registration No", "Student Name", "Date of Birth", "Program",
        "Transcript Date", "Credit Hours Attempted", "Credit Hours Earned", "CGPA",
        "Credit Hours Accepted Towards Degree", "Passed Courses Count",
        "Passed Credit Hours Toward Degree", "Core Passed Courses Count",
        "Core Passed Credit Hours", "Elective Passed Courses Count",
        "Elective Passed Credit Hours", "Thesis Passed Courses Count",
        "Thesis Passed Credit Hours", "Deficiency Passed Courses Count",
        "Deficiency Passed Credit Hours", "Thesis I Passed", "Thesis II Passed",
        "Software Engineering Domain Electives Passed",
        "AI & Data Science Domain Electives Passed",
        "Network and Information Security Domain Electives Passed",
        "Software Engineering Specialization Course Count",
        "Software Engineering Specialization Credit Hours",
        "Software Engineering Specialization Status",
        "AI & Data Science Specialization Course Count",
        "AI & Data Science Specialization Credit Hours",
        "AI & Data Science Specialization Status",
        "Network and Information Security Specialization Course Count",
        "Network and Information Security Specialization Credit Hours",
        "Network and Information Security Specialization Status",
        "MSCS Specialization Awarded", "MSCS Specialization Eligibility Details",
    ]
    dark_blue = "17365D"
    light_blue = "D9E2F3"
    for student_index, student in enumerate(students):
        if student_index:
            # Start each student's profile on a new printed page.
            from openpyxl.worksheet.pagebreak import Break
            ws.row_breaks.append(Break(id=ws.max_row))
            ws.append([])
        title_row = 1 if ws.max_row == 1 and ws.cell(1, 1).value is None else ws.max_row + 1
        ws.cell(title_row, 1, f"STUDENT INFORMATION — {student.get('Student Name') or 'Student'}")
        ws.cell(title_row, 1).font = Font(color="FFFFFF", bold=True, size=12)
        ws.cell(title_row, 1).fill = PatternFill("solid", fgColor=dark_blue)
        ws.cell(title_row, 1).alignment = Alignment(vertical="center", wrap_text=True)
        ws.merge_cells(start_row=title_row, start_column=1, end_row=title_row, end_column=2)
        ws.row_dimensions[title_row].height = 26
        header_row = title_row + 1
        ws.cell(header_row, 1, "Field")
        ws.cell(header_row, 2, "Value")
        for cell in ws[header_row]:
            cell.fill = PatternFill("solid", fgColor=dark_blue)
            cell.font = Font(color="FFFFFF", bold=True)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = Border(bottom=Side(style="thin", color=light_blue))
        for field in fields:
            row = ws.max_row + 1
            ws.cell(row, 1, field)
            ws.cell(row, 2, student.get(field, ""))
            ws.cell(row, 1).font = Font(bold=True)
            ws.cell(row, 1).fill = PatternFill("solid", fgColor="EAF0F8")
            for cell in ws[row]:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.border = Border(bottom=Side(style="hair", color="D9E2F3"))

    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 38
    ws.freeze_panes = "A3"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = 0.35
    ws.page_margins.right = 0.35
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.5
    ws.print_options.horizontalCentered = True
    if ws.max_row:
        ws.print_area = f"A1:B{ws.max_row}"

    # The remaining sheets keep their usual row-per-record format.
    definitions = [
        ("Courses", ["Source File","Registration No","Student Name","Program","Course Section","Semester","Course Code","Course Title","Credit Hours","Grade","Category"], courses),
        ("Semester Summary", ["Source File","Registration No","Student Name","Semester","Semester Credit Hours","Semester GPA","Cumulative Credit Hours","Cumulative GPA"], semesters),
        ("Specialization Summary", [
            "Source File", "Registration No", "Student Name", "Program",
            "Thesis I Passed", "Thesis II Passed",
            "Software Engineering Status", "Software Engineering Passed Elective Count", "Software Engineering Passed Elective Credit Hours",
            "AI & Data Science Status", "AI & Data Science Passed Elective Count", "AI & Data Science Passed Elective Credit Hours",
            "Network and Information Security Status", "Network and Information Security Passed Elective Count", "Network and Information Security Passed Elective Credit Hours",
            "Specialization(s) Awarded", "Eligibility Details"
        ], [{
            "Source File": student.get("Source File", ""),
            "Registration No": student.get("Registration No", ""),
            "Student Name": student.get("Student Name", ""),
            "Program": student.get("Program", ""),
            "Thesis I Passed": student.get("Thesis I Passed", ""),
            "Thesis II Passed": student.get("Thesis II Passed", ""),
            "Software Engineering Status": student.get("Software Engineering Specialization Status", ""),
            "Software Engineering Passed Elective Count": student.get("Software Engineering Specialization Course Count", 0),
            "Software Engineering Passed Elective Credit Hours": student.get("Software Engineering Specialization Credit Hours", 0),
            "AI & Data Science Status": student.get("AI & Data Science Specialization Status", ""),
            "AI & Data Science Passed Elective Count": student.get("AI & Data Science Specialization Course Count", 0),
            "AI & Data Science Passed Elective Credit Hours": student.get("AI & Data Science Specialization Credit Hours", 0),
            "Network and Information Security Status": student.get("Network and Information Security Specialization Status", ""),
            "Network and Information Security Passed Elective Count": student.get("Network and Information Security Specialization Course Count", 0),
            "Network and Information Security Passed Elective Credit Hours": student.get("Network and Information Security Specialization Credit Hours", 0),
            "Specialization(s) Awarded": student.get("MSCS Specialization Awarded", ""),
            "Eligibility Details": student.get("MSCS Specialization Eligibility Details", ""),
        } for student in students]),
        ("Conversion Log", ["Source File","Status / Notes"], [{"Source File": f, "Status / Notes": n} for f,n in logs]),
    ]
    for name, headers, rows in definitions:
        sheet = wb.create_sheet(name)
        sheet.append(headers)
        for row in rows:
            sheet.append([row.get(h, "") for h in headers])
        style_sheet(sheet)
        sheet.page_setup.orientation = "landscape"
        sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 0
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.page_margins.left = 0.25
        sheet.page_margins.right = 0.25
    wb.save(output)

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("780x520")
        self.minsize(680, 420)
        self.files = []
        self.category_memory = load_category_memory()
        self.selected_program = tk.StringVar(value="MSCS")
        self.output = tk.StringVar(value=str(Path.home() / "transcripts_converted.xlsx"))
        frame = ttk.Frame(self, padding=18)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text=APP_TITLE, font=("TkDefaultFont", 16, "bold")).pack(anchor="w")
        ttk.Label(frame, text="Select a program, then categorize transcript courses as Core, Elective, Thesis, or Deficiency.").pack(anchor="w", pady=(8,14))
        buttons = ttk.Frame(frame); buttons.pack(fill="x")
        ttk.Button(buttons, text="Select PDF files…", command=self.select_files).pack(side="left")
        ttk.Button(buttons, text="Clear list", command=self.clear_files).pack(side="left", padx=8)
        self.file_label = ttk.Label(frame, text="No files selected"); self.file_label.pack(anchor="w", pady=8)
        program_row = ttk.Frame(frame); program_row.pack(fill="x", pady=(8, 4))
        ttk.Label(program_row, text="Program for this conversion:").pack(side="left")
        self.program_combo = ttk.Combobox(
            program_row, textvariable=self.selected_program,
            values=["MSCS", "MSDS", "MSSE", "MSCY", "Other"],
            state="readonly", width=12
        )
        self.program_combo.pack(side="left", padx=(8, 0))
        ttk.Label(frame, text="Choose the program before converting. Course categories are remembered separately for each program.").pack(anchor="w", pady=(0, 6))
        ttk.Label(frame, text="Excel output file").pack(anchor="w", pady=(10,4))
        outrow = ttk.Frame(frame); outrow.pack(fill="x")
        ttk.Entry(outrow, textvariable=self.output).pack(side="left", fill="x", expand=True)
        ttk.Button(outrow, text="Browse…", command=self.choose_output).pack(side="left", padx=(8,0))
        action_row = ttk.Frame(frame)
        action_row.pack(anchor="w", pady=16)
        ttk.Button(action_row, text="Convert to Excel", command=self.convert).pack(side="left")
        ttk.Button(action_row, text="Manage Course Categories…", command=self.manage_categories).pack(side="left", padx=(10, 0))
        self.log_box = tk.Text(frame, height=12, wrap="word", state="disabled")
        self.log_box.pack(fill="both", expand=True)
        ttk.Label(frame, text="Review the Conversion Log and verify student identifiers before official use.").pack(anchor="w", pady=(10,0))

    def manage_categories(self):
        """Allow the user to correct saved course categories for a specific program."""
        known_programs = {"MSCS", "MSDS", "MSSE", "MSCY"}
        for key in self.category_memory:
            if "|" in key:
                known_programs.add(key.split("|", 1)[0])
        current = self.selected_program.get().strip()
        if current == "Other":
            current = ""
        programs = sorted(known_programs | ({current} if current else set()))

        win = tk.Toplevel(self)
        win.title("Manage Saved Course Categories")
        win.geometry("720x470")
        win.minsize(600, 380)
        win.transient(self)

        outer = ttk.Frame(win, padding=14)
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="Correct a saved category", font=("TkDefaultFont", 12, "bold")).pack(anchor="w")
        ttk.Label(
            outer,
            text="Changes are saved for future conversions in the selected program. Existing Excel files are not changed.",
            wraplength=670,
        ).pack(anchor="w", pady=(5, 10))

        program_row = ttk.Frame(outer)
        program_row.pack(fill="x", pady=(0, 8))
        ttk.Label(program_row, text="Program:").pack(side="left")
        program_var = tk.StringVar(value=current if current in programs else (programs[0] if programs else "MSCS"))
        program_combo = ttk.Combobox(
            program_row, textvariable=program_var, values=programs,
            state="readonly", width=14,
        )
        program_combo.pack(side="left", padx=(8, 0))

        list_frame = ttk.Frame(outer)
        list_frame.pack(fill="both", expand=True)
        columns = ("course", "category")
        tree = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="browse", height=13)
        tree.heading("course", text="Saved Course Title")
        tree.heading("category", text="Current Category")
        tree.column("course", width=470, anchor="w")
        tree.column("category", width=160, anchor="w")
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="left", fill="y")

        edit_frame = ttk.Frame(outer)
        edit_frame.pack(fill="x", pady=(10, 0))
        ttk.Label(edit_frame, text="Change selected course to:").pack(side="left")
        category_var = tk.StringVar(value="Core")
        category_combo = ttk.Combobox(
            edit_frame, textvariable=category_var, values=CATEGORIES,
            state="readonly", width=14,
        )
        category_combo.pack(side="left", padx=(8, 8))
        save_button = ttk.Button(edit_frame, text="Save Correction", state="disabled")
        save_button.pack(side="left")

        key_by_iid = {}

        def refresh_rows(*_):
            for iid in tree.get_children():
                tree.delete(iid)
            key_by_iid.clear()
            program = program_var.get().strip()
            for key in sorted(self.category_memory, key=lambda k: k.casefold()):
                if "|" not in key:
                    continue
                key_program, normalized_title = key.split("|", 1)
                category = self.category_memory.get(key)
                if key_program == program and category in CATEGORIES:
                    display_title = " ".join(word.capitalize() for word in normalized_title.split())
                    iid = tree.insert("", "end", values=(display_title, category))
                    key_by_iid[iid] = key
            save_button.configure(state="disabled")
            category_var.set("Core")

        def on_select(_event=None):
            selected = tree.selection()
            if not selected:
                save_button.configure(state="disabled")
                return
            values = tree.item(selected[0], "values")
            if len(values) > 1:
                category_var.set(values[1])
                save_button.configure(state="normal")

        def save_correction():
            selected = tree.selection()
            if not selected:
                messagebox.showwarning(APP_TITLE, "Select a course to correct.", parent=win)
                return
            key = key_by_iid.get(selected[0])
            category = category_var.get()
            if not key or category not in CATEGORIES:
                messagebox.showerror(APP_TITLE, "Select a valid course and category.", parent=win)
                return
            old_category = self.category_memory.get(key)
            self.category_memory[key] = category
            try:
                save_category_memory(self.category_memory)
            except OSError as exc:
                self.category_memory[key] = old_category
                messagebox.showerror(APP_TITLE, f"Could not save category correction: {exc}", parent=win)
                return
            refresh_rows()
            messagebox.showinfo(
                APP_TITLE,
                f"Saved: {key.split('|', 1)[1]}\n{old_category} → {category}\nProgram: {program_var.get()}",
                parent=win,
            )

        tree.bind("<<TreeviewSelect>>", on_select)
        program_combo.bind("<<ComboboxSelected>>", refresh_rows)
        save_button.configure(command=save_correction)
        refresh_rows()

        if not self.category_memory:
            ttk.Label(outer, text="No saved course categories yet. Categorize courses during a conversion first.").pack(anchor="w", pady=(8, 0))

    def log(self, message):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", message + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def select_files(self):
        chosen = filedialog.askopenfilenames(title="Select transcript PDFs", filetypes=[("PDF files", "*.pdf")])
        if chosen:
            self.files = list(chosen)
            self.file_label.configure(text=f"{len(self.files)} PDF file(s) selected")

    def clear_files(self):
        self.files = []
        self.file_label.configure(text="No files selected")

    def choose_output(self):
        chosen = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel workbook", "*.xlsx")],
                                              initialfile=Path(self.output.get()).name)
        if chosen:
            self.output.set(chosen)

    def convert(self):
        if not self.files:
            messagebox.showwarning(APP_TITLE, "Select at least one PDF transcript.")
            return
        selected_program = self.selected_program.get().strip()
        if selected_program not in {"MSCS", "MSDS", "MSSE", "MSCY", "Other"}:
            messagebox.showwarning(APP_TITLE, "Select a program before converting.")
            return
        if selected_program == "Other":
            custom = simpledialog.askstring(APP_TITLE, "Enter the program code (for example, MSIT):")
            if not custom or not custom.strip():
                messagebox.showwarning(APP_TITLE, "Enter a program code or select a listed program.")
                return
            selected_program = custom.strip().upper()
        output = Path(self.output.get()).expanduser()
        if output.suffix.lower() != ".xlsx":
            output = output.with_suffix(".xlsx")
        students, courses, semesters, logs = [], [], [], []
        for filename in self.files:
            try:
                student, crs, sems = parse_transcript(filename)
                crs, thesis_1_dropped = apply_thesis_completion_rule(crs)
                if not categorize_courses(self, student, crs, self.category_memory, selected_program):
                    self.log("Conversion cancelled while categorizing courses.")
                    return
                passed_summary = calculate_passed_course_summary(crs)
                student.update(passed_summary)
                student.update(calculate_mscs_specialization(student, crs))
                students.append(student); courses.extend(crs); semesters.extend(sems)
                warnings = []
                if not student["Registration No"]: warnings.append("Registration No not detected")
                if not student["Student Name"]: warnings.append("Student Name not detected")
                if not crs: warnings.append("No courses detected")
                note = f"Converted {len(crs)} courses and {len(sems)} semester summaries."
                if thesis_1_dropped:
                    note += f" Dropped {thesis_1_dropped} Thesis I course(s) because no eligible Thesis II was found."
                if warnings: note += " CHECK: " + "; ".join(warnings)
                logs.append((Path(filename).name, note))
            except Exception as e:
                logs.append((Path(filename).name, f"ERROR: {type(e).__name__}: {e}"))
        try:
            save_category_memory(self.category_memory)
            create_excel(output, students, courses, semesters, logs)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Could not save workbook: {e}")
            return
        for filename, note in logs: self.log(f"{filename}: {note}")
        messagebox.showinfo(APP_TITLE, f"Conversion complete.\nSaved to:\n{output}")

if __name__ == "__main__":
    App().mainloop()
