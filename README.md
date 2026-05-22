
# Talent Acquisition System

Talent Acquisition System is a Flask-based web application that analyzes resumes and provides a structured recruitment score, criteria breakdown, and job recommendations. Users can upload resumes in PDF, DOCX, or TXT format, optionally add a target job title and job description, and receive a tailored match analysis based on extracted skills and resume content.

## Features

- Resume upload and parsing for PDF, DOCX, and TXT files
- Skill and keyword extraction from resume content
- Target job matching using title and job description
- TAS match score with criteria-based breakdown
- Job recommendation support using predefined job data
- Secure upload handling with temporary file cleanup
- SQLite-based report storage
- Responsive and user-friendly web interface

## Technologies Used

- Python
- Flask
- HTML, CSS, JavaScript
- SQLite
- Pandas, NumPy
- pypdf / PyPDF2
- python-docx
- spaCy
- pytesseract and pdf2image for optional OCR support

## Project Objective

The objective of this project is to simplify and improve the recruitment screening process by using resume parsing, keyword analysis, and job matching techniques. It helps candidates and recruiters quickly understand how well a resume aligns with a target role.

## Installation

```bash
git clone https://github.com/SA-Sanush/Main_Project_ICT.git
cd Main_Project_ICT
pip install -r requirements.txt
```

## Run the Project

```bash
python app.py
```

Then open the local Flask URL in your browser.

## Author

Developed by [SA-Sanush] and Team (https://github.com/SA-Sanush)

## License

This project is created for educational and academic purposes.
