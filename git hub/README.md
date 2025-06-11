# 🏆 Certificate Generator from Excel (DOCX to PDF)

Automate your certificate generation process with this Python tool!  
It reads data from an Excel file, fills a Word template with personalized content using placeholders, and exports both `.docx` and `.pdf` versions of each certificate.

---

## ✨ Features

- ✅ Load bulk data from an Excel spreadsheet
- ✅ Fill a `.docx` template with dynamic fields like `{{name}}`, `{{date}}`, etc.
- ✅ Automatically generate personalized certificates
- ✅ Convert `.docx` certificates to `.pdf`
- ✅ Organize output neatly in folders

---

## 📦 Requirements

Make sure you have Python installed and then install the following packages:

```bash
pip install pandas docxtpl openpyxl docx2pdf

Step-by-Step Guide
Prepare Excel Data
Populate your Excel file (data.xlsx) with all required fields like name, date, message, etc.

Design the Word Template
Create a Word document (message.docx) using placeholders that match your Excel column names.

Run the Script
Execute the Python script. It will:[make sure to add a path in each script ]

Read each row from Excel

Replace placeholders in the Word template

Save each filled template as a .docx file

Automatic PDF Conversion
After generating .docx files, the script will convert them into .pdf and save them into a dedicated PDF folder.

