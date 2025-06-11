import pandas as pd
from docxtpl import DocxTemplate
import os

# Path to the Excel file
excel_path = r"./InputData/excel.xlsx"

# Path to the Word template
template_path = r"./InputData/template.docx"

# Output directory for generated documents
output_dir = r"./generated_certificates"
os.makedirs(output_dir, exist_ok=True)

# Load Excel data into a DataFrame
df = pd.read_excel(excel_path)

# Convert DataFrame rows into dictionaries
bulk_data = df.to_dict(orient="records")

# Generate documents
for entry in bulk_data:
    doc = DocxTemplate(template_path)
    doc.render(entry)
    
    # Safe filename
    filename = f"certificate_{entry['name'].replace(' ', '_')}.docx"
    doc.save(os.path.join(output_dir, filename))

print("✅ All certificates generated successfully from Excel!")
