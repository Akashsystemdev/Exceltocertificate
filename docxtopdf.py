import os
from docx2pdf import convert

# Hardcoded path to the folder with .docx files
source_folder = r"./generated_certificates"  # 🔁 Change this to your actual path

# Create 'PDF' directory inside the source folder
pdf_folder = os.path.join(source_folder, "PDF")
os.makedirs(pdf_folder, exist_ok=True)

# Loop through all files in the folder
for file in os.listdir(source_folder):
    if file.endswith(".docx"):
        docx_path = os.path.join(source_folder, file)
        pdf_path = os.path.join(pdf_folder, file.replace(".docx", ".pdf"))

        try:
            convert(docx_path, pdf_path)
            print(f"Converted: {file} → PDF")
        except Exception as e:
            print(f"Failed to convert {file}: {e}")

print("✅ Conversion completed.")
