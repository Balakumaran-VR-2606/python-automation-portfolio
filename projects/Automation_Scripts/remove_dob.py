import fitz
import os
import re

input_folder = "input_pdfs"
output_folder = "output_pdfs"
os.makedirs(output_folder, exist_ok=True)


patterns = [
    r"Date of Birth\s+([0-9]{2}\s[A-Za-z]{3}\s[0-9]{4})",  
    r"Date of Birth\s+([0-9]{2}-[A-Za-z]{3}-[0-9]{4})"     
]

for file in os.listdir(input_folder):
    if file.endswith(".pdf"):
        pdf_path = os.path.join(input_folder, file)
        doc = fitz.open(pdf_path)

        for page in doc:
            text = page.get_text("text")

            for pattern in patterns:
                matches = re.findall(pattern, text)

                for date_text in matches:
                    instances = page.search_for(date_text)
                    for inst in instances:
                        page.add_redact_annot(inst, fill=(1, 1, 1))

            page.apply_redactions()

        doc.save(os.path.join(output_folder, file))

print("✅ All DOB formats removed macha!")