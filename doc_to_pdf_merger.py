import os
from docx2pdf import convert
from PyPDF2 import PdfMerger

# Set paths
docx_folder = "docs_to_merge"
pdf_folder = "converted_pdfs"
os.makedirs(pdf_folder, exist_ok=True)
# Before cleaning th previous files.
for file in os.listdir(pdf_folder):
    file_path = os.path.join(pdf_folder, file)
    if os.path.isfile(file_path):
        os.remove(file_path)  # delete each PDF in converted_pdfs


# Step 1: Convert .docx to .pdf
for file in os.listdir(docx_folder):
    if file.endswith(".docx"):
        input_path = os.path.join(docx_folder, file)
        output_path = os.path.join(pdf_folder, file.replace(".docx", ".pdf"))
        convert(input_path, output_path)

# Step 2: Merge all PDFs into one
merger = PdfMerger()

for file in sorted(os.listdir(pdf_folder)):
    if file.endswith(".pdf"):
        merger.append(os.path.join(pdf_folder, file))

# Save the merged PDF
merger.write("final_combined.pdf")
merger.close()

# Step 3: Clean up files
for file in os.listdir(docx_folder):
    file_path = os.path.join(docx_folder, file)
    if os.path.isfile(file_path):
        os.remove(file_path)  # delete each DOCX in docs_to_merge

print("✅ All Word docs have been converted and merged into final_combined.pdf")
print("🧹 Cleanup complete: All source and intermediate files deleted.")
