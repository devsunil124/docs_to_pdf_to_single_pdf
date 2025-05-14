import os
from tkinter import Tk, filedialog, messagebox, Button, Label
from docx2pdf import convert
from PyPDF2 import PdfMerger

def process_files():
    docx_files = filedialog.askopenfilenames(filetypes=[("Word Documents", "*.docx")])
    if not docx_files:
        return

    pdf_folder = "converted_pdfs"
    os.makedirs(pdf_folder, exist_ok=True)

    # Clear previous PDFs
    for file in os.listdir(pdf_folder):
        file_path = os.path.join(pdf_folder, file)
        if os.path.isfile(file_path):
            os.remove(file_path)

    # Convert .docx to .pdf
    for docx_file in docx_files:
        filename = os.path.basename(docx_file).replace(".docx", ".pdf")
        output_path = os.path.join(pdf_folder, filename)
        convert(docx_file, output_path)

    # Merge all PDFs
    merger = PdfMerger()
    for file in sorted(os.listdir(pdf_folder)):
        if file.endswith(".pdf"):
            merger.append(os.path.join(pdf_folder, file))
    merger.write("final_combined.pdf")
    merger.close()


    messagebox.showinfo("Done", "✅ Word files converted and merged into final_combined.pdf")

# GUI Setup
root = Tk()
root.title("DOCX to Merged PDF Converter")
Label(root, text="Click to select Word files and convert them").pack(pady=10)
Button(root, text="Select Files and Convert", command=process_files).pack(pady=10)
root.mainloop()
