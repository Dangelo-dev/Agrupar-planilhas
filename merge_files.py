import PyPDF2
import os
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

directory = filedialog.askdirectory(title="Select a directory")
print("Selected directory:", directory)

list_files = os.listdir(directory)
list_files.sort()

merger = PyPDF2.PdfMerger()

for file in list_files:
    file_path = os.path.join(directory, file)
    try:
        if file.endswith(".pdf"):
            print(f"Adding file: {file_path}")
            merger.append(file_path)
    except Exception as e:
        print(f"Error processing file {file}: {e}")

output_file_path = os.path.join(directory, "PDF Final.pdf")
print("Output file path:", output_file_path)
merger.write(output_file_path)
print("Merged PDF file created successfully.")