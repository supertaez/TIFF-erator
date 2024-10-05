import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
import fitz  # PyMuPDF
from PIL import Image, ImageDraw
from docx import Document
import pandas as pd

stop_conversion = False

def convert_pdf_to_tiff(pdf_path, tiff_path, dpi=150, compression="tiff_lzw"):
    doc = fitz.open(pdf_path)
    images = []
    for page_num in range(len(doc)):
        if stop_conversion:
            return
        page = doc.load_page(page_num)
        zoom = dpi / 72  # 72 is the default DPI for PDFs
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append(img)

    if images:
        images[0].save(tiff_path, save_all=True, append_images=images[1:], compression=compression)

def convert_image_to_tiff(image_path, tiff_path, compression="tiff_lzw"):
    img = Image.open(image_path)
    img.save(tiff_path, compression=compression)

def convert_docx_to_tiff(docx_path, tiff_path, dpi=150, compression="tiff_lzw"):
    doc = Document(docx_path)
    images = []
    for para in doc.paragraphs:
        if stop_conversion:
            return
        img = Image.new("RGB", (2480, 3508), (255, 255, 255))  # A4 size at 300 DPI
        d = ImageDraw.Draw(img)
        d.text((10, 10), para.text, fill=(0, 0, 0))
        images.append(img)
    if images:
        images[0].save(tiff_path, save_all=True, append_images=images[1:], compression=compression)

def convert_excel_to_tiff(excel_path, tiff_path, dpi=150, compression="tiff_lzw"):
    df = pd.read_excel(excel_path)
    text = df.to_string()
    img = Image.new("RGB", (2480, 3508), (255, 255, 255))  # A4 size at 300 DPI
    d = ImageDraw.Draw(img)
    d.text((10, 10), text, fill=(0, 0, 0))
    img.save(tiff_path, compression=compression)

def convert_files(input_folder, output_folder, dpi=150, compression="tiff_lzw"):
    global stop_conversion
    stop_conversion = False
    files = os.listdir(input_folder)
    total_files = len(files)
    for index, filename in enumerate(files):
        if stop_conversion:
            messagebox.showinfo("Stopped", "Conversion process has been stopped.")
            return
        file_path = os.path.join(input_folder, filename)
        tiff_path = os.path.join(output_folder, os.path.splitext(filename)[0] + ".tiff")
        if filename.endswith(".pdf"):
            convert_pdf_to_tiff(file_path, tiff_path, dpi, compression)
        elif filename.lower().endswith((".png", ".jpg", ".jpeg", ".bmp", ".gif")):
            convert_image_to_tiff(file_path, tiff_path, compression)
        elif filename.endswith(".docx"):
            convert_docx_to_tiff(file_path, tiff_path, dpi, compression)
        elif filename.endswith(".xlsx"):
            convert_excel_to_tiff(file_path, tiff_path, dpi, compression)
        counter_label.config(text=f"Converting {index + 1} out of {total_files} files")
    messagebox.showinfo("Success", "All files have been converted to TIFF.")

def select_input_folder():
    folder_selected = filedialog.askdirectory()
    input_folder_var.set(folder_selected)

def select_output_folder():
    folder_selected = filedialog.askdirectory()
    output_folder_var.set(folder_selected)

def start_conversion():
    input_folder = input_folder_var.get()
    output_folder = output_folder_var.get()
    if not input_folder or not output_folder:
        messagebox.showwarning("Input Error", "Please select both input and output folders.")
        return
    conversion_thread = threading.Thread(target=convert_files, args=(input_folder, output_folder))
    conversion_thread.start()

def stop_conversion_process():
    global stop_conversion
    stop_conversion = True

# Create the main window
root = tk.Tk()
root.title("Tiff-erator")

input_folder_var = tk.StringVar()
output_folder_var = tk.StringVar()

# Create and place the widgets
tk.Label(root, text="Select Input Folder:").pack(pady=5)
tk.Entry(root, textvariable=input_folder_var, width=50).pack(pady=5)
tk.Button(root, text="Browse", command=select_input_folder).pack(pady=5)

tk.Label(root, text="Select Output Folder:").pack(pady=5)
tk.Entry(root, textvariable=output_folder_var, width=50).pack(pady=5)
tk.Button(root, text="Browse", command=select_output_folder).pack(pady=5)

tk.Button(root, text="Start Conversion", command=start_conversion).pack(pady=20)
tk.Button(root, text="Stop Conversion", command=stop_conversion_process).pack(pady=5)

# Add a label to show the conversion progress
counter_label = tk.Label(root, text="Converted 0 out of 0 files")
counter_label.pack(pady=10)

# Run the application
root.mainloop()
