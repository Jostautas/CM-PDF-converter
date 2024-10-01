import os
from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from PIL import Image
import tkinter as tk
from tkinter import filedialog

output_file_path = "C:/Users/josta/Downloads/out.pdf"
pdf = None

def add_text_to_pdf(pdf, document):
    text_y_position = 750  # Starting Y position on the page

    for para in document.paragraphs:
        text = para.text

        if text:  # Only render non-empty paragraphs
            pdf.drawString(100, text_y_position, text)  # Draw the paragraph text
            text_y_position -= 20  # Move to the next line

            if text_y_position < 50:  # Start a new page if we're too far down
                pdf.showPage()
                text_y_position = 750  # Reset position for new page


def select_input_file():
    global pdf
    global output_file_path
    file_path = filedialog.askopenfilename()
    print(f"Selected file: {file_path}")
    document = Document(file_path)
    pdf = canvas.Canvas(output_file_path, pagesize=A4)
    add_text_to_pdf(pdf, document)


def select_image_folder():
    folder_path = filedialog.askdirectory()
    print(f"Selected image folder: {folder_path}")


def select_output_folder():
    global output_file_path
    folder_path = filedialog.askdirectory()
    output_file_path = f"{folder_path}/out.pdf"
    print(f"Selected output folder: {output_file_path}")


def save_pdf():
    global pdf
    if pdf:
        pdf.save()
        print("PDF saved!")
    else:
        print("No PDF to save.")



if __name__ == '__main__':
    root = tk.Tk()
    root.title("CM PDF Converter")

    btn_select_output_folder = tk.Button(root, text="Select PDF output folder", command=select_output_folder)
    btn_select_docx = tk.Button(root, text="Select Word Document", command=lambda: select_input_file())
    btn_select_image_folder = tk.Button(root, text="Select Image Folder", command=select_image_folder)
    btn_generate_pdf = tk.Button(root, text="Generate PDF", command=save_pdf)

    btn_select_output_folder.pack(pady=10)
    btn_select_docx.pack(pady=10)
    btn_select_image_folder.pack(pady=10)
    btn_generate_pdf.pack(pady=10)

    root.mainloop()
