import os
import threading
import re
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from PIL import Image
import tkinter as tk
from tkinter import filedialog
import fitz

output_file_name = "pretenzija.docx"
output_file_path = f"C:/Users/josta/Downloads/{output_file_name}"
pdf = None
output_doc = Document()
header_image_path = "CM_logo.png"
page_bottom_limit = 100


def register_fonts():
    pdfmetrics.registerFont(TTFont('Arial', 'Arial-Unicode-Regular.ttf'))
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'Arial-Unicode-Bold.ttf'))
    pdfmetrics.registerFont(TTFont('Arial-Bold-Italic', 'Arial-Unicode-Bold-Italic.ttf'))
    pdfmetrics.registerFont(TTFont('Arial-Italic', 'Arial-Unicode-Italic.ttf'))


def add_text_to_pdf(document):
    global pdf
    text_y_position = 750  # Starting Y position on the page
    font_size = 12  # Default font size
    line_spacing = 14  # Line spacing between paragraphs

    add_header_footer()

    page_width, page_height = A4

    for para in document.paragraphs:
        if text_y_position < page_bottom_limit:  # Start a new page if we're too far down
            pdf.showPage()
            add_header_footer()
            text_y_position = 750  # Reset position for new page

        # Check if the paragraph is a heading (you can modify this as needed)
        if para.style.name.startswith('Heading'):
            font_size = 16  # Increase font size for headings
            pdf.setFont("Arial", font_size)
        else:
            font_size = 12  # Normal font size for body text
            pdf.setFont("Arial", font_size)

        # Determine alignment safely (None means default to left alignment)
        try:
            alignment = para.alignment
        except ValueError:
            alignment = None  # Handle invalid alignment gracefully
        if alignment == 1:  # Center alignment
            text_x_position = page_width / 2  # Center horizontally
            draw_function = pdf.drawCentredString
        elif alignment == 2:  # Right alignment (if needed)
            text_x_position = page_width - 100
            draw_function = pdf.drawRightString
        else:  # Default to left alignment
            text_x_position = 100
            draw_function = pdf.drawString

        # Loop through runs within the paragraph to handle formatting
        for run in para.runs:
            text = run.text

            # Apply formatting for bold and italic
            if run.bold:
                pdf.setFont("Arial-Bold", font_size)
            elif run.italic:
                pdf.setFont("Arial-Italic", font_size)
            else:
                pdf.setFont("Arial", font_size)

            # Draw the text using the appropriate draw function (center, right, or left)
            draw_function(text_x_position, text_y_position, text)

            # Adjust position if there are multiple runs in the paragraph
            if alignment != 1:  # For non-centered text, increment x position
                text_x_position += pdf.stringWidth(text, "Arial", font_size)

        # Move to the next line after the paragraph
        text_y_position -= line_spacing

        # Start a new page if we're too far down
        if text_y_position < page_bottom_limit:
            pdf.showPage()
            add_header_footer()
            text_y_position = 750


def add_text_to_docx(input_doc):
    global output_doc

    for para in input_doc.paragraphs:
        new_para = output_doc.add_paragraph()

        new_para.style = para.style

        for run in para.runs:
            new_run = new_para.add_run(run.text)

            new_run.bold = run.bold
            new_run.italic = run.italic
            new_run.underline = run.underline

            if run.font.size:
                new_run.font.size = run.font.size

            if run.font.name:
                new_run.font.name = run.font.name

        if para.alignment is not None:
            new_para.alignment = para.alignment


def select_input_file():
    global output_file_path
    file_path = filedialog.askopenfilename()

    if not file_path:
        docx_file_label.config(text=f"Error opening file {file_path}")
        loading_label.update_idletasks()
        return

    docx_file_label.config(text="Word document selected")
    loading_label.update_idletasks()

    input_document = Document(file_path)
    add_text_to_docx(input_document)


def remove_table_borders(table):
    tbl = table._element
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'none')
        tblBorders.append(border)
    tbl.tblPr.append(tblBorders)


def add_header_footer():
    global output_doc
    global header_image_path

    for section in output_doc.sections:

        header = section.header
        header_para = header.add_paragraph()

        try:
            header_para.add_run().add_picture(header_image_path, width=Pt(100))
        except Exception as e:
            print(f"Error loading header image: {e}")

        footer = section.footer
        footer_table = footer.add_table(rows=1, cols=2, width=Pt(500))  # Create a table with 1 row and 2 columns

        remove_table_borders(footer_table)

        footer_table.columns[0].width = Pt(250)
        footer_table.columns[1].width = Pt(250)

        cell_1 = footer_table.cell(0, 0)
        cell_1.text = "Kauno g. 16-308, LT-03212 Vilnius, Lietuva\n" \
                      "Įm. k. 305594385 \n" \
                      "UAB „Claims management“"

        cell_2 = footer_table.cell(0, 1)
        cell_2.text = "www.claimsmanagement.lt\n" \
                      "El. p.: paulius@claimsmanagement.lt\n" \
                      "Tel.nr.: +370 6 877 63 30"


def paste_images_to_pdf_4x4(image_files, subfolder_path, num_of_images):
    global pdf
    image_count = 0
    positions = [(80, 450), (330, 450), (80, 160), (330, 160)]  # Positions for up to 4 images

    for i, image_file in enumerate(image_files):
        image_path = os.path.join(subfolder_path, image_file)

        try:
            with Image.open(image_path) as img:
                img = img.convert("RGB")
                img_width, img_height = img.size

                max_width = 216
                max_height = 260
                ratio = min(max_width / img_width, max_height / img_height)
                new_size = (int(img_width * ratio), int(img_height * ratio))

                # Use the in-memory ImageReader to embed the image directly
                img_buffer = ImageReader(img)

                x_pos, y_pos = positions[image_count % 4]
                pdf.drawImage(img_buffer, x_pos, y_pos, width=new_size[0], height=new_size[1])

        except Exception as e:
            print(f"Error processing image {image_file}: {e}")
            continue

        image_count += 1

        # After 4 images, add a new page
        if image_count % 4 == 0 and image_count != num_of_images:
            pdf.showPage()
            add_header_footer()


def paste_images_to_pdf_1pic(image_files, subfolder_path):
    position = (100, 160)

    for i, image_file in enumerate(image_files):
        image_path = os.path.join(subfolder_path, image_file)

        try:
            with Image.open(image_path) as img:
                img = img.convert("RGB")
                img_width, img_height = img.size

                max_width = 400
                max_height = 560
                ratio = min(max_width / img_width, max_height / img_height)
                new_size = (int(img_width * ratio), int(img_height * ratio))

                # Use the in-memory ImageReader to embed the image directly
                img_buffer = ImageReader(img)

                x_pos, y_pos = position
                pdf.drawImage(img_buffer, x_pos, y_pos, width=new_size[0], height=new_size[1])

        except Exception as e:
            print(f"Error processing image {image_file}: {e}")
            continue


def extract_images_from_pdf(pdf_path, output_folder):
    pdf_document = fitz.open(pdf_path)
    images = []
    image_xrefs = set()

    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        for img in page.get_images(full=True):
            xref = img[0]  # The image reference (xref)

            # Check if this image has already been extracted
            if xref in image_xrefs:
                continue  # Skip this image since it's already processed

            image_xrefs.add(xref)  # Mark the xref as processed
            base_image = pdf_document.extract_image(xref)
            image_bytes = base_image["image"]
            image_ext = base_image["ext"]
            image_name = f"pdf_image_{page_num}_{xref}.{image_ext}"

            # Save the extracted image temporarily to the output folder
            image_path = os.path.join(output_folder, image_name)
            with open(image_path, "wb") as f:
                f.write(image_bytes)

            images.append(image_name)

    pdf_document.close()
    return images


def process_images():
    global pdf

    # Select the main folder
    folder_path = filedialog.askdirectory()
    if not folder_path:
        loading_label.config(text="No folder selected!")
        return

    subfolders = [f for f in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, f))]

    # Sort subfolders by the number at the beginning of the folder name
    def extract_number(folder_name):
        match = re.match(r"(\d+)", folder_name)
        return int(match.group(0)) if match else float('inf')  # Sort non-matching folders last
    def extract_folder_free_text(folder_name):
        match = re.match(r"\d+ (.+)", folder_name)
        return match.group(1) if match else ""


    sorted_subfolders = sorted(subfolders, key=extract_number)

    # Loop through subfolders in the selected folder
    for subfolder_name in sorted_subfolders:
        subfolder_path = os.path.join(folder_path, subfolder_name)

        if os.path.isdir(subfolder_path):
            pdf.showPage()  # Page break for each new subfolder

            add_header_footer()

            pdf.setFont("Arial-Bold", 14)
            pdf.drawString(100, 740, f"{extract_folder_free_text(subfolder_name)}")

            image_files = [f for f in os.listdir(subfolder_path)
                           if f.lower().endswith(('.png', '.jpg', '.jpeg', '.webp', '.heic', '.pdf'))]

            # Extract images from PDFs if any
            for image_file in image_files:
                if image_file.lower().endswith('.pdf'):
                    pdf_path = os.path.join(subfolder_path, image_file)
                    extracted_images = extract_images_from_pdf(pdf_path, subfolder_path)
                    image_files.extend(extracted_images)

                # Remove the pdf files from the list after extracting images
            image_files = [f for f in image_files if not f.lower().endswith('.pdf')]

            if not image_files:
                print(f"No images found in folder: {subfolder_name}")
                continue

            num_of_images = len(image_files)

            if num_of_images > 1:
                paste_images_to_pdf_4x4(image_files, subfolder_path, num_of_images)
            elif num_of_images == 1:
                paste_images_to_pdf_1pic(image_files, subfolder_path)

    btn_generate_pdf.pack(pady=0)
    save_status_label.pack(pady=2)
    loading_label.config(text="Images processed successfully!")
    loading_label.update_idletasks()


def select_image_folder():
    loading_label.config(text="Processing images, please wait...")
    loading_label.update_idletasks()

    thread = threading.Thread(target=process_images)
    thread.start()


def select_output_folder():
    global output_file_path
    folder_path = filedialog.askdirectory()

    if folder_path:
        output_file_path = f"{folder_path}/{output_file_name}"
        output_folder_label.config(text=f"Output folder selected: {folder_path}")
        loading_label.update_idletasks()
    else:
        output_folder_label.config(text="Error selecting output folder")
        loading_label.update_idletasks()

def save_word():
    global output_file_path

    add_header_footer()
    output_doc.save(output_file_path)
    save_status_label.config(text="PDF saved")
    loading_label.update_idletasks()

def save_pdf():
    global pdf
    if pdf:
        pdf.save()
        save_status_label.config(text="PDF saved")
        loading_label.update_idletasks()
    else:
        save_status_label.config(text="Error saving PDF")
        loading_label.update_idletasks()


if __name__ == '__main__':
    root = tk.Tk()
    root.title("CM PDF Converter")

    register_fonts()

    btn_select_output_folder = tk.Button(root, text="Select PDF Output Folder", command=select_output_folder)
    btn_select_docx = tk.Button(root, text="Select Input Word Document", command=select_input_file)
    btn_select_image_folder = tk.Button(root, text="Select Image Folder", command=select_image_folder)
    btn_generate_pdf = tk.Button(root, text=f"Generate {output_file_name}", command=save_word)

    output_folder_label = tk.Label(root, text="")
    docx_file_label = tk.Label(root, text="")
    loading_label = tk.Label(root, text="")
    save_status_label = tk.Label(root, text="")

    btn_select_output_folder.pack(pady=10)
    output_folder_label.pack(pady=0)
    btn_select_docx.pack(pady=10)
    docx_file_label.pack(pady=2)
    btn_select_image_folder.pack(pady=10)
    loading_label.pack(pady=2)
    btn_generate_pdf.pack(pady=2)#btn_generate_pdf.forget()
    save_status_label.pack(pady=2)#save_status_label.forget()

    root.mainloop()
