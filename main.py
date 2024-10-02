import os
import threading
import re
from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from PIL import Image
import tkinter as tk
from tkinter import filedialog
import fitz

output_file_name = "pretenzija.pdf"
output_file_path = f"C:/Users/josta/Downloads/{output_file_name}"
pdf = None
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


def select_input_file():
    global pdf
    global output_file_path
    file_path = filedialog.askopenfilename()
    print(f"Selected file: {file_path}")
    document = Document(file_path)
    pdf = canvas.Canvas(output_file_path, pagesize=A4)
    add_text_to_pdf(document)


def add_header_footer():
    global pdf
    global header_image_path
    page_width, page_height = A4

    # Header:
    try:
        header_image_width = 110
        header_image_height = 36
        pdf.drawImage(header_image_path, 70, page_height - 64,
                      width=header_image_width, height=header_image_height)
    except Exception as e:
        print(f"Error loading header image: {e}")

    # Footer (lines are written bottom-up):
    footer_text_column_1 = [
        "Kauno g. 16-308, LT-03212 Vilnius, Lietuva",
        "Įm. k. 305594385 ",
        "UAB „Claims management“"
    ]

    footer_text_column_2 = [
        "www.claimsmanagement.lt",
        "El. p.: paulius@claimsmanagement.lt",
        "Tel.nr.: +370 6 877 63 30"
    ]

    pdf.setFont("Arial", 10)

    x_col1 = 80
    y_start = 50
    line_height = 15

    for i, text in enumerate(footer_text_column_1):
        y_position = y_start + i * line_height
        pdf.drawString(x_col1, y_position, text)

    x_col2 = 360
    for i, text in enumerate(footer_text_column_2):
        y_position = y_start + i * line_height
        pdf.drawString(x_col2, y_position, text)


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

    print(f"Selected folder: {folder_path}")

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

            print(image_files)
            print("-------------------")

            # Extract images from PDFs if any
            for image_file in image_files:
                if image_file.lower().endswith('.pdf'):
                    pdf_path = os.path.join(subfolder_path, image_file)
                    extracted_images = extract_images_from_pdf(pdf_path, subfolder_path)
                    image_files.extend(extracted_images)

                # Remove the pdf files from the list after extracting images
            image_files = [f for f in image_files if not f.lower().endswith('.pdf')]

            print(image_files)

            if not image_files:
                print(f"No images found in folder: {subfolder_name}")
                continue

            num_of_images = len(image_files)

            if num_of_images > 1:
                paste_images_to_pdf_4x4(image_files, subfolder_path, num_of_images)
            elif num_of_images == 1:
                paste_images_to_pdf_1pic(image_files, subfolder_path)

    # Hide the loading label when the process is done
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
    output_file_path = f"{folder_path}/{output_file_name}"
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

    register_fonts()

    btn_select_output_folder = tk.Button(root, text="Select PDF output folder", command=select_output_folder)
    btn_select_docx = tk.Button(root, text="Select Word Document", command=select_input_file)
    btn_select_image_folder = tk.Button(root, text="Select Image Folder", command=select_image_folder)
    btn_generate_pdf = tk.Button(root, text="Generate PDF", command=save_pdf)

    btn_select_output_folder.pack(pady=10)
    btn_select_docx.pack(pady=10)
    btn_select_image_folder.pack(pady=10)

    loading_label = tk.Label(root, text="")
    loading_label.pack(pady=10)

    btn_generate_pdf.pack(pady=10)

    root.mainloop()
