import os
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_ORIENTATION as WD_ORIENT
import re
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO

# Define paths
input_dir = './data/Ebook/'
output_dir = './data/Processed_Ebooks/'
os.makedirs(output_dir, exist_ok=True)

# Fonts mapping
FONT_MAP = {
    'Bengali': 'SutonnyMJ',
    'English': 'Times New Roman',
    'Arabic': 'Al Majeed Quranic',
    'Mixed (Nirmala UI)': 'Kalpurus'
}

# Utility function to set font for a run
def set_run_font(run, font_name, font_size):
    run.font.name = font_name
    run.font.size = Pt(font_size)

# Utility function to convert a table to an image if it exceeds margin constraints
def convert_table_to_image(table):
    # Render table content as an image
    table_content = "\n".join(["\t".join([cell.text for cell in row.cells]) for row in table.rows])
    font = ImageFont.load_default()
    width, height = 600, 200  # Example dimensions, adjust as needed
    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)
    draw.text((10, 10), table_content, fill="black", font=font)
    
    # Save to a stream
    image_stream = BytesIO()
    image.save(image_stream, format="PNG")
    image_stream.seek(0)
    return image_stream

# Function to process a Word document
def process_word_file(input_path, output_path):
    try:
        doc = Document(input_path)

        # 1. Set page layout
        for section in doc.sections:
            section.page_height = Inches(11)
            section.page_width = Inches(8.5)
            section.orientation = WD_ORIENT.PORTRAIT  # Explicitly set to Portrait
            section.top_margin = Inches(1)
            section.bottom_margin = Inches(1)
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)

        # 2. Remove headers and footers
        for section in doc.sections:
            section.header.is_linked_to_previous = False
            section.header.paragraphs[0].clear()
            section.footer.is_linked_to_previous = False
            section.footer.paragraphs[0].clear()

        # 3. Process paragraphs
        for para in doc.paragraphs:
            # Remove URLs, contact details, distributor names, and price lines
            if re.search(r'http[s]?://|www\.|@|[Rr]okomari', para.text):
                para.clear()
                continue

            # Apply font settings
            for run in para.runs:
                if re.search(r'[\u0980-\u09FF]', run.text):  # Bengali characters
                    set_run_font(run, FONT_MAP['Bengali'], 16)
                elif re.search(r'[\u0600-\u06FF]', run.text):  # Arabic characters
                    set_run_font(run, FONT_MAP['Arabic'], 16)
                else:  # English or default
                    set_run_font(run, FONT_MAP['English'], 16)

            # Center-align images
            if para.text == '':
                para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        # 4. Adjust tables
        for table in doc.tables:
            # If the table width exceeds the allowed limit, convert it to an image
            table_width = sum(cell.width for cell in table.rows[0].cells)
            if table_width > Inches(6):  # Exceeds margin constraints
                image_stream = convert_table_to_image(table)
                para = doc.add_paragraph()
                para.add_run().add_picture(image_stream, width=Inches(6))  # Adjust width as needed
                # Remove original table
                table._element.getparent().remove(table._element)

        # 5. Process text boxes
        for shape in doc.inline_shapes:
            if shape.type == 3:  # Type 3 indicates a text box
                text_box_content = shape.text_frame.text
                if text_box_content:
                    para = doc.add_paragraph(text_box_content)
                    para.style = doc.styles['Normal']
                shape._element.getparent().remove(shape._element)  # Remove the text box

        # Save the processed document
        doc.save(output_path)
    except Exception as e:
        print(f"Error processing file {input_path}: {e}")

# Process all .docx files in the input directory
for file_name in os.listdir(input_dir):
    if file_name.endswith('.docx'):
        input_file_path = os.path.join(input_dir, file_name)
        output_file_path = os.path.join(output_dir, file_name)
        process_word_file(input_file_path, output_file_path)

print(f"Processing completed. Files saved to {output_dir}")
