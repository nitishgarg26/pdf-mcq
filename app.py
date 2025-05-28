# app.py
import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches
import io
import pytesseract
from PIL import Image
import re
import os
from pix2tex import cli as pix2tex

# Configure checkpoint path
CHECKPOINT_DIR = os.path.join(os.path.expanduser("~"), ".pix2tex")
os.makedirs(CHECKPOINT_DIR, exist_ok=True)

def initialize_model():
    """Initialize LatexOCR model with custom checkpoint handling"""
    try:
        if not os.path.exists(os.path.join(CHECKPOINT_DIR, "checkpoints")):
            st.error("""Model checkpoints missing! Follow these steps:
                    1. Download from https://github.com/lukas-blecher/LaTeX-OCR/releases
                    2. Create folder: ~/.pix2tex/checkpoints
                    3. Place model.pth in checkpoints folder""")
            return None
        return pix2tex.LatexOCR()
    except Exception as e:
        st.error(f"Model initialization failed: {str(e)}")
        return None

model = initialize_model()

def clean_text(text, is_question=True):
    """Remove question numbers and option labels"""
    patterns = [
        r'^\s*([A-Za-z]?\d+[\.\)]\s*|Q\d+\s*|\))',  # Question numbers
        r'^\s*([\(\[]?[A-D1-4][\.\)\]]\s*|•\s*)'     # Option labels
    ]
    return re.sub(patterns[0] if is_question else patterns[1], '', text).strip()

def extract_pdf_content(pdf_bytes):
    """Extract content with layout preservation"""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    content = []
    
    for page in doc:
        blocks = page.get_text("dict", sort=True)["blocks"]
        for block in blocks:
            if block["type"] == 0:  # Text
                for line in block["lines"]:
                    for span in line["spans"]:
                        text = clean_text(span["text"])
                        if text:
                            content.append(("text", text, block["bbox"]))
            elif block["type"] == 1:  # Image
                xref = block["xref"]
                try:
                    base_image = doc.extract_image(xref)
                    content.append(("image", base_image["image"], block["bbox"]))
                except Exception as e:
                    st.warning(f"Could not extract image: {e}")
    return content

def process_image(image_bytes):
    """Process images with equation detection and OCR"""
    img = None
    try:
        img = Image.open(io.BytesIO(image_bytes)).convert('RGB')
        if model:
            return model(img)
        else:
            return pytesseract.image_to_string(img)
    except Exception as e:
        if img is not None:
            return pytesseract.image_to_string(img, config='--psm 6')
        return "Image could not be processed"

def create_word_document(content):
    """Generate clean Word document"""
    doc = Document()
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    
    # Set column widths
    table.columns[0].width = Inches(3.5)
    for col in table.columns[1:]:
        col.width = Inches(1.5)
        
    # Headers
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Question'
    hdr_cells[1].text = 'Option 1'
    hdr_cells[2].text = 'Option 2'
    hdr_cells[3].text = 'Option 3'
    hdr_cells[4].text = 'Option 4'

    current_row = table.add_row().cells
    col_idx = 0
    question_buffer = []
    options = []

    for item_type, data, _ in content:
        if item_type == "text":
            if col_idx == 0 and not options:
                question_buffer.append(data)
            else:
                options.append(clean_text(data, False))
        elif item_type == "image":
            question_buffer.append(f"Equation: {process_image(data)}")

        # Create new row when 4 options collected
        if len(options) == 4:
            current_row[0].text = '\n'.join(question_buffer)
            for i, opt in enumerate(options, 1):
                current_row[i].text = opt
            current_row = table.add_row().cells
            question_buffer = []
            options = []
            col_idx = 0

    # Ensure the last question/options are added if present
    if question_buffer and options:
        current_row[0].text = '\n'.join(question_buffer)
        for i, opt in enumerate(options, 1):
            current_row[i].text = opt

    return doc

def main():
    st.title("PDF MCQ to Word Converter")
    
    uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])
    
    if uploaded_file and model:
        with st.spinner("Processing..."):
            content = extract_pdf_content(uploaded_file.read())
            doc = create_word_document(content)
            
            # Save to buffer
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            st.download_button(
                label="Download Word Document",
                data=buffer,
                file_name="converted_mcqs.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

if __name__ == "__main__":
    main()
