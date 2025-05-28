# app.py
import streamlit as st
import pymupdf as fitz  # Correct import for PyMuPDF
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
    """Initialize LatexOCR model with proper error handling"""
    try:
        checkpoint_path = os.path.join(CHECKPOINT_DIR, "checkpoints", "model.pth")
        if not os.path.exists(checkpoint_path):
            st.error("Model checkpoints missing! Follow these steps:\n"
                     "1. Download weights.pth from: https://github.com/lukas-blecher/LaTeX-OCR/releases\n"
                     "2. Create folder: ~/.pix2tex/checkpoints\n"
                     "3. Place model.pth in checkpoints folder")
            return None
        return pix2tex.LatexOCR()
    except Exception as e:
        st.error(f"Model initialization failed: {str(e)}")
        return None

model = initialize_model()

def clean_content(text, is_question=True):
    """Clean question numbers and option labels using regex"""
    patterns = [
        r'^\s*([A-Za-z]?\d+[\.\)]\s*|Q\d+\s*|\))',  # Question numbers
        r'^\s*([\(\[]?[A-D1-4][\.\)\]]\s*|•\s*)'     # Option labels
    ]
    return re.sub(patterns[0] if is_question else patterns[1], '', text).strip()

def extract_pdf_elements(pdf_bytes):
    """Extract structured content from PDF with layout preservation"""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    content = []
    
    for page in doc:
        blocks = page.get_text("dict", sort=True)["blocks"]
        for block in blocks:
            if block["type"] == 0:  # Text block
                for line in block["lines"]:
                    for span in line["spans"]:
                        text = clean_content(span["text"])
                        if text:
                            content.append(("text", text, block["bbox"]))
            elif block["type"] == 1:  # Image block
                try:
                    xref = block["xref"]
                    base_image = doc.extract_image(xref)
                    content.append(("image", base_image["image"], block["bbox"]))
                except Exception as e:
                    st.warning(f"Skipped image: {str(e)}")
    return content

def process_image_content(image_bytes):
    """Process images with hybrid OCR/equation detection"""
    try:
        img = Image.open(io.BytesIO(image_bytes)).convert('RGB')
        return model(img) if model else pytesseract.image_to_string(img)
    except Exception as e:
        return pytesseract.image_to_string(img, config='--psm 6')

def generate_word_document(content):
    """Create structured Word document with cleaned content"""
    doc = Document()
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    
    # Configure column widths
    table.columns[0].width = Inches(3.5)
    for col in table.columns[1:]:
        col.width = Inches(1.5)
        
    # Table headers
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Question'
    hdr_cells[1].text = 'Option 1'
    hdr_cells[2].text = 'Option 2'
    hdr_cells[3].text = 'Option 3'
    hdr_cells[4].text = 'Option 4'

    current_row = table.add_row().cells
    question_buffer = []
    options = []

    for item_type, data, _ in content:
        if item_type == "text":
            if not options:  # Question text
                question_buffer.append(data)
            else:  # Option text
                options.append(clean_content(data, False))
                
        elif item_type == "image":
            question_buffer.append(f"Equation: {process_image_content(data)}")

        # Create new row when 4 options are accumulated
        if len(options) == 4:
            current_row[0].text = '\n'.join(question_buffer)
            for i, opt in enumerate(options[:4], 1):
                current_row[i].text = opt
            current_row = table.add_row().cells
            question_buffer = []
            options = []

    return doc

def main():
    st.title("PDF MCQ to Word Converter")
    
    uploaded_file = st.file_uploader("Upload PDF File", type=["pdf"])
    
    if uploaded_file:
        if not model:
            st.error("OCR model not initialized. Check setup instructions.")
            return
            
        with st.spinner("Processing PDF..."):
            try:
                content = extract_pdf_elements(uploaded_file.read())
                doc = generate_word_document(content)
                
                # Save to in-memory buffer
                buffer = io.BytesIO()
                doc.save(buffer)
                buffer.seek(0)
                
                st.success("Conversion successful!")
                st.download_button(
                    label="Download Word Document",
                    data=buffer,
                    file_name="converted_questions.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Processing failed: {str(e)}")

if __name__ == "__main__":
    main()
