import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches
import io
import pytesseract
from PIL import Image
import numpy as np
from pix2tex import cli as pix2tex
import re

# Initialize pix2tex model
model = pix2tex.LatexOCR()

def extract_pdf_content(pdf_bytes):
    """Extract text and images from PDF with equation detection"""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    content = []
    
    for page in doc:
        # Extract text blocks with coordinates
        text_blocks = page.get_text("blocks")
        for block in text_blocks:
            if block[6] == 0:  # text block
                content.append(("text", block[4], block[:4]))
        
        # Extract images
        img_list = page.get_images(full=True)
        for img_index, img in enumerate(img_list):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            content.append(("image", image_bytes, img_index))
    
    # Sort content by vertical position
    content.sort(key=lambda x: x[2][1] if len(x) > 2 else 0)
    return content

def process_image(image_bytes):
    """Process image with OCR and equation detection"""
    try:
        # Try equation detection first
        equation = model(Image.open(io.BytesIO(image_bytes)))
        return f"${equation}$"
    except:
        # Fallback to Tesseract OCR
        image = Image.open(io.BytesIO(image_bytes)).convert('L')
        return pytesseract.image_to_string(image, config='--psm 6')

def create_word_document(content):
    doc = Document()
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    
    # Set column widths
    table.columns[0].width = Inches(3.5)
    for col in table.columns[1:]:
        col.width = Inches(1.5)
        
    # Add headers
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Question/Image'
    hdr_cells[1].text = 'Option A'
    hdr_cells[2].text = 'Option B'
    hdr_cells[3].text = 'Option C'
    hdr_cells[4].text = 'Option D'

    current_row = table.add_row().cells
    col_index = 0
    
    for item_type, data, _ in content:
        if item_type == "image":
            # Process image with OCR/equation detection
            text = process_image(data)
            current_row[col_index].text = text
            col_index = (col_index + 1) % 5
        elif item_type == "text":
            # Clean and structure text
            cleaned_text = re.sub(r'\s+', ' ', data).strip()
            if any(opt in cleaned_text for opt in ['(A)', '(B)', '(C)', '(D)']):
                # Split options
                for option in re.findall(r'\([A-D]\)\s*[^\(]+', cleaned_text):
                    current_row[col_index].text = option
                    col_index = (col_index + 1) % 5
            else:
                current_row[col_index].text = cleaned_text
                col_index = (col_index + 1) % 5
    
    return doc

def main():
    st.title("Open-Source MCQ Processor")
    
    uploaded_file = st.file_uploader("Upload PDF with MCQs", type=["pdf"])
    
    if uploaded_file:
        pdf_bytes = uploaded_file.read()
        content = extract_pdf_content(pdf_bytes)
        
        doc = create_word_document(content)
        
        # Save to buffer
        doc_buffer = io.BytesIO()
        doc.save(doc_buffer)
        doc_buffer.seek(0)
        
        st.download_button(
            label="Download Document",
            data=doc_buffer,
            file_name="processed_mcqs.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

if __name__ == "__main__":
    main()
