import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches
import io
import pytesseract
from PIL import Image
import re
from pix2tex import cli as pix2tex

# Initialize pix2tex model
model = pix2tex.LatexOCR()

def clean_question(text):
    """Remove question numbers and leading special characters"""
    return re.sub(r'^\s*([A-Za-z]?\d+[\.\)]\s*|Q\d+\s*|\))', '', text).strip()

def clean_option(text):
    """Remove option labels and leading special characters"""
    return re.sub(r'^\s*([\(\[]?[A-D1-4][\.\)\]]\s*|•\s*)', '', text).strip()

def extract_pdf_content(pdf_bytes):
    """Extract content with layout preservation"""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    content = []
    
    for page in doc:
        # Extract text blocks with coordinates
        blocks = page.get_text("dict", sort=True)["blocks"]
        
        for block in blocks:
            if block["type"] == 0:  # Text block
                for line in block["lines"]:
                    for span in line["spans"]:
                        text = span["text"]
                        content.append(("text", text, block["bbox"]))
            
            elif block["type"] == 1:  # Image block
                xref = block["xref"]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                content.append(("image", image_bytes, block["bbox"]))
    
    return content

def process_image_content(image_bytes):
    """Process images with equation detection and OCR"""
    try:
        img = Image.open(io.BytesIO(image_bytes)).convert('RGB')
        equation = model(img)
        return f"\\({equation}\\)"
    except Exception as e:
        return pytesseract.image_to_string(img, config='--psm 6').strip()

def create_clean_word_document(content):
    """Generate Word document with cleaned content"""
    doc = Document()
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    
    # Set column widths
    table.columns[0].width = Inches(3.5)
    for col in table.columns[1:]:
        col.width = Inches(1.5)
        
    # Add headers
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Question'
    hdr_cells[1].text = 'Option 1'
    hdr_cells[2].text = 'Option 2'
    hdr_cells[3].text = 'Option 3'
    hdr_cells[4].text = 'Option 4'

    current_row = None
    current_question = []
    current_options = []

    for item_type, data, _ in content:
        if item_type == "text":
            cleaned = clean_question(data) if not current_options else clean_option(data)
            
            if not current_row:
                current_question.append(cleaned)
            else:
                if len(current_options) < 4:
                    current_options.append(cleaned)
                else:
                    current_question.append(cleaned)
        
        elif item_type == "image":
            processed_text = process_image_content(data)
            current_question.append(processed_text)
        
        # Create new row when we have 4 options
        if len(current_options) == 4:
            if current_row is None:
                current_row = table.add_row().cells
                
            # Add question with images
            current_row[0].text = '\n'.join(current_question)
            
            # Add options
            for i, opt in enumerate(current_options, 1):
                current_row[i].text = opt
            
            # Reset for next question
            current_row = None
            current_question = []
            current_options = []

    return doc

def main():
    st.title("MCQ PDF to Clean Word Converter")
    
    uploaded_file = st.file_uploader("Upload PDF with MCQs", type=["pdf"])
    
    if uploaded_file:
        with st.spinner("Processing PDF..."):
            pdf_bytes = uploaded_file.read()
            content = extract_pdf_content(pdf_bytes)
            
            doc = create_clean_word_document(content)
            
            # Save to buffer
            doc_buffer = io.BytesIO()
            doc.save(doc_buffer)
            doc_buffer.seek(0)
            
            st.success("PDF processed successfully!")
            
            st.download_button(
                label="Download Clean Document",
                data=doc_buffer,
                file_name="cleaned_mcqs.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

if __name__ == "__main__":
    main()
