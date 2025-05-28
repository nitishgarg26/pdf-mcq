import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches
from PIL import Image
import io
import re

st.title("MCQ PDF to Word Table Converter (with Images)")

def extract_mcqs_and_images(pdf_file):
    """
    Extracts MCQs and images from the PDF.
    Associates images to the closest preceding question.
    Returns a list of dicts: {'question': ..., 'options': [...], 'image': ...}
    """
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    mcqs = []
    img_map = {}  # page_num -> list of (bbox, image_bytes, ext)
    # Extract images per page
    for page_num in range(len(doc)):
        page = doc[page_num]
        img_map[page_num] = []
        for block in page.get_text("dict")["blocks"]:
            if block["type"] == 1:  # image block
                img_bytes = block["image"]
                ext = block["ext"]
                bbox = block["bbox"]
                img_map[page_num].append((bbox, img_bytes, ext))
    # Extract text and associate images
    for page_num in range(len(doc)):
        page = doc[page_num]
        text_blocks = []
        for block in page.get_text("dict")["blocks"]:
            if block["type"] == 0:  # text block
                text_blocks.append((block["bbox"], block["lines"]))
        # Flatten text lines into paragraphs
        paragraphs = []
        for bbox, lines in text_blocks:
            para = " ".join("".join(span["text"] for span in line["spans"]) for line in lines)
            if para.strip():
                paragraphs.append((bbox, para.strip()))
        # Use regex to extract MCQs from paragraphs
        for i, (bbox, para) in enumerate(paragraphs):
            m = re.match(r'(\d+)\.\s*(.*)', para)
            if not m:
                continue
            q_text = m.group(2)
            # Find options
            options = re.findall(r'\([A-D1-4]\)\s*([^\(]+)', q_text)
            q_text_clean = re.split(r'\([A-D1-4]\)', q_text)[0].strip()
            if options:
                # Try to find the nearest image below this question
                q_ymax = bbox[3]
                img_bytes = None
                img_ext = None
                min_dist = float("inf")
                for img_bbox, ib, iext in img_map[page_num]:
                    img_ymin = img_bbox[1]
                    dist = img_ymin - q_ymax
                    if 0 <= dist < min_dist:
                        min_dist = dist
                        img_bytes = ib
                        img_ext = iext
                mcqs.append({
                    "question": q_text_clean,
                    "options": [opt.strip() for opt in options],
                    "image": (img_bytes, img_ext) if img_bytes else None
                })
    return mcqs

def create_word_table(mcqs):
    """
    Creates a Word document with a table:
    - First column: Question text and image (if any)
    - Next columns: Options
    """
    doc = Document()
    max_options = max(len(mcq["options"]) for mcq in mcqs)
    table = doc.add_table(rows=1, cols=1+max_options)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Question'
    for i in range(max_options):
        hdr_cells[i+1].text = f"Option {chr(65+i)}"
    for mcq in mcqs:
        row_cells = table.add_row().cells
        # Add question text and image
        p = row_cells[0].paragraphs[0]
        p.add_run(mcq["question"])
        if mcq["image"]:
            img_bytes, img_ext = mcq["image"]
            image_stream = io.BytesIO(img_bytes)
            try:
                # Optionally resize image for better fit
                pil_img = Image.open(image_stream)
                pil_img.thumbnail((250, 250))
                img_buffer = io.BytesIO()
                pil_img.save(img_buffer, format="PNG")
                img_buffer.seek(0)
                p.add_run().add_picture(img_buffer, width=Inches(1.5))
            except Exception:
                pass  # If image fails, skip it
        # Add options
        for i, opt in enumerate(mcq["options"]):
            row_cells[i+1].text = opt
    return doc

uploaded_file = st.file_uploader("Upload your MCQ PDF", type=["pdf"])
if uploaded_file:
    with st.spinner("Extracting questions and images..."):
        mcqs = extract_mcqs_and_images(uploaded_file)
    if not mcqs:
        st.error("No MCQs found. Please check your PDF format.")
    else:
        st.success(f"Found {len(mcqs)} questions.")
        doc = create_word_table(mcqs)
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        st.download_button(
            label="Download Word Table",
            data=buffer,
            file_name="mcq_table_with_images.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        if st.checkbox("Preview extracted questions"):
            for i, mcq in enumerate(mcqs, 1):
                st.markdown(f"**Q{i}:** {mcq['question']}")
                if mcq['image']:
                    st.image(mcq['image'][0], caption=f"Image for Q{i}", width=150)
                for j, opt in enumerate(mcq['options']):
                    st.markdown(f"- {chr(65+j)}. {opt}")

st.markdown("""
---
**Instructions:**  
- The PDF should contain MCQs in a format similar to your sample files:  
  Each question starts with a number and options are marked (A), (B), etc.
- Images are associated with the nearest preceding question.
- Images are resized for better fit in the Word table.
""")
