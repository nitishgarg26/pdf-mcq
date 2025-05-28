import streamlit as st
import io
import re

import pymupdf as fitz
import pdfplumber
import pytesseract
from PIL import Image

from docx import Document
from docx.shared import Inches

# Add this near the top of your app.py
_invalid_xml_chars_re = re.compile(r'[\x00-\x08\x0B-\x0C\x0E-\x1F]')

def cleanup_text(s: str) -> str:
    """
    Strip out control characters that Python-docx cannot serialize.
    """
    return _invalid_xml_chars_re.sub("", s)


# In your create_word_document, update every run of add_run(part) to:
for part in parts:
    if part.startswith("$") and part.endswith("$"):
        # … equation handling …
    else:
        clean = cleanup_text(part)
        if clean:
            qpara.add_run(clean)


# Configure Streamlit page
st.set_page_config(page_title="MCQ Processor", layout="wide")

def clean_labels(text: str) -> str:
    """
    Remove question numbers (e.g. "1.") and option labels "(A)", "(B)", etc.
    """
    # strip leading question numbers
    text = re.sub(r'^\s*\d+\.\s*', '', text, flags=re.MULTILINE)
    # replace option labels with line breaks
    text = re.sub(r'\s*\([A-D1-4]\)\s*', '\n', text)
    return text.strip()

def extract_text_advanced(uploaded_file) -> str:
    """
    Use PyMuPDF + OCR fallback + pdfplumber to get high-fidelity text.
    """
    pdf_bytes = uploaded_file.read()
    text = ""
    # 1. Use PyMuPDF for text
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    for page in doc:
        block = page.get_text("text")
        if block and len(block.strip()) > 20:
            text += block + "\n"
        else:
            # OCR fallback
            pix = page.get_pixmap(dpi=300)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            ocr = pytesseract.image_to_string(img, config="--psm 1")
            text += ocr + "\n"
    doc.close()
    # 2. Supplement with pdfplumber (tables, columns)
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf2:
            for p in pdf2.pages:
                tbl_text = p.extract_text()
                if tbl_text:
                    text += tbl_text + "\n"
    except Exception:
        pass
    return text

def parse_mcq_text(raw: str):
    """
    Split raw text into MCQ blocks, clean labels, extract questions and options.
    """
    mcqs = []
    cleaned = clean_labels(raw)
    # question blocks like "1. ..." -> after cleaning, look for splits by blank lines and numbering
    blocks = re.split(r'\n(?=\d+\.)', raw)
    for block in blocks:
        block = block.strip()
        if not block:
            continue
        # remove original numbering and option labels
        block = clean_labels(block)
        # split question vs options
        parts = [p.strip() for p in block.split('\n') if p.strip()]
        if len(parts) < 2:
            continue
        question_text = parts[0]
        option_texts = parts[1:]
        options = [{"text": o} for o in option_texts]
        mcqs.append({"question": question_text, "options": options})
    return mcqs

def latex_to_image(latex: str) -> io.BytesIO:
    """
    Render LaTeX string to an image buffer.
    """
    import matplotlib.pyplot as plt
    buf = io.BytesIO()
    fig = plt.figure(figsize=(0.01, 0.01))
    fig.text(0, 0, f'${latex}$', fontsize=14)
    plt.axis('off')
    fig.savefig(buf, dpi=300, transparent=True, bbox_inches='tight', pad_inches=0)
    plt.close(fig)
    buf.seek(0)
    return buf

def detect_and_extract_equations(text: str):
    """
    Find inline LaTeX equations demarcated by $...$ and return mapping.
    """
    eqns = re.findall(r'\$(.+?)\$', text)
    return eqns

def create_word_document(mcqs):
    """
    Build a .docx in memory with questions in col1, options in subsequent cols,
    converting any inline $...$ to images.
    """
    doc = Document()
    doc.add_heading("Multiple Choice Questions", level=1)
    doc.add_paragraph(f"Total Questions: {len(mcqs)}\n")

    for mcq in mcqs:
        cols = 1 + len(mcq["options"])
        table = doc.add_table(rows=1, cols=cols, style="Table Grid")
        row = table.rows[0]

        # process question text: detect equations
        qcell = row.cells[0]
        qpara = qcell.paragraphs[0]
        qtext = mcq["question"]
        eqs = detect_and_extract_equations(qtext)
        # split by inline equations
        parts = re.split(r'(\$.+?\$)', qtext)
        for part in parts:
            if part.startswith("$") and part.endswith("$"):
                latex = part.strip("$")
                img_buf = latex_to_image(latex)
                run = qpara.add_run()
                run.add_picture(img_buf, width=Inches(2))
            else:
                clean = cleanup_text(part)
                if clean:
                    qpara.add_run(clean)

        # options
        for idx, opt in enumerate(mcq["options"]):
            cell = row.cells[idx + 1]
            para = cell.paragraphs[0]
            otext = opt["text"]
            eqs_o = detect_and_extract_equations(otext)
            parts_o = re.split(r'(\$.+?\$)', otext)
            for part in parts_o:
                if part.startswith("$") and part.endswith("$"):
                    buf = latex_to_image(part.strip("$"))
                    run = para.add_run()
                    run.add_picture(buf, width=Inches(1.2))
                else:
                    clean = cleanup_text(part)
                    if clean:
                        qpara.add_run(clean)

        doc.add_paragraph()  # spacer

    mem = io.BytesIO()
    doc.save(mem)
    mem.seek(0)
    return mem

def main():
    st.title("📝 Advanced MCQ Processor")
    uploaded = st.file_uploader("Upload PDF/DOCX/TXT", type=["pdf", "docx", "txt"])
    if not uploaded:
        st.info("Please upload a file.")
        return

    if uploaded.type == "application/pdf":
        raw_text = extract_text_advanced(uploaded)
    elif uploaded.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        import docx
        doc = docx.Document(uploaded)
        raw_text = "\n".join(p.text for p in doc.paragraphs)
    else:
        raw_text = uploaded.read().decode("utf-8")

    mcqs = parse_mcq_text(raw_text)
    if not mcqs:
        st.error("No MCQs detected.")
        return

    st.success(f"Extracted {len(mcqs)} questions")
    if st.button("Download Word"):
        doc_io = create_word_document(mcqs)
        st.download_button(
            "Download .docx",
            data=doc_io.getvalue(),
            file_name="mcqs_processed.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

if __name__ == "__main__":
    main()
