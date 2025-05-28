import streamlit as st
import pdfplumber
from docx import Document
import io
import re

# Page configuration
st.set_page_config(
    page_title="MCQ Document Processor",
    page_icon="📝",
    layout="wide"
)

def parse_mcq_text(text):
    """Parse raw text into a list of MCQs."""
    mcqs = []
    # Normalize whitespace
    text = re.sub(r'\s+', ' ', text)
    # Split into question blocks like "1. …"
    pattern = r'(\d+)\.\s*(.*?)(?=\d+\.|$)'
    for qnum, content in re.findall(pattern, text, flags=re.DOTALL):
        # Find options (A), (B), ...
        opts = re.findall(r'\(([A-D])\)\s*([^()]+)', content)
        if len(opts) >= 2:
            # Extract question text up to first option
            first_opt = content.find(f'({opts[0][0]})')
            qtext = content[:first_opt].strip()
            # Clean text
            qtext = re.sub(r'\s+', ' ', qtext)
            options = [{'label': lab, 'text': re.sub(r'\s+', ' ', txt).strip()}
                       for lab, txt in opts]
            mcqs.append({
                'number': qnum,
                'question': qtext,
                'options': options
            })
    return mcqs

def extract_mcq_from_pdf(uploaded_file):
    """Extract text from PDF and parse MCQs."""
    text = ""
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    return parse_mcq_text(text)

def extract_mcq_from_docx(uploaded_file):
    """Extract text from DOCX and parse MCQs."""
    doc = Document(uploaded_file)
    text = "\n".join(para.text for para in doc.paragraphs if para.text.strip())
    return parse_mcq_text(text)

def create_word_document(mcqs):
    """Generate a Word document in memory from parsed MCQs."""
    doc = Document()
    doc.add_heading('Multiple Choice Questions', level=1)
    doc.add_paragraph(f'Total Questions: {len(mcqs)}\n')
    for mcq in mcqs:
        # Create table: 1 column for question + len(options) columns
        cols = 1 + len(mcq['options'])
        table = doc.add_table(rows=1, cols=cols, style='Table Grid')
        row = table.rows[0]
        # Question cell
        row.cells[0].text = f"Q{mcq['number']}: {mcq['question']}"
        # Option cells
        for i, opt in enumerate(mcq['options']):
            row.cells[i+1].text = f"({opt['label']}) {opt['text']}"
        doc.add_paragraph()  # spacing
    # Save to bytes
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

def main():
    st.title("📝 MCQ Document Processor")
    st.markdown("Upload a PDF, DOCX, or TXT file of MCQs to generate a Word document.")

    uploaded = st.file_uploader(
        "Choose MCQ file",
        type=['pdf', 'docx', 'txt']
    )
    if not uploaded:
        st.info("Awaiting file upload…")
        return

    file_type = uploaded.name.split('.')[-1].lower()
    with st.spinner("Processing file…"):
        if file_type == 'pdf':
            mcqs = extract_mcq_from_pdf(uploaded)
        elif file_type == 'docx':
            mcqs = extract_mcq_from_docx(uploaded)
        else:  # txt
            content = uploaded.read().decode('utf-8')
            mcqs = parse_mcq_text(content)

    if not mcqs:
        st.error("No MCQs found. Please check file format and content.")
        return

    st.success(f"✅ Parsed {len(mcqs)} questions!")
    # Preview first 3 questions
    with st.expander("Preview Questions", expanded=True):
        for mcq in mcqs[:3]:
            st.markdown(f"**Q{mcq['number']}:** {mcq['question']}")
            cols = st.columns(len(mcq['options']))
            for col, opt in zip(cols, mcq['options']):
                col.write(f"({opt['label']}) {opt['text']}")
            st.markdown("---")
        if len(mcqs) > 3:
            st.write(f"...and {len(mcqs) - 3} more questions")

    if st.button("📥 Download Word Document"):
        doc_buffer = create_word_document(mcqs)
        st.download_button(
            label="Download .docx",
            data=doc_buffer,
            file_name=f"MCQs_{uploaded.name.rsplit('.',1)[0]}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

if __name__ == "__main__":
    main()
