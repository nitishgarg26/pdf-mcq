import streamlit as st
import pymupdf as fitz  # PyMuPDF
import io
import re
from docx import Document
from docx.shared import Inches

st.title("MCQ PDF to Word Converter")

# Upload the PDF
pdf_file = st.file_uploader("Upload a PDF file with MCQs", type=["pdf"])
if pdf_file is not None:
    try:
        pdf_bytes = pdf_file.read()
        pdf_doc = fitz.open(stream=pdf_bytes, filetype="pdf")

        full_text = ""
        all_images = []

        # Collect full text and all images
        for page in pdf_doc:
            full_text += page.get_text()
            images = page.get_images(full=True)
            for img in images:
                try:
                    if isinstance(img, tuple) and len(img) > 0:
                        xref = img[0]
                        base_image = pdf_doc.extract_image(xref)
                        image_bytes = base_image.get("image", None)
                        if image_bytes:
                            all_images.append(image_bytes)
                except Exception as img_err:
                    st.warning(f"Skipping image due to error: {img_err}")


        # Step 1: Merge text lines into paragraphs
        lines = full_text.splitlines()
        paragraphs = []
        current_para = ""
        for line in lines:
            if line.strip() == "":
                if current_para:
                    paragraphs.append(current_para.strip())
                    current_para = ""
            else:
                current_para += " " + line.strip()
        if current_para:
            paragraphs.append(current_para.strip())

        # Step 2: Extract questions and options
        question_pattern = re.compile(r'^\d{1,3}\.\s+(.*)', re.DOTALL)
        option_pattern = re.compile(r'^\(?[A-D]\)?[\.:]?\s+(.*)', re.DOTALL)

        questions = []
        current_q = None

        for para in paragraphs:
            if re.match(r'^\d{1,3}\.', para.strip()):
                # Start of a new question
                if current_q:
                    questions.append(current_q)
                q_text = re.sub(r'^\d{1,3}\.\s*', '', para).strip()
                current_q = {"question_text": q_text, "options": [], "images": []}
            elif option_pattern.match(para.strip()):
                if current_q:
                    opt_text = option_pattern.sub(r'\1', para).strip()
                    current_q["options"].append(opt_text)
            else:
                # Continuation
                if current_q:
                    if current_q["options"]:
                        current_q["options"][-1] += " " + para.strip()
                    else:
                        current_q["question_text"] += " " + para.strip()

        if current_q:
            questions.append(current_q)

        # Assign 1 image per question, in order (if needed)
        for i, q in enumerate(questions):
            if i < len(all_images):
                q["images"].append(all_images[i])

        # Display extracted content
        if questions:
            st.header("Extracted Questions")
            for idx, q in enumerate(questions, start=1):
                st.markdown(f"**Question {idx}:** {q['question_text']}")
                if q["images"]:
                    for img in q["images"]:
                        st.image(img)
                for j, opt in enumerate(q["options"], start=1):
                    st.write(f"{chr(64 + j)}. {opt}")
        else:
            st.warning("No questions found. Please check PDF formatting.")

        # Build Word doc
        if questions:
            for q in questions:
                while len(q["options"]) < 4:
                    q["options"].append("")

            doc = Document()
            table = doc.add_table(rows=1, cols=5)

            for i, q in enumerate(questions):
                if i == 0:
                    row_cells = table.rows[0].cells
                else:
                    row_cells = table.add_row().cells
                cell = row_cells[0]
                para = cell.paragraphs[0]
                para.add_run(q["question_text"])
                if q["images"]:
                    for img_bytes in q["images"]:
                        para.add_run("\n")
                        run = para.add_run()
                        run.add_picture(io.BytesIO(img_bytes), width=Inches(2.5))
                for j, opt in enumerate(q["options"], start=1):
                    row_cells[j].text = opt

            # Save doc
            output = io.BytesIO()
            doc.save(output)
            output.seek(0)

            st.download_button(
                label="Download Word Document",
                data=output.getvalue(),
                file_name="questions.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    except Exception as e:
        st.error(f"Error processing PDF: {e}")
