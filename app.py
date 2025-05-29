import streamlit as st
import fitz  # PyMuPDF
import io
import re
from docx import Document
from docx.shared import Inches

st.title("MCQ PDF to Word Converter")

pdf_file = st.file_uploader("Upload a PDF file with MCQs", type=["pdf"])
if pdf_file is not None:
    try:
        pdf_bytes = pdf_file.read()
        pdf_doc = fitz.open(stream=pdf_bytes, filetype="pdf")

        full_text = ""
        all_images = []

        # --- Extract text and image blocks ---
        for page in pdf_doc:
            # Extract text
            full_text += page.get_text()

            # Extract images from block dictionary
            page_dict = page.get_text("dict")
            for block in page_dict.get("blocks", []):
                if block["type"] == 1 and "image" in block:
                    try:
                        image_bytes = block["image"]
                        if image_bytes:
                            all_images.append(image_bytes)
                    except Exception as e:
                        st.warning(f"Skipping image: {e}")

        # --- Clean and structure text into paragraphs ---
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

        # --- Parse questions and options ---
        question_pattern = re.compile(r'^\d{1,3}\.\s+')
        option_pattern = re.compile(r'^\(?[A-D]\)?[\.:]?\s+')

        questions = []
        current_q = None

        for para in paragraphs:
            if question_pattern.match(para):
                if current_q:
                    questions.append(current_q)
                q_text = question_pattern.sub('', para).strip()
                current_q = {"question_text": q_text, "options": [], "images": []}
            elif option_pattern.match(para):
                if current_q:
                    opt_text = option_pattern.sub('', para).strip()
                    current_q["options"].append(opt_text)
            else:
                if current_q:
                    if current_q["options"]:
                        current_q["options"][-1] += " " + para.strip()
                    else:
                        current_q["question_text"] += " " + para.strip()

        if current_q:
            questions.append(current_q)

        # --- Attach one image per question (optional, naive match) ---
        for i, q in enumerate(questions):
            if i < len(all_images):
                q["images"].append(all_images[i])

        # --- Display parsed questions ---
        if questions:
            st.header("Extracted Questions")
            for idx, q in enumerate(questions, start=1):
                st.markdown(f"**Question {idx}:** {q['question_text']}")
                for img in q["images"]:
                    st.image(img)
                for j, opt in enumerate(q["options"], start=1):
                    st.write(f"{chr(64 + j)}. {opt}")
        else:
            st.warning("No questions found. Check PDF formatting.")

        # --- Build Word Document ---
        if questions:
            for q in questions:
                while len(q["options"]) < 4:
                    q["options"].append("")

            doc = Document()
            table = doc.add_table(rows=1, cols=5)

            for i, q in enumerate(questions):
                row = table.add_row().cells if i > 0 else table.rows[0].cells
                para = row[0].paragraphs[0]
                para.add_run(q["question_text"])
                for img_bytes in q["images"]:
                    para.add_run("\n")
                    para.add_run().add_picture(io.BytesIO(img_bytes), width=Inches(2.5))
                for j in range(4):
                    row[j + 1].text = q["options"][j]

            # Save to buffer and offer download
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
