import streamlit as st
import pymupdf as fitz  # PyMuPDF
import io
import re
from docx import Document
from docx.shared import Inches

st.title("MCQ PDF to Word Converter")

# File uploader for PDF input
pdf_file = st.file_uploader("Upload a PDF file with MCQs", type=['pdf'])
if pdf_file is not None:
    try:
        # Read PDF into memory
        pdf_bytes = pdf_file.read()
        pdf_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        questions = []
        current_q = None

        # Regular expressions to detect questions and options
        question_pattern = re.compile(r'^(?:Q(?:uestion)?\s*)?\d+[\.\)]', re.IGNORECASE)
        option_pattern = re.compile(r'^[A-D][\.\)]\s*', re.IGNORECASE)

        # Extract text and images from each page
        for page in pdf_doc:
            page_dict = page.get_text("dict")
            blocks = page_dict["blocks"]
            # Sort blocks by vertical position (top to bottom)
            blocks.sort(key=lambda b: b['bbox'][1])
            for block in blocks:
                if block["type"] == 0:  # text block
                    lines = block["text"].splitlines()
                    for line in lines:
                        text = line.strip()
                        if not text:
                            continue
                        # Detect start of a new question
                        if question_pattern.match(text):
                            # Save previous question
                            if current_q:
                                questions.append(current_q)
                            # Remove leading number (e.g., "1.") from question text
                            q_text = question_pattern.sub("", text).strip()
                            current_q = {"question_text": q_text, "options": [], "images": []}
                        # Detect an option line
                        elif option_pattern.match(text):
                            if current_q:
                                # Remove the leading letter (e.g., "A.") from option text
                                opt_text = option_pattern.sub("", text).strip()
                                current_q["options"].append(opt_text)
                        else:
                            # Continuation line: append to question or last option
                            if current_q:
                                if current_q["options"]:
                                    # Append to last option
                                    current_q["options"][-1] += " " + text
                                else:
                                    # Append to question text
                                    current_q["question_text"] += " " + text
                elif block["type"] == 1:  # image block
                    # Attach any found image to the current question
                    if current_q:
                        image_bytes = block["image"]  # raw image bytes from PDF
                        current_q["images"].append(image_bytes)
            # End of page blocks
        # Append the last question if exists
        if current_q:
            questions.append(current_q)

        # Preview extracted questions and options
        if questions:
            st.header("Extracted Questions")
            for idx, q in enumerate(questions, start=1):
                st.markdown(f"**Question {idx}:** {q['question_text']}")
                # Display associated image(s) if any
                if q["images"]:
                    for img_bytes in q["images"]:
                        st.image(img_bytes)
                # Display options
                for opt_idx, opt in enumerate(q["options"], start=1):
                    st.write(f"{chr(64+opt_idx)}. {opt}")
        else:
            st.warning("No questions found in the PDF. Please ensure the format is correct.")

        # Generate Word document with questions in a table
        if questions:
            # Ensure exactly 4 options per question (pad with empty if needed)
            for q in questions:
                while len(q["options"]) < 4:
                    q["options"].append("")

            doc = Document()
            table = doc.add_table(rows=1, cols=5)
            # Populate the table rows
            for i, q in enumerate(questions):
                if i == 0:
                    row_cells = table.rows[0].cells
                else:
                    row_cells = table.add_row().cells
                # Fill question text and image in the first cell
                cell = row_cells[0]
                para = cell.paragraphs[0]
                para.add_run(q["question_text"])
                if q["images"]:
                    for img_bytes in q["images"]:
                        para.add_run("\n")  # newline before image
                        run = para.add_run()
                        run.add_picture(io.BytesIO(img_bytes), width=Inches(2.5))
                # Fill the four options (A-D) in the next cells
                for j, opt in enumerate(q["options"], start=1):
                    row_cells[j].text = opt
            # Save Word document to a bytes buffer
            output = io.BytesIO()
            doc.save(output)
            output.seek(0)
            # Download button for the Word file
            st.download_button(
                label="Download Word Document",
                data=output.getvalue(),
                file_name="questions.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    except Exception as e:
        st.error(f"Error processing PDF: {e}")
