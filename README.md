# MCQ PDF to Word Converter

This Streamlit app lets you upload a PDF file containing multiple-choice questions (MCQs) with text and images. It extracts each question’s text, any associated image, and the four answer options, then generates a downloadable Word document where each question is a row in a table. The columns are: **Question (and image)**, **Option A**, **Option B**, **Option C**, **Option D**. Even if the original question has fewer than four options, the output table always has four option columns (blank if not available).

## Features

- Upload a PDF with MCQs (text + images).
- Automatic extraction of questions, images, and options A–D.
- Preview extracted questions in the app.
- Download a formatted Word (`.docx`) file with questions in a table.

## Setup Instructions

1. **Clone the repository** to your local machine.
2. Install the required libraries:
   ```bash
   pip install -r requirements.txt
