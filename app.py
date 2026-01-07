import streamlit as st
from google.cloud import vision
from PIL import Image
import pandas as pd
import io
from docx import Document
from langdetect import detect

st.set_page_config(page_title="Smart Handwritten OCR", layout="wide")

st.title("🧠 Smart Handwritten OCR (Image / PDF)")
st.caption("Google Vision OCR | Bulk Upload | Auto Language | Word & Excel Download")

# ---------- GOOGLE VISION CLIENT ----------
client = vision.ImageAnnotatorClient()

# ---------- FILE UPLOAD ----------
uploaded_files = st.file_uploader(
    "📤 Upload Handwritten Images / PDFs",
    type=["jpg", "jpeg", "png", "pdf"],
    accept_multiple_files=True
)

# ---------- PROCESS BUTTON ----------
if uploaded_files:
    if st.button("🚀 Process OCR"):
        all_results = []

        for file in uploaded_files:
            st.write(f"📄 Processing: **{file.name}**")

            image = Image.open(file)
            img_byte_arr = io.BytesIO()
            image.save(img_byte_arr, format="PNG")

            content = img_byte_arr.getvalue()
            image_vision = vision.Image(content=content)

            response = client.document_text_detection(image=image_vision)
            text = response.full_text_annotation.text

            try:
                language = detect(text)
            except:
                language = "unknown"

            all_results.append({
                "File Name": file.name,
                "Language": language,
                "Extracted Text": text
            })

            st.text_area("📝 Extracted Text", text, height=200)

        # ---------- EXCEL DOWNLOAD ----------
        df = pd.DataFrame(all_results)
        excel_buffer = io.BytesIO()
        df.to_excel(excel_buffer, index=False)

        st.download_button(
            "⬇ Download Excel",
            excel_buffer.getvalue(),
            file_name="handwritten_ocr.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # ---------- WORD DOWNLOAD ----------
        doc = Document()
        for row in all_results:
            doc.add_heading(row["File Name"], level=2)
            doc.add_paragraph(row["Extracted Text"])

        word_buffer = io.BytesIO()
        doc.save(word_buffer)

        st.download_button(
            "⬇ Download Word",
            word_buffer.getvalue(),
            file_name="handwritten_ocr.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
