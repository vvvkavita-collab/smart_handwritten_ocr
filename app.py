import streamlit as st
from google.cloud import vision
from PIL import Image
from pdf2image import convert_from_bytes
from docx import Document
import pandas as pd
import io
from langdetect import detect

st.set_page_config(
    page_title="Smart Handwritten OCR",
    layout="wide"
)

st.title("🧠 Smart Handwritten OCR (Image / PDF)")
st.caption("Google Vision OCR | Bulk Upload | Auto Language | Word & Excel Download")

client = vision.ImageAnnotatorClient()

uploaded_files = st.file_uploader(
    "📤 Upload Images or PDF (Multiple allowed)",
    type=["jpg", "jpeg", "png", "pdf"],
    accept_multiple_files=True
)

def extract_text_from_image(image_bytes):
    image = vision.Image(content=image_bytes)
    response = client.document_text_detection(image=image)
    return response.full_text_annotation.text

all_text = ""

if uploaded_files:
    if st.button("🚀 Process Files"):
        with st.spinner("Extracting text using Google Vision OCR..."):
            for file in uploaded_files:
                if file.type == "application/pdf":
                    pages = convert_from_bytes(file.read())
                    for page in pages:
                        img_byte_arr = io.BytesIO()
                        page.save(img_byte_arr, format='PNG')
                        text = extract_text_from_image(img_byte_arr.getvalue())
                        all_text += text + "\n"
                else:
                    img = Image.open(file)
                    img_byte_arr = io.BytesIO()
                    img.save(img_byte_arr, format='PNG')
                    text = extract_text_from_image(img_byte_arr.getvalue())
                    all_text += text + "\n"

        st.success("OCR Completed Successfully!")

        # 🌐 Language Detection
        try:
            lang = detect(all_text)
        except:
            lang = "Unknown"

        st.subheader("🌐 Detected Language")
        st.write(lang)

        # 📝 Display Text
        st.subheader("📝 Extracted Text")
        st.text_area("Result", all_text, height=350)

        # 📄 Word Download
        doc = Document()
        for line in all_text.split("\n"):
            doc.add_paragraph(line)

        word_buffer = io.BytesIO()
        doc.save(word_buffer)

        st.download_button(
            "⬇️ Download as Word",
            data=word_buffer.getvalue(),
            file_name="OCR_Text.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

        # 📊 Excel (Table Auto Columns)
        rows = []
        for line in all_text.split("\n"):
            parts = [p for p in line.split(" ") if p.strip()]
            if len(parts) > 1:
                rows.append(parts)

        if rows:
            max_len = max(len(r) for r in rows)
            for r in rows:
                r.extend([""] * (max_len - len(r)))

            df = pd.DataFrame(rows)
            df.columns = [f"Column_{i+1}" for i in range(df.shape[1])]

            excel_buffer = io.BytesIO()
            df.to_excel(excel_buffer, index=False)

            st.download_button(
                "⬇️ Download as Excel",
                data=excel_buffer.getvalue(),
                file_name="OCR_Table.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
