import streamlit as st
import pdfplumber

st.title("🔍 Debug Teks PDF")

files = st.file_uploader("Upload 1 PDF SDLG", type=["pdf"], accept_multiple_files=False)

if files:
    with pdfplumber.open(files) as pdf:
        teks = "\n".join(p.extract_text() or "" for p in pdf.pages)
    st.success(f"✅ {len(teks)} karakter terekstrak")
    st.text_area("📋 Copy semua teks di bawah:", teks, height=600)
