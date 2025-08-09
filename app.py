
import streamlit as st
import pandas as pd
from io import BytesIO
from header_utils import parse_docx
from chunker import build_csv_rows

st.set_page_config(page_title="DOCX → CSV Chapter Chunker v3.1", layout="wide")
st.title("📚 DOCX → CSV Chapter Chunker (v3.1)")
st.caption("Paragraph-based chunks • 20% overlap within chapter • Quotation rows separated.")

book = st.text_input("Book Name", "Spirit of Islam")
author = st.text_input("Author Name", "Unknown Author")
min_w = st.slider("Min words per chunk", 50, 600, 180, step=10)
max_w = st.slider("Max words per chunk", 80, 800, 250, step=10)
uploaded = st.file_uploader("Upload .docx", type=["docx"])

if uploaded:
    rows = parse_docx(uploaded, h1_min=14, h2_min=13, h3_min=13, require_h1_bold=True, max_header_words=15)
    st.success("File parsed. Review detection below (read-only in this simple build).")
    df = pd.DataFrame(rows)
    st.dataframe(df.head(30), use_container_width=True)
    if st.button("Generate CSV"):
        out = build_csv_rows(rows, book, author, min_w, max_w, 0.20)
        buf = BytesIO(); out.to_csv(buf, index=False, encoding="utf-8-sig"); buf.seek(0)
        st.download_button("⬇️ Download CSV", data=buf, file_name="output.csv", mime="text/csv")
        st.success(f"CSV ready with {len(out)} rows.")
