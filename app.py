import streamlit as st
import pandas as pd
from io import BytesIO
from header_utils import parse_docx
from chunker import build_csv_rows

st.set_page_config(page_title="DOCX → CSV Chapter Chunker v3 (20% overlap)", layout="wide")
st.title("📚 DOCX → CSV Chapter Chunker (v3 + 20% overlap within chapter)")
st.caption("Natural paragraph chunking • Quotations isolated • H1/H2/H3 editing • 20% overlap INSIDE a chapter only.")

# Book/Author
cba1, cba2 = st.columns(2)
book_name = cba1.text_input("Book Name", value="Spirit of Islam")
author_name = cba2.text_input("Author Name", value="Unknown Author")

# Detection settings
st.subheader("Header Detection Rules")
c1, c2, c3, c4 = st.columns(4)
max_header_words = c1.slider("Max header words", 3, 20, 15)
suppress_sentences = c2.checkbox("Downweight sentence-like lines", True)
suppress_quotes = c3.checkbox("Downweight quoted one-liners", True)
auto_detect = c4.checkbox("Enable Auto-detect", True)

st.markdown("**H1 (Main chapter)**")
h1c1, h1c2, h1c3, h1c4 = st.columns(4)
h1_enabled = h1c1.checkbox("Enable H1", True)
h1_min_size = h1c2.number_input("H1 min font size (pt)", value=14, step=1)
h1_require_bold = h1c3.checkbox("H1 require bold", True)
h1_align = h1c4.multiselect("H1 align allowed", ["left", "center", "right"],
                            default=["left", "center", "right"])

st.markdown("**H2**")
h2c1, h2c2, h2c3, h2c4 = st.columns(4)
h2_enabled = h2c1.checkbox("Enable H2", True)
h2_min_size = h2c2.number_input("H2 min font size (pt)", value=13, step=1)
h2_require_bold = h2c3.checkbox("H2 require bold", False)
h2_align = h2c4.multiselect("H2 align allowed", ["left", "center", "right"],
                            default=["left", "center", "right"])

st.markdown("**H3**")
h3c1, h3c2, h3c3, h3c4 = st.columns(4)
h3_enabled = h3c1.checkbox("Enable H3", True)
h3_min_size = h3c2.number_input("H3 min font size (pt)", value=13, step=1)
h3_require_bold = h3c3.checkbox("H3 require bold", False)
h3_align = h3c4.multiselect("H3 align allowed", ["left", "center", "right"],
                            default=["left", "center", "right"])

# Quotation rules
st.subheader("Quotation Detection")
qc1, qc2, qc3, qc4 = st.columns(4)
short_word_cutoff = qc1.slider("Short paragraph cutoff (words)", 10, 200, 60, step=5)
quote_centered = qc2.checkbox("Treat centered short lines as quotation", True)
quote_bold = qc3.checkbox("Treat bold short lines as quotation", True)
quote_italic = qc4.checkbox("Treat italic short lines as quotation", True)
quote_ql = st.checkbox('Treat quoted one-liners ("...") as quotation', True)

# Chunk settings
st.subheader("Chunk Settings (Paragraph-based)")
cc1, cc2 = st.columns(2)
min_words = cc1.slider("Min words per chunk (merge until ≥ this)", 50, 600, 180, step=10)
max_words = cc2.slider("Max words per chunk (aim for this cap)", 80, 800, 250, step=10)
overlap_ratio = 0.20  # <- FIXED 20% inside chapter only

uploaded = st.file_uploader("Upload a DOCX file", type=["docx"])

def build_rules():
    return {
        "auto_detect": auto_detect,
        "max_header_words": max_header_words,
        "suppress_sentences": suppress_sentences,
        "suppress_quotes": suppress_quotes,
        "levels": {
            "h1": {"enabled": h1_enabled, "min_size": float(h1_min_size), "require_bold": bool(h1_require_bold),
                   "allowed_align": [a.lower() for a in h1_align], "require_short_phrase": True},
            "h2": {"enabled": h2_enabled, "min_size": float(h2_min_size), "require_bold": bool(h2_require_bold),
                   "allowed_align": [a.lower() for a in h2_align], "require_short_phrase": True},
            "h3": {"enabled": h3_enabled, "min_size": float(h3_min_size), "require_bold": bool(h3_require_bold),
                   "allowed_align": [a.lower() for a in h3_align], "require_short_phrase": True},
        }
    }

def build_quote_rules():
    return {
        "short_word_cutoff": int(short_word_cutoff),
        "centered_short": bool(quote_centered),
        "bold_short": bool(quote_bold),
        "italic_short": bool(quote_italic),
        "quoted_one_liners": bool(quote_ql),
    }

if uploaded is not None and st.button("Preview & Edit Headers/Quotations", type="primary"):
    try:
        rules = build_rules()
        quote_rules = build_quote_rules()
        from docx import Document
        if rules["auto_detect"]:
            rows = parse_docx(uploaded, rules, quote_rules)
        else:
            # Manual mode with no auto-detect
            doc = Document(uploaded)
            rows = []
            idx = 0
            for p in doc.paragraphs:
                t = (p.text or "").strip()
                if not t:
                    continue
                rows.append({
                    "idx": idx, "text": t, "is_h1": False, "is_h2": False, "is_h3": False,
                    "is_header": False, "is_quote": False, "score": 0,
                    "all_caps": t.isupper(), "short_phrase": len(t.split()) <= max_header_words,
                    "avg_font_size": None, "max_font_size": None, "bold_fraction": 0.0, "any_bold": False,
                    "italic_fraction": 0.0, "any_italic": False, "align": "left", "style": "",
                    "sentence_like": False, "quoted_oneliner": False, "word_count": len(t.split())
                })
                idx += 1
        st.session_state["rows"] = rows
    except Exception as e:
        st.error(f"Failed to parse DOCX: {e}")

if "rows" in st.session_state:
    st.subheader("Review detection (toggle H1/H2/H3/Quotation if needed)")
    df = pd.DataFrame(st.session_state["rows"])
    cols = ["idx","text","is_h1","is_h2","is_h3","is_quote","score","align","any_bold","any_italic",
            "avg_font_size","max_font_size","all_caps","short_phrase","word_count"]
    df = df[cols]
    edited = st.data_editor(
        df,
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "text": st.column_config.TextColumn(width="large"),
            "is_h1": st.column_config.CheckboxColumn(help="Mark as Heading 1"),
            "is_h2": st.column_config.CheckboxColumn(help="Mark as Heading 2"),
            "is_h3": st.column_config.CheckboxColumn(help="Mark as Heading 3"),
            "is_quote": st.column_config.CheckboxColumn(help="Mark paragraph as quotation"),
            "score": st.column_config.NumberColumn(format="%.0f"),
            "avg_font_size": st.column_config.NumberColumn(format="%.1f"),
            "max_font_size": st.column_config.NumberColumn(format="%.1f"),
        },
        disabled=["idx","score","align","any_bold","any_italic","avg_font_size","max_font_size",
                  "all_caps","short_phrase","word_count","text"]
    )

    st.markdown("**Current headers preview:**")
    st.dataframe(
        edited[(edited["is_h1"]) | (edited["is_h2"]) | (edited["is_h3"])][["idx","is_h1","is_h2","is_h3","text"]].head(50),
        use_container_width=True
    )

    if st.button("Apply Edits"):
        idx_h1 = edited.set_index("idx")["is_h1"].to_dict()
        idx_h2 = edited.set_index("idx")["is_h2"].to_dict()
        idx_h3 = edited.set_index("idx")["is_h3"].to_dict()
        idx_qt = edited.set_index("idx")["is_quote"].to_dict()
        for r in st.session_state["rows"]:
            i = r["idx"]
            r["is_h1"] = bool(idx_h1.get(i, r["is_h1"]))
            r["is_h2"] = bool(idx_h2.get(i, r["is_h2"]))
            r["is_h3"] = bool(idx_h3.get(i, r["is_h3"]))
            r["is_header"] = r["is_h1"] or r["is_h2"] or r["is_h3"]
            r["is_quote"] = bool(idx_qt.get(i, r["is_quote"]))
        st.success("Edits applied.")

    if st.button("Generate CSV"):
        final_rows = st.session_state["rows"]
        out_df = build_csv_rows(final_rows, book_name, author_name, min_words, max_words, overlap_ratio)
        if out_df.empty:
            st.warning("No content produced. Adjust detection/chunk settings.")
        else:
            st.write(out_df.head(5))
            buf = BytesIO()
            out_df.to_csv(buf, index=False, encoding="utf-8-sig")
            buf.seek(0)
            st.download_button("⬇️ Download CSV", data=buf, file_name="output.csv", mime="text/csv")
            st.success(f"CSV ready with {len(out_df)} rows.")
