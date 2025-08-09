
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

ALIGN_MAP = {
    WD_ALIGN_PARAGRAPH.LEFT: "left",
    WD_ALIGN_PARAGRAPH.CENTER: "center",
    WD_ALIGN_PARAGRAPH.RIGHT: "right",
    WD_ALIGN_PARAGRAPH.JUSTIFY: "justify",
    None: "left",
}

def _pt(size): 
    try: return size.pt if size else None
    except: return None

def _any_bold(p): return any(r.bold for r in p.runs if r.text)
def _bold_fraction(p):
    total = sum(len(r.text or "") for r in p.runs) or 1
    bold = sum(len(r.text or "") for r in p.runs if r.bold)
    return bold/total

def _any_italic(p): return any(r.italic for r in p.runs if r.text)
def _italic_fraction(p):
    total = sum(len(r.text or "") for r in p.runs) or 1
    it = sum(len(r.text or "") for r in p.runs if r.italic)
    return it/total

def _avg_font(p):
    sizes = [_pt(r.font.size) for r in p.runs if r.font and r.font.size]
    return sum(sizes)/len(sizes) if sizes else None

def _max_font(p):
    sizes = [_pt(r.font.size) for r in p.runs if r.font and r.font.size]
    return max(sizes) if sizes else None

def _align(p): return ALIGN_MAP.get(p.alignment, "left")
def _style(p):
    try: return (p.style.name or "").lower()
    except: return ""

def parse_docx(file, h1_min=14, h2_min=13, h3_min=13, require_h1_bold=True, max_header_words=15):
    doc = Document(file)
    rows = []
    for i, p in enumerate(doc.paragraphs):
        text = (p.text or "").strip()
        if not text: 
            continue
        words = text.split()
        avg = _avg_font(p); mx = _max_font(p)
        any_b = _any_bold(p); bfrac = _bold_fraction(p)
        any_i = _any_italic(p); ifrac = _italic_fraction(p)
        align = _align(p); sty = _style(p)
        short = len(words) <= max_header_words and len(text) <= 120

        is_h1 = ((avg and avg>=h1_min) or (mx and mx>=h1_min)) and short and (any_b or bfrac>=0.4 if require_h1_bold else True)
        is_h2 = ((avg and avg>=h2_min) or (mx and mx>=h2_min)) and short
        is_h3 = ((avg and avg>=h3_min) or (mx and mx>=h3_min)) and short

        # Style hints
        stl = "heading" in sty
        if "heading 1" in sty: is_h1=True; is_h2=is_h3=False
        elif "heading 2" in sty: is_h2=True; is_h1=is_h3=False
        elif "heading 3" in sty: is_h3=True; is_h1=is_h2=False

        # Quotation rules
        quoted_oneliner = (len(words)<=20 and ((text.startswith('"') and text.endswith('"')) or (text.startswith("'") and text.endswith("'"))))
        is_quote = quoted_oneliner or ((len(words)<=60) and (align=="center" or any_b or any_i))

        rows.append({
            "idx": i,
            "text": text,
            "is_h1": bool(is_h1),
            "is_h2": bool(is_h2),
            "is_h3": bool(is_h3),
            "is_header": bool(is_h1 or is_h2 or is_h3),
            "is_quote": bool(is_quote and not (is_h1 or is_h2 or is_h3)),
            "avg_font": avg, "max_font": mx,
            "any_bold": any_b, "bold_fraction": round(bfrac,2),
            "any_italic": any_i, "italic_fraction": round(ifrac,2),
            "align": align, "style": sty,
            "word_count": len(words)
        })
    return rows
