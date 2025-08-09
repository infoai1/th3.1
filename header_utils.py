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
    try:
        return float(size.pt) if size else None
    except Exception:
        return None

def _is_num(x): return isinstance(x, (int, float))
def _f(x, d=None): return float(x) if _is_num(x) else d
def _r(x): return round(float(x), 2) if _is_num(x) else None

def _any_bold(p): return any(r.bold for r in p.runs if r.text)
def _bold_fraction(p):
    total = sum(len(r.text or "") for r in p.runs) or 1
    bold = sum(len(r.text or "") for r in p.runs if r.bold)
    return bold / total

def _any_italic(p): return any(r.italic for r in p.runs if r.text)
def _italic_fraction(p):
    total = sum(len(r.text or "") for r in p.runs) or 1
    it = sum(len(r.text or "") for r in p.runs if r.italic)
    return it / total

def _avg_font(p):
    sizes = [_pt(r.font.size) for r in p.runs if r.font and r.font.size]
    sizes = [s for s in sizes if _is_num(s)]
    return sum(sizes) / len(sizes) if sizes else None

def _max_font(p):
    sizes = [_pt(r.font.size) for r in p.runs if r.font and r.font.size]
    sizes = [s for s in sizes if _is_num(s)]
    return max(sizes) if sizes else None

def _align(p): return ALIGN_MAP.get(p.alignment, "left")
def _style(p):
    try: return (p.style.name or "").lower()
    except Exception: return ""

def _looks_sentence_like(text):
    t = (text or "").strip()
    if not t: return True
    end = t.endswith((".", "!", "?", '"', "'"))
    has_punct = sum(t.count(x) for x in ".!?") >= 1
    longish = len(t.split()) >= 10
    return end or has_punct or longish

def _is_quoted_oneliner(text):
    t = (text or "").strip()
    return len(t.split()) <= 20 and (
        (t.startswith('"') and t.endswith('"')) or (t.startswith("'") and t.endswith("'"))
    )

def _quote_flag(text, f, qr):
    wc, align = f["word_count"], f["align"]
    if qr.get("quoted_one_liners", True) and f["quoted_oneliner"]:
        return True
    if wc <= qr.get("short_word_cutoff", 60):
        if qr.get("centered_short", True) and align == "center": return True
        if qr.get("bold_short", True) and (f["any_bold"] or f["bold_fraction"] >= 0.6): return True
        if qr.get("italic_short", True) and (f["any_italic"] or f["italic_fraction"] >= 0.6): return True
    return False

def classify_levels_and_features(paragraph, text, rules, quote_rules):
    words = text.split()
    word_count = len(words)
    all_caps = text.isupper() and any(c.isalpha() for c in text)
    short_phrase = word_count <= rules.get("max_header_words", 15) and len(text) <= 120
    avg_size = _avg_font(paragraph)
    max_size = _max_font(paragraph)
    any_bold = _any_bold(paragraph)
    bold_frac = _bold_fraction(paragraph)
    any_italic = _any_italic(paragraph)
    italic_frac = _italic_fraction(paragraph)
    align = _align(paragraph)
    style = _style(paragraph)
    sentence_like = _looks_sentence_like(text) if rules.get("suppress_sentences", True) else False
    quoted_oneliner = _is_quoted_oneliner(text) if rules.get("suppress_quotes", True) else False

    style_h1 = "heading 1" in style or style.strip() == "heading1"
    style_h2 = "heading 2" in style or style.strip() == "heading2"
    style_h3 = "heading 3" in style or style.strip() == "heading3"

    def level_match(key):
        lvl = rules["levels"][key]
        min_size = _f(lvl.get("min_size"), 13)
        avg_ok = _is_num(avg_size) and _is_num(min_size) and avg_size >= min_size
        max_ok = _is_num(max_size) and _is_num(min_size) and max_size >= min_size
        size_ok = avg_ok or max_ok
        if not lvl.get("enabled", True) or not size_ok: return False, 0
        if align not in lvl.get("allowed_align", ["left", "center", "right"]): return False, 0
        if lvl.get("require_bold", False) and not any_bold and bold_frac < 0.4: return False, 0
        if lvl.get("require_short_phrase", True) and not short_phrase: return False, 0

        score = 0
        if any_bold or bold_frac >= 0.6: score += 1
        if all_caps: score += 1
        if align == "center": score += 1
        if short_phrase: score += 1
        if "heading" in style: score += 2
        if sentence_like: score -= 2
        if quoted_oneliner: score -= 2
        return True, score

    is_h1, is_h2, is_h3, score = style_h1, style_h2, style_h3, 0
    if not any([is_h1, is_h2, is_h3]):
        for k in ["h1", "h2", "h3"]:
            ok, sc = level_match(k)
            if ok:
                if k == "h1": is_h1 = True
                if k == "h2": is_h2 = True
                if k == "h3": is_h3 = True
                score = max(score, sc)

    features = {
        "all_caps": all_caps,
        "short_phrase": short_phrase,
        "avg_font_size": _r(avg_size),
        "max_font_size": _r(max_size),
        "bold_fraction": _r(bold_frac),
        "any_bold": any_bold,
        "italic_fraction": _r(italic_frac),
        "any_italic": any_italic,
        "align": align,
        "style": style,
        "sentence_like": sentence_like,
        "quoted_oneliner": quoted_oneliner,
        "word_count": word_count,
    }

    is_quote = _quote_flag(text, features, quote_rules)
    is_header = any([is_h1, is_h2, is_h3])
    return {"is_h1": is_h1, "is_h2": is_h2, "is_h3": is_h3,
            "is_header": is_header, "is_quote": is_quote,
            "score": score, "features": features}

def parse_docx(docx_file, rules, quote_rules):
    doc = Document(docx_file)
    rows = []
    for idx, p in enumerate(doc.paragraphs):
        text = (p.text or "").strip()
        if not text: continue
        c = classify_levels_and_features(p, text, rules, quote_rules)
        f = c["features"]
        rows.append({
            "idx": idx, "text": text,
            "is_h1": c["is_h1"], "is_h2": c["is_h2"], "is_h3": c["is_h3"],
            "is_header": c["is_header"], "is_quote": c["is_quote"],
            "score": c["score"],
            "all_caps": f["all_caps"], "short_phrase": f["short_phrase"],
            "avg_font_size": f["avg_font_size"], "max_font_size": f["max_font_size"],
            "bold_fraction": f["bold_fraction"], "any_bold": f["any_bold"],
            "italic_fraction": f["italic_fraction"], "any_italic": f["any_italic"],
            "align": f["align"], "style": f["style"],
            "sentence_like": f["sentence_like"], "quoted_oneliner": f["quoted_oneliner"],
            "word_count": f["word_count"],
        })
    return rows
