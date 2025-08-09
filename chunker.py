import pandas as pd

def _make_chapter_name(h1, h2, h3):
    parts = [p for p in [h1, h2, h3] if p]
    return " | ".join(parts) if parts else "Introduction"

def _apply_overlap(chunks, ratio):
    if ratio <= 0 or len(chunks) <= 1:
        return chunks
    out = []
    for i, ch in enumerate(chunks):
        if i == 0:
            out.append(ch)
            continue
        prev_words = chunks[i-1].split()
        take = max(1, int(len(prev_words) * ratio))
        prefix = " ".join(prev_words[-take:]) if prev_words else ""
        out.append((prefix + "\n\n" if prefix else "") + ch)
    return out

def _flush_paragraph_group(par_group, book_name, author_name, h1, h2, h3,
                           min_words, max_words, overlap_ratio, out_rows):
    # Build non-overlapping paragraph chunks
    tmp, i = [], 0
    while i < len(par_group):
        if par_group[i]["words"] > max_words:
            tmp.append(par_group[i]["text"])
            i += 1
            continue
        merged = par_group[i]["text"]
        total = par_group[i]["words"]
        j = i + 1
        while j < len(par_group) and (total + par_group[j]["words"]) <= max_words:
            merged += "\n\n" + par_group[j]["text"]
            total += par_group[j]["words"]
            j += 1
        if total < min_words and j < len(par_group):
            merged += "\n\n" + par_group[j]["text"]
            total += par_group[j]["words"]
            j += 1
        tmp.append(merged)
        i = j

    # 20% overlap inside THIS section only
    final_chunks = _apply_overlap(tmp, overlap_ratio)

    for ch in final_chunks:
        out_rows.append({
            "book_name": book_name or "Unknown Book",
            "author_name": author_name or "Unknown Author",
            "h1": h1 or "", "h2": h2 or "", "h3": h3 or "",
            "chapter_name": _make_chapter_name(h1, h2, h3),
            "text_chunk": ch,
            "quotation": False
        })

def build_csv_rows(rows, book_name, author_name, min_words, max_words, overlap_ratio=0.20):
    out_rows = []
    h1 = h2 = h3 = None
    par_group = []

    def flush_group():
        nonlocal par_group
        if par_group:
            _flush_paragraph_group(par_group, book_name, author_name, h1, h2, h3,
                                   min_words, max_words, overlap_ratio, out_rows)
            par_group.clear()

    for row in rows:
        # New chapter → flush first (so there is NO overlap across chapters)
        if row.get("is_h1") or row.get("is_h2") or row.get("is_h3"):
            flush_group()
            if row.get("is_h1"):
                h1, h2, h3 = row["text"], None, None
            elif row.get("is_h2"):
                h2, h3 = row["text"], None
            elif row.get("is_h3"):
                h3 = row["text"]
            continue

        # Quotation → emit alone and break grouping
        if row.get("is_quote"):
            flush_group()
            out_rows.append({
                "book_name": book_name or "Unknown Book",
                "author_name": author_name or "Unknown Author",
                "h1": h1 or "", "h2": h2 or "", "h3": h3 or "",
                "chapter_name": _make_chapter_name(h1, h2, h3),
                "text_chunk": row["text"],
                "quotation": True
            })
            continue

        # Normal paragraph
        t = row["text"].strip()
        if t:
            par_group.append({"text": t, "words": len(t.split())})

    flush_group()
    return pd.DataFrame(out_rows)
