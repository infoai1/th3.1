
import pandas as pd

def _chapter_name(h1,h2,h3):
    parts=[p for p in [h1,h2,h3] if p]
    return " | ".join(parts) if parts else "Introduction"

def _emit_group(group, book, author, h1,h2,h3, min_w, max_w, overlap_ratio, out):
    # Build chunks by merging paragraphs
    tmp=[]; i=0
    while i < len(group):
        if group[i]["w"]>max_w:
            tmp.append(group[i]["t"]); i+=1; continue
        merged=group[i]["t"]; total=group[i]["w"]; j=i+1
        while j<len(group) and total+group[j]["w"]<=max_w:
            merged += "\n\n"+group[j]["t"]; total += group[j]["w"]; j+=1
        if total<min_w and j<len(group):
            merged += "\n\n"+group[j]["t"]; total += group[j]["w"]; j+=1
        tmp.append(merged); i=j

    # 20% overlap inside section
    final=[]
    for k,ch in enumerate(tmp):
        if k==0: final.append(ch); continue
        prev_words=tmp[k-1].split()
        take=max(1,int(len(prev_words)*overlap_ratio))
        prefix=" ".join(prev_words[-take:])
        final.append(prefix+"\n\n"+ch)

    for ch in final:
        out.append({
            "book_name": book, "author_name": author,
            "h1": h1 or "", "h2": h2 or "", "h3": h3 or "",
            "chapter_name": _chapter_name(h1,h2,h3),
            "text_chunk": ch, "quotation": False
        })

def build_csv_rows(rows, book, author, min_w=180, max_w=250, overlap_ratio=0.2):
    out=[]; h1=h2=h3=None; group=[]
    def flush():
        nonlocal group
        if group:
            _emit_group(group, book, author, h1,h2,h3, min_w, max_w, overlap_ratio, out)
            group=[]
    for r in rows:
        if r.get("is_h1") or r.get("is_h2") or r.get("is_h3"):
            flush()
            if r.get("is_h1"): h1=r["text"]; h2=h3=None
            elif r.get("is_h2"): h2=r["text"]; h3=None
            elif r.get("is_h3"): h3=r["text"]
            continue
        if r.get("is_quote"):
            flush()
            out.append({
                "book_name": book, "author_name": author,
                "h1": h1 or "", "h2": h2 or "", "h3": h3 or "",
                "chapter_name": _chapter_name(h1,h2,h3),
                "text_chunk": r["text"], "quotation": True
            })
            continue
        t=r["text"].strip()
        if t: group.append({"t":t,"w":len(t.split())})
    flush()
    return pd.DataFrame(out)
