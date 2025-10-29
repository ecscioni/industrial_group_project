 # src/extract_docx.py
from pathlib import Path
from hashlib import sha1
import pandas as pd
from docx import Document

def _pt(val):
    try:
        return float(val.pt)
    except Exception:
        return None

def _rgb(color):
    try:
        if color and color.rgb:
            return str(color.rgb)  # e.g., '000000'
    except Exception:
        pass
    return None

def extract_features_docx(path: Path) -> pd.DataFrame:
    doc = Document(path)
    rows = []

    # paragraphs
    for p_idx, para in enumerate(doc.paragraphs):
        p_font = getattr(para, "style", None)
        p_font = getattr(p_font, "font", None)

        for r_idx, run in enumerate(para.runs):
            rf = run.font
            fam = rf.name or (p_font.name if p_font else None)
            size = _pt(rf.size) or (_pt(p_font.size) if p_font else None)
            rows.append({
                "file": path.name,
                "kind": "docx_run",
                "page_like": None,  # Word has no fixed pages
                "para_idx": p_idx,
                "run_idx": r_idx,
                "text": run.text,
                "font_family": fam,
                "font_size_pt": size,
                "bold": bool(rf.bold),
                "italic": bool(rf.italic),
                "underline": bool(rf.underline),
                "color_rgb": _rgb(rf.color),
            })

    # pictures (logo detection support)
    # basic: hash all embedded images
    try:
        part = doc.part
        for rel in part.rels.values():
            if rel.reltype.endswith("/image"):
                blob = rel._target.blob
                rows.append({
                    "file": path.name,
                    "kind": "docx_image",
                    "sha1": sha1(blob).hexdigest(),
                    "para_idx": None,
                    "run_idx": None,
                    "text": None,
                    "font_family": None,
                    "font_size_pt": None,
                    "bold": None,
                    "italic": None,
                    "underline": None,
                    "color_rgb": None,
                })
    except Exception:
        pass

    return pd.DataFrame(rows)

