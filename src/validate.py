 # src/validate.py
from pathlib import Path
import json
import pandas as pd

POLICY = Path("reports/policy/policy.json")
FEATURES_DIR = Path("reports/features")
OUT_DIR = Path("reports/validation")

def _load_policy():
    with open(POLICY, "r", encoding="utf-8") as f:
        return json.load(f)

def _ok_color(c, allowed):
    return (c is None) or (c.upper() in {x.upper() for x in allowed})

def validate_file(features_csv: Path, policy: dict) -> pd.DataFrame:
    df = pd.read_csv(features_csv)
    vio_rows = []

    # DOCX text
    docx_rules = policy["fonts"].get("docx", {})
    body_rule = docx_rules.get("body", {})
    allowed_colors = policy.get("colors_rgb", [])

    is_docx = features_csv.suffix == ".csv" and "docx" in features_csv.stem.lower()

    for _, r in df.iterrows():
        if r["kind"] in ("docx_run", "pptx_run"):
            fam = str(r.get("font_family") or "")
            size = r.get("font_size_pt")
            col = r.get("color_rgb")

            if r["kind"] == "docx_run" and body_rule:
                if fam and body_rule.get("family") and fam != body_rule["family"]:
                    vio_rows.append((r["file"], "font_family", fam, body_rule["family"]))
                if pd.notna(size) and body_rule.get("size_pt") and float(size) != float(body_rule["size_pt"]):
                    vio_rows.append((r["file"], "font_size_pt", size, body_rule["size_pt"]))
                if not _ok_color(col, allowed_colors):
                    vio_rows.append((r["file"], "color_rgb", col, "palette"))

            if r["kind"] == "pptx_run":
                pptx = policy["fonts"].get("pptx", {})
                min_body = pptx.get("body_min_pt")
                fam_req = pptx.get("family")
                # naive heuristic: treat short runs as body unless size is big; tune in notebook
                if fam_req and fam and fam != fam_req:
                    vio_rows.append((r["file"], "font_family", fam, fam_req))
                if pd.notna(size) and min_body and float(size) < float(min_body):
                    vio_rows.append((r["file"], "font_size_pt", size, f">={min_body}"))
                if not _ok_color(col, allowed_colors):
                    vio_rows.append((r["file"], "color_rgb", col, "palette"))

        if r["kind"] == "pptx_image" and policy.get("logo", {}).get("pptx_title_slide_required"):
            if int(r.get("slide_idx", 0)) == 1:
                ref = policy["logo"].get("sha1")
                if ref:
                    if str(r.get("sha1")) != ref:
                        vio_rows.append((r["file"], "logo_sha1", r.get("sha1"), ref))

    out = pd.DataFrame(vio_rows, columns=["file", "rule", "observed", "expected"])
    return out

def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    policy = _load_policy()
    summary = []
    for csv in FEATURES_DIR.glob("*.features.csv"):
        vdf = validate_file(csv, policy)
        vdf.to_csv(OUT_DIR / f"{csv.stem}.violations.csv", index=False)
        summary.append({"file": csv.stem, "violations": len(vdf)})
    pd.DataFrame(summary).to_csv(OUT_DIR / "summary.csv", index=False)

if __name__ == "__main__":
    main()

