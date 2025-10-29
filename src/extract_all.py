 # src/extract_all.py
from pathlib import Path
import pandas as pd
from tqdm import tqdm
from extract_docx import extract_features_docx
from extract_pptx import extract_features_pptx

RAW = Path("data/raw")
OUT = Path("reports/features")

def main():
    OUT.mkdir(parents=True, exist_ok=True)
    files = list(RAW.glob("*.docx")) + list(RAW.glob("*.pptx"))
    for f in tqdm(files, desc="Extracting"):
        if f.suffix.lower() == ".docx":
            df = extract_features_docx(f)
        else:
            df = extract_features_pptx(f)
        df.to_csv(OUT / f"{f.stem}.features.csv", index=False)

if __name__ == "__main__":
    main()

