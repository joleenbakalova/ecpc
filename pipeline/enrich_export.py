#!/usr/bin/env python3
"""Export rows needing enrichment (blank core fields + have abstract) into
batch JSON files for in-session extraction."""
import csv, json
from pathlib import Path

HERE = Path(__file__).parent
SRC = HERE / "ecpc_unga_2023_2026.csv"
OUTDIR = HERE / "enrich_batches"
OUTDIR.mkdir(exist_ok=True)
BATCH = 30

def blank(r, k): return not r.get(k, "").strip()

rows = list(csv.DictReader(open(SRC, encoding="utf-8-sig")))
work = []
for idx, r in enumerate(rows):
    core_blank = blank(r, "Country") or blank(r, "Sample size") or blank(r, "Key demographics")
    if core_blank and r.get("Abstract", "").strip():
        key = r.get("OpenAlex ID", "").strip() or r.get("DOI", "").strip() or f"row{idx}"
        work.append({
            "key": key,
            "member": r.get("ECPC member", ""),
            "title": r.get("Study name", "").split(" — ", 2)[-1],
            "journal": r.get("Journal", ""),
            "year": r.get("Year", ""),
            "abstract": r.get("Abstract", ""),
        })

# write batches
n = 0
for i in range(0, len(work), BATCH):
    n += 1
    batch = work[i:i+BATCH]
    (OUTDIR / f"batch_{n:02d}.json").write_text(
        json.dumps(batch, ensure_ascii=False, indent=1), encoding="utf-8")
print(f"{len(work)} rows -> {n} batches of up to {BATCH} in {OUTDIR.name}/")
