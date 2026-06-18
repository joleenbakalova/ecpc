#!/usr/bin/env python3
"""Merge in-session extraction JSON (enrich_out/*.json) back into the
intermediate CSV. Only fills blank cells; refines design/type when current
value is 'unclear'. Computes lat/long from country centroid when blank."""
import csv, json, glob
from pathlib import Path

HERE = Path(__file__).parent
SRC = HERE / "ecpc_unga_2023_2026.csv"
OUTDIR = HERE / "enrich_out"

# country -> (lat, lon) approximate centroids (covers this dataset)
CENTROID = {
    "afghanistan": (33.94, 67.71), "azerbaijan": (40.14, 47.58),
    "bangladesh": (23.68, 90.36), "bhutan": (27.51, 90.43),
    "bolivia": (-16.29, -63.59), "brazil": (-14.24, -51.93),
    "burundi": (-3.37, 29.92), "cambodia": (12.57, 104.99),
    "canada": (56.13, -106.35), "chile": (-35.68, -71.54),
    "china": (35.86, 104.20), "colombia": (4.57, -74.30),
    "democratic republic of the congo": (-4.04, 21.76), "drc": (-4.04, 21.76),
    "egypt": (26.82, 30.80), "ethiopia": (9.15, 40.49),
    "germany": (51.17, 10.45), "ghana": (7.95, -1.02),
    "haiti": (18.97, -72.29), "india": (20.59, 78.96),
    "indonesia": (-0.79, 113.92), "iran": (32.43, 53.69),
    "iraq": (33.22, 43.68), "ireland": (53.41, -8.24),
    "northern ireland": (54.79, -6.49), "israel": (31.05, 34.85),
    "italy": (41.87, 12.57), "jordan": (30.59, 36.24),
    "kenya": (-0.02, 37.91), "kyrgyzstan": (41.20, 74.77),
    "lebanon": (33.85, 35.86), "liberia": (6.43, -9.43),
    "moldova": (47.41, 28.37), "nepal": (28.39, 84.12),
    "niger": (17.61, 8.08), "nigeria": (9.08, 8.68),
    "norway": (60.47, 8.47), "pakistan": (30.38, 69.35),
    "rwanda": (-1.94, 29.87), "serbia": (44.02, 21.01),
    "sierra leone": (8.46, -11.78), "south africa": (-30.56, 22.94),
    "south sudan": (6.88, 31.31), "sri lanka": (7.87, 80.77),
    "syria": (34.80, 38.997), "thailand": (15.87, 100.99),
    "turkey": (38.96, 35.24), "uganda": (1.37, 32.29),
    "ukraine": (48.38, 31.17), "united arab emirates": (23.42, 53.85),
    "uk": (55.38, -3.44), "united kingdom": (55.38, -3.44),
    "england": (52.36, -1.17), "u.s.": (37.09, -95.71),
    "usa": (37.09, -95.71), "united states": (37.09, -95.71),
    "vietnam": (14.06, 108.28), "zambia": (-13.13, 27.85),
    "global": ("", ""), "multiple": ("", ""),
}

# extraction-key -> intermediate CSV column
FIELD_COL = {
    "country": "Country", "setting": "Setting/location",
    "longitudinal": "Longitudinal", "design": "Design",
    "followup_months": "Follow-up (months)", "quant_qual": "Quant/qual",
    "type": "Type", "intervention_name": "Intervention name",
    "control": "Control", "randomized": "Randomized", "sample_size": "Sample size",
    "age_0_8": "Age 0-8", "age_9_25": "Age 9-25", "age_adult": "Adult",
    "age_caregiver": "Caregiver", "age_teacher": "Teacher",
    "demographics": "Key demographics", "caregiver_detail": "Caregiver outcomes (detail)",
    "teacher_detail": "Teacher/classroom outcomes", "biomarkers": "Biomarkers",
    "kf1": "Key finding 1", "kf2": "Key finding 2", "kf3": "Key finding 3",
    "effect_sizes": "Effect sizes",
}
# fields allowed to OVERRIDE an existing 'unclear'/blank value
REFINABLE = {"design", "quant_qual", "type", "longitudinal", "randomized"}


def rowkey(r, idx):
    return r.get("OpenAlex ID", "").strip() or r.get("DOI", "").strip() or f"row{idx}"


def main():
    rows = list(csv.DictReader(open(SRC, encoding="utf-8-sig")))
    fieldnames = rows[0].keys()
    # load all extraction outputs
    ext = {}
    for fp in sorted(glob.glob(str(OUTDIR / "*.json"))):
        for k, v in json.load(open(fp, encoding="utf-8")).items():
            ext[k] = v
    print(f"Loaded {len(ext)} extraction records from {OUTDIR.name}/")

    applied = 0; cells = 0
    for idx, r in enumerate(rows):
        e = ext.get(rowkey(r, idx))
        if not e:
            continue
        applied += 1
        for fk, col in FIELD_COL.items():
            val = (e.get(fk) or "").strip() if isinstance(e.get(fk), str) else e.get(fk)
            if not val:
                continue
            cur = r.get(col, "").strip()
            if not cur or (fk in REFINABLE and cur.lower() in ("unclear", "")):
                if r.get(col, "").strip() != str(val):
                    r[col] = str(val); cells += 1
        # lat/long from country
        if r.get("Country", "").strip() and not r.get("Latitude", "").strip():
            c0 = r["Country"].split(";")[0].split(",")[0].strip().lower()
            if c0 in CENTROID:
                lat, lon = CENTROID[c0]
                if lat != "":
                    r["Latitude"] = str(lat); r["Longitude"] = str(lon); cells += 1

    with open(SRC, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=fieldnames)
        w.writeheader(); w.writerows(rows)
    print(f"Applied extractions to {applied} rows, updated {cells} cells.")
    print(f"Rewrote {SRC.name}. Now re-run: python3 to_unga_exact.py")


if __name__ == "__main__":
    main()
