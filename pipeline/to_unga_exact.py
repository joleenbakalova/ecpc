#!/usr/bin/env python3
"""
Transform the intermediate ECPC dataset (ecpc_unga_2023_2026.csv) into the
EXACT 99-column UNGA78 coding layout, so rows paste directly under the
existing UNGA78 sheet.

- The 8 UNGA78 header rows are copied verbatim from the source file.
- Each ECPC field is mapped to its precise UNGA78 column index.
- Open Access + bibliographic identifiers are appended as EXTRA columns to the
  right of column 99 (harmless to the paste alignment; delete if unwanted).
"""
import csv
from pathlib import Path

HERE = Path(__file__).parent
SRC_UNGA = Path("/Users/joleenbakalova/Downloads/UNGA78 study coding_02SEP2023.xlsx - All (1).csv")
INTERMEDIATE = HERE / "ecpc_unga_2023_2026.csv"
OUT = HERE / "ecpc_unga_FILLED.csv"

NCOLS = 99
HEADER_ROWS = 8

# extra reference columns appended after the 99 UNGA columns
EXTRA = ["Open Access", "ECPC member", "ECPC affiliation", "Authors", "Year",
         "Journal", "DOI", "OA URL", "OpenAlex ID", "Cited by", "Abstract"]


def flag(v):
    return "Yes" if str(v).strip() else ""


def notes_cell(r):
    parts = []
    for k in ("Key finding 1", "Key finding 2", "Key finding 3"):
        if r.get(k, "").strip():
            parts.append(r[k].strip())
    if r.get("Effect sizes", "").strip():
        parts.append("Effect sizes: " + r["Effect sizes"].strip())
    if r.get("Notes", "").strip():
        parts.append(r["Notes"].strip())
    return " | ".join(parts)


def to_row(r):
    c = [""] * NCOLS
    c[0] = r.get("Study name", "")
    # 1 Same sample as -> blank
    c[2] = r.get("Country", "")
    c[3] = r.get("Setting/location", "")
    c[4] = r.get("Latitude", "")
    c[5] = r.get("Longitude", "")
    c[6] = r.get("Longitudinal", "")
    c[7] = r.get("Design", "")
    c[8] = r.get("Follow-up (months)", "")
    c[9] = r.get("Quant/qual", "")
    c[10] = r.get("Type", "")
    c[11] = r.get("Intervention name", "")
    c[12] = r.get("Control", "")
    c[13] = r.get("Randomized", "")
    c[16] = r.get("Sample size", "")
    c[17] = r.get("Age 0-8", "")
    c[18] = r.get("Age 9-25", "")
    c[19] = r.get("Adult", "")
    c[20] = r.get("Caregiver", "")
    c[21] = r.get("Teacher", "")
    c[22] = r.get("Key demographics", "")
    # child outcomes
    c[23] = r.get("Child: Reading", "")
    c[24] = r.get("Child: Writing", "")
    c[25] = r.get("Child: Mathematics", "")
    c[26] = r.get("Child: Cognitive", "")
    c[27] = r.get("Child: Language", "")
    c[28] = r.get("Child: Executive function", "")
    c[29] = r.get("Child: Motor skills", "")
    c[30] = r.get("Child: Social-emotional", "")
    c[31] = r.get("Child: Mental health", "")
    c[32] = r.get("Child: Behavioral", "")
    # 33 internalizing / 34 externalizing -> left blank (not separately coded)
    c[35] = r.get("Child: Physical health", "")
    c[36] = r.get("Child: Victimization", "")
    # caregiver / family
    c[39] = r.get("Caregiver: Mental health & well-being", "")
    c[43] = r.get("Caregiver: positive/responsive", "")
    c[44] = r.get("Caregiver: harsh/violent discipline", "")
    c[45] = r.get("Caregiver: family unit", "")
    c[46] = r.get("Caregiver outcomes (detail)", "")
    # teacher
    c[54] = flag(r.get("Teacher/classroom outcomes", ""))
    c[56] = r.get("Teacher/classroom outcomes", "")
    # biomarkers
    c[60] = r.get("Biomarkers", "")
    # findings type flags
    c[62] = flag(r.get("Findings: Psychometrics", ""))
    c[63] = flag(r.get("Findings: Descriptive", ""))
    c[64] = flag(r.get("Findings: Group differences", ""))
    c[66] = flag(r.get("Findings: Associations", ""))   # Associations > Positive col
    c[67] = flag(r.get("Findings: Moderation", ""))
    c[68] = flag(r.get("Findings: Mediation", ""))
    c[69] = flag(r.get("Findings: Intervention impact", ""))
    # 95 figure / 96 framework -> unknown, blank
    c[97] = notes_cell(r)
    # extras
    extra = [
        r.get("Open Access", ""), r.get("ECPC member", ""), r.get("ECPC affiliation", ""),
        r.get("Authors", ""), r.get("Year", ""), r.get("Journal", ""), r.get("DOI", ""),
        r.get("OA URL", ""), r.get("OpenAlex ID", ""), r.get("Cited by", ""), r.get("Abstract", ""),
    ]
    return c + extra


def main():
    # 1) verbatim header rows
    src = list(csv.reader(open(SRC_UNGA, encoding="utf-8")))
    headers = []
    for i in range(HEADER_ROWS):
        row = src[i][:NCOLS]
        row += [""] * (NCOLS - len(row))
        if i == 0:
            row = row + EXTRA           # label extras on the first header row
        else:
            row = row + [""] * len(EXTRA)
        headers.append(row)

    # 2) data
    data = list(csv.DictReader(open(INTERMEDIATE, encoding="utf-8-sig")))
    out_rows = [to_row(r) for r in data]

    with open(OUT, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        for h in headers:
            w.writerow(h)
        for r in out_rows:
            w.writerow(r)

    print(f"Wrote {len(out_rows)} data rows + {HEADER_ROWS} header rows -> {OUT.name}")
    print(f"Columns: {NCOLS} UNGA78 + {len(EXTRA)} reference = {NCOLS + len(EXTRA)}")
    # sanity: every row width identical
    widths = {len(r) for r in out_rows} | {len(h) for h in headers}
    print("All row widths:", widths)


if __name__ == "__main__":
    main()
