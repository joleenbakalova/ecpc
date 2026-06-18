#!/usr/bin/env python3
"""
Build a UNGA78-formatted, deduplicated, completeness-expanded CSV of ECPC
members' 2023-2026 publications.

- Discovery: free OpenAlex (no API key). For every member we union works found
  via author.orcid AND via the OpenAlex author.id, with a name+affiliation
  fallback when the ORCID is not linked in OpenAlex.
- Existing rows in ecpc_2023_2026.csv (already coded, often via Claude) are
  PRESERVED; only newly discovered works are coded with the keyword extractor.
- Output columns follow the UNGA78 coding sheet, with an Open Access column
  inserted right after the study-name column.
- Text cells are whitespace-cleaned; rows are sorted by surname -> year -> title.
"""
import csv
import re
import sys
import time
import datetime
from pathlib import Path

import requests

import pipeline  # reuse helpers
# Discovery window per user: 1 Aug 2023 -> today.
pipeline.DATE_FROM = "2023-08-01"
from pipeline import (
    load_members,
    find_openalex_author_id,
    fetch_works,
    reconstruct_abstract,
    keyword_extract,
)

HERE = Path(__file__).parent
EXISTING_CSV = HERE / "ecpc_2023_2026.csv"
OUT_CSV = HERE / "ecpc_unga_2023_2026.csv"
EMAIL = pipeline.OPENALEX_EMAIL

# ── text cleaning ────────────────────────────────────────────────────────────
def clean(s) -> str:
    if s is None:
        return ""
    s = str(s)
    s = s.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    s = s.replace("\xa0", " ").replace("​", "")
    s = re.sub(r"\s+", " ", s)               # collapse runs of whitespace
    s = re.sub(r"\s+([,.;:%])", r"\1", s)     # no space before punctuation
    return s.strip()

def yn(v) -> str:
    return "Yes" if str(v).strip().lower() in ("yes", "y", "true", "1") else ""

def norm_key(openalex_id="", doi="", title="", year=""):
    if openalex_id:
        return openalex_id.rsplit("/", 1)[-1].lower()
    if doi:
        return re.sub(r"^https?://(dx\.)?doi\.org/", "", doi.strip().lower())
    return f"{clean(title).lower()}|{year}"

# ── output schema (UNGA78-aligned leaf columns) ──────────────────────────────
COLUMNS = [
    "Study name",            # "Surname, First — Year — Title"
    "Open Access",           # NEW (inserted next to study name)
    "ECPC member", "ECPC affiliation",
    "Same sample as",
    "Country", "Setting/location", "Latitude", "Longitude",
    "Longitudinal", "Design", "Follow-up (months)", "Quant/qual", "Type",
    "Intervention name", "Control", "Randomized", "Sample size",
    "Age 0-8", "Age 9-25", "Adult", "Caregiver", "Teacher", "Key demographics",
    # Child outcomes
    "Child: Reading", "Child: Writing", "Child: Mathematics", "Child: Cognitive",
    "Child: Language", "Child: Executive function", "Child: Motor skills",
    "Child: Social-emotional", "Child: Mental health", "Child: Behavioral",
    "Child: Physical health", "Child: Victimization",
    # Caregiver / family
    "Caregiver: Mental health & well-being", "Caregiver: positive/responsive",
    "Caregiver: harsh/violent discipline", "Caregiver: family unit",
    "Caregiver outcomes (detail)",
    # Teacher / classroom / school
    "Teacher/classroom outcomes",
    # Biomarkers
    "Biomarkers",
    # Findings type
    "Findings: Psychometrics", "Findings: Descriptive", "Findings: Group differences",
    "Findings: Associations", "Findings: Moderation", "Findings: Mediation",
    "Findings: Intervention impact",
    # Key findings text
    "Key finding 1", "Key finding 2", "Key finding 3", "Effect sizes", "Notes",
    # Bibliographic tail (not in UNGA78 but needed for review/verification)
    "Authors", "Year", "Journal", "DOI", "OA URL", "OpenAlex ID", "Cited by", "Abstract",
]

# map existing-CSV header -> internal field
EXISTING_MAP = {
    "Ecpc Last": "last", "Ecpc First": "first", "Ecpc Affiliation": "affiliation",
    "Title": "title", "Authors": "authors", "Year": "year", "Journal": "journal",
    "Doi": "doi", "Oa Url": "oa_url", "Openalex Id": "openalex_id",
    "Cited By Count": "cited_by", "Country": "country",
    "Setting Location": "setting", "Latitude": "lat", "Longitude": "lon",
    "Longitudinal": "longitudinal", "Design": "design",
    "Followup Months": "followup", "Quant Qual": "quant_qual", "Type": "type",
    "Intervention Name": "intervention", "Control": "control",
    "Randomized": "randomized", "Sample Size": "sample_size",
    "Age 0 8": "age_0_8", "Age 9 25": "age_9_25", "Age Adult": "age_adult",
    "Age Caregiver": "age_caregiver", "Age Teacher": "age_teacher",
    "Demographics": "demographics",
    "Child Reading": "child_reading", "Child Writing": "child_writing",
    "Child Math": "child_math", "Child Cognitive": "child_cognitive",
    "Child Language": "child_language",
    "Child Executive Function": "child_executive_function",
    "Child Motor": "child_motor", "Child Social Emotional": "child_social_emotional",
    "Child Mental Health": "child_mental_health", "Child Behavioral": "child_behavioral",
    "Child Physical": "child_physical", "Child Victimization": "child_victimization",
    "Caregiver Outcomes": "caregiver_outcomes", "Teacher Outcomes": "teacher_outcomes",
    "Biomarkers": "biomarkers",
    "Findings Psychometrics": "f_psy", "Findings Descriptive": "f_desc",
    "Findings Group Differences": "f_grp", "Findings Associations": "f_assoc",
    "Findings Moderation": "f_mod", "Findings Mediation": "f_med",
    "Findings Intervention Impact": "f_int",
    "Key Finding 1": "kf1", "Key Finding 2": "kf2", "Key Finding 3": "kf3",
    "Effect Sizes": "effect_sizes", "Notes": "notes", "Abstract": "abstract",
}


def study_name(last, first, year, title):
    t = clean(title)
    if len(t) > 90:
        t = t[:90].rsplit(" ", 1)[0] + "…"
    return f"{clean(last)}, {clean(first)} — {year or 'n.d.'} — {t}"


def caregiver_cols(text):
    """Split caregiver_outcomes free text into UNGA78-style flags."""
    t = (text or "").lower()
    mh = "Yes" if any(k in t for k in ("mental health", "wellbeing", "well-being", "distress", "stress", "depress", "anxiety")) else ""
    pos = "Yes" if any(k in t for k in ("responsive", "sensitiv", "positive parenting", "warmth", "stimulation", "parenting practice")) else ""
    harsh = "Yes" if any(k in t for k in ("harsh", "violent", "punishment", "maltreat")) else ""
    fam = "Yes" if any(k in t for k in ("co-parent", "family unit", "family functioning", "paternal", "father involve", "connectedness", "conflict")) else ""
    return mh, pos, harsh, fam


def row_from_existing(f):
    last, first = f.get("last", ""), f.get("first", "")
    mh, pos, harsh, fam = caregiver_cols(f.get("caregiver_outcomes", ""))
    oa = "Yes" if clean(f.get("oa_url", "")) else "No"
    return {
        "Study name": study_name(last, first, f.get("year", ""), f.get("title", "")),
        "Open Access": oa,
        "ECPC member": clean(f"{last}, {first}"),
        "ECPC affiliation": clean(f.get("affiliation", "")),
        "Same sample as": "",
        "Country": clean(f.get("country", "")),
        "Setting/location": clean(f.get("setting", "")),
        "Latitude": clean(f.get("lat", "")), "Longitude": clean(f.get("lon", "")),
        "Longitudinal": clean(f.get("longitudinal", "")),
        "Design": clean(f.get("design", "")),
        "Follow-up (months)": clean(f.get("followup", "")),
        "Quant/qual": clean(f.get("quant_qual", "")),
        "Type": clean(f.get("type", "")),
        "Intervention name": clean(f.get("intervention", "")),
        "Control": clean(f.get("control", "")),
        "Randomized": clean(f.get("randomized", "")),
        "Sample size": clean(f.get("sample_size", "")),
        "Age 0-8": clean(f.get("age_0_8", "")), "Age 9-25": clean(f.get("age_9_25", "")),
        "Adult": clean(f.get("age_adult", "")), "Caregiver": clean(f.get("age_caregiver", "")),
        "Teacher": clean(f.get("age_teacher", "")),
        "Key demographics": clean(f.get("demographics", "")),
        "Child: Reading": yn(f.get("child_reading")), "Child: Writing": yn(f.get("child_writing")),
        "Child: Mathematics": yn(f.get("child_math")), "Child: Cognitive": yn(f.get("child_cognitive")),
        "Child: Language": yn(f.get("child_language")),
        "Child: Executive function": yn(f.get("child_executive_function")),
        "Child: Motor skills": yn(f.get("child_motor")),
        "Child: Social-emotional": yn(f.get("child_social_emotional")),
        "Child: Mental health": yn(f.get("child_mental_health")),
        "Child: Behavioral": yn(f.get("child_behavioral")),
        "Child: Physical health": yn(f.get("child_physical")),
        "Child: Victimization": yn(f.get("child_victimization")),
        "Caregiver: Mental health & well-being": mh,
        "Caregiver: positive/responsive": pos,
        "Caregiver: harsh/violent discipline": harsh,
        "Caregiver: family unit": fam,
        "Caregiver outcomes (detail)": clean(f.get("caregiver_outcomes", "")),
        "Teacher/classroom outcomes": clean(f.get("teacher_outcomes", "")),
        "Biomarkers": clean(f.get("biomarkers", "")),
        "Findings: Psychometrics": yn(f.get("f_psy")), "Findings: Descriptive": yn(f.get("f_desc")),
        "Findings: Group differences": yn(f.get("f_grp")), "Findings: Associations": yn(f.get("f_assoc")),
        "Findings: Moderation": yn(f.get("f_mod")), "Findings: Mediation": yn(f.get("f_med")),
        "Findings: Intervention impact": yn(f.get("f_int")),
        "Key finding 1": clean(f.get("kf1", "")), "Key finding 2": clean(f.get("kf2", "")),
        "Key finding 3": clean(f.get("kf3", "")), "Effect sizes": clean(f.get("effect_sizes", "")),
        "Notes": clean(f.get("notes", "")),
        "Authors": clean(f.get("authors", "")), "Year": clean(f.get("year", "")),
        "Journal": clean(f.get("journal", "")), "DOI": clean(f.get("doi", "")),
        "OA URL": clean(f.get("oa_url", "")), "OpenAlex ID": clean(f.get("openalex_id", "")),
        "Cited by": clean(f.get("cited_by", "")), "Abstract": clean(f.get("abstract", "")),
        "_sortkey": (clean(last).lower(), clean(first).lower(), str(f.get("year", "")), clean(f.get("title", "")).lower()),
    }


def row_from_work(member, work):
    last, first = member["last"], member["first"]
    abstract = reconstruct_abstract(work.get("abstract_inverted_index"))
    ex = keyword_extract(work, abstract)
    # bibliographic
    title = work.get("title") or ""
    year = work.get("publication_year") or ""
    doi = (work.get("doi") or "").replace("https://doi.org/", "")
    loc = work.get("primary_location") or {}
    src = loc.get("source") or {}
    journal = src.get("display_name") or ""
    oa = work.get("open_access") or {}
    is_oa = oa.get("is_oa")
    oa_url = oa.get("oa_url") or ""
    authors = ", ".join(
        (a.get("author") or {}).get("display_name", "")
        for a in (work.get("authorships") or [])
    )
    mh, pos, harsh, fam = caregiver_cols(ex.get("caregiver_outcomes", ""))
    first_sentence = re.split(r"(?<=[.!?])\s+", abstract.strip(), maxsplit=1)[0] if abstract else ""
    return {
        "Study name": study_name(last, first, year, title),
        "Open Access": "Yes" if is_oa else "No",
        "ECPC member": clean(f"{last}, {first}"),
        "ECPC affiliation": clean(member.get("affiliation", "")),
        "Same sample as": "",
        "Country": "", "Setting/location": "", "Latitude": "", "Longitude": "",
        "Longitudinal": clean(ex.get("longitudinal", "")),
        "Design": clean(ex.get("design", "")),
        "Follow-up (months)": "",
        "Quant/qual": clean(ex.get("quant_qual", "")),
        "Type": clean(ex.get("type", "")),
        "Intervention name": "", "Control": "",
        "Randomized": clean(ex.get("randomized", "")),
        "Sample size": "",
        "Age 0-8": "", "Age 9-25": "", "Adult": "", "Caregiver": "", "Teacher": "",
        "Key demographics": "",
        "Child: Reading": yn(ex.get("child_reading")), "Child: Writing": yn(ex.get("child_writing")),
        "Child: Mathematics": yn(ex.get("child_math")), "Child: Cognitive": yn(ex.get("child_cognitive")),
        "Child: Language": yn(ex.get("child_language")),
        "Child: Executive function": yn(ex.get("child_executive_function")),
        "Child: Motor skills": yn(ex.get("child_motor")),
        "Child: Social-emotional": yn(ex.get("child_social_emotional")),
        "Child: Mental health": yn(ex.get("child_mental_health")),
        "Child: Behavioral": yn(ex.get("child_behavioral")),
        "Child: Physical health": yn(ex.get("child_physical")),
        "Child: Victimization": yn(ex.get("child_victimization")),
        "Caregiver: Mental health & well-being": mh,
        "Caregiver: positive/responsive": pos,
        "Caregiver: harsh/violent discipline": harsh,
        "Caregiver: family unit": fam,
        "Caregiver outcomes (detail)": clean(ex.get("caregiver_outcomes", "")),
        "Teacher/classroom outcomes": clean(ex.get("teacher_outcomes", "")),
        "Biomarkers": clean(ex.get("biomarkers", "")),
        "Findings: Psychometrics": yn(ex.get("findings_psychometrics")),
        "Findings: Descriptive": yn(ex.get("findings_descriptive")),
        "Findings: Group differences": yn(ex.get("findings_group_differences")),
        "Findings: Associations": yn(ex.get("findings_associations")),
        "Findings: Moderation": yn(ex.get("findings_moderation")),
        "Findings: Mediation": yn(ex.get("findings_mediation")),
        "Findings: Intervention impact": yn(ex.get("findings_intervention_impact")),
        "Key finding 1": clean(first_sentence), "Key finding 2": "", "Key finding 3": "",
        "Effect sizes": "",
        "Notes": "[auto-coded from title+abstract; verify location/sample/findings against full text]",
        "Authors": clean(authors), "Year": clean(str(year)), "Journal": clean(journal),
        "DOI": clean(doi), "OA URL": clean(oa_url),
        "OpenAlex ID": clean(work.get("id", "")), "Cited by": clean(str(work.get("cited_by_count", ""))),
        "Abstract": clean(abstract),
        "_sortkey": (clean(last).lower(), clean(first).lower(), str(year), clean(title).lower()),
    }


def author_id_from_orcid(orcid):
    orcid = orcid.replace("https://orcid.org/", "").strip()
    try:
        r = requests.get("https://api.openalex.org/authors",
                         params={"filter": f"orcid:{orcid}", "mailto": EMAIL}, timeout=20)
        r.raise_for_status()
        res = r.json().get("results", [])
        if res:
            return res[0]["id"].rsplit("/", 1)[-1]
    except Exception as e:
        print(f"    [warn] orcid->author failed: {e}")
    return None


def discover(member):
    """Return list of OpenAlex work dicts for a member (orcid ∪ author.id)."""
    orcid = (member.get("orcid") or "").strip()
    aid = None
    if orcid:
        aid = author_id_from_orcid(orcid)
    if not aid:
        aid = find_openalex_author_id(member["last"], member["first"], member.get("affiliation", ""))
        if aid:
            aid = aid.rsplit("/", 1)[-1]
    works = {}
    if orcid:
        for w in fetch_works("", orcid):
            works[w["id"]] = w
    if aid:
        for w in fetch_works(aid, None):
            works.setdefault(w["id"], w)
    return list(works.values()), orcid, aid


def main():
    members = load_members(str(HERE / "members.csv"), str(HERE / "orcid_ids.csv"))

    # 1) load existing coded rows
    existing = {}
    with open(EXISTING_CSV, newline="", encoding="utf-8") as fh:
        for raw in csv.DictReader(fh):
            f = {EXISTING_MAP[k]: v for k, v in raw.items() if k in EXISTING_MAP}
            key = norm_key(f.get("openalex_id", ""), f.get("doi", ""), f.get("title", ""), f.get("year", ""))
            existing[key] = f

    records = {}   # key -> output row
    # seed with existing (preserve richer coding)
    for key, f in existing.items():
        records[key] = row_from_existing(f)

    # 2) discover for every member, add NEW works only
    summary = []
    for i, m in enumerate(members, 1):
        name = f"{m['last']}, {m['first']}"
        try:
            works, orcid, aid = discover(m)
        except Exception as e:
            print(f"[{i:2}/{len(members)}] {name:28s} ERROR {e}")
            summary.append((name, "ERR", 0, 0))
            continue
        added = 0
        for w in works:
            doi = (w.get("doi") or "").replace("https://doi.org/", "")
            key = norm_key(w.get("id", ""), doi, w.get("title", ""), w.get("publication_year", ""))
            if key in records:
                continue
            records[key] = row_from_work(m, w)
            added += 1
        print(f"[{i:2}/{len(members)}] {name:28s} found={len(works):3d} new={added:3d} "
              f"(orcid={'y' if orcid else '-'} aid={'y' if aid else '-'})")
        summary.append((name, "ok", len(works), added))
        time.sleep(0.15)

    # 3) sort + write
    rows = sorted(records.values(), key=lambda r: r["_sortkey"])
    with open(OUT_CSV, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=COLUMNS, extrasaction="ignore")
        w.writeheader()
        for r in rows:
            w.writerow(r)

    print("\n" + "=" * 60)
    print(f"Wrote {len(rows)} rows -> {OUT_CSV.name}")
    members_with = len({r['ECPC member'] for r in rows})
    print(f"Members represented: {members_with} / {len(members)}")
    print("=" * 60)


if __name__ == "__main__":
    main()
