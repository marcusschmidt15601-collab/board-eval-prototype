"""calc.py — raw questionnaire export + structure.json  ->  results.json
All figures in the report come from here. Method (agreed with the consultant's calculation sheet):
  * mean = arithmetic mean of answers; spread = SAMPLE standard deviation (Excel STDEV / STDEV.S)
  * subarea scores pool all answers of the subarea's items
  * peer ratings exclude the member's rating of themselves -> divided by (n-1)
Deliberate corrections vs. the legacy sheet are listed in METHOD_NOTES and written to results.json.
"""
import json, re, sys, statistics as st, openpyxl

RAW = sys.argv[1] if len(sys.argv) > 1 else "input/raw.xlsx"
S = json.load(open("structure.json"))
ws = openpyxl.load_workbook(RAW, data_only=True).worksheets[0]
HDR = [c.value for c in ws[1]]
ROWS = [[c.value for c in r] for r in ws.iter_rows(min_row=2) if r[S["respondent_name_col"]].value]
NAMES = [r[S["respondent_name_col"]] for r in ROWS]
N = len(ROWS)

METHOD_NOTES = [
  "Spredning = stikprøvens standardafvigelse (STDEV.S).",
  "Anciennitet: gennemsnit over alle medlemmer (regnearket brugte F6:F12 og udelod to medlemmer).",
  "Spørgsmålet om skriftlige spilleregler (ja/nej) indgår ikke i gennemsnit på 0-5-skalaen.",
  "'Følger spillereglerne' beregnes kun for medlemmer, der har svaret (tomme svar tælles ikke som 0).",
  "0 på 1-5-skalaen behandles som manglende svar (beslutning 2026-10-31); ikke-svar tælles ikke med i gennemsnit.",
  "Kumuleret aldersfordeling divideres med antal medlemmer og ender derfor på 100 %.",
]

def num(v):
    if v is None or v == "": return None
    if isinstance(v, (int, float)): return float(v)
    m = re.match(r"^\s*(\d)", str(v)); return float(m.group(1)) if m else None

def answers(col, rows=ROWS):
    """Scale answers on 1-5. A 0 is not a scale label, so it is treated as missing (decision 2026-10-31)."""
    if str(col) in S["conditional_items"]:
        gate = S["conditional_items"][str(col)]
        rows = [r for r in rows if r[gate] == 1]
    return [x for x in (num(r[col]) for r in rows) if x is not None and x != 0]

def summary(vals):
    if not vals: return {"n": 0, "mean": None, "sd": None}
    return {"n": len(vals), "mean": st.mean(vals), "sd": st.stdev(vals) if len(vals) > 1 else 0.0}

def pooled(cols, rows=ROWS):
    vals = [v for c in cols for v in answers(c, rows)]
    return summary(vals)

def band(m):   # mean -> Danish interpretation (from 'Baggrund for beregninger')
    if m is None: return None
    for lo, txt in [(4.0, "Høj"), (3.1, "Moderat til høj"), (2.5, "Moderat"), (2.0, "Lav til moderat")]:
        if m >= lo - 1e-9: return txt
    return "Lav"

def spread_band(sd):
    if sd is None: return None
    for hi, txt in [(0.5, "Meget lille"), (1.0, "Lille"), (1.5, "Middel"), (2.0, "Stor")]:
        if sd <= hi + 1e-9: return txt
    return "Meget stor"

def item(col):
    s = summary(answers(col)); s.update(band=band(s["mean"]), spread=spread_band(s["sd"]), question=HDR[col]); return s

D = S["demographics"]
ages = [r[D["age"]] for r in ROWS]; ten = [r[D["tenure"]] for r in ROWS]
bins = [(0, 30)] + [(a, a + 5) for a in range(30, 90, 5)]
age_counts = [sum(1 for a in ages if lo <= a < hi) for lo, hi in bins]
cum, run = [], 0
for c in age_counts: run += c; cum.append(run / N)
R = {"n": N, "names": NAMES, "method_notes": METHOD_NOTES}
if "Submit Date (UTC)" in HDR:   # response period, used on the introduction slide
    _d = [r[HDR.index("Submit Date (UTC)")] for r in ROWS if r[HDR.index("Submit Date (UTC)")]]
    R["period"] = {"start": min(_d).isoformat(), "end": max(_d).isoformat()}
R["demographics"] = {
  "age_mean": st.mean(ages), "age_min": min(ages), "age_max": max(ages),
  "age_bins": [f"[{lo}:{hi}[" for lo, hi in bins], "age_counts": age_counts, "age_cum": cum,
  "tenure_mean": st.mean(ten), "tenure_mode": st.mode(ten), "tenure_mode_count": ten.count(st.mode(ten)),
  "gender": {g: sum(1 for r in ROWS if r[D["gender"]] == g) for g in ["Mand", "Kvinde"]},
  "board_training": sum(1 for r in ROWS if r[D["board_training"]] == 1),
  "want_other_size": sum(1 for r in ROWS if r[D["right_size"]] == 0),
}
R["diversity"] = {k: item(c) for k, c in S["diversity_items"].items()}
R["written_rules"] = {"yes": sum(1 for r in ROWS if r[16] == 1), "n": N}
R["general_work"] = [dict(label=g["label"], **pooled(g["cols"])) for g in S["general_work"]]
# item level (used by the narrative step to point at concrete questions)
R["items"] = []
for sec in ("general_work", "thematic_work"):
    for g in S[sec]:
        for c in g["cols"]:
            it = item(c); it.update(section=sec, subarea=g["label"], col=c); R["items"].append(it)
R["thematic_work"] = [dict(label=g["label"], **pooled(g["cols"])) for g in S["thematic_work"]]

# Competency areas: subarea pooled stats, area pooled mean, radar max/min over members' area means
areas = []
for a in S["competency_areas"]:
    cols = [c for s in a["subareas"] for c in s["cols"]]
    member_means = []
    for r in ROWS:   # a member's area mean uses only their valid answers; members with none are left out
        vals = [v for c in cols for v in answers(c, [r])]
        if vals: member_means.append(st.mean(vals))
    subs = [dict(label=s["label"], **pooled(s["cols"])) for s in a["subareas"]]
    area = dict(area=a["area"], subareas=subs, **pooled(cols),
                max=max(member_means), min=min(member_means),
                mean_sd=st.mean(s["sd"] for s in subs))
    area["band"] = band(area["mean"]); area["spread_band"] = spread_band(area["mean_sd"])
    areas.append(area)
R["competency_areas"] = areas
R["area_ratings"] = sorted([dict(area=k, **summary(answers(c))) for k, c in S["area_ratings"].items()], key=lambda x: -x["mean"])
R["team_competence"] = item(S["team_competence_col"])
R["priority_alignment"] = item(S["priority_alignment_col"])

# Ranking: comma-separated list, position 1 = most important
pos = {}
for r in ROWS:
    for p, t in enumerate(str(r[S["ranking_col"]]).split(","), 1): pos.setdefault(t.strip(), []).append(p)
R["ranking"] = sorted([{"topic": t, "avg_rank": st.mean(v), "n": len(v)} for t, v in pos.items()], key=lambda x: x["avg_rank"])

R["independent_members"] = sum(1 for r in ROWS if all(r[c] == 0 for c in S["independence_cols"]))
R["independence"] = [{"question": HDR[c], "yes": sum(1 for r in ROWS if r[c] == 1)} for c in S["independence_cols"]]

# Peer ratings: locate each block from the header row, not from fixed column numbers
def match_member(header):
    """Return the respondent index a peer-rating header refers to, or None."""
    if header is None: return None
    h = str(header).strip()
    cands = [i for i, nm in enumerate(NAMES) if nm.split()[-1] == h.split()[-1]]
    if len(cands) > 1: cands = [i for i in cands if NAMES[i][:4] == h[:4]]
    return cands[0] if len(cands) == 1 else None

def find_peer_blocks():
    """A block is the run of consecutive member-name headers. Searching starts after the last
    fixed question column, so question text can never be mistaken for a name."""
    blocks, pos = {}, max(S["independence_cols"]) + 1
    for dim in S["peer_rating_blocks"]:              # dimension order as in structure.json
        while pos < len(HDR) and match_member(HDR[pos]) is None: pos += 1
        start = pos
        while pos < len(HDR) and match_member(HDR[pos]) is not None: pos += 1
        blocks[dim] = list(range(start, pos))
    return blocks

peer = {}
for dim, cols in find_peer_blocks().items():
    targets = [match_member(HDR[c]) for c in cols]
    matched = sorted(t for t in targets if t is not None)
    if len(cols) != N or matched != list(range(N)):   # fail loudly, never silently
        raise ValueError(f"Peer block '{dim}' (columns {cols[0]}-{cols[-1]}): {len(cols)} headers, "
                         f"{len(matched)} matched members. Expected each of the {N} respondents exactly once.")
    for col, tgt_idx in zip(cols, targets):
        others = [num(r[col]) for i, r in enumerate(ROWS) if i != tgt_idx and num(r[col]) is not None]
        own = num(ROWS[tgt_idx][col])
        peer.setdefault(NAMES[tgt_idx], {})[dim] = {"peer_mean": st.mean(others), "n": len(others), "own": own}
R["peer"] = [{"name": n, **peer[n]} for n in NAMES]

# Individual competency profiles (member's own answers): per area 'Samlet' + subareas
prof = []
for r in ROWS:
    cats, vals = [], []
    for a in S["competency_areas"]:
        cols = [c for s in a["subareas"] for c in s["cols"]]
        _v = [x for x in (num(r[c]) for c in cols) if x not in (None, 0)]
        cats.append("Samlet"); vals.append(st.mean(_v) if _v else 0)  # 0 = no valid answers (same as subareas; flagged in profile check)
        for s in a["subareas"]:
            v = [x for x in (num(r[c]) for c in s["cols"]) if x not in (None, 0)]
            cats.append(s["label"]); vals.append(st.mean(v) if v else 0)
    prof.append({"name": r[S["respondent_name_col"]], "categories": cats, "values": vals})
R["profiles"] = prof

json.dump(R, open("results.json", "w"), ensure_ascii=False, indent=1, default=float)
f = lambda x: f"{x:.2f}"
print("n", N, "| tenure", f(R["demographics"]["tenure_mean"]), "| age", f(R["demographics"]["age_mean"]))
print("general work:", [(g["label"][:12], f(g["mean"]), f(g["sd"])) for g in R["general_work"]])
print("thematic:", [(g["label"][:12], f(g["mean"]), f(g["sd"])) for g in R["thematic_work"]])
for a in areas: print(f"{a['area'][:28]:<28} mean {f(a['mean'])} max {f(a['max'])} min {f(a['min'])} meanSD {f(a['mean_sd'])} |", [(s['label'][:14], f(s['mean']), f(s['sd'])) for s in a['subareas']])
print("ranking:", [(x["topic"][:14], f(x["avg_rank"])) for x in R["ranking"]])
print("peer[0]:", R["peer"][0])
print("independence:", [x["yes"] for x in R["independence"]])
