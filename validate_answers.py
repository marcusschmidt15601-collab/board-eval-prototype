"""validate_answers.py - check raw answers against question_bank.json.
usage: python3 validate_answers.py raw.xlsx [question_bank.json]
Errors block the import; warnings are reported but allowed (used for legacy exports)."""
import json, re, sys, openpyxl
from collections import Counter

RAW = sys.argv[1]
QB = json.load(open(sys.argv[2] if len(sys.argv) > 2 else "question_bank.json"))
ws = openpyxl.load_workbook(RAW, data_only=True).active
hdr = [c.value for c in ws[1]]
rows = [[c.value for c in r] for r in ws.iter_rows(min_row=2) if r[1].value]   # completed responses only
names = [r[1] for r in rows]
errors, warnings = [], []
RANKING_TOPICS = {"Forretningsudvikling", "Ledelse", "Økonomistyring", "Interessenter og stakeholders",
                  "Organisatoriske forhold", "Riskostyring", "Digitalisering og teknologi",
                  "Internationalisering", "Bæredygtighed og ESG", "Juridiske forhold"}
GENDERS = {"Mand", "Kvinde", "Andet", "Ønsker ikke at oplyse"}
DOUBLE = re.compile(r"^\s*([1-5])\b")

def legacy_scale(v):
    """Legacy export stores '4 - I høj grad' text for matrices; plain numbers elsewhere."""
    if v is None: return None
    if isinstance(v, (int, float)): return int(v)
    m = DOUBLE.match(str(v)); return int(m.group(1)) if m else "bad"

for r_i, r in enumerate(rows, 1):
    who = r[1]
    gate_ok = {q["legacy_col"]: q for q in QB["questions"] if isinstance(q["legacy_col"], int)}
    for q in QB["questions"]:
        c = q["legacy_col"]
        if not isinstance(c, int) or q["type"] in ("member_select",): continue
        v = r[c]
        # conditional: must be blank unless gate answered 1
        if q.get("condition"):
            gate = next(x for x in QB["questions"] if x["id"] == q["condition"]["question"])
            gv = r[gate["legacy_col"]]
            if gv != 1 and v is not None:
                errors.append(f"{who}: '{q['id']}' answered although gate '{gate['id']}' is not Ja")
            if gv == 1 and v is None and q["required"]:
                errors.append(f"{who}: '{q['id']}' missing (gate is Ja)")
            if v is None: continue
        elif v is None:
            if q["required"] and q["type"] != "ranking": errors.append(f"{who}: '{q['id']}' (form {q['form_no']}) missing")
            continue
        t = q["type"]
        if t == "scale":
            s = legacy_scale(v)
            if s == "bad" or s is None: errors.append(f"{who}: '{q['id']}' value {v!r} is not a scale value")
            elif s == 0: warnings.append(f"{who}: '{q['id']}' is 0 (outside 1-5; treated as missing in calc)")
            elif s not in (1, 2, 3, 4, 5): errors.append(f"{who}: '{q['id']}' value {s} outside 1-5")
        elif t == "yes_no":
            if v not in (0, 1): errors.append(f"{who}: '{q['id']}' value {v!r} is not 0/1")
        elif t == "number":
            try: n = float(v)
            except: errors.append(f"{who}: '{q['id']}' value {v!r} is not a number"); continue
            lo, hi = q.get("validate", {}).get("min"), q.get("validate", {}).get("max")
            if (lo is not None and n < lo) or (hi is not None and n > hi):
                errors.append(f"{who}: '{q['id']}' value {n:g} outside {lo}-{hi}")
        elif t == "single_choice":
            if v not in GENDERS: errors.append(f"{who}: '{q['id']}' value {v!r} not an option")
        elif t == "ranking":
            items = [x.strip() for x in str(v).split(",")]
            if len(items) != 10 or len(set(items)) != 10:
                errors.append(f"{who}: ranking has {len(items)} items, {len(set(items))} unique (need 10)")
            unknown = [x for x in items if x not in RANKING_TOPICS]
            if unknown: errors.append(f"{who}: ranking topics not recognised: {unknown}")
        elif t == "matrix_scale":
            for row in q["rows"]:
                s = legacy_scale(r[row["legacy_col"]])
                if s in (None, "bad") or s not in (1, 2, 3, 4, 5):
                    errors.append(f"{who}: matrix '{row['label']}' value {r[row['legacy_col']]!r} invalid")

# member matrices: every member must be rated by the others (self-rating allowed, blocks must be complete)
blocks = {q["subarea"]: q["legacy_col"] for q in QB["questions"] if q["type"] == "member_matrix"}
for label, start in blocks.items():
    for j in range(8):
        col = start + j
        if not str(hdr[col] or "").strip():
            errors.append(f"peer block '{label}': column {col + 1} has no header"); continue
        if hdr[col] not in names and hdr[col].split()[-1] not in [n.split()[-1] for n in names]:
            errors.append(f"peer block '{label}': column '{hdr[col]}' matches no respondent")
    n_filled = [sum(1 for r in rows if r[start + j] is not None) for j in range(8)]
    if len(set(n_filled)) > 1: warnings.append(f"peer block '{label}': uneven ratings per member {n_filled}")

print(f"Responses checked: {len(rows)} | questions per response: "
      f"{sum(1 for q in QB['questions'] if isinstance(q['legacy_col'], int))}")
print(f"Errors: {len(errors)} | Warnings: {len(warnings)}")
for e in errors: print("  ERROR  ", e)
zeros = Counter(w.split("'")[1] for w in warnings if "is 0" in w)
if zeros: print("  WARN   0 on 1-5 items:", sum(zeros.values()), "answers across", len(zeros), "questions")
for w in warnings:
    if "is 0" not in w: print("  WARN  ", w)
sys.exit(1 if errors else 0)
