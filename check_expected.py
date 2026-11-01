"""check_expected.py - recompute key figures straight from the synthetic export and compare with results.json.
Run after calc.py, in the same folder as results.json. Expected: general-work check PASS, peer check FAIL (see notes)."""
import json, statistics as st, openpyxl

S = json.load(open("structure.json")); R = json.load(open("results.json"))
ws = openpyxl.load_workbook("synthetic/raw.xlsx", data_only=True).worksheets[0]
rows = [list(r) for r in ws.iter_rows(min_row=2, values_only=True) if r[1]]
N = len(rows); PEER_START, STEP = 135, N + 1

# 1. Conditional question: Beslutningsprocesser (cols 13, 14, 15, 17); col 17 only where col 16 == 1
def valid(c, row):
    if c == 17 and row[16] != 1: return False
    return isinstance(row[c], (int, float))
vals = [row[c] for row in rows for c in (13, 14, 15, 17) if valid(c, row)]
exp_mean, exp_sd = st.mean(vals), st.stdev(vals)
got = R["general_work"][0]
ok1 = abs(exp_mean - got["mean"]) < 1e-9 and abs(exp_sd - got["sd"]) < 1e-9 and len(vals) == got["n"]
print(f"[1] Beslutningsprocesser  expected mean {exp_mean:.4f} sd {exp_sd:.4f} n {len(vals)}"
      f" | pipeline mean {got['mean']:.4f} sd {got['sd']:.4f} n {got['n']} -> {'PASS' if ok1 else 'FAIL'}")

# 2. Independent members (all seven independence answers = 0)
exp_ind = sum(1 for row in rows if all(row[c] == 0 for c in range(128, 135)))
print(f"[2] Independent members   expected {exp_ind} | pipeline {R['independent_members']}"
      f" -> {'PASS' if exp_ind == R['independent_members'] else 'FAIL'}")

# 3. Peer ratings: mean of the other members' ratings, read from the ACTUAL column positions
dims = list(S["peer_rating_blocks"].keys())
by_name = {p["name"]: p for p in R["peer"]}
bad = 0
for j, row_j in enumerate(rows):
    for k, dim in enumerate(dims):
        col = PEER_START + k * STEP + j
        others = [r[col] for i, r in enumerate(rows) if i != j and isinstance(r[col], (int, float))]
        exp = st.mean(others)
        got_v = by_name.get(row_j[1], {}).get(dim, {}).get("peer_mean")
        if got_v is None or abs(exp - got_v) > 1e-9: bad += 1
print(f"[3] Peer means            {bad} of {N * len(dims)} member-dimension values differ -> "
      f"{'PASS' if bad == 0 else 'FAIL'}")
print(f"    members found in results.json: {len(R['peer'])} of {N}")
