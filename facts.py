"""facts.py — results.json (+ context.json) -> facts.json
All comparisons and rankings the narrative may refer to are computed HERE, in code.
Numbers are pre-formatted in Danish so the text can copy them verbatim, and the guardrail in
narrative.py rejects any number that is not in this file."""
import json
from decimal import Decimal, ROUND_HALF_UP
R = json.load(open("results.json"))
CTX = json.load(open("context.json")) if __import__("os").path.exists("context.json") else {}

def f(x, d=2): return f"{Decimal(str(x)).quantize(Decimal(1).scaleb(-d), rounding=ROUND_HALF_UP):.{d}f}".replace(".", ",")
def pct(x): return f"{Decimal(str(x*100)).quantize(Decimal(1), rounding=ROUND_HALF_UP)}%"
def spread(sd):
    for hi, t in [(0.5, "meget lille"), (1.0, "lille"), (1.5, "middel"), (2.0, "stor")]:
        if sd <= hi + 1e-9: return t
    return "meget stor"

# one key per topic, so priority / competence / work can be compared
STEMS = [("disciplin", "Bestyrelsesdisciplin"), ("diciplin", "Bestyrelsesdisciplin"), ("økonomi", "Økonomistyring"),
         ("ledelse", "Ledelsesudvikling"), ("forretning", "Forretningsudvikling"), ("interessent", "Interessenter og stakeholdere"),
         ("organisator", "Organisatoriske forhold"), ("risi", "Risikostyring"), ("risk", "Risikostyring"), ("digital", "Digitalisering og teknologi"),
         ("international", "Internationalisering"), ("bæredygt", "Bæredygtighed og ESG"), ("juridisk", "Juridiske forhold"),
         ("direktion", "Direktionen")]
def topic(name):
    n = name.lower()
    for stem, t in STEMS:
        if stem in n: return t
    raise KeyError(name)

topics = {}
def put(name, **kw): topics.setdefault(topic(name), {"emne": topic(name)}).update(kw)
for i, x in enumerate(R["ranking"], 1): put(x["topic"], prioritet_placering=i, prioritet_gns_rang=f(x["avg_rank"]))
comp = sorted(R["competency_areas"], key=lambda a: -a["mean"])
for i, a in enumerate(comp, 1): put(a["area"], kompetence_placering=i, kompetence_gns=f(a["mean"]))
for i, a in enumerate(R["area_ratings"], 1): put(a["area"], vurderet_tilstede_placering=i, vurderet_tilstede_gns=f(a["mean"]))
tw = sorted(R["thematic_work"], key=lambda x: -x["mean"])
for i, x in enumerate(tw, 1): put(x["label"], arbejde_placering=i, arbejde_gns=f(x["mean"]), arbejde_spredning=spread(x["sd"]))
n_comp, n_rank, n_work = len(comp), len(R["ranking"]), len(tw)

d = R["demographics"]; n = R["n"]; div = R["diversity"]
items = R["items"]
def low_items(section, k):
    return [{"spørgsmål": it["question"].strip(), "område": it["subarea"], "gns": f(it["mean"]), "spredning": spread(it["sd"])}
            for it in sorted([x for x in items if x["section"] == section], key=lambda x: x["mean"])[:k]]
def high_items(section, k):
    return [{"spørgsmål": it["question"].strip(), "område": it["subarea"], "gns": f(it["mean"])}
            for it in sorted([x for x in items if x["section"] == section], key=lambda x: -x["mean"])[:k]]

DIMS = ["Aktivt bidrag", "Viden og kompetencer", "Deler tanker, indsigter og perspektiver", "Samarbejdsvillig, konstruktiv og reflekteret"]
dim_avg = {dm: sum(m[dm]["peer_mean"] for m in R["peer"]) / n for dm in DIMS}
members = []
for m in R["peer"]:
    dims = []
    for dm in DIMS:
        x = m[dm]; dims.append({"dimension": dm, "andres_vurdering": f(x["peer_mean"]), "egen_vurdering": f(x["own"], 0),
                                "bestyrelsens_gns": f(dim_avg[dm]),
                                "under_bestyrelsens_gns_med_mindst_0,3": x["peer_mean"] <= dim_avg[dm] - 0.3,
                                "egen_vurdering_mindst_0,5_højere_end_andres": x["own"] - x["peer_mean"] >= 0.5})
    members.append({"navn": m["name"], "samlet_andres_vurdering": f(sum(m[dm]["peer_mean"] for dm in DIMS) / 4), "dimensioner": dims})
board_avg_all = sum(dim_avg.values()) / 4

F = {
 "generelt": {"antal_medlemmer": n, "skala": "0-5, hvor 5 er i meget høj grad",
              "fortolkning_gns": "1,0-1,9 lav; 2,0-2,4 lav til moderat; 2,5-3,0 moderat; 3,1-3,9 moderat til høj; 4,0-5,0 høj",
              "fortolkning_spredning": "0-0,5 meget lille; 0,5-1,0 lille; 1,0-1,5 middel; 1,5-2,0 stor; over 2,0 meget stor"},
 "sammensaetning": {
   "kvinder": d["gender"]["Kvinde"], "mænd": d["gender"]["Mand"],
   "andel_kvinder": pct(d["gender"]["Kvinde"] / n), "andel_mænd": pct(d["gender"]["Mand"] / n),
   "alder_gns": f(d["age_mean"], 0), "alder_yngste": d["age_min"], "alder_ældste": d["age_max"],
   "anciennitet_gns": f(d["tenure_mean"], 1), "medlemmer_med_bestyrelsesuddannelse": d["board_training"],
   "ønsker_anden_størrelse": d["want_other_size"],
   "mangfoldighed": {k: {"gns": f(v["mean"], 1), "spredning": spread(v["sd"]), "spørgsmål": v["question"].strip()} for k, v in div.items()}},
 "kompetencer": {
   "områder_efter_niveau": [{"område": topic(a["area"]), "gns": f(a["mean"]), "niveau": a["band"].lower(), "spredning": spread(a["mean_sd"])} for a in comp],
   "samlet_gns": f(sum(a["mean"] for a in comp) / n_comp), "antal_områder": n_comp,
   "vurderet_tilstede_efter_niveau": [{"område": topic(a["area"]), "gns": f(a["mean"])} for a in R["area_ratings"]],
   "opmærksom_på_kompetencesammensætning_gns": f(R["team_competence"]["mean"], 1),
   "laveste_4_niveau": [topic(a["area"]) for a in comp[-4:]],
   "laveste_4_vurderet_tilstede": [topic(a["area"]) for a in R["area_ratings"][-4:]]},
 "prioritet_arbejde_kompetence": {
   "forklaring": f"placering 1 = højest. prioritet: 1-{n_rank} (lav gns. rang = høj prioritet); kompetence: 1-{n_comp}; arbejde: 1-{n_work}",
   "emner": sorted(topics.values(), key=lambda t: t.get("prioritet_placering", 99)),
   "prioritet_følger_nuværende_fokus_gns": f(R["priority_alignment"]["mean"], 1)},
 "generelt_arbejde": {
   "områder": [{"område": x["label"], "gns": f(x["mean"]), "spredning": spread(x["sd"])} for x in R["general_work"]],
   "laveste_spørgsmål": low_items("general_work", 5), "højeste_spørgsmål": high_items("general_work", 3),
   "har_skriftlig_forretningsorden_ja": R["written_rules"]["yes"], "har_skriftlig_forretningsorden_af": R["written_rules"]["n"]},
 "faglige_temaer": {
   "temaer_efter_niveau": [{"tema": topic(x["label"]), "gns": f(x["mean"]), "spredning": spread(x["sd"])} for x in tw],
   "laveste_spørgsmål": low_items("thematic_work", 5)},
 "uafhaengighed": {
   "uafhængige_medlemmer": R["independent_members"], "andel_uafhængige": pct(R["independent_members"] / n),
   "anbefaling": "mindst halvdelen (50%) bør være uafhængige (Komitéen for god Selskabsledelse)",
   "kriterier_antal_ja": [{"kriterium": x["question"].strip(), "antal_ja": x["yes"]} for x in R["independence"]],
   "baggrund_fra_konsulent": CTX.get("board_background", [])},
 "bidrag": {"bestyrelsens_samlede_gns": f(board_avg_all), "medlemmer": members},
 "kontekst_fra_konsulent": {k: v for k, v in CTX.items() if k != "board_background"},
}
json.dump(F, open("facts.json", "w"), ensure_ascii=False, indent=1)
print("facts.json:", {k: len(json.dumps(v, ensure_ascii=False)) for k, v in F.items()})
