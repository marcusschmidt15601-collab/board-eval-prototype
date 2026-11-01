"""narrative.py — facts.json + context.json + template text -> narrative.json (drafts for review)

  python narrative.py            draft every section with the Claude API (needs ANTHROPIC_API_KEY)
  python narrative.py --prompts  only write the prompts to prompts/ (no API call)
  python narrative.py --check    run the guardrails on an existing narrative.json (e.g. after manual edits)

Each draft is checked before it is accepted: format and length, every number must exist in the facts
given to that section, member names only where allowed, no claims about earlier evaluations.
A section that fails is retried with the error; after 3 attempts it is kept but marked needs_review."""
import json, os, re, sys
from pptx import Presentation

MODEL = os.environ.get("NARRATIVE_MODEL", "claude-sonnet-5-5")
TEMPLATE = os.environ.get("REPORT_TEMPLATE", "input/template.pptx")
SPEC = json.load(open("narrative_spec.json"))
FACTS = json.load(open("facts.json"))
CTX = json.load(open("context.json"))
SYSTEM = open("narrative_prompt.md").read()
NAMES = [m["navn"] for m in FACTS["bidrag"]["medlemmer"]]

def template_text(slide, shape_id):
    prs = Presentation(TEMPLATE)
    def walk(shs):
        for s in shs:
            if s.shape_type == 6: yield from walk(s.shapes)
            else: yield s
    sh = next(s for s in walk(prs.slides[slide - 1].shapes) if s.shape_id == shape_id)
    return "\n".join(p.text for p in sh.text_frame.paragraphs if p.text.strip())

FORMATS = {
  "lead_bullets": '{"lead": "...", "bullets": ["...", "..."]}',
  "paragraphs": '{"paragraphs": ["...", "..."]}',
  "headed_items": '{"items": [{"heading": "...", "text": "..."}]}',
  "lead_intro_bullets": '{"lead": "...", "intro": "...", "bullets": ["..."]}',
}
def build_prompt(sec):
    facts = {k: FACTS[k] for k in sec["facts"]}
    ctx = {k: v for k, v in CTX.items() if not k.startswith("_") and v}
    rng = sec.get("bullets") or sec.get("items")
    count = f"Antal punkter: {rng[0]}-{rng[1]}. " if rng else f"Antal afsnit: {len(sec['para_idx'])}. "
    return (f"<opgave>{sec['task']} {count}Maks {sec['max_words']} ord pr. punkt/afsnit.</opgave>\n"
            f"<format>{FORMATS[sec['layout']]}</format>\n"
            f"<eksempel>\n{template_text(sec['slide'], sec['shape'])}\n</eksempel>\n"
            f"<kontekst>{json.dumps(ctx, ensure_ascii=False)}</kontekst>\n"
            f"<facts>{json.dumps(facts, ensure_ascii=False)}</facts>")

# ------------------------------------------------------------------ guardrails
def budget(sec):
    """Character budget = 125% of the template text this section replaces, so the draft fits the text box."""
    prs = Presentation(TEMPLATE)
    def walk(shs):
        for s in shs:
            if s.shape_type == 6: yield from walk(s.shapes)
            else: yield s
    ps = next(s for s in walk(prs.slides[sec["slide"] - 1].shapes) if s.shape_id == sec["shape"]).text_frame.paragraphs
    if sec["layout"] == "paragraphs": idx = sec["para_idx"]
    elif sec["layout"] == "headed_items": idx = [i for i, p in enumerate(ps) if p.text.strip()]
    else:
        idx = [sec["lead_idx"]] + ([sec["intro_idx"]] if "intro_idx" in sec else [])
        i = sec["bullet_idx"]
        while i < len(ps) and ps[i].text.strip():
            if not sec.get("sub_from_context") or ps[i].level == 0: idx.append(i)
            i += 1
    return int(1.25 * sum(len(ps[i].text.strip()) for i in idx))
NUM = re.compile(r"\d+(?:[.,]\d+)?")
def allowed_numbers(sec):
    blob = json.dumps({k: FACTS[k] for k in sec["facts"]}, ensure_ascii=False) + json.dumps(CTX, ensure_ascii=False)
    return set(NUM.findall(blob))
def texts_of(draft):
    out = [draft.get("lead", ""), draft.get("intro", "")] + draft.get("bullets", []) + draft.get("paragraphs", [])
    for it in draft.get("items", []): out += [it.get("heading", ""), it.get("text", "")]
    return [t for t in out if t]
def check(sec, draft):
    errs = []
    lay = sec["layout"]
    need = {"lead_bullets": ["lead", "bullets"], "paragraphs": ["paragraphs"], "headed_items": ["items"],
            "lead_intro_bullets": ["lead", "intro", "bullets"]}[lay]
    for k in need:
        if k not in draft: errs.append(f"mangler feltet '{k}'")
    if errs: return errs
    lo_hi = sec.get("bullets") or sec.get("items")
    n = len(draft.get("bullets", draft.get("items", draft.get("paragraphs", []))))
    if lo_hi and not lo_hi[0] <= n <= lo_hi[1]: errs.append(f"{n} punkter, skal være {lo_hi[0]}-{lo_hi[1]}")
    if lay == "paragraphs" and n != len(sec["para_idx"]): errs.append(f"{n} afsnit, skal være {len(sec['para_idx'])}")
    total = sum(len(t) for t in texts_of(draft)); cap = budget(sec)
    if total > cap: errs.append(f"teksten fylder {total} tegn, men pladsen på sliden rækker til ca. {cap}; forkort")
    ok_nums = allowed_numbers(sec)
    for t in texts_of(draft):
        if len(t.split()) > sec["max_words"] + 10: errs.append(f"for lang ({len(t.split())} ord): {t[:60]}...")
        for x in NUM.findall(t):
            if x not in ok_nums: errs.append(f"tallet {x} findes ikke i fakta: {t[:70]}...")
        low = t.lower()
        if not CTX.get("forrige_evaluering") and re.search(r"forrige|tidligere evaluering|sidste år|fremgang|tilbagegang", low):
            errs.append(f"omtaler en tidligere evaluering, som ikke findes i data: {t[:70]}...")
        if not sec.get("names_allowed"):
            for nm in NAMES:
                if re.search(rf"\b{re.escape(nm.split()[0])}\b", t): errs.append(f"nævner {nm} i et afsnit, hvor medlemmer ikke må nævnes")
        else:
            for nm in NAMES:
                first = nm.split()[0]
                if re.search(rf"\b{first}\b", t) and nm not in t and not any(o != nm and o in t for o in NAMES if o.split()[0] == first):
                    errs.append(f"brug det fulde navn '{nm}'")
    return errs

# ------------------------------------------------------------------ API
def call_claude(system, prompt):
    import anthropic
    client = anthropic.Anthropic()
    msg = client.messages.create(model=MODEL, max_tokens=1500, system=system,
                                 messages=[{"role": "user", "content": prompt}])
    text = "".join(b.text for b in msg.content if b.type == "text").strip()
    text = re.sub(r"^```(?:json)?|```$", "", text).strip()
    return json.loads(text)

def draft_all():
    out = {}
    for sec in SPEC:
        prompt, feedback = build_prompt(sec), ""
        for attempt in range(3):
            draft = call_claude(SYSTEM, prompt + feedback)
            errs = check(sec, draft)
            if not errs: break
            feedback = "\n<rettelser>Dit forrige udkast blev afvist:\n- " + "\n- ".join(errs) + "\nRet fejlene.</rettelser>"
        out[sec["id"]] = {"status": "draft" if not errs else "needs_review", "errors": errs, **draft}
        print(sec["id"], out[sec["id"]]["status"], errs)
    json.dump(out, open("narrative.json", "w"), ensure_ascii=False, indent=1)

if __name__ == "__main__":
    if "--prompts" in sys.argv:
        os.makedirs("prompts", exist_ok=True)
        for sec in SPEC: open(f"prompts/{sec['id']}.txt", "w").write(SYSTEM + "\n\n" + build_prompt(sec))
        print("prompts written:", [s["id"] for s in SPEC])
    elif "--check" in sys.argv:
        N = json.load(open("narrative.json")); bad = 0
        for sec in SPEC:
            errs = check(sec, N[sec["id"]]); bad += bool(errs)
            print(f"{sec['id']:<4} {'OK' if not errs else 'AFVIST'}", *errs, sep="\n      ")
        sys.exit(1 if bad else 0)
    else:
        draft_all()
