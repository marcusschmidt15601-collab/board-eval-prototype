"""generate_synthetic.py - invented board-evaluation export for testing. Contains NO real data.
Board of 10. Respondent index 6 answers 'no' to col 16 and correctly skips col 17.
Peer blocks follow the real layout for 10 members: 10 columns per block, one spacer column.
Run from the project folder:  python3 generate_synthetic.py
"""
import datetime, json, os, random
import openpyxl

SEED, N, SKIP_IDX = 20261101, 10, 6
random.seed(SEED)
S = json.load(open("structure.json"))
NAMES = ["Kim Eksempel", "Sam Fiktiv", "Lee Testsen", "Jo Demo", "Max Prøvemann",
         "Eva Syntetisk", "Ole Dummy", "Ida Mock", "Rune Kunstig", "Mia Modelsen"]
TOPICS = ["Forretningsudvikling", "Ledelsesudvikling", "Økonomistyring", "Interessenter og stakeholder",
          "Organisatoriske forhold", "Risikostyring", "Digitalisering og teknologi",
          "Internationalisering", "Bæredygtighed og ESG", "Juridiske forhold"]
LABEL = {1: "Slet ikke", 2: "I nogen grad", 3: "I moderat grad", 4: "I høj grad", 5: "I meget høj grad"}
PEER_START = 135
BLOCK_STEP = N + 1                       # 10 columns + 1 spacer
TOTAL = PEER_START + 3 * BLOCK_STEP + N  # last peer column + 1

def scale(): return random.randint(1, 5)

headers = [f"Kolonne {i}" for i in range(TOTAL)]
headers[0], headers[1] = "Submit Date (UTC)", "Navn"
for k, q in S["questions"].items(): headers[int(k)] = q
for c in range(128, 135): headers[c] = f"Uafhængighed kriterium {c}"
for k, start in enumerate(S["peer_rating_blocks"].values()):
    for j in range(N): headers[PEER_START + k * BLOCK_STEP + j] = NAMES[j]

rows = []
for i, name in enumerate(NAMES):
    r = [None] * TOTAL
    r[0] = datetime.datetime(2026, 4, 15 + i, 10, 0)
    r[1] = name
    r[2] = random.randint(44, 69)
    r[3] = random.choice(["Mand", "Kvinde"])
    r[4] = random.randint(1, 6)
    r[5] = random.randint(0, 1)
    for c in range(6, 11): r[c] = scale()
    r[11] = 0 if i == 0 else 1           # one member would like a different board size
    r[12] = random.choice([8, 9, 10])
    for c in range(13, 58):
        if c not in (16, 17): r[c] = scale()
    r[16] = 0 if i == SKIP_IDX else 1
    r[17] = None if i == SKIP_IDX else scale()
    r[58] = ", ".join(random.sample(TOPICS, len(TOPICS)))
    r[59] = scale()
    for c in range(60, 116): r[c] = scale()
    for c in range(116, 127):
        v = scale(); r[c] = f"{v} - {LABEL[v]}"
    r[127] = scale()
    for c in range(128, 135): r[c] = 1 if random.random() < 0.2 else 0
    for k in range(4):
        for j in range(N): r[PEER_START + k * BLOCK_STEP + j] = scale()
    rows.append(r)

wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Svar"
ws.append(headers)
for r in rows: ws.append(r)
os.makedirs("synthetic", exist_ok=True)
wb.save("synthetic/raw.xlsx")
print(f"wrote synthetic/raw.xlsx: {N} respondents, {TOTAL} columns, skipped conditional for {NAMES[SKIP_IDX]}")
