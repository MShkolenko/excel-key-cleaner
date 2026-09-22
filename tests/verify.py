import sys, csv, pathlib, re
sys.stdout.reconfigure(encoding='utf-8')
HERE = pathlib.Path(sys.argv[1])   # the gen.py out dir, after Run-CleaningAudit.ps1 wrote results.csv there
# the reports file is normally reports.tsv next to the rest; a second argument names it when the
# evidence was copied elsewhere under another name (Codex round 6 ran into exactly that)
REPORTS = HERE / (sys.argv[2] if len(sys.argv) > 2 else 'reports.tsv')
def load(name, key):
    with open(HERE / name, encoding='utf-8-sig', newline='') as f:
        return {tuple(r[k] for k in key): r for r in csv.DictReader(f)}
cases = load('cases.csv', ('id',))
exp = load('expected.csv', ('opt', 'id'))
act = load('results.csv', ('opt', 'id'))
ok = fail = 0
print(f"{'opt':4} {'id':4} {'input':40} {'actual':40} {'expected(mirror)':40} verdict")
for (opt, cid), e in exp.items():
    a = act[(opt, cid)]['actual']
    inp = cases[(cid,)]['input']
    same = a == e['expected']
    ok += same; fail += (not same)
    if not same or opt == 'D':
        print(f"{opt:4} {cid:4} {inp!r:40} {a!r:40} {e['expected']!r:40} {'OK' if same else 'MISMATCH'}")
print(f"\nmirror agreement: {ok} OK, {fail} MISMATCH out of {ok+fail}")

# ---- invariants over the whole result set ----
HOMO = set('АВЕКМНОРСТХаеорсх')
BAD_WS = set(chr(c) for c in [9, 13, 10, 11, 12, 0x85, 0xA0, 0x1680, 0x2028, 0x2029, 0x202F, 0x205F, 0x3000, 0x200B, 0x200C, 0x200D, 0x2060, 0xFEFF]) | set(chr(c) for c in range(0x2000, 0x200B))
ESC = {'{LF}': '\n', '{CR}': '\r', '{TAB}': '\t', '{NBSP}': '\u00a0', '{ZWSP}': '\u200b', '{ZWNJ}': '\u200c', '{ZWJ}': '\u200d', '{BOM}': '\ufeff',
       '{ENSP}': '\u2002', '{THIN}': '\u2009', '{IDEO}': '\u3000', '{WJ}': '\u2060', '{LSEP}': '\u2028', '{PSEP}': '\u2029', '{OGH}': '\u1680', '{NEL}': '\u0085', '{VT}': '\u000b', '{FF}': '\u000c'}
def unesc(s):
    for k, v in ESC.items(): s = s.replace(k, v)
    return s
inv_fail = 0
# I1: with the override set (O = 1,1,0,1,1) no selected text cell keeps a homoglyph, a bad whitespace char, or an edge space
for (opt, cid), r in act.items():
    if opt != 'O': continue
    v = unesc(r['actual'])
    if set(v) & HOMO: print(f"INVARIANT I1 FAIL {cid}: homoglyph left in {v!r}"); inv_fail += 1
    if set(v) & BAD_WS: print(f"INVARIANT I1 FAIL {cid}: bad whitespace left in {v!r}"); inv_fail += 1
    if v != v.strip(' '): print(f"INVARIANT I1 FAIL {cid}: edge space left in {v!r}"); inv_fail += 1
# I2: with the default set (D) every case whose result still holds Cyrillic is NAMED in the report and SELECTED
reports = {}
with open(REPORTS, encoding='utf-8-sig') as f:
    for line in f:
        if '\t' in line:
            k, v = line.rstrip('\n').split('\t', 1); reports[k] = v
ids = [r['id'] for r in csv.DictReader(open(HERE / 'cases.csv', encoding='utf-8-sig', newline=''))]
row_of = {cid: i + 2 for i, cid in enumerate(ids)}        # cases sit in B2.. in case order
def expand(addr):
    # "B4:B6,B9" -> {"B4","B5","B6","B9"} (single column B, as the matrix is laid out)
    out = set()
    for tok in addr.split(','):
        tok = tok.strip().replace('$', '')
        if ':' in tok:
            a, b = tok.split(':')
            ra, rb = int(re.sub(r'[A-Z]', '', a)), int(re.sub(r'[A-Z]', '', b))
            out |= {f"B{i}" for i in range(ra, rb + 1)}
        elif tok:
            out.add(tok)
    return out
MAX_LISTED = 12
for opt in ('D', 'O'):
    msg = reports.get(opt, '')
    m = re.search(r'[|]SELECTED=([^|]+)$', msg)
    selected = expand(m.group(1)) if m else set()
    left = [(cid, e) for (o, cid), e in exp.items() if o == opt and e['cyr_left'] == '1']
    for k, (cid, e) in enumerate(left):
        addr = f"B{row_of[cid]}"
        if k < MAX_LISTED and not re.search(r'(?<![A-Z])' + addr + r'(?![0-9])', msg):
            print(f"INVARIANT I2 FAIL [{opt}] {cid}: {addr} not named in the report"); inv_fail += 1
        if addr not in selected:
            print(f"INVARIANT I2 FAIL [{opt}] {cid}: {addr} not in the final selection ({sorted(selected)})"); inv_fail += 1
    if not left and opt == 'D':
        print("INVARIANT I2: no Cyrillic-left cases in D - check the case table"); inv_fail += 1
print(f"invariants: {'all OK' if inv_fail == 0 else str(inv_fail) + ' FAIL'}")
# a verifier that prints failures and exits 0 is not a gate (Codex round 6)
sys.exit(1 if (fail or inv_fail) else 0)
