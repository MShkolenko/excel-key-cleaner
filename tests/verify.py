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
# I2 (v3.3, operator 2026-09-24): with D and O every case with Cyrillic left is COUNTED in the report; the
# SUSPICIOUS ones (a Latin and a Cyrillic letter in one word, or a letter with no Latin pair left in a
# converted word) are SELECTED, Russian text is NOT selected; the report text names no cell at all.
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
for opt in ('D', 'O'):
    msg = reports.get(opt, '')
    text, _, sel = msg.partition('|SELECTED=')
    selected = expand(sel.split('|')[0]) if sel else set()
    rows = [(cid, e) for (o, cid), e in exp.items() if o == opt]
    n_left = sum(e['cyr_left'] == '1' for _, e in rows)
    n_susp = sum(e['suspicious'] == '1' for _, e in rows)
    # exact, not just "contains" (Codex, v3.3 review): the final selection IS the suspicious set, the
    # report text names no cell of any column, and each count line appears once with the right number
    want_sel = {f"B{row_of[cid]}" for cid, e in rows if e['suspicious'] == '1'}
    if selected != want_sel:
        print(f"INVARIANT I2 FAIL [{opt}]: selection {sorted(selected)} != suspicious {sorted(want_sel)}"); inv_fail += 1
    named = re.findall(r'(?<![A-Za-z0-9])[A-Z]{1,3}[1-9][0-9]{0,6}(?![A-Za-z0-9])', text)
    if named:
        print(f"INVARIANT I2 FAIL [{opt}]: the report text names cells {named} (no addresses since v3.3)"); inv_fail += 1
    mixed_head = 'БУКВЫ БЕЗ ЛАТИНСКОЙ ПАРЫ (Ж, Ш, У, Ы...) остались в ' if opt == 'O' else 'ЛАТИНИЦА И КИРИЛЛИЦА В ОДНОМ СЛОВЕ: '
    for head, n in ((mixed_head, n_susp), ('Русский текст оставлен как есть: ', n_left - n_susp)):
        got = re.findall(re.escape(head) + r'(\d+) яч\.', text)
        if got != ([str(n)] if n else []):
            print(f"INVARIANT I2 FAIL [{opt}]: {head!r} lines {got}, expected {[n] if n else 'none'}"); inv_fail += 1
    if opt == 'D' and not (n_susp and n_left - n_susp):
        print("INVARIANT I2: D needs both suspicious and Russian-text cases - check the case table"); inv_fail += 1
print(f"invariants: {'all OK' if inv_fail == 0 else str(inv_fail) + ' FAIL'}")
# a verifier that prints failures and exits 0 is not a gate (Codex round 6)
sys.exit(1 if (fail or inv_fail) else 0)
