import sys, csv, pathlib
sys.stdout.reconfigure(encoding='utf-8')
HERE = pathlib.Path(sys.argv[1])   # the gen.py out dir, after Run-CleaningAudit.ps1 wrote results.csv there
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
