"""Build the importable set from src/.

src/  holds the readable UTF-8 sources (what you review on GitHub).
dist/ holds what Excel actually imports: the same text re-encoded to cp1251 (the VBA editor is
      ANSI on a Russian Windows), CRLF line ends, plus the binary CleaningForm.frx unchanged.
Also writes dist/excel-key-cleaner.zip with the four files a recipient needs.

    python build.py
"""
import pathlib, shutil, zipfile

ROOT = pathlib.Path(__file__).parent
SRC, DIST = ROOT / 'src', ROOT / 'dist'
DIST.mkdir(exist_ok=True)

for name in ('Cleaning.bas', 'CleaningForm.frm'):
    text = (SRC / name).read_text(encoding='utf-8').replace('\r\n', '\n').replace('\n', '\r\n')
    (DIST / name).write_bytes(text.encode('cp1251'))
    print(f'{name}: {len(text.splitlines())} lines -> dist/ (cp1251)')
shutil.copy(SRC / 'CleaningForm.frx', DIST / 'CleaningForm.frx')

with zipfile.ZipFile(DIST / 'excel-key-cleaner.zip', 'w', zipfile.ZIP_DEFLATED) as z:
    for name in ('Cleaning.bas', 'CleaningForm.frm', 'CleaningForm.frx'):
        z.write(DIST / name, name)
    z.write(ROOT / 'docs' / 'INSTALL.md', 'INSTALL.md')
    z.write(ROOT / 'docs' / 'INSTALL-FOR-CLAUDE.md', 'INSTALL-FOR-CLAUDE.md')
print('dist/excel-key-cleaner.zip written')
