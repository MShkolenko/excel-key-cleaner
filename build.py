"""Build the importable set and the add-in from src/.

src/  holds the readable UTF-8 sources (what you review on GitHub).
dist/ holds what Excel actually imports: the same text re-encoded to cp1251 (the VBA editor is
      ANSI on a Russian Windows), CRLF line ends, plus the binary CleaningForm.frx unchanged;
      dist/excel-key-cleaner.xlam - the add-in: the three modules, built by a private Excel
      instance, with the ribbon (src/customUI14.xml) written into the package;
      dist/excel-key-cleaner.zip - the add-in, the loose files and the install guides.

    python build.py              needs Windows, Excel and pywin32 for the add-in
    python build.py --no-xlam    text files and zip only
"""
import pathlib, shutil, sys, zipfile

ROOT = pathlib.Path(__file__).parent
SRC, DIST = ROOT / 'src', ROOT / 'dist'
DIST.mkdir(exist_ok=True)
XLAM = DIST / 'excel-key-cleaner.xlam'
MODULES = ('Cleaning.bas', 'CleaningForm.frm', 'Ribbon.bas')
# what Excel shows in File > Options > Add-ins and in the Add-ins dialog
TITLE = 'Очистка ключей'
COMMENTS = ('Кириллица вместо латиницы в кодах, переносы строк, невидимые символы, лишние пробелы - '
            'в выделенных ячейках. Вкладка "Очистка ключей" на ленте. github.com/MShkolenko/excel-key-cleaner')
REL_UI = 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility'   # customUI14.xml (Office 2010+)


def build_text():
    for name in MODULES:
        text = (SRC / name).read_text(encoding='utf-8').replace('\r\n', '\n').replace('\n', '\r\n')
        (DIST / name).write_bytes(text.encode('cp1251'))
        print(f'{name}: {len(text.splitlines())} lines -> dist/ (cp1251)')
    shutil.copy(SRC / 'CleaningForm.frx', DIST / 'CleaningForm.frx')


def build_xlam():
    import os, xml.etree.ElementTree as ET
    import pythoncom, win32com.client
    raw, packed = DIST / '_excel.xlam', DIST / '_packed.xlam'
    for f in (raw, packed):
        f.unlink(missing_ok=True)
    pythoncom.CoInitialize()
    xl = win32com.client.DispatchEx('Excel.Application')   # a private instance, never the user's Excel
    wb = None
    try:
        xl.Visible = False
        xl.DisplayAlerts = False
        wb = xl.Workbooks.Add(-4167)                        # xlWBATWorksheet: one sheet
        for name in MODULES:                                # the .frx is picked up next to the .frm
            wb.VBProject.VBComponents.Import(str((DIST / name).resolve()))
        refs = {wb.VBProject.References.Item(i).Name for i in range(1, wb.VBProject.References.Count + 1)}
        if 'Office' not in refs:                            # IRibbonControl lives in the Office library
            raise RuntimeError(f'no Office library reference in a new workbook: {sorted(refs)}')
        wb.BuiltinDocumentProperties('Title').Value = TITLE
        wb.BuiltinDocumentProperties('Comments').Value = COMMENTS
        wb.SaveAs(str(raw.resolve()), 55)                   # xlOpenXMLAddIn
        wb.Close(False)
        wb = None
    finally:
        try:
            if wb is not None:
                wb.Close(False)
        finally:
            xl.Quit()
    # the ribbon is not reachable through the object model: it is a part of the package itself.
    # Written to a temporary file and checked first - a failure never leaves a half-built add-in
    # in place of the previous good one.
    try:
        with zipfile.ZipFile(raw) as src, zipfile.ZipFile(packed, 'w', zipfile.ZIP_DEFLATED) as out:
            for item in src.infolist():
                data = src.read(item.filename)
                if item.filename == '_rels/.rels':
                    rels = data.decode('utf-8')
                    if REL_UI in rels or '</Relationships>' not in rels:
                        raise RuntimeError('unexpected _rels/.rels: ' + rels[:200])
                    rels = rels.replace('</Relationships>', f'<Relationship Id="rIdKeyCleanerUI" Type="{REL_UI}" '
                                                            'Target="customUI/customUI14.xml"/></Relationships>')
                    data = rels.encode('utf-8')
                out.writestr(item, data)
            out.writestr('customUI/customUI14.xml', (SRC / 'customUI14.xml').read_bytes())
        with zipfile.ZipFile(packed) as z:
            ns = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships',
                  'c': 'http://schemas.openxmlformats.org/package/2006/content-types'}
            rels = ET.fromstring(z.read('_rels/.rels')).findall('r:Relationship', ns)
            if not any(r.get('Type') == REL_UI and r.get('Target') == 'customUI/customUI14.xml' for r in rels):
                raise RuntimeError('the ribbon relationship did not land in _rels/.rels')
            types = ET.fromstring(z.read('[Content_Types].xml')).findall('c:Default', ns)
            if not any(d.get('Extension', '').lower() == 'xml' and d.get('ContentType') == 'application/xml' for d in types):
                raise RuntimeError('no Default xml -> application/xml in [Content_Types].xml')
            ET.fromstring(z.read('customUI/customUI14.xml'))
        os.replace(packed, XLAM)
    finally:
        raw.unlink(missing_ok=True)
        packed.unlink(missing_ok=True)
    print(f'{XLAM.name}: {XLAM.stat().st_size} bytes -> dist/')


def build_zip(with_xlam):
    with zipfile.ZipFile(DIST / 'excel-key-cleaner.zip', 'w', zipfile.ZIP_DEFLATED) as z:
        if with_xlam:
            z.write(XLAM, XLAM.name)
        for name in ('Cleaning.bas', 'CleaningForm.frm', 'CleaningForm.frx'):
            z.write(DIST / name, name)
        z.write(ROOT / 'docs' / 'INSTALL.md', 'INSTALL.md')
        z.write(ROOT / 'docs' / 'ribbon.png', 'ribbon.png')              # INSTALL.md shows it
        z.write(ROOT / 'docs' / 'INSTALL-FOR-CLAUDE.md', 'INSTALL-FOR-CLAUDE.md')
    print('dist/excel-key-cleaner.zip written')


if __name__ == '__main__':
    xlam = '--no-xlam' not in sys.argv
    build_text()
    if xlam:
        build_xlam()
    elif XLAM.exists():
        print(f'note: {XLAM.name} in dist/ is from an earlier build and is NOT in the zip')
    build_zip(xlam)
