from zipfile import ZipFile
from xml.etree import ElementTree as ET
from pathlib import Path
import json

root = Path.cwd()
base_path = root / '直发器.xlsx'
final_path = root / 'outputs' / 'straightener-20260825' / '产品上传表格.xlsx'
ns = {'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

def get_xml(z, name):
    return ET.fromstring(z.read(name))

def xml_text(element):
    return None if element is None else ET.tostring(element)

with ZipFile(base_path) as b, ZipFile(final_path) as f:
    bw, fw = get_xml(b,'xl/workbook.xml'), get_xml(f,'xl/workbook.xml')
    base_names = [x.attrib['name'] for x in bw.findall('.//m:sheets/m:sheet',ns)]
    final_names = [x.attrib['name'] for x in fw.findall('.//m:sheets/m:sheet',ns)]
    checks = {'sheet_names':base_names==final_names}
    for sheet_file, key in [('xl/worksheets/sheet1.xml','project'),('xl/worksheets/sheet2.xml','instructions')]:
        bx, fx = get_xml(b,sheet_file), get_xml(f,sheet_file)
        for tag in ['sheetPr','sheetViews','sheetFormatPr','cols','mergeCells','dataValidations','conditionalFormatting','tableParts','autoFilter']:
            be = bx.find(f'm:{tag}',ns)
            fe = fx.find(f'm:{tag}',ns)
            checks[f'{key}_{tag}'] = xml_text(be) == xml_text(fe)
        if key == 'project':
            brow = bx.find(".//m:sheetData/m:row[@r='3']",ns)
            frow = fx.find(".//m:sheetData/m:row[@r='3']",ns)
            checks['project_header_row_style'] = ET.tostring(brow) == ET.tostring(frow)
        else:
            checks['instructions_sheetdata'] = ET.tostring(bx.find('m:sheetData',ns)) == ET.tostring(fx.find('m:sheetData',ns))
    checks['styles'] = b.read('xl/styles.xml') == f.read('xl/styles.xml')
print(json.dumps(checks, ensure_ascii=False, indent=2))
if not all(checks.values()):
    raise SystemExit(1)
