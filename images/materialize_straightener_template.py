from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from html import escape
import json

root = Path.cwd()
out = root / 'outputs' / 'straightener-20260825'
template = root / '直发器.xlsx'
target = out / '产品上传表格.xlsx'
temp = out / '_materialized_upload.xlsx'
rows = json.loads((out / 'final_upload_values.json').read_text(encoding='utf-8'))

def letter(index):
    result = ''
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        result = chr(65 + rem) + result
    return result

def cell_xml(ref, value):
    if value is None:
        return ''
    if isinstance(value, bool):
        return f'<c r="{ref}" t="b"><v>{1 if value else 0}</v></c>'
    if isinstance(value, (int, float)):
        return f'<c r="{ref}"><v>{value}</v></c>'
    text = escape(str(value), quote=False)
    return f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'

rows_xml = []
for row_index, values in enumerate(rows, start=5):
    cells = ''.join(cell_xml(f'{letter(column)}{row_index}', value) for column, value in enumerate(values))
    rows_xml.append(f'<row r="{row_index}">{cells}</row>')
rows_xml = ''.join(rows_xml)

with ZipFile(template, 'r') as source, ZipFile(temp, 'w', ZIP_DEFLATED) as destination:
    for info in source.infolist():
        data = source.read(info.filename)
        if info.filename == 'xl/worksheets/sheet1.xml':
            xml = data.decode('utf-8')
            xml = xml.replace('<dimension ref="A1:AV4"/>', f'<dimension ref="A1:AV{4+len(rows)}"/>', 1)
            xml = xml.replace('</sheetData>', f'{rows_xml}</sheetData>', 1)
            data = xml.encode('utf-8')
        destination.writestr(info, data)
temp.replace(target)
print(f'Materialized {len(rows)} rows into {target}')
