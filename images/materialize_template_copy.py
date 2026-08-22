from copy import deepcopy
from pathlib import Path
import json, os, shutil, tempfile, zipfile
import xml.etree.ElementTree as ET

base = Path(r"H:\New project\wb-product-images\images")
source = base / "化妆舞会服装.xlsx"
outdir = base / "outputs" / "01a00ef6-c440-74e3-83c4-ae1f8250c789"
target = outdir / "产品上传表格.xlsx"
rows = json.loads((outdir / "data_rows.json").read_text(encoding="utf-8"))
main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
xmlns = "http://www.w3.org/XML/1998/namespace"
ET.register_namespace("", main)
ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
ET.register_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
ET.register_namespace("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")
ET.register_namespace("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision")
ET.register_namespace("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2")
ET.register_namespace("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3")

def colname(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

with zipfile.ZipFile(source) as zin:
    root = ET.fromstring(zin.read("xl/worksheets/sheet1.xml"))
    dim = root.find(f"{{{main}}}dimension")
    dim.set("ref", "A1:EIV110")
    data = root.find(f"{{{main}}}sheetData")
    rowmap = {int(r.attrib["r"]): r for r in data.findall(f"{{{main}}}row")}
    for ridx, values in enumerate(rows, start=5):
        row = rowmap.get(ridx)
        if row is None:
            row = ET.Element(f"{{{main}}}row", {"r": str(ridx)})
            data.append(row)
            rowmap[ridx] = row
        existing = {c.attrib["r"]: c for c in row.findall(f"{{{main}}}c")}
        for cidx, value in enumerate(values, start=1):
            if value is None or value == "":
                continue
            address = f"{colname(cidx)}{ridx}"
            cell = existing.get(address)
            if cell is None:
                cell = ET.Element(f"{{{main}}}c", {"r": address})
                row.append(cell)
            else:
                for child in list(cell):
                    cell.remove(child)
            if isinstance(value, bool):
                cell.set("t", "b")
                ET.SubElement(cell, f"{{{main}}}v").text = "1" if value else "0"
            elif isinstance(value, (int, float)):
                cell.attrib.pop("t", None)
                ET.SubElement(cell, f"{{{main}}}v").text = str(value)
            else:
                cell.set("t", "inlineStr")
                inline = ET.SubElement(cell, f"{{{main}}}is")
                text = ET.SubElement(inline, f"{{{main}}}t")
                if str(value) != str(value).strip() or "\n" in str(value):
                    text.set(f"{{{xmlns}}}space", "preserve")
                text.text = str(value)
        def cnum(c):
            letters = ''.join(x for x in c.attrib.get("r", "") if x.isalpha())
            n = 0
            for ch in letters: n = n * 26 + ord(ch.upper()) - 64
            return n
        row[:] = sorted(list(row), key=cnum)
    data[:] = sorted(list(data), key=lambda r: int(r.attrib.get("r", "0")))
    sheet_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    fd, temp_name = tempfile.mkstemp(suffix=".xlsx", dir=outdir)
    os.close(fd)
    try:
        with zipfile.ZipFile(temp_name, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                zout.writestr(item, sheet_xml if item.filename == "xl/worksheets/sheet1.xml" else zin.read(item.filename))
        shutil.copyfile(temp_name, target)
    finally:
        if os.path.exists(temp_name): os.unlink(temp_name)

print(target)
