from copy import deepcopy
from pathlib import Path
import os, shutil, tempfile, zipfile
import xml.etree.ElementTree as ET

base = Path(r"H:\New project\wb-product-images\images")
source = base / "化妆舞会服装.xlsx"
target = base / "outputs" / "01a00ef6-c440-74e3-83c4-ae1f8250c789" / "产品上传表格.xlsx"
main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
ET.register_namespace("", main)
ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
ET.register_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
ET.register_namespace("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")
ET.register_namespace("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision")
ET.register_namespace("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2")
ET.register_namespace("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3")

with zipfile.ZipFile(source) as zs, zipfile.ZipFile(target) as zt:
    replacements = {}
    for name in ("xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"):
        original = ET.fromstring(zs.read(name))
        edited = ET.fromstring(zt.read(name))
        edited_dimension = edited.find(f"{{{main}}}dimension")
        edited_sheet_data = edited.find(f"{{{main}}}sheetData")
        if edited_sheet_data is None:
            raise RuntimeError(f"Missing sheetData in {name}")
        rebuilt = ET.Element(original.tag, original.attrib)
        for child in original:
            local = child.tag.rsplit("}", 1)[-1]
            if local == "dimension":
                if edited_dimension is not None:
                    rebuilt.append(deepcopy(edited_dimension))
            elif local == "sheetData":
                rebuilt.append(deepcopy(edited_sheet_data))
            else:
                rebuilt.append(deepcopy(child))
        replacements[name] = ET.tostring(rebuilt, encoding="utf-8", xml_declaration=True)
    fd, temp_name = tempfile.mkstemp(suffix=".xlsx", dir=target.parent)
    os.close(fd)
    try:
        with zipfile.ZipFile(temp_name, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zt.infolist():
                zout.writestr(item, replacements.get(item.filename, zt.read(item.filename)))
        shutil.copyfile(temp_name, target)
    finally:
        if os.path.exists(temp_name):
            os.unlink(temp_name)

print(target)
