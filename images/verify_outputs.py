from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import json, re

base=Path(r"H:\New project\wb-product-images\images")
orig=base/"化妆舞会服装.xlsx"
out=base/"outputs"/"01a00ef6-c440-74e3-83c4-ae1f8250c789"/"产品上传表格.xlsx"
NS={"m":"http://schemas.openxmlformats.org/spreadsheetml/2006/main","r":"http://schemas.openxmlformats.org/officeDocument/2006/relationships"}

def xml(z,n): return ET.fromstring(z.read(n))
def snapshot(p):
    with ZipFile(p) as z:
        wb=xml(z,"xl/workbook.xml")
        sheets=[(s.attrib.get("name"),s.attrib.get("state","visible")) for s in wb.find("m:sheets",NS)]
        result={"sheets":sheets,"worksheets":{}}
        for n in ["xl/worksheets/sheet1.xml","xl/worksheets/sheet2.xml"]:
            root=xml(z,n)
            cols=[tuple(sorted(c.attrib.items())) for c in root.findall("m:cols/m:col",NS)]
            merges=[x.attrib.get("ref") for x in root.findall("m:mergeCells/m:mergeCell",NS)]
            pane=root.find("m:sheetViews/m:sheetView/m:pane",NS)
            dvs=root.find("m:dataValidations",NS)
            result["worksheets"][n]={
                "cols":cols,"merges":merges,"pane":None if pane is None else dict(pane.attrib),
                "dataValidationCount":0 if dvs is None else int(dvs.attrib.get("count",len(dvs))),
                "hiddenCols":[dict(c.attrib) for c in root.findall("m:cols/m:col",NS) if c.attrib.get("hidden")=="1"]
            }
        return result

a,b=snapshot(orig),snapshot(out)
comparison={
    "sheetNamesStatesEqual":a["sheets"]==b["sheets"],
    "sheet1ColsEqual":a["worksheets"]["xl/worksheets/sheet1.xml"]["cols"]==b["worksheets"]["xl/worksheets/sheet1.xml"]["cols"],
    "sheet1MergesEqual":a["worksheets"]["xl/worksheets/sheet1.xml"]["merges"]==b["worksheets"]["xl/worksheets/sheet1.xml"]["merges"],
    "sheet1PaneEqual":a["worksheets"]["xl/worksheets/sheet1.xml"]["pane"]==b["worksheets"]["xl/worksheets/sheet1.xml"]["pane"],
    "sheet1ValidationCountEqual":a["worksheets"]["xl/worksheets/sheet1.xml"]["dataValidationCount"]==b["worksheets"]["xl/worksheets/sheet1.xml"]["dataValidationCount"],
    "sheet2ColsEqual":a["worksheets"]["xl/worksheets/sheet2.xml"]["cols"]==b["worksheets"]["xl/worksheets/sheet2.xml"]["cols"],
    "sheet2MergesEqual":a["worksheets"]["xl/worksheets/sheet2.xml"]["merges"]==b["worksheets"]["xl/worksheets/sheet2.xml"]["merges"],
    "sheet2PaneEqual":a["worksheets"]["xl/worksheets/sheet2.xml"]["pane"]==b["worksheets"]["xl/worksheets/sheet2.xml"]["pane"],
}
print(json.dumps({"comparison":comparison,"original":a,"output":b},ensure_ascii=False,indent=2))
