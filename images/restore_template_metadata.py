from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import re, tempfile, shutil

base=Path(r"H:\New project\wb-product-images\images")
src=base/"化妆舞会服装.xlsx"
dst=base/"outputs"/"01a00ef6-c440-74e3-83c4-ae1f8250c789"/"产品上传表格.xlsx"
targets=["xl/worksheets/sheet1.xml","xl/worksheets/sheet2.xml"]

with ZipFile(src,"r") as z:
    originals={n:z.read(n) for n in targets}

temp=dst.with_suffix(".metadata-fix.tmp.xlsx")
with ZipFile(dst,"r") as zin, ZipFile(temp,"w",ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data=zin.read(item.filename)
        if item.filename in targets:
            old=originals[item.filename]
            for tag in (b"sheetViews",b"cols"):
                pattern=rb"<(?:[A-Za-z0-9_]+:)?"+tag+rb"(?:\s[^>]*)?>.*?</(?:[A-Za-z0-9_]+:)?"+tag+rb">"
                source_match=re.search(pattern,old,flags=re.S)
                if source_match:
                    source_block=re.sub(rb"<(/?)([A-Za-z][A-Za-z0-9]*)(?=[\s>/])",rb"<\1x:\2",source_match.group(0))
                    if re.search(pattern,data,flags=re.S):
                        data=re.sub(pattern,lambda _:source_block,data,count=1,flags=re.S)
                    else:
                        anchor=re.search(rb"<(?:(?:[A-Za-z0-9_]+:)?worksheet)\b[^>]*>",data)
                        if not anchor:
                            raise RuntimeError(f"No insertion anchor in {item.filename}")
                        pos=anchor.end()
                        data=data[:pos]+source_block+data[pos:]
        zout.writestr(item,data)
shutil.move(temp,dst)
print(dst)
