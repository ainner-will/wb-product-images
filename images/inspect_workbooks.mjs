import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const base = "H:/New project/wb-product-images/images";
const out = path.join(base, "work_preview");
await fs.mkdir(out, { recursive: true });

for (const name of ["化妆舞会服装.xlsx", "COS服装及配饰清单.xlsx", "WB产品颜色-中文版.xlsx"]) {
  const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(base, name)));
  const summary = await wb.inspect({kind:"workbook,sheet,table", maxChars:18000, tableMaxRows:20, tableMaxCols:80, tableMaxCellChars:300});
  await fs.writeFile(path.join(out, `${name}.inspect.txt`), summary.ndjson, "utf8");
  const sheets = await wb.inspect({kind:"sheet", include:"id,name", maxChars:4000});
  await fs.writeFile(path.join(out, `${name}.sheets.txt`), sheets.ndjson, "utf8");
  const parsed = sheets.ndjson.split(/\r?\n/).filter(Boolean).map(x=>JSON.parse(x));
  for (let i=0; i<parsed.length; i++) {
    const sheetName = parsed[i].name ?? parsed[i].sheetName;
    if (!sheetName) continue;
    try {
      const png = await wb.render({sheetName, autoCrop:"all", scale:0.8, format:"png"});
      await fs.writeFile(path.join(out, `${name}.${i}.${sheetName.replace(/[\\/:*?\"<>|]/g,"_")}.png`), new Uint8Array(await png.arrayBuffer()));
    } catch (e) {
      await fs.appendFile(path.join(out, "render_errors.txt"), `${name} | ${sheetName} | ${e}\n`);
    }
  }
}

const dirs = (await fs.readdir(base, {withFileTypes:true})).filter(d=>d.isDirectory() && d.name!=="work_preview" && d.name!=="outputs");
const inventory = [];
for (const d of dirs) {
  const folder = path.join(base, d.name);
  const entries = await fs.readdir(folder, {withFileTypes:true});
  for (const e of entries) {
    const full = path.join(folder,e.name);
    if (e.isFile()) inventory.push({skuFolder:d.name, subfolder:"", file:e.name, ext:path.extname(e.name).toLowerCase(), full});
    if (e.isDirectory()) {
      const files = await fs.readdir(full,{withFileTypes:true});
      for (const f of files) if (f.isFile()) inventory.push({skuFolder:d.name, subfolder:e.name, file:f.name, ext:path.extname(f.name).toLowerCase(), full:path.join(full,f.name)});
    }
  }
}
await fs.writeFile(path.join(out,"inventory.json"), JSON.stringify(inventory,null,2),"utf8");
console.log(`inspected ${dirs.length} folders, ${inventory.length} files`);
