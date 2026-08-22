import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";
const base="H:/New project/wb-product-images/images";
const out=path.join(base,"work_preview");
await fs.mkdir(out,{recursive:true});
for(const name of ["化妆舞会服装.xlsx","COS服装及配饰清单.xlsx","WB产品颜色-中文版.xlsx"]){
 const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(base,name)));
 const x=await wb.inspect({kind:"workbook,sheet,table",maxChars:20000,tableMaxRows:40,tableMaxCols:80,tableMaxCellChars:300});
 await fs.writeFile(path.join(out,`${name}.inspect.txt`),x.ndjson,"utf8");
 const sheets=await wb.inspect({kind:"sheet",include:"id,name",maxChars:3000});
 const parsed=sheets.ndjson.split(/\r?\n/).filter(Boolean).map(JSON.parse);
 for(let i=0;i<parsed.length;i++){
   const s=parsed[i].name??parsed[i].sheetName;
   if(!s)continue;
   try{const png=await wb.render({sheetName:s,range:s==="项目"?"A1:BB8":undefined,autoCrop:s==="项目"?undefined:"all",scale:1,format:"png"});await fs.writeFile(path.join(out,`${name}.${i}.png`),new Uint8Array(await png.arrayBuffer()));}catch(e){await fs.appendFile(path.join(out,"errors.txt"),`${name}|${s}|${e}\n`);}
 }
}
const dirs=(await fs.readdir(base,{withFileTypes:true})).filter(d=>d.isDirectory()&&!['work_preview','outputs','node_modules'].includes(d.name));
const inv=[];
for(const d of dirs){for(const e of await fs.readdir(path.join(base,d.name),{withFileTypes:true})){if(e.isFile())inv.push({folder:d.name,file:e.name,ext:path.extname(e.name).toLowerCase(),full:path.join(base,d.name,e.name)});}}
await fs.writeFile(path.join(out,"inventory.json"),JSON.stringify(inv,null,2),"utf8");
console.log(JSON.stringify({folders:dirs.length,files:inv.length}));
