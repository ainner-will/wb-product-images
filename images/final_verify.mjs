import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";
const outDir="H:/New project/wb-product-images/images/outputs/01a00ef6-c440-74e3-83c4-ae1f8250c789";
const upload=await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(outDir,"产品上传表格.xlsx")));
const sheet=upload.worksheets.getItem("项目");
const vals=sheet.getRange("A5:BB143").values;
const titles=vals.map(r=>String(r[3]??""));
const desc=vals.map(r=>String(r[6]??""));
const codes=vals.map(r=>String(r[1]??""));
const shoeRows=vals.filter(r=>String(r[1]??"").match(/(SHOE|SHOES|BOOT|HEEL)/));
const imagesByBase=new Map();
for (const r of vals) {
  const base=String(r[1]).replace(/-(?:ONE|XS|S|M|L|XL|XXL|XXXL|\d+)$/,'');
  const photos=String(r[7]??"");
  if (!imagesByBase.has(base)) imagesByBase.set(base,photos);
  else if (imagesByBase.get(base)!==photos) throw new Error(`Inconsistent photos within ${base}`);
}
const bases=[...imagesByBase.entries()];
for(let i=0;i<bases.length;i++) for(let j=i+1;j<bases.length;j++) if(bases[i][1] && bases[i][1]===bases[j][1]) throw new Error(`Cross-SKU duplicate photos: ${bases[i][0]} ${bases[j][0]}`);
const errors=await upload.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",options:{useRegex:true,maxResults:300},summary:"final formula error scan",maxChars:3000});
const checks={
  rows:vals.length,uniqueCodes:new Set(codes).size,minDescriptionLength:Math.min(...desc.map(x=>x.length)),maxDescriptionLength:Math.max(...desc.map(x=>x.length)),
  maxTitleLength:Math.max(...titles.map(x=>x.length)),allTitlesRussianNames:titles.every(t=>/Фокалорс|Фурина|Флинс|Сангономия Кокоми|Коломбина|Навия|Дурин|Николь Рейн|Сёгун Райдэн/.test(t)),
  shoeRows:shoeRows.length,allShoeSizesCorrect:shoeRows.every(r=>Number(r[21])===Number(r[22])-1),
  crossSkuDuplicatePhotos:false,formulaErrorScan:errors.ndjson
};
for (const [name,sheetName,range] of [["final_project","项目","A1:N16"],["final_instructions","使用说明","A1:D20"]]) {
  const png=await upload.render({sheetName,range,scale:1,format:"png"});
  await fs.writeFile(path.join(outDir,`${name}.png`),new Uint8Array(await png.arrayBuffer()));
}
const map=await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(outDir,"图片URL映射表.xlsx")));
const mpng=await map.render({sheetName:"图片URL映射表",range:"A1:E12",scale:1,format:"png"});
await fs.writeFile(path.join(outDir,"final_mapping.png"),new Uint8Array(await mpng.arrayBuffer()));
await fs.writeFile(path.join(outDir,"final_checks.json"),JSON.stringify(checks,null,2),"utf8");
console.log(JSON.stringify(checks,null,2));
