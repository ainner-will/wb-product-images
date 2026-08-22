import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const base="H:/New project/wb-product-images/images";
const out=path.join(base,"outputs/01a00ef6-c440-74e3-83c4-ae1f8250c789");
const mainPath=path.join(out,"产品上传表格.xlsx");
const mapPath=path.join(out,"图片URL映射表.xlsx");
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(mainPath));
const table=await wb.inspect({kind:"table",range:"项目!A1:BB110",maxChars:900000,tableMaxRows:120,tableMaxCols:60,tableMaxCellChars:4000});
const parsed=table.ndjson.split(/\r?\n/).filter(Boolean).map(JSON.parse);
const item=parsed.find(x=>Array.isArray(x.values));
if(!item) throw new Error("Could not inspect project table");
const data=item.values.slice(4).filter(r=>r.some(v=>v!==null && v!==""));
if(data.length!==106) throw new Error(`Row count ${data.length}`);
await fs.writeFile(path.join(out,"data_rows.json"),JSON.stringify(data),"utf8");
const groups=new Map();
for(const r of data){const k=String(r[1]);if(!groups.has(k))groups.set(k,[]);groups.get(k).push(r);}
if(groups.size!==16) throw new Error(`SKU count ${groups.size}`);
const sizeCols=new Set([21,22,37]);
for(const [sku,rs] of groups){
 const a=rs[0];
 for(const r of rs.slice(1))for(let i=0;i<54;i++)if(!sizeCols.has(i)&&JSON.stringify(r[i])!==JSON.stringify(a[i]))throw new Error(`Mismatch ${sku} col ${i+1}`);
}
const shoeRows=data.filter(r=>["SDR-SHOES","ARL-SHOES"].includes(r[1]));
if(shoeRows.length!==20 || shoeRows.some(r=>Number(r[21])!==Number(r[22])-1))throw new Error("Shoe conversion failed");
const titles=[...groups.values()].map(rs=>String(rs[0][3]));
if(Math.max(...titles.map(x=>x.length))>60)throw new Error("Title >60");
const descriptions=[...groups.values()].map(rs=>String(rs[0][6]));
if(new Set(descriptions).size!==16 || descriptions.some(x=>x.length<1600||x.length>2000))throw new Error("Description uniqueness/length failed");
const urlsBySku=new Map(); const allUrls=[];
for(const [sku,rs] of groups){const urls=String(rs[0][7]||"").split(";").filter(Boolean);urlsBySku.set(sku,urls);for(const u of urls){if(!u.startsWith("https://raw.githubusercontent.com/ainner-will/wb-product-images/main/images/")||u.includes(" "))throw new Error(`Bad URL ${u}`);allUrls.push([sku,u]);}}
const seen=new Map();
for(const [sku,u] of allUrls){if(seen.has(u)&&seen.get(u)!==sku)throw new Error(`Cross SKU URL ${u}`);seen.set(u,sku);}
const errors=await wb.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",options:{useRegex:true,maxResults:200},summary:"formula error scan"});
await fs.writeFile(path.join(out,"formula_error_scan.ndjson"),errors.ndjson,"utf8");
const renderDir=path.join(out,"renders"); await fs.mkdir(renderDir,{recursive:true});
for(const [name,sheetName,range] of [["project_left","项目","A1:Q12"],["project_middle","项目","R1:AJ12"],["project_right","项目","AK1:BB12"],["project_tail","项目","A100:BB110"],["instructions","使用说明","A1:D20"]]){
 const png=await wb.render({sheetName,range,scale:1,format:"png"}); await fs.writeFile(path.join(renderDir,`${name}.png`),new Uint8Array(await png.arrayBuffer()));
}
const mwb=await SpreadsheetFile.importXlsx(await FileBlob.load(mapPath));
const mpng=await mwb.render({sheetName:"图片URL映射表",range:"A1:E25",scale:1,format:"png"});
await fs.writeFile(path.join(renderDir,"mapping.png"),new Uint8Array(await mpng.arrayBuffer()));
const mapInspect=await mwb.inspect({kind:"table",range:"图片URL映射表!A1:E120",maxChars:300000,tableMaxRows:130,tableMaxCols:8,tableMaxCellChars:2000});
const mitem=mapInspect.ndjson.split(/\r?\n/).filter(Boolean).map(JSON.parse).find(x=>Array.isArray(x.values));
const mapRows=mitem.values.slice(1).filter(r=>r.some(v=>v!==null&&v!==""));
if(mapRows.length!==113)throw new Error(`Map rows ${mapRows.length}`);
const errorRecords=errors.ndjson.split(/\r?\n/).filter(Boolean).map(JSON.parse).filter(x=>x.kind==="match");
const result={rows:data.length,skus:groups.size,articlesUnique:new Set(data.map(r=>r[1])).size,shoeRows:shoeRows.length,shoeConversion:true,maxTitle:Math.max(...titles.map(x=>x.length)),uniqueDescriptions:new Set(descriptions).size,descriptionMin:Math.min(...descriptions.map(x=>x.length)),descriptionMax:Math.max(...descriptions.map(x=>x.length)),imageUrls:allUrls.length,mapRows:mapRows.length,videoUrls:0,formulaErrors:errorRecords.length,missing:[...groups.keys()].filter(k=>!(urlsBySku.get(k)||[]).length)};
await fs.writeFile(path.join(out,"final_verification.json"),JSON.stringify(result,null,2),"utf8");
console.log(JSON.stringify(result,null,2));
