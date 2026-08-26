import fs from 'node:fs/promises';
import path from 'node:path';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const root = process.cwd();
const out = path.join(root,'outputs','straightener-20260825');
const uploadPath = path.join(out,'产品上传表格.xlsx');
const rows = JSON.parse(await fs.readFile(path.join(out,'straightener_rows.json'),'utf8'));
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(uploadPath));
const base = await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(root,'直发器.xlsx')));
const sheet = wb.worksheets.getItem('项目');
const finalRows = sheet.getRange(`A5:AV${4+rows.length}`).values;
await fs.writeFile(path.join(out,'final_upload_values.json'), JSON.stringify(finalRows));
const failures = [];
const allowedColors = new Set(['灰色','粉红色','黑色','金色','蓝色','橙色','银色','白色']);
const blankColumns = [5,9,11,12,15,17,19,20,21,22,23,27,38,43];
let imageUrls = 0;
for (let i=0;i<rows.length;i++) {
  const e=rows[i], a=finalRows[i];
  const add=(cond,msg)=>{if(!cond)failures.push(`${e.sku}: ${msg}`)};
  add(a[0]===1,'团体'); add(a[1]===e.sku,'编码'); add(a[4]==='直发器','物品'); add(a[10]===e.weight,'物流重量');
  add(a[18]===e.price,'价格取整'); add(a[31]==='Китай','生产国'); add(a[34]===e.width && a[35]===e.length && a[36]===e.height,'包装尺寸'); add(a[37]===e.weight*1000,'毛重');
  add(typeof a[3]==='string' && a[3].length<=60 && !a[3].includes(e.sku),'标题');
  add(typeof a[6]==='string' && a[6].length>=1600 && a[6].length<=2000 && !a[6].includes(e.sku),'说明书');
  add(a[16].split(';').every(x=>allowedColors.has(x)),'颜色');
  for (const c of blankColumns) add(a[c]===null || a[c]==='',`应留空列 ${c+1}`);
  const urls = a[7] ? a[7].split(';') : [];
  imageUrls += urls.length;
  add((e.images===0) === (urls.length===0),'图片数量');
  for (const url of urls) add(url.startsWith('https://raw.githubusercontent.com/ainner-will/wb-product-images/main/images/') && !url.includes(' ') && url.includes(encodeURIComponent(e.sku)+'/'),'图片URL');
  add(a[8]===null || a[8]==='','视频留空');
  for (let c=0;c<a.length;c++) if (![6,7,8].includes(c) && typeof a[c]==='string') add(a[c].length<=100,`字符数列 ${c+1}`);
}
const descriptions=finalRows.map(r=>r[6]);
if(new Set(descriptions).size!==descriptions.length) failures.push('说明书重复');
if(new Set(finalRows.map(r=>r[1])).size!==finalRows.length) failures.push('编码重复');
const errors = await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A',options:{useRegex:true,maxResults:300},maxChars:5000});
const hasFormulaErrors = !errors.ndjson.includes('matched 0 entries');
if(hasFormulaErrors) failures.push(`公式错误 ${errors.ndjson}`);
const parseNames = x => x.split('\n').filter(Boolean).map(line=>JSON.parse(line).name);
const baseSheets = parseNames((await base.inspect({kind:'sheet',include:'id,name',maxChars:1000})).ndjson);
const finalSheets = parseNames((await wb.inspect({kind:'sheet',include:'id,name',maxChars:1000})).ndjson);
const baseHeaders=base.worksheets.getItem('项目').getRange('A1:AV4').values;
const finalHeaders=sheet.getRange('A1:AV4').values;
if(JSON.stringify(baseHeaders)!==JSON.stringify(finalHeaders)) failures.push('表头内容改变');
if(JSON.stringify(baseSheets)!==JSON.stringify(finalSheets)) failures.push('工作表结构改变');
const baseInstructions=base.worksheets.getItem('使用说明').getRange('A1:H20').values;
const finalInstructions=wb.worksheets.getItem('使用说明').getRange('A1:H20').values;
if(JSON.stringify(baseInstructions)!==JSON.stringify(finalInstructions)) failures.push('说明页内容改变');
await fs.mkdir(path.join(out,'renders'),{recursive:true});
for (const [name,range] of [['left','A1:N24'],['middle','O1:AB24'],['right','AC1:AV24'],['instructions','A1:H20']]) {
  const img=await wb.render({sheetName:name==='instructions'?'使用说明':'项目',range,scale:1,format:'png'});
  await fs.writeFile(path.join(out,'renders',`${name}.png`),new Uint8Array(await img.arrayBuffer()));
}
const mapWb=await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(out,'图片URL映射表.xlsx')));
const mapImg=await mapWb.render({sheetName:'图片URL映射表',range:'A1:G20',scale:1,format:'png'});
await fs.writeFile(path.join(out,'renders','mapping.png'),new Uint8Array(await mapImg.arrayBuffer()));
console.log(JSON.stringify({rows:rows.length,imageUrls,formulaErrors:hasFormulaErrors?1:0,uniqueCodes:new Set(finalRows.map(r=>r[1])).size,uniqueDescriptions:new Set(descriptions).size,descriptions:[Math.min(...descriptions.map(x=>x.length)),Math.max(...descriptions.map(x=>x.length))],failures},null,2));
if(failures.length) process.exitCode=1;
