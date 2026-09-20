import fs from 'node:fs/promises';
import path from 'node:path';
import JSZip from 'jszip';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const root=path.resolve('.');
const out=path.join(root,'outputs','phone-case-2026-09-20');
const original=await JSZip.loadAsync(await fs.readFile(path.join(root,'.work_phone_case','before-width-fix.xlsx')));
const result=await JSZip.loadAsync(await fs.readFile(path.join(out,'产品上传表格.xlsx')));
const names=Object.keys(original.files).filter(n=>!original.files[n].dir).sort();
const resultNames=Object.keys(result.files).filter(n=>!result.files[n].dir).sort();
if(JSON.stringify(names)!==JSON.stringify(resultNames)) throw Error('Template package entries changed');
for(const n of names){if(n==='xl/worksheets/sheet1.xml')continue;const a=await original.file(n).async('nodebuffer'),b=await result.file(n).async('nodebuffer');if(!a.equals(b))throw Error('Modified template part '+n)}
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(out,'产品上传表格.xlsx')));
if(wb.worksheets.items.map(x=>x.name).join('|')!=='项目|使用说明') throw Error('Sheet structure changed');
const rows=wb.worksheets.getItem('项目').getRange('A5:AQ18').values;
if(rows.length!==14)throw Error('Wrong product row count');
const prior=await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(root,'.work_phone_case','before-width-fix.xlsx')));
const oldRows=prior.worksheets.getItem('项目').getRange('A5:AQ18').values;
const changedColumns=new Set([35,36,37,38,39,40]);
for(let i=0;i<14;i++){
 for(let j=0;j<43;j++)if(!changedColumns.has(j)&&rows[i][j]!==oldRows[i][j])throw Error(`Unexpected change ${i+5}, ${j+1}`);
 const expectedWidth=9,expectedLength=i<8?12:13;
 if(rows[i][35]!=='1 чехол'||rows[i][36]!==expectedWidth||rows[i][37]!==expectedLength||rows[i][38]!==2||rows[i][39]!==200||rows[i][40]!==null)throw Error(`Wrong packaging positions row ${i+5}`);
 if(typeof rows[i][36]!=='number')throw Error(`Width is not numeric in row ${i+5}`);
}
const sheetXml=await result.file('xl/worksheets/sheet1.xml').async('string');
for(let r=5;r<=18;r++)if(!new RegExp(`<c r="AK${r}"><v>9<\\/v><\\/c>`).test(sheetXml))throw Error(`AK${r} is not stored as numeric XML`);
const ids=new Set(),urls=new Set();let imageAssignments=0;
for(const row of rows){
 const sku=row[1];if(ids.has(sku))throw Error('Duplicate sku '+sku);ids.add(sku);
 if(!sku||!row[3]||!row[6]||!row[7]||!row[8])throw Error('Missing core field '+sku);
 if(row[3].length>60||row[6].length>2000)throw Error('Length limit '+sku);
 const images=row[7].split(';'); imageAssignments+=images.length;
 for(const u of images){if(urls.has(u))throw Error('Duplicate image URL '+u);urls.add(u)}
 const videoExpected=sku.startsWith('FC-')?'/视频1/':'/视频/';
 const videoDecoded=decodeURIComponent(new URL(row[8]).pathname);
 if(!videoDecoded.includes(videoExpected))throw Error('Wrong video '+sku);
 for(const ci of [2,5,9,15,17,19,20,21,22,23,24,26])if(row[ci]!==null&&row[ci]!=='')throw Error('Restricted/unsupported field '+sku+' '+ci+' '+row[ci]);
}
if(imageAssignments!==114)throw Error('Wrong image total');
const mapping=await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(out,'图片URL映射表.xlsx')));
const pr=mapping.worksheets.getItem('图片URL映射').getRange('A2:E115').values;
const vr=mapping.worksheets.getItem('视频URL映射').getRange('A2:F15').values;
if(pr.length!==114||vr.length!==14)throw Error('Wrong mapping row count');
for(const p of pr){if(!urls.has(p[3]))throw Error('Unmatched mapped URL '+p[3]);await fs.access(p[2]);}
for(const v of vr){await fs.access(v[2]);if(!ids.has(v[5]))throw Error('Unmatched video sku '+v[5]);}
const preview=await wb.render({sheetName:'项目',range:'A1:F8',scale:0.7,format:'png'});
await fs.writeFile(path.join(root,'.work_phone_case','final-preview.png'),new Uint8Array(await preview.arrayBuffer()));
console.log(JSON.stringify({products:rows.length,uniqueSKUs:ids.size,images:imageAssignments,videoAssignments:vr.length,uniqueVideos:new Set(vr.map(r=>r[3])).size,templateOtherPartsUnchanged:names.length-1,sheets:wb.worksheets.items.map(x=>x.name),outputFiles:await fs.readdir(out)},null,2));
