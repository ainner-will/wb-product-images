import fs from 'node:fs/promises';
import path from 'node:path';
import JSZip from 'jszip';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';

const root=path.resolve('.');
const baseline=path.join(root,'.work_phone_case','before-width-fix.xlsx');
const target=path.join(root,'outputs','phone-case-2026-09-20','产品上传表格.xlsx');
const workbook=await SpreadsheetFile.importXlsx(await FileBlob.load(baseline));
const sheet=workbook.worksheets.getItem('项目');
for(let r=5;r<=18;r++){
 const length=r<=12?12:13;
 sheet.getRange(`AJ${r}:AO${r}`).values=[['1 чехол',9,length,2,200,null]];
}
workbook.recalculate();
const draft=path.join(root,'.work_phone_case','width-artifact-draft.xlsx');
await (await SpreadsheetFile.exportXlsx(workbook)).save(draft);
const zip=await JSZip.loadAsync(await fs.readFile(baseline));
let xml=await zip.file('xl/worksheets/sheet1.xml').async('string');
for(let r=5;r<=18;r++){
 const length=r<=12?12:13;
 const rowRe=new RegExp(`<row r="${r}"[^>]*>[\\s\\S]*?<\\/row>`);
 const original=xml.match(rowRe)?.[0];
 if(!original)throw Error(`Missing row ${r}`);
 const priorWidth=original.match(new RegExp(`<c r="AK${r}"[^>]*>[\\s\\S]*?<\\/c>`))?.[0];
 if(!priorWidth)throw Error(`Missing original width in AK${r}`);
 let updated=original.replace(new RegExp(`<c r="(?:AJ|AK|AL|AM|AN|AO)${r}"[^>]*>[\\s\\S]*?<\\/c>`,'g'),'');
 const newCells=`<c r="AJ${r}" t="inlineStr"><is><t>1 чехол</t></is></c><c r="AK${r}"><v>9</v></c><c r="AL${r}"><v>${length}</v></c><c r="AM${r}"><v>2</v></c><c r="AN${r}"><v>200</v></c>`;
 updated=updated.replace('</row>',newCells+'</row>');
 xml=xml.replace(original,updated);
}
zip.file('xl/worksheets/sheet1.xml',xml);
await fs.writeFile(target,await zip.generateAsync({type:'nodebuffer',compression:'DEFLATE'}));
await fs.rm(draft);
await fs.rm(draft+'.inspect.ndjson',{force:true});
console.log('Updated AJ:AN for rows 5:18 from existing workbook');
