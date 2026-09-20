import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
for(const [name,label] of [['手机壳.xlsx','template'],['适用iPhone18Fold折叠屏支架手机壳磁吸无线.xlsx','data'],['WB产品颜色-中文版.xlsx','colors']]){
 const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(name));
 console.log('BOOK',label,(await wb.inspect({kind:'sheet',include:'id,name',maxChars:5000})).ndjson);
 for(const ws of wb.worksheets.items){
  const ur=ws.getUsedRange(true);
  console.log('SHEET',label,ws.name,ur?.address);
  if(label==='template'){
   console.log('HEADERS',JSON.stringify(ws.getRange('A1:CL4').values));
   const im=await wb.render({sheetName:ws.name,range:ws.name==='项目'?'A1:P8':'A1:D20',scale:.7,format:'png'});
   await fs.writeFile(`.work_phone_case/${label}-${encodeURIComponent(ws.name)}.png`,new Uint8Array(await im.arrayBuffer()));
  }else if(label==='data') console.log('VALUES',JSON.stringify(ur.values));
  else console.log('COLORS',JSON.stringify(ur.values.flat().filter(Boolean).filter(x=>/^(黑|白|灰|金|橙|紫|红|透|蓝|褐|咖啡|深)/.test(String(x)))));
 }
}
