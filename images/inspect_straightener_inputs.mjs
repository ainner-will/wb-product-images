import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';
import fs from 'node:fs/promises';

const files = ['直发器.xlsx', '产品信息表-202410.xlsx', 'WB产品颜色-中文版.xlsx'];
for (const name of files) {
  const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(name));
  console.log(`\n### ${name}`);
  const summary = await wb.inspect({ kind: 'workbook,sheet,table', maxChars: 14000, tableMaxRows: 30, tableMaxCols: 70, tableMaxCellChars: 140 });
  console.log(summary.ndjson);
  for (const sheetName of (await wb.inspect({ kind:'sheet', include:'id,name', maxChars:2000 })).ndjson.split('\n').filter(Boolean).map(x => JSON.parse(x).name || '').filter(Boolean)) {
    const data = await wb.inspect({kind:'table', sheetId:sheetName, range:'A1:BE35', include:'values,formulas', maxChars:18000, tableMaxRows:35, tableMaxCols:57, tableMaxCellChars:160});
    console.log(data.ndjson);
  }
  if (name === '直发器.xlsx') {
    for (const sheetName of ['项目','使用说明']) {
      try {
        const img = await wb.render({sheetName, autoCrop:'all', scale:1, format:'png'});
        await fs.mkdir('straightener_preview', {recursive:true});
        await fs.writeFile(`straightener_preview/${sheetName}.png`, new Uint8Array(await img.arrayBuffer()));
      } catch (e) { console.log(`render ${sheetName}: ${e.message}`); }
    }
  }
}

const sourceWb = await SpreadsheetFile.importXlsx(await FileBlob.load('产品信息表-202410.xlsx'));
console.log('\n### SOURCE TARGETED');
console.log((await sourceWb.inspect({kind:'table', sheetId:'Sheet1', range:'A1:Z100', include:'values,formulas', maxChars:30000, tableMaxRows:100, tableMaxCols:26, tableMaxCellChars:500})).ndjson);
const tplWb = await SpreadsheetFile.importXlsx(await FileBlob.load('直发器.xlsx'));
const image = await tplWb.render({sheetName:'项目', range:'A1:AV18', scale:1, format:'png'});
await fs.writeFile('straightener_preview/project.png', new Uint8Array(await image.arrayBuffer()));
const colorWb = await SpreadsheetFile.importXlsx(await FileBlob.load('WB产品颜色-中文版.xlsx'));
console.log('\n### COLOR CANDIDATES');
console.log((await colorWb.inspect({kind:'match', searchTerm:'湖蓝|玫红|桔色|橙色|粉红色|灰色|黑色|金色|银色', options:{useRegex:true,maxResults:100}, maxChars:8000})).ndjson);
