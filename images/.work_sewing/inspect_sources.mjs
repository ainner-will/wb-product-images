import fs from 'node:fs/promises';
import { FileBlob, SpreadsheetFile, Workbook } from '@oai/artifact-tool';

const [templatePath, colorPath, csvPath] = process.argv.slice(2);

async function inspectBook(path, label) {
  const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(path));
  const summary = await wb.inspect({ kind: 'workbook,sheet,table', maxChars: 12000, tableMaxRows: 8, tableMaxCols: 120, tableMaxCellChars: 140 });
  console.log(`=== ${label} SUMMARY ===`);
  console.log(summary.ndjson);
  for (const ws of wb.worksheets.items) {
    const used = label === 'TEMPLATE' && ws.name === '项目' ? ws.getRange('A1:AV15') : ws.getUsedRange(true);
    console.log(`=== ${label} SHEET ${ws.name} USED ===`);
    console.log(JSON.stringify({ address: used?.address, values: used?.values }, null, 2));
    if (label === 'TEMPLATE') {
      const range = ws.name === '项目' ? 'A1:AV10' : 'A1:D20';
      const img = await wb.render({ sheetName: ws.name, range, scale: 0.8, format: 'png' });
      await fs.writeFile(`.work_sewing/template-${encodeURIComponent(ws.name)}.png`, new Uint8Array(await img.arrayBuffer()));
    }
  }
}

await inspectBook(templatePath, 'TEMPLATE');
await inspectBook(colorPath, 'COLORS');

const bytes = await fs.readFile(csvPath);
for (const enc of ['utf8', 'gb18030']) {
  const text = new TextDecoder(enc).decode(bytes);
  const wb = await Workbook.fromCSV(text, { sheetName: enc });
  const used = wb.worksheets.getItem(enc).getUsedRange();
  console.log(`=== CSV ${enc} ===`);
  console.log(JSON.stringify(used.values, null, 2));
}
