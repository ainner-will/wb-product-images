import fs from 'node:fs/promises';
import path from 'node:path';
import JSZip from 'jszip';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const root = process.cwd();
const outDir = path.join(root, 'outputs', '01a04c8f-2214-74f1-959d-5098f7c43bb6-sewing');
const templatePath = path.join(root, '缝纫机.xlsx');
const productPath = path.join(outDir, '产品上传表格.xlsx');
const mappingPath = path.join(outDir, '图片URL映射表.xlsx');

const [origZip, outZip] = await Promise.all([
  JSZip.loadAsync(await fs.readFile(templatePath)),
  JSZip.loadAsync(await fs.readFile(productPath)),
]);
const origEntries = Object.keys(origZip.files).sort();
const outEntries = Object.keys(outZip.files).sort();
const exactEntries = JSON.stringify(origEntries) === JSON.stringify(outEntries);
const onlyOriginalEntries = origEntries.filter(x => !outEntries.includes(x));
const onlyOutputEntries = outEntries.filter(x => !origEntries.includes(x));

const sameFiles = ['xl/workbook.xml', 'xl/styles.xml', 'xl/_rels/workbook.xml.rels', 'xl/worksheets/sheet2.xml', '[Content_Types].xml'];
const equality = {};
for (const f of sameFiles) equality[f] = (await origZip.file(f).async('string')) === (await outZip.file(f).async('string'));

const origSheet = await origZip.file('xl/worksheets/sheet1.xml').async('string');
const outSheet = await outZip.file('xl/worksheets/sheet1.xml').async('string');
const section = (xml, tag) => xml.match(new RegExp(`<${tag}\\b[\\s\\S]*?<\\/${tag}>`))?.[0] ?? xml.match(new RegExp(`<${tag}\\b[^>]*/>`))?.[0] ?? '';
const structural = {};
for (const tag of ['cols', 'mergeCells', 'dataValidations', 'conditionalFormatting', 'hyperlinks', 'pageMargins', 'pageSetup']) structural[tag] = section(origSheet, tag) === section(outSheet, tag);

const productWb = await SpreadsheetFile.importXlsx(await FileBlob.load(productPath));
const mapWb = await SpreadsheetFile.importXlsx(await FileBlob.load(mappingPath));
const productRows = productWb.worksheets.getItem('项目').getRange('A5:AV10').values;
const mapRows = mapWb.worksheets.getItem('图片URL映射').getRange('A2:F66').values;
const formulaErrors = [];
for (const ws of [...productWb.worksheets.items, ...mapWb.worksheets.items]) {
  const used = ws.getUsedRange(true);
  if (!used) continue;
  used.values.forEach((row, ri) => row.forEach((v, ci) => {
    if (typeof v === 'string' && /^#(REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|NULL!)/.test(v)) formulaErrors.push(`${ws.name}:${ri + 1},${ci + 1}:${v}`);
  }));
}

const skus = productRows.map(r => r[1]);
const titles = productRows.map(r => r[3]);
const descriptions = productRows.map(r => r[6]);
const photoLists = productRows.map(r => String(r[7] || '').split(';').filter(Boolean));
const allUrls = mapRows.map(r => r[5]).filter(Boolean).map(String);
const badUrls = allUrls.filter(u => /[\s()\u0080-\uFFFF]/.test(u));
const foldersCorrect = productRows.every((r, i) => photoLists[i].every(u => u.includes(`/images/${encodeURIComponent(r[1])}/`)));
const result = {
  exactEntries,
  onlyOriginalEntries,
  onlyOutputEntries,
  equality,
  structural,
  sheets: productWb.worksheets.items.map(s => s.name),
  rowCount: productRows.filter(r => r[1]).length,
  skuUnique: new Set(skus).size === skus.length,
  titleUnique: new Set(titles).size === titles.length,
  titleLengths: titles.map(x => x.length),
  descriptionLengths: descriptions.map(x => x.length),
  imageCounts: photoLists.map(x => x.length),
  mappingCount: allUrls.length,
  badUrlCount: badUrls.length,
  foldersCorrect,
  brandBlank: productRows.every(r => r[5] == null),
  videoBlank: productRows.every(r => r[8] == null),
  fixedFields: productRows.every(r => r[0] === 1 && r[4] === '缝纫机' && r[28] === 'Китай'),
  formulaErrors,
};
const allGood = exactEntries && Object.values(equality).every(Boolean) && Object.values(structural).every(Boolean) &&
  result.sheets.join('|') === '项目|使用说明' && result.rowCount === 6 && result.skuUnique && result.titleUnique &&
  result.titleLengths.every(x => x <= 60) && result.descriptionLengths.every(x => x >= 1500 && x <= 2000) &&
  result.imageCounts.reduce((a, b) => a + b, 0) === 65 && result.mappingCount === 65 && result.badUrlCount === 0 &&
  result.foldersCorrect && result.brandBlank && result.videoBlank && result.fixedFields && result.formulaErrors.length === 0;
console.log(JSON.stringify({ allGood, ...result }, null, 2));
if (!allGood) process.exitCode = 1;
