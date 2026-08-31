import fs from 'node:fs/promises';
import path from 'node:path';
import JSZip from 'jszip';
import { FileBlob, SpreadsheetFile, Workbook } from '@oai/artifact-tool';

const root = process.cwd();
const outDir = path.join(root, 'outputs', '01a04c8f-2214-74f1-959d-5098f7c43bb6-sewing');
await fs.mkdir(outDir, { recursive: true });
const templatePath = path.join(root, '缝纫机.xlsx');
const colorPath = path.join(root, 'WB产品颜色-中文版.xlsx');
const csvPath = path.join(root, 'ku_list.csv');
const interimPath = path.join(outDir, '.artifact-intermediate.xlsx');
const productPath = path.join(outDir, '产品上传表格.xlsx');
const mappingPath = path.join(outDir, '图片URL映射表.xlsx');

const csvBytes = await fs.readFile(csvPath);
const csvText = new TextDecoder('gb18030').decode(csvBytes);
const csvWb = await Workbook.fromCSV(csvText, { sheetName: 'ku_list' });
const csvRows = csvWb.worksheets.getItem('ku_list').getUsedRange(true).values;
const headers = csvRows[0];
const idx = Object.fromEntries(headers.map((h, i) => [String(h).trim(), i]));
const parsed = csvRows.slice(1).filter(r => r[idx['建议英文SKU']]).map(r => ({
  sku: String(r[idx['建议英文SKU']]).trim().toUpperCase(),
  name: String(r[idx['SKU名称']]).trim(),
  id: String(r[idx['SKU ID']]).trim(),
  length: Number(String(r[idx['长(cm)']]).trim()),
  width: Number(String(r[idx['宽(cm)']]).trim()),
  height: Number(String(r[idx['高(cm)']]).trim()),
  weightKg: Number(String(r[idx['重量(g)']]).replaceAll(',', '').trim()) / 1000,
  packageWeight: Number(String(r[idx['快递包装重量']]).trim()),
  salePrice: Math.ceil(Number(String(r[idx['跨境售价']]).trim())),
  meaning: String(r[idx['编码含义']]).trim(),
}));

const colorWb = await SpreadsheetFile.importXlsx(await FileBlob.load(colorPath));
const allowedColors = new Set(colorWb.worksheets.getItemAt(0).getUsedRange(true).values.flat().filter(Boolean).map(String));
if (!allowedColors.has('橙色')) throw new Error('颜色表中未找到“橙色”');

const collator = new Intl.Collator('zh-CN', { numeric: true, sensitivity: 'base' });
const imageExt = new Set(['.png', '.jpg', '.jpeg', '.webp']);
const videoExt = new Set(['.mp4', '.mov', '.avi', '.webm', '.mkv']);
const base = 'https://raw.githubusercontent.com/ainner-will/wb-product-images/main/images';
const strictEncode = s => encodeURIComponent(s).replace(/[!'()*]/g, ch => `%${ch.charCodeAt(0).toString(16).toUpperCase()}`);
const encPath = parts => parts.map(strictEncode).join('/');
const assetsBySku = {};
const mappingRows = [];
for (const item of parsed) {
  const folder = path.join(root, item.sku);
  const names = (await fs.readdir(folder, { withFileTypes: true }))
    .filter(x => x.isFile()).map(x => x.name).sort(collator.compare);
  const images = names.filter(n => imageExt.has(path.extname(n).toLowerCase()));
  const videos = names.filter(n => videoExt.has(path.extname(n).toLowerCase()));
  const imageUrls = images.map(n => `${base}/${encPath([item.sku, n])}`);
  const videoUrls = videos.map(n => `${base}/${encPath([item.sku, n])}`);
  assetsBySku[item.sku] = { images, videos, imageUrls, videoUrls };
  for (let i = 0; i < images.length; i++) mappingRows.push([item.sku, item.sku, images[i], '图片', i + 1, imageUrls[i]]);
  for (let i = 0; i < videos.length; i++) mappingRows.push([item.sku, item.sku, videos[i], '视频', i + 1, videoUrls[i]]);
}

const commonCorded = `Это переносная мешкозашивочная машина для закрытия заполненных мешков на складе, в цехе, сельском хозяйстве или зоне отгрузки. Она подходит для тканых полипропиленовых, джутовых, бумажных мешков, нетканого полотна, геотекстиля и упаковочных материалов. Встроенная обрезка нити помогает быстро завершить шов, а регулируемый натяжитель упрощает настройку под материал. Корпус выполнен из алюминиевого сплава.\n\nЭта версия работает от сети 220 В через кабель длиной 2,5 м. В отличие от аккумуляторной машины, сетевой вариант не требует зарядки батареи и удобен для продолжительной работы на постоянном посту. При этом рядом должна быть исправная розетка, а кабель необходимо размещать так, чтобы он не попадал под движущиеся части и не мешал оператору. Аккумуляторная версия свободнее перемещается между мешками, но ее время работы зависит от заряда и емкости батареи.\n\nПо данным материалов, допустимая толщина прошивки составляет 0,2–11 мм, шаг стежка — 7–10 мм. Заправьте подходящую нить, проверьте натяжение на образце и ведите машину без рывков. После окончания используйте встроенный механизм обрезки.\n\nВ комплект входят кабель, пять игл, масленка, приводной ремень, ключ, шестигранники, запасной нож и инструкция. Перед обслуживанием, заменой иглы, ножа или ремня отключайте питание. Не касайтесь иглы и режущего узла во время работы.`;

const commonBattery = (cap, choice) => `Эта аккумуляторная мешкозашивочная машина предназначена для закрытия мешков на складе, в логистике, сельском хозяйстве и производстве. Она работает с тканым полипропиленом, джутом, бумагой, неткаными и другими упаковочными материалами. Одноигольная однониточная схема образует цепной шов, а автоматическая обрезка нити помогает аккуратно завершать работу. Регулятор натяжения позволяет подстроить подачу нити.\n\nГлавное отличие этой версии — литий-ионный аккумулятор 36 В емкостью ${cap} мАч. ${choice} По сравнению с сетевой машиной аккумуляторный вариант не привязан к розетке и кабелю, поэтому удобен при переходе между палетами, транспортом и удаленными рабочими точками. Сетевая версия лучше подходит для длительной работы на одном месте без пауз на зарядку, а аккумуляторная дает мобильность, но ее фактическое время работы зависит от нагрузки, материала, состояния батареи и температуры. Заряжайте аккумулятор штатным устройством.\n\nПо изображениям указаны мощность двигателя 210 Вт, скорость двигателя до 12000 об/мин и скорость шитья до 1700 стежков в минуту. Перед основной работой выполните пробный шов, отрегулируйте натяжение и проверьте установку иглы. Ведите машину плавно, не закрывайте вентиляционные отверстия и держите руки вне рабочей зоны.\n\nКомплект включает аккумулятор, зарядное устройство, две катушки нити, пять игл, масленку, приводной ремень, ключи и запасной нож. Перед заменой иглы, очисткой или обслуживанием снимайте аккумулятор. Не работайте во влажной среде и не используйте поврежденную батарею. Эта комплектация рассчитана на одну иглу и одну нить; от сетевых вариантов с двумя нитями или двумя иглами она отличается мобильным питанием и однониточной схемой.`;

const spec = {
  '370-1N1T': {
    title: 'Мешкозашивочная машина сетевая, 1 игла, 1 нить',
    desc: `${commonCorded}\n\nКонкретная комплектация использует одну иглу и одну нить. Это наиболее простая схема заправки среди трех сетевых вариантов: она формирует одну линию цепного стежка и требует одной рабочей нити. От версии с одной иглой и двумя нитями отличается количеством нитей и способом формирования шва, а от двухигольной — отсутствием второй иглы и второй параллельной линии. Такой вариант удобен, когда приоритетом являются простая подготовка и экономичный расход нити.`,
    stitchLength: 10, speed: null, rows: 1,
  },
  '370-1N2T': {
    title: 'Мешкозашивочная машина сетевая, 1 игла, 2 нити',
    desc: `${commonCorded}\n\nЭта комплектация рассчитана на одну иглу и две нити. В отличие от одноигольной однониточной версии в формировании шва участвуют две нити, поэтому заправка и контроль натяжения выполняются для обеих нитей. При этом линия прошивки остается одной; это отличает вариант от двухигольной двухниточной машины, которая выполняет две линии. Выбирайте эту конфигурацию, если нужна одноигольная схема с двухниточным формированием шва и вы готовы учитывать больший расход нити и более внимательную настройку натяжения.`,
    stitchLength: 10, speed: null, rows: 1,
  },
  '370-2N2T': {
    title: 'Мешкозашивочная машина сетевая, 2 иглы, 2 нити',
    desc: `${commonCorded}\n\nЭта комплектация использует две иглы и две нити и предназначена для выполнения двух линий прошивки за один проход. В этом ее основное отличие от одноигольных вариантов: у базовой версии одна игла и одна нить, у промежуточной — одна игла и две нити, а здесь задействованы два игольных канала. Перед работой особенно важно одинаково проверить положение обеих игл и натяжение обеих нитей на пробном образце. Конфигурация подходит, когда по технологической задаче требуется двухрядное закрытие мешка.`,
    stitchLength: 10, speed: null, rows: 2,
  },
  'GK9-BAT-1500MAH': {
    title: 'Мешкозашивочная машина аккумуляторная 1500 мАч',
    desc: commonBattery(1500, 'Это самая небольшая емкость среди трех аккумуляторных комплектаций; она подходит, когда нужен базовый запас энергии и предусмотрена регулярная подзарядка.'),
    stitchLength: null, speed: 1700, rows: 1,
  },
  'GK9-BAT-2000MAH': {
    title: 'Мешкозашивочная машина аккумуляторная 2000 мАч',
    desc: commonBattery(2000, 'Это средняя емкость в линейке: она дает больший запас энергии, чем версия 1500 мАч, но меньше, чем вариант 2500 мАч, при одинаковой одноигольной однониточной схеме.'),
    stitchLength: null, speed: 1700, rows: 1,
  },
  'GK9-BAT-2500MAH': {
    title: 'Мешкозашивочная машина аккумуляторная 2500 мАч',
    desc: commonBattery(2500, 'Это максимальная емкость среди представленных аккумуляторных вариантов; при сопоставимой нагрузке она обеспечивает больший запас энергии, чем комплектации 1500 и 2000 мАч, без изменения одноигольной однониточной схемы.'),
    stitchLength: null, speed: 1700, rows: 1,
  },
};

const cordedKit = 'машина;кабель 2,5 м;5 игл;масленка;ремень;ключ;шестигранники;нож;инструкция';
const batteryKit = 'машина;аккумулятор 36 В;зарядка;2 катушки;5 игл;масленка;ремень;ключи;нож';
const rows = parsed.map(item => {
  const s = spec[item.sku];
  const corded = item.sku.startsWith('370-');
  return [
    1, item.sku, null, s.title, '缝纫机', null, s.desc,
    assetsBySku[item.sku].imageUrls.join(';'), assetsBySku[item.sku].videoUrls.join(';') || null,
    '不需要', item.packageWeight, null, null, 1, 1, null, '橙色', null, item.salePrice,
    null, null, null, null, null, null, null, null, s.stitchLength, 'Китай', null, null,
    'автоматическая обрезка нити', s.speed, s.rows, corded ? cordedKit : batteryKit,
    item.weightKg, Math.ceil(item.width), item.packageWeight, Math.ceil(item.length), Math.ceil(item.height),
    null, null, null, null, corded ? 'алюминиевый сплав' : null, null, 'ручная', 'кнопочное'
  ];
});

if (rows.some(r => r.length !== 48)) throw new Error('Количество столбцов строки не равно 48');
for (const item of parsed) {
  const s = spec[item.sku];
  if (s.title.length > 60) throw new Error(`Длинный заголовок ${item.sku}: ${s.title.length}`);
  if (s.desc.length > 2000 || s.desc.length < 1500) throw new Error(`Описание ${item.sku}: ${s.desc.length}`);
  if (!/^[A-Z0-9-]+$/.test(item.sku)) throw new Error(`Некорректный SKU: ${item.sku}`);
  if (assetsBySku[item.sku].images.length === 0) throw new Error(`Нет изображений: ${item.sku}`);
}
for (const row of rows) for (let c = 0; c < row.length; c++) {
  const v = row[c];
  if (typeof v === 'string' && c !== 6 && c !== 7 && c !== 8 && v.length > 100) throw new Error(`Ячейка >100: col ${c + 1}`);
}

// Author the upload workbook with artifact-tool first.
const uploadWb = await SpreadsheetFile.importXlsx(await FileBlob.load(templatePath));
const project = uploadWb.worksheets.getItem('项目');
project.getRange('A5:AV10').values = rows;
const uploadBlob = await SpreadsheetFile.exportXlsx(uploadWb);
await uploadBlob.save(interimPath);

// Create the URL mapping workbook with artifact-tool.
const mapWb = Workbook.create();
const mapSheet = mapWb.worksheets.add('图片URL映射');
mapSheet.showGridLines = false;
mapSheet.getRange('A1:F1').values = [['建议英文SKU', '素材文件夹', '文件名', '类型', '排序', 'Raw URL']];
mapSheet.getRange(`A2:F${mappingRows.length + 1}`).values = mappingRows;
mapSheet.getRange('A1:F1').format = { fill: '#6D28D9', font: { bold: true, color: '#FFFFFF' }, horizontalAlignment: 'center', verticalAlignment: 'center' };
mapSheet.getRange(`A2:F${mappingRows.length + 1}`).format = { verticalAlignment: 'top' };
mapSheet.getRange(`E2:E${mappingRows.length + 1}`).format.numberFormat = '0';
mapSheet.getRange(`A1:F${mappingRows.length + 1}`).format.borders = { preset: 'insideHorizontal', style: 'thin', color: '#E5E7EB' };
mapSheet.getRange('A:A').format.columnWidth = 24;
mapSheet.getRange('B:B').format.columnWidth = 24;
mapSheet.getRange('C:C').format.columnWidth = 58;
mapSheet.getRange('D:E').format.columnWidth = 12;
mapSheet.getRange('F:F').format.columnWidth = 95;
mapSheet.getRange('A1:F1').format.rowHeight = 26;
mapSheet.freezePanes.freezeRows(1);
mapSheet.tables.add(`A1:F${mappingRows.length + 1}`, true, 'ImageUrlMap');
const mapBlob = await SpreadsheetFile.exportXlsx(mapWb);
await mapBlob.save(mappingPath);

function colName(index) {
  let n = index + 1, s = '';
  while (n) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
function esc(s) { return String(s).replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;'); }
function cellXml(ref, value, old = '') {
  const open = old.match(/^<c\b([^>]*)/s)?.[1] ?? ` r="${ref}"`;
  const attrs = open.replace(/\s+r="[^"]*"/g, '').replace(/\s+t="[^"]*"/g, '');
  if (typeof value === 'number') return `<c r="${ref}"${attrs}><v>${value}</v></c>`;
  return `<c r="${ref}"${attrs} t="inlineStr"><is><t xml:space="preserve">${esc(value)}</t></is></c>`;
}
function injectRows(xml, dataRows) {
  for (let ri = 0; ri < dataRows.length; ri++) {
    const rowNum = ri + 5;
    const re = new RegExp(`<row\\b([^>]*)\\br="${rowNum}"([^>]*)>([\\s\\S]*?)<\\/row>`);
    const m = xml.match(re);
    let attrs = m ? `${m[1]}r="${rowNum}"${m[2]}` : ` r="${rowNum}"`;
    let body = m ? m[3] : '';
    const cells = new Map();
    for (const cm of body.matchAll(/<c\b[^>]*\br="([A-Z]+\d+)"[^>]*(?:\/>|>[\s\S]*?<\/c>)/g)) cells.set(cm[1], cm[0]);
    for (let ci = 0; ci < dataRows[ri].length; ci++) {
      const value = dataRows[ri][ci];
      if (value === null || value === '') continue;
      const ref = `${colName(ci)}${rowNum}`;
      cells.set(ref, cellXml(ref, value, cells.get(ref) || ''));
    }
    const sorted = [...cells.entries()].sort((a, b) => {
      const ca = a[0].match(/[A-Z]+/)[0], cb = b[0].match(/[A-Z]+/)[0];
      const toNum = c => [...c].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
      return toNum(ca) - toNum(cb);
    }).map(x => x[1]).join('');
    const newRow = `<row${attrs}>${sorted}</row>`;
    xml = m ? xml.replace(re, newRow) : xml.replace('</sheetData>', `${newRow}</sheetData>`);
  }
  return xml;
}

// Restore the exact original template package and inject only populated cells.
const originalZip = await JSZip.loadAsync(await fs.readFile(templatePath));
const workbookXml = await originalZip.file('xl/workbook.xml').async('string');
const relsXml = await originalZip.file('xl/_rels/workbook.xml.rels').async('string');
const projectRel = workbookXml.match(/<sheet\b[^>]*name="项目"[^>]*r:id="([^"]+)"/)[1];
const target = relsXml.match(new RegExp(`<Relationship\\b[^>]*Id="${projectRel}"[^>]*Target="([^"]+)"`))[1];
const sheetPath = target.startsWith('/') ? target.slice(1) : `xl/${target.replace(/^\.\//, '')}`;
const originalSheet = await originalZip.file(sheetPath).async('string');
originalZip.file(sheetPath, injectRows(originalSheet, rows), { createFolders: false });
await fs.writeFile(productPath, await originalZip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
await fs.rm(interimPath);

// Re-open, inspect, and render final workbooks with artifact-tool.
const finalUpload = await SpreadsheetFile.importXlsx(await FileBlob.load(productPath));
const finalMap = await SpreadsheetFile.importXlsx(await FileBlob.load(mappingPath));
const inspectUpload = await finalUpload.inspect({ kind: 'table,formula', sheetId: '项目', range: 'A1:AV10', maxChars: 25000, tableMaxRows: 12, tableMaxCols: 48, tableMaxCellChars: 220 });
const inspectMap = await finalMap.inspect({ kind: 'table,formula', sheetId: '图片URL映射', range: `A1:F${mappingRows.length + 1}`, maxChars: 10000, tableMaxRows: 8, tableMaxCols: 6, tableMaxCellChars: 160 });
await fs.writeFile(path.join(outDir, 'inspect-upload.ndjson'), inspectUpload.ndjson);
await fs.writeFile(path.join(outDir, 'inspect-mapping.ndjson'), inspectMap.ndjson);

const renderSpecs = [
  ['项目', 'A1:H10', 'upload-01.png'], ['项目', 'I1:P10', 'upload-02.png'],
  ['项目', 'Q1:X10', 'upload-03.png'], ['项目', 'Y1:AF10', 'upload-04.png'],
  ['项目', 'AG1:AN10', 'upload-05.png'], ['项目', 'AO1:AV10', 'upload-06.png'],
  ['使用说明', 'A1:D20', 'upload-instructions.png'],
];
for (const [sheetName, range, file] of renderSpecs) {
  const img = await finalUpload.render({ sheetName, range, scale: 1, format: 'png' });
  await fs.writeFile(path.join(outDir, file), new Uint8Array(await img.arrayBuffer()));
}
for (const [range, file] of [['A1:F35', 'mapping-01.png'], [`A36:F${mappingRows.length + 1}`, 'mapping-02.png']]) {
  const img = await finalMap.render({ sheetName: '图片URL映射', range, scale: 0.9, format: 'png' });
  await fs.writeFile(path.join(outDir, file), new Uint8Array(await img.arrayBuffer()));
}

const descLens = parsed.map(x => ({ sku: x.sku, title: spec[x.sku].title.length, description: spec[x.sku].desc.length, images: assetsBySku[x.sku].images.length, videos: assetsBySku[x.sku].videos.length }));
console.log(JSON.stringify({ productPath, mappingPath, skuCount: parsed.length, mappingRows: mappingRows.length, descLens }, null, 2));
