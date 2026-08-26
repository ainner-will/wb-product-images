import fs from 'node:fs/promises';
import path from 'node:path';
import { FileBlob, SpreadsheetFile, Workbook } from '@oai/artifact-tool';

const root = process.cwd();
const outputDir = path.join(root, 'outputs', 'straightener-20260825');
const rawPrefix = 'https://raw.githubusercontent.com/ainner-will/wb-product-images/main/images/';
const imageExtensions = new Set(['.png','.jpg','.jpeg','.webp']);
const videoExtensions = new Set(['.mp4','.mov','.avi','.webm','.mkv']);

const products = [
  ['灰色 / 平板夹+欧规','HS-GY-FLAT',0.5,278.378378378378,35,10,6,'灰色','серый','Фен выпрямитель для волос с плоскими пластинами','плоские пластины',15,200,4,'керамика','сетевой'],
  ['灰色 / 弧形夹板+欧规','HS-GY-CURVE',0.5,278.378378378378,35,10,6,'灰色','серый','Фен выпрямитель для волос с изогнутыми пластинами','изогнутые пластины',15,220,5,'керамика','сетевой'],
  ['灰色 / 玉米夹+欧规','HS-GY-CRIMP',0.5,278.378378378378,35,10,6,'灰色','серый','Фен выпрямитель для волос для гофре','пластины для гофре',15,220,5,'керамика','сетевой'],
  ['调速热蒸汽直发夹板+欧规','HS-STEAM-EU',0.5,294.594594594595,35,10,6,'粉红色','розовый','Паровой фен выпрямитель для волос','паровая рабочая зона',null,200,3,null,'сетевой'],
  ['五齿+欧规','HS-CRIMP-5T',0.6,317.837837837838,35,10,6,'黑色;金色','чёрный и золотистый','Фен выпрямитель для волос гофре с пятью зубцами','пятизубцовая пластина',15,200,4,'керамика','сетевой'],
  ['六齿+欧规','HS-CRIMP-6T',0.6,317.837837837838,35,10,6,'黑色;金色','чёрный и золотистый','Фен выпрямитель для волос гофре с шестью зубцами','шестизубцовая пластина',15,200,4,'керамика','сетевой'],
  ['九齿+欧规','HS-CRIMP-9T',0.6,317.837837837838,35,10,6,'黑色;金色','чёрный и золотистый','Фен выпрямитель для волос гофре с девятью зубцами','девятизубцовая пластина',15,200,4,'керамика','сетевой'],
  ['平板+欧规','HS-CRIMP-FLAT',0.6,317.837837837838,35,10,6,'黑色;金色','чёрный и золотистый','Фен выпрямитель для волос с широкой пластиной','широкая рабочая пластина',15,200,4,'керамика','сетевой'],
  ['湖蓝色平板夹+欧规','HS-LB-FLAT',0.6,301.621621621622,30,4,3,'蓝色','голубой','Голубой фен выпрямитель для волос с плоскими пластинами','плоские пластины',null,null,null,null,'сетевой'],
  ['玫红色平板夹+欧规','HS-RPK-FLAT',0.6,301.621621621622,30,4,3,'粉红色','розовый','Розовый фен выпрямитель для волос с плоскими пластинами','плоские пластины',null,null,null,null,'сетевой'],
  ['桔色平板夹+欧规','HS-OR-FLAT',0.6,301.621621621622,30,4,3,'橙色','оранжевый','Оранжевый фен выпрямитель для волос с плоскими пластинами','плоские пластины',null,null,null,null,'сетевой'],
  ['湖蓝色玉米夹+欧规','HS-LB-CRIMP',0.6,301.621621621622,30,4,3,'蓝色','голубой','Голубой фен выпрямитель для волос для гофре','пластины для гофре',null,null,null,null,'сетевой'],
  ['玫红色玉米夹+欧规','HS-RPK-CRIMP',0.6,301.621621621622,30,4,3,'粉红色','розовый','Розовый фен выпрямитель для волос для гофре','пластины для гофре',null,null,null,null,'сетевой'],
  ['桔色玉米夹+欧规','HS-OR-CRIMP',0.6,301.621621621622,30,4,3,'橙色','оранжевый','Оранжевый фен выпрямитель для волос для гофре','пластины для гофре',null,null,null,null,'сетевой'],
  ['KD382-充电款+欧规','KD382-CORDLESS',0.9,2008.09523809524,32,14,10,'银色;粉红色;黑色','серебристый, розовый и чёрный','Беспроводной фен выпрямитель расчёска для волос','расчёска выпрямитель',15,200,6,null,'аккумулятор'],
  ['KD680欧规-KSKIN','KD680-EU-KSKIN',0.6,615.135135135135,26,10,7,'粉红色','розовый','Паровой фен выпрямитель расчёска для волос','паровая расчёска выпрямитель',null,200,3,null,'сетевой'],
  ['黑蓝色 / 中规+（物流选配欧规插头）','HS-BBK-CN',0.6,1101.62162162162,32,10,7,'黑色;蓝色','чёрный и синий','Чёрно синий фен выпрямитель расчёска для волос','расчёска выпрямитель',30,200,5,null,'сетевой'],
  ['白色 / 中规+（物流选配欧规插头）','HS-WH-CN',0.6,1101.62162162162,32,10,7,'白色','белый','Белый фен выпрямитель расчёска для волос','расчёска выпрямитель',30,200,5,null,'сетевой'],
  ['KD880 粉色 / 中规国内插头+（物流选配欧规插头）','KD880-PK-CN',0.6,1507.02702702703,32,10,7,'粉红色','розовый','Розовый фен выпрямитель расчёска для волос','расчёска выпрямитель',null,200,5,null,'сетевой'],
].map(([sourceName, sku, weight, price, length, width, height, color, colorRu, title, type, heatSeconds, maxTemp, modes, material, power]) => ({sourceName, sku, weight, price, length, width, height, color, colorRu, title, type, heatSeconds, maxTemp, modes, material, power}));

function compactDescription(p, index, hasImages) {
  const temperature = p.maxTemp ? `Максимальная температура, указанная для этой версии, составляет ${p.maxTemp} градусов.` : 'Температуру следует выбирать постепенно, начиная с комфортного уровня для своего типа волос.';
  const modes = p.modes ? `Предусмотрено ${p.modes} режимов работы, поэтому настройку удобно подобрать для повседневной укладки.` : 'Перед первой укладкой рекомендуется проверить результат на небольшой пряди.';
  const heat = p.heatSeconds ? `Быстрый нагрев занимает около ${p.heatSeconds} секунд, что помогает начать укладку без долгого ожидания.` : 'Устройство рассчитано на аккуратную последовательную укладку без лишней спешки.';
  const visual = hasImages ? `По материалам для данной версии подтверждены ${p.type}, цвет ${p.colorRu} и конструкция, подходящая для домашней укладки.` : `Для этой версии подтверждены ${p.type}, цвет ${p.colorRu} и сетевое подключение с европейской вилкой.`;
  const power = p.power === 'аккумулятор' ? 'Беспроводной формат удобен дома, в поездке и там, где важно свободно двигаться во время укладки.' : 'Сетевое питание подходит для стабильного использования дома и в рабочей зоне у зеркала.';
  const variants = [
    'Рабочую часть ведут по сухим и предварительно расчёсанным прядям плавным движением, не задерживаясь на одном участке.',
    'Для более естественного результата удобно делить волосы на небольшие пряди и подбирать скорость движения по их плотности.',
    'При создании объёма у корней или выраженной текстуры лучше начинать с нижних прядей и контролировать форму поэтапно.',
    'После укладки прибору дают остыть на термостойкой поверхности, затем очищают рабочую зону только в отключённом состоянии.',
  ];
  const care = [
    'Не используйте прибор рядом с водой и не касайтесь нагретой зоны руками. Перед хранением дождитесь полного остывания.',
    'Для сохранения аккуратного вида держите корпус сухим, не наматывайте кабель вокруг горячего прибора и храните его отдельно.',
    'Настройку температуры выбирают с учётом состояния волос. Для осветлённых, тонких или повреждённых прядей подходит более бережный режим.',
    'Результат зависит от подготовки прядей, поэтому перед укладкой волосы должны быть чистыми, сухими и свободно расчесанными.',
  ];
  const paragraphs = [
    `${p.title} предназначен для аккуратной укладки волос дома. Эта версия использует ${p.type}. ${visual} Прибор помогает оформить гладкую причёску или выразительный рельеф в зависимости от формы рабочей части.`,
    `${heat} ${modes} ${temperature} Такое сочетание удобно для ежедневного ухода, когда важно работать с отдельными прядями и контролировать итоговый вид причёски.`,
    `${power} ${variants[index % variants.length]} Не давите на корпус и не соединяйте рабочие части на пряди слишком резко, чтобы сохранить естественное движение волос.`,
    `Перед включением проверьте целостность кабеля и корпуса. Используйте прибор в сухом помещении и подключайте к подходящей розетке. ${care[index % care.length]} После завершения отключите прибор от питания.`,
    `Модель подойдёт для самостоятельной укладки перед работой, учёбой, встречей или поездкой. Компактный формат упаковки облегчает хранение. Точный эффект зависит от длины, густоты и исходной текстуры волос, поэтому режим и скорость движения выбирают индивидуально.`,
  ];
  let text = paragraphs.join('\n\n');
  const filler = 'Плавная техника укладки и разумный выбор температуры помогают поддерживать аккуратный вид волос без излишнего воздействия.';
  while (text.length < 1650) text += `\n\n${filler}`;
  return text.slice(0, 1950).trim();
}

function encodedUrl(folder, file) { return `${rawPrefix}${encodeURIComponent(folder)}/${encodeURIComponent(file)}`; }
function ceil(n) { return Math.ceil(n); }

await fs.mkdir(outputDir, {recursive:true});
const dirEntries = await fs.readdir(root, {withFileTypes:true});
const folderImages = new Map();
const mapRows = [];
let videoFolderCount = 0;
let videoUrlCount = 0;
for (const entry of dirEntries.filter(x => x.isDirectory())) {
  if (entry.name === 'node_modules' || entry.name === 'outputs' || entry.name === 'straightener_preview') continue;
  const folder = path.join(root, entry.name);
  const names = (await fs.readdir(folder)).sort((a,b)=>a.localeCompare(b,'en'));
  const images = names.filter(n => imageExtensions.has(path.extname(n).toLowerCase()));
  const videos = names.filter(n => videoExtensions.has(path.extname(n).toLowerCase()));
  if (videos.length) { videoFolderCount++; videoUrlCount += videos.length; }
  folderImages.set(entry.name, images);
  images.forEach((file, i) => mapRows.push([entry.name, file, path.join(folder,file), encodedUrl(entry.name,file), i+1, '图片', entry.name]));
  videos.forEach((file, i) => mapRows.push([entry.name, file, path.join(folder,file), encodedUrl(entry.name,file), i+1, '视频', entry.name]));
}

for (let i=0;i<products.length;i++) {
  const p = products[i];
  p.images = folderImages.get(p.sku) || [];
  p.description = compactDescription(p, i, p.images.length > 0);
  if (p.title.length > 60) throw new Error(`Title too long: ${p.title}`);
  if (p.description.length > 2000) throw new Error(`Description too long: ${p.sku}`);
}

const uploadWb = await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(root,'直发器.xlsx')));
const uploadSheet = uploadWb.worksheets.getItem('项目');
const values = products.map(p => {
  const row = Array(48).fill(null);
  row[0] = 1;
  row[1] = p.sku;
  row[3] = p.title;
  row[4] = '直发器';
  row[6] = p.description;
  row[7] = p.images.map(file => encodedUrl(p.sku,file)).join(';') || null;
  row[10] = p.weight;
  row[13] = 1;
  row[14] = 1;
  row[16] = p.color;
  row[18] = ceil(p.price);
  row[25] = p.heatSeconds;
  row[30] = p.maxTemp;
  row[31] = 'Китай';
  row[32] = p.maxTemp ? 'регулировка температуры' : null;
  row[34] = ceil(p.width);
  row[35] = ceil(p.length);
  row[36] = ceil(p.height);
  row[37] = p.weight * 1000;
  row[41] = p.material;
  row[44] = 'профессиональный';
  row[46] = p.power;
  row[47] = p.modes;
  return row;
});
uploadSheet.getRange(`A5:AV${4+products.length}`).values = values;
const uploadXlsx = await SpreadsheetFile.exportXlsx(uploadWb);
await uploadXlsx.save(path.join(outputDir,'产品上传表格.xlsx'));

const mapWb = Workbook.create();
const mapSheet = mapWb.worksheets.add('图片URL映射表');
const headers = ['SKU文件夹名','图片文件名','本地路径','图片URL','图片排序号','文件类型','适用SKU范围'];
mapSheet.getRange(`A1:G${mapRows.length+1}`).values = [headers, ...mapRows];
mapSheet.getRange('A1:G1').format = {fill:'#5B21B6',font:{bold:true,color:'#FFFFFF'},horizontalAlignment:'center'};
mapSheet.getRange(`A2:G${mapRows.length+1}`).format.wrapText = false;
mapSheet.getRange('A:A').format.columnWidth = 22;
mapSheet.getRange('B:B').format.columnWidth = 42;
mapSheet.getRange('C:C').format.columnWidth = 58;
mapSheet.getRange('D:D').format.columnWidth = 100;
mapSheet.getRange('E:E').format.columnWidth = 12;
mapSheet.getRange('F:G').format.columnWidth = 18;
mapSheet.freezePanes.freezeRows(1);
const mapXlsx = await SpreadsheetFile.exportXlsx(mapWb);
await mapXlsx.save(path.join(outputDir,'图片URL映射表.xlsx'));

await fs.writeFile(path.join(outputDir,'straightener_rows.json'), JSON.stringify(products.map(p=>({sku:p.sku,weight:p.weight,price:ceil(p.price),length:ceil(p.length),width:ceil(p.width),height:ceil(p.height),images:p.images.length,title:p.title,descriptionLength:p.description.length})),null,2));
await fs.writeFile(path.join(outputDir,'straightener_metadata.json'), JSON.stringify({sourceProducts:products.length,skuFolders:folderImages.size,imageUrls:mapRows.filter(r=>r[5]==='图片').length,videoFolderCount,videoUrlCount,unmatched:products.filter(p=>!p.images.length).map(p=>p.sku)},null,2));
console.log(JSON.stringify({products:products.length,skuFolders:folderImages.size,imageUrls:mapRows.filter(r=>r[5]==='图片').length,videoFolderCount,videoUrlCount,unmatched:products.filter(p=>!p.images.length).map(p=>p.sku),descriptionLengths:products.map(p=>[p.sku,p.description.length])},null,2));
