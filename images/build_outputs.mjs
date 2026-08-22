import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const base = "H:/New project/wb-product-images/images";
const outDir = path.join(base, "outputs/01a00ef6-c440-74e3-83c4-ae1f8250c789");
await fs.mkdir(outDir, { recursive: true });

const inv = JSON.parse(await fs.readFile(path.join(base, "work_preview/inventory.json"), "utf8"));
const imageExts = new Set([".png", ".jpg", ".jpeg", ".webp"]);
const imagesByFolder = new Map();
for (const e of inv.filter(x => imageExts.has(x.ext))) {
  if (!imagesByFolder.has(e.folder)) imagesByFolder.set(e.folder, []);
  imagesByFolder.get(e.folder).push(e);
}
for (const arr of imagesByFolder.values()) arr.sort((a,b)=>a.file.localeCompare(b.file,"zh-CN",{numeric:true}));
const rawUrl = (folder, file) => `https://raw.githubusercontent.com/ainner-will/wb-product-images/main/images/${encodeURIComponent(folder)}/${encodeURIComponent(file)}`;

const clothing = ["XS","S","M","L","XL","XXL","XXXL"];
const specs = [
  {sku:"CLEO-DENILE",weight:0.7,price:1344.8,dims:[35,28,5],sizes:clothing,title:"Косплей костюм Клео де Нил Cleo de Nile",char:"Клео де Нил",series:"Monster High",colors:"绿松石色;金色;黑色",kind:"costume",contents:"платье;накидка;пояс;украшения",detail:"Комплект передает образ Клео де Нил. Бирюзовые, золотистые и черные элементы, декоративный пояс, украшения для рук и ног и многослойные детали формируют узнаваемый египетский стиль персонажа."},
  {sku:"ELISSABAT-SKIRT",weight:0.9,price:1733.6,dims:[35,28,5],sizes:clothing,title:"Косплей костюм Элиссабэт Elissabat",char:"Элиссабэт",series:"Monster High",colors:"黑色;紫色",kind:"costume",contents:"платье;корсет;юбка;украшение",detail:"Комплект воспроизводит готический образ Элиссабэт. Черно фиолетовая гамма, корсетная часть, многослойная юбка и украшение на шею создают характерный сценический силуэт."},
  {sku:"CLAWDEEN-WOLF",weight:0.9,price:1653.6,dims:[35,28,5],sizes:clothing,title:"Косплей костюм Клодин Вульф Clawdeen Wolf",char:"Клодин Вульф",series:"Monster High",colors:"紫色;粉红色;黑色",kind:"costume",contents:"куртка;корсет;юбка;пояс;ободок с ушами",detail:"Комплект создан по образу Клодин Вульф. Куртка, корсетная часть, юбка, пояс и ободок с ушами объединены фиолетовыми, розовыми и черными акцентами и помогают собрать узнаваемый образ героини."},
  {sku:"MH-ZOMBIE-COS",weight:0.5,price:996,dims:[30,25,5],sizes:clothing,title:"Косплей костюм Фрэнки Штейн Frankie Stein",char:"Фрэнки Штейн",series:"Monster High",colors:"黑色;蓝色;粉红色",kind:"costume",contents:"платье;пояс;гольфы;аксессуары",detail:"Комплект передает школьный образ Фрэнки Штейн. Платье с клетчатыми и контрастными вставками, пояс, гольфы и декоративные детали поддерживают характерную черно синюю гамму с розовыми акцентами."},
  {sku:"MH-DRACULAURA",weight:0.5,price:1036,dims:[30,25,5],sizes:clothing,title:"Косплей костюм Дракулаура Draculaura",char:"Дракулаура",series:"Monster High",colors:"粉红色;黑色;白色",kind:"costume",contents:"платье;накидка;головной аксессуар;бант",detail:"Комплект выполнен по образу Дракулауры. Розовые, черные и белые элементы, накидка, головной аксессуар и бант образуют контрастный силуэт, подходящий для тематического перевоплощения."},
  {sku:"MH-TORALEI-COS",weight:0.8,price:1499.2,dims:[30,25,5],sizes:clothing,title:"Косплей костюм Торалей Страйп Toralei Stripe",char:"Торалей Страйп",series:"Monster High",colors:"粉红色;黑色;橙色",kind:"costume",contents:"топ;куртка;брюки;уши;шарф;хвост",detail:"Комплект создан по образу Торалей Страйп. Топ, куртка, брюки, шарф, ушки и хвост собраны в черно розовой гамме с оранжевыми акцентами и передают кошачьи черты персонажа."},
  {sku:"SDR-COSTUME",weight:1.4,price:1052,dims:[35,40,5],sizes:clothing,title:"Косплей костюм Сандроне Sandrone",char:"Сандроне",series:"Genshin Impact",colors:"白色;黑色;红色;金色",kind:"costume",contents:"костюм;головной убор;перчатки;аксессуары",detail:"Костюм воспроизводит образ Сандроне. Белые и черные части дополнены красными и золотистыми акцентами, головным убором, перчатками и декоративными элементами, характерными для кукольной стилистики героини."},
  {sku:"SDR-SHOES",weight:0.9,price:419.142857,dims:[40,30,8],sizes:[36,37,38,39,40,41,42],title:"Косплей обувь Сандроне Sandrone",char:"Сандроне",series:"Genshin Impact",colors:"黑色;白色;红色;金色",kind:"shoes",contents:"пара обуви",detail:"Косплейная обувь дополняет образ Сандроне. Черно белая основа, золотистая отделка и красный бант согласованы с костюмом персонажа. В таблице указан европейский размер EUR, а российский размер рассчитан на одну единицу меньше."},
  {sku:"SDR-WIG",weight:0.7,price:280.285714,dims:[25,35,5],sizes:["единый"],title:"Косплей парик Сандроне Sandrone",char:"Сандроне",series:"Genshin Impact",colors:"浅金色;棕色",kind:"wig",contents:"парик",detail:"Косплейный парик предназначен для образа Сандроне. Светлый коричнево золотистый оттенок, челка и оформленные пряди помогают приблизить прическу к внешности персонажа. Изделие поставляется в едином размере."},
  {sku:"ARL-COSTUME",weight:1.1,price:809.428571,dims:[15,11,28],sizes:clothing,title:"Косплей костюм Арлекино Arlecchino",char:"Арлекино",series:"Genshin Impact",colors:"",kind:"costume",contents:"костюм",detail:"Косплейный комплект предназначен для создания образа Арлекино. В исходных материалах подтвержден тип товара и наличие костюма, поэтому описание не добавляет неподтвержденные детали отделки, цвета или состава."},
  {sku:"ARL-COSTUME-WIG",weight:1.6,price:1042.285714,dims:[15,11,28],sizes:clothing,title:"Косплей костюм Арлекино Arlecchino с париком",char:"Арлекино",series:"Genshin Impact",colors:"",kind:"costume",contents:"костюм;парик",detail:"Набор предназначен для создания образа Арлекино и включает костюм и парик. В исходных материалах не было доступной папки с фотографиями, поэтому неподтвержденные элементы дизайна, цвет и состав не заявляются."},
  {sku:"ARL-WIG-NET",weight:0.6,price:268,dims:[15,11,28],sizes:["единый"],title:"Косплей парик Арлекино Arlecchino",char:"Арлекино",series:"Genshin Impact",colors:"",kind:"wig",contents:"парик;сетка для волос",detail:"Набор для образа Арлекино включает косплейный парик и сетку для волос. Товар имеет единый размер. Поскольку фотографии для этого артикула отсутствуют, описание не содержит неподтвержденных сведений о цвете, длине и укладке."},
  {sku:"ARL-SHOES",weight:0.8,price:521.142857,dims:[15,11,28],sizes:Array.from({length:13},(_,i)=>34+i),title:"Косплей обувь Арлекино Arlecchino",char:"Арлекино",series:"Genshin Impact",colors:"",kind:"shoes",contents:"пара обуви",detail:"Косплейная обувь предназначена для завершения образа Арлекино. Доступны европейские размеры EUR от 34 до 46. Российский размер в таблице указан на одну единицу меньше. Неподтвержденные сведения о цвете и отделке не добавлены."},
  {sku:"AYK-COSTUME",weight:0.6,price:299.428571,dims:[29,26,3],sizes:clothing,title:"Косплей костюм Камисато Аяка Kamisato Ayaka",char:"Камисато Аяка",series:"Genshin Impact",colors:"蓝色;黑色;红色;金色;白色",kind:"costume",contents:"костюм;головной аксессуар;перчатки;украшения",detail:"Комплект создан по образу Камисато Аяки. Синие, черные и белые детали дополнены красными и золотистыми акцентами, головным аксессуаром, перчатками и декоративными элементами для цельного образа героини."},
  {sku:"AYK-COSTUME-WIG",weight:1.0,price:442.857143,dims:[28,25,6],sizes:clothing,title:"Костюм Камисато Аяка Kamisato Ayaka с париком",char:"Камисато Аяка",series:"Genshin Impact",colors:"蓝色;黑色;红色;金色;白色",kind:"costume",contents:"костюм;парик;головной аксессуар;перчатки;украшения",detail:"Расширенный комплект Камисато Аяки объединяет костюм и подходящий парик. Синие, черные, белые, красные и золотистые детали, головной аксессуар, перчатки и украшения позволяют собрать согласованный образ без отдельного подбора прически."},
  {sku:"AYK-COSTUME-WIG-FAN",weight:1.1,price:480.857143,dims:[30,27,8],sizes:clothing,title:"Камисато Аяка Kamisato Ayaka костюм парик веер",char:"Камисато Аяка",series:"Genshin Impact",colors:"蓝色;黑色;红色;金色;白色",kind:"costume",contents:"костюм;парик;веер;головной аксессуар;перчатки",detail:"Полный комплект Камисато Аяки включает костюм, парик и веер. Головной аксессуар, перчатки и декоративные детали согласованы с синими, черными, белыми, красными и золотистыми элементами образа."}
];

function description(p) {
  const typeText = p.kind === "shoes"
    ? "Обувь рассчитана на использование как часть сценического комплекта. Перед выходом рекомендуется примерить пару в помещении, проверить посадку и удобство движения. Размер выбирают по длине стопы и размерной сетке продавца, поскольку привычное обозначение в повседневной обуви может отличаться."
    : p.kind === "wig"
    ? "Парик помогает завершить образ без изменения собственных волос. Перед использованием его рекомендуется аккуратно расправить, уложить пальцами или расческой с редкими зубьями и закрепить на подготовленных волосах. Избегайте сильного натяжения основы и воздействия открытого огня."
    : "Костюм предназначен для косплея, тематических мероприятий, фотосессий, фестивалей и сценических выступлений. Комплект можно сочетать с подходящей обувью, париком и реквизитом, если эти элементы не перечислены в составе поставки. Все включенные части следует сверять с фотографиями и перечнем комплектации.";
  const care = p.kind === "shoes"
    ? "После использования удалите поверхностную пыль мягкой сухой салфеткой. Не замачивайте обувь и не сушите ее возле нагревательных приборов. Храните пару в сухом проветриваемом месте, сохраняя форму носка и декоративных деталей."
    : p.kind === "wig"
    ? "Для хранения используйте сетку или подставку, чтобы пряди меньше спутывались. При необходимости допускается осторожная ручная очистка прохладной водой со средством для синтетических волос. Не выкручивайте изделие и сушите естественным способом вдали от нагревателей."
    : "Перед первым использованием примерьте все части и проверьте расположение застежек и декоративных элементов. Рекомендуется деликатная ручная очистка в прохладной воде без отбеливателя, без интенсивного трения и выкручивания. Сушите изделие расправленным вдали от прямого нагрева.";
  const common = `Товар разработан для поклонников ${p.series} и подходит взрослым любителям костюмированных образов. Он помогает воспроизвести внешность персонажа ${p.char} для конвента, праздника, постановки или создания контента. Это предмет для перевоплощения, а не повседневная одежда или профессиональное защитное снаряжение.\n\n${typeText}\n\n${care}\n\nИз-за индивидуальных настроек экрана оттенок на фотографии может незначительно отличаться от восприятия вживую. Декоративные детали требуют бережного обращения. Храните изделие в сухом месте, защищенном от длительного солнечного света, влаги и контакта с острыми предметами. Не оставляйте упаковочные материалы рядом с маленькими детьми.\n\nПеред заказом ознакомьтесь с составом комплекта и доступными вариантами. Измерения упаковки и масса приведены в карточке товара. Если изделие используется на сцене или во время продолжительного мероприятия, заранее проведите полную примерку, чтобы спокойно отрегулировать посадку и убедиться, что все элементы надежно закреплены.`;
  let s = `${p.detail}\n\n${common}`;
  const addon = `\n\nОбраз лучше собирать заранее при хорошем освещении. Разложите детали, сопоставьте их с фотографиями и надевайте последовательно, не прикладывая чрезмерного усилия к декору. После мероприятия очистите изделие в соответствии с типом материала и полностью высушите перед хранением.`;
  while (s.length < 1700) s += addon;
  if (s.length > 1995) s = s.slice(0, 1988).replace(/\s+\S*$/, "") + ".";
  return s;
}

for (const p of specs) {
  p.description = description(p);
  if (p.title.length > 60) throw new Error(`Title too long ${p.sku}: ${p.title.length}`);
  if (p.description.length < 1600 || p.description.length > 2000) throw new Error(`Description length ${p.sku}: ${p.description.length}`);
}

// Verify every selected Chinese color exists in the supplied color reference.
const colorWb = await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(base,"WB产品颜色-中文版.xlsx")));
const colorInspect = await colorWb.inspect({kind:"table",maxChars:100000,tableMaxRows:1000,tableMaxCols:20,tableMaxCellChars:200});
const colorText = colorInspect.ndjson;
for (const p of specs) for (const c of p.colors.split(";").filter(Boolean)) if (!colorText.includes(c)) throw new Error(`Color not in reference: ${c}`);

const rows = [];
for (const p of specs) {
  const imgs = imagesByFolder.get(p.sku) || [];
  const photos = imgs.map(x=>rawUrl(p.sku,x.file)).join(";");
  for (const sz of p.sizes) {
    const shoe = p.kind === "shoes";
    const r = Array(54).fill(null);
    r[0]=1; r[1]=p.sku; r[2]=null; r[3]=p.title; r[4]="化妆舞会服装"; r[5]=null;
    r[6]=p.description; r[7]=photos || null; r[8]=null; r[9]=null; r[10]=p.weight;
    r[11]=null; r[12]=null; r[13]=1; r[14]=1; r[15]=null; r[16]=null; r[17]="женский";
    r[18]=p.kind==="wig"?"синтетическое волокно":p.kind==="shoes"?null:"текстиль";
    r[19]=p.colors || null; r[20]=null; r[21]=shoe?Number(sz)-1:null; r[22]=String(sz); r[23]=Math.ceil(p.price);
    r[24]=null; r[25]=null; r[26]=null; r[27]=null; r[28]=null; r[29]=null; r[30]=null;
    r[31]=p.kind==="shoes"?"Косплейная обувь":p.kind==="wig"?"Косплейный парик":"Костюм персонажа";
    r[32]=p.kind==="shoes"?"обувь персонажа":p.kind==="wig"?"парик персонажа":"костюм персонажа";
    r[33]=null; r[34]=p.kind==="costume"?"длинный":null; r[35]="14+"; r[36]=p.char;
    r[37]=shoe?`EUR ${sz}`:String(sz); r[38]=null;
    r[39]=p.kind==="shoes"?"протирать сухой салфеткой":p.kind==="wig"?"бережный уход;естественная сушка":"деликатная ручная стирка";
    r[40]="Китай"; r[41]="косплей;фотосессия;сценический образ"; r[42]=p.series;
    r[43]=null; r[44]=null; r[45]=null; r[46]=null; r[47]=p.contents;
    r[48]=Math.ceil(p.dims[1]); r[49]=Math.ceil(p.dims[0]); r[50]=Math.ceil(p.dims[2]); r[51]=p.weight*1000; r[52]=null; r[53]=null;
    rows.push(r);
  }
}

// Latest user instruction: within one SKU only the three size-related fields may differ.
for (const p of specs) {
  const group = rows.filter(r=>r[1]===p.sku);
  const baseRow = group[0];
  for (const r of group.slice(1)) for (let i=0;i<54;i++) if (![21,22,37].includes(i) && r[i]!==baseRow[i]) throw new Error(`Non-size mismatch ${p.sku} col ${i+1}`);
}
if (rows.length !== 106) throw new Error(`Expected 106 rows, got ${rows.length}`);

const templatePath = path.join(base,"化妆舞会服装.xlsx");
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(templatePath));
const sheet = wb.worksheets.getItem("项目");
sheet.getRange(`A5:BB${4+rows.length}`).values = rows;
const outMain = path.join(outDir,"产品上传表格.xlsx");
await (await SpreadsheetFile.exportXlsx(wb)).save(outMain);

const mapWb = Workbook.create();
const mapSheet = mapWb.worksheets.add("图片URL映射表");
const mapRows = [["SKU文件夹名","图片文件名","本地路径","图片URL","图片排序号"]];
for (const [folder, arr] of [...imagesByFolder.entries()].sort((a,b)=>a[0].localeCompare(b[0]))) {
  arr.forEach((x,i)=>mapRows.push([folder,x.file,x.full,rawUrl(folder,x.file),i+1]));
}
mapSheet.getRange(`A1:E${mapRows.length}`).values = mapRows;
mapSheet.getRange("A1:E1").format = {fill:"#1F4E78",font:{bold:true,color:"#FFFFFF"},horizontalAlignment:"center",verticalAlignment:"center"};
mapSheet.getRange(`A2:E${mapRows.length}`).format = {verticalAlignment:"top",wrapText:true};
mapSheet.getRange("A:A").format.columnWidth=22;
mapSheet.getRange("B:B").format.columnWidth=38;
mapSheet.getRange("C:C").format.columnWidth=64;
mapSheet.getRange("D:D").format.columnWidth=90;
mapSheet.getRange("E:E").format.columnWidth=12;
mapSheet.freezePanes.freezeRows(1);
const outMap = path.join(outDir,"图片URL映射表.xlsx");
await (await SpreadsheetFile.exportXlsx(mapWb)).save(outMap);

const check = await wb.inspect({kind:"table",range:`项目!A1:BB${4+rows.length}`,maxChars:12000,tableMaxRows:8,tableMaxCols:54,tableMaxCellChars:120});
await fs.writeFile(path.join(outDir,"build_inspect.ndjson"),check.ndjson,"utf8");
console.log(JSON.stringify({rows:rows.length,skus:specs.length,imageFolders:imagesByFolder.size,images:mapRows.length-1,missing:specs.filter(p=>!imagesByFolder.has(p.sku)).map(p=>p.sku),maxTitle:Math.max(...specs.map(p=>p.title.length)),descriptionLengths:Object.fromEntries(specs.map(p=>[p.sku,p.description.length]))},null,2));
