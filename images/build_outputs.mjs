import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const base = "H:/New project/wb-product-images/images";
const outputDir = path.join(base, "outputs", "01a00ef6-c440-74e3-83c4-ae1f8250c789");
const templatePath = path.join(base, "化妆舞会服装.xlsx");
const sourcePath = path.join(base, "COS服装及配饰清单.xlsx");
const colorPath = path.join(base, "WB产品颜色-中文版.xlsx");
const imageExts = new Set([".png", ".jpg", ".jpeg", ".webp"]);
const videoExts = new Set([".mp4", ".mov", ".avi", ".webm", ".mkv"]);

const meta = {
  "BLK-FUR-SWIG-CAP": { character:"Фурина", kind:"wig", colors:["白色","浅蓝色"], gender:"женский", contents:["парик","сетка"], features:["короткая версия","волнистые пряди","термостойкое волокно"], care:"бережное расчесывание;ручная укладка", composition:"термостойкое синтетическое волокно", decor:"голубые пряди" },
  "CMB-HEEL-CLR": { character:"Коломбина", kind:"shoes", colors:["透明","裸色"], gender:"женский", contents:["пара туфель"], features:["прозрачный верх","устойчивый каблук","квадратный мыс"], care:"протирать мягкой тканью", composition:"полиуретан", decor:"прозрачные ремешки", upper:"искусственная кожа;полиуретан" },
  "CMB-SET": { character:"Коломбина", kind:"costume", colors:["白色","淡蓝色","黑色","红色"], gender:"женский", contents:["платье","рукава","головной убор","украшения"], features:["многослойный крой","декоративные детали","сценический силуэт"], care:"ручная стирка;не отбеливать", composition:"текстиль", decor:"ленты;подвески;головной убор", sleeve:"длинный рукав" },
  "CMB-WIG": { character:"Коломбина", kind:"wig", colors:["黑色","红色"], gender:"женский", contents:["парик","сетка"], features:["длинные пряди","черно-красный градиент","ровная челка"], care:"бережное расчесывание;ручная укладка", composition:"термостойкое синтетическое волокно", decor:"красный градиент" },
  "DUR-SET": { character:"Дурин", kind:"costume", colors:["黑色","紫色","白色","红色"], gender:"мужской", contents:["верх","брюки","перчатки","хвост","пояс","декор"], features:["контрастная отделка","многослойный крой","образ дракона"], care:"ручная стирка;не отбеливать", composition:"текстиль", decor:"крылья;подвески;контрастные вставки", sleeve:"длинный рукав" },
  "DUR-SHOES": { character:"Дурин", kind:"shoes", colors:["黑色","银色","红色"], gender:"мужской", contents:["пара сапог"], features:["высокое голенище","контрастные вставки","устойчивая подошва"], care:"протирать мягкой тканью", composition:"искусственная кожа", decor:"серебристые детали;красные вставки", upper:"искусственная кожа" },
  "DUR-WIG": { character:"Дурин", kind:"wig", colors:["淡紫色","红色"], gender:"мужской", contents:["парик","сетка"], features:["многослойная стрижка","фиолетовый градиент","красные акценты"], care:"бережное расчесывание;ручная укладка", composition:"термостойкое синтетическое волокно", decor:"красные пряди" },
  "FLI-BOOT": { character:"Флинс", kind:"shoes", colors:["黑色","银色","紫色"], gender:"мужской", contents:["пара сапог"], features:["высокое голенище","шнуровка спереди","серебристый декор"], care:"протирать мягкой тканью", composition:"искусственная кожа", decor:"фиолетовые вставки;серебристые детали", upper:"искусственная кожа" },
  "FLI-COS-SET": { character:"Флинс", kind:"costume", colors:["黑色","紫色","银色"], gender:"мужской", contents:["рубашка","жилет","брюки","пояс","перчатки","подвески"], features:["асимметричный силуэт","серебристая отделка","сценический образ"], care:"ручная стирка;не отбеливать", composition:"текстиль", decor:"подвески;металлизированные детали", sleeve:"длинный рукав" },
  "FLI-WIG": { character:"Флинс", kind:"wig", colors:["深蓝色","银白色"], gender:"мужской", contents:["парик","сетка"], features:["длинные пряди","сине-белый градиент","многослойная челка"], care:"бережное расчесывание;ручная укладка", composition:"термостойкое синтетическое волокно", decor:"светлый градиент" },
  "FOC-COS-HAT": { character:"Фокалорс", kind:"costume", colors:["蓝色","白色","黑色","金色"], gender:"женский", contents:["жакет","жилет","рубашка","шорты","шляпа","перчатки","украшения"], features:["асимметричный крой","детальная отделка","шляпа в комплекте"], care:"ручная стирка;не отбеливать", composition:"текстиль", decor:"банты;подвески;золотистые детали", sleeve:"длинный рукав", hat:true },
  "FOC-COS-HAT-WIG": { character:"Фокалорс", kind:"costume_wig", colors:["蓝色","白色","黑色","金色"], gender:"женский", contents:["жакет","жилет","рубашка","шорты","шляпа","парик","сетка","аксессуары"], features:["полный образ","парик и шляпа","детальная отделка"], care:"ручная стирка;парик расчесывать бережно", composition:"текстиль;синтетическое волокно", decor:"банты;подвески;золотистые детали", sleeve:"длинный рукав", hat:true },
  "FOC-SHOE": { character:"Фурина", kind:"shoes", colors:["黑色","蓝色"], gender:"женский", contents:["пара туфель"], features:["декоративный бант","устойчивый каблук","закрытый мыс"], care:"протирать мягкой тканью", composition:"искусственная кожа", decor:"синий бант;пряжка", upper:"искусственная кожа" },
  "FOC-WIG-CAP": { character:"Фокалорс", kind:"wig", colors:["白色","浅蓝色"], gender:"женский", contents:["парик","сетка"], features:["длинные волнистые пряди","бело-голубой градиент","термостойкое волокно"], care:"бережное расчесывание;ручная укладка", composition:"термостойкое синтетическое волокно", decor:"голубой градиент" },
  "FUR-COS-HAT": { character:"Фурина", kind:"costume", colors:["蓝色","白色","黑色","金色"], gender:"женский", contents:["жакет","жилет","рубашка","шорты","шляпа","перчатки","украшения"], features:["многослойный крой","декоративная отделка","шляпа в комплекте"], care:"ручная стирка;не отбеливать", composition:"текстиль", decor:"банты;подвески;золотистые детали", sleeve:"длинный рукав", hat:true },
  "FUR-COS-HAT-WIG": { character:"Фурина", kind:"costume_wig", colors:["蓝色","白色","黑色","金色"], gender:"женский", contents:["жакет","жилет","рубашка","шорты","шляпа","парик","сетка","аксессуары"], features:["полный образ","парик и шляпа","многослойный крой"], care:"ручная стирка;парик расчесывать бережно", composition:"текстиль;синтетическое волокно", decor:"банты;подвески;золотистые детали", sleeve:"длинный рукав", hat:true },
  "KOK-COS-SET": { character:"Сангономия Кокоми", kind:"costume", colors:["紫色","粉红色","白色","蓝色"], gender:"женский", contents:["топ","накидка","юбка","перчатки","украшения","гольфы"], features:["морская палитра","многослойный силуэт","градиентная отделка"], care:"ручная стирка;не отбеливать", composition:"текстиль", decor:"банты;ленты;украшения", sleeve:"открытые плечи" },
  "KOK-SHOE": { character:"Сангономия Кокоми", kind:"shoes", colors:["紫色","金色","白色"], gender:"женский", contents:["пара босоножек"], features:["платформа","декоративный бант","открытый мыс"], care:"протирать мягкой тканью", composition:"искусственная кожа", decor:"фиолетовый бант;контрастный рисунок", upper:"искусственная кожа" },
  "KOK-WIG": { character:"Сангономия Кокоми", kind:"wig", colors:["粉红色","浅蓝色"], gender:"женский", contents:["парик"], features:["длинные пряди","розово-голубой градиент","боковой декор"], care:"бережное расчесывание;ручная укладка", composition:"термостойкое синтетическое волокно", decor:"градиентные пряди" },
  "NRE-SET": { character:"Николь Рейн", kind:"costume", colors:["白色","淡紫色","黑色","金色"], gender:"женский", contents:["платье","накидка","перчатки","пояс","украшения"], features:["длинный силуэт","градиентная отделка","декоративные акценты"], care:"ручная стирка;не отбеливать", composition:"текстиль", decor:"цепочки;подвески;контрастные вставки", sleeve:"длинный рукав" },
  "NRE-SHOES": { character:"Николь Рейн", kind:"shoes", colors:["白色","红色","黑色"], gender:"женский", contents:["пара ботильонов"], features:["высокий каблук","острый мыс","акцентный бант"], care:"протирать мягкой тканью", composition:"искусственная кожа", decor:"красный бант;ремешок", upper:"искусственная кожа" },
  "NRE-WIG": { character:"Николь Рейн", kind:"wig", colors:["浅金色"], gender:"женский", contents:["парик"], features:["длинные светлые пряди","мягкая челка","плетение по ободку"], care:"бережное расчесывание;ручная укладка", composition:"термостойкое синтетическое волокно", decor:"плетеный ободок" },
  "NVA-SET": { character:"Навия", kind:"costume", colors:["黑色","黄色","白色","金色"], gender:"женский", contents:["платье","шляпа","чокер","нарукавники","перчатки","чулки","серьги"], features:["асимметричная юбка","головной убор","контрастная отделка"], care:"ручная стирка;не отбеливать", composition:"текстиль", decor:"банты;кружево;золотистые детали", sleeve:"без рукавов" },
  "NVA-SET-WIG": { character:"Навия", kind:"costume_wig", colors:["黑色","黄色","白色","金色"], gender:"женский", contents:["платье","шляпа","чокер","перчатки","чулки","парик","аксессуары"], features:["полный образ","парик в комплекте","контрастная отделка"], care:"ручная стирка;парик расчесывать бережно", composition:"текстиль;синтетическое волокно", decor:"банты;кружево;золотистые детали", sleeve:"без рукавов" },
  "RDS-SET": { character:"Сёгун Райдэн", kind:"costume", colors:["紫色","白色","红色","黑色"], gender:"женский", contents:["верх","юбка","пояс","рукава","декор","чулки"], features:["многослойный крой","пояс оби","контрастные детали"], care:"ручная стирка;не отбеливать", composition:"текстиль", decor:"банты;пояс;подвески", sleeve:"отдельные рукава" },
  "RDS-SHOES": { character:"Сёгун Райдэн", kind:"shoes", colors:["红色","米色"], gender:"женский", contents:["пара сандалий"], features:["открытый мыс","устойчивый каблук","ремешок"], care:"протирать мягкой тканью", composition:"искусственная кожа", decor:"красный ремешок", upper:"искусственная кожа" },
  "RDS-WIG": { character:"Сёгун Райдэн", kind:"wig", colors:["深浓紫色"], gender:"женский", contents:["парик"], features:["длинная коса","ровная челка","фиолетовый оттенок"], care:"бережное расчесывание;ручная укладка", composition:"термостойкое синтетическое волокно", decor:"длинная коса" },
};

function normSku(s) {
  return String(s ?? "").normalize("NFKC").replace(/[‐‑‒–—―_ *×]/g, "-").replace(/-+/g,"-").trim().toUpperCase();
}
function sizesFromRemark(remark) {
  const s = String(remark ?? "");
  if (s.includes("均码")) return ["ONE"];
  const numeric = s.match(/(\d+)\s*-\s*(\d+)/);
  if (numeric) return Array.from({length:+numeric[2]-+numeric[1]+1},(_,i)=>String(+numeric[1]+i));
  const letter = s.match(/(XS|S|M|L|XL|XXL|XXXL)\s*-\s*(XS|S|M|L|XL|XXL|XXXL)/i);
  if (letter) {
    const order=["XS","S","M","L","XL","XXL","XXXL"];
    return order.slice(order.indexOf(letter[1].toUpperCase()), order.indexOf(letter[2].toUpperCase())+1);
  }
  return ["ONE"];
}
function cleanCode(s) { return normSku(s).replace(/[^A-Z0-9-]/g,""); }
function urlSegment(s) { return encodeURIComponent(s).replace(/'/g,"%27"); }
function titleFor(m, size) {
  const sizeText = size === "ONE" ? "единый размер" : size;
  let t;
  if (m.kind === "wig") t = `Косплей-парик ${m.character}, ${sizeText}`;
  else if (m.kind === "shoes") t = `Косплей-обувь ${m.character}, EUR ${size}`;
  else if (m.kind === "costume_wig") t = `Косплей-костюм ${m.character} с париком, ${size}`;
  else if (m.hat) t = `Косплей-костюм ${m.character} со шляпой, ${size}`;
  else t = `Косплей-костюм ${m.character}, размер ${size}`;
  if (t.length > 60) t = t.replace("Косплей-костюм", "Костюм").replace("Косплей-обувь","Обувь");
  return t;
}
function descFor(m, size) {
  const parts = [];
  const sizeText = size === "ONE" ? "единый размер" : (m.kind === "shoes" ? `размер EUR ${size}, российский размер ${+size-1}` : `размер ${size}`);
  if (m.kind === "wig") {
    parts.push(`Косплей-парик для образа ${m.character} передает узнаваемую форму прически и характерную цветовую гамму персонажа. Модель рассчитана на создание цельного сценического образа, фотосессии, тематического мероприятия или коллекционного костюма. Вариант исполнения соответствует фотографиям именно этого товара: ${m.features.join(", ")}.`);
    parts.push(`Парик выполнен из термостойкого синтетического волокна. Пряди уложены слоями, поэтому силуэт читается спереди, сбоку и сзади. Волокно имеет аккуратный блеск и сохраняет заданное направление при бережном обращении. Комплектация: ${m.contents.join(", ")}. Цвета и переходы оттенков подобраны под образ ${m.character}.`);
    parts.push(`Посадку можно отрегулировать внутренней сеткой. Перед первым использованием осторожно встряхните изделие, расправьте пряди руками и подправьте челку. Для более точного повторения прически допустима легкая ручная укладка при низкой температуре, если режим подходит для синтетического волокна. Не приближайте материал к открытому огню.`);
    parts.push(`После использования расчешите пряди от концов к корням редким гребнем. Храните парик на подставке или в свободной упаковке, избегая сильного сминания. Для очистки используйте прохладную воду и мягкое средство, не выкручивайте и сушите естественным способом вдали от нагревательных приборов.`);
    parts.push(`Размер модели: ${sizeText}. Перед оформлением заказа сравните желаемую длину, форму челки и расположение цветовых акцентов с изображениями товара. Оттенок может незначительно восприниматься иначе при разном освещении и настройках экрана.`);
  } else if (m.kind === "shoes") {
    parts.push(`Косплей-обувь для образа ${m.character} создана как завершающая часть костюма и повторяет заметные детали обуви персонажа. Модель подходит для фотосессий, фестивалей, тематических выступлений и костюмированных мероприятий. Особенности этого варианта: ${m.features.join(", ")}. Цвет и декоративные элементы соответствуют представленным изображениям.`);
    parts.push(`Верх выполнен из материалов с аккуратной поверхностью и контрастной отделкой. Конструкция поддерживает выразительный силуэт и позволяет сочетать обувь с соответствующим костюмом. В комплект входит ${m.contents.join(", ")}. Декор: ${m.decor}. Перед использованием проверьте фиксацию ремешков, шнуровки или застежек и убедитесь, что обувь сидит устойчиво.`);
    parts.push(`Указан европейский размер EUR. Российский размер для этой модели на одну единицу меньше европейского. Выбранный вариант: ${sizeText}. Для точного выбора измерьте длину стопы и сопоставьте ее с таблицей продавца, поскольку посадка косплейной обуви может отличаться от повседневной. При промежуточном значении ориентируйтесь на комфорт и особенности носка.`);
    parts.push(`Надевайте обувь на ровной сухой поверхности. Перед длительной съемкой или выступлением рекомендуется короткая примерка дома. Изделие предназначено прежде всего для создания образа, поэтому избегайте длительной ходьбы по неровному покрытию, влаги и сильного механического воздействия. Каблук и подошву следует регулярно осматривать.`);
    parts.push(`Для ухода удаляйте пыль мягкой сухой тканью. Не применяйте абразивные средства и растворители, особенно на прозрачных, окрашенных и металлизированных деталях. Сушите при комнатной температуре, не ставьте рядом с батареями и храните так, чтобы декоративные элементы не деформировались.`);
  } else {
    const withWig = m.kind === "costume_wig";
    parts.push(`Косплей-костюм ${m.character} предназначен для создания узнаваемого образа персонажа на фестивале, фотосессии, сценическом выступлении или тематическом мероприятии. Крой, сочетание оттенков и расположение декоративных деталей соответствуют представленным фотографиям товара. Особенности комплекта: ${m.features.join(", ")}.`);
    parts.push(`В комплект входят: ${m.contents.join(", ")}. Элементы образа сочетаются между собой по цвету и отделке, а многослойная конструкция помогает передать характерный силуэт персонажа. Декоративное оформление включает ${m.decor}. Обувь не входит в комплект, если она не показана в перечне комплектации как отдельный элемент.`);
    parts.push(`Материалы подобраны для костюмированного использования и позволяют сохранить форму деталей при аккуратной носке. Перед первым выходом расправьте одежду после упаковки, разложите аксессуары по назначению и выполните примерку всего комплекта. Мелкие украшения закрепляйте без чрезмерного усилия, чтобы не повредить ткань и крепления.`);
    parts.push(`Выбранный вариант: ${sizeText}. Для правильной посадки сверяйте рост, обхват груди, талии и бедер с таблицей размеров на изображениях. Не выбирайте размер только по привычной маркировке повседневной одежды. Если планируется нижний слой одежды или длительное выступление, учитывайте дополнительный запас для движения.`);
    parts.push(`Костюм подходит взрослым поклонникам косплея и участникам тематических проектов. Он рассчитан на выразительный внешний вид в кадре и на сцене. При движении следите за длинными, свободными и подвесными деталями, не допускайте их попадания в двери, механизмы и источники огня. Украшения не предназначены для силовой нагрузки.`);
    parts.push(`Рекомендуется ручная стирка в прохладной воде с мягким средством. Съемные аксессуары, жесткие элементы и украшения перед очисткой необходимо снять. Не используйте отбеливатель и интенсивный отжим. Сушите изделие в расправленном виде вдали от прямого нагрева, а храните на вешалке или аккуратно сложенным.`);
    if (withWig) parts.push(`Парик и сетка входят в этот вариант комплекта. Перед использованием аккуратно расправьте пряди, подстройте внутреннюю сетку и выполните легкую укладку. Парик следует хранить отдельно от металлических и жестких деталей костюма, чтобы волокно не спутывалось и не цеплялось за декор.`);
  }
  let text = parts.join("\n\n");
  if (text.length > 1995) text = text.slice(0,1990).replace(/\s+\S*$/,"") + ".";
  return text;
}

await fs.mkdir(outputDir,{recursive:true});

// Read source and approved colors.
const srcWb = await SpreadsheetFile.importXlsx(await FileBlob.load(sourcePath));
const srcValues = srcWb.worksheets.getItem("Sheet1").getRange("A1:J28").values;
const colorWb = await SpreadsheetFile.importXlsx(await FileBlob.load(colorPath));
const approvedColors = new Set(colorWb.worksheets.getItem("Sheet1").getRange("A1:A931").values.flat().filter(Boolean));

// Inventory folders and build exact raw URLs.
const dirents = (await fs.readdir(base,{withFileTypes:true})).filter(d=>d.isDirectory() && !["outputs","work_preview","node_modules"].includes(d.name));
const folders = new Map(dirents.map(d=>[normSku(d.name),d.name]));
const imageMap = new Map();
const mappingRows=[];
let videoFolderCount=0, videoUrlCount=0;
for (const [normalized,folderName] of folders) {
  const folderPath=path.join(base,folderName);
  const files=(await fs.readdir(folderPath,{withFileTypes:true})).filter(x=>x.isFile()).sort((a,b)=>a.name.localeCompare(b.name,"zh-Hans-CN",{numeric:true}));
  const urls=[];
  let imageNo=0;
  for (const f of files) {
    const ext=path.extname(f.name).toLowerCase();
    if (!imageExts.has(ext)) continue;
    imageNo++;
    const url=`https://raw.githubusercontent.com/ainner-will/wb-product-images/main/images/${urlSegment(folderName)}/${urlSegment(f.name)}`;
    urls.push(url);
    mappingRows.push([folderName,f.name,path.join(folderPath,f.name),url,imageNo,"","","","","",""]);
  }
  imageMap.set(normalized,urls);
}

// Build upload rows from source, expanding every size to an individual article.
const rows=[];
const stats={matched:0,unmatched:[],ambiguous:[],codes:new Set(),imageUrls:0};
for (let i=1;i<srcValues.length;i++) {
  const [,cnName,rawSku,weight,price,len,wid,hei,,remark]=srcValues[i];
  if (!rawSku) continue;
  const sku=normSku(rawSku);
  const m=meta[sku];
  if (!m) throw new Error(`Missing metadata for ${sku}`);
  const folderName=folders.get(sku);
  const urls=imageMap.get(sku) ?? [];
  if (folderName) stats.matched++; else stats.unmatched.push(String(rawSku));
  stats.imageUrls += urls.length;
  for (const c of m.colors) if (!approvedColors.has(c)) throw new Error(`Unapproved color ${c} for ${sku}`);
  for (const size of sizesFromRemark(remark)) {
    const codeBase=cleanCode(rawSku);
    const code=`${codeBase}-${size}`;
    if (!/^[A-Z0-9-]+$/.test(code) || stats.codes.has(code)) throw new Error(`Invalid or duplicate code ${code}`);
    stats.codes.add(code);
    const title=titleFor(m,size);
    if (title.length>60) throw new Error(`Title too long: ${title}`);
    const desc=descFor(m,size);
    if (desc.length>2000) throw new Error(`Description too long: ${code} ${desc.length}`);
    const isShoes=m.kind==="shoes";
    const ruSize=isShoes ? Number(size)-1 : "";
    const generalSize=size==="ONE" ? "единый" : (isShoes ? Number(size) : size);
    const purpose="косплей;фотосессия;сценический образ";
    const productType=isShoes?"Косплейная обувь":m.kind==="wig"?"Косплейный парик":"Костюм персонажа";
    const modelFeatures=m.features.join(";");
    const photoUrls=urls.join(";");
    const row = Array(54).fill("");
    row[0]=1;
    row[1]=code;
    row[3]=title;
    row[4]="化妆舞会服装";
    row[6]=desc;
    row[7]=photoUrls;
    row[10]=Number(weight);
    row[13]=1;
    row[17]=m.gender;
    row[18]=m.composition;
    row[19]=m.colors.join(";");
    row[21]=ruSize;
    row[22]=generalSize;
    row[23]=Math.ceil(Number(price));
    row[31]=title;
    row[32]=productType;
    row[33]=m.character;
    row[34]=m.sleeve ?? "";
    row[36]=m.character;
    row[37]=size==="ONE"?"единый":(isShoes?`EUR ${size}`:size);
    row[38]=modelFeatures;
    row[39]=m.care;
    row[40]="Китай";
    row[41]=purpose;
    row[42]="Genshin Impact";
    row[44]=isShoes?(m.upper??m.composition):m.composition;
    row[46]=m.decor;
    row[47]=m.contents.join(";");
    row[48]=Math.ceil(Number(wid));
    row[49]=Math.ceil(Number(len));
    row[50]=Math.ceil(Number(hei));
    row[51]=Number(weight)*1000;
    rows.push(row);
  }
}

// Enforce the 100-character rule outside description and URLs.
for (const [ri,row] of rows.entries()) for (let ci=0;ci<row.length;ci++) {
  if ([6,7,8].includes(ci) || row[ci]===null || row[ci]===undefined) continue;
  if (String(row[ci]).length>100) throw new Error(`Cell over 100 chars at row ${ri+5}, col ${ci+1}`);
}

// Edit only the data region of the imported template.
const uploadWb = await SpreadsheetFile.importXlsx(await FileBlob.load(templatePath));
const uploadSheet=uploadWb.worksheets.getItem("项目");
uploadSheet.getRangeByIndexes(4,0,rows.length,54).values=rows;
uploadSheet.getRangeByIndexes(4,0,rows.length,54).format.rowHeight=45;
uploadSheet.freezePanes.freezeRows(4);
uploadSheet.freezePanes.freezeColumns(2);
const uploadOut=await SpreadsheetFile.exportXlsx(uploadWb);
const uploadPath=path.join(outputDir,"产品上传表格.xlsx");
await uploadOut.save(uploadPath);

// Create the required URL mapping workbook.
const mapWb=Workbook.create();
const mapSheet=mapWb.worksheets.add("图片URL映射表");
const headers=["SKU文件夹名","图片文件名","本地路径","图片URL","图片排序号","视频文件夹名","视频文件名","视频本地路径","视频URL","视频排序号","适用SKU范围"];
mapSheet.getRange("A1:K1").values=[headers];
if (mappingRows.length) mapSheet.getRangeByIndexes(1,0,mappingRows.length,headers.length).values=mappingRows;
mapSheet.getRange(`A1:K${mappingRows.length+1}`).format={font:{name:"Microsoft YaHei",size:10},verticalAlignment:"center"};
mapSheet.getRange("A1:K1").format={fill:"#1F4E78",font:{name:"Microsoft YaHei",size:10,bold:true,color:"#FFFFFF"},horizontalAlignment:"center",verticalAlignment:"center",wrapText:true};
mapSheet.getRange(`A2:B${mappingRows.length+1}`).format.wrapText=true;
mapSheet.getRange(`C2:D${mappingRows.length+1}`).format.wrapText=false;
const widths=[22,38,62,95,12,20,30,55,90,12,20];
widths.forEach((w,i)=>mapSheet.getRangeByIndexes(0,i,mappingRows.length+1,1).format.columnWidth=w);
mapSheet.freezePanes.freezeRows(1);
const mapOut=await SpreadsheetFile.exportXlsx(mapWb);
const mapPath=path.join(outputDir,"图片URL映射表.xlsx");
await mapOut.save(mapPath);

// Compact inspections and visual verification assets.
const uploadCheck=await uploadWb.inspect({kind:"table",range:`项目!A1:BB12`,include:"values,formulas",tableMaxRows:12,tableMaxCols:54,maxChars:12000});
const errorCheck=await uploadWb.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",options:{useRegex:true,maxResults:300},summary:"final formula error scan",maxChars:4000});
await fs.writeFile(path.join(outputDir,"verification.txt"),uploadCheck.ndjson+"\n"+errorCheck.ndjson,"utf8");
for (const [name,range] of [["upload_left","A1:N16"],["upload_right","U1:BB16"],["upload_tail","A132:BB143"]]) {
  const png=await uploadWb.render({sheetName:"项目",range,scale:1,format:"png"});
  await fs.writeFile(path.join(outputDir,`${name}.png`),new Uint8Array(await png.arrayBuffer()));
}
const helpPng=await uploadWb.render({sheetName:"使用说明",range:"A1:D20",scale:1,format:"png"});
await fs.writeFile(path.join(outputDir,"instructions.png"),new Uint8Array(await helpPng.arrayBuffer()));
const mapPng=await mapWb.render({sheetName:"图片URL映射表",range:"A1:E12",scale:1,format:"png"});
await fs.writeFile(path.join(outputDir,"mapping.png"),new Uint8Array(await mapPng.arrayBuffer()));

// Structural and content report.
const report={
  skuFolders:folders.size,videoFolders:videoFolderCount,imageUrls:mappingRows.length,videoUrls:videoUrlCount,
  matchedSkus:stats.matched,unmatchedSkus:stats.unmatched,ambiguousSkus:stats.ambiguous,
  productRows:rows.length,uniqueCodes:stats.codes.size,allCodesUnique:stats.codes.size===rows.length,
  sheetNames:[uploadWb.worksheets.getItemAt(0).name,uploadWb.worksheets.getItemAt(1).name],
  outputPaths:{uploadPath,mapPath},descriptionsUnique:new Set(rows.map(r=>r[6])).size===rows.length,
  colorsApproved:rows.every(r=>String(r[19]).split(";").filter(Boolean).every(c=>approvedColors.has(c))),
  urlChecks:{allRaw:rows.every(r=>!r[7]||r[7].split(";").every(u=>u.startsWith("https://raw.githubusercontent.com/ainner-will/wb-product-images/main/images/"))),noSpaces:rows.every(r=>!r[7]||!r[7].includes(" "))}
};
await fs.writeFile(path.join(outputDir,"report.json"),JSON.stringify(report,null,2),"utf8");
console.log(JSON.stringify(report,null,2));
