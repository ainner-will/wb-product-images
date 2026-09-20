import fs from 'node:fs/promises';
import path from 'node:path';
import JSZip from 'jszip';
import {FileBlob, SpreadsheetFile, Workbook} from '@oai/artifact-tool';

const root=path.resolve('.');
const out=path.join(root,'outputs','phone-case-2026-09-20');
await fs.mkdir(out,{recursive:true});
const source=await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(root,'适用iPhone18Fold折叠屏支架手机壳磁吸无线.xlsx')));
const sourceRows=source.worksheets.getItemAt(0).getRange('A2:M15').values;
const colorsBook=await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(root,'WB产品颜色-中文版.xlsx')));
const allowed=new Set(colorsBook.worksheets.getItemAt(0).getUsedRange(true).values.flat().filter(Boolean).map(String));
const caseBook=await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(root,'手机壳.xlsx')));
const sheet=caseBook.worksheets.getItem('项目');
const raw='https://raw.githubusercontent.com/ainner-will/wb-product-images/main/images/';
const enc=s=>encodeURIComponent(s).replace(/[!'()*]/g,c=>'%'+c.charCodeAt(0).toString(16).toUpperCase());
const url=(...parts)=>raw+parts.map(enc).join('/');
const extImage=/\.(png|jpe?g|webp)$/i;
const extVideo=/\.(mp4|mov|avi|webm|mkv)$/i;
const skuSet=new Set(sourceRows.map(r=>String(r[6])));
const imageRows=[],videoRows=[],productRows=[];
const colorMap={TBK:'透明;黑色',DBL:'深蓝色',TGY:'灰色',GLD:'金色',CRD:'樱桃红',DPU:'紫色',ORG:'橙色',BLK:'黑色',CLR:'透明',MBK:'黑色'};
const colorRu={TBK:'прозрачно-чёрный',DBL:'тёмно-синий',TGY:'титаново-серый',GLD:'золотистый',CRD:'вишнёво-красный',DPU:'тёмно-фиолетовый',ORG:'оранжевый',BLK:'чёрный',CLR:'прозрачный',MBK:'матовый чёрный'};
const colorDistinct={TBK:'Полупрозрачная тёмная отделка оставляет заметным силуэт устройства, при этом делает его внешний вид более сдержанным.',DBL:'Глубокий синий оттенок создаёт спокойный акцент и хорошо сочетается с повседневными аксессуарами.',TGY:'Сдержанный титаново-серый тон подчёркивает геометрию складного корпуса и не выглядит слишком ярко.',GLD:'Золотистое исполнение добавляет тёплый металлический акцент к форме складного смартфона.',CRD:'Насыщенный вишнёво-красный цвет делает чехол заметной деталью повседневного комплекта.',DPU:'Тёмно-фиолетовый оттенок выглядит выразительно, сохраняя спокойный характер оформления.',ORG:'Оранжевый цвет выделяет смартфон и помогает быстро найти его среди других вещей.',BLK:'Классический чёрный цвет подходит к деловому и повседневному стилю.',CLR:'Прозрачное исполнение позволяет видеть оформление самого телефона и не перегружает его внешний вид.',MBK:'Матовая чёрная поверхность даёт более сдержанный внешний вид и меньше акцентирует внимание на бликах.'};
function description({model,colour,key,mag,rotate,fold}){
 const intro=`Чехол для Apple iPhone 18 ${model} в цвете ${colour} рассчитан на соответствующую модель складного смартфона. Он закрывает заднюю часть и боковые грани, помогает снизить риск повседневных потёртостей при контакте с сумкой, столом или другими предметами. Конструкция учитывает складной форм-фактор: область центрального шарнира прикрыта, а телефон можно открывать и закрывать без снятия чехла. Вырез под блок камер сохраняет доступ к объективам; расположение элементов корпуса оставляет доступ к основным кнопкам и разъёмам. Перед покупкой сверьте модель устройства с названием, поскольку форма и расположение камер у разных телефонов могут отличаться.`;
 let feature;
 if(fold) feature=`Эта версия относится к серии для iPhone 18 Fold. На задней панели расположен магнитный круг для совместимых магнитных аксессуаров и беспроводной зарядки. Встроенная откидная опора раскладывается, когда удобно поставить смартфон на стол для просмотра видео, чтения или видеосвязи, и складывается обратно для переноски. Поверхность корпуса визуально имеет мягкую, приятную на вид отделку. Чехол сочетает защиту шарнира с опорой, но сам смартфон и зарядное устройство в комплект не входят.`;
 else if(rotate) feature=`Это исполнение для iPhone 18 Duo с магнитным кольцом и поворотной опорой на 360°. Кольцо можно повернуть и использовать как подставку, подбирая удобную ориентацию телефона для просмотра видео или общения. Магнитная зона предназначена для совместимых магнитных аксессуаров и беспроводной зарядки. В отличие от базовой и обычной магнитной версии, здесь есть именно поворотный опорный механизм. Перед использованием проверьте совместимость зарядного устройства и аксессуаров с вашим телефоном.`;
 else if(mag) feature=`Это магнитная версия для iPhone 18 Duo. На задней части предусмотрено магнитное кольцо для совместимых магнитных аксессуаров и беспроводной зарядки. В отличие от базового исполнения, здесь добавлена магнитная зона; поворотной подставки в этой версии нет. При зарядке и креплении используйте аксессуары, подходящие для вашего смартфона. Чехол не заменяет защитное стекло экрана и не является водонепроницаемым изделием.`;
 else feature=`Это базовая версия для iPhone 18 Duo без заявленного магнитного крепления и без поворотной подставки. Её назначение — защита корпуса и шарнира в повседневном использовании, когда дополнительные функции магнитного держателя не нужны. На изображениях показана аккуратная спинка и оформленная зона камеры. Если требуется магнитное крепление или опора для установки телефона на стол, выбирайте соответствующий вариант этой серии, а не данное базовое исполнение.`;
 const finish=`${colorDistinct[key]} Чехол удобен для повседневного ношения: складная конструкция сохраняет характерный формат устройства, а продуманная форма позволяет пользоваться камерой, кнопками и портом. Для ухода достаточно протирать поверхность мягкой сухой или слегка влажной салфеткой; не используйте абразивные составы. Комплектация — один чехол. Изображения помогают оценить оттенок и детали конструкции, однако восприятие цвета может немного зависеть от настроек экрана. Убедитесь, что выбран нужный вариант цвета и функций именно для указанной модели телефона.`;
 const result=[intro,feature,finish].join('\n\n');
 if(result.length>2000||result.length<1200) throw Error(`description length ${result.length}`);
 return result;
}
function title({model,colour,mag,rotate,fold}){
 const part=fold?'магнитный':rotate?'опора 360°':mag?'магнитный':'защита шарнира';
 const t=`Чехол Apple 18 ${model}, ${part}, ${colour}`;
 if(t.length>60) throw Error(`title length ${t.length}: ${t}`);
 return t;
}
for(let i=0;i<sourceRows.length;i++){
 const src=sourceRows[i], sku=String(src[6]),fold=sku.startsWith('FPC-IP18F-MAG-ST-'),model=fold?'Fold':'Duo';
 const key=fold?sku.split('-').at(-1):sku.includes('MBK')?'MBK':'CLR';
 const mag=fold||sku.includes('-MAG'); const rotate=sku.endsWith('-R360');
 const c=colorMap[key]; if(!c) throw Error(`missing color ${sku}`);
 for(const part of c.split(';')) if(!allowed.has(part)) throw Error(`color not allowed ${part}`);
 const dir=path.join(root,sku);
 const files=(await fs.readdir(dir,{withFileTypes:true})).filter(e=>e.isFile()&&extImage.test(e.name)).map(e=>e.name).sort((a,b)=>a.localeCompare(b,'en'));
 if(!files.length) throw Error(`no images ${sku}`);
 const imageUrls=[];
 for(let j=0;j<files.length;j++){
  const name=files[j],u=url(sku,name);imageUrls.push(u);
  imageRows.push([sku,name,path.join(dir,name),u,j+1]);
 }
 const vf=fold?'视频':'视频1';
 const vfiles=(await fs.readdir(path.join(root,vf),{withFileTypes:true})).filter(e=>e.isFile()&&extVideo.test(e.name)).map(e=>e.name).sort((a,b)=>a.localeCompare(b,'en'));
 if(vfiles.length!==1) throw Error(`expected one video in ${vf}`);
 const vu=url(vf,vfiles[0]); videoRows.push([vf,vfiles[0],path.join(root,vf,vfiles[0]),vu,1,sku]);
 const vals=Array(43).fill(null);
 const set=(col,val)=>{vals[col-1]=val};
 set(1,1);set(2,sku);set(4,title({model,colour:colorRu[key],mag,rotate,fold}));set(5,'Чехол для телефона');
 set(7,description({model,colour:colorRu[key],key,mag,rotate,fold}));set(8,imageUrls.join(';'));set(9,vu);
 set(11,src[11]);set(14,1);set(17,c);set(19,Math.ceil(Number(src[12])));
 set(26,`Apple iPhone 18 ${model}`);set(28,fold?'защита шарнира;магнитное крепление;подставка':rotate?'защита шарнира;магнитное крепление;поворотная опора':mag?'защита шарнира;магнитное крепление':'защита шарнира');
 set(30,'чехол-накладка');set(32,'Китай');set(33,'Apple');set(36,'1 чехол');
 set(37,Math.ceil(Number(src[8])));set(38,Math.ceil(Number(src[7])));set(39,Math.ceil(Number(src[9])));set(40,Math.ceil(Number(src[11])*1000));
 productRows.push(vals);
}
if(productRows.length!==14||imageRows.length!==114||videoRows.length!==14) throw Error(`unexpected counts ${productRows.length}/${imageRows.length}/${videoRows.length}`);
sheet.getRange('A5:AQ18').values=productRows;
caseBook.recalculate();
const intermediate=path.join(out,'_artifact_draft.xlsx');
await (await SpreadsheetFile.exportXlsx(caseBook)).save(intermediate);

// Patch only new data rows into the original OOXML package, retaining its sheet instructions,
// metadata, validation, formatting, hidden columns, print settings and all other package parts.
const zip=await JSZip.loadAsync(await fs.readFile(path.join(root,'手机壳.xlsx')));
const xmlPath='xl/worksheets/sheet1.xml';
let xml=await zip.file(xmlPath).async('string');
if(!xml.includes('</sheetData>')) throw Error('template sheetData missing');
const esc=s=>String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');
function col(n){let x='';while(n){n--;x=String.fromCharCode(65+n%26)+x;n=Math.floor(n/26)}return x}
let rowXml='';
for(let ri=0;ri<productRows.length;ri++){
 const rn=ri+5;
 rowXml+=`<row r="${rn}">`;
 for(let ci=0;ci<43;ci++){
  const v=productRows[ri][ci]; if(v===null||v==='') continue;
  const ref=col(ci+1)+rn;
  rowXml+=typeof v==='number'?`<c r="${ref}"><v>${v}</v></c>`:`<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
 }
 rowXml+='</row>';
}
xml=xml.replace('</sheetData>',rowXml+'</sheetData>');
xml=xml.replace(/<dimension ref="[^"]*"\s*\/>/, '<dimension ref="A1:AQ18"/>');
zip.file(xmlPath,xml);
await fs.writeFile(path.join(out,'产品上传表格.xlsx'),await zip.generateAsync({type:'nodebuffer',compression:'DEFLATE'}));
await fs.rm(intermediate);
await fs.rm(intermediate+'.inspect.ndjson',{force:true});

const map=Workbook.create();
const pics=map.worksheets.add('图片URL映射');
pics.getRange('A1:E1').values=[['SKU文件夹名','图片文件名','本地路径','图片URL','图片排序号']];
pics.getRange(`A2:E${imageRows.length+1}`).values=imageRows;
pics.getRange('A1:E1').format.fill='#DCEAF7';
pics.getRange('A1:E1').format.font.bold=true;
const vids=map.worksheets.add('视频URL映射');
vids.getRange('A1:F1').values=[['视频文件夹名','视频文件名','本地路径','视频URL','视频排序号','适用SKU范围']];
vids.getRange(`A2:F${videoRows.length+1}`).values=videoRows;
vids.getRange('A1:F1').format.fill='#DCEAF7';
vids.getRange('A1:F1').format.font.bold=true;
map.recalculate();
await (await SpreadsheetFile.exportXlsx(map)).save(path.join(out,'图片URL映射表.xlsx'));
await fs.rm(path.join(out,'图片URL映射表.xlsx.inspect.ndjson'),{force:true});

console.log(JSON.stringify({out,products:productRows.length,images:imageRows.length,videoAssignments:videoRows.length,uniqueVideos:new Set(videoRows.map(r=>r[3])).size,skuFolders:[...skuSet],titleLengths:productRows.map(r=>r[3].length),descriptionLengths:productRows.map(r=>r[6].length),colors:productRows.map(r=>r[16])},null,2));
