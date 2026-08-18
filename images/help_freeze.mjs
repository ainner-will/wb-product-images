import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load("H:/New project/wb-product-images/images/化妆舞会服装.xlsx"));
console.log(wb.help("worksheet.freezePanes",{include:"index,examples,notes",maxChars:5000}).ndjson);
