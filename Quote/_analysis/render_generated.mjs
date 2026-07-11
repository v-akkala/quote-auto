import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "C:/Users/Vikram/Desktop/Coding Stuff/SKC/Quote/outputs/SQ26270001.xlsx";
const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);

const summary = await workbook.inspect({
  kind: "sheet,region,match",
  searchTerm: "SQ26270001|Vibration|High Temperature|Grand Total|Customer",
  maxChars: 8000,
  tableMaxRows: 30,
  tableMaxCols: 16,
  options: { maxResults: 40 },
});
await fs.writeFile("_analysis/generated_inspect.ndjson", summary.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  maxChars: 4000,
});
await fs.writeFile("_analysis/generated_errors.ndjson", errors.ndjson);

const preview = await workbook.render({
  sheetName: "SQ26270001",
  range: "A1:AX30",
  scale: 1,
  format: "png",
});
await fs.writeFile("_analysis/generated_preview.png", new Uint8Array(await preview.arrayBuffer()));
console.log("rendered _analysis/generated_preview.png");
