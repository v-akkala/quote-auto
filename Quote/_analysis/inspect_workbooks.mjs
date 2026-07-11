import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const books = [
  {
    name: "template",
    path: "C:/Users/Vikram/Downloads/Template Quote (2).xlsx",
  },
  {
    name: "example",
    path: "C:/Users/Vikram/Downloads/Quote Copy 203.xlsx",
  },
];

await fs.mkdir("_analysis", { recursive: true });

for (const book of books) {
  const input = await FileBlob.load(book.path);
  const workbook = await SpreadsheetFile.importXlsx(input);

  const summary = await workbook.inspect({
    kind: "workbook,sheet,region,match",
    searchTerm:
      "Service Quote|Customer|Designation|Company|Tel|E-mail|Address|Test Sample|Type of test|SAC|Batch|SQ|Dear|email|Dated|Total",
    maxChars: 22000,
    tableMaxRows: 95,
    tableMaxCols: 24,
    tableMaxCellChars: 140,
    options: { maxResults: 120 },
  });
  await fs.writeFile(`_analysis/${book.name}_inspect.ndjson`, summary.ndjson);

  const sheetSummary = await workbook.inspect({
    kind: "sheet",
    include: "id,name",
    maxChars: 5000,
  });
  await fs.writeFile(`_analysis/${book.name}_sheets.ndjson`, sheetSummary.ndjson);

  const png = await workbook.render({
    sheetName: "Quote",
    autoCrop: "all",
    scale: 1,
    format: "png",
  });
  await fs.writeFile(
    `_analysis/${book.name}_quote.png`,
    new Uint8Array(await png.arrayBuffer()),
  );
}
