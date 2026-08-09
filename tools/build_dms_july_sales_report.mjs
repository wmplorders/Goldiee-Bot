import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = "C:/Users/goenk/Documents/Codex/2026-05-12/Goldiee Purchase Bot";
const inputPath = path.join(root, "downloads", "dms_july_2026_sales_aggregate.json");
const outputPath = path.join(root, "downloads", "DMS Item Sales July 2026.xlsx");

const payload = JSON.parse(await fs.readFile(inputPath, "utf8"));
const { summary, items, raw_headers: rawHeaders, raw_rows: rawRows } = payload;

function setValues(sheet, row, col, values) {
  sheet.getRangeByIndexes(row, col, values.length, values[0].length).values = values;
}

function styleHeader(range, color = "#1F4E78") {
  range.format = {
    fill: color,
    font: { bold: true, color: "#FFFFFF" },
    wrapText: true,
  };
}

function styleSheet(sheet, widths) {
  sheet.showGridLines = false;
  sheet.freezePanes.freezeRows(1);
  for (const [col, width] of Object.entries(widths)) {
    sheet.getRange(`${col}:${col}`).format.columnWidth = width;
  }
}

const workbook = Workbook.create();

const summarySheet = workbook.worksheets.add("Summary");
setValues(summarySheet, 0, 0, [
  ["DMS Item Sales - July 2026", ""],
  ["Date From", summary.date_from],
  ["Date To", summary.date_to],
  ["Source Rows", summary.raw_rows],
  ["Items Sold", summary.item_count],
  ["Total Quantity (Cases)", summary.total_cases],
  ["Total Net Value", summary.total_net_value],
  ["Quantity Field Used", summary.quantity_field],
  ["Value Field Used", summary.value_field],
  ["Source CSV", summary.source_csv],
]);
summarySheet.getRange("A1:B1").format = {
  fill: "#0F766E",
  font: { bold: true, color: "#FFFFFF", size: 14 },
};
summarySheet.getRange("A1:B1").merge();
summarySheet.getRange("A2:A10").format = { font: { bold: true } };
summarySheet.getRange("B6").format.numberFormat = "#,##0.00";
summarySheet.getRange("B7").format.numberFormat = "#,##0.00";
styleSheet(summarySheet, { A: 28, B: 110 });

const reportSheet = workbook.worksheets.add("Item Sales");
const reportRows = [
  ["S. No.", "Item Name", "Quantity (Cases)", "Net Value", "Invoice Rows"],
  ...items.map((row) => [
    row.s_no,
    row.item_name,
    row.quantity_cases,
    row.net_value,
    row.invoice_rows,
  ]),
  ["", "TOTAL", summary.total_cases, summary.total_net_value, ""],
];
setValues(reportSheet, 0, 0, reportRows);
styleHeader(reportSheet.getRange("A1:E1"));
const totalRow = reportRows.length;
reportSheet.getRange(`A${totalRow}:E${totalRow}`).format = {
  fill: "#E2F0D9",
  font: { bold: true },
};
reportSheet.getRange(`C2:C${totalRow}`).format.numberFormat = "#,##0.00";
reportSheet.getRange(`D2:D${totalRow}`).format.numberFormat = "#,##0.00";
reportSheet.getRange(`A1:E${totalRow}`).format.borders = {
  preset: "all",
  style: "thin",
  color: "#D9E2F3",
};
styleSheet(reportSheet, { A: 8, B: 58, C: 18, D: 18, E: 12 });

const rawSheet = workbook.worksheets.add("Raw Dump");
const rawMatrix = [
  rawHeaders,
  ...rawRows.map((row) => rawHeaders.map((header) => row[header] ?? "")),
];
setValues(rawSheet, 0, 0, rawMatrix);
styleHeader(rawSheet.getRangeByIndexes(0, 0, 1, rawHeaders.length));
styleSheet(rawSheet, { A: 16, B: 26, C: 16, D: 16, E: 16, F: 16 });

const inspect = await workbook.inspect({
  kind: "table",
  sheetId: "Item Sales",
  range: "A1:E12",
  include: "values,formulas",
  tableMaxRows: 12,
  tableMaxCols: 5,
});
console.log(inspect.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
});
console.log(errors.ndjson);

await workbook.render({ sheetName: "Summary", range: "A1:B10", scale: 1, format: "png" });
await workbook.render({ sheetName: "Item Sales", range: "A1:E25", scale: 1, format: "png" });
await workbook.render({ sheetName: "Raw Dump", range: "A1:H20", scale: 1, format: "png" });

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(JSON.stringify({ outputPath, rows: items.length, totalCases: summary.total_cases, totalNetValue: summary.total_net_value }));
