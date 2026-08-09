import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = "C:/Users/goenk/Documents/Codex/2026-05-12/Goldiee Purchase Bot";
const inputPath = path.join(root, "downloads", "physical_stock_source_2026-08-05.json");
const outputPath = path.join(root, "downloads", "Tally Physical Stock Sheet 2026-08-05.xlsx");

const payload = JSON.parse(await fs.readFile(inputPath, "utf8"));
const { summary, rows } = payload;

function writeMatrix(sheet, row, col, matrix) {
  sheet.getRangeByIndexes(row, col, matrix.length, matrix[0].length).values = matrix;
}

function styleHeader(range, fill = "#1F4E78") {
  range.format = {
    fill,
    font: { bold: true, color: "#FFFFFF" },
    wrapText: true,
  };
}

function setColumnWidths(sheet, widths) {
  for (const [col, width] of Object.entries(widths)) {
    sheet.getRange(`${col}:${col}`).format.columnWidth = width;
  }
}

const workbook = Workbook.create();

const report = workbook.worksheets.add("Physical Stock Sheet");
const reportRows = [
  ["S. No.", "Item Name", "PTR", "MRP", "Full Cases", "Loose PCS", "PCS per Case", "Total PCS", "Check PCS"],
  ...rows.map((row, index) => [
    index + 1,
    row.item_name,
    row.ptr,
    row.mrp,
    row.full_cases,
    row.loose_pcs,
    row.pcs_per_case,
    row.total_pcs,
    row.full_cases * row.pcs_per_case + row.loose_pcs,
  ]),
];
writeMatrix(report, 0, 0, reportRows);
styleHeader(report.getRange("A1:I1"));
report.freezePanes.freezeRows(1);
report.showGridLines = false;
setColumnWidths(report, {
  A: 8,
  B: 52,
  C: 12,
  D: 10,
  E: 13,
  F: 12,
  G: 13,
  H: 12,
  I: 12,
});
report.getRange(`A1:I${reportRows.length}`).format.borders = {
  preset: "all",
  style: "thin",
  color: "#D9E2F3",
};
report.getRange(`D2:D${reportRows.length}`).format.numberFormat = "#,##0.00";
report.getRange(`E2:I${reportRows.length}`).format.numberFormat = "#,##0";

const summarySheet = workbook.worksheets.add("Summary");
writeMatrix(summarySheet, 0, 0, [
  ["Tally Physical Stock Sheet", ""],
  ["Generated Date", "2026-08-05"],
  ["Source Stock Export", summary.source],
  ["Rows", summary.rows],
  ["Total PCS", summary.total_pcs],
  ["Total Full Cases", summary.total_full_cases],
  ["Total Loose PCS", summary.total_loose_pcs],
  ["Rows Missing Conversion", summary.missing_conversion_rows],
  ["Conversion Rule", "Full cases and loose PCS are calculated from item master PCS per Case."],
]);
summarySheet.getRange("A1:B1").merge();
summarySheet.getRange("A1:B1").format = {
  fill: "#0F766E",
  font: { bold: true, color: "#FFFFFF", size: 14 },
};
summarySheet.getRange("A2:A9").format = { font: { bold: true } };
summarySheet.getRange("B4:B8").format.numberFormat = "#,##0";
summarySheet.showGridLines = false;
setColumnWidths(summarySheet, { A: 28, B: 110 });

const inspect = await workbook.inspect({
  kind: "table",
  sheetId: "Physical Stock Sheet",
  range: "A1:I15",
  include: "values,formulas",
  tableMaxRows: 15,
  tableMaxCols: 9,
});
console.log(inspect.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
});
console.log(errors.ndjson);

await workbook.render({ sheetName: "Physical Stock Sheet", range: "A1:I25", scale: 1, format: "png" });
await workbook.render({ sheetName: "Summary", range: "A1:B9", scale: 1, format: "png" });

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(JSON.stringify({ outputPath, rows: summary.rows, totalPcs: summary.total_pcs }));
