import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = "C:/Users/goenk/Documents/Codex/2026-05-12/Goldiee Purchase Bot";
const inputPath = path.join(root, "downloads", "physical_stock_source_2026-08-06_corrected.json");
const outputPath = path.join(root, "downloads", "Tally Physical Stock Sheet 2026-08-06 corrected.xlsx");

const payload = JSON.parse(await fs.readFile(inputPath, "utf8"));
const { summary, positive_rows: positives, negative_rows: negatives } = payload;

function writeMatrix(sheet, row, col, matrix) {
  sheet.getRangeByIndexes(row, col, matrix.length, matrix[0].length).values = matrix;
}

function styleHeader(range, fill = "#1F4E78") {
  range.format = { fill, font: { bold: true, color: "#FFFFFF" }, wrapText: true };
}

function setColumnWidths(sheet, widths) {
  for (const [col, width] of Object.entries(widths)) {
    sheet.getRange(`${col}:${col}`).format.columnWidth = width;
  }
}

function styleGrid(sheet, range) {
  sheet.getRange(range).format.borders = { preset: "all", style: "thin", color: "#D9E2F3" };
  sheet.showGridLines = false;
  sheet.freezePanes.freezeRows(1);
}

const workbook = Workbook.create();

const report = workbook.worksheets.add("Physical Stock Sheet");
const reportRows = [
  ["S. No.", "Item Name", "PTR", "MRP", "Full Cases", "Loose PCS", "PCS per Case", "Total PCS", "Net Same MRP PCS", "Remarks"],
  ...positives.map((row, index) => [
    index + 1,
    row.item_name,
    row.ptr,
    row.mrp,
    row.full_cases,
    row.loose_pcs,
    row.pcs_per_case,
    row.total_pcs,
    row.net_same_mrp_pcs,
    row.remarks,
  ]),
];
writeMatrix(report, 0, 0, reportRows);
styleHeader(report.getRange("A1:J1"));
styleGrid(report, `A1:J${reportRows.length}`);
setColumnWidths(report, { A: 8, B: 52, C: 12, D: 10, E: 13, F: 12, G: 13, H: 12, I: 16, J: 72 });
report.getRange(`D2:D${reportRows.length}`).format.numberFormat = "#,##0.00";
report.getRange(`E2:I${reportRows.length}`).format.numberFormat = "#,##0";

const negative = workbook.worksheets.add("Negative Stock Review");
const negativeRows = [
  ["Item Name", "PTR", "MRP", "Full Cases", "Loose PCS", "PCS per Case", "Total PCS", "Net Same MRP PCS", "Remarks"],
  ...negatives.map((row) => [
    row.item_name,
    row.ptr,
    row.mrp,
    row.full_cases,
    row.loose_pcs,
    row.pcs_per_case,
    row.total_pcs,
    row.net_same_mrp_pcs,
    row.remarks,
  ]),
];
writeMatrix(negative, 0, 0, negativeRows);
styleHeader(negative.getRange("A1:I1"), "#9C0006");
styleGrid(negative, `A1:I${negativeRows.length}`);
setColumnWidths(negative, { A: 52, B: 12, C: 10, D: 13, E: 12, F: 13, G: 12, H: 16, I: 58 });
negative.getRange(`C2:C${Math.max(2, negativeRows.length)}`).format.numberFormat = "#,##0.00";
negative.getRange(`D2:H${Math.max(2, negativeRows.length)}`).format.numberFormat = "#,##0";

const summarySheet = workbook.worksheets.add("Summary");
writeMatrix(summarySheet, 0, 0, [
  ["Tally Physical Stock Sheet - Corrected", ""],
  ["Generated Date", "2026-08-06"],
  ["Source Stock Export", summary.source],
  ["Positive Rows", summary.positive_rows],
  ["Negative Rows", summary.negative_rows],
  ["Rows Flagged For Review", summary.review_rows],
  ["Total Positive PCS", summary.total_positive_pcs],
  ["Total Negative PCS", summary.total_negative_pcs],
  ["Net PCS", summary.net_pcs],
  ["Rows Missing Conversion", summary.missing_conversion_rows],
  ["Correction Note", "Positive rows are still shown for physical checking, but any offsetting negative item/MRP batch is flagged in Remarks and listed separately."],
]);
summarySheet.getRange("A1:B1").merge();
summarySheet.getRange("A1:B1").format = { fill: "#0F766E", font: { bold: true, color: "#FFFFFF", size: 14 } };
summarySheet.getRange("A2:A11").format = { font: { bold: true } };
summarySheet.getRange("B4:B10").format.numberFormat = "#,##0";
summarySheet.showGridLines = false;
setColumnWidths(summarySheet, { A: 30, B: 120 });

const inspect = await workbook.inspect({
  kind: "table",
  sheetId: "Physical Stock Sheet",
  range: "A1:J25",
  include: "values,formulas",
  tableMaxRows: 25,
  tableMaxCols: 10,
});
console.log(inspect.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
});
console.log(errors.ndjson);

await workbook.render({ sheetName: "Physical Stock Sheet", range: "A1:J25", scale: 1, format: "png" });
await workbook.render({ sheetName: "Negative Stock Review", range: "A1:I5", scale: 1, format: "png" });
await workbook.render({ sheetName: "Summary", range: "A1:B11", scale: 1, format: "png" });

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(JSON.stringify({ outputPath, positiveRows: summary.positive_rows, negativeRows: summary.negative_rows }));
