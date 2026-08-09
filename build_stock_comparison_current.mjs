import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const downloads = "C:/Users/goenk/Documents/Codex/2026-05-12/Goldiee Purchase Bot/downloads";
const data = JSON.parse(await fs.readFile(`${downloads}/stock_comparison_data_current.json`, "utf8"));

function parseTsv(text) {
  const lines = text.replace(/^\uFEFF/, "").trim().split(/\r?\n/);
  const headers = lines[0].split("\t");
  return lines.slice(1).map((line) => {
    const values = line.split("\t");
    return Object.fromEntries(headers.map((header, i) => [header, values[i] ?? ""]));
  });
}

function toNumber(value) {
  if (value === null || value === undefined || value === "") return "";
  const numeric = Number(String(value).replaceAll(",", ""));
  return Number.isFinite(numeric) ? numeric : "";
}

const tally = parseTsv(await fs.readFile(`${downloads}/actual_tally_stock_current.tsv`, "utf8"));
const dms = parseTsv(await fs.readFile(`${downloads}/dms_stock_current.tsv`, "utf8"));
const wb = Workbook.create();

const titleFill = "#17365D";
const headerFill = "#1F4E78";
const greenFill = "#C6EFCE";
const redFill = "#FFC7CE";
const paleBlue = "#D9EAF7";
const borderColor = "#B7C9D6";

function styleHeader(range) {
  range.format.fill = headerFill;
  range.format.font = { bold: true, color: "#FFFFFF" };
  range.format.wrapText = true;
  range.format.verticalAlignment = "center";
}

function addTitle(sheet, title, endColumn) {
  sheet.mergeCells(`A1:${endColumn}1`);
  sheet.getRange("A1").values = [[title]];
  const titleRange = sheet.getRange(`A1:${endColumn}1`);
  titleRange.format.fill = titleFill;
  titleRange.format.font = { bold: true, color: "#FFFFFF", size: 16 };
  titleRange.format.rowHeight = 28;
}

const summary = wb.worksheets.add("Summary");
addTitle(summary, "Tally vs DMS Current Stock Comparison", "F");
summary.getRange("A3:B15").values = [
  ["Metric", "Value"],
  ["Generated at", data.summary.generated_at],
  ["Comparison basis", data.summary.comparison_basis],
  ["Tally item/MRP combinations", data.summary.tally_combinations],
  ["DMS item/MRP combinations", data.summary.dms_combinations],
  ["Comparison rows", data.summary.comparison_rows],
  ["Exact matches", data.summary.exact_matches],
  ["Mismatches / review rows", data.summary.mismatches],
  ["Exact match percentage", data.summary.exact_matches / data.summary.comparison_rows],
  ["Tally detail rows", data.summary.tally_detail_rows],
  ["DMS detail rows", data.summary.dms_detail_rows],
  ["Tally source", data.summary.tally_source],
  ["DMS source", data.summary.dms_source],
];
styleHeader(summary.getRange("A3:B3"));
summary.getRange("B11").format.numberFormat = "0.00%";
summary.getRange("D3:F8").values = [
  ["Legend", "", ""],
  ["Green", "Exact item name, MRP, and quantity match", ""],
  ["Red", "Quantity mismatch, MRP/item presence difference, or review needed", ""],
  ["Tally qty", "Net quantity across all nonzero PTR/batch rows, including negative rows", ""],
  ["DMS qty", "Total Current Stock in Pcs from DMS current-stock report", ""],
  ["Batch display", "Tally shows PTR/batch with qty; DMS shows batch/MRP bucket with qty", ""],
];
styleHeader(summary.getRange("D3:F3"));
summary.getRange("D4:F4").format.fill = greenFill;
summary.getRange("D5:F5").format.fill = redFill;
summary.getRange("D6:F8").format.fill = paleBlue;
summary.getRange("A3:F15").format.borders = { preset: "all", style: "thin", color: borderColor };
summary.getRange("A1:F15").format.wrapText = true;
summary.getRange("A:A").format.columnWidth = 34;
summary.getRange("B:B").format.columnWidth = 76;
summary.getRange("C:C").format.columnWidth = 3;
summary.getRange("D:D").format.columnWidth = 16;
summary.getRange("E:F").format.columnWidth = 45;

const comp = wb.worksheets.add("Stock Comparison");
addTitle(comp, "Stock Comparison by Exact Item Name and MRP", "J");
const compHeaders = [
  "Status",
  "Item Name",
  "MRP",
  "Tally Qty (PCS)",
  "DMS Qty (PCS)",
  "Difference (PCS)",
  "Tally PTR / Batches",
  "DMS Batch / MRP Buckets",
  "Remarks",
  "Review",
];
comp.getRange("A3:J3").values = [compHeaders];
styleHeader(comp.getRange("A3:J3"));
const compValues = data.rows.map((r) => [
  r.status,
  r.item,
  r.mrp ?? "",
  r.tally_qty_pcs,
  r.dms_qty_pcs,
  r.difference_pcs,
  r.tally_ptr_batches,
  r.dms_mrp_batches,
  r.remarks,
  "",
]);
comp.getRange(`A4:J${compValues.length + 3}`).values = compValues;
for (let i = 0; i < data.rows.length; i++) {
  const row = i + 4;
  comp.getRange(`A${row}:J${row}`).format.fill =
    data.rows[i].status === "EXACT MATCH" ? greenFill : redFill;
}
comp.getRange(`C4:F${compValues.length + 3}`).format.numberFormat = "0.00";
comp.getRange(`A3:J${compValues.length + 3}`).format.wrapText = true;
comp.getRange(`A3:J${compValues.length + 3}`).format.borders = {
  preset: "all",
  style: "thin",
  color: borderColor,
};
comp.freezePanes.freezeRows(3);
comp.getRange("A:A").format.columnWidth = 18;
comp.getRange("B:B").format.columnWidth = 48;
comp.getRange("C:F").format.columnWidth = 16;
comp.getRange("G:H").format.columnWidth = 34;
comp.getRange("I:I").format.columnWidth = 64;
comp.getRange("J:J").format.columnWidth = 18;
comp.tables.add(`A3:J${compValues.length + 3}`, true, "StockComparisonTable");

const tallySheet = wb.worksheets.add("Tally Current Stock");
addTitle(tallySheet, "Actual Tally Stock - Batch/PTR Detail", "H");
const tallyHeaders = [
  "Item Name",
  "PTR / Batch",
  "MRP",
  "Quantity",
  "Unit",
  "Closing Rate",
  "Closing Value",
  "MRP Source Date",
];
tallySheet.getRange("A3:H3").values = [tallyHeaders];
styleHeader(tallySheet.getRange("A3:H3"));
tallySheet.getRange(`A4:H${tally.length + 3}`).values = tally.map((r) => [
  r.item,
  r.ptr_batch,
  toNumber(r.mrp),
  toNumber(r.qty),
  r.unit,
  toNumber(r.closing_rate),
  toNumber(r.closing_value),
  r.mrp_source_date,
]);
tallySheet.getRange(`C4:G${tally.length + 3}`).format.numberFormat = "0.00";
tallySheet.getRange(`A3:H${tally.length + 3}`).format.wrapText = true;
tallySheet.getRange(`A3:H${tally.length + 3}`).format.borders = {
  preset: "all",
  style: "thin",
  color: borderColor,
};
tallySheet.freezePanes.freezeRows(3);
tallySheet.getRange("A:A").format.columnWidth = 52;
tallySheet.getRange("B:B").format.columnWidth = 18;
tallySheet.getRange("C:G").format.columnWidth = 16;
tallySheet.getRange("H:H").format.columnWidth = 19;
tallySheet.tables.add(`A3:H${tally.length + 3}`, true, "TallyCurrentStockTable");

const dmsSheet = wb.worksheets.add("DMS Current Stock");
addTitle(dmsSheet, "DMS Current Stock - Detailed Inventory Rows", "I");
const dmsHeaders = [
  "ERP ID",
  "Item Name",
  "MRP Batch No",
  "MRP",
  "Current Qty (PCS)",
  "Cases",
  "Pieces",
  "Purchase Price",
  "Net Value",
];
dmsSheet.getRange("A3:I3").values = [dmsHeaders];
styleHeader(dmsSheet.getRange("A3:I3"));
dmsSheet.getRange(`A4:I${dms.length + 3}`).values = dms.map((r) => [
  r.erp_id,
  r.item,
  r.mrp_batch,
  toNumber(r.mrp),
  toNumber(r.qty),
  toNumber(r.cases),
  toNumber(r.pieces),
  toNumber(r.purchase_price),
  toNumber(r.net_value),
]);
dmsSheet.getRange(`D4:I${dms.length + 3}`).format.numberFormat = "0.00";
dmsSheet.getRange(`A3:I${dms.length + 3}`).format.wrapText = true;
dmsSheet.getRange(`A3:I${dms.length + 3}`).format.borders = {
  preset: "all",
  style: "thin",
  color: borderColor,
};
dmsSheet.freezePanes.freezeRows(3);
dmsSheet.getRange("A:A").format.columnWidth = 16;
dmsSheet.getRange("B:B").format.columnWidth = 52;
dmsSheet.getRange("C:C").format.columnWidth = 24;
dmsSheet.getRange("D:I").format.columnWidth = 17;
dmsSheet.tables.add(`A3:I${dms.length + 3}`, true, "DmsCurrentStockTable");

for (const sheet of [summary, comp, tallySheet, dmsSheet]) {
  sheet.showGridlines = false;
}

const summaryInspect = await wb.inspect({
  kind: "table",
  sheetId: "Summary",
  range: "A1:F15",
  include: "values,formulas",
  tableMaxRows: 15,
  tableMaxCols: 6,
});
console.log(summaryInspect.ndjson);

const comparisonInspect = await wb.inspect({
  kind: "table",
  sheetId: "Stock Comparison",
  range: "A3:J12",
  include: "values,formulas",
  tableMaxRows: 10,
  tableMaxCols: 10,
});
console.log(comparisonInspect.ndjson);

const errorInspect = await wb.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "final formula error scan",
});
console.log(errorInspect.ndjson);

const summaryPreview = await wb.render({
  sheetName: "Summary",
  autoCrop: "all",
  scale: 1,
  format: "png",
});
await fs.writeFile(
  `${downloads}/Tally vs DMS Stock Comparison 2026-08-06 Summary Preview.png`,
  new Uint8Array(await summaryPreview.arrayBuffer()),
);

const comparisonPreview = await wb.render({
  sheetName: "Stock Comparison",
  range: "A1:J22",
  scale: 1,
  format: "png",
});
await fs.writeFile(
  `${downloads}/Tally vs DMS Stock Comparison 2026-08-06 Comparison Preview.png`,
  new Uint8Array(await comparisonPreview.arrayBuffer()),
);

const output = await SpreadsheetFile.exportXlsx(wb);
const outPath = `${downloads}/Tally vs DMS Stock Comparison 2026-08-06.xlsx`;
await output.save(outPath);
console.log(outPath);
