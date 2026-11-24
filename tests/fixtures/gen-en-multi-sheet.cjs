// v3.2.1: writes tests/fixtures/en-multi-sheet.xlsx, the English multi-sheet workbook of the language
// tests (an English interface with a file whose worksheets, headers and data are all English).
// Usage: node tests/fixtures/gen-en-multi-sheet.cjs [output folder] — the default is this folder.
// Uses the repo's own SheetJS; dates are integer serials with a yyyy-mm-dd format (never Date
// objects, which SheetJS writes in the local time zone). Two runs write the same bytes.
const fs = require("node:fs");
const path = require("node:path");
const XLSX = require(path.join(__dirname, "..", "..", "vendor", "xlsx-0.20.3.full.min.js"));

const out = process.argv[2] || __dirname;
const DATE_FORMAT = "yyyy-mm-dd";

function sheet(rows, dateColumn) {
  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  for (let row = 1; row < rows.length; row += 1) {
    const address = XLSX.utils.encode_cell({ r: row, c: dateColumn });
    if (worksheet[address] && worksheet[address].t === "n") worksheet[address].z = DATE_FORMAT;
  }
  return worksheet;
}

const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, sheet([
  ["Order", "Date", "Region", "Amount", "Note"],
  ["O-1001", 45658, "East", 120, ""],
  ["O-1002", 45665, "West", 80, "Rush"],
  ["O-1003", 45672, "East", "N/A", ""],
  ["O-1004", 45679, "North", 200, ""],
  ["O-1005", 45686, "West", 95, "Rush"],
  ["O-1006", 45693, "East", 150, ""]
], 1), "Orders");
XLSX.utils.book_append_sheet(workbook, sheet([
  ["Customer", "Since", "Region"],
  ["C-01", 45292, "East"],
  ["C-02", 45323, "West"],
  ["C-03", 45352, "North"]
], 1), "Customers");
XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([["Topic", "Text"]]), "Notes");

// No document properties with a time stamp: the bytes depend only on the content above.
const bytes = XLSX.write(workbook, { type: "buffer", bookType: "xlsx", bookSST: true, compression: true });
fs.writeFileSync(path.join(out, "en-multi-sheet.xlsx"), bytes);
console.log("en-multi-sheet.xlsx", bytes.length, require("node:crypto").createHash("sha256").update(bytes).digest("hex"));
