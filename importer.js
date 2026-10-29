// 文件导入：CSV / Excel 的读取、编码识别、解析、表头保护、导入上限校验和导入并发保护，最后交给 commitTabularData() 提交。

const ENCODING_LABELS = {
  auto: "自动识别",
  "utf-8": "UTF-8",
  gbk: "GBK",
  gb18030: "GB18030"
};

const IMPORT_LIMITS = Object.freeze({
  maxFileSizeBytes: 25 * 1024 * 1024,
  maxRows: 100000,
  maxColumns: 200
});

async function parseCsvFile(file) {
  const importId = beginImport();
  try {
    assertRuntimeDependency("Papa", "CSV 解析组件加载失败，请刷新页面重试。");
    const selectedEncoding = dom.fileEncoding.value || "auto";
    setStatus(`正在读取 ${file.name} ...`);
    const buffer = await readFileAsArrayBuffer(file, importId);
    if (!isCurrentImport(importId)) return;
    const decoded = decodeCsvBuffer(buffer, selectedEncoding);
    setStatus(`正在解析 ${file.name}（${decoded.label}）...`);
    parseCsvText(decoded.text, file.name, importId);
  } catch (error) {
    showError(error.message, importId);
  }
}

function parseDataFile(file) {
  if (Number.isFinite(file?.size) && file.size > IMPORT_LIMITS.maxFileSizeBytes) {
    const importId = beginImport();
    showError(`文件大小超过 ${formatImportLimitSize(IMPORT_LIMITS.maxFileSizeBytes)}，请拆分文件后重试。`, importId);
    return;
  }
  if (/\.csv$/i.test(file.name)) return parseCsvFile(file);
  if (/\.(xlsx|xls)$/i.test(file.name)) return parseExcelFile(file);
  const importId = beginImport();
  showError("不支持的文件格式。请上传 CSV、XLSX 或 XLS 文件。", importId);
}

async function parseExcelFile(file) {
  const importId = beginImport();
  try {
    assertRuntimeDependency("XLSX", "Excel 解析组件加载失败，请刷新页面重试。");
    setStatus(`正在读取 ${file.name} ...`);
    const buffer = await readFileAsArrayBuffer(file, importId);
    if (!isCurrentImport(importId)) return;
    setStatus(`正在解析 ${file.name} ...`);
    validateExcelFileSignature(buffer, file.name);
    let workbook;
    try {
      workbook = readExcelWorkbook(buffer);
    } catch {
      throw new Error("Excel 文件格式异常或文件已损坏，无法解析。请确认文件是有效的 XLSX 或 XLS 文件。");
    }

    const sheetNames = Array.isArray(workbook.SheetNames) ? workbook.SheetNames.filter(Boolean) : [];
    if (!sheetNames.length) throw new Error("Excel 文件中没有可读取的工作表。");

    if (sheetNames.length > 1) {
      showExcelSheetSelection(workbook, file.name, sheetNames, importId);
      return;
    }

    importExcelSheet(workbook, file.name, sheetNames[0], importId);
  } catch (error) {
    showError(error.message, importId);
  }
}

function readExcelWorkbook(buffer) {
  // cellNF keeps each cell's number format string (cell.z) for format hints.
  return XLSX.read(buffer, { type: "array", cellDates: true, cellNF: true });
}

function showExcelSheetSelection(workbook, fileName, sheetNames, importId) {
  if (!isCurrentImport(importId)) return;
  setImportBusy(false);
  state.pendingExcel = { workbook, fileName, importId };
  dom.excelWorkbookName.textContent = `${fileName} · 共 ${sheetNames.length} 个工作表`;
  dom.excelSheetSelect.innerHTML = "";
  sheetNames.forEach((sheetName) => {
    const option = document.createElement("option");
    option.value = sheetName;
    option.textContent = sheetName;
    dom.excelSheetSelect.appendChild(option);
  });
  dom.sheetSelectionMessage.className = "sheet-selection-message";
  dom.sheetSelectionMessage.textContent = "请选择一个工作表，确认后再生成分析结果。";
  dom.sheetSelectionPanel.classList.remove("hidden");
  dom.statusPanel.className = "status-panel hidden";
  dom.statusPanel.textContent = "";
}

function analyzeSelectedExcelSheet() {
  const pending = state.pendingExcel;
  if (!pending || !isCurrentImport(pending.importId)) {
    showError("Excel 工作簿状态已失效，请重新上传文件。");
    return;
  }

  const sheetName = dom.excelSheetSelect.value;
  if (!sheetName) {
    showSheetSelectionError("请选择一个工作表后再继续。");
    return;
  }

  try {
    importExcelSheet(pending.workbook, pending.fileName, sheetName, pending.importId);
  } catch (error) {
    showSheetSelectionError(error.message);
  }
}

function importExcelSheet(workbook, fileName, sheetName, importId) {
  if (!isCurrentImport(importId)) return;
  const worksheet = workbook.Sheets?.[sheetName];
  if (!worksheet || !worksheet["!ref"]) {
    throw new Error(`工作表「${sheetName}」为空，请选择包含表头和数据的工作表。`);
  }

  const { fields, rows, rowNumbers, numberFormatHints, headerDetection = null } = readExcelSheetTable(worksheet, sheetName);

  state.pendingExcel = null;
  dom.sheetSelectionPanel.classList.add("hidden");
  setStatus(`正在分析 ${fileName} · ${sheetName} ...`);
  commitTabularData(fields, rows, fileName, importId, [], sheetName, numberFormatHints, headerDetection, rowNumbers);
}

// Reads one worksheet into { fields, rows, rowNumbers, numberFormatHints } plus
// headerDetection when title rows above the header were skipped. Row numbers
// follow the !ref start.
function readExcelSheetTable(worksheet, sheetName) {
  validateExcelWorksheetLimits(worksheet, sheetName);

  const rawRows = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: "",
    raw: true,
    blankrows: true
  });
  return buildExcelTabularData(
    rawRows,
    rawRows,
    sheetName,
    createExcelNumberFormatReader(worksheet),
    XLSX.utils.decode_range(worksheet["!ref"]).s.r + 1
  );
}

// Percent hints inspect only the first data rows below the header, so large
// sheets do not pay for a second full cell scan. The cost: a percent format
// that first appears after this many rows is not recognized.
const EXCEL_FORMAT_SAMPLE_ROWS = 1000;

// Returns a reader for one sheet_to_json({ header: 1 }) row index: numeric
// cells get cell.z, every other position stays empty. Formatted text (cell.w)
// is never used because it depends on the runtime time zone and can shift
// dates by one day.
function createExcelNumberFormatReader(worksheet) {
  const range = XLSX.utils.decode_range(worksheet["!ref"]);
  const columnNames = [];
  for (let column = range.s.c; column <= range.e.c; column += 1) {
    columnNames.push(XLSX.utils.encode_col(column));
  }
  return (rowIndex) => {
    const formats = [];
    columnNames.forEach((columnName, columnIndex) => {
      const cell = worksheet[`${columnName}${range.s.r + rowIndex + 1}`];
      if (cell?.t === "n") formats[columnIndex] = String(cell.z ?? "General");
    });
    return formats;
  };
}

function isPercentNumberFormat(format) {
  // Quoted text and backslash-escaped characters are literals, not a percent scale.
  return String(format ?? "").replace(/\\.|"[^"]*"/g, "").includes("%");
}

function buildExcelNumberFormatHints(fields, usedColumnIndexes, dataPairs, readNumberFormats) {
  const sampledFormats = dataPairs
    .slice(0, EXCEL_FORMAT_SAMPLE_ROWS)
    .map((pair) => readNumberFormats(pair.index));
  const hints = Object.create(null);
  fields.forEach((field, fieldIndex) => {
    const columnIndex = usedColumnIndexes[fieldIndex];
    let numericCells = 0;
    let percentCells = 0;
    sampledFormats.forEach((formats) => {
      const format = formats[columnIndex];
      if (typeof format !== "string") return;
      numericCells += 1;
      if (isPercentNumberFormat(format)) percentCells += 1;
    });
    if (percentCells * 2 > numericCells) hints[field] = "percent";
  });
  return hints;
}

// firstRowNumber is the worksheet row number of formattedRows[0] (the !ref start row);
// rowNumbers[i] is the worksheet row of rows[i].
function buildExcelTabularData(formattedRows, rawRows, sheetName = "当前工作表", readNumberFormats = null, firstRowNumber = 1) {
  const rowCount = Math.max(formattedRows?.length || 0, rawRows?.length || 0);
  const rowPairs = Array.from({ length: rowCount }, (_, index) => ({
    index,
    formatted: Array.isArray(formattedRows?.[index]) ? formattedRows[index] : [],
    raw: Array.isArray(rawRows?.[index]) ? rawRows[index] : []
  }));

  while (rowPairs.length && !excelRowHasValue(rowPairs[0].formatted) && !excelRowHasValue(rowPairs[0].raw)) {
    rowPairs.shift();
  }
  while (rowPairs.length && !excelRowHasValue(rowPairs[rowPairs.length - 1].formatted) && !excelRowHasValue(rowPairs[rowPairs.length - 1].raw)) {
    rowPairs.pop();
  }

  if (!rowPairs.length) {
    throw new Error(`工作表「${sheetName}」为空，请选择包含表头和数据的工作表。`);
  }

  const { position: headerPosition, attempted: titleSkipAttempted } = findExcelHeaderPosition(rowPairs);

  // The row limit counts data rows below the header, not skipped title rows.
  if (rowPairs.length - 1 - headerPosition > IMPORT_LIMITS.maxRows) {
    throw new Error(`工作表「${sheetName}」超过 ${formatInteger(IMPORT_LIMITS.maxRows)} 行数据限制，请拆分后重试。`);
  }

  const headerPair = rowPairs[headerPosition];
  const dataPairs = rowPairs.slice(headerPosition + 1);
  const usedColumnIndexes = getExcelUsedColumnIndexes(headerPair, dataPairs);

  if (usedColumnIndexes.length > IMPORT_LIMITS.maxColumns) {
    throw new Error(`工作表「${sheetName}」超过 ${formatInteger(IMPORT_LIMITS.maxColumns)} 列限制，请减少字段后重试。`);
  }

  if (!usedColumnIndexes.length) {
    throw new Error(`工作表「${sheetName}」为空，请选择包含表头和数据的工作表。`);
  }

  const headerCheck = checkExcelHeaderFields(headerPair, usedColumnIndexes);
  if (headerCheck.error) {
    const message = excelHeaderErrorMessage(headerCheck.error, sheetName);
    throw new Error(titleSkipAttempted
      ? `${message}${formatTitleRowFallbackHint(firstRowNumber + headerPair.index)}`
      : message);
  }
  const { fields } = headerCheck;

  const keptPairs = dataPairs
    .filter((pair) => usedColumnIndexes.some((columnIndex) => !isMissing(pair.formatted[columnIndex]) || !isMissing(pair.raw[columnIndex])));
  const rows = keptPairs
    .map((pair) => fields.reduce((row, field, fieldIndex) => {
      const columnIndex = usedColumnIndexes[fieldIndex];
      row[field] = normalizeExcelCellValue(
        pair.formatted[columnIndex],
        pair.raw[columnIndex]
      );
      return row;
    }, Object.create(null)));

  if (!rows.length) {
    throw new Error(`工作表「${sheetName}」只有表头，没有可分析的数据行。`);
  }

  const result = { fields, rows, rowNumbers: keptPairs.map((pair) => firstRowNumber + pair.index) };
  if (headerPosition) {
    // Rows above the header count as headerRowNumber - 1; only non-empty ones are kept.
    result.headerDetection = {
      headerRowNumber: firstRowNumber + headerPair.index,
      skippedRows: rowPairs.slice(0, headerPosition)
        .map((pair) => ({
          rowNumber: firstRowNumber + pair.index,
          cells: pair.formatted.filter((value) => !isMissing(value)).map(displayValue)
        }))
        .filter((row) => row.cells.length)
    };
  }
  if (readNumberFormats) {
    result.numberFormatHints = buildExcelNumberFormatHints(fields, usedColumnIndexes, dataPairs, readNumberFormats);
  }
  return result;
}

// Position in rowPairs of the header row. It stays 0 unless the first non-empty
// row fails the header checks, detectHeaderRow() finds a header below title rows
// within the first rows, and that row passes the same checks against every row.
// attempted tells whether the first row was a skippable title row.
function findExcelHeaderPosition(rowPairs) {
  if (!findExcelHeaderErrorAt(rowPairs, 0)) return { position: 0, attempted: false };
  const sampleRows = rowPairs.slice(0, HEADER_SCAN_ROWS + HEADER_DETECTION_DATA_ROWS).map((pair) => pair.raw);
  const detection = detectHeaderRow(sampleRows, "excel");
  const accepted = detection.status === "skipped" && detection.headerIndex
    && !findExcelHeaderErrorAt(rowPairs, detection.headerIndex);
  return { position: accepted ? detection.headerIndex : 0, attempted: detection.attempted };
}

// Appended to a header error when the first non-empty row looked like a title
// but no header row below it could be confirmed.
function formatTitleRowFallbackHint(rowNumber) {
  return `第 ${rowNumber} 行看起来像标题。如果表头不在第 ${rowNumber} 行，请删除表头上方的标题行后重新上传。`;
}

function findExcelHeaderErrorAt(rowPairs, position) {
  const headerPair = rowPairs[position];
  const usedColumnIndexes = getExcelUsedColumnIndexes(headerPair, rowPairs.slice(position + 1));
  return usedColumnIndexes.length ? checkExcelHeaderFields(headerPair, usedColumnIndexes).error : null;
}

// Columns that hold a header cell or any data value; empty trailing columns drop out.
function getExcelUsedColumnIndexes(headerPair, dataPairs) {
  const maxColumns = [headerPair, ...dataPairs]
    .reduce((max, pair) => Math.max(max, pair.formatted.length, pair.raw.length), 0);
  return Array.from({ length: maxColumns }, (_, index) => index).filter((columnIndex) =>
    !isMissing(headerPair.formatted[columnIndex])
      || dataPairs.some((pair) => !isMissing(pair.formatted[columnIndex]) || !isMissing(pair.raw[columnIndex]))
  );
}

// Excel header rules: every used column needs a non-empty text header, and the
// trimmed names must be unique. Returns { fields } or { error }, checked column by column.
function checkExcelHeaderFields(headerPair, usedColumnIndexes) {
  const fields = [];
  for (const columnIndex of usedColumnIndexes) {
    const formattedHeader = headerPair.formatted[columnIndex];
    if (isMissing(formattedHeader)) return { error: { code: "missingHeader", columnIndex } };
    if (typeof headerPair.raw[columnIndex] !== "string") return { error: { code: "nonTextHeader", columnIndex } };
    fields.push(normalizeFieldName(formattedHeader));
  }
  const duplicateFields = fields.filter((field, index) => fields.indexOf(field) !== index);
  if (duplicateFields.length) return { error: { code: "duplicateHeader", field: duplicateFields[0] } };
  return { fields };
}

function excelHeaderErrorMessage(error, sheetName) {
  if (error.code === "missingHeader") {
    return `工作表「${sheetName}」的第 ${error.columnIndex + 1} 列缺少表头，请补充字段名后重新上传。`;
  }
  if (error.code === "nonTextHeader") {
    return `工作表「${sheetName}」未识别到有效文本表头；第 ${error.columnIndex + 1} 列首行不是字段名。`;
  }
  return `工作表「${sheetName}」包含重复表头「${error.field}」，请修改后重新上传。`;
}

function excelRowHasValue(row) {
  return Array.isArray(row) && row.some((value) => !isMissing(value));
}

function normalizeExcelDateValue(value) {
  const date = toDate(value);
  return date ? date.toISOString().slice(0, 10) : null;
}

function normalizeExcelCellValue(formattedValue, rawValue) {
  if (isDateObject(rawValue)) {
    return normalizeExcelDateValue(rawValue) || "";
  }
  if (["string", "number", "boolean", "bigint"].includes(typeof rawValue)) {
    return String(rawValue);
  }
  return formattedValue === null || formattedValue === undefined
    ? ""
    : String(formattedValue);
}

function validateExcelFileSignature(buffer, fileName) {
  const bytes = new Uint8Array(buffer);
  const isXlsx = /\.xlsx$/i.test(fileName);
  const isXls = /\.xls$/i.test(fileName);
  const hasZipSignature = bytes.length >= 4
    && bytes[0] === 0x50
    && bytes[1] === 0x4b
    && [[0x03, 0x04], [0x05, 0x06], [0x07, 0x08]]
      .some(([third, fourth]) => bytes[2] === third && bytes[3] === fourth);
  const oleSignature = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  const hasOleSignature = bytes.length >= oleSignature.length
    && oleSignature.every((value, index) => bytes[index] === value);

  if ((isXlsx && !hasZipSignature) || (isXls && !hasOleSignature)) {
    throw new Error("Excel 文件格式与扩展名不匹配或文件已损坏。请重新导出为有效的 XLSX 或 XLS 文件后再上传。");
  }
}

function validateExcelWorksheetLimits(worksheet, sheetName) {
  let range;
  try {
    range = XLSX.utils.decode_range(worksheet["!ref"]);
  } catch {
    throw new Error(`工作表「${sheetName}」的使用区域无效，无法解析。`);
  }

  const dataRowCount = range.e.r - range.s.r;
  const columnCount = range.e.c - range.s.c + 1;
  // An early, coarse check: title rows above the header may still be skipped, so
  // allow HEADER_SCAN_ROWS extra rows; buildExcelTabularData() applies the exact limit.
  if (dataRowCount > IMPORT_LIMITS.maxRows + HEADER_SCAN_ROWS) {
    throw new Error(`工作表「${sheetName}」超过 ${formatInteger(IMPORT_LIMITS.maxRows)} 行数据限制，请拆分后重试。`);
  }
  if (columnCount > IMPORT_LIMITS.maxColumns) {
    throw new Error(`工作表「${sheetName}」超过 ${formatInteger(IMPORT_LIMITS.maxColumns)} 列限制，请减少字段后重试。`);
  }
}

function showSheetSelectionError(message) {
  dom.sheetSelectionMessage.className = "sheet-selection-message error";
  dom.sheetSelectionMessage.textContent = message;
}

// retry is set only for the one re-parse below skipped title rows:
// { lineOffset, headerDetection, firstError }. wholeText drops the row preview, for
// the one re-parse after blank lines filled it (see handleParsedData()).
function parseCsvText(text, sourceName, importId = beginImport(), retry = null, wholeText = false) {
  if (!isCurrentImport(importId)) return;
  assertRuntimeDependency("Papa", "CSV 解析组件加载失败，请刷新页面重试。");
  Papa.parse(text, {
    header: true,
    skipEmptyLines: true,
    ...(wholeText ? {} : { preview: IMPORT_LIMITS.maxRows + 1 }),
    // PapaParse workers cannot clone transformHeader functions. Keep parsing on
    // the main thread so protected and duplicate header handling remains intact.
    worker: false,
    transformHeader: encodeCsvHeaderForParser,
    complete: (results) => handleParsedData(results, sourceName, importId, { text, retry, wholeText }),
    error: (error) => showError(retry ? retry.firstError : error.message, importId)
  });
}

// Finds the header below title rows of a CSV whose first row failed the header
// checks. Returns { offset, lineOffset, headerDetection, titleHint }: offset is the
// header row's first character and lineOffset the physical lines above it.
// headerDetection is null when no header was found; titleHint is the fallback
// hint for the first non-empty row, or "" when skipping was not tried.
function findCsvHeaderBelowTitleRows(text) {
  const rows = [];
  const starts = [];
  let cursor = 0;
  Papa.parse(text, {
    header: false,
    skipEmptyLines: false,
    preview: HEADER_SCAN_ROWS + HEADER_DETECTION_DATA_ROWS,
    worker: false,
    step: (results) => {
      rows.push(results.data);
      starts.push(cursor);
      cursor = results.meta.cursor;
    }
  });
  const detection = detectHeaderRow(rows, "csv");
  const lineNumbers = lineNumbersAtOffsets(text, starts);
  const lineNumberAt = (index) => lineNumbers[index] ?? 1;
  const firstIndex = Math.max(0, rows.findIndex((row) => countNonEmptyCells(row) > 0));
  const titleHint = detection.attempted ? formatTitleRowFallbackHint(lineNumberAt(firstIndex)) : "";
  if (detection.status !== "skipped" || !detection.headerIndex) return { headerDetection: null, titleHint };
  const headerRowNumber = lineNumberAt(detection.headerIndex);
  return {
    titleHint,
    offset: starts[detection.headerIndex],
    lineOffset: headerRowNumber - 1,
    headerDetection: {
      headerRowNumber,
      skippedRows: detection.skippedIndexes
        .map((index) => ({
          rowNumber: lineNumberAt(index),
          cells: rows[index].filter((value) => !isMissing(value)).map(displayValue)
        }))
        .filter((row) => row.cells.length)
    }
  };
}

// File line (1-based) at each text offset: the line breaks before it plus one,
// where "\r\n", a lone "\r" and a lone "\n" each end one line. Counting carries on
// from the previous offset, so ascending offsets cost one pass over the text.
function lineNumbersAtOffsets(text, offsets) {
  let position = 0;
  let breaks = 0;
  return offsets.map((offset) => {
    if (offset < position) {
      position = 0;
      breaks = 0;
    }
    for (; position < offset; position += 1) {
      const code = text.charCodeAt(position);
      // A "\n" right after "\r" closes the same line break.
      if (code === 13 || (code === 10 && text.charCodeAt(position - 1) !== 13)) breaks += 1;
    }
    return breaks + 1;
  });
}

// File line where each data record of a parsed CSV starts. A second PapaParse pass
// (header:false, synchronous step) reads the same text with the main parse's
// delimiter and line break; a record starts at the previous step's cursor.
// Records that skipEmptyLines:true drops (a single empty cell) are dropped here
// too, before the first remaining one is taken as the header, as PapaParse does.
// Returns null when the record count differs from dataRowCount.
function findCsvRecordLineNumbers(text, meta, dataRowCount) {
  const starts = [];
  let cursor = 0;
  Papa.parse(text, {
    header: false,
    skipEmptyLines: false,
    worker: false,
    delimiter: meta?.delimiter,
    newline: meta?.linebreak,
    step: (results) => {
      const record = results.data;
      if (!(record.length === 1 && record[0] === "")) starts.push(cursor);
      cursor = results.meta.cursor;
    }
  });
  starts.shift();
  return starts.length === dataRowCount ? lineNumbersAtOffsets(text, starts) : null;
}

// Whether text holds more than limit data records: non-empty records (those
// skipEmptyLines:true keeps) after the first one, the header. A header:false pass
// with the main parse's delimiter and line break counts them without keeping any
// row and stops once the count is past the limit.
function csvHasMoreDataRecordsThan(text, meta, limit) {
  let records = 0;
  Papa.parse(text, {
    header: false,
    skipEmptyLines: false,
    worker: false,
    delimiter: meta?.delimiter,
    newline: meta?.linebreak,
    step: (results, parser) => {
      const record = results.data;
      if (record.length === 1 && record[0] === "") return;
      records += 1;
      if (records > limit + 1) parser.abort();
    }
  });
  return records > limit + 1;
}

// On a CSV header error, re-parse once from the detected header row. Returns the
// error to show now (with the title hint when skipping was tried), or "" when the
// re-parse took over; a failed re-parse then shows that same first error.
function retryCsvBelowTitleRows(source, sourceName, importId, message) {
  if (!source || source.retry) return message;
  const found = findCsvHeaderBelowTitleRows(source.text);
  const firstError = `${message}${found.titleHint}`;
  if (!found.headerDetection) return firstError;
  parseCsvText(source.text.slice(found.offset), sourceName, importId, {
    lineOffset: found.lineOffset,
    headerDetection: found.headerDetection,
    firstError
  });
  return "";
}

const CSV_HEADER_KEY_PREFIX = "__SMART_CSV_HEADER__";

function encodeCsvHeaderForParser(header, index) {
  return `${CSV_HEADER_KEY_PREFIX}${index}:${encodeURIComponent(String(header ?? ""))}`;
}

function decodeCsvHeaderFromParser(header) {
  const value = String(header ?? "");
  if (!value.startsWith(CSV_HEADER_KEY_PREFIX)) return value;
  const separatorIndex = value.indexOf(":", CSV_HEADER_KEY_PREFIX.length);
  if (separatorIndex < 0) return value;
  try {
    return decodeURIComponent(value.slice(separatorIndex + 1));
  } catch {
    return value;
  }
}

function beginImport() {
  state.activeImportId += 1;
  abortActiveFileRead();
  resetAnalysisState();
  setImportBusy(true);
  return state.activeImportId;
}

function isCurrentImport(importId) {
  return importId === state.activeImportId;
}

function assertRuntimeDependency(globalName, message) {
  if (!window[globalName]) throw new Error(message);
}

function abortActiveFileRead() {
  const reader = state.activeFileReader;
  state.activeFileReader = null;
  if (reader && reader.readyState === 1 && typeof reader.abort === "function") reader.abort();
}

function readFileAsArrayBuffer(file, importId = state.activeImportId) {
  if (typeof FileReader === "undefined" && typeof file.arrayBuffer === "function") {
    return file.arrayBuffer();
  }

  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    state.activeFileReader = reader;
    const clearReader = () => {
      if (state.activeFileReader === reader && isCurrentImport(importId)) {
        state.activeFileReader = null;
      }
    };
    reader.onload = () => {
      clearReader();
      resolve(reader.result);
    };
    reader.onerror = () => {
      clearReader();
      reject(reader.error || new Error("文件读取失败"));
    };
    reader.onabort = () => {
      clearReader();
      reject(new Error("文件读取已取消"));
    };
    reader.readAsArrayBuffer(file);
  });
}

function decodeCsvBuffer(buffer, selectedEncoding) {
  if (selectedEncoding === "auto") {
    return autoDecodeCsvBuffer(buffer);
  }

  if (!Object.prototype.hasOwnProperty.call(ENCODING_LABELS, selectedEncoding)) {
    throw new Error("不支持的文件编码。");
  }

  return {
    text: decodeText(buffer, selectedEncoding),
    label: ENCODING_LABELS[selectedEncoding]
  };
}

function autoDecodeCsvBuffer(buffer) {
  if (hasUtf8Bom(buffer)) {
    return {
      text: decodeText(buffer, "utf-8"),
      label: "UTF-8"
    };
  }

  try {
    return {
      text: decodeText(buffer, "utf-8", true),
      label: "UTF-8"
    };
  } catch {
    return {
      text: decodeText(buffer, "gb18030"),
      label: "GB18030 自动识别"
    };
  }
}

function hasUtf8Bom(buffer) {
  const bytes = new Uint8Array(buffer);
  return bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf;
}

function decodeText(buffer, encoding, fatal = false) {
  try {
    return new TextDecoder(encoding, { fatal }).decode(buffer);
  } catch (error) {
    if (fatal) throw error;
    throw new Error(`当前浏览器不支持 ${ENCODING_LABELS[encoding] || encoding} 解码，请尝试选择其他编码或将文件另存为 CSV UTF-8。`);
  }
}

// source is { text, retry } from parseCsvText(); without it no header retry runs.
function handleParsedData(results, sourceName, importId, source = null) {
  if (!isCurrentImport(importId)) return;
  const retry = source?.retry || null;
  // Physical lines skipped above the header, so row numbers match the file.
  const lineOffset = retry?.lineOffset || 0;
  // A failed re-parse below skipped title rows shows the first parse's error.
  const fail = (message) => showError(retry ? retry.firstError : message, importId);
  const failHeader = (message) => {
    const error = retryCsvBelowTitleRows(source, sourceName, importId, message);
    if (error) fail(error);
  };
  const parseErrors = Array.isArray(results.errors) ? results.errors : [];
  const extraFieldError = parseErrors.find((error) => error.code === "TooManyFields");
  if (extraFieldError) {
    const rowNumber = Number.isInteger(extraFieldError.row) ? extraFieldError.row + 2 + lineOffset : null;
    failHeader(`CSV${rowNumber ? ` 第 ${rowNumber} 行` : ""}的数据列多于表头，继续分析会丢失额外单元格。请补全表头或修正该行后重试。`);
    return;
  }
  const fatalError = parseErrors.find((error) =>
    error.code === "MissingQuotes"
    || (error.type === "Delimiter" && error.code !== "UndetectableDelimiter")
  );
  if (fatalError) {
    fail(`CSV 解析失败：${fatalError.message}`);
    return;
  }

  const allFieldPairs = (results.meta?.fields || []).map((rawField) => ({
    raw: rawField,
    clean: normalizeFieldName(decodeCsvHeaderFromParser(rawField))
  }));
  if (allFieldPairs.length > IMPORT_LIMITS.maxColumns) {
    fail(`CSV 超过 ${formatInteger(IMPORT_LIMITS.maxColumns)} 列限制，请减少字段后重试。`);
    return;
  }
  const fieldPairs = allFieldPairs.filter((pair) => pair.clean !== "");
  const fields = fieldPairs.map((pair) => pair.clean);
  const headerNames = allFieldPairs.map((pair) => pair.clean);
  const duplicateFields = findDuplicateCsvHeaders(headerNames);

  if (duplicateFields.length) {
    failHeader(`CSV 表头规范化后存在重复字段：${duplicateFields.join("、")}。请修改表头后重试。`);
    return;
  }

  const sourceRows = Array.isArray(results.data) ? results.data : [];
  const rowLimitError = `CSV 超过 ${formatInteger(IMPORT_LIMITS.maxRows)} 行数据限制，请拆分文件后重试。`;
  if (sourceRows.length > IMPORT_LIMITS.maxRows) {
    fail(rowLimitError);
    return;
  }
  // The row preview counts blank records too: when they filled it, the parse stopped
  // early under the limit. Count the data records, then fail as over the limit or
  // parse the whole text once more.
  if (results.meta?.truncated && source && !source.wholeText) {
    if (csvHasMoreDataRecordsThan(source.text, results.meta, IMPORT_LIMITS.maxRows)) fail(rowLimitError);
    else parseCsvText(source.text, sourceName, importId, retry, true);
    return;
  }
  const blankHeaderHasData = csvBlankHeaderHasData(
    headerNames,
    sourceRows,
    (row, columnIndex) => row?.[allFieldPairs[columnIndex].raw]
  );
  if (blankHeaderHasData) {
    failHeader("CSV 存在缺少表头的数据列。请补全每一列的字段名后重试。");
    return;
  }

  const parseWarnings = parseErrors
    .filter((error) => !(error.code === "UndetectableDelimiter" && fields.length === 1))
    .slice(0, 20)
    .map((error) => ({
      row: Number.isInteger(error.row) ? error.row + 2 + lineOffset : null,
      code: error.code || "ParseWarning",
      message: error.message || "CSV 行结构异常"
    }));
  // Indexes into sourceRows of the kept rows, so row numbers drop with blank rows.
  const keptIndexes = [];
  const rows = sourceRows.reduce((normalizedRows, row, index) => {
    const normalized = normalizeParsedRow(row, fieldPairs);
    if (fields.some((field) => !isMissing(normalized[field]))) {
      normalizedRows.push(normalized);
      keptIndexes.push(index);
    }
    return normalizedRows;
  }, []);

  if (!fields.length || !rows.length) {
    fail("CSV 中没有可分析的数据。请确认文件包含表头和至少一行数据。");
    return;
  }

  // Only a file about to be committed gets the line-number pass; a re-parse below
  // skipped title rows adds the physical lines above its header.
  const recordLineNumbers = source ? findCsvRecordLineNumbers(source.text, results.meta, sourceRows.length) : null;
  const rowNumbers = recordLineNumbers ? keptIndexes.map((index) => recordLineNumbers[index] + lineOffset) : null;
  commitTabularData(fields, rows, sourceName, importId, parseWarnings, "", null, retry?.headerDetection || null, rowNumbers);
}

// Header names that repeat after normalizing; blank names are never compared.
function findDuplicateCsvHeaders(headerNames) {
  const fields = headerNames.filter((name) => name !== "");
  return [...new Set(fields.filter((field, index) => fields.indexOf(field) !== index))];
}

// Whether a column with a blank header holds data. readCell(row, columnIndex)
// reads one cell, so parsed row objects and plain arrays share the rule.
function csvBlankHeaderHasData(headerNames, dataRows, readCell) {
  return headerNames.some((name, columnIndex) => name === ""
    && dataRows.some((row) => !isMissing(readCell(row, columnIndex))));
}

// Header detection scans at most this many rows, starting at the first non-empty row.
const HEADER_SCAN_ROWS = 10;
// Detection reads only this many rows beyond the scan; the full checks then see every row.
const HEADER_DETECTION_DATA_ROWS = 200;
// A column below a header candidate counts as numeric or date data from this share on.
const HEADER_GUARD_TYPE_RATIO = 0.8;

// Picks the header row when title rows sit above it. rows is a 2-D array: Excel
// raw values, or CSV strings from PapaParse header:false. Pure: no DOM, no state.
// - "unchanged": rows[0] passes the header checks and stays the header;
// - "skipped": rows[headerIndex] is the header and every row above it is skipped;
// - "fallback": keep the first non-empty row; attempted tells whether that row
//   was a skippable title row, i.e. skipping was actually tried.
function detectHeaderRow(rows, kind) {
  const matrix = (Array.isArray(rows) ? rows : []).map((row) => (Array.isArray(row) ? row : []));
  const firstIndex = Math.max(0, matrix.findIndex((row) => countNonEmptyCells(row) > 0));
  const result = (status, headerIndex, attempted) => ({
    status,
    headerIndex,
    skippedIndexes: status === "skipped" ? Array.from({ length: headerIndex }, (_, index) => index) : [],
    attempted
  });

  if (!findHeaderRowError(matrix, firstIndex, kind)) {
    return result(firstIndex === 0 ? "unchanged" : "skipped", firstIndex, false);
  }
  if (!isSkippableHeaderRow(matrix[firstIndex])) return result("fallback", firstIndex, false);

  const lastIndex = Math.min(matrix.length, firstIndex + HEADER_SCAN_ROWS) - 1;
  for (let index = firstIndex + 1; index <= lastIndex; index += 1) {
    if (isSkippableHeaderRow(matrix[index])) continue;
    const accepted = !findHeaderRowError(matrix, index, kind)
      && modeNonEmptyCount(matrix.slice(index + 1)) >= 2
      && !vetoesHeaderCandidate(matrix, index);
    return accepted ? result("skipped", index, true) : result("fallback", firstIndex, true);
  }
  return result("fallback", firstIndex, true);
}

// The import's own header checks, applied with rows[headerIndex] as the header.
// Returns an error code, or "" when the row would pass.
function findHeaderRowError(rows, headerIndex, kind) {
  const header = rows[headerIndex] || [];
  const dataRows = rows.slice(headerIndex + 1);
  if (kind === "excel") {
    const toPair = (row) => ({ formatted: row, raw: row });
    const headerPair = toPair(header);
    const dataPairs = dataRows.map(toPair);
    const usedColumnIndexes = getExcelUsedColumnIndexes(headerPair, dataPairs);
    if (!usedColumnIndexes.length) return "emptySheet";
    return checkExcelHeaderFields(headerPair, usedColumnIndexes).error?.code || "";
  }
  // PapaParse header mode reports TooManyFields for a row longer than the header.
  if (dataRows.some((row) => row.length > header.length)) return "tooManyFields";
  const headerNames = header.map(normalizeFieldName);
  if (findDuplicateCsvHeaders(headerNames).length) return "duplicateHeader";
  if (csvBlankHeaderHasData(headerNames, dataRows, (row, columnIndex) => row[columnIndex])) return "missingHeader";
  return "";
}

// Rows a title area may hold: blank rows (CSV ",,," included), one text cell, or
// two text cells where one is a "label：value" note. Numbers and dates are not text.
function isSkippableHeaderRow(row) {
  const values = (row || []).filter((value) => !isMissing(value));
  if (values.length > 2 || !values.every(isHeaderTextValue)) return false;
  return values.length < 2 || values.some((value) => /[:：]/.test(value));
}

function isHeaderTextValue(value) {
  return typeof value === "string" && toNumber(value) === null && toDate(value) === null;
}

function countNonEmptyCells(row) {
  return (row || []).filter((value) => !isMissing(value)).length;
}

// Most common non-empty cell count among non-blank rows; a tie takes the smaller count.
function modeNonEmptyCount(rows) {
  const frequencies = new Map();
  rows.forEach((row) => {
    const count = countNonEmptyCells(row);
    if (count) frequencies.set(count, (frequencies.get(count) || 0) + 1);
  });
  let mode = 0;
  let modeFrequency = 0;
  frequencies.forEach((frequency, count) => {
    if (frequency > modeFrequency || (frequency === modeFrequency && count < mode)) {
      mode = count;
      modeFrequency = frequency;
    }
  });
  return mode;
}

// Rejects a data row taken for the header: a candidate cell that reads as a number
// above a mostly numeric column, as a date above a mostly date column, or whose
// trimmed value appears again in its column below.
function vetoesHeaderCandidate(rows, headerIndex) {
  const below = rows.slice(headerIndex + 1);
  return rows[headerIndex].some((cell, columnIndex) => {
    if (isMissing(cell)) return false;
    const values = below.map((row) => row[columnIndex]).filter((value) => !isMissing(value));
    if (!values.length) return false;
    const mostly = (convert) => (
      values.filter((value) => convert(value) !== null).length >= values.length * HEADER_GUARD_TYPE_RATIO
    );
    if (toNumber(cell) !== null && mostly(toNumber)) return true;
    if (toDate(cell) !== null && mostly((value) => toDate(value))) return true;
    const text = normalizeValue(cell);
    return values.some((value) => normalizeValue(value) === text);
  });
}

function normalizeFieldName(field) {
  return String(field ?? "").replace(/^\uFEFF/, "").trim();
}

function normalizeParsedRow(row, fieldPairs) {
  return fieldPairs.reduce((normalized, pair) => {
    if (Object.prototype.hasOwnProperty.call(row, pair.raw)) {
      normalized[pair.clean] = row[pair.raw];
    } else if (Object.prototype.hasOwnProperty.call(row, pair.clean)) {
      normalized[pair.clean] = row[pair.clean];
    } else {
      // A short row omits trailing cells; never read inherited prototype members.
      normalized[pair.clean] = undefined;
    }
    return normalized;
  }, Object.create(null));
}

function formatImportLimitSize(bytes) {
  return `${Math.round(bytes / (1024 * 1024))} MB`;
}
