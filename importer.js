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

  validateExcelWorksheetLimits(worksheet, sheetName);

  const rawRows = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: "",
    raw: true,
    blankrows: true
  });
  const { fields, rows, numberFormatHints } = buildExcelTabularData(
    rawRows,
    rawRows,
    sheetName,
    createExcelNumberFormatReader(worksheet)
  );

  state.pendingExcel = null;
  dom.sheetSelectionPanel.classList.add("hidden");
  setStatus(`正在分析 ${fileName} · ${sheetName} ...`);
  commitTabularData(fields, rows, fileName, importId, [], sheetName, numberFormatHints);
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

function buildExcelTabularData(formattedRows, rawRows, sheetName = "当前工作表", readNumberFormats = null) {
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

  if (rowPairs.length - 1 > IMPORT_LIMITS.maxRows) {
    throw new Error(`工作表「${sheetName}」超过 ${formatInteger(IMPORT_LIMITS.maxRows)} 行数据限制，请拆分后重试。`);
  }

  const headerPair = rowPairs[0];
  const dataPairs = rowPairs.slice(1);
  const maxColumns = rowPairs.reduce((max, pair) => Math.max(max, pair.formatted.length, pair.raw.length), 0);
  const usedColumnIndexes = Array.from({ length: maxColumns }, (_, index) => index).filter((columnIndex) =>
    !isMissing(headerPair.formatted[columnIndex])
      || dataPairs.some((pair) => !isMissing(pair.formatted[columnIndex]) || !isMissing(pair.raw[columnIndex]))
  );

  if (usedColumnIndexes.length > IMPORT_LIMITS.maxColumns) {
    throw new Error(`工作表「${sheetName}」超过 ${formatInteger(IMPORT_LIMITS.maxColumns)} 列限制，请减少字段后重试。`);
  }

  if (!usedColumnIndexes.length) {
    throw new Error(`工作表「${sheetName}」为空，请选择包含表头和数据的工作表。`);
  }

  const fields = usedColumnIndexes.map((columnIndex) => {
    const formattedHeader = headerPair.formatted[columnIndex];
    const rawHeader = headerPair.raw[columnIndex];
    if (isMissing(formattedHeader)) {
      throw new Error(`工作表「${sheetName}」的第 ${columnIndex + 1} 列缺少表头，请补充字段名后重新上传。`);
    }
    if (typeof rawHeader !== "string") {
      throw new Error(`工作表「${sheetName}」未识别到有效文本表头；第 ${columnIndex + 1} 列首行不是字段名。`);
    }
    return normalizeFieldName(formattedHeader);
  });

  const duplicateFields = fields.filter((field, index) => fields.indexOf(field) !== index);
  if (duplicateFields.length) {
    throw new Error(`工作表「${sheetName}」包含重复表头「${duplicateFields[0]}」，请修改后重新上传。`);
  }

  const rows = dataPairs
    .filter((pair) => usedColumnIndexes.some((columnIndex) => !isMissing(pair.formatted[columnIndex]) || !isMissing(pair.raw[columnIndex])))
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

  if (!readNumberFormats) return { fields, rows };
  return {
    fields,
    rows,
    numberFormatHints: buildExcelNumberFormatHints(fields, usedColumnIndexes, dataPairs, readNumberFormats)
  };
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
  if (dataRowCount > IMPORT_LIMITS.maxRows) {
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

function parseCsvText(text, sourceName, importId = beginImport()) {
  if (!isCurrentImport(importId)) return;
  assertRuntimeDependency("Papa", "CSV 解析组件加载失败，请刷新页面重试。");
  Papa.parse(text, {
    header: true,
    skipEmptyLines: true,
    preview: IMPORT_LIMITS.maxRows + 1,
    // PapaParse workers cannot clone transformHeader functions. Keep parsing on
    // the main thread so protected and duplicate header handling remains intact.
    worker: false,
    transformHeader: encodeCsvHeaderForParser,
    complete: (results) => handleParsedData(results, sourceName, importId),
    error: (error) => showError(error.message, importId)
  });
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

function handleParsedData(results, sourceName, importId) {
  if (!isCurrentImport(importId)) return;
  const parseErrors = Array.isArray(results.errors) ? results.errors : [];
  const extraFieldError = parseErrors.find((error) => error.code === "TooManyFields");
  if (extraFieldError) {
    const rowNumber = Number.isInteger(extraFieldError.row) ? extraFieldError.row + 2 : null;
    showError(`CSV${rowNumber ? ` 第 ${rowNumber} 行` : ""}的数据列多于表头，继续分析会丢失额外单元格。请补全表头或修正该行后重试。`, importId);
    return;
  }
  const fatalError = parseErrors.find((error) =>
    error.code === "MissingQuotes"
    || (error.type === "Delimiter" && error.code !== "UndetectableDelimiter")
  );
  if (fatalError) {
    showError(`CSV 解析失败：${fatalError.message}`, importId);
    return;
  }

  const allFieldPairs = (results.meta?.fields || []).map((rawField) => ({
    raw: rawField,
    clean: normalizeFieldName(decodeCsvHeaderFromParser(rawField))
  }));
  if (allFieldPairs.length > IMPORT_LIMITS.maxColumns) {
    showError(`CSV 超过 ${formatInteger(IMPORT_LIMITS.maxColumns)} 列限制，请减少字段后重试。`, importId);
    return;
  }
  const fieldPairs = allFieldPairs.filter((pair) => pair.clean !== "");
  const fields = fieldPairs.map((pair) => pair.clean);
  const duplicateFields = [...new Set(
    fields.filter((field, index) => fields.indexOf(field) !== index)
  )];

  if (duplicateFields.length) {
    showError(`CSV 表头规范化后存在重复字段：${duplicateFields.join("、")}。请修改表头后重试。`, importId);
    return;
  }

  const sourceRows = Array.isArray(results.data) ? results.data : [];
  if (sourceRows.length > IMPORT_LIMITS.maxRows) {
    showError(`CSV 超过 ${formatInteger(IMPORT_LIMITS.maxRows)} 行数据限制，请拆分文件后重试。`, importId);
    return;
  }
  const blankHeaderHasData = allFieldPairs
    .filter((pair) => pair.clean === "")
    .some((pair) => sourceRows.some((row) => !isMissing(row?.[pair.raw])));
  if (blankHeaderHasData) {
    showError("CSV 存在缺少表头的数据列。请补全每一列的字段名后重试。", importId);
    return;
  }

  const parseWarnings = parseErrors
    .filter((error) => !(error.code === "UndetectableDelimiter" && fields.length === 1))
    .slice(0, 20)
    .map((error) => ({
      row: Number.isInteger(error.row) ? error.row + 2 : null,
      code: error.code || "ParseWarning",
      message: error.message || "CSV 行结构异常"
    }));
  const rows = sourceRows.reduce((normalizedRows, row) => {
    const normalized = normalizeParsedRow(row, fieldPairs);
    if (fields.some((field) => !isMissing(normalized[field]))) {
      normalizedRows.push(normalized);
    }
    return normalizedRows;
  }, []);

  if (!fields.length || !rows.length) {
    showError("CSV 中没有可分析的数据。请确认文件包含表头和至少一行数据。", importId);
    return;
  }

  commitTabularData(fields, rows, sourceName, importId, parseWarnings);
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
