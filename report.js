// 报告与导出：报告快照、HTML 报告生成、处理后 CSV 导出、问题行 CSV 导出和写法统一的表的导出（整表检查、写法统一的 CSV 和改动清单）、导出门禁、文件名和下载。

// v3.2.0: [中文, English] pairs (Lp), read with Lpick(); the report, its file names and its CSV
// columns follow the language at the moment of the export.
const REPORT_CHART_TITLES = Object.freeze({
  numericDistributionChart: Lp("数值字段分布", "Distribution of number columns"),
  categoryTopChart: Lp("分类字段 Top 10", "Top 10 values of text columns"),
  dateTrendChart: Lp("日期字段趋势", "Trend of date columns"),
  v2TopGroupChart: Lp("自定义分析 Top 10 分组", "Breakdown: top 10 groups"),
  v2SelectedTrendChart: Lp("自定义分析时间趋势", "Breakdown: trend over time")
});

const REPORT_FILTER_MAX_VALUES = 10;

const REPORT_NO_FILTER_TEXT = Lp("未设置筛选，结果基于完整数据。", "No filters are set; the results use all the data.");

const REPORT_QUALITY_SCOPE_TEXT = Lp("数据质量部分始终基于原始数据，不受筛选影响。", "The data quality part always uses the original data; filters do not change it.");

const REPORT_NO_CUSTOM_ANALYSIS = Lp("当前没有可导出的自定义分组分析结果。", "There is no breakdown by group to export yet.");

// What ensureFieldConfigurationApplied() refused (app.js writes the sentence).
const DRAFT_ACTION_REPORT = Lp("导出报告", "export the report");

const REPORT_CHART_HEADING_IDS = Object.freeze({
  numericDistributionChart: "numericChartTitle",
  categoryTopChart: "categoryChartTitle",
  dateTrendChart: "dateChartTitle",
  v2TopGroupChart: "v2TopGroupTitle",
  v2SelectedTrendChart: "v2TrendTitle"
});

function canExportReport() {
  return Boolean(
    state.rows.length
      && state.fields.length
      && state.profiles.length
      && state.analysisCompletedAt
      && !hasFieldConfigDraftChanges()
  );
}

function canExportProcessedData() {
  return Boolean(state.rows.length && state.fields.length);
}

function updateExportButtons() {
  const disabled = !canExportReport();
  [dom.exportHtmlReport].filter(Boolean).forEach((button) => {
    button.disabled = disabled;
    button.setAttribute("aria-disabled", String(disabled));
  });
  const dataExportDisabled = !canExportProcessedData();
  if (dom.exportFilteredCsv) {
    dom.exportFilteredCsv.disabled = dataExportDisabled;
    dom.exportFilteredCsv.setAttribute("aria-disabled", String(dataExportDisabled));
  }
  // 检查标准化改动 has the report gate; its exports also need a check, and the change list a cell.
  const checkDisabled = !canCheckStandardizedExport();
  const result = state.standardizedExport;
  [
    [dom.checkStandardizedExport, checkDisabled],
    [dom.exportStandardizedCsv, checkDisabled || !result],
    [dom.exportStandardizedChanges, checkDisabled || !result || !result.cellCount]
  ].forEach(([button, isDisabled]) => {
    if (!button) return;
    button.disabled = isDisabled;
    button.setAttribute("aria-disabled", String(isDisabled));
  });
  if (dom.standardizedExportGate) {
    dom.standardizedExportGate.textContent = checkDisabled
      ? formatStandardizedExportGate(state.rows.length > 0, hasFieldConfigDraftChanges(), Boolean(state.analysisCompletedAt))
      : "";
  }
  renderDataExportSummary();
}

function renderDataExportSummary() {
  const originalCount = state.rows.length;
  const filteredCount = getAnalysisRows().length;
  dom.dataExportOriginalCount.textContent = formatInteger(originalCount);
  dom.dataExportFilteredCount.textContent = formatInteger(filteredCount);
  // v3.2.0 commit 10: “共 42 行” / “42 rows”, the words written around the number.
  [[dom.dataExportFilteredBefore, dom.dataExportFilteredAfter, filteredCount], [dom.dataExportOriginalBefore, dom.dataExportOriginalAfter, originalCount]].forEach(([before, after, count]) => {
    const words = formatExportRowCount(count);
    before.textContent = words.before;
    after.textContent = words.after;
  });
  // v3.1.0: the rows left out are not in this export either.
  const excludedCount = state.rows.length ? countExcludedRows() : 0;
  dom.dataExportExcludedNote.textContent = excludedCount ? formatDataExportExcludedNote(excludedCount) : "";
}

function exportProcessedCsv(mode) {
  if (!canExportProcessedData()) {
    setDataExportStatus(() => L("请先上传并解析有效的数据文件。", "Upload and read a valid data file first."), "warning");
    return;
  }

  try {
    const rows = getProcessedExportRows(mode);
    const exportedAt = new Date();
    const content = buildCsvExportContent(rows, state.fields);
    const fileName = buildDataExportFileName(mode, exportedAt);
    downloadUtf8TextFile(content, fileName, "text/csv;charset=utf-8");
    setDataExportStatus(() => L(`已导出当前分析的 ${formatInteger(rows.length)} 行：${fileName}`, `Exported ${enCount(rows.length, "row in analysis", "rows in analysis")}: ${fileName}`), "success");
  } catch (error) {
    setDataExportStatus(() => L(`数据导出失败：${errorText(error)}`, `The data export failed: ${errorText(error)}`), "error");
  }
}

function getProcessedExportRows(mode) {
  // v3.1.0: one data export, the rows of the current analysis (the one without repeats is gone).
  if (mode === "filtered") return getAnalysisRows();
  throw languageError(Lp("不支持的数据导出类型。", "This data export type is not supported."));
}

function buildCsvExportContent(rows, fields) {
  const sourceFields = [...fields];
  return buildCsvTableContent(sourceFields, rows.map((row) => sourceFields.map((field) => readCsvExportCell(row, field))));
}

// CSV text of a header array and rows of cell arrays: every header and cell goes
// through sanitizeCsvExportValue(), lines end with CRLF. The data export (the rows of the
// current analysis), the problem-row export and both standardized exports write through here.
function buildCsvTableContent(headers, cellRows) {
  if (!window.Papa || typeof window.Papa.unparse !== "function") {
    throw languageError(Lp("CSV 导出组件未加载，请刷新页面重试。", "The CSV writer failed to load; refresh the page and try again."));
  }
  return window.Papa.unparse(
    { fields: headers.map(sanitizeCsvExportValue), data: cellRows.map((cells) => cells.map(sanitizeCsvExportValue)) },
    { header: true, newline: "\r\n", skipEmptyLines: false }
  );
}

// A row's cell for export: "" when the row lacks the field or holds null / undefined.
function readCsvExportCell(row, field) {
  return row && Object.prototype.hasOwnProperty.call(row, field) && row[field] !== null && row[field] !== undefined
    ? row[field]
    : "";
}

// The two columns the export adds, and how a taken name is numbered, in both languages.
const PROBLEM_ROWS_CSV_COLUMNS = Lp(
  { line: "文件行号", note: "说明", numbered: (name, suffix) => `${name}（${suffix}）` },
  { line: "File row", note: "Note", numbered: (name, suffix) => `${name} (${suffix})` }
);

// Column names of a problem-row export: 文件行号 and 说明 ahead of the fields, each
// numbered "（2）", "（3）"… (English " (2)") past any field already named so; field names stay.
function buildProblemRowsCsvHeaders(fields) {
  const columns = Lpick(PROBLEM_ROWS_CSV_COLUMNS);
  const taken = new Set(fields);
  const unique = (name) => {
    let candidate = name;
    for (let suffix = 2; taken.has(candidate); suffix += 1) candidate = columns.numbered(name, suffix);
    taken.add(candidate);
    return candidate;
  };
  return [unique(columns.line), unique(columns.note), ...fields];
}

// Every item of a problem-row list (listProblemRows()), not only the page on screen:
// its file row, its note and the original value of every field ("" for a parse
// warning without an imported row).
function buildProblemRowsCsvContent(list, fields) {
  // v3.1.0: a left-out row's note ends in （不计入分析）, as in the viewer.
  const excluded = getExcludedIndexSet();
  return buildCsvTableContent(buildProblemRowsCsvHeaders(fields), list.items.map((item) => {
    const row = item.index === null ? null : state.rows[item.index];
    return [item.line ?? "", markExcludedNote(item, excluded) || "", ...fields.map((field) => readCsvExportCell(row, field))];
  }));
}

const PROBLEM_KIND_FILE_NAMES = Object.freeze({
  missing: Lp("空值", "blank"),
  conversion: Lp("认不出", "unrecognized"),
  outlier: Lp("偏大或偏小", "high-or-low"),
  duplicateRows: Lp("整行重复", "duplicate-rows"),
  idDuplicate: Lp("编号重复", "duplicate-ids"),
  placeholder: Lp("像是空值的写法", "blank-like"),
  parseWarnings: Lp("读取有疑问", "reading-issues"),
  suspectedTotal: Lp("疑似合计", "possible-totals"),
  formulaError: Lp("公式错误", "formula-errors"),
  dateStyle: Lp("日期写法", "mixed-date-formats"),
  lookAlike: Lp("写法不一样", "variants")
});

// <原文件名>_问题行_<类别>[_<字段>][_<写法>]_<yyyymmdd>.csv (English <name>_problem-rows_<kind>…);
// only the part before the date is sanitized and shortened, so the date always stays.
function buildProblemRowsFileName(category, now = new Date()) {
  const originalName = getReportSourceFileName().split(/[\\/]/).pop() || "data";
  const stem = originalName.replace(/\.(?:csv|xlsx|xls)$/i, "") || "data";
  const parts = [
    { text: stem, cut: true },
    { text: L("问题行", "problem-rows"), cut: false },
    { text: PROBLEM_KIND_FILE_NAMES[category.kind] && Lpick(PROBLEM_KIND_FILE_NAMES[category.kind]), cut: false },
    { text: category.field, cut: true },
    { text: category.key, cut: true }
  ].filter((part) => part.text !== undefined && part.text !== null && part.text !== "");
  return `${buildLabeledFileName(parts)}_${formatReportFileDate(now)}.csv`;
}

// "导出这一类": the open category, every row, as a UTF-8 CSV with BOM.
function exportProblemRows() {
  const view = state.problemView;
  if (!canExportProcessedData() || !view) {
    setProblemRowsExportStatus(() => L("请先打开一类问题行。", "Open a category of problem rows first."), "warning");
    return;
  }
  try {
    const content = buildProblemRowsCsvContent(view.list, state.fields);
    const fileName = buildProblemRowsFileName(view.category, new Date());
    downloadUtf8TextFile(content, fileName, "text/csv;charset=utf-8");
    const count = view.list.items.length;
    setProblemRowsExportStatus(() => L(`已导出 ${formatInteger(count)} 条：${fileName}`, `Exported ${enCount(count, "row", "rows")}: ${fileName}`), "success");
  } catch (error) {
    setProblemRowsExportStatus(() => L(`导出失败：${errorText(error)}`, `The export failed: ${errorText(error)}`), "error");
  }
}

// ---- Standardized export (v2.10.0) ----
// Action codes stored per listed cell: 1–4 change the cell, the kept ones export it as written.
const STANDARDIZE_ACTIONS = Object.freeze(["", "date", "number", "placeholder", "trim", ...STANDARDIZE_KEPT_ACTIONS]);
const STANDARDIZE_ACTION_CODES = new Map(STANDARDIZE_ACTIONS.map((action, code) => [action, code]));
const STANDARDIZE_LAST_CHANGE_CODE = 4;
const STANDARDIZE_SAMPLE_LIMIT = 3;

// The same gate as the reports: data, a finished analysis and no unapplied field drafts.
function canCheckStandardizedExport() {
  return canExportReport();
}

// "mixed" when the cells the field reads as numbers (percent and scientific ones too) carry
// two or more currency signs, "single" for one, "" for none. Each distinct text is read once.
function findCurrencyMode(rows, profile, confirmedKeys) {
  const parse = getNumberParser(profile.numberStrategy);
  const fullwidth = Boolean(profile.numberStrategy && profile.numberStrategy.fullwidth);
  const currencies = new Set();
  const seen = new Set();
  for (let index = 0; index < rows.length; index += 1) {
    const value = rows[index][profile.field];
    if (typeof value !== "string" || seen.has(value)) continue;
    seen.add(value);
    if (!CURRENCY_SYMBOL_PATTERN.test(value) || isMissingOrConfirmed(value, confirmedKeys) || parse(value) === null) continue;
    const text = value.trim();
    currencies.add(currencyOfNumberText(fullwidth ? toHalfwidthNumberText(text) : text));
    if (currencies.size > 1) return "mixed";
  }
  return currencies.size ? "single" : "";
}

// The one-cell rule of a field, taken once outside the row loop: { kind, currencyMode, standardize }.
// A field read by year is not rewritten as dates; its cells follow the text rule.
function buildStandardizeRule(profile, rows, confirmedKeys) {
  const typeKey = profile ? profile.typeKey : "ignore";
  if (typeKey === "ignore") return { kind: "ignored", currencyMode: "", standardize: null };
  if (typeKey === "date") {
    if (profile.dateStrategy.granularity === "year") {
      return { kind: "year", currencyMode: "", standardize: (value) => standardizeTextCell(value, confirmedKeys) };
    }
    return { kind: "date", currencyMode: "", standardize: (value) => standardizeDateCell(value, profile.dateStrategy, confirmedKeys) };
  }
  if (typeKey === "numeric") {
    const currencyMode = findCurrencyMode(rows, profile, confirmedKeys);
    return { kind: "number", currencyMode, standardize: (value) => standardizeNumberCell(value, profile.numberStrategy, confirmedKeys, currencyMode) };
  }
  return { kind: "text", currencyMode: "", standardize: (value) => standardizeTextCell(value, confirmedKeys) };
}

// Every cell of the original rows by the applied configuration. counts: one entry per field,
// in field order; cells: the listed cells as four parallel arrays, by row then field (row index,
// field index, action code, exported value; the value is undefined for a kept cell). orderHint:
// a date field read in dmy / mdy order whose exported cells no longer give that order to
// buildDateStrategy(), and some of which a second import would then not read.
function buildStandardizedExport(rows, fields, profiles, confirmations) {
  const fieldProfiles = fields.map((field) => profiles.find((item) => item.field === field));
  const rules = fields.map((field, index) => buildStandardizeRule(fieldProfiles[index], rows, getConfirmedPlaceholderKeys(field, confirmations)));
  const counts = fields.map((field, index) => ({
    field,
    typeKey: fieldProfiles[index] ? fieldProfiles[index].typeKey : "ignore",
    kind: rules[index].kind,
    currencyMode: rules[index].currencyMode,
    unit: rules[index].kind === "number" ? fieldProfiles[index].numberStrategy.unit || "" : "",
    date: 0, number: 0, placeholder: 0, trim: 0, kept: 0,
    keptByAction: {},
    keptSamples: [],
    orderHint: false
  }));
  // Exported non-missing values of the date fields read in an order, for orderHint.
  const orderValues = fields.map((field, index) => (
    rules[index].kind === "date" && ["dmy", "mdy"].includes(fieldProfiles[index].dateStrategy.order) ? [] : null
  ));
  const memos = fields.map(() => new Map());
  const cells = { rows: [], fields: [], actions: [], values: [] };
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex];
    for (let fieldIndex = 0; fieldIndex < fields.length; fieldIndex += 1) {
      const rule = rules[fieldIndex];
      if (!rule.standardize) continue;
      const value = row[fields[fieldIndex]];
      let result;
      if (typeof value === "string") {
        result = memos[fieldIndex].get(value);
        if (!result) {
          result = rule.standardize(value);
          memos[fieldIndex].set(value, result);
        }
      } else result = rule.standardize(value);
      if (orderValues[fieldIndex] && !isMissing(result.value)) orderValues[fieldIndex].push(result.value);
      if (!result.action) continue;
      const count = counts[fieldIndex];
      const code = STANDARDIZE_ACTION_CODES.get(result.action);
      cells.rows.push(rowIndex);
      cells.fields.push(fieldIndex);
      cells.actions.push(code);
      if (code <= STANDARDIZE_LAST_CHANGE_CODE) {
        cells.values.push(result.value);
        count[result.action] += 1;
      } else {
        cells.values.push(undefined);
        count.kept += 1;
        count.keptByAction[result.action] = (count.keptByAction[result.action] || 0) + 1;
        if (count.keptSamples.length < STANDARDIZE_SAMPLE_LIMIT && !count.keptSamples.includes(value)) count.keptSamples.push(value);
      }
    }
  }
  // A hint only when the order changes and some exported cell that reads now would not read
  // with the order a second import finds.
  orderValues.forEach((values, index) => {
    if (!values) return;
    const strategy = fieldProfiles[index].dateStrategy;
    const exported = buildDateStrategy(fields[index], values);
    counts[index].orderHint = exported.order !== strategy.order
      && values.some((value) => toDate(value, strategy) !== null && toDate(value, exported) === null);
  });
  const totals = { date: 0, number: 0, placeholder: 0, trim: 0, kept: 0 };
  counts.forEach((count) => Object.keys(totals).forEach((key) => { totals[key] += count[key]; }));
  return { counts, cells, totals, cellCount: cells.rows.length };
}

// Every row's cells for export, the standardized values laid over the original ones.
function buildStandardizedCellRows(result, rows, fields) {
  const cellRows = rows.map((row) => fields.map((field) => readCsvExportCell(row, field)));
  const { cells } = result;
  for (let index = 0; index < cells.rows.length; index += 1) {
    if (cells.values[index] !== undefined) cellRows[cells.rows[index]][cells.fields[index]] = cells.values[index];
  }
  return cellRows;
}

// At most this many cells go into one change list: an Excel sheet holds 1,048,576 rows,
// one of them the header.
const STANDARDIZED_CHANGE_LIMIT = 1048575;
const STANDARDIZED_CHANGE_HEADERS = Lp(Object.freeze(["文件行号", "字段", "原值", "导出值", "处理"]), Object.freeze(["File row", "Column", "Original", "Exported", "Action"]));

// The standardized CSV: every original row and field, in order, original headers.
function buildStandardizedCsvContent(result, rows, fields) {
  return buildCsvTableContent([...fields], buildStandardizedCellRows(result, rows, fields));
}

// The change list: the first `limit` listed cells, by row then field. The original value and
// the file row are read now from the rows (a row without a file row gets ""); a kept cell
// exports its original value.
function buildStandardizedChangesCsvContent(result, rows, fields, getRowNumber, limit = STANDARDIZED_CHANGE_LIMIT) {
  const { cells } = result;
  const count = Math.min(result.cellCount, limit);
  const data = new Array(count);
  for (let index = 0; index < count; index += 1) {
    const row = rows[cells.rows[index]];
    const field = fields[cells.fields[index]];
    const original = readCsvExportCell(row, field);
    const exported = cells.values[index] === undefined ? original : cells.values[index];
    data[index] = [getRowNumber(row) ?? "", field, original, exported, formatStandardizeAction(STANDARDIZE_ACTIONS[cells.actions[index]])];
  }
  return buildCsvTableContent([...Lpick(STANDARDIZED_CHANGE_HEADERS)], data);
}

// <原文件名>_写法统一_<yyyymmdd>.csv, <原文件名>_写法统一改动清单[_前N条]_<yyyymmdd>.csv (English
// _standardized_, _standardized-changes[_first-N]_); only the part before the date is sanitized
// and shortened, so the date always stays.
function buildStandardizedFileName(kind, now = new Date(), truncatedTo = 0) {
  const originalName = getReportSourceFileName().split(/[\\/]/).pop() || "data";
  const stem = originalName.replace(/\.(?:csv|xlsx|xls)$/i, "") || "data";
  const labels = kind === "changes"
    ? [L("写法统一改动清单", "standardized-changes"), ...(truncatedTo ? [L(`前${truncatedTo}条`, `first-${truncatedTo}`)] : [])]
    : [L("写法统一", "standardized")];
  return `${buildLabeledFileName([{ text: stem, cut: true }, ...labels.map((text) => ({ text, cut: false }))])}_${formatReportFileDate(now)}.csv`;
}

// 导出标准化 CSV / 导出改动清单: from the result of the last check, never computed again.
function exportStandardizedData(kind, limit = STANDARDIZED_CHANGE_LIMIT) {
  const result = state.standardizedExport;
  if (!result || !canCheckStandardizedExport()) {
    setStandardizedExportStatus(() => L("请先点“检查会改哪些格”。", "Click “Check which cells will change” first."), "warning");
    return;
  }
  if (kind === "changes" && !result.cellCount) {
    setStandardizedExportStatus(() => L("没有需要列出的单元格。", "There are no cells to list."), "warning");
    return;
  }
  try {
    const now = new Date();
    if (kind === "changes") {
      const truncatedTo = result.cellCount > limit ? limit : 0;
      const content = buildStandardizedChangesCsvContent(result, state.rows, state.fields, getSourceRowNumber, limit);
      const fileName = buildStandardizedFileName("changes", now, truncatedTo);
      downloadUtf8TextFile(content, fileName, "text/csv;charset=utf-8");
      const listed = Math.min(result.cellCount, limit);
      const total = result.cellCount;
      setStandardizedExportStatus(() => formatStandardizedChangesStatus(listed, total, fileName), truncatedTo ? "warning" : "success");
      return;
    }
    const content = buildStandardizedCsvContent(result, state.rows, state.fields);
    const fileName = buildStandardizedFileName("data", now);
    downloadUtf8TextFile(content, fileName, "text/csv;charset=utf-8");
    const rowCount = state.rows.length;
    setStandardizedExportStatus(() => L(`已导出写法统一的表 ${formatInteger(rowCount)} 行：${fileName}`, `Exported the standardized table, ${enCount(rowCount, "row", "rows")}: ${fileName}`), "success");
  } catch (error) {
    setStandardizedExportStatus(() => L(`导出失败：${errorText(error)}`, `The export failed: ${errorText(error)}`), "error");
  }
}

// 检查标准化改动: computed here only, never on a refresh path.
function checkStandardizedExport() {
  if (!canCheckStandardizedExport()) {
    updateExportButtons();
    return;
  }
  state.standardizedExport = buildStandardizedExport(state.rows, state.fields, state.profiles, state.placeholderConfirmations);
  setStandardizedExportStatus("");
  renderStandardizedExport();
  updateExportButtons();
}

// v3.2.0: the four export status lines are one-off messages (app.js): a message is a text or a
// function returning one, written again in the other language after a switch.
function setStandardizedExportStatus(message, tone = "") {
  if (!dom.standardizedExportStatus) return;
  writeOneOffMessage(dom.standardizedExportStatus, "inline-status", message, tone);
}

function sanitizeCsvExportValue(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "number" || typeof value === "boolean") return value;
  const text = String(value);
  const candidate = text.replace(/^[ \t\r]+/, "");
  const startsWithControl = /^[\t\r]/.test(text);
  const startsWithFormulaToken = /^[=+\-@]/.test(candidate);
  if (!startsWithControl && !startsWithFormulaToken) return value;
  if (/^[+\-]/.test(candidate) && toNumber(candidate) !== null) return value;
  return `'${text}`;
}

function exportHtmlReport() {
  exportAnalysisReport();
}

// v3.2.0: one report, HTML (the Markdown report is gone).
function exportAnalysisReport() {
  if (!canExportReport()) {
    if (state.rows.length && hasFieldConfigDraftChanges()) {
      ensureFieldConfigurationApplied(DRAFT_ACTION_REPORT);
      setReportExportStatus(() => L("字段配置有未应用的修改，请先应用配置后再导出。", "Column settings have changes that are not applied; apply them before exporting."), "warning");
    } else {
      setStatus(() => L("请先完成数据分析，再导出报告。", "Finish the analysis before exporting the report."), "warning");
      setReportExportStatus(() => L("请先完成数据分析，再导出报告。", "Finish the analysis before exporting the report."), "warning");
    }
    updateExportButtons();
    return;
  }

  const buttons = [dom.exportHtmlReport].filter(Boolean);
  const labels = buttons.map((button) => button.textContent);
  buttons.forEach((button) => {
    button.disabled = true;
  });
  if (dom.exportHtmlReport) dom.exportHtmlReport.textContent = L("正在导出...", "Exporting...");
  setReportExportStatus(() => L("正在生成报告...", "Creating the report..."), "loading");

  try {
    const exportedAt = new Date();
    const reportData = buildReportData(exportedAt);
    const content = buildHtmlReport(reportData, captureCurrentChartImages());
    const fileName = buildReportFileName("html", exportedAt);
    downloadUtf8TextFile(content, fileName, "text/html;charset=utf-8");
    setStatus(() => L(`报告已生成：${fileName}`, `The report is ready: ${fileName}`), "success");
    setReportExportStatus(() => L(`报告已生成：${fileName}`, `The report is ready: ${fileName}`), "success");
  } catch (error) {
    setStatus(() => L(`报告导出失败：${errorText(error)}`, `The report export failed: ${errorText(error)}`), "error");
    setReportExportStatus(() => L(`报告导出失败：${errorText(error)}`, `The report export failed: ${errorText(error)}`), "error");
  } finally {
    buttons.forEach((button, index) => {
      button.textContent = labels[index];
    });
    updateExportButtons();
  }
}

function createUtf8BomBlob(content, mimeType) {
  return new Blob(["\uFEFF", content], { type: mimeType });
}

function downloadUtf8TextFile(content, fileName, mimeType) {
  const blob = createUtf8BomBlob(content, mimeType);
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function captureCurrentChartImages() {
  const knownIds = Object.keys(REPORT_CHART_TITLES);
  const chartIds = [
    ...knownIds,
    ...Object.keys(state.charts).filter((id) => !knownIds.includes(id))
  ];

  return chartIds.flatMap((id) => {
    const chart = state.charts[id];
    if (!chart) return [];
    const canvas = chart.canvas || document.getElementById(id);
    if (!canvas) return [];
    if (typeof canvas.closest === "function" && canvas.closest(".hidden")) return [];

    try {
      if (typeof chart.stop === "function") chart.stop();
      if (typeof chart.update === "function") chart.update("none");
      const dataUrl = typeof chart.toBase64Image === "function"
        ? chart.toBase64Image()
        : canvas.toDataURL("image/png");
      if (!isSafeChartDataUrl(dataUrl)) return [];
      return [{
        id,
        title: document.getElementById(REPORT_CHART_HEADING_IDS[id])?.textContent?.trim()
          || (REPORT_CHART_TITLES[id] && Lpick(REPORT_CHART_TITLES[id]))
          || chart.data?.datasets?.[0]?.label
          || L("分析图表", "Chart"),
        dataUrl
      }];
    } catch {
      return [];
    }
  });
}

function isSafeChartDataUrl(value) {
  return /^data:image\/(?:png|jpeg|webp);base64,[a-z0-9+/=]+$/i.test(String(value || ""));
}

function buildReportData(now = new Date()) {
  const exportedAt = normalizeReportDate(now, new Date());
  const analysisAt = normalizeReportDate(state.analysisCompletedAt, exportedAt);
  const analysisRows = getAnalysisRows();
  const qualityNumericStats = getQualityNumericStats();
  const activeProfiles = getActiveProfiles();
  const fieldTypes = state.profiles.map((profile) => ({
    field: profile.field,
    // The names in the language of the export (in Chinese the same as profile.inferredType / type).
    inferredType: fieldTypeLabel(profile.inferredTypeKey),
    appliedType: fieldTypeLabel(profile.typeKey),
    typeKey: profile.typeKey,
    nonMissingCount: profile.nonMissingCount,
    uniqueCount: profile.uniqueCount,
    missingCount: profile.missingCount,
    missingRate: formatPercent(profile.missingRate),
    conversionFailureCount: profile.conversionFailureCount || 0,
    conversionFailureSamples: [...(profile.conversionFailureSamples || [])],
    sampleValues: [...profile.sampleValues],
    numberUnit: getFieldNumberUnit(profile.field)
  }));
  const numericStats = state.numericStats.map((stat) => ({
    field: stat.field,
    count: stat.count,
    invalidCount: stat.invalidCount || 0,
    mean: formatFieldNumber(stat.field, stat.mean),
    median: formatFieldNumber(stat.field, stat.median),
    min: formatFieldNumber(stat.field, stat.min),
    max: formatFieldNumber(stat.field, stat.max),
    standardDeviation: formatFieldNumber(stat.field, stat.std),
    outlierCount: stat.outlierCount
  }));
  const categoryStats = state.categoryStats.map((category) => ({
    field: category.field,
    total: category.total,
    distinctCount: category.distinctCount,
    top: category.top.slice(0, 10).map((item) => ({
      name: displayGroupName(item.name),
      count: item.count,
      ratio: formatPercent(item.ratio)
    }))
  }));

  return {
    version: "v3.1.0",
    fileName: getReportSourceFileName(),
    sheetName: state.sourceSheetName || "",
    sourceType: state.sourceType || (state.sourceSheetName ? "excel" : "csv"),
    headerDetection: state.headerDetection
      ? {
        headerRowNumber: state.headerDetection.headerRowNumber,
        skippedRows: state.headerDetection.skippedRows.map((row) => ({ ...row }))
      }
      : null,
    analysisTime: formatReportDateTime(analysisAt),
    analysisTimeIso: analysisAt.toISOString(),
    exportTime: formatReportDateTime(exportedAt),
    dataScale: {
      rows: analysisRows.length,
      originalRows: state.rows.length,
      filteredOutRows: Math.max(0, getIncludedRows().length - analysisRows.length),
      excludedRows: countExcludedRows(),
      fields: state.fields.length,
      activeFields: activeProfiles.length,
      numericFields: getProfilesByType("numeric").length,
      categoryFields: getProfilesByType("category").length,
      dateFields: getProfilesByType("date").length,
      idFields: getProfilesByType("id").length,
      ignoredFields: getProfilesByType("ignore").length
    },
    filters: buildReportFilters(),
    exclusions: buildReportExclusions(),
    fieldTypes,
    quality: {
      totalMissing: state.totalMissing,
      duplicateRows: state.duplicateRows,
      parseWarningCount: state.parseWarnings.length,
      parseWarnings: state.parseWarnings.slice(0, 50).map((warning) => ({
        row: warning.row ?? "",
        message: formatParseWarning(warning)
      })),
      missingByField: state.profiles.map((profile) => ({
        field: profile.field,
        type: fieldTypeLabel(profile.typeKey),
        missingCount: profile.missingCount,
        missingRate: formatPercent(profile.missingRate)
      })),
      outliers: qualityNumericStats.map((stat) => ({
        field: stat.field,
        outlierCount: stat.outlierCount,
        lowerBound: formatFieldNumber(stat.field, stat.lowerBound),
        upperBound: formatFieldNumber(stat.field, stat.upperBound)
      })),
      idDuplicates: state.idDuplicates.map((item) => ({
        ...item,
        values: item.values.map((value) => ({ ...value, rowNumbers: [...value.rowNumbers] }))
      })),
      placeholders: buildReportPlaceholders(),
      numberConversions: buildNumberConversionSummary(state.profiles),
      numericLooking: state.numericLooking.map((item) => ({ ...item, samples: [...item.samples] })),
      suspectedTotals: state.suspectedTotals.map((item) => ({ ...item, matches: item.matches.map((match) => ({ ...match, fields: [...match.fields] })) })),
      suspectedTotalsExcluded: state.exclusionSourceRows === state.rows ? state.excludedRows.suspectedTotal.length : 0,
      suspectedTotalExtras: getSuspectedTotalExtras(),
      excludedScope: countExcludedRows()
        ? { totalRows: state.rows.length, excludedRows: countExcludedRows(), excludedMissing: countExcludedMissing() }
        : null,
      formulaErrors: state.formulaErrors.map((item) => ({ ...item, byError: { ...item.byError }, lines: [...item.lines], fieldIgnored: getFieldProfile(item.field)?.typeKey === "ignore" })),
      dateStyles: state.dateStyles.map((item) => ({ ...item, styles: item.styles.map((style) => ({ ...style })) })),
      constantFields: state.constantFields.map((item) => ({ ...item })),
      lookAlikes: state.lookAlikes.map((item) => ({ ...item, groups: item.groups.map((group) => ({ ...group, reasons: [...group.reasons], variants: group.variants.map((variant) => ({ ...variant })) })) }))
    },
    numericStats,
    categoryStats,
    dateTrends: buildReportDateTrends(),
    customAnalysis: buildReportCustomAnalysis(),
    insights: [state.chartInsights.category, state.chartInsights.trend].filter(Boolean)
  };
}

// v3.1.0: the rows left out of the analysis, by kind: how many and the file rows of the first
// REPORT_EXCLUSION_LINE_LIMIT (null when one of them has no file row). null when none.
function buildReportExclusions() {
  if (!countExcludedRows()) return null;
  const describe = (indexes) => {
    const lines = indexes.slice(0, REPORT_EXCLUSION_LINE_LIMIT).map((index) => getSourceRowNumber(state.rows[index]));
    return { count: indexes.length, lines: lines.some((line) => line === null) ? null : lines };
  };
  // Rows of both kinds, in row order; null when there are none.
  const totals = new Set(state.excludedRows.suspectedTotal);
  const overlap = state.excludedRows.duplicateRows.filter((index) => totals.has(index));
  return {
    suspectedTotal: describe(state.excludedRows.suspectedTotal),
    duplicateRows: describe(state.excludedRows.duplicateRows),
    overlap: overlap.length ? describe(overlap) : null
  };
}

// Confirmed keys, and pending keys with the effect text of the quality page.
function buildReportPlaceholders() {
  return {
    confirmed: state.placeholders.filter((entry) => entry.confirmed).map((entry) => ({
      field: entry.field,
      key: entry.key,
      rawValues: [...entry.rawValues],
      count: entry.count,
      fieldIgnored: entry.fieldIgnored
    })),
    pending: state.placeholders.filter((entry) => !entry.confirmed).map((entry) => {
      const profile = getFieldProfile(entry.field);
      const preview = previewPlaceholderConfirmation(entry.field, [entry.key]);
      return {
        field: entry.field,
        key: entry.key,
        tier: entry.tier,
        rawValues: [...entry.rawValues],
        count: entry.count,
        rowNumbers: [...entry.rowNumbers],
        effect: L(`按空值处理后，${formatPlaceholderChange(entry.field, profile.missingCount, profile.typeKey, preview)}`, `If treated as blank: ${formatPlaceholderChange(entry.field, profile.missingCount, profile.typeKey, preview)}`)
      };
    })
  };
}

// Describes the applied filters (state.filters only changes on "应用筛选",
// never on drafts) in the fixed order category, numeric, date.
function buildReportFilters(filters = state.filters) {
  const items = [];
  if (filters?.category) {
    const { field, values } = filters.category;
    const options = getCategoryFilterOptions(field);
    const keptKeys = new Set(values);
    // Kept labels follow the filter panel order (display labels, zh-CN sort).
    const keptOptions = options.filter((option) => keptKeys.has(option.value));
    const kept = keptOptions.map((option) => option.label);
    const keptValues = kept.slice(0, REPORT_FILTER_MAX_VALUES);
    // English lists the values after a colon, each in curly quotes; the blank one is its label (Blank).
    const quotedValues = () => keptOptions.slice(0, REPORT_FILTER_MAX_VALUES).map((option) => (option.value === "" ? option.label : `“${option.label}”`)).join(", ");
    const more = kept.length > keptValues.length ? L(`……等 ${formatInteger(kept.length)} 个`, `, … ${formatInteger(kept.length)} in all`) : "";
    items.push({
      type: "category",
      field,
      keptCount: kept.length,
      totalCount: options.length,
      keptValues,
      description: L(
        `${field}：保留 ${formatInteger(kept.length)} / 共 ${formatInteger(options.length)} 个类别${kept.length ? `（${keptValues.join("、")}${more}）` : ""}`,
        `${field}: keeps ${formatInteger(kept.length)} of ${enCount(options.length, "value", "values")}${kept.length ? `: ${quotedValues()}${more}` : ""}`
      )
    });
  }
  if (filters?.numeric) {
    const { field, min, max } = filters.numeric;
    items.push({
      type: "numeric",
      field,
      min,
      max,
      description: L(`${field}：${formatFieldNumber(field, min)} 至 ${formatFieldNumber(field, max)}`, `${field}: ${formatFieldNumber(field, min)} to ${formatFieldNumber(field, max)}`)
    });
  }
  if (filters?.date) {
    const { field, start, end } = filters.date;
    items.push({ type: "date", field, start, end, description: L(`${field}：${start} 至 ${end}`, `${field}: ${start} to ${end}`) });
  }
  return items;
}

// v3.2.0 (R10): #sourceName holds the display name alone since v3.0.0 (“name · sheet”), so the
// old “已载入 · ” prefix is gone from the fallback.
function getReportSourceFileName() {
  if (state.sourceFileName) return state.sourceFileName;
  const sourceTitle = dom.sourceName?.title || dom.sourceName?.textContent || "";
  return sourceTitle.split(" · ")[0] || "data-report";
}

// Every date field uses the trend chart's current metric / aggregate selection.
function buildReportDateTrends() {
  const { metricField, aggregate } = getDateChartSelection();
  return getProfilesByType("date").map((profile) => {
    const trend = buildDateChartTrend(profile.field, metricField, aggregate);
    if (!trend.labels.length) {
      return {
        field: profile.field,
        metricField,
        pointCount: 0,
        labels: [],
        values: [],
        start: "",
        end: "",
        summary: L(`日期字段「${profile.field}」里没有认得出的日期。`, `The date column “${profile.field}” has no recognized dates.`)
      };
    }

    const firstValue = trend.values[0];
    const lastValue = trend.values[trend.values.length - 1];
    const valueLabel = metricField
      ? L(`${formatFieldLabel(metricField)}${getDateChartAggregateLabel(aggregate)}`, `the ${getDateChartAggregateLabel(aggregate)} of ${formatFieldLabel(metricField)}`)
      : L("记录数", "the record count");
    const firstDisplay = metricField
      ? formatFieldNumber(metricField, firstValue)
      : formatInteger(firstValue);
    const lastDisplay = metricField
      ? formatFieldNumber(metricField, lastValue)
      : formatInteger(lastValue);
    const direction = formatReportTrendDirection(trend.values.length, firstValue, lastValue);

    return {
      field: profile.field,
      metricField,
      pointCount: trend.labels.length,
      labels: [...trend.labels],
      values: [...trend.values],
      start: trend.labels[0],
      end: trend.labels[trend.labels.length - 1],
      summary: L(
        `日期字段「${profile.field}」包含 ${formatInteger(trend.labels.length)} 个时间点；${valueLabel}从 ${trend.labels[0]} 的 ${firstDisplay} 变为 ${trend.labels[trend.labels.length - 1]} 的 ${lastDisplay}，${direction}。`,
        `The date column “${profile.field}” has ${enCount(trend.labels.length, "point in time", "points in time")}; ${valueLabel} goes from ${firstDisplay} in ${trend.labels[0]} to ${lastDisplay} in ${trend.labels[trend.labels.length - 1]}${direction}`
      )
    };
  });
}

// The direction at the end of a trend sentence: Chinese "整体上升" (the sentence adds "。"),
// English with its own punctuation (", rising overall." or " (only one point in time).").
function formatReportTrendDirection(pointCount, first, last) {
  if (pointCount < 2) return L("仅有一个时间点", " (only one point in time).");
  if (last > first) return L("整体上升", ", rising overall.");
  if (last < first) return L("整体下降", ", falling overall.");
  return L("整体持平", ", about the same overall.");
}

function buildReportCustomAnalysis() {
  const current = state.customAnalysis;

  if (!current || !Array.isArray(current.grouped) || !current.grouped.length) {
    return {
      completed: false,
      message: Lpick(REPORT_NO_CUSTOM_ANALYSIS)
    };
  }

  const rows = current.grouped.map((item) => ({
    group: displayGroupName(item.name),
    count: item.count,
    sum: formatFieldNumber(current.metricField, item.sum),
    average: formatFieldNumber(current.metricField, item.avg),
    median: formatFieldNumber(current.metricField, item.median),
    minimum: formatFieldNumber(current.metricField, item.min),
    maximum: formatFieldNumber(current.metricField, item.max)
  }));
  const aggregateMethod = current.aggregateMethod || "sum";
  const trendSummary = current.dateField
    ? summarizeCustomTrend(current.dateField, current.metricField, current.trend, aggregateMethod)
    : "";

  return {
    completed: true,
    metricField: current.metricField,
    groupField: current.groupField,
    aggregateMethod,
    aggregateLabel: getAggregateMethodLabel(aggregateMethod),
    dateField: current.dateField || "",
    metricInvalidCount: current.metricInvalidCount ?? countInvalidNumericValues(current.metricField, getAnalysisRows()),
    rows,
    trendSummary
  };
}

function summarizeCustomTrend(dateField, metricField, trend, aggregateMethod = "sum") {
  if (!trend?.labels?.length) {
    return L(`日期字段「${dateField}」没有足够的有效记录生成自定义趋势。`, `The date column “${dateField}” does not have enough valid records for a trend.`);
  }
  const first = trend.values[0];
  const last = trend.values[trend.values.length - 1];
  if (!Number.isFinite(first) || !Number.isFinite(last)) {
    return L(`按「${dateField}」观察时，「${metricField}」的聚合结果超出浏览器可表示的有限数值范围，无法可靠判断趋势。`, `By “${dateField}”, the aggregated “${metricField}” is beyond the numbers the browser can represent, so the trend cannot be judged reliably.`);
  }
  const direction = formatReportTrendDirection(trend.values.length, first, last);
  return L(
    `按「${dateField}」观察，「${metricField}」${getAggregateMethodLabel(aggregateMethod)}从 ${trend.labels[0]} 的 ${formatFieldNumber(metricField, first)} 变为 ${trend.labels[trend.labels.length - 1]} 的 ${formatFieldNumber(metricField, last)}，${direction}。`,
    `By “${dateField}”, the ${getAggregateMethodLabel(aggregateMethod)} of “${metricField}” goes from ${formatFieldNumber(metricField, first)} in ${trend.labels[0]} to ${formatFieldNumber(metricField, last)} in ${trend.labels[trend.labels.length - 1]}${direction}`
  );
}

function normalizeReportSnapshot(reportData = {}) {
  const quality = reportData.quality || {};
  const dataScale = reportData.dataScale || {};
  return {
    ...reportData,
    version: reportData.version || "v3.1.0",
    fileName: reportData.fileName || "data-report",
    sheetName: reportData.sheetName || "",
    // Snapshots from before v2.2.0 have no header detection.
    headerDetection: reportData.headerDetection || null,
    analysisTime: reportData.analysisTime || "—",
    exportTime: reportData.exportTime || "—",
    dataScale: {
      rows: 0,
      originalRows: dataScale.rows ?? 0,
      filteredOutRows: 0,
      // Snapshots from before v3.1.0 leave no rows out.
      excludedRows: 0,
      fields: 0,
      activeFields: 0,
      numericFields: 0,
      categoryFields: 0,
      dateFields: 0,
      idFields: 0,
      ignoredFields: 0,
      ...dataScale
    },
    filters: Array.isArray(reportData.filters) ? reportData.filters : [],
    // Snapshots of v3.1.0 before the fix have no overlap: null leaves the 共 N 行 part out.
    exclusions: reportData.exclusions ? { ...reportData.exclusions, overlap: reportData.exclusions.overlap || null } : null,
    fieldTypes: (reportData.fieldTypes || []).map((field) => ({
      ...field,
      inferredType: field.inferredType || field.type || "—",
      appliedType: field.appliedType || field.type || "—",
      nonMissingCount: field.nonMissingCount ?? 0,
      uniqueCount: field.uniqueCount ?? 0,
      missingCount: field.missingCount ?? 0,
      conversionFailureCount: field.conversionFailureCount ?? 0,
      conversionFailureSamples: Array.isArray(field.conversionFailureSamples) ? field.conversionFailureSamples : [],
      sampleValues: Array.isArray(field.sampleValues) ? field.sampleValues : [],
      // Snapshots from before v2.5.0 have no number units.
      numberUnit: typeof field.numberUnit === "string" ? field.numberUnit : ""
    })),
    quality: {
      ...quality,
      totalMissing: quality.totalMissing ?? 0,
      duplicateRows: quality.duplicateRows ?? 0,
      parseWarnings: Array.isArray(quality.parseWarnings) ? quality.parseWarnings : [],
      parseWarningCount: quality.parseWarningCount ?? quality.parseWarnings?.length ?? 0,
      missingByField: Array.isArray(quality.missingByField)
        ? quality.missingByField
        : Array.isArray(quality.fields) ? quality.fields : [],
      outliers: Array.isArray(quality.outliers) ? quality.outliers : [],
      // Snapshots from before v2.3.0 have no ID duplicate check: null leaves it out.
      idDuplicates: Array.isArray(quality.idDuplicates) ? quality.idDuplicates : null,
      // Snapshots from before v2.4.0 have no placeholder check: null leaves it out.
      placeholders: Array.isArray(quality.placeholders?.confirmed) && Array.isArray(quality.placeholders?.pending)
        ? quality.placeholders
        : null,
      // Snapshots from before v2.5.0 have no number sections: null leaves them out.
      numberConversions: Array.isArray(quality.numberConversions) ? quality.numberConversions : null,
      numericLooking: Array.isArray(quality.numericLooking) ? quality.numericLooking : null,
      // Snapshots from before v2.8.0 have neither check: null leaves them out.
      suspectedTotals: Array.isArray(quality.suspectedTotals) ? quality.suspectedTotals : null,
      formulaErrors: Array.isArray(quality.formulaErrors) ? quality.formulaErrors : null,
      // Snapshots from before v3.1.0: nothing left out, no sentences under the totals.
      suspectedTotalsExcluded: quality.suspectedTotalsExcluded ?? 0,
      suspectedTotalExtras: quality.suspectedTotalExtras || null,
      excludedScope: quality.excludedScope || null,
      // Snapshots from before v2.9.0 have none of the three: null leaves them out.
      dateStyles: Array.isArray(quality.dateStyles) ? quality.dateStyles : null,
      constantFields: Array.isArray(quality.constantFields) ? quality.constantFields : null,
      lookAlikes: Array.isArray(quality.lookAlikes) ? quality.lookAlikes : null
    },
    numericStats: (Array.isArray(reportData.numericStats) ? reportData.numericStats : []).map((item) => ({
      ...item,
      invalidCount: item.invalidCount ?? 0
    })),
    categoryStats: Array.isArray(reportData.categoryStats) ? reportData.categoryStats : [],
    dateTrends: Array.isArray(reportData.dateTrends) ? reportData.dateTrends : [],
    customAnalysis: reportData.customAnalysis
      ? {
        ...reportData.customAnalysis,
        ...(reportData.customAnalysis.completed ? { metricInvalidCount: reportData.customAnalysis.metricInvalidCount ?? 0 } : {})
      }
      : {
        completed: false,
        message: Lpick(REPORT_NO_CUSTOM_ANALYSIS)
      },
    insights: Array.isArray(reportData.insights) ? reportData.insights : []
  };
}

const REPORT_ID_DUPLICATE_HEADERS = Lp(["字段", "重复值", "出现次数", "文件行号", "说明"], ["Column", "Duplicate value", "Count", "File rows", "Note"]);

// Plain-text rows and notes of the report's ID duplicate section: listed values and
// mostly repeated fields as table rows, then overflow and clean-field notes.
function buildReportIdDuplicateSection(idDuplicates) {
  if (!idDuplicates.length) return { rows: [], notes: [L("未识别到 ID 字段。", "No ID columns were found.")] };
  const rows = [];
  const notes = [];
  const cleanFields = [];
  idDuplicates.forEach((item) => {
    if (item.mostlyRepeated) {
      rows.push([item.field, "—", "—", "—", formatMostlyRepeatedIdHint(item)]);
      return;
    }
    if (!item.values.length) {
      cleanFields.push(item.field);
      return;
    }
    item.values.forEach((value) => rows.push([
      item.field,
      value.value,
      value.count,
      formatRowNumberList(value.rowNumbers, value.count),
      formatIdenticalRows(value.identicalRowCount, value.identicalGroupCount)
    ]));
    const overflow = formatIdDuplicateOverflow(item.duplicateValueCount, item.values.length);
    if (overflow) notes.push(L(`「${item.field}」${overflow}`, `“${item.field}”: ${overflow}.`));
  });
  if (cleanFields.length) notes.push(formatCleanIdFieldsNote(cleanFields, rows.length > 0));
  return { rows, notes };
}

const REPORT_PLACEHOLDER_CONFIRMED_HEADERS = Lp(["字段", "写法", "出现次数", "状态"], ["Column", "Value", "Count", "Status"]);
const REPORT_PLACEHOLDER_PENDING_HEADERS = Lp(["字段", "写法", "档位", "出现次数", "文件行号", "按空值处理后的影响"], ["Column", "Value", "Kind", "Count", "File rows", "If treated as blank"]);
const REPORT_PLACEHOLDER_TIER_LABELS = { common: Lp("常见", "Common"), maybe: Lp("可能是有效值", "May be a real value") };

// Plain-text tables and notes of the report's placeholder section. Every note
// starts with fixed text, never with a field name or value.
function buildReportPlaceholderSection({ confirmed, pending }) {
  if (!confirmed.length && !pending.length) return { confirmedRows: [], pendingRows: [], notes: [L("没有发现像是空值的写法。", "No values that may mean blank were found.")] };
  const notes = formatPlaceholderNotes(pending.map((entry) => entry.key));
  return {
    confirmedRows: confirmed.map((entry) => [
      entry.field,
      formatPlaceholderValue(entry),
      entry.count,
      entry.fieldIgnored ? L("已按空值处理（字段已忽略）", "Treated as blank (column ignored)") : L("已按空值处理", "Treated as blank")
    ]),
    pendingRows: pending.map((entry) => [
      entry.field,
      formatPlaceholderValue(entry),
      REPORT_PLACEHOLDER_TIER_LABELS[entry.tier] ? Lpick(REPORT_PLACEHOLDER_TIER_LABELS[entry.tier]) : "—",
      entry.count,
      formatRowNumberList(entry.rowNumbers, entry.count),
      entry.effect
    ]),
    notes: notes ? [L(`说明：${notes}`, `Note: ${notes}`)] : []
  };
}

// Field labels of a normalized snapshot, with the unit its fieldTypes recorded.
function buildReportFieldLabeler(reportData) {
  const units = new Map(reportData.fieldTypes.map((field) => [field.field, field.numberUnit]));
  return (field) => formatFieldLabelWithUnit(field, units.get(field) || "");
}

// The lines of "按单位读取的数字" and "疑似数值列", the quality page's texts; an
// empty list for a missing (pre-v2.5.0) or empty section.
function buildReportNumberSections(quality) {
  return {
    conversions: (quality.numberConversions || []).map(formatNumberConversion),
    numericLooking: (quality.numericLooking || []).map(formatNumericLookingHint)
  };
}

// The counts and sections of "疑似合计 / 小计行" and "公式错误单元格", the quality page's
// texts; no section for a missing (pre-v2.8.0) or empty check. The report has no 查看, so
// the suspected totals beyond the limit are only counted.
function buildReportRowNoteSections(quality) {
  const totals = quality.suspectedTotals || [];
  const errors = quality.formulaErrors || [];
  return {
    suspectedTotalCount: totals.length,
    formulaErrorCount: formatFormulaErrorCount(errors),
    sections: [
      {
        title: L("疑似合计 / 小计行", "Possible total or subtotal rows"),
        // The page's note: another one once the totals are left out (v3.1.0).
        note: quality.suspectedTotalsExcluded ? formatSuspectedTotalExcludedNote(quality.suspectedTotalsExcluded) : Lpick(SUSPECTED_TOTAL_NOTE),
        lines: totals.slice(0, SUSPECTED_TOTAL_LIST_LIMIT).map((item) => formatSuspectedTotal(item)),
        overflow: formatSuspectedTotalOverflow(totals.length),
        extras: formatSuspectedTotalExtras(quality.suspectedTotalExtras, totals.length)
      },
      { title: L("公式错误单元格", "Cells with formula errors"), note: Lpick(FORMULA_ERROR_NOTE), lines: errors.map((item) => formatFormulaError(item)), overflow: "", extras: [] }
    ].filter((section) => section.lines.length)
  };
}

// The sections of "日期写法混用", "常量列" and "写法细节不同的值", the quality page's texts,
// and the two counts of the summary; no section for a missing (pre-v2.9.0) or empty check.
// Constant columns are no problem in themselves, so the summary does not count them.
function buildReportFormatSections(quality) {
  const dateStyles = quality.dateStyles || [];
  const constantFields = quality.constantFields || [];
  const lookAlikes = quality.lookAlikes || [];
  return {
    dateStyleCount: dateStyles.length,
    lookAlikeCount: lookAlikes.length,
    sections: [
      { title: L("日期写法混用", "Mixed date formats"), note: Lpick(DATE_STYLE_NOTE), lines: dateStyles.map(formatDateStyleLine) },
      { title: L("整列只有一个值", "Single-value columns"), note: Lpick(CONSTANT_FIELD_NOTE), lines: constantFields.map(formatConstantField) },
      { title: L("同一个值，写法不一样", "Same value, written differently"), note: Lpick(LOOK_ALIKE_NOTE), lines: lookAlikes.map(formatLookAlikeLine) }
    ].filter((section) => section.lines.length)
  };
}

// The unit after a count in <strong> of the English summary ("1 column", "2 columns").
function reportCountUnit(count, one, many) {
  return Number(count) === 1 ? one : many;
}

function buildHtmlReport(reportData, chartImages = []) {
  reportData = normalizeReportSnapshot(reportData);
  const fieldLabel = buildReportFieldLabeler(reportData);
  const numberSections = buildReportNumberSections(reportData.quality);
  // The summary counts problems only: numbers read by unit get their section, no count.
  const numericLookingSummary = numberSections.numericLooking.length
    ? L(`；像是数字的文本列：<strong>${escapeHtml(numberSections.numericLooking.length)}</strong> 个字段`, `; text columns that look numeric: <strong>${escapeHtml(numberSections.numericLooking.length)}</strong> ${reportCountUnit(numberSections.numericLooking.length, "column", "columns")}`)
    : "";
  const numberSectionHtml = [
    [L("按单位读取的数字", "Numbers read with a unit"), numberSections.conversions],
    [L("像是数字的文本列", "Text columns that look numeric"), numberSections.numericLooking]
  ].filter(([, lines]) => lines.length).map(([title, lines]) => `<h3>${title}</h3>
    <ul class="summary-list">${lines.map((line) => `<li>${escapeHtml(line)}</li>`).join("")}</ul>`).join("\n    ");
  const formatSections = buildReportFormatSections(reportData.quality);
  const formatSummary = `${formatSections.dateStyleCount ? L(`；日期写法混用：<strong>${escapeHtml(formatSections.dateStyleCount)}</strong> 个字段`, `; mixed date formats: <strong>${escapeHtml(formatSections.dateStyleCount)}</strong> ${reportCountUnit(formatSections.dateStyleCount, "column", "columns")}`) : ""}${formatSections.lookAlikeCount ? L(`；同一个值，写法不一样：<strong>${escapeHtml(formatSections.lookAlikeCount)}</strong> 个字段`, `; same value, written differently: <strong>${escapeHtml(formatSections.lookAlikeCount)}</strong> ${reportCountUnit(formatSections.lookAlikeCount, "column", "columns")}`) : ""}`;
  const formatSectionHtml = formatSections.sections.map(({ title, note, lines }) => `<h3>${title}</h3>
    <p>${escapeHtml(note)}</p>
    <ul class="summary-list">${lines.map((line) => `<li>${escapeHtml(line)}</li>`).join("")}</ul>`).join("\n    ");
  const rowNotes = buildReportRowNoteSections(reportData.quality);
  const rowNoteSummary = `${rowNotes.suspectedTotalCount ? L(`；疑似合计 / 小计行：<strong>${escapeHtml(rowNotes.suspectedTotalCount)}</strong> 行`, `; possible total or subtotal rows: <strong>${escapeHtml(rowNotes.suspectedTotalCount)}</strong>`) : ""}${rowNotes.formulaErrorCount ? L(`；公式错误：<strong>${escapeHtml(rowNotes.formulaErrorCount)}</strong> 个单元格`, `; formula errors: <strong>${escapeHtml(rowNotes.formulaErrorCount)}</strong> ${reportCountUnit(rowNotes.formulaErrorCount, "cell", "cells")}`) : ""}`;
  const rowNoteSectionHtml = rowNotes.sections.map(({ title, note, lines, overflow, extras }) => `<h3>${title}</h3>${note ? `
    <p>${escapeHtml(note)}</p>` : ""}
    <ul class="summary-list">${lines.map((line) => `<li>${escapeHtml(line)}</li>`).join("")}</ul>${overflow ? `
    <p>${escapeHtml(overflow)}</p>` : ""}${extras.map((extra) => `
    <p>${escapeHtml(extra)}</p>`).join("")}`).join("\n    ");
  // v3.1.0: the rows left out, each sentence only when there are any.
  const exclusionsText = formatReportExclusions(reportData.exclusions);
  const noFilterText = formatReportNoFilterText(reportData.dataScale.rows, reportData.dataScale.excludedRows) || Lpick(REPORT_NO_FILTER_TEXT);
  const excludedScope = reportData.quality.excludedScope;
  const qualityScopeText = excludedScope ? formatQualityScopeNote(excludedScope.totalRows, excludedScope.excludedRows, excludedScope.excludedMissing) : "";
  const safeCharts = chartImages.filter((chart) => isSafeChartDataUrl(chart.dataUrl));
  const sampleSeparator = L("、", ", ");
  const sheetMeta = reportData.sheetName
    ? `<div><dt>${L("工作表名称", "Worksheet")}</dt><dd>${escapeHtml(reportData.sheetName)}</dd></div>`
    : "";
  const headerNote = formatHeaderDetectionNote(reportData.headerDetection);
  const headerMeta = headerNote
    ? `<div><dt>${L("表头位置", "Header row")}</dt><dd>${escapeHtml(headerNote)}</dd></div>`
    : "";
  const idDuplicates = reportData.quality.idDuplicates;
  const idSummary = idDuplicates
    ? L(`；编号重复：<strong>${escapeHtml(countListedIdDuplicateFields(idDuplicates))}</strong> 个字段`, `; duplicate IDs: <strong>${escapeHtml(countListedIdDuplicateFields(idDuplicates))}</strong> ${reportCountUnit(countListedIdDuplicateFields(idDuplicates), "column", "columns")}`)
    : "";
  const idSection = idDuplicates ? buildReportIdDuplicateSection(idDuplicates) : null;
  const placeholders = reportData.quality.placeholders;
  const placeholderSummary = placeholders && (placeholders.confirmed.length || placeholders.pending.length)
    ? L(`；像是空值的写法：<strong>${escapeHtml(countPendingPlaceholderFields(placeholders.pending))}</strong> 个字段`, `; values that may mean blank: <strong>${escapeHtml(countPendingPlaceholderFields(placeholders.pending))}</strong> ${reportCountUnit(countPendingPlaceholderFields(placeholders.pending), "column", "columns")}`)
    : "";
  const placeholderSection = placeholders ? buildReportPlaceholderSection(placeholders) : null;
  const placeholderSectionHtml = placeholderSection
    ? `<h3>${L("像是空值的写法", "Values that may mean blank")}</h3>
    ${placeholderSection.confirmedRows.length ? `<p>${L("已按空值处理的写法：", "Treated as blank:")}</p>${buildHtmlReportTable(Lpick(REPORT_PLACEHOLDER_CONFIRMED_HEADERS), placeholderSection.confirmedRows, "")}` : ""}
    ${placeholderSection.pendingRows.length ? `<p>${L("还没确认的写法：", "Not treated as blank yet:")}</p>${buildHtmlReportTable(Lpick(REPORT_PLACEHOLDER_PENDING_HEADERS), placeholderSection.pendingRows, "")}` : ""}
    ${placeholderSection.notes.map((note) => `<p>${escapeHtml(note)}</p>`).join("")}`
    : "";
  const idSectionHtml = idSection
    ? `<h3>${L("编号重复", "Duplicate IDs")}</h3>
    ${idSection.rows.length ? buildHtmlReportTable(Lpick(REPORT_ID_DUPLICATE_HEADERS), idSection.rows, "") : ""}
    ${idSection.notes.map((note) => `<p>${escapeHtml(note)}</p>`).join("")}`
    : "";
  const categorySections = reportData.categoryStats.length
    ? reportData.categoryStats.map((category) => `
      <section class="subsection">
        <h3>${escapeHtml(category.field)} Top 10</h3>
        ${buildHtmlReportTable(
          Lpick(REPORT_CATEGORY_HEADERS),
          category.top.map((item) => [item.name, item.count, item.ratio]),
          L("该分类字段没有可统计的非空值。", "This text column has no non-blank values to count.")
        )}
      </section>
    `).join("")
    : `<p class="empty">${L("无适用的分类字段。", "There are no text columns.")}</p>`;
  const dateTrendItems = reportData.dateTrends.length
    ? `<ul class="summary-list">${reportData.dateTrends.map((trend) => `<li>${escapeHtml(trend.summary)}</li>`).join("")}</ul>`
    : `<p class="empty">${L("无适用的日期字段。", "There are no date columns.")}</p>`;
  const customSection = reportData.customAnalysis.completed
    ? `
      <p><strong>${L("指标字段：", "Metric column:")}</strong>${L("", " ")}${escapeHtml(fieldLabel(reportData.customAnalysis.metricField))}${L("　", " &nbsp;")}
      <strong>${L("分组字段：", "Group by:")}</strong>${L("", " ")}${escapeHtml(reportData.customAnalysis.groupField)}
      ${reportData.customAnalysis.aggregateLabel ? `${L("　", "&nbsp; ")}<strong>${L("聚合方式：", "Aggregation:")}</strong>${L("", " ")}${escapeHtml(reportData.customAnalysis.aggregateLabel)}` : ""}
      ${reportData.customAnalysis.dateField ? `${L("　", "&nbsp; ")}<strong>${L("日期字段：", "Date column:")}</strong>${L("", " ")}${escapeHtml(reportData.customAnalysis.dateField)}` : ""}</p>
      ${reportData.customAnalysis.trendSummary ? `<p>${escapeHtml(reportData.customAnalysis.trendSummary)}</p>` : ""}
      ${buildHtmlReportTable(
        [reportData.customAnalysis.groupField, ...Lpick(REPORT_CUSTOM_HEADERS)],
        reportData.customAnalysis.rows.map((row) => [row.group, row.count, row.sum, row.average, row.median, row.minimum, row.maximum]),
        Lpick(REPORT_NO_CUSTOM_ANALYSIS)
      )}
    `
    : `<p class="empty">${escapeHtml(reportData.customAnalysis.message)}</p>`;
  const chartSection = safeCharts.length
    ? `<div class="chart-list">${safeCharts.map((chart) => `
        <figure>
          <figcaption>${escapeHtml(chart.title)}</figcaption>
          <img src="${chart.dataUrl}" alt="${escapeHtml(chart.title)}">
        </figure>
      `).join("")}</div>`
    : `<p class="empty">${L("当前没有可嵌入的图表。", "There are no charts to include.")}</p>`;

  return `<!DOCTYPE html>
<html lang="${L("zh-CN", "en")}">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${escapeHtml(L(`${reportData.fileName} - 表格体检报告`, `${reportData.fileName} - Sheet Checkup report`))}</title>
  <style>
    :root { color-scheme: light; --ink: #23263a; --muted: #686f82; --line: #dfe2ea; --panel: #ffffff; --accent: #625bde; }
    * { box-sizing: border-box; }
    body { margin: 0; padding: 32px 20px; color: var(--ink); background: #f4f5f9; font-family: "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif; line-height: 1.6; }
    main { max-width: 1120px; margin: 0 auto; }
    header, section { margin-bottom: 20px; padding: 24px; border: 1px solid var(--line); border-radius: 14px; background: var(--panel); }
    header { border-top: 5px solid var(--accent); }
    h1, h2, h3 { margin: 0 0 14px; line-height: 1.3; }
    h1 { font-size: 28px; }
    h2 { padding-bottom: 9px; border-bottom: 1px solid var(--line); font-size: 20px; }
    h3 { font-size: 16px; }
    .meta, .scale-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(190px, 1fr)); gap: 12px; }
    .meta div, .scale-card { padding: 12px; border-radius: 9px; background: #f7f7fb; }
    dt, .label { color: var(--muted); font-size: 12px; }
    dd { margin: 3px 0 0; font-weight: 650; overflow-wrap: anywhere; }
    .value { margin-top: 3px; font-size: 20px; font-weight: 720; }
    .table-wrap { width: 100%; overflow-x: auto; }
    table { width: 100%; border-collapse: collapse; font-size: 13px; }
    th, td { padding: 9px 10px; border: 1px solid var(--line); text-align: left; vertical-align: top; overflow-wrap: anywhere; }
    th { background: #f0efff; }
    .subsection + .subsection { margin-top: 22px; }
    .summary-list, .insights { margin: 0; padding-left: 22px; }
    .summary-list li, .insights li { margin: 7px 0; }
    .empty { margin: 0; padding: 12px; border-radius: 8px; color: var(--muted); background: #f7f7fb; }
    .chart-list { display: grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap: 16px; }
    figure { margin: 0; padding: 14px; border: 1px solid var(--line); border-radius: 10px; }
    figcaption { margin-bottom: 10px; font-weight: 650; }
    img { display: block; width: 100%; height: auto; }
    footer { color: var(--muted); text-align: center; font-size: 12px; }
    @media (max-width: 640px) { body { padding: 14px 8px; } header, section { padding: 16px; } .chart-list { grid-template-columns: 1fr; } }
    @media print { body { padding: 0; background: #fff; } header, section { break-inside: avoid; box-shadow: none; } }
  </style>
</head>
<body>
<main>
  <header>
    <h1>${L("表格体检报告", "Sheet Checkup report")}</h1>
    <dl class="meta">
      <div><dt>${L("文件名称", "File")}</dt><dd>${escapeHtml(reportData.fileName)}</dd></div>
      ${sheetMeta}
      ${headerMeta}
      <div><dt>${L("分析时间", "Analyzed at")}</dt><dd>${escapeHtml(reportData.analysisTime)}</dd></div>
      <div><dt>${L("报告生成时间", "Report created at")}</dt><dd>${escapeHtml(reportData.exportTime)}</dd></div>
    </dl>
  </header>
  <section>
    <h2>${L("数据规模", "Data size")}</h2>
    <div class="scale-grid">
      ${[
        [L("当前分析行数", "Rows in analysis"), reportData.dataScale.rows],
        [L("原始数据行数", "Rows in the original data"), reportData.dataScale.originalRows],
        [L("筛除行数", "Rows filtered out"), reportData.dataScale.filteredOutRows],
        ...(reportData.dataScale.excludedRows ? [[L("不计入分析行数", "Rows excluded from analysis"), reportData.dataScale.excludedRows]] : []),
        [L("字段总数", "Columns"), reportData.dataScale.fields],
        [L("参与分析字段", "Columns in analysis"), reportData.dataScale.activeFields],
        [L("数值字段", "Number columns"), reportData.dataScale.numericFields],
        [L("分类字段", "Text columns"), reportData.dataScale.categoryFields],
        [L("日期字段", "Date columns"), reportData.dataScale.dateFields],
        [L("ID 字段", "ID columns"), reportData.dataScale.idFields],
        [L("忽略字段", "Ignored columns"), reportData.dataScale.ignoredFields]
      ].map(([label, value]) => `<div class="scale-card"><div class="label">${escapeHtml(label)}</div><div class="value">${escapeHtml(value)}</div></div>`).join("")}
    </div>
  </section>
  <section>
    <h2>${L("筛选条件", "Filters")}</h2>
    ${reportData.filters.length
      ? buildHtmlReportTable(Lpick(REPORT_FILTER_HEADERS), reportData.filters.map((item) => [item.field, item.description]), Lpick(REPORT_NO_FILTER_TEXT))
      : `<p>${escapeHtml(noFilterText)}</p>`}${exclusionsText ? `
    <p>${escapeHtml(exclusionsText)}</p>` : ""}
    <p>${escapeHtml(Lpick(REPORT_QUALITY_SCOPE_TEXT))}</p>
  </section>
  <section>
    <h2>${L("字段类型结果", "Column types")}</h2>
    ${buildHtmlReportTable(
      Lpick(REPORT_FIELD_TYPE_HEADERS),
      reportData.fieldTypes.map((field) => [
        field.field,
        field.inferredType,
        field.appliedType,
        field.nonMissingCount,
        field.uniqueCount,
        field.missingCount,
        formatConversionFailures(field.conversionFailureCount, field.conversionFailureSamples),
        field.sampleValues.join(sampleSeparator) || "—"
      ]),
      L("无字段类型结果。", "No column types.")
    )}
  </section>
  <section>
    <h2>${L("数据质量报告", "Data quality")}</h2>${qualityScopeText ? `
    <p>${escapeHtml(qualityScopeText)}</p>` : ""}
    <p>${L(
      `空值总数：<strong>${escapeHtml(reportData.quality.totalMissing)}</strong>；重复行：<strong>${escapeHtml(reportData.quality.duplicateRows)}</strong>；读取时有疑问：<strong>${escapeHtml(reportData.quality.parseWarningCount)}</strong> 处`,
      `Blanks: <strong>${escapeHtml(reportData.quality.totalMissing)}</strong>; duplicate rows: <strong>${escapeHtml(reportData.quality.duplicateRows)}</strong>; reading issues: <strong>${escapeHtml(reportData.quality.parseWarningCount)}</strong>`
    )}${idSummary}${placeholderSummary}${numericLookingSummary}${rowNoteSummary}${formatSummary}${L("。", ".")}</p>
    ${rowNoteSectionHtml}
    <h3>${L("各字段空值情况", "Blanks by column")}</h3>
    ${buildHtmlReportTable(
      Lpick(REPORT_MISSING_HEADERS),
      reportData.quality.missingByField.map((item) => [item.field, item.type, item.missingCount, item.missingRate]),
      L("无字段质量结果。", "No column results.")
    )}
    ${numberSectionHtml}
    ${formatSectionHtml}
    ${placeholderSectionHtml}
    <h3>${L("偏大或偏小的值", "Unusually high or low values")}</h3>
    ${buildHtmlReportTable(
      Lpick(REPORT_OUTLIER_HEADERS),
      reportData.quality.outliers.map((item) => [fieldLabel(item.field), item.outlierCount, item.lowerBound, item.upperBound]),
      Lpick(REPORT_NO_NUMBER_FIELDS)
    )}
    ${idSectionHtml}
  </section>
  <section>
    <h2>${L("数值字段描述性统计", "Number column statistics")}</h2>
    ${buildHtmlReportTable(
      Lpick(REPORT_NUMERIC_HEADERS),
      reportData.numericStats.map((item) => [fieldLabel(item.field), item.count, item.invalidCount, item.mean, item.median, item.min, item.max, item.standardDeviation]),
      Lpick(REPORT_NO_NUMBER_FIELDS)
    )}
  </section>
  <section>
    <h2>${L("分类字段 Top 10", "Top 10 values of text columns")}</h2>
    ${categorySections}
  </section>
  <section>
    <h2>${L("日期趋势摘要", "Date trends")}</h2>
    ${dateTrendItems}
  </section>
  <section>
    <h2>${L("自定义分组分析结果", "Breakdown by group")}</h2>
    ${customSection}
  </section>
  <section>
    <h2>${L("当前图表", "Charts")}</h2>
    ${chartSection}
  </section>
  <section>
    <h2>${L("自动分析结论", "Automatic conclusions")}</h2>
    ${reportData.insights.length
      ? `<ol class="insights">${reportData.insights.map((insight) => `<li>${escapeHtml(insight)}</li>`).join("")}</ol>`
      : `<p class="empty">${L("暂无自动分析结论。", "There are no automatic conclusions.")}</p>`}
  </section>
  <footer>${L(
    `由<a href="https://github.com/wy-data-30/smart-tabular-analyzer">表格体检</a> ${escapeHtml(reportData.version)} 在浏览器本地生成；数据未上传。`,
    `Created by <a href="https://github.com/wy-data-30/smart-tabular-analyzer">Sheet Checkup</a> ${escapeHtml(reportData.version)} in your browser; no data was uploaded.`
  )}</footer>
</main>
</body>
</html>`;
}

// The report's table headers in both languages (v3.2.0).
const REPORT_CATEGORY_HEADERS = Lp(["类别", "数量", "占比"], ["Value", "Count", "Share"]);
const REPORT_CUSTOM_HEADERS = Lp(["记录数", "求和", "平均值", "中位数", "最小值", "最大值"], ["Records", "Sum", "Average", "Median", "Minimum", "Maximum"]);
const REPORT_FILTER_HEADERS = Lp(["字段", "条件"], ["Column", "Condition"]);
const REPORT_FIELD_TYPE_HEADERS = Lp(["字段名", "自动识别类型", "最终类型", "非空值", "唯一值", "空值", "认不出", "示例值"], ["Column", "Detected type", "Applied type", "Non-blank", "Unique", "Blank", "Unrecognized", "Sample values"]);
const REPORT_MISSING_HEADERS = Lp(["字段", "类型", "空值", "空值占比"], ["Column", "Type", "Blank", "Blank share"]);
const REPORT_OUTLIER_HEADERS = Lp(["字段", "偏大或偏小的个数", "下界", "上界"], ["Column", "High/low values", "Lower bound", "Upper bound"]);
const REPORT_NUMERIC_HEADERS = Lp(["字段", "有效数量", "认不出", "平均值", "中位数", "最小值", "最大值", "标准差"], ["Column", "Valid", "Unrecognized", "Average", "Median", "Minimum", "Maximum", "Standard deviation"]);
const REPORT_NO_NUMBER_FIELDS = Lp("未识别到数值字段，暂无可用数值统计。", "No number columns were found, so there are no number statistics.");

function buildHtmlReportTable(headers, rows, emptyMessage) {
  if (!rows.length) return `<p class="empty">${escapeHtml(emptyMessage)}</p>`;
  return `<div class="table-wrap"><table>
    <thead><tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}</tr></thead>
    <tbody>${rows.map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>`).join("")}</tbody>
  </table></div>`;
}

// v3.2.0: always .html; the extension parameter stays so the callers do not change.
function buildReportFileName(extension, now = new Date()) {
  const originalName = getReportSourceFileName().split(/[\\/]/).pop() || "data-report";
  const stem = originalName.replace(/\.(?:csv|xlsx|xls)$/i, "") || "data-report";
  return `${sanitizeFileName(stem)}-${formatReportFileTimestamp(now)}.html`;
}

function buildDataExportFileName(mode, now = new Date()) {
  const suffix = L("当前分析", "in-analysis");
  const originalName = getReportSourceFileName().split(/[\\/]/).pop() || "data";
  const stem = originalName.replace(/\.(?:csv|xlsx|xls)$/i, "") || "data";
  return `${sanitizeFileName(stem)}_${suffix}_${formatReportFileDate(now)}.csv`;
}

function normalizeReportDate(value, fallback = new Date()) {
  const parsed = value instanceof Date ? new Date(value.getTime()) : value ? new Date(value) : null;
  if (parsed && !Number.isNaN(parsed.getTime())) return parsed;
  const fallbackDate = fallback instanceof Date ? new Date(fallback.getTime()) : new Date(fallback);
  return Number.isNaN(fallbackDate.getTime()) ? new Date() : fallbackDate;
}

function formatReportDateTime(value) {
  const date = normalizeReportDate(value);
  return `${date.getFullYear()}-${padReportNumber(date.getMonth() + 1)}-${padReportNumber(date.getDate())} ${padReportNumber(date.getHours())}:${padReportNumber(date.getMinutes())}:${padReportNumber(date.getSeconds())}`;
}

function formatReportFileTimestamp(value) {
  const date = normalizeReportDate(value);
  return `${date.getFullYear()}${padReportNumber(date.getMonth() + 1)}${padReportNumber(date.getDate())}-${padReportNumber(date.getHours())}${padReportNumber(date.getMinutes())}${padReportNumber(date.getSeconds())}`;
}

function formatReportFileDate(value) {
  const date = normalizeReportDate(value);
  return `${date.getFullYear()}${padReportNumber(date.getMonth() + 1)}${padReportNumber(date.getDate())}`;
}

function padReportNumber(value) {
  return String(value).padStart(2, "0");
}

// v3.2.0 (R1): at most this many characters before the date and the extension.
const FILE_NAME_LIMIT = 80;

// What sanitizeFileName() does before trimming and cutting: unsafe characters become "-", bidi
// controls go, runs of white space become one space.
function cleanFileNamePart(value) {
  return String(value)
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, "-")
    .replace(/[\u202A-\u202E\u2066-\u2069]/g, "")
    .replace(/\s+/g, " ");
}

function sanitizeFileName(value) {
  return cleanFileNamePart(value || "data-report").trim().slice(0, FILE_NAME_LIMIT) || "data-report";
}

// The original name and the labels joined by "_", for the problem-row and standardized files
// (v3.2.0, R1). A name that fits in FILE_NAME_LIMIT is exactly sanitizeFileName() over the whole,
// as before. A longer one is cut in the parts marked { cut: true } (the original name, the
// user's column and writing), the longest first, one character at a time; the labels (问题行,
// the category word, 写法统一改动清单, 前N条) are never cut.
function buildLabeledFileName(parts) {
  const whole = cleanFileNamePart(parts.map((part) => part.text).join("_")).trim();
  if (whole.length <= FILE_NAME_LIMIT) return whole || "data-report";
  const pieces = parts.map((part) => ({ cut: part.cut, chars: Array.from(cleanFileNamePart(part.text)) }));
  pieces[0].chars = Array.from(pieces[0].chars.join("").trimStart());
  pieces[pieces.length - 1].chars = Array.from(pieces[pieces.length - 1].chars.join("").trimEnd());
  const textLength = (piece) => piece.chars.join("").length;
  let length = pieces.reduce((sum, piece) => sum + textLength(piece), pieces.length - 1);
  while (length > FILE_NAME_LIMIT) {
    let longest = null;
    pieces.forEach((piece) => {
      if (piece.cut && piece.chars.length > 1 && (!longest || textLength(piece) > textLength(longest))) longest = piece;
    });
    if (!longest) break;
    length -= longest.chars.pop().length;
  }
  return pieces
    .map((piece) => (piece.cut ? piece.chars.join("").trim() : piece.chars.join("")))
    .filter(Boolean)
    .join("_");
}

function setReportExportStatus(message, tone = "") {
  if (!dom.reportExportStatus) return;
  writeOneOffMessage(dom.reportExportStatus, "inline-status", message, tone);
}

function setProblemRowsExportStatus(message, tone = "") {
  if (!dom.problemRowsExportStatus) return;
  writeOneOffMessage(dom.problemRowsExportStatus, "inline-status", message, tone);
}

function setDataExportStatus(message, tone = "") {
  if (!dom.dataExportStatus) return;
  writeOneOffMessage(dom.dataExportStatus, "inline-status", message, tone);
}
