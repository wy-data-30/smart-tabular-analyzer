// 报告与导出：报告快照、HTML / Markdown 报告生成、处理后 CSV 导出、问题行 CSV 导出和标准化导出（整表检查、标准化 CSV 和改动清单）、导出门禁、文件名和下载。

const REPORT_CHART_TITLES = Object.freeze({
  numericDistributionChart: "数值字段分布",
  categoryTopChart: "分类字段 Top 10",
  dateTrendChart: "日期字段趋势",
  v2TopGroupChart: "自定义分析 Top 10 分组",
  v2SelectedTrendChart: "自定义分析时间趋势"
});

const REPORT_FILTER_MAX_VALUES = 10;

const REPORT_NO_FILTER_TEXT = "未设置筛选，结果基于完整数据。";

const REPORT_QUALITY_SCOPE_TEXT = "数据质量部分始终基于原始数据，不受筛选影响。";

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
  [dom.exportHtmlReport, dom.exportMarkdownReport].filter(Boolean).forEach((button) => {
    button.disabled = disabled;
    button.setAttribute("aria-disabled", String(disabled));
  });
  const dataExportDisabled = !canExportProcessedData();
  [dom.exportFilteredCsv, dom.exportDeduplicatedCsv].filter(Boolean).forEach((button) => {
    button.disabled = dataExportDisabled;
    button.setAttribute("aria-disabled", String(dataExportDisabled));
  });
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
  const deduplicatedCount = Math.max(0, originalCount - state.duplicateRows);
  dom.dataExportOriginalCount.textContent = formatInteger(originalCount);
  dom.dataExportFilteredCount.textContent = formatInteger(filteredCount);
  dom.dataExportDeduplicatedCount.textContent = formatInteger(deduplicatedCount);
}

function exportProcessedCsv(mode) {
  if (!canExportProcessedData()) {
    setDataExportStatus("请先上传并解析有效的数据文件。", "warning");
    return;
  }

  try {
    const rows = getProcessedExportRows(mode);
    const exportedAt = new Date();
    const content = buildCsvExportContent(rows, state.fields);
    const fileName = buildDataExportFileName(mode, exportedAt);
    downloadUtf8TextFile(content, fileName, "text/csv;charset=utf-8");
    const exportLabel = mode === "deduplicated" ? "去重后数据" : "筛选后数据";
    setDataExportStatus(`已导出${exportLabel} ${formatInteger(rows.length)} 行：${fileName}`, "success");
  } catch (error) {
    setDataExportStatus(`数据导出失败：${error.message}`, "error");
  }
}

function getProcessedExportRows(mode) {
  if (mode === "filtered") return getAnalysisRows();
  if (mode === "deduplicated") return deduplicateRows(state.rows, state.fields);
  throw new Error("不支持的数据导出类型。");
}

function buildCsvExportContent(rows, fields) {
  const sourceFields = [...fields];
  return buildCsvTableContent(sourceFields, rows.map((row) => sourceFields.map((field) => readCsvExportCell(row, field))));
}

// CSV text of a header array and rows of cell arrays: every header and cell goes
// through sanitizeCsvExportValue(), lines end with CRLF. Both data exports and the
// problem-row export write through here.
function buildCsvTableContent(headers, cellRows) {
  if (!window.Papa || typeof window.Papa.unparse !== "function") {
    throw new Error("CSV 导出组件未加载，请刷新页面重试。");
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

// Column names of a problem-row export: 文件行号 and 说明 ahead of the fields, each
// numbered "（2）", "（3）"… past any field already named so; field names stay.
function buildProblemRowsCsvHeaders(fields) {
  const taken = new Set(fields);
  const unique = (name) => {
    let candidate = name;
    for (let suffix = 2; taken.has(candidate); suffix += 1) candidate = `${name}（${suffix}）`;
    taken.add(candidate);
    return candidate;
  };
  return [unique("文件行号"), unique("说明"), ...fields];
}

// Every item of a problem-row list (listProblemRows()), not only the page on screen:
// its file row, its note and the original value of every field ("" for a parse
// warning without an imported row).
function buildProblemRowsCsvContent(list, fields) {
  return buildCsvTableContent(buildProblemRowsCsvHeaders(fields), list.items.map((item) => {
    const row = item.index === null ? null : state.rows[item.index];
    return [item.line ?? "", item.note || "", ...fields.map((field) => readCsvExportCell(row, field))];
  }));
}

const PROBLEM_KIND_FILE_NAMES = Object.freeze({
  missing: "空值",
  conversion: "认不出",
  outlier: "偏大或偏小",
  duplicateRows: "整行重复",
  idDuplicate: "编号重复",
  placeholder: "像是空值的写法",
  parseWarnings: "读取有疑问",
  suspectedTotal: "疑似合计",
  formulaError: "公式错误",
  dateStyle: "日期写法",
  lookAlike: "写法不一样"
});

// <原文件名>_问题行_<类别>[_<字段>][_<写法>]_<yyyymmdd>.csv; only the part before the
// date is sanitized and shortened, so the date always stays.
function buildProblemRowsFileName(category, now = new Date()) {
  const originalName = getReportSourceFileName().split(/[\\/]/).pop() || "data";
  const stem = originalName.replace(/\.(?:csv|xlsx|xls)$/i, "") || "data";
  const parts = [stem, "问题行", PROBLEM_KIND_FILE_NAMES[category.kind], category.field, category.key]
    .filter((part) => part !== undefined && part !== null && part !== "");
  return `${sanitizeFileName(parts.join("_"))}_${formatReportFileDate(now)}.csv`;
}

// "导出这一类": the open category, every row, as a UTF-8 CSV with BOM.
function exportProblemRows() {
  const view = state.problemView;
  if (!canExportProcessedData() || !view) {
    setProblemRowsExportStatus("请先打开一类问题行。", "warning");
    return;
  }
  try {
    const content = buildProblemRowsCsvContent(view.list, state.fields);
    const fileName = buildProblemRowsFileName(view.category, new Date());
    downloadUtf8TextFile(content, fileName, "text/csv;charset=utf-8");
    setProblemRowsExportStatus(`已导出 ${formatInteger(view.list.items.length)} 条：${fileName}`, "success");
  } catch (error) {
    setProblemRowsExportStatus(`导出失败：${error.message}`, "error");
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
const STANDARDIZED_CHANGE_HEADERS = Object.freeze(["文件行号", "字段", "原值", "导出值", "处理"]);

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
  return buildCsvTableContent([...STANDARDIZED_CHANGE_HEADERS], data);
}

// <原文件名>_标准化_<yyyymmdd>.csv, <原文件名>_标准化改动清单[_前N条]_<yyyymmdd>.csv; only the part
// before the date is sanitized and shortened, so the date always stays.
function buildStandardizedFileName(kind, now = new Date(), truncatedTo = 0) {
  const originalName = getReportSourceFileName().split(/[\\/]/).pop() || "data";
  const stem = originalName.replace(/\.(?:csv|xlsx|xls)$/i, "") || "data";
  const label = kind === "changes" ? `写法统一改动清单${truncatedTo ? `_前${truncatedTo}条` : ""}` : "写法统一";
  return `${sanitizeFileName(`${stem}_${label}`)}_${formatReportFileDate(now)}.csv`;
}

// 导出标准化 CSV / 导出改动清单: from the result of the last check, never computed again.
function exportStandardizedData(kind, limit = STANDARDIZED_CHANGE_LIMIT) {
  const result = state.standardizedExport;
  if (!result || !canCheckStandardizedExport()) {
    setStandardizedExportStatus("请先点“检查会改哪些格”。", "warning");
    return;
  }
  if (kind === "changes" && !result.cellCount) {
    setStandardizedExportStatus("没有需要列出的单元格。", "warning");
    return;
  }
  try {
    const now = new Date();
    if (kind === "changes") {
      const truncatedTo = result.cellCount > limit ? limit : 0;
      const content = buildStandardizedChangesCsvContent(result, state.rows, state.fields, getSourceRowNumber, limit);
      const fileName = buildStandardizedFileName("changes", now, truncatedTo);
      downloadUtf8TextFile(content, fileName, "text/csv;charset=utf-8");
      setStandardizedExportStatus(formatStandardizedChangesStatus(Math.min(result.cellCount, limit), result.cellCount, fileName), truncatedTo ? "warning" : "success");
      return;
    }
    const content = buildStandardizedCsvContent(result, state.rows, state.fields);
    const fileName = buildStandardizedFileName("data", now);
    downloadUtf8TextFile(content, fileName, "text/csv;charset=utf-8");
    setStandardizedExportStatus(`已导出写法统一的表 ${formatInteger(state.rows.length)} 行：${fileName}`, "success");
  } catch (error) {
    setStandardizedExportStatus(`导出失败：${error.message}`, "error");
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

function setStandardizedExportStatus(message, tone = "") {
  if (!dom.standardizedExportStatus) return;
  dom.standardizedExportStatus.className = `inline-status${tone ? ` ${tone}` : ""}`;
  dom.standardizedExportStatus.textContent = message;
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
  exportAnalysisReport("html");
}

function exportMarkdownReport() {
  exportAnalysisReport("md");
}

function exportAnalysisReport(format) {
  if (!canExportReport()) {
    if (state.rows.length && hasFieldConfigDraftChanges()) {
      ensureFieldConfigurationApplied("导出报告");
      setReportExportStatus("字段配置有未应用的修改，请先应用配置后再导出。", "warning");
    } else {
      setStatus("请先完成数据分析，再导出报告。", "warning");
      setReportExportStatus("请先完成数据分析，再导出报告。", "warning");
    }
    updateExportButtons();
    return;
  }

  const buttons = [dom.exportHtmlReport, dom.exportMarkdownReport].filter(Boolean);
  const labels = buttons.map((button) => button.textContent);
  buttons.forEach((button) => {
    button.disabled = true;
  });
  const activeButton = format === "md" ? dom.exportMarkdownReport : dom.exportHtmlReport;
  if (activeButton) activeButton.textContent = "正在导出...";
  setReportExportStatus(`正在生成${format === "html" ? " HTML" : " Markdown"}报告...`, "loading");

  try {
    const exportedAt = new Date();
    const reportData = buildReportData(exportedAt);
    const chartImages = format === "html" ? captureCurrentChartImages() : [];
    const content = format === "html"
      ? buildHtmlReport(reportData, chartImages)
      : buildMarkdownReport(reportData);
    const fileName = buildReportFileName(format, exportedAt);
    const mimeType = format === "html"
      ? "text/html;charset=utf-8"
      : "text/markdown;charset=utf-8";
    downloadUtf8TextFile(content, fileName, mimeType);
    setStatus(`${format === "html" ? "HTML" : "Markdown"} 报告已生成：${fileName}`, "success");
    setReportExportStatus(`报告已生成：${fileName}`, "success");
  } catch (error) {
    setStatus(`报告导出失败：${error.message}`, "error");
    setReportExportStatus(`报告导出失败：${error.message}`, "error");
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
          || REPORT_CHART_TITLES[id]
          || chart.data?.datasets?.[0]?.label
          || "分析图表",
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
    inferredType: profile.inferredType,
    appliedType: profile.type,
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
      name: item.name,
      count: item.count,
      ratio: formatPercent(item.ratio)
    }))
  }));

  return {
    version: "v2.11.0",
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
      filteredOutRows: Math.max(0, state.rows.length - analysisRows.length),
      fields: state.fields.length,
      activeFields: activeProfiles.length,
      numericFields: getProfilesByType("numeric").length,
      categoryFields: getProfilesByType("category").length,
      dateFields: getProfilesByType("date").length,
      idFields: getProfilesByType("id").length,
      ignoredFields: getProfilesByType("ignore").length
    },
    filters: buildReportFilters(),
    fieldTypes,
    quality: {
      totalMissing: state.totalMissing,
      duplicateRows: state.duplicateRows,
      parseWarningCount: state.parseWarnings.length,
      parseWarnings: state.parseWarnings.slice(0, 50).map((warning) => ({
        row: warning.row ?? "",
        message: warning.message || warning.code || warning.type || "读取时有疑问"
      })),
      missingByField: state.profiles.map((profile) => ({
        field: profile.field,
        type: profile.type,
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
        effect: `按空值处理后，${formatPlaceholderChange(entry.field, profile.missingCount, profile.typeKey, preview)}`
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
    const kept = options.filter((option) => keptKeys.has(option.value)).map((option) => option.label);
    const keptValues = kept.slice(0, REPORT_FILTER_MAX_VALUES);
    const more = kept.length > keptValues.length ? `……等 ${formatInteger(kept.length)} 个` : "";
    items.push({
      type: "category",
      field,
      keptCount: kept.length,
      totalCount: options.length,
      keptValues,
      description: `${field}：保留 ${formatInteger(kept.length)} / 共 ${formatInteger(options.length)} 个类别${kept.length ? `（${keptValues.join("、")}${more}）` : ""}`
    });
  }
  if (filters?.numeric) {
    const { field, min, max } = filters.numeric;
    items.push({
      type: "numeric",
      field,
      min,
      max,
      description: `${field}：${formatFieldNumber(field, min)} 至 ${formatFieldNumber(field, max)}`
    });
  }
  if (filters?.date) {
    const { field, start, end } = filters.date;
    items.push({ type: "date", field, start, end, description: `${field}：${start} 至 ${end}` });
  }
  return items;
}

function getReportSourceFileName() {
  if (state.sourceFileName) return state.sourceFileName;
  const sourceTitle = dom.sourceName?.title || dom.sourceName?.textContent || "";
  return sourceTitle.replace(/^已载入\s*·\s+/, "").split(" · ")[0] || "data-report";
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
        summary: `日期字段「${profile.field}」里没有认得出的日期。`
      };
    }

    const firstValue = trend.values[0];
    const lastValue = trend.values[trend.values.length - 1];
    const valueLabel = metricField ? `${formatFieldLabel(metricField)}${getDateChartAggregateLabel(aggregate)}` : "记录数";
    const firstDisplay = metricField
      ? formatFieldNumber(metricField, firstValue)
      : formatInteger(firstValue);
    const lastDisplay = metricField
      ? formatFieldNumber(metricField, lastValue)
      : formatInteger(lastValue);
    let direction = "仅有一个时间点";
    if (trend.values.length > 1) {
      direction = lastValue > firstValue ? "整体上升" : lastValue < firstValue ? "整体下降" : "整体持平";
    }

    return {
      field: profile.field,
      metricField,
      pointCount: trend.labels.length,
      labels: [...trend.labels],
      values: [...trend.values],
      start: trend.labels[0],
      end: trend.labels[trend.labels.length - 1],
      summary: `日期字段「${profile.field}」包含 ${formatInteger(trend.labels.length)} 个时间点；${valueLabel}从 ${trend.labels[0]} 的 ${firstDisplay} 变为 ${trend.labels[trend.labels.length - 1]} 的 ${lastDisplay}，${direction}。`
    };
  });
}

function buildReportCustomAnalysis() {
  const current = state.customAnalysis;

  if (!current || !Array.isArray(current.grouped) || !current.grouped.length) {
    return {
      completed: false,
      message: "当前没有可导出的自定义分组分析结果。"
    };
  }

  const rows = current.grouped.map((item) => ({
    group: item.name,
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
    return `日期字段「${dateField}」没有足够的有效记录生成自定义趋势。`;
  }
  const first = trend.values[0];
  const last = trend.values[trend.values.length - 1];
  if (!Number.isFinite(first) || !Number.isFinite(last)) {
    return `按「${dateField}」观察时，「${metricField}」的聚合结果超出浏览器可表示的有限数值范围，无法可靠判断趋势。`;
  }
  const direction = trend.values.length < 2
    ? "仅有一个时间点"
    : last > first ? "整体上升" : last < first ? "整体下降" : "整体持平";
  return `按「${dateField}」观察，「${metricField}」${getAggregateMethodLabel(aggregateMethod)}从 ${trend.labels[0]} 的 ${formatFieldNumber(metricField, first)} 变为 ${trend.labels[trend.labels.length - 1]} 的 ${formatFieldNumber(metricField, last)}，${direction}。`;
}

function normalizeReportSnapshot(reportData = {}) {
  const quality = reportData.quality || {};
  const dataScale = reportData.dataScale || {};
  return {
    ...reportData,
    version: reportData.version || "v2.11.0",
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
        message: "当前没有可导出的自定义分组分析结果。"
      },
    insights: Array.isArray(reportData.insights) ? reportData.insights : []
  };
}

const REPORT_ID_DUPLICATE_HEADERS = ["字段", "重复值", "出现次数", "文件行号", "说明"];

// Plain-text rows and notes of the report's ID duplicate section: listed values and
// mostly repeated fields as table rows, then overflow and clean-field notes.
function buildReportIdDuplicateSection(idDuplicates) {
  if (!idDuplicates.length) return { rows: [], notes: ["未识别到 ID 字段。"] };
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
    if (overflow) notes.push(`「${item.field}」${overflow}`);
  });
  if (cleanFields.length) notes.push(formatCleanIdFieldsNote(cleanFields, rows.length > 0));
  return { rows, notes };
}

const REPORT_PLACEHOLDER_CONFIRMED_HEADERS = ["字段", "写法", "出现次数", "状态"];
const REPORT_PLACEHOLDER_PENDING_HEADERS = ["字段", "写法", "档位", "出现次数", "文件行号", "按空值处理后的影响"];
const REPORT_PLACEHOLDER_TIER_LABELS = { common: "常见", maybe: "可能是有效值" };

// Plain-text tables and notes of the report's placeholder section. Every note
// starts with fixed text, never with a field name or value.
function buildReportPlaceholderSection({ confirmed, pending }) {
  if (!confirmed.length && !pending.length) return { confirmedRows: [], pendingRows: [], notes: ["没有发现像是空值的写法。"] };
  const notes = formatPlaceholderNotes(pending.map((entry) => entry.key));
  return {
    confirmedRows: confirmed.map((entry) => [
      entry.field,
      formatPlaceholderValue(entry),
      entry.count,
      entry.fieldIgnored ? "已按空值处理（字段已忽略）" : "已按空值处理"
    ]),
    pendingRows: pending.map((entry) => [
      entry.field,
      formatPlaceholderValue(entry),
      REPORT_PLACEHOLDER_TIER_LABELS[entry.tier] || "—",
      entry.count,
      formatRowNumberList(entry.rowNumbers, entry.count),
      entry.effect
    ]),
    notes: notes ? [`说明：${notes}`] : []
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
        title: "疑似合计 / 小计行",
        note: SUSPECTED_TOTAL_NOTE,
        lines: totals.slice(0, SUSPECTED_TOTAL_LIST_LIMIT).map((item) => formatSuspectedTotal(item)),
        overflow: formatSuspectedTotalOverflow(totals.length)
      },
      { title: "公式错误单元格", note: FORMULA_ERROR_NOTE, lines: errors.map((item) => formatFormulaError(item)), overflow: "" }
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
      { title: "日期写法混用", note: DATE_STYLE_NOTE, lines: dateStyles.map(formatDateStyleLine) },
      { title: "整列只有一个值", note: CONSTANT_FIELD_NOTE, lines: constantFields.map(formatConstantField) },
      { title: "同一个值，写法不一样", note: LOOK_ALIKE_NOTE, lines: lookAlikes.map(formatLookAlikeLine) }
    ].filter((section) => section.lines.length)
  };
}

function buildHtmlReport(reportData, chartImages = []) {
  reportData = normalizeReportSnapshot(reportData);
  const fieldLabel = buildReportFieldLabeler(reportData);
  const numberSections = buildReportNumberSections(reportData.quality);
  // The summary counts problems only: numbers read by unit get their section, no count.
  const numericLookingSummary = numberSections.numericLooking.length
    ? `；像是数字的文本列：<strong>${escapeHtml(numberSections.numericLooking.length)}</strong> 个字段`
    : "";
  const numberSectionHtml = [
    ["按单位读取的数字", numberSections.conversions],
    ["像是数字的文本列", numberSections.numericLooking]
  ].filter(([, lines]) => lines.length).map(([title, lines]) => `<h3>${title}</h3>
    <ul class="summary-list">${lines.map((line) => `<li>${escapeHtml(line)}</li>`).join("")}</ul>`).join("\n    ");
  const formatSections = buildReportFormatSections(reportData.quality);
  const formatSummary = `${formatSections.dateStyleCount ? `；日期写法混用：<strong>${escapeHtml(formatSections.dateStyleCount)}</strong> 个字段` : ""}${formatSections.lookAlikeCount ? `；同一个值，写法不一样：<strong>${escapeHtml(formatSections.lookAlikeCount)}</strong> 个字段` : ""}`;
  const formatSectionHtml = formatSections.sections.map(({ title, note, lines }) => `<h3>${title}</h3>
    <p>${escapeHtml(note)}</p>
    <ul class="summary-list">${lines.map((line) => `<li>${escapeHtml(line)}</li>`).join("")}</ul>`).join("\n    ");
  const rowNotes = buildReportRowNoteSections(reportData.quality);
  const rowNoteSummary = `${rowNotes.suspectedTotalCount ? `；疑似合计 / 小计行：<strong>${escapeHtml(rowNotes.suspectedTotalCount)}</strong> 行` : ""}${rowNotes.formulaErrorCount ? `；公式错误：<strong>${escapeHtml(rowNotes.formulaErrorCount)}</strong> 个单元格` : ""}`;
  const rowNoteSectionHtml = rowNotes.sections.map(({ title, note, lines, overflow }) => `<h3>${title}</h3>${note ? `
    <p>${escapeHtml(note)}</p>` : ""}
    <ul class="summary-list">${lines.map((line) => `<li>${escapeHtml(line)}</li>`).join("")}</ul>${overflow ? `
    <p>${escapeHtml(overflow)}</p>` : ""}`).join("\n    ");
  const safeCharts = chartImages.filter((chart) => isSafeChartDataUrl(chart.dataUrl));
  const sheetMeta = reportData.sheetName
    ? `<div><dt>工作表名称</dt><dd>${escapeHtml(reportData.sheetName)}</dd></div>`
    : "";
  const headerNote = formatHeaderDetectionNote(reportData.headerDetection);
  const headerMeta = headerNote
    ? `<div><dt>表头位置</dt><dd>${escapeHtml(headerNote)}</dd></div>`
    : "";
  const idDuplicates = reportData.quality.idDuplicates;
  const idSummary = idDuplicates
    ? `；编号重复：<strong>${escapeHtml(countListedIdDuplicateFields(idDuplicates))}</strong> 个字段`
    : "";
  const idSection = idDuplicates ? buildReportIdDuplicateSection(idDuplicates) : null;
  const placeholders = reportData.quality.placeholders;
  const placeholderSummary = placeholders && (placeholders.confirmed.length || placeholders.pending.length)
    ? `；像是空值的写法：<strong>${escapeHtml(countPendingPlaceholderFields(placeholders.pending))}</strong> 个字段`
    : "";
  const placeholderSection = placeholders ? buildReportPlaceholderSection(placeholders) : null;
  const placeholderSectionHtml = placeholderSection
    ? `<h3>像是空值的写法</h3>
    ${placeholderSection.confirmedRows.length ? `<p>已按空值处理的写法：</p>${buildHtmlReportTable(REPORT_PLACEHOLDER_CONFIRMED_HEADERS, placeholderSection.confirmedRows, "")}` : ""}
    ${placeholderSection.pendingRows.length ? `<p>还没确认的写法：</p>${buildHtmlReportTable(REPORT_PLACEHOLDER_PENDING_HEADERS, placeholderSection.pendingRows, "")}` : ""}
    ${placeholderSection.notes.map((note) => `<p>${escapeHtml(note)}</p>`).join("")}`
    : "";
  const idSectionHtml = idSection
    ? `<h3>编号重复</h3>
    ${idSection.rows.length ? buildHtmlReportTable(REPORT_ID_DUPLICATE_HEADERS, idSection.rows, "") : ""}
    ${idSection.notes.map((note) => `<p>${escapeHtml(note)}</p>`).join("")}`
    : "";
  const categorySections = reportData.categoryStats.length
    ? reportData.categoryStats.map((category) => `
      <section class="subsection">
        <h3>${escapeHtml(category.field)} Top 10</h3>
        ${buildHtmlReportTable(
          ["类别", "数量", "占比"],
          category.top.map((item) => [item.name, item.count, item.ratio]),
          "该分类字段没有可统计的非空值。"
        )}
      </section>
    `).join("")
    : `<p class="empty">无适用的分类字段。</p>`;
  const dateTrendItems = reportData.dateTrends.length
    ? `<ul class="summary-list">${reportData.dateTrends.map((trend) => `<li>${escapeHtml(trend.summary)}</li>`).join("")}</ul>`
    : `<p class="empty">无适用的日期字段。</p>`;
  const customSection = reportData.customAnalysis.completed
    ? `
      <p><strong>指标字段：</strong>${escapeHtml(fieldLabel(reportData.customAnalysis.metricField))}　
      <strong>分组字段：</strong>${escapeHtml(reportData.customAnalysis.groupField)}
      ${reportData.customAnalysis.aggregateLabel ? `　<strong>聚合方式：</strong>${escapeHtml(reportData.customAnalysis.aggregateLabel)}` : ""}
      ${reportData.customAnalysis.dateField ? `　<strong>日期字段：</strong>${escapeHtml(reportData.customAnalysis.dateField)}` : ""}</p>
      ${reportData.customAnalysis.trendSummary ? `<p>${escapeHtml(reportData.customAnalysis.trendSummary)}</p>` : ""}
      ${buildHtmlReportTable(
        [reportData.customAnalysis.groupField, "记录数", "求和", "平均值", "中位数", "最小值", "最大值"],
        reportData.customAnalysis.rows.map((row) => [row.group, row.count, row.sum, row.average, row.median, row.minimum, row.maximum]),
        "当前没有可导出的自定义分组分析结果。"
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
    : `<p class="empty">当前没有可嵌入的图表。</p>`;

  return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${escapeHtml(`${reportData.fileName} - 数据分析报告`)}</title>
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
    <h1>数据分析报告</h1>
    <dl class="meta">
      <div><dt>文件名称</dt><dd>${escapeHtml(reportData.fileName)}</dd></div>
      ${sheetMeta}
      ${headerMeta}
      <div><dt>分析时间</dt><dd>${escapeHtml(reportData.analysisTime)}</dd></div>
      <div><dt>报告生成时间</dt><dd>${escapeHtml(reportData.exportTime)}</dd></div>
    </dl>
  </header>
  <section>
    <h2>数据规模</h2>
    <div class="scale-grid">
      ${[
        ["当前分析行数", reportData.dataScale.rows],
        ["原始数据行数", reportData.dataScale.originalRows],
        ["筛除行数", reportData.dataScale.filteredOutRows],
        ["字段总数", reportData.dataScale.fields],
        ["参与分析字段", reportData.dataScale.activeFields],
        ["数值字段", reportData.dataScale.numericFields],
        ["分类字段", reportData.dataScale.categoryFields],
        ["日期字段", reportData.dataScale.dateFields],
        ["ID 字段", reportData.dataScale.idFields],
        ["忽略字段", reportData.dataScale.ignoredFields]
      ].map(([label, value]) => `<div class="scale-card"><div class="label">${escapeHtml(label)}</div><div class="value">${escapeHtml(value)}</div></div>`).join("")}
    </div>
  </section>
  <section>
    <h2>筛选条件</h2>
    ${reportData.filters.length
      ? buildHtmlReportTable(["字段", "条件"], reportData.filters.map((item) => [item.field, item.description]), REPORT_NO_FILTER_TEXT)
      : `<p>${escapeHtml(REPORT_NO_FILTER_TEXT)}</p>`}
    <p>${escapeHtml(REPORT_QUALITY_SCOPE_TEXT)}</p>
  </section>
  <section>
    <h2>字段类型结果</h2>
    ${buildHtmlReportTable(
      ["字段名", "自动识别类型", "最终类型", "非空值", "唯一值", "空值", "认不出", "示例值"],
      reportData.fieldTypes.map((field) => [
        field.field,
        field.inferredType,
        field.appliedType,
        field.nonMissingCount,
        field.uniqueCount,
        field.missingCount,
        formatConversionFailures(field.conversionFailureCount, field.conversionFailureSamples),
        field.sampleValues.join("、") || "—"
      ]),
      "无字段类型结果。"
    )}
  </section>
  <section>
    <h2>数据质量报告</h2>
    <p>空值总数：<strong>${escapeHtml(reportData.quality.totalMissing)}</strong>；重复行：<strong>${escapeHtml(reportData.quality.duplicateRows)}</strong>；读取时有疑问：<strong>${escapeHtml(reportData.quality.parseWarningCount)}</strong> 处${idSummary}${placeholderSummary}${numericLookingSummary}${rowNoteSummary}${formatSummary}。</p>
    ${rowNoteSectionHtml}
    <h3>各字段空值情况</h3>
    ${buildHtmlReportTable(
      ["字段", "类型", "空值", "空值占比"],
      reportData.quality.missingByField.map((item) => [item.field, item.type, item.missingCount, item.missingRate]),
      "无字段质量结果。"
    )}
    ${numberSectionHtml}
    ${formatSectionHtml}
    ${placeholderSectionHtml}
    <h3>偏大或偏小的值</h3>
    ${buildHtmlReportTable(
      ["字段", "偏大或偏小的个数", "下界", "上界"],
      reportData.quality.outliers.map((item) => [fieldLabel(item.field), item.outlierCount, item.lowerBound, item.upperBound]),
      "未识别到数值字段，暂无可用数值统计。"
    )}
    ${idSectionHtml}
  </section>
  <section>
    <h2>数值字段描述性统计</h2>
    ${buildHtmlReportTable(
      ["字段", "有效数量", "认不出", "平均值", "中位数", "最小值", "最大值", "标准差"],
      reportData.numericStats.map((item) => [fieldLabel(item.field), item.count, item.invalidCount, item.mean, item.median, item.min, item.max, item.standardDeviation]),
      "未识别到数值字段，暂无可用数值统计。"
    )}
  </section>
  <section>
    <h2>分类字段 Top 10</h2>
    ${categorySections}
  </section>
  <section>
    <h2>日期趋势摘要</h2>
    ${dateTrendItems}
  </section>
  <section>
    <h2>自定义分组分析结果</h2>
    ${customSection}
  </section>
  <section>
    <h2>当前图表</h2>
    ${chartSection}
  </section>
  <section>
    <h2>自动分析结论</h2>
    ${reportData.insights.length
      ? `<ol class="insights">${reportData.insights.map((insight) => `<li>${escapeHtml(insight)}</li>`).join("")}</ol>`
      : `<p class="empty">暂无自动分析结论。</p>`}
  </section>
  <footer>由 <a href="https://github.com/wy-data-30/smart-tabular-analyzer">Smart Tabular Analyzer</a> ${escapeHtml(reportData.version)} 在浏览器本地生成；数据未上传。</footer>
</main>
</body>
</html>`;
}

function buildHtmlReportTable(headers, rows, emptyMessage) {
  if (!rows.length) return `<p class="empty">${escapeHtml(emptyMessage)}</p>`;
  return `<div class="table-wrap"><table>
    <thead><tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}</tr></thead>
    <tbody>${rows.map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>`).join("")}</tbody>
  </table></div>`;
}

function buildMarkdownReport(reportData) {
  reportData = normalizeReportSnapshot(reportData);
  const fieldLabel = buildReportFieldLabeler(reportData);
  const numberSections = buildReportNumberSections(reportData.quality);
  const numberSectionLines = [
    ["按单位读取的数字", numberSections.conversions],
    ["像是数字的文本列", numberSections.numericLooking]
  ].filter(([, lines]) => lines.length).flatMap(([title, lines]) => [`### ${title}`, "", ...lines.map((line) => `- ${escapeMarkdownText(line)}`), ""]);
  const formatSections = buildReportFormatSections(reportData.quality);
  const formatSectionLines = formatSections.sections.flatMap(({ title, note, lines }) => [`### ${title}`, "", escapeMarkdownText(note), "", ...lines.map((line) => `- ${escapeMarkdownText(line)}`), ""]);
  const rowNotes = buildReportRowNoteSections(reportData.quality);
  const rowNoteSectionLines = rowNotes.sections.flatMap(({ title, note, lines, overflow }) => [
    `### ${title}`,
    "",
    ...(note ? [escapeMarkdownText(note), ""] : []),
    ...lines.map((line) => `- ${escapeMarkdownText(line)}`),
    "",
    ...(overflow ? [escapeMarkdownText(overflow), ""] : [])
  ]);
  const idDuplicates = reportData.quality.idDuplicates;
  const idSection = idDuplicates ? buildReportIdDuplicateSection(idDuplicates) : null;
  const idSectionLines = idSection
    ? [
      "### 编号重复",
      "",
      ...(idSection.rows.length ? [buildMarkdownTable(REPORT_ID_DUPLICATE_HEADERS, idSection.rows), ""] : []),
      ...idSection.notes.flatMap((note) => [escapeMarkdownText(note), ""])
    ]
    : [];
  const placeholders = reportData.quality.placeholders;
  const placeholderSection = placeholders ? buildReportPlaceholderSection(placeholders) : null;
  const placeholderSectionLines = placeholderSection
    ? [
      "### 像是空值的写法",
      "",
      ...(placeholderSection.confirmedRows.length ? ["已按空值处理的写法：", "", buildMarkdownTable(REPORT_PLACEHOLDER_CONFIRMED_HEADERS, placeholderSection.confirmedRows), ""] : []),
      ...(placeholderSection.pendingRows.length ? ["还没确认的写法：", "", buildMarkdownTable(REPORT_PLACEHOLDER_PENDING_HEADERS, placeholderSection.pendingRows), ""] : []),
      ...placeholderSection.notes.flatMap((note) => [escapeMarkdownText(note), ""])
    ]
    : [];
  const lines = [
    "# 数据分析报告",
    "",
    `- 文件名称：${escapeMarkdownText(reportData.fileName)}`,
    ...(reportData.sheetName ? [`- 工作表名称：${escapeMarkdownText(reportData.sheetName)}`] : []),
    ...(reportData.headerDetection
      ? [`- 表头位置：${escapeMarkdownText(formatHeaderDetectionNote(reportData.headerDetection))}`]
      : []),
    `- 分析时间：${escapeMarkdownText(reportData.analysisTime)}`,
    `- 报告生成时间：${escapeMarkdownText(reportData.exportTime)}`,
    "",
    "## 数据规模",
    "",
    buildMarkdownTable(
      ["指标", "数量"],
      [
        ["当前分析行数", reportData.dataScale.rows],
        ["原始数据行数", reportData.dataScale.originalRows],
        ["筛除行数", reportData.dataScale.filteredOutRows],
        ["字段总数", reportData.dataScale.fields],
        ["参与分析字段", reportData.dataScale.activeFields],
        ["数值字段", reportData.dataScale.numericFields],
        ["分类字段", reportData.dataScale.categoryFields],
        ["日期字段", reportData.dataScale.dateFields],
        ["ID 字段", reportData.dataScale.idFields],
        ["忽略字段", reportData.dataScale.ignoredFields]
      ]
    ),
    "",
    "## 筛选条件",
    "",
    reportData.filters.length
      ? buildMarkdownTable(["字段", "条件"], reportData.filters.map((item) => [item.field, item.description]))
      : escapeMarkdownText(REPORT_NO_FILTER_TEXT),
    "",
    escapeMarkdownText(REPORT_QUALITY_SCOPE_TEXT),
    "",
    "## 字段类型结果",
    "",
    reportData.fieldTypes.length
      ? buildMarkdownTable(
        ["字段名", "自动识别类型", "最终类型", "非空值", "唯一值", "空值", "认不出", "示例值"],
        reportData.fieldTypes.map((field) => [
          field.field,
          field.inferredType,
          field.appliedType,
          field.nonMissingCount,
          field.uniqueCount,
          field.missingCount,
          formatConversionFailures(field.conversionFailureCount, field.conversionFailureSamples),
          field.sampleValues.join("、") || "—"
        ])
      )
      : "无字段类型结果。",
    "",
    "## 数据质量报告",
    "",
    `- 空值总数：${reportData.quality.totalMissing}`,
    `- 重复行：${reportData.quality.duplicateRows}`,
    `- 读取时有疑问：${reportData.quality.parseWarningCount} 处`,
    ...(idDuplicates ? [`- 编号重复：${countListedIdDuplicateFields(idDuplicates)} 个字段`] : []),
    ...(placeholders && (placeholders.confirmed.length || placeholders.pending.length)
      ? [`- 像是空值的写法：${countPendingPlaceholderFields(placeholders.pending)} 个字段`]
      : []),
    ...(numberSections.numericLooking.length ? [`- 像是数字的文本列：${numberSections.numericLooking.length} 个字段`] : []),
    ...(rowNotes.suspectedTotalCount ? [`- 疑似合计 / 小计行：${rowNotes.suspectedTotalCount} 行`] : []),
    ...(rowNotes.formulaErrorCount ? [`- 公式错误：${rowNotes.formulaErrorCount} 个单元格`] : []),
    ...(formatSections.dateStyleCount ? [`- 日期写法混用：${formatSections.dateStyleCount} 个字段`] : []),
    ...(formatSections.lookAlikeCount ? [`- 同一个值，写法不一样：${formatSections.lookAlikeCount} 个字段`] : []),
    "",
    ...rowNoteSectionLines,
    "### 各字段空值情况",
    "",
    reportData.quality.missingByField.length
      ? buildMarkdownTable(
        ["字段", "类型", "空值", "空值占比"],
        reportData.quality.missingByField.map((item) => [item.field, item.type, item.missingCount, item.missingRate])
      )
      : "无字段质量结果。",
    "",
    ...numberSectionLines,
    ...formatSectionLines,
    ...placeholderSectionLines,
    "### 偏大或偏小的值",
    "",
    reportData.quality.outliers.length
      ? buildMarkdownTable(
        ["字段", "偏大或偏小的个数", "下界", "上界"],
        reportData.quality.outliers.map((item) => [fieldLabel(item.field), item.outlierCount, item.lowerBound, item.upperBound])
      )
      : "未识别到数值字段，暂无可用数值统计。",
    "",
    ...idSectionLines,
    "## 数值字段描述性统计",
    "",
    reportData.numericStats.length
      ? buildMarkdownTable(
        ["字段", "有效数量", "认不出", "平均值", "中位数", "最小值", "最大值", "标准差"],
        reportData.numericStats.map((item) => [fieldLabel(item.field), item.count, item.invalidCount, item.mean, item.median, item.min, item.max, item.standardDeviation])
      )
      : "未识别到数值字段，暂无可用数值统计。",
    "",
    "## 分类字段 Top 10",
    ""
  ];

  if (reportData.categoryStats.length) {
    reportData.categoryStats.forEach((category) => {
      lines.push(
        `### ${escapeMarkdownText(category.field)}`,
        "",
        category.top.length
          ? buildMarkdownTable(
            ["类别", "数量", "占比"],
            category.top.map((item) => [item.name, item.count, item.ratio])
          )
          : "该分类字段没有可统计的非空值。",
        ""
      );
    });
  } else {
    lines.push("无适用的分类字段。", "");
  }

  lines.push("## 日期趋势摘要", "");
  if (reportData.dateTrends.length) {
    reportData.dateTrends.forEach((trend) => {
      lines.push(`- ${escapeMarkdownText(trend.summary)}`);
    });
  } else {
    lines.push("无适用的日期字段。");
  }

  lines.push("", "## 自定义分组分析结果", "");
  if (reportData.customAnalysis.completed) {
    lines.push(
      `- 指标字段：${escapeMarkdownText(fieldLabel(reportData.customAnalysis.metricField))}`,
      `- 分组字段：${escapeMarkdownText(reportData.customAnalysis.groupField)}`,
      ...(reportData.customAnalysis.aggregateLabel ? [`- 聚合方式：${escapeMarkdownText(reportData.customAnalysis.aggregateLabel)}`] : []),
      ...(reportData.customAnalysis.dateField ? [`- 日期字段：${escapeMarkdownText(reportData.customAnalysis.dateField)}`] : []),
      ...(reportData.customAnalysis.trendSummary ? [`- ${escapeMarkdownText(reportData.customAnalysis.trendSummary)}`] : []),
      "",
      buildMarkdownTable(
        [reportData.customAnalysis.groupField, "记录数", "求和", "平均值", "中位数", "最小值", "最大值"],
        reportData.customAnalysis.rows.map((row) => [row.group, row.count, row.sum, row.average, row.median, row.minimum, row.maximum])
      )
    );
  } else {
    lines.push(escapeMarkdownText(reportData.customAnalysis.message));
  }

  lines.push("", "## 自动分析结论", "");
  if (reportData.insights.length) {
    reportData.insights.forEach((insight, index) => {
      lines.push(`${index + 1}. ${escapeMarkdownText(insight)}`);
    });
  } else {
    lines.push("暂无自动分析结论。");
  }
  lines.push("", `> 由 [Smart Tabular Analyzer](https://github.com/wy-data-30/smart-tabular-analyzer) ${escapeMarkdownText(reportData.version)} 在浏览器本地生成；数据未上传。`, "");
  return lines.join("\n");
}

function buildMarkdownTable(headers, rows) {
  const headerLine = `| ${headers.map(escapeMarkdownCell).join(" | ")} |`;
  const separatorLine = `| ${headers.map(() => "---").join(" | ")} |`;
  const bodyLines = rows.map((row) => `| ${row.map(escapeMarkdownCell).join(" | ")} |`);
  return [headerLine, separatorLine, ...bodyLines].join("\n");
}

function escapeMarkdownCell(value) {
  // Escape ASCII punctuation so table data cannot become inline Markdown,
  // autolinks or raw HTML. Add only our own line-break markup afterwards.
  return String(value ?? "")
    .replace(/[!-/:-@\[-`{-~]/g, (character) => {
      if (character === "&") return "&amp;";
      if (character === "<") return "&lt;";
      if (character === ">") return "&gt;";
      return `\\${character}`;
    })
    .replace(/\r\n|[\r\n]/g, "<br>");
}

function escapeMarkdownText(value) {
  return String(value ?? "")
    .replace(/\\/g, "\\\\")
    .replace(/\|/g, "\\|")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/([*_`#[\]])/g, "\\$1")
    // Add our own line-break markup only after user text is escaped.
    .replace(/\r\n|[\r\n]/g, "<br>");
}

function buildReportFileName(extension, now = new Date()) {
  const safeExtension = String(extension || "html").replace(/^\./, "").toLowerCase() === "md"
    ? "md"
    : "html";
  const originalName = getReportSourceFileName().split(/[\\/]/).pop() || "data-report";
  const stem = originalName.replace(/\.(?:csv|xlsx|xls)$/i, "") || "data-report";
  return `${sanitizeFileName(stem)}-${formatReportFileTimestamp(now)}.${safeExtension}`;
}

function buildDataExportFileName(mode, now = new Date()) {
  const suffix = mode === "deduplicated" ? "deduplicated" : "filtered";
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

function sanitizeFileName(value) {
  return String(value || "data-report")
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, "-")
    .replace(/[\u202A-\u202E\u2066-\u2069]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 80) || "data-report";
}

function setReportExportStatus(message, tone = "") {
  if (!dom.reportExportStatus) return;
  dom.reportExportStatus.className = `inline-status${tone ? ` ${tone}` : ""}`;
  dom.reportExportStatus.textContent = message;
}

function setProblemRowsExportStatus(message, tone = "") {
  if (!dom.problemRowsExportStatus) return;
  dom.problemRowsExportStatus.className = `inline-status${tone ? ` ${tone}` : ""}`;
  dom.problemRowsExportStatus.textContent = message;
}

function setDataExportStatus(message, tone = "") {
  if (!dom.dataExportStatus) return;
  dom.dataExportStatus.className = `inline-status${tone ? ` ${tone}` : ""}`;
  dom.dataExportStatus.textContent = message;
}
