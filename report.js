// 报告与导出：报告快照、HTML / Markdown 报告生成、处理后 CSV 导出和问题行 CSV 导出、导出门禁、文件名和下载。

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
  missing: "缺失",
  conversion: "格式问题",
  outlier: "异常值",
  duplicateRows: "整行重复",
  idDuplicate: "ID重复",
  placeholder: "疑似缺失值",
  parseWarnings: "解析警告"
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
    version: "v2.6.1",
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
        message: warning.message || warning.code || warning.type || "CSV 解析警告"
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
      numericLooking: state.numericLooking.map((item) => ({ ...item, samples: [...item.samples] }))
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
        effect: `确认后${formatPlaceholderChange(entry.field, profile.missingCount, profile.typeKey, preview)}`
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
        summary: `日期字段「${profile.field}」没有可安全转换的有效日期。`
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
    version: reportData.version || "v2.6.1",
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
      numericLooking: Array.isArray(quality.numericLooking) ? quality.numericLooking : null
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
const REPORT_PLACEHOLDER_PENDING_HEADERS = ["字段", "写法", "档位", "出现次数", "文件行号", "确认后的影响"];
const REPORT_PLACEHOLDER_TIER_LABELS = { common: "常见", maybe: "可能是有效值" };

// Plain-text tables and notes of the report's placeholder section. Every note
// starts with fixed text, never with a field name or value.
function buildReportPlaceholderSection({ confirmed, pending }) {
  if (!confirmed.length && !pending.length) return { confirmedRows: [], pendingRows: [], notes: ["没有发现疑似缺失值。"] };
  const notes = formatPlaceholderNotes(pending.map((entry) => entry.key));
  return {
    confirmedRows: confirmed.map((entry) => [
      entry.field,
      formatPlaceholderValue(entry),
      entry.count,
      entry.fieldIgnored ? "已按缺失处理（字段已忽略）" : "已按缺失处理"
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

function buildHtmlReport(reportData, chartImages = []) {
  reportData = normalizeReportSnapshot(reportData);
  const fieldLabel = buildReportFieldLabeler(reportData);
  const numberSections = buildReportNumberSections(reportData.quality);
  // The summary counts problems only: numbers read by unit get their section, no count.
  const numericLookingSummary = numberSections.numericLooking.length
    ? `；疑似数值列：<strong>${escapeHtml(numberSections.numericLooking.length)}</strong> 个字段`
    : "";
  const numberSectionHtml = [
    ["按单位读取的数字", numberSections.conversions],
    ["疑似数值列", numberSections.numericLooking]
  ].filter(([, lines]) => lines.length).map(([title, lines]) => `<h3>${title}</h3>
    <ul class="summary-list">${lines.map((line) => `<li>${escapeHtml(line)}</li>`).join("")}</ul>`).join("\n    ");
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
    ? `；ID 重复字段：<strong>${escapeHtml(countListedIdDuplicateFields(idDuplicates))}</strong>`
    : "";
  const idSection = idDuplicates ? buildReportIdDuplicateSection(idDuplicates) : null;
  const placeholders = reportData.quality.placeholders;
  const placeholderSummary = placeholders && (placeholders.confirmed.length || placeholders.pending.length)
    ? `；疑似缺失值：<strong>${escapeHtml(countPendingPlaceholderFields(placeholders.pending))}</strong> 个字段`
    : "";
  const placeholderSection = placeholders ? buildReportPlaceholderSection(placeholders) : null;
  const placeholderSectionHtml = placeholderSection
    ? `<h3>疑似缺失值</h3>
    ${placeholderSection.confirmedRows.length ? `<p>已按缺失处理的写法：</p>${buildHtmlReportTable(REPORT_PLACEHOLDER_CONFIRMED_HEADERS, placeholderSection.confirmedRows, "")}` : ""}
    ${placeholderSection.pendingRows.length ? `<p>还没确认的疑似值：</p>${buildHtmlReportTable(REPORT_PLACEHOLDER_PENDING_HEADERS, placeholderSection.pendingRows, "")}` : ""}
    ${placeholderSection.notes.map((note) => `<p>${escapeHtml(note)}</p>`).join("")}`
    : "";
  const idSectionHtml = idSection
    ? `<h3>ID 字段重复值</h3>
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
      ["字段名", "自动识别类型", "最终类型", "非空值", "唯一值", "缺失值", "转换失败", "示例值"],
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
    <p>缺失值总数：<strong>${escapeHtml(reportData.quality.totalMissing)}</strong>；重复行：<strong>${escapeHtml(reportData.quality.duplicateRows)}</strong>；解析警告：<strong>${escapeHtml(reportData.quality.parseWarningCount)}</strong>${idSummary}${placeholderSummary}${numericLookingSummary}。</p>
    <h3>各字段缺失情况</h3>
    ${buildHtmlReportTable(
      ["字段", "类型", "缺失值", "缺失率"],
      reportData.quality.missingByField.map((item) => [item.field, item.type, item.missingCount, item.missingRate]),
      "无字段质量结果。"
    )}
    ${numberSectionHtml}
    ${placeholderSectionHtml}
    <h3>数值异常值</h3>
    ${buildHtmlReportTable(
      ["字段", "异常值数量", "下界", "上界"],
      reportData.quality.outliers.map((item) => [fieldLabel(item.field), item.outlierCount, item.lowerBound, item.upperBound]),
      "未识别到数值字段，暂无可用数值统计。"
    )}
    ${idSectionHtml}
  </section>
  <section>
    <h2>数值字段描述性统计</h2>
    ${buildHtmlReportTable(
      ["字段", "有效数量", "无法转换", "平均值", "中位数", "最小值", "最大值", "标准差"],
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
    ["疑似数值列", numberSections.numericLooking]
  ].filter(([, lines]) => lines.length).flatMap(([title, lines]) => [`### ${title}`, "", ...lines.map((line) => `- ${escapeMarkdownText(line)}`), ""]);
  const idDuplicates = reportData.quality.idDuplicates;
  const idSection = idDuplicates ? buildReportIdDuplicateSection(idDuplicates) : null;
  const idSectionLines = idSection
    ? [
      "### ID 字段重复值",
      "",
      ...(idSection.rows.length ? [buildMarkdownTable(REPORT_ID_DUPLICATE_HEADERS, idSection.rows), ""] : []),
      ...idSection.notes.flatMap((note) => [escapeMarkdownText(note), ""])
    ]
    : [];
  const placeholders = reportData.quality.placeholders;
  const placeholderSection = placeholders ? buildReportPlaceholderSection(placeholders) : null;
  const placeholderSectionLines = placeholderSection
    ? [
      "### 疑似缺失值",
      "",
      ...(placeholderSection.confirmedRows.length ? ["已按缺失处理的写法：", "", buildMarkdownTable(REPORT_PLACEHOLDER_CONFIRMED_HEADERS, placeholderSection.confirmedRows), ""] : []),
      ...(placeholderSection.pendingRows.length ? ["还没确认的疑似值：", "", buildMarkdownTable(REPORT_PLACEHOLDER_PENDING_HEADERS, placeholderSection.pendingRows), ""] : []),
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
        ["字段名", "自动识别类型", "最终类型", "非空值", "唯一值", "缺失值", "转换失败", "示例值"],
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
    `- 缺失值总数：${reportData.quality.totalMissing}`,
    `- 重复行：${reportData.quality.duplicateRows}`,
    `- 解析警告：${reportData.quality.parseWarningCount}`,
    ...(idDuplicates ? [`- ID 重复字段：${countListedIdDuplicateFields(idDuplicates)}`] : []),
    ...(placeholders && (placeholders.confirmed.length || placeholders.pending.length)
      ? [`- 疑似缺失值：${countPendingPlaceholderFields(placeholders.pending)} 个字段`]
      : []),
    ...(numberSections.numericLooking.length ? [`- 疑似数值列：${numberSections.numericLooking.length} 个字段`] : []),
    "",
    "### 各字段缺失情况",
    "",
    reportData.quality.missingByField.length
      ? buildMarkdownTable(
        ["字段", "类型", "缺失值", "缺失率"],
        reportData.quality.missingByField.map((item) => [item.field, item.type, item.missingCount, item.missingRate])
      )
      : "无字段质量结果。",
    "",
    ...numberSectionLines,
    ...placeholderSectionLines,
    "### 数值异常值",
    "",
    reportData.quality.outliers.length
      ? buildMarkdownTable(
        ["字段", "异常值数量", "下界", "上界"],
        reportData.quality.outliers.map((item) => [fieldLabel(item.field), item.outlierCount, item.lowerBound, item.upperBound])
      )
      : "未识别到数值字段，暂无可用数值统计。",
    "",
    ...idSectionLines,
    "## 数值字段描述性统计",
    "",
    reportData.numericStats.length
      ? buildMarkdownTable(
        ["字段", "有效数量", "无法转换", "平均值", "中位数", "最小值", "最大值", "标准差"],
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
