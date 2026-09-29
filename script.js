const state = {
  rows: [],
  filteredRows: null,
  filterSourceRows: null,
  fields: [],
  profiles: [],
  numericStats: [],
  qualityNumericStats: [],
  categoryStats: [],
  duplicateRows: 0,
  totalMissing: 0,
  parseWarnings: [],
  fieldTypeDrafts: Object.create(null),
  pendingExcel: null,
  activeFileReader: null,
  activeImportId: 0,
  charts: {},
  sourceFileName: "",
  sourceSheetName: "",
  sourceType: "",
  analysisCompletedAt: null,
  insights: [],
  customAnalysis: null,
  filters: {
    category: null,
    numeric: null,
    date: null
  },
  categoryFilterOptions: [],
  categoryFilterDraftValues: new Set()
};

const dom = {
  fileInput: document.getElementById("csvFile"),
  fileEncoding: document.getElementById("fileEncoding"),
  loadSample: document.getElementById("loadSample"),
  dropZone: document.getElementById("dropZone"),
  statusPanel: document.getElementById("statusPanel"),
  sheetSelectionPanel: document.getElementById("sheetSelectionPanel"),
  excelWorkbookName: document.getElementById("excelWorkbookName"),
  excelSheetSelect: document.getElementById("excelSheetSelect"),
  confirmSheetSelection: document.getElementById("confirmSheetSelection"),
  sheetSelectionMessage: document.getElementById("sheetSelectionMessage"),
  results: document.getElementById("results"),
  reportTitle: document.getElementById("reportTitle"),
  sourceName: document.getElementById("sourceName"),
  mobileSectionNav: document.getElementById("mobileSectionNav"),
  globalFilterStatus: document.getElementById("globalFilterStatus"),
  globalFilterSummary: document.getElementById("globalFilterSummary"),
  globalClearFilters: document.getElementById("globalClearFilters"),
  overviewCards: document.getElementById("overviewCards"),
  insights: document.getElementById("insights"),
  previewTable: document.getElementById("previewTable"),
  schemaTable: document.getElementById("schemaTable"),
  fieldConfigMessage: document.getElementById("fieldConfigMessage"),
  applyFieldConfig: document.getElementById("applyFieldConfig"),
  restoreAutoFieldTypes: document.getElementById("restoreAutoFieldTypes"),
  categoryFilterField: document.getElementById("categoryFilterField"),
  categoryFilterControls: document.getElementById("categoryFilterControls"),
  categoryFilterSearch: document.getElementById("categoryFilterSearch"),
  categoryFilterValues: document.getElementById("categoryFilterValues"),
  categoryFilterEmpty: document.getElementById("categoryFilterEmpty"),
  categoryFilterSelectionCount: document.getElementById("categoryFilterSelectionCount"),
  selectAllCategoryValues: document.getElementById("selectAllCategoryValues"),
  clearCategoryValues: document.getElementById("clearCategoryValues"),
  numericFilterField: document.getElementById("numericFilterField"),
  numericFilterMin: document.getElementById("numericFilterMin"),
  numericFilterMax: document.getElementById("numericFilterMax"),
  numericFilterRange: document.getElementById("numericFilterRange"),
  dateFilterField: document.getElementById("dateFilterField"),
  dateFilterStart: document.getElementById("dateFilterStart"),
  dateFilterEnd: document.getElementById("dateFilterEnd"),
  dateFilterRange: document.getElementById("dateFilterRange"),
  filterOriginalCount: document.getElementById("filterOriginalCount"),
  filterCurrentCount: document.getElementById("filterCurrentCount"),
  filterExcludedCount: document.getElementById("filterExcludedCount"),
  filterMessage: document.getElementById("filterMessage"),
  applyFilters: document.getElementById("applyFilters"),
  clearAllFilters: document.getElementById("clearAllFilters"),
  qualitySummary: document.getElementById("qualitySummary"),
  qualityTable: document.getElementById("qualityTable"),
  outlierTable: document.getElementById("outlierTable"),
  numericStatsTable: document.getElementById("numericStatsTable"),
  categoryStats: document.getElementById("categoryStats"),
  numericChartField: document.getElementById("numericChartField"),
  categoryChartField: document.getElementById("categoryChartField"),
  dateChartField: document.getElementById("dateChartField"),
  numericChartTitle: document.getElementById("numericChartTitle"),
  categoryChartTitle: document.getElementById("categoryChartTitle"),
  dateChartTitle: document.getElementById("dateChartTitle"),
  metricField: document.getElementById("metricField"),
  groupField: document.getElementById("groupField"),
  aggregateMethod: document.getElementById("aggregateMethod"),
  v2MetricField: document.getElementById("v2MetricField"),
  v2GroupField: document.getElementById("v2GroupField"),
  v2AggregateMethod: document.getElementById("v2AggregateMethod"),
  v2DateField: document.getElementById("v2DateField"),
  v2AnalyzeButton: document.getElementById("v2AnalyzeButton"),
  v2FieldPrompt: document.getElementById("v2FieldPrompt"),
  v2AnalysisResults: document.getElementById("v2AnalysisResults"),
  v2SummaryCards: document.getElementById("v2SummaryCards"),
  v2ResultSummary: document.getElementById("v2ResultSummary"),
  v2PrimaryTable: document.getElementById("v2PrimaryTable"),
  v2PrimaryTableTitle: document.getElementById("v2PrimaryTableTitle"),
  v2TopGroupTitle: document.getElementById("v2TopGroupTitle"),
  v2TrendTitle: document.getElementById("v2TrendTitle"),
  v2SumTable: document.getElementById("v2SumTable"),
  v2AvgTable: document.getElementById("v2AvgTable"),
  v2TrendCard: document.getElementById("v2TrendCard"),
  scenarioTemplate: document.getElementById("scenarioTemplate"),
  templateMappingForm: document.getElementById("templateMappingForm"),
  runTemplateAnalysis: document.getElementById("runTemplateAnalysis"),
  templatePrompt: document.getElementById("templatePrompt"),
  templateResults: document.getElementById("templateResults"),
  templateSummaryCards: document.getElementById("templateSummaryCards"),
  templateInsights: document.getElementById("templateInsights"),
  templatePrimaryTable: document.getElementById("templatePrimaryTable"),
  templateMainChartTitle: document.getElementById("templateMainChartTitle"),
  templateSecondaryChartTitle: document.getElementById("templateSecondaryChartTitle"),
  templateTrendCard: document.getElementById("templateTrendCard"),
  templateTrendChartTitle: document.getElementById("templateTrendChartTitle"),
  customAnalysisChartTitle: document.getElementById("customAnalysisChartTitle"),
  dataExportOriginalCount: document.getElementById("dataExportOriginalCount"),
  dataExportFilteredCount: document.getElementById("dataExportFilteredCount"),
  dataExportDeduplicatedCount: document.getElementById("dataExportDeduplicatedCount"),
  exportFilteredCsv: document.getElementById("exportFilteredCsv"),
  exportDeduplicatedCsv: document.getElementById("exportDeduplicatedCsv"),
  dataExportStatus: document.getElementById("dataExportStatus"),
  exportHtmlReport: document.getElementById("exportHtmlReport"),
  exportMarkdownReport: document.getElementById("exportMarkdownReport"),
  reportExportStatus: document.getElementById("reportExportStatus"),
  workflowStepUpload: document.getElementById("workflowStepUpload"),
  workflowStepSchema: document.getElementById("workflowStepSchema"),
  workflowStepAnalysis: document.getElementById("workflowStepAnalysis"),
  workflowStepExport: document.getElementById("workflowStepExport")
};

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

const FIELD_TYPE_LABELS = {
  numeric: "数值",
  category: "分类",
  date: "日期",
  id: "ID",
  ignore: "忽略"
};

const DATE_TIME_SUFFIX_PATTERN = "(?:[ T](?:[01]?\\d|2[0-3]):[0-5]\\d(?::[0-5]\\d(?:\\.\\d{1,3})?)?(?:Z|[+-]\\d{2}:?\\d{2})?)?";

const CHART_THEME = {
  primary: "#625bde",
  primaryFill: "rgba(98, 91, 222, 0.12)",
  secondary: "#0f9f8f",
  secondaryFill: "rgba(15, 159, 143, 0.12)",
  warning: "#e09a32",
  text: "#697084",
  grid: "rgba(211, 215, 227, 0.72)"
};

if (window.Chart) {
  Chart.defaults.color = CHART_THEME.text;
  Chart.defaults.font.family = 'Inter, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif';
  Chart.defaults.font.size = 11;
  Chart.defaults.elements.bar.borderRadius = 6;
  Chart.defaults.elements.bar.borderSkipped = false;
  Chart.defaults.elements.line.borderWidth = 2.25;
  Chart.defaults.elements.point.radius = 0;
  Chart.defaults.elements.point.hoverRadius = 4;
}

const TEMPLATE_DEFINITIONS = {
  generic: {
    name: "通用数据分析",
    fields: [
      { key: "metric", label: "分析指标字段", required: true, expected: "numeric" },
      { key: "group", label: "分组维度字段", required: true, expected: "category" },
      { key: "date", label: "时间字段（可选）", required: false, expected: "date" }
    ]
  },
  sales: {
    name: "销售数据分析",
    fields: [
      { key: "salesAmount", label: "销售额字段（可选）", required: false, expected: "numeric" },
      { key: "unitPrice", label: "单价字段（可选）", required: false, expected: "numeric" },
      { key: "quantity", label: "数量字段（可选）", required: false, expected: "numeric" },
      { key: "category", label: "品类字段", required: true, expected: "category" },
      { key: "region", label: "地区字段", required: true, expected: "category" },
      { key: "date", label: "日期字段（可选）", required: false, expected: "date" }
    ]
  },
  score: {
    name: "成绩数据分析",
    fields: [
      { key: "score", label: "成绩字段", required: true, expected: "numeric" },
      { key: "class", label: "班级字段（可选）", required: false, expected: "category" },
      { key: "subject", label: "科目字段（可选）", required: false, expected: "category" },
      { key: "examDate", label: "考试日期字段（可选）", required: false, expected: "date" }
    ]
  },
  usedGoods: {
    name: "二手商品价格分析",
    fields: [
      { key: "price", label: "价格字段", required: true, expected: "numeric" },
      { key: "item", label: "商品字段（可选）", required: false, expected: "category" },
      { key: "platform", label: "平台字段（可选）", required: false, expected: "category" },
      { key: "city", label: "城市字段（可选）", required: false, expected: "category" },
      { key: "condition", label: "成色字段（可选）", required: false, expected: "category" },
      { key: "date", label: "日期字段（可选）", required: false, expected: "date" }
    ]
  },
  survey: {
    name: "问卷调查分析",
    fields: [
      { key: "satisfaction", label: "满意度字段（可选）", required: false, expected: "numeric" },
      { key: "age", label: "年龄字段（可选）", required: false, expected: "numeric" },
      { key: "gender", label: "性别字段（可选）", required: false, expected: "category" },
      { key: "city", label: "城市字段（可选）", required: false, expected: "category" },
      { key: "choice", label: "选择题字段（可选）", required: false, expected: "category" }
    ]
  },
  behavior: {
    name: "用户行为分析",
    fields: [
      { key: "userId", label: "用户 ID 字段", required: true, expected: "any" },
      { key: "event", label: "事件字段", required: true, expected: "category" },
      { key: "channel", label: "渠道字段（可选）", required: false, expected: "category" },
      { key: "device", label: "设备字段（可选）", required: false, expected: "category" },
      { key: "date", label: "日期字段（可选）", required: false, expected: "date" }
    ]
  }
};

setupResultNavigation();
setupWorkspaceNavigation();

dom.fileInput.addEventListener("change", (event) => {
  const file = event.target.files[0];
  if (file) parseDataFile(file);
  event.target.value = "";
});

dom.loadSample.addEventListener("click", async () => {
  const importId = beginImport();
  try {
    setStatus("正在加载 sample-data.csv ...");
    const response = await fetch("examples/sample-data.csv");
    if (!response.ok) throw new Error("示例数据加载失败");
    const text = await response.text();
    if (!isCurrentImport(importId)) return;
    parseCsvText(text, "sample-data.csv", importId);
  } catch (error) {
    showError(error.message, importId);
  }
});

["dragenter", "dragover"].forEach((eventName) => {
  dom.dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dom.dropZone.classList.add("dragging");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  dom.dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dom.dropZone.classList.remove("dragging");
  });
});

dom.dropZone.addEventListener("drop", (event) => {
  const file = event.dataTransfer.files[0];
  if (!file) return;
  parseDataFile(file);
});

[
  dom.numericChartField,
  dom.categoryChartField,
  dom.dateChartField,
  dom.metricField,
  dom.groupField,
  dom.aggregateMethod
].forEach((control) => control.addEventListener("change", renderCharts));

dom.v2AnalyzeButton.addEventListener("click", renderV2Analysis);
dom.globalClearFilters.addEventListener("click", clearAllFilters);

dom.scenarioTemplate.addEventListener("change", () => {
  renderTemplateMappingForm();
  resetTemplateResults(`已切换到「${getCurrentTemplate().name}」，请选择字段并生成分析。`);
});

dom.runTemplateAnalysis.addEventListener("click", runTemplateAnalysis);
dom.exportFilteredCsv.addEventListener("click", () => exportProcessedCsv("filtered"));
dom.exportDeduplicatedCsv.addEventListener("click", () => exportProcessedCsv("deduplicated"));
dom.exportHtmlReport.addEventListener("click", exportHtmlReport);
dom.exportMarkdownReport.addEventListener("click", exportMarkdownReport);
dom.confirmSheetSelection.addEventListener("click", analyzeSelectedExcelSheet);
dom.excelSheetSelect.addEventListener("change", () => {
  dom.sheetSelectionMessage.className = "sheet-selection-message";
  dom.sheetSelectionMessage.textContent = "";
});

dom.schemaTable.addEventListener("change", (event) => {
  const select = event.target.closest("select[data-field-type]");
  if (!select) return;
  stageFieldTypeChange(select.dataset.fieldType, select.value);
});
dom.applyFieldConfig.addEventListener("click", applyFieldConfiguration);
dom.restoreAutoFieldTypes.addEventListener("click", restoreAutomaticFieldTypes);
dom.categoryFilterField.addEventListener("change", updateCategoryFilterField);
dom.categoryFilterSearch.addEventListener("input", renderCategoryFilterValues);
dom.categoryFilterValues.addEventListener("change", updateCategoryFilterDraftSelection);
dom.selectAllCategoryValues.addEventListener("click", () => setAllCategoryFilterValues(true));
dom.clearCategoryValues.addEventListener("click", () => setAllCategoryFilterValues(false));
dom.numericFilterField.addEventListener("change", updateNumericFilterField);
dom.dateFilterField.addEventListener("change", updateDateFilterField);
dom.applyFilters.addEventListener("click", applyFilters);
dom.clearAllFilters.addEventListener("click", clearAllFilters);
updateExportButtons();

[dom.v2MetricField, dom.v2GroupField, dom.v2DateField].forEach((control) => {
  control.addEventListener("change", () => {
    if (state.rows.length) {
      dom.v2FieldPrompt.className = "analysis-prompt";
      dom.v2FieldPrompt.textContent = "字段已更新，点击“生成自定义分析”生成新的结果。";
    }
  });
});

function setupResultNavigation() {
  const links = Array.from(document.querySelectorAll(".results-nav nav a"));
  const sections = links
    .map((link) => document.querySelector(link.getAttribute("href")))
    .filter(Boolean);

  const setActiveLink = (sectionId) => {
    links.forEach((link) => {
      const isActive = link.getAttribute("href") === `#${sectionId}`;
      link.classList.toggle("active", isActive);
      if (isActive) {
        link.setAttribute("aria-current", "location");
      } else {
        link.removeAttribute("aria-current");
      }
    });
    if (dom.mobileSectionNav && dom.mobileSectionNav.value !== sectionId) {
      dom.mobileSectionNav.value = sectionId;
    }
  };

  const navigateToSection = (sectionId) => {
    const section = document.getElementById(sectionId);
    if (!section) return;
    setActiveLink(sectionId);
    window.history.replaceState(null, "", `#${sectionId}`);
    section.scrollIntoView({ behavior: "smooth", block: "start" });
  };

  links.forEach((link) => {
    link.addEventListener("click", (event) => {
      event.preventDefault();
      navigateToSection(link.getAttribute("href").slice(1));
    });
  });

  dom.mobileSectionNav?.addEventListener("change", () => {
    navigateToSection(dom.mobileSectionNav.value);
  });

  let navigationFrame = null;
  const updateActiveLinkFromScroll = () => {
    navigationFrame = null;
    const isAtPageEnd = window.innerHeight + window.scrollY >= document.documentElement.scrollHeight - 4;
    if (isAtPageEnd && sections.length) {
      setActiveLink(sections[sections.length - 1].id);
      return;
    }

    const navigationHeight = document.querySelector(".results-nav")?.getBoundingClientRect?.().height || 0;
    const readingLine = Math.max(150, Math.min(320, window.innerHeight * 0.35, navigationHeight + 130));
    const currentSection = sections.reduce((current, section) => (
      section.getBoundingClientRect().top <= readingLine ? section : current
    ), sections[0]);
    if (currentSection) setActiveLink(currentSection.id);
  };

  if (typeof window.addEventListener === "function") {
    window.addEventListener("scroll", () => {
      if (navigationFrame !== null) return;
      navigationFrame = window.requestAnimationFrame(updateActiveLinkFromScroll);
    }, { passive: true });
  }

  if (sections[0]) setActiveLink(sections[0].id);
}

function setupWorkspaceNavigation() {
  document.querySelectorAll(".header-cta, .nav-new-analysis").forEach((link) => {
    link.addEventListener("click", () => {
      document.body.classList.remove("has-results");
      dom.results.classList.add("hidden");
      dom.statusPanel.className = "status-panel hidden";
      dom.statusPanel.textContent = "";
      dom.statusPanel.removeAttribute("aria-busy");
      setImportBusy(false);
    });
  });
}

function setImportBusy(isBusy) {
  document.body.classList.toggle("is-loading", isBusy);
  dom.fileInput.disabled = isBusy;
  dom.loadSample.disabled = isBusy;
  dom.confirmSheetSelection.disabled = isBusy;
  dom.dropZone.setAttribute("aria-busy", String(isBusy));
}

function showResultsWorkspace() {
  document.body.classList.add("has-results");
  setImportBusy(false);
  updateWorkflowSteps();
  window.requestAnimationFrame(() => {
    dom.reportTitle?.focus({ preventScroll: true });
    dom.results.scrollIntoView({ block: "start" });
  });
}

function updateWorkflowSteps() {
  if (!dom.workflowStepUpload) return;

  const setStep = (element, stateKey, label) => {
    if (!element) return;
    element.dataset.state = stateKey;
    const stateLabel = typeof element.querySelector === "function"
      ? element.querySelector(".step-state")
      : null;
    if (stateLabel) stateLabel.textContent = label;
  };

  if (!state.rows.length) {
    setStep(dom.workflowStepUpload, "current", "等待文件");
    setStep(dom.workflowStepSchema, "pending", "等待上传");
    setStep(dom.workflowStepAnalysis, "pending", "等待配置");
    setStep(dom.workflowStepExport, "pending", "等待分析");
    return;
  }

  const hasDraftChanges = hasFieldConfigDraftChanges();
  setStep(dom.workflowStepUpload, "complete", "已完成");

  if (hasDraftChanges) {
    setStep(dom.workflowStepSchema, "current", "有未应用修改");
    setStep(dom.workflowStepAnalysis, "pending", "等待重新计算");
    setStep(dom.workflowStepExport, "pending", "暂不可导出");
    return;
  }

  setStep(
    dom.workflowStepSchema,
    "complete",
    hasAutomaticTypeDifferences() ? "配置已应用" : "自动识别完成"
  );
  setStep(dom.workflowStepAnalysis, state.analysisCompletedAt ? "complete" : "current", state.analysisCompletedAt ? "已生成" : "正在生成");
  setStep(dom.workflowStepExport, canExportReport() ? "current" : "pending", canExportReport() ? "可导出" : "等待分析");
}

async function parseCsvFile(file) {
  const importId = beginImport();
  try {
    assertRuntimeDependency("Papa", "CSV 解析组件加载失败，请检查网络连接后刷新页面重试。");
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
    assertRuntimeDependency("XLSX", "Excel 解析组件加载失败，请检查网络连接后刷新页面重试。");
    setStatus(`正在读取 ${file.name} ...`);
    const buffer = await readFileAsArrayBuffer(file, importId);
    if (!isCurrentImport(importId)) return;
    setStatus(`正在解析 ${file.name} ...`);
    validateExcelFileSignature(buffer, file.name);
    let workbook;
    try {
      workbook = XLSX.read(buffer, { type: "array", cellDates: true });
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
  const { fields, rows } = buildExcelTabularData(rawRows, rawRows, sheetName);

  state.pendingExcel = null;
  dom.sheetSelectionPanel.classList.add("hidden");
  setStatus(`正在分析 ${fileName} · ${sheetName} ...`);
  commitTabularData(fields, rows, fileName, importId, [], sheetName);
}

function buildExcelTabularData(formattedRows, rawRows, sheetName = "当前工作表") {
  const rowCount = Math.max(formattedRows?.length || 0, rawRows?.length || 0);
  const rowPairs = Array.from({ length: rowCount }, (_, index) => ({
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

  return { fields, rows };
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
  assertRuntimeDependency("Papa", "CSV 解析组件加载失败，请检查网络连接后刷新页面重试。");
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

function commitTabularData(fields, rows, sourceName, importId, parseWarnings = [], sheetName = "") {
  if (!isCurrentImport(importId)) return;
  if (!fields.length || !rows.length) {
    showError("文件中没有可分析的数据。请确认数据包含表头和至少一行数据。", importId);
    return;
  }
  if (fields.length > IMPORT_LIMITS.maxColumns || rows.length > IMPORT_LIMITS.maxRows) {
    showError(`数据超过 ${formatInteger(IMPORT_LIMITS.maxRows)} 行或 ${formatInteger(IMPORT_LIMITS.maxColumns)} 列限制，请拆分后重试。`, importId);
    return;
  }

  state.rows = rows;
  state.filteredRows = rows.slice();
  state.filterSourceRows = state.rows;
  state.fields = fields;
  state.parseWarnings = parseWarnings;
  state.pendingExcel = null;
  state.sourceFileName = sourceName;
  state.sourceSheetName = sheetName;
  state.sourceType = sheetName ? "excel" : "csv";
  state.analysisCompletedAt = new Date();
  state.insights = [];
  state.customAnalysis = null;
  state.filters = createEmptyFilterConfig();
  state.categoryFilterOptions = [];
  state.categoryFilterDraftValues = new Set();
  state.profiles = fields.map((field) => buildColumnProfile(field, rows));
  state.fieldTypeDrafts = Object.fromEntries(
    state.profiles.map((profile) => [profile.field, profile.inferredTypeKey])
  );
  state.duplicateRows = countDuplicateRows(rows, fields);
  state.totalMissing = state.profiles.reduce((sum, profile) => sum + profile.missingCount, 0);
  state.numericStats = buildNumericStats(state.filteredRows);
  state.qualityNumericStats = buildNumericStats(state.rows);
  state.categoryStats = buildCategoryStats(state.filteredRows);

  const sourceDisplayName = sheetName ? `${sourceName} · ${sheetName}` : sourceName;
  dom.sourceName.textContent = `已载入 · ${sourceDisplayName}`;
  dom.sourceName.title = sourceDisplayName;
  dom.sheetSelectionPanel.classList.add("hidden");
  dom.results.classList.remove("hidden");
  if (state.parseWarnings.length) {
    const affectedRows = new Set(state.parseWarnings.map((warning) => warning.row).filter(Boolean)).size;
    setStatus(`CSV 已载入，但发现 ${state.parseWarnings.length} 个解析警告${affectedRows ? `，涉及 ${affectedRows} 行` : ""}。请核对数据预览和字段数量。`, "warning");
  } else {
    setStatus(`${sourceDisplayName} 已成功载入，分析结果已生成。`, "success");
  }

  renderOverview();
  renderInsights();
  renderPreview();
  renderSchema();
  initializeFilterControls();
  renderQuality();
  renderNumericStats();
  renderCategoryStats();
  populateControls();
  renderCharts();
  resetV2AnalysisState();
  renderTemplateMappingForm();
  resetTemplateResults("字段识别已完成。请选择分析模板并映射字段。");
  updateExportButtons();
  showResultsWorkspace();
}

function buildColumnProfile(field, rows) {
  const values = rows.map((row) => row[field]);
  const nonMissingValues = values.filter((value) => !isMissing(value));
  const total = rows.length;
  const nonMissingCount = nonMissingValues.length;
  const uniqueValues = new Set(nonMissingValues.map((value) => normalizeValue(value)));
  const uniqueCount = uniqueValues.size;
  const uniqueRatio = nonMissingCount ? uniqueCount / nonMissingCount : 0;
  const numericValues = nonMissingValues.map(toNumber).filter((value) => value !== null);
  const dateStrategy = buildDateStrategy(field, nonMissingValues);
  const dateValues = nonMissingValues.map((value) => toDate(value, dateStrategy)).filter(Boolean);
  const numericRatio = nonMissingCount ? numericValues.length / nonMissingCount : 0;
  const dateRatio = nonMissingCount ? dateValues.length / nonMissingCount : 0;
  const idName = isIdFieldName(field);

  let typeKey = "category";

  if (idName || (uniqueRatio >= 0.98 && nonMissingCount >= 12 && numericRatio < 0.7 && dateRatio < 0.7)) {
    typeKey = "id";
  } else if (dateRatio >= 0.8) {
    typeKey = "date";
  } else if (numericRatio >= 0.8) {
    typeKey = "numeric";
  }

  const type = FIELD_TYPE_LABELS[typeKey];
  const conversionFailuresByType = {
    numeric: nonMissingCount - numericValues.length,
    date: nonMissingCount - dateValues.length
  };
  const numericFormat = detectNumericFormat(nonMissingValues);
  const sampleValues = [];
  const sampleKeys = new Set();
  nonMissingValues.forEach((value) => {
    const displayed = displayValue(value);
    const key = normalizeValue(value);
    if (sampleValues.length >= 3 || sampleKeys.has(key)) return;
    sampleKeys.add(key);
    sampleValues.push(displayed);
  });

  return {
    field,
    type,
    typeKey,
    inferredType: type,
    inferredTypeKey: typeKey,
    total,
    nonMissingCount,
    missingCount: total - nonMissingCount,
    missingRate: total ? (total - nonMissingCount) / total : 0,
    uniqueCount,
    uniqueRatio,
    numericRatio,
    dateRatio,
    dateStrategy,
    numericFormat,
    sampleValues,
    conversionFailuresByType,
    conversionFailureCount: conversionFailuresByType[typeKey] || 0
  };
}

function countConversionFailuresForValues(values, typeKey) {
  if (typeKey !== "numeric" && typeKey !== "date") return 0;
  const convert = typeKey === "numeric" ? toNumber : toDate;
  return values.reduce((count, value) => (
    count + (!isMissing(value) && convert(value) === null ? 1 : 0)
  ), 0);
}

function countFieldConversionFailures(field, typeKey) {
  if (typeKey !== "numeric" && typeKey !== "date") return 0;
  const profile = getFieldProfile(field);
  if (profile?.conversionFailuresByType) {
    return profile.conversionFailuresByType[typeKey] || 0;
  }
  return countConversionFailuresForValues(
    state.rows.map((row) => row[field]).filter((value) => !isMissing(value)),
    typeKey
  );
}

function buildNumericStats(rows = getAnalysisRows()) {
  return getProfilesByType("numeric").flatMap((profile) => {
    const values = getNumericValues(profile.field, rows);
    if (!values.length) return [];
    const sorted = [...values].sort((a, b) => a - b);
    const avg = mean(values);
    const min = sorted[0];
    const max = sorted[sorted.length - 1];
    const q1 = quantile(sorted, 0.25);
    const q3 = quantile(sorted, 0.75);
    const iqr = q3 - q1;
    const lowerBound = q1 - 1.5 * iqr;
    const upperBound = q3 + 1.5 * iqr;
    const outliers = values.filter((value) => value < lowerBound || value > upperBound);

    return [{
      field: profile.field,
      count: values.length,
      mean: avg,
      median: median(values),
      min,
      max,
      std: standardDeviation(values, avg),
      q1,
      q3,
      lowerBound,
      upperBound,
      outlierCount: outliers.length
    }];
  });
}

function buildCategoryStats(rows = getAnalysisRows()) {
  return getProfilesByType("category").map((profile) => {
    const counts = countValues(profile.field, rows);
    const nonMissingCount = rows.filter((row) => !isMissing(row[profile.field])).length;
    const top = Array.from(counts.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10)
      .map(([name, count]) => ({
        name,
        count,
        ratio: nonMissingCount ? count / nonMissingCount : 0
      }));

    return {
      field: profile.field,
      total: nonMissingCount,
      top
    };
  });
}

function renderOverview() {
  const analysisRows = getAnalysisRows();
  const activeProfiles = getActiveProfiles();
  const numericCount = getProfilesByType("numeric").length;
  const categoryCount = getProfilesByType("category").length;
  const activeMissing = analysisRows.reduce((total, row) => (
    total + activeProfiles.filter((profile) => isMissing(row[profile.field])).length
  ), 0);
  const rowValue = hasActiveFilters()
    ? `${formatInteger(state.rows.length)} → ${formatInteger(analysisRows.length)}`
    : state.rows.length;
  const cards = [
    [hasActiveFilters() ? "总记录 → 当前分析" : "总记录数", rowValue],
    ["字段数", state.fields.length],
    ["数值字段数", numericCount],
    ["分类字段数", categoryCount],
    ["缺失值概览", activeMissing]
  ];

  dom.overviewCards.innerHTML = cards.map(([label, value]) => `
    <div class="metric-card">
      <div class="metric-label">${escapeHtml(label)}</div>
      <div class="metric-value">${typeof value === "number" ? formatInteger(value) : escapeHtml(value)}</div>
    </div>
  `).join("");
}

function renderInsights() {
  const analysisRows = getAnalysisRows();
  const insights = [];
  const activeProfiles = getActiveProfiles();
  const activeFields = activeProfiles.map((profile) => profile.field);
  const ignoredCount = state.profiles.length - activeProfiles.length;
  const activeMissing = analysisRows.reduce((sum, row) => (
    sum + activeFields.filter((field) => isMissing(row[field])).length
  ), 0);
  const analysisDuplicates = activeFields.length ? countDuplicateRows(analysisRows, activeFields) : 0;
  const missingRate = analysisRows.length && activeFields.length
    ? activeMissing / (analysisRows.length * activeFields.length)
    : 0;
  const filteredOutCount = Math.max(0, state.rows.length - analysisRows.length);
  const scopeText = hasActiveFilters()
    ? `原始数据包含 ${formatInteger(state.rows.length)} 行；应用筛选后保留 ${formatInteger(analysisRows.length)} 行，筛除 ${formatInteger(filteredOutCount)} 行。`
    : `数据集包含 ${formatInteger(state.rows.length)} 行，当前分析使用全部记录。`;

  if (!activeFields.length) {
    insights.push(scopeText);
    insights.push(`数据集包含 ${formatInteger(state.fields.length)} 个原始字段。`);
    insights.push("所有字段均已设为忽略，因此未生成统计、图表或字段相关的自动结论。");
    state.insights = insights;
    renderInsightGroups(insights);
    return;
  }

  if (!analysisRows.length) {
    insights.push(scopeText);
    insights.push("当前筛选条件下没有记录，因此未生成描述性统计、分类分布或趋势结论。请调整或清除筛选条件。");
    state.insights = insights;
    renderInsightGroups(insights);
    return;
  }

  const topNumeric = state.numericStats
    .slice()
    .sort((a, b) => Math.abs(b.max - b.min) - Math.abs(a.max - a.min))[0];
  const firstCategory = state.categoryStats.find((item) => item.top.length);
  const dateProfile = getProfilesByType("date")
    .find((profile) => fieldHasDateData(profile.field));

  insights.push(scopeText);
  insights.push(`当前有 ${formatInteger(activeFields.length)} 个字段参与分析${ignoredCount ? `，另有 ${formatInteger(ignoredCount)} 个字段已忽略` : ""}；其中数值字段 ${getProfilesByType("numeric").length} 个，分类字段 ${getProfilesByType("category").length} 个，日期字段 ${getProfilesByType("date").length} 个。`);
  insights.push(`参与分析的字段共发现 ${formatInteger(activeMissing)} 个缺失值，缺失率为 ${formatPercent(missingRate)}；按参与分析字段计算的重复行数量为 ${formatInteger(analysisDuplicates)}。`);

  if (topNumeric) {
    insights.push(`数值字段「${topNumeric.field}」的平均值为 ${formatFieldNumber(topNumeric.field, topNumeric.mean)}，最低值为 ${formatFieldNumber(topNumeric.field, topNumeric.min)}，最高值为 ${formatFieldNumber(topNumeric.field, topNumeric.max)}。`);
  } else {
    insights.push("未识别到可用于统计的数值字段。");
  }

  if (firstCategory) {
    const top = firstCategory.top[0];
    insights.push(`分类字段「${firstCategory.field}」中占比最高的类别是「${top.name}」，出现 ${formatInteger(top.count)} 次，占比 ${formatPercent(top.ratio)}。`);
  } else {
    insights.push("未识别到适合生成频数统计的分类字段。");
  }

  if (dateProfile) {
    insights.push(buildTrendInsight(dateProfile.field));
  } else {
    insights.push("未识别到日期字段，因此未生成时间趋势结论。");
  }

  state.insights = insights;
  renderInsightGroups(insights);
}

function renderInsightGroups(insights) {
  const findings = insights.slice(1);
  const keyFindings = findings.slice(0, 3);
  const otherFindings = findings.slice(3);
  const scopeLabel = hasActiveFilters()
    ? `基于当前筛选数据：${formatInteger(getAnalysisRows().length)} / ${formatInteger(state.rows.length)} 行`
    : `基于完整数据：${formatInteger(state.rows.length)} 行`;
  const renderItems = (items) => items.map((text) => `
    <div class="insight-item">${escapeHtml(text)}</div>
  `).join("");

  dom.insights.innerHTML = `
    <p class="insight-scope">${escapeHtml(scopeLabel)}</p>
    <h3 class="insight-group-title">关键发现</h3>
    <div class="insights">${renderItems(keyFindings)}</div>
    ${otherFindings.length ? `
      <details class="insights-more">
        <summary>其他发现（${formatInteger(otherFindings.length)}）</summary>
        <div class="insights">${renderItems(otherFindings)}</div>
      </details>
    ` : ""}
  `;
}

function renderPreview() {
  const previewRows = getAnalysisRows().slice(0, 10);
  if (!previewRows.length) {
    dom.previewTable.innerHTML = emptyTable("当前筛选条件下没有可预览的数据。");
    return;
  }
  dom.previewTable.innerHTML = `
    <thead>
      <tr>${state.fields.map((field) => `<th>${escapeHtml(field)}</th>`).join("")}</tr>
    </thead>
    <tbody>
      ${previewRows.map((row) => `
        <tr>${state.fields.map((field) => `<td>${escapeHtml(displayValue(row[field]))}</td>`).join("")}</tr>
      `).join("")}
    </tbody>
  `;
}

function renderSchema() {
  dom.schemaTable.innerHTML = `
    <thead>
      <tr>
        <th>字段名</th>
        <th>确认类型</th>
        <th>示例值</th>
        <th>字段详情</th>
      </tr>
    </thead>
    <tbody>
      ${state.profiles.map((profile) => {
        const draftTypeKey = getDraftTypeKey(profile);
        const failureCount = countFieldConversionFailures(profile.field, draftTypeKey);
        const requiresConversion = draftTypeKey === "numeric" || draftTypeKey === "date";
        const exampleText = profile.sampleValues.join("、") || "—";
        const failureLabel = requiresConversion
          ? `${profile.field} 转换失败 ${formatInteger(failureCount)} 个`
          : `${profile.field} 当前类型无需转换`;
        return `
          <tr>
            <td>${escapeHtml(profile.field)}</td>
            <td>
              <select class="schema-type-select" data-field-type="${escapeHtml(profile.field)}" aria-label="调整 ${escapeHtml(profile.field)} 的字段类型">
                ${Object.entries(FIELD_TYPE_LABELS).map(([typeKey, label]) => `
                  <option value="${typeKey}"${draftTypeKey === typeKey ? " selected" : ""}>${escapeHtml(label)}</option>
                `).join("")}
              </select>
            </td>
            <td><span class="schema-example" title="${escapeHtml(exampleText)}">${escapeHtml(exampleText)}</span></td>
            <td>
              <details class="schema-details">
                <summary>查看详情</summary>
                <dl class="schema-detail-list">
                  <dt>自动类型</dt><dd><span class="type-badge ${profile.inferredTypeKey}">${escapeHtml(profile.inferredType)}</span></dd>
                  <dt>非空值</dt><dd>${formatInteger(profile.nonMissingCount)}</dd>
                  <dt>唯一值</dt><dd>${formatInteger(profile.uniqueCount)}</dd>
                  <dt>转换失败</dt>
                  <dd><span class="conversion-failure-count${failureCount ? " warning" : ""}" title="${requiresConversion ? "非空值中无法安全转换的数量" : "该类型不需要转换"}" aria-label="${escapeHtml(failureLabel)}">${requiresConversion ? formatInteger(failureCount) : "—"}</span></dd>
                </dl>
              </details>
            </td>
          </tr>
        `;
      }).join("")}
    </tbody>
  `;
  updateFieldConfigActions();
  if (!dom.fieldConfigMessage.textContent) {
    setFieldConfigMessage("自动识别已完成。可修改类型后统一应用。", "");
  }
}

function getDraftTypeKey(profile) {
  return state.fieldTypeDrafts[profile.field] || profile.typeKey;
}

function stageFieldTypeChange(field, typeKey) {
  if (!Object.prototype.hasOwnProperty.call(FIELD_TYPE_LABELS, typeKey)) return;
  const profile = state.profiles.find((item) => item.field === field);
  if (!profile) return;

  state.fieldTypeDrafts[field] = typeKey;
  updateSchemaDraftRow(field, typeKey);
  updateFieldConfigActions();
  if (!hasFieldConfigDraftChanges()) {
    setFieldConfigMessage("当前选择与已应用的字段配置一致。", "");
    return;
  }
  const totalFailures = state.profiles.reduce((sum, item) => (
    sum + countFieldConversionFailures(item.field, getDraftTypeKey(item))
  ), 0);
  const failureText = totalFailures
    ? `；当前配置预计有 ${formatInteger(totalFailures)} 个非空值无法安全转换，应用后会从对应统计中跳过`
    : "";
  setFieldConfigMessage(`字段配置有未应用的修改${failureText}。`, "warning");
}

function updateSchemaDraftRow(field, typeKey) {
  const select = Array.from(dom.schemaTable.querySelectorAll("select[data-field-type]"))
    .find((item) => item.dataset.fieldType === field);
  if (!select) return;
  select.value = typeKey;

  const failureElement = select.closest("tr")?.querySelector(".conversion-failure-count");
  if (!failureElement) return;
  const failureCount = countFieldConversionFailures(field, typeKey);
  const requiresConversion = typeKey === "numeric" || typeKey === "date";
  failureElement.textContent = requiresConversion ? formatInteger(failureCount) : "—";
  failureElement.classList.toggle("warning", failureCount > 0);
  failureElement.title = requiresConversion
    ? "非空值中无法安全转换的数量"
    : "该类型不需要转换";
  failureElement.setAttribute(
    "aria-label",
    requiresConversion
      ? `${field} 转换失败 ${formatInteger(failureCount)} 个`
      : `${field} 当前类型无需转换`
  );
}

function applyFieldConfiguration() {
  if (!state.rows.length) return;
  state.profiles.forEach((profile) => {
    const typeKey = getDraftTypeKey(profile);
    profile.typeKey = typeKey;
    profile.type = FIELD_TYPE_LABELS[typeKey];
    profile.conversionFailureCount = countFieldConversionFailures(profile.field, typeKey);
  });
  state.fieldTypeDrafts = Object.fromEntries(
    state.profiles.map((profile) => [profile.field, profile.typeKey])
  );

  const failureCount = state.profiles.reduce((sum, profile) => sum + profile.conversionFailureCount, 0);
  const message = failureCount
    ? `字段配置已应用；${formatInteger(failureCount)} 个非空值转换失败，已从对应统计和图表中跳过。`
    : "字段配置已应用，统计、图表和分析结论已重新计算。";
  rebuildAnalysisFromProfiles(message);
}

function restoreAutomaticFieldTypes() {
  if (!state.rows.length) return;
  state.profiles.forEach((profile) => {
    profile.typeKey = profile.inferredTypeKey;
    profile.type = profile.inferredType;
    profile.conversionFailureCount = countFieldConversionFailures(profile.field, profile.inferredTypeKey);
  });
  state.fieldTypeDrafts = Object.fromEntries(
    state.profiles.map((profile) => [profile.field, profile.inferredTypeKey])
  );
  rebuildAnalysisFromProfiles("已恢复自动识别，统计、图表和分析结论已重新计算。");
}

function hasFieldConfigDraftChanges() {
  return state.profiles.some((profile) => getDraftTypeKey(profile) !== profile.typeKey);
}

function hasAutomaticTypeDifferences() {
  return state.profiles.some((profile) => (
    profile.typeKey !== profile.inferredTypeKey
      || getDraftTypeKey(profile) !== profile.inferredTypeKey
  ));
}

function updateFieldConfigActions() {
  dom.applyFieldConfig.disabled = !hasFieldConfigDraftChanges();
  dom.restoreAutoFieldTypes.disabled = !hasAutomaticTypeDifferences();
  updateExportButtons();
}

function setFieldConfigMessage(message, tone = "") {
  dom.fieldConfigMessage.className = `field-config-message${tone ? ` ${tone}` : ""}`;
  dom.fieldConfigMessage.textContent = message;
}

function ensureFieldConfigurationApplied(actionLabel) {
  if (!hasFieldConfigDraftChanges()) return true;
  setFieldConfigMessage(`字段配置有未应用的修改。请先点击“应用字段配置”，再${actionLabel}。`, "warning");
  return false;
}

function createEmptyFilterConfig() {
  return {
    category: null,
    numeric: null,
    date: null
  };
}

function getAnalysisRows() {
  return state.filterSourceRows === state.rows && Array.isArray(state.filteredRows)
    ? state.filteredRows
    : state.rows;
}

function getQualityNumericStats() {
  return state.filterSourceRows === state.rows
    ? state.qualityNumericStats
    : buildNumericStats(state.rows);
}

function hasActiveFilters(filters = state.filters) {
  return Boolean(filters?.category || filters?.numeric || filters?.date);
}

function resetFilterState() {
  state.filteredRows = state.rows.slice();
  state.filterSourceRows = state.rows;
  state.filters = createEmptyFilterConfig();
  state.categoryFilterOptions = [];
  state.categoryFilterDraftValues = new Set();
}

function initializeFilterControls() {
  const categoryFields = getProfilesByType("category").map((profile) => profile.field);
  const numericFields = getProfilesByType("numeric")
    .filter((profile) => fieldHasNumericData(profile.field, state.rows))
    .map((profile) => profile.field);
  const dateFields = getProfilesByType("date")
    .filter((profile) => fieldHasDateData(profile.field, state.rows))
    .map((profile) => profile.field);

  setSelectOptions(dom.categoryFilterField, categoryFields, "不筛选分类字段", true);
  setSelectOptions(dom.numericFilterField, numericFields, "不筛选数值字段", true);
  setSelectOptions(dom.dateFilterField, dateFields, "不筛选日期字段", true);
  dom.categoryFilterField.disabled = categoryFields.length === 0;
  dom.numericFilterField.disabled = numericFields.length === 0;
  dom.dateFilterField.disabled = dateFields.length === 0;
  dom.categoryFilterField.value = "";
  dom.numericFilterField.value = "";
  dom.dateFilterField.value = "";
  updateCategoryFilterField();
  updateNumericFilterField();
  updateDateFilterField();
  renderFilterCounts();
  setFilterMessage("尚未应用筛选，当前分析使用全部记录。", "");
}

function getCategoryFilterOptions(field) {
  if (!field) return [];
  const counts = new Map();
  state.rows.forEach((row) => {
    const value = normalizeValue(row[field]);
    counts.set(value, (counts.get(value) || 0) + 1);
  });
  return Array.from(counts.entries())
    .map(([value, count]) => ({
      value,
      count,
      label: value || "（空值）"
    }))
    .sort((a, b) => a.label.localeCompare(b.label, "zh-CN", { numeric: true }));
}

function updateCategoryFilterField() {
  const field = dom.categoryFilterField.value;
  dom.categoryFilterSearch.value = "";
  state.categoryFilterOptions = getCategoryFilterOptions(field);
  state.categoryFilterDraftValues = new Set(
    state.categoryFilterOptions.map((option) => option.value)
  );
  dom.categoryFilterControls.classList.toggle("hidden", !field);
  dom.categoryFilterEmpty.classList.toggle("hidden", Boolean(field));
  dom.categoryFilterEmpty.textContent = dom.categoryFilterField.disabled
    ? "没有可用的分类字段。"
    : "请选择分类字段。";
  renderCategoryFilterValues();
}

function renderCategoryFilterValues() {
  const query = dom.categoryFilterSearch.value.trim().toLocaleLowerCase("zh-CN");
  const visibleOptions = state.categoryFilterOptions.filter((option) => (
    !query || option.label.toLocaleLowerCase("zh-CN").includes(query)
  ));
  const selectedCount = state.categoryFilterDraftValues.size;
  dom.categoryFilterSelectionCount.textContent = dom.categoryFilterField.value
    ? `已选 ${formatInteger(selectedCount)} / ${formatInteger(state.categoryFilterOptions.length)}`
    : "未选择字段";

  if (!visibleOptions.length) {
    dom.categoryFilterValues.innerHTML = `<p class="filter-value-empty">没有匹配的类别。</p>`;
    return;
  }

  dom.categoryFilterValues.innerHTML = visibleOptions.map((option) => {
    const index = state.categoryFilterOptions.indexOf(option);
    return `
      <label class="filter-value-option">
        <input type="checkbox" data-category-filter-index="${index}"${state.categoryFilterDraftValues.has(option.value) ? " checked" : ""}>
        <span class="value-label" title="${escapeHtml(option.label)}">${escapeHtml(option.label)}</span>
        <span class="value-count">${formatInteger(option.count)}</span>
      </label>
    `;
  }).join("");
}

function updateCategoryFilterDraftSelection(event) {
  const checkbox = event.target.closest("input[data-category-filter-index]");
  if (!checkbox) return;
  const option = state.categoryFilterOptions[Number(checkbox.dataset.categoryFilterIndex)];
  if (!option) return;
  if (checkbox.checked) state.categoryFilterDraftValues.add(option.value);
  else state.categoryFilterDraftValues.delete(option.value);
  dom.categoryFilterSelectionCount.textContent = `已选 ${formatInteger(state.categoryFilterDraftValues.size)} / ${formatInteger(state.categoryFilterOptions.length)}`;
}

function setAllCategoryFilterValues(selected) {
  state.categoryFilterDraftValues = selected
    ? new Set(state.categoryFilterOptions.map((option) => option.value))
    : new Set();
  renderCategoryFilterValues();
}

function updateNumericFilterField() {
  const field = dom.numericFilterField.value;
  const values = field ? getNumericValues(field, state.rows) : [];
  const enabled = values.length > 0;
  dom.numericFilterMin.disabled = !enabled;
  dom.numericFilterMax.disabled = !enabled;
  if (!enabled) {
    dom.numericFilterMin.value = "";
    dom.numericFilterMax.value = "";
    dom.numericFilterRange.textContent = dom.numericFilterField.disabled ? "没有可用字段" : "请选择字段";
    return;
  }
  const { min, max } = numericExtents(values);
  dom.numericFilterMin.value = String(min);
  dom.numericFilterMax.value = String(max);
  dom.numericFilterRange.textContent = `${formatFieldNumber(field, min)} 至 ${formatFieldNumber(field, max)}`;
}

function updateDateFilterField() {
  const field = dom.dateFilterField.value;
  const dates = field
    ? state.rows.map((row) => parseFieldDate(field, row[field])).filter(Boolean).sort((a, b) => a - b)
    : [];
  const enabled = dates.length > 0;
  dom.dateFilterStart.disabled = !enabled;
  dom.dateFilterEnd.disabled = !enabled;
  if (!enabled) {
    dom.dateFilterStart.value = "";
    dom.dateFilterEnd.value = "";
    dom.dateFilterRange.textContent = dom.dateFilterField.disabled ? "没有可用字段" : "请选择字段";
    return;
  }
  const start = dates[0].toISOString().slice(0, 10);
  const end = dates[dates.length - 1].toISOString().slice(0, 10);
  dom.dateFilterStart.value = start;
  dom.dateFilterEnd.value = end;
  dom.dateFilterRange.textContent = `${start} 至 ${end}`;
}

function readFilterConfig() {
  const filters = createEmptyFilterConfig();
  const categoryField = dom.categoryFilterField.value;
  if (categoryField) {
    filters.category = {
      field: categoryField,
      values: Array.from(state.categoryFilterDraftValues)
    };
  }

  const numericField = dom.numericFilterField.value;
  if (numericField) {
    if (dom.numericFilterMin.value === "" || dom.numericFilterMax.value === "") {
      return { error: "请填写数值筛选的最小值和最大值。", filters };
    }
    const min = Number(dom.numericFilterMin.value);
    const max = Number(dom.numericFilterMax.value);
    if (!Number.isFinite(min) || !Number.isFinite(max)) {
      return { error: "数值筛选范围必须是有效数字。", filters };
    }
    if (min > max) {
      return { error: "数值筛选的最小值不能大于最大值。", filters };
    }
    filters.numeric = { field: numericField, min, max };
  }

  const dateField = dom.dateFilterField.value;
  if (dateField) {
    const start = dom.dateFilterStart.value;
    const end = dom.dateFilterEnd.value;
    if (!start || !end) {
      return { error: "请填写日期筛选的开始日期和结束日期。", filters };
    }
    if (start > end) {
      return { error: "日期筛选的开始日期不能晚于结束日期。", filters };
    }
    filters.date = { field: dateField, start, end };
  }

  return { error: "", filters };
}

function filterRowsByConfig(rows, filters) {
  const categoryValues = filters.category ? new Set(filters.category.values) : null;
  const startTime = filters.date ? Date.parse(`${filters.date.start}T00:00:00Z`) : null;
  const endTime = filters.date ? Date.parse(`${filters.date.end}T23:59:59.999Z`) : null;

  return rows.filter((row) => {
    if (filters.category) {
      const value = normalizeValue(row[filters.category.field]);
      if (!categoryValues.has(value)) return false;
    }
    if (filters.numeric) {
      const value = toNumber(row[filters.numeric.field]);
      if (value === null || value < filters.numeric.min || value > filters.numeric.max) return false;
    }
    if (filters.date) {
      const value = parseFieldDate(filters.date.field, row[filters.date.field]);
      if (!value || value.getTime() < startTime || value.getTime() > endTime) return false;
    }
    return true;
  });
}

function applyFilters() {
  if (!state.rows.length) return;
  if (!ensureFieldConfigurationApplied("应用筛选")) {
    setFilterMessage("字段配置有未应用的修改，请先应用字段配置。", "warning");
    return;
  }
  const { error, filters } = readFilterConfig();
  if (error) {
    setFilterMessage(error, "warning");
    return;
  }

  state.filters = filters;
  state.filteredRows = filterRowsByConfig(state.rows, filters);
  state.filterSourceRows = state.rows;
  refreshFilteredAnalysis();

  const activeCount = [filters.category, filters.numeric, filters.date].filter(Boolean).length;
  const filteredOutCount = state.rows.length - state.filteredRows.length;
  if (!activeCount) {
    setFilterMessage("未选择筛选字段，当前分析继续使用全部记录。", "");
  } else if (!state.filteredRows.length) {
    setFilterMessage(`已应用 ${formatInteger(activeCount)} 个筛选条件，但没有记录符合条件。`, "warning");
  } else {
    setFilterMessage(`已应用 ${formatInteger(activeCount)} 个筛选条件，保留 ${formatInteger(state.filteredRows.length)} 行，筛除 ${formatInteger(filteredOutCount)} 行。`, "success");
  }
}

function clearAllFilters() {
  if (!state.rows.length) return;
  resetFilterState();
  initializeFilterControls();
  refreshFilteredAnalysis();
  setFilterMessage("已清除全部筛选，当前分析恢复使用全部记录。", "success");
}

function renderFilterCounts() {
  const currentCount = getAnalysisRows().length;
  dom.filterOriginalCount.textContent = formatInteger(state.rows.length);
  dom.filterCurrentCount.textContent = formatInteger(currentCount);
  dom.filterExcludedCount.textContent = formatInteger(Math.max(0, state.rows.length - currentCount));
  updateGlobalFilterStatus();
}

function updateGlobalFilterStatus() {
  if (!dom.globalFilterSummary) return;
  const activeCount = [state.filters.category, state.filters.numeric, state.filters.date].filter(Boolean).length;
  const currentCount = getAnalysisRows().length;
  const totalCount = state.rows.length;
  dom.globalFilterSummary.textContent = activeCount
    ? `已筛选 ${formatInteger(activeCount)} 项 · ${formatInteger(currentCount)} / ${formatInteger(totalCount)} 行`
    : `全部 ${formatInteger(totalCount)} 行`;
  dom.globalFilterStatus.classList.toggle("active", activeCount > 0);
  dom.globalClearFilters.classList.toggle("hidden", activeCount === 0);
}

function setFilterMessage(message, tone = "") {
  dom.filterMessage.className = `filter-message${tone ? ` ${tone}` : ""}`;
  dom.filterMessage.textContent = message;
}

function refreshFilteredAnalysis() {
  const rerunV2Analysis = Boolean(
    state.customAnalysis && !dom.v2AnalysisResults.classList.contains("hidden")
  );
  const rerunTemplateAnalysis = !dom.templateResults.classList.contains("hidden");
  setDataExportStatus("");
  state.analysisCompletedAt = null;
  updateExportButtons();
  state.numericStats = buildNumericStats(getAnalysisRows());
  state.categoryStats = buildCategoryStats(getAnalysisRows());
  renderOverview();
  renderFilterCounts();
  renderInsights();
  renderPreview();
  renderNumericStats();
  renderCategoryStats();
  renderCharts();
  if (rerunV2Analysis) renderV2Analysis();
  else resetV2AnalysisState();
  if (rerunTemplateAnalysis) runTemplateAnalysis();
  state.analysisCompletedAt = new Date();
  updateExportButtons();
}

function rebuildAnalysisFromProfiles(fieldConfigMessage = "字段配置已更新。") {
  state.analysisCompletedAt = null;
  updateExportButtons();
  resetFilterState();
  setDataExportStatus("");
  state.numericStats = buildNumericStats(getAnalysisRows());
  state.qualityNumericStats = buildNumericStats(state.rows);
  state.categoryStats = buildCategoryStats(getAnalysisRows());
  renderOverview();
  renderInsights();
  renderPreview();
  renderSchema();
  initializeFilterControls();
  renderQuality();
  renderNumericStats();
  renderCategoryStats();
  populateControls();
  renderCharts();
  resetV2AnalysisState();
  renderTemplateMappingForm();
  resetTemplateResults("字段类型已更新。请重新选择分析字段或模板映射。");
  setFieldConfigMessage(`${fieldConfigMessage} 已清除现有筛选条件。`, "success");
  state.analysisCompletedAt = new Date();
  updateFieldConfigActions();
}

function renderQuality() {
  const qualityNumericStats = getQualityNumericStats();
  const missingClass = state.totalMissing === 0 ? "success" : missingRateLevel();
  const duplicateClass = state.duplicateRows === 0 ? "success" : "warning";
  dom.qualitySummary.innerHTML = `
    <span class="quality-pill ${missingClass}">缺失值总数：${formatInteger(state.totalMissing)}</span>
    <span class="quality-pill ${duplicateClass}">重复行：${formatInteger(state.duplicateRows)}</span>
    <span class="quality-pill">异常值字段：${formatInteger(qualityNumericStats.filter((item) => item.outlierCount > 0).length)}</span>
  `;

  dom.qualityTable.innerHTML = `
    <thead>
      <tr>
        <th>字段</th>
        <th>类型</th>
        <th>缺失值数量</th>
        <th>缺失率</th>
      </tr>
    </thead>
    <tbody>
      ${state.profiles.map((profile) => `
        <tr>
          <td>${escapeHtml(profile.field)}</td>
          <td>${escapeHtml(profile.type)}</td>
          <td>${formatInteger(profile.missingCount)}</td>
          <td>${formatPercent(profile.missingRate)}</td>
        </tr>
      `).join("")}
    </tbody>
  `;

  if (!qualityNumericStats.length) {
    dom.outlierTable.innerHTML = emptyTable("未识别到数值字段。");
    return;
  }

  dom.outlierTable.innerHTML = `
    <thead>
      <tr>
        <th>字段</th>
        <th>异常值数量</th>
        <th>下界</th>
        <th>上界</th>
      </tr>
    </thead>
    <tbody>
      ${qualityNumericStats.map((stat) => `
        <tr>
          <td>${escapeHtml(stat.field)}</td>
          <td>${formatInteger(stat.outlierCount)}</td>
          <td>${formatFieldNumber(stat.field, stat.lowerBound)}</td>
          <td>${formatFieldNumber(stat.field, stat.upperBound)}</td>
        </tr>
      `).join("")}
    </tbody>
  `;
}

function renderNumericStats() {
  if (!state.numericStats.length) {
    dom.numericStatsTable.innerHTML = emptyTable("未识别到数值字段。");
    return;
  }

  dom.numericStatsTable.innerHTML = `
    <thead>
      <tr>
        <th>字段</th>
        <th>有效数量</th>
        <th>平均值</th>
        <th>中位数</th>
        <th>最大值</th>
        <th>最小值</th>
        <th>标准差</th>
      </tr>
    </thead>
    <tbody>
      ${state.numericStats.map((stat) => `
        <tr>
          <td>${escapeHtml(stat.field)}</td>
          <td>${formatInteger(stat.count)}</td>
          <td>${formatFieldNumber(stat.field, stat.mean)}</td>
          <td>${formatFieldNumber(stat.field, stat.median)}</td>
          <td>${formatFieldNumber(stat.field, stat.max)}</td>
          <td>${formatFieldNumber(stat.field, stat.min)}</td>
          <td>${formatFieldNumber(stat.field, stat.std)}</td>
        </tr>
      `).join("")}
    </tbody>
  `;
}

function renderCategoryStats() {
  if (!state.categoryStats.length) {
    dom.categoryStats.innerHTML = `<div class="empty-note">未识别到分类字段。</div>`;
    return;
  }

  dom.categoryStats.innerHTML = state.categoryStats.map((category) => `
    <div class="frequency-card">
      <h3>${escapeHtml(category.field)}</h3>
      ${category.top.length ? category.top.map((item) => `
        <div class="frequency-row">
          <span class="category-name" title="${escapeHtml(item.name)}">${escapeHtml(item.name)}</span>
          <strong>${formatInteger(item.count)}</strong>
          <span>${formatPercent(item.ratio)}</span>
        </div>
      `).join("") : `<div class="empty-note">当前筛选结果中没有非空类别值。</div>`}
    </div>
  `).join("");
}

function populateControls() {
  const numericProfiles = getProfilesByType("numeric")
    .filter((profile) => fieldHasNumericData(profile.field));
  const categoryProfiles = getProfilesByType("category");
  const dateProfiles = getProfilesByType("date")
    .filter((profile) => fieldHasDateData(profile.field));

  setSelectOptions(dom.numericChartField, numericProfiles.map((profile) => profile.field));
  setSelectOptions(dom.categoryChartField, categoryProfiles.map((profile) => profile.field));
  setSelectOptions(dom.dateChartField, dateProfiles.map((profile) => profile.field));
  setSelectOptions(dom.metricField, numericProfiles.map((profile) => profile.field));
  setSelectOptions(dom.groupField, categoryProfiles.map((profile) => profile.field));
  setSelectOptions(dom.v2MetricField, numericProfiles.map((profile) => profile.field), "请选择数值字段");
  setSelectOptions(dom.v2GroupField, categoryProfiles.map((profile) => profile.field), "请选择分类字段");
  setSelectOptions(dom.v2DateField, dateProfiles.map((profile) => profile.field), "不选择时间字段", true);
}

function renderCharts() {
  if (!window.Chart) {
    ["numericDistributionChart", "categoryTopChart", "dateTrendChart", "customAnalysisChart"].forEach((id) => {
      drawEmptyChart(document.getElementById(id), "图表组件加载失败，请检查网络连接");
    });
    return;
  }
  renderNumericDistributionChart();
  renderCategoryTopChart();
  renderDateTrendChart();
  renderCustomAnalysisChart();
}

function renderNumericDistributionChart() {
  const field = dom.numericChartField.value;
  const canvas = document.getElementById("numericDistributionChart");
  destroyChart("numericDistributionChart");
  dom.numericChartTitle.textContent = field ? `${field} 数值分布（记录频数）` : "数值字段分布（记录频数）";
  canvas.setAttribute("aria-label", dom.numericChartTitle.textContent);
  if (!field) return drawEmptyChart(canvas, "没有可用的数值字段");

  const values = getNumericValues(field);
  if (!values.length) return drawEmptyChart(canvas, "该字段没有可安全转换的有效数值");
  const bins = buildHistogram(values, 10, field);
  state.charts.numericDistributionChart = new Chart(canvas, {
    type: "bar",
    data: {
      labels: bins.map((bin) => bin.label),
      datasets: [{
        label: `${field} 分布（记录频数）`,
        data: bins.map((bin) => bin.count),
        backgroundColor: CHART_THEME.primary
      }]
    },
    options: chartOptions("记录频数")
  });
}

function renderCategoryTopChart() {
  const field = dom.categoryChartField.value;
  const canvas = document.getElementById("categoryTopChart");
  destroyChart("categoryTopChart");
  dom.categoryChartTitle.textContent = field ? `${field} Top 10（记录数）` : "分类字段 Top 10（记录数）";
  canvas.setAttribute("aria-label", dom.categoryChartTitle.textContent);
  if (!field) return drawEmptyChart(canvas, "没有可用的分类字段");

  const counts = Array.from(countValues(field).entries())
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10);
  if (!counts.length) return drawEmptyChart(canvas, "当前筛选结果中没有可统计的分类值");
  state.charts.categoryTopChart = new Chart(canvas, {
    type: "bar",
    data: {
      labels: counts.map(([name]) => name),
      datasets: [{
        label: `${field} Top 10（记录数）`,
        data: counts.map(([, count]) => count),
        backgroundColor: CHART_THEME.secondary
      }]
    },
    options: chartOptions("记录数")
  });
}

function renderDateTrendChart() {
  const field = dom.dateChartField.value;
  const canvas = document.getElementById("dateTrendChart");
  destroyChart("dateTrendChart");
  if (!field) return drawEmptyChart(canvas, "没有可用的日期字段");

  const numericField = dom.metricField.value || getProfilesByType("numeric")
    .find((profile) => fieldHasNumericData(profile.field))?.field;
  const trend = buildDateTrend(field, numericField);
  if (!trend.labels.length) return drawEmptyChart(canvas, "该字段没有可安全转换的有效日期");
  const timeUnit = getDateGranularity(field) === "year"
    ? "年份"
    : trend.labels.some((label) => /^\d{4}-\d{2}$/.test(label)) ? "月份" : "日期";
  const aggregationLabel = numericField ? "平均值" : "记录数";
  dom.dateChartTitle.textContent = numericField
    ? `${numericField} 按${timeUnit}趋势（${aggregationLabel}）`
    : `${field} 按${timeUnit}趋势（${aggregationLabel}）`;
  canvas.setAttribute("aria-label", dom.dateChartTitle.textContent);
  state.charts.dateTrendChart = new Chart(canvas, {
    type: "line",
    data: {
      labels: trend.labels,
      datasets: [{
        label: trend.label,
        data: trend.values,
        borderColor: CHART_THEME.primary,
        backgroundColor: CHART_THEME.primaryFill,
        tension: 0.25,
        fill: true
      }]
    },
    options: chartOptions(numericAxisLabel(numericField, trend.axisLabel), numericField)
  });
}

function renderCustomAnalysisChart() {
  const metric = dom.metricField.value;
  const group = dom.groupField.value;
  const method = dom.aggregateMethod.value;
  const canvas = document.getElementById("customAnalysisChart");
  destroyChart("customAnalysisChart");
  const methodLabel = { sum: "合计", avg: "平均值", median: "中位数" }[method];
  dom.customAnalysisChartTitle.textContent = metric && group
    ? `${metric} 按 ${group} 分组（${methodLabel}）`
    : "分组聚合结果";
  canvas.setAttribute("aria-label", dom.customAnalysisChartTitle.textContent);

  if (!metric || !group) {
    return drawEmptyChart(canvas, "请选择数值字段和分类字段");
  }

  const grouped = aggregateByCategory(metric, group, method).slice(0, 12);
  if (!grouped.length) return drawEmptyChart(canvas, "当前筛选结果中没有可计算的分组记录");
  state.charts.customAnalysisChart = new Chart(canvas, {
    type: "bar",
    data: {
      labels: grouped.map((item) => item.name),
      datasets: [{
        label: `${group} 分组的 ${metric} ${methodLabel}`,
        data: grouped.map((item) => item.value),
        backgroundColor: CHART_THEME.warning
      }]
    },
    options: chartOptions(numericAxisLabel(metric, methodLabel), metric)
  });
}

function resetV2AnalysisState() {
  destroyChart("v2TopGroupChart");
  destroyChart("v2SelectedTrendChart");
  state.customAnalysis = null;
  dom.v2AnalysisResults.classList.add("hidden");
  dom.v2FieldPrompt.className = "analysis-prompt";
  dom.v2FieldPrompt.textContent = "请选择指标字段、分组维度和聚合方式，再点击“生成自定义分析”。时间字段为可选。";
}

function renderV2Analysis() {
  const metricField = dom.v2MetricField.value;
  const groupField = dom.v2GroupField.value;
  const aggregateMethod = dom.v2AggregateMethod.value;
  const dateField = dom.v2DateField.value;
  const aggregateLabel = getAggregateMethodLabel(aggregateMethod);

  destroyChart("v2TopGroupChart");
  destroyChart("v2SelectedTrendChart");

  const validationMessage = getV2ValidationMessage(metricField, groupField, dateField);
  if (validationMessage) {
    state.customAnalysis = null;
    dom.v2AnalysisResults.classList.add("hidden");
    dom.v2FieldPrompt.className = "analysis-prompt warning";
    dom.v2FieldPrompt.textContent = validationMessage;
    return;
  }

  const grouped = buildV2GroupedStats(metricField, groupField);
  if (!grouped.length) {
    state.customAnalysis = null;
    dom.v2AnalysisResults.classList.add("hidden");
    dom.v2FieldPrompt.className = "analysis-prompt warning";
    dom.v2FieldPrompt.textContent = "所选字段没有可计算的有效记录，请重新选择字段后再分析。";
    return;
  }

  dom.v2FieldPrompt.className = "analysis-prompt";
  dom.v2FieldPrompt.textContent = dateField
    ? `已生成分组${aggregateLabel}汇总和时间趋势。`
    : `已生成分组${aggregateLabel}汇总；未选择时间字段，因此不生成时间趋势图。`;
  dom.v2AnalysisResults.classList.remove("hidden");

  state.customAnalysis = {
    metricField,
    groupField,
    aggregateMethod,
    dateField,
    grouped: grouped.map((item) => ({ ...item })),
    trend: dateField ? buildMetricDateTrend(dateField, metricField, aggregateMethod) : null
  };

  renderV2SummaryCards(metricField, groupField, aggregateMethod, dateField, grouped);
  renderV2AggregationTables(metricField, groupField, grouped);
  renderV2PrimaryTable(metricField, groupField, grouped, aggregateMethod);
  renderV2TopGroupChart(metricField, groupField, grouped, aggregateMethod);
  dom.v2ResultSummary.textContent = hasActiveFilters()
    ? `已按 ${groupField} 对 ${metricField} 进行${aggregateLabel}，结果基于当前筛选数据 ${formatInteger(getAnalysisRows().length)} / ${formatInteger(state.rows.length)} 行。`
    : `已按 ${groupField} 对 ${metricField} 进行${aggregateLabel}，结果基于完整数据 ${formatInteger(state.rows.length)} 行。`;

  if (dateField) {
    dom.v2TrendCard.classList.remove("hidden");
    renderV2TrendChart(metricField, dateField, aggregateMethod);
  } else {
    dom.v2TrendCard.classList.add("hidden");
  }
}

function getV2ValidationMessage(metricField, groupField, dateField) {
  if (!state.rows.length) return "请先上传或加载 CSV 数据。";
  if (!getAnalysisRows().length) return "当前筛选条件下没有可分析的记录，请调整或清除筛选条件。";
  if (!metricField) return "请选择一个数值字段作为分析指标。";
  if (!groupField) return "请选择一个分类字段作为分组维度。";
  if (!getProfilesByType("numeric").some((profile) => profile.field === metricField)) {
    return "分析指标必须是系统识别出的数值字段。";
  }
  if (!getProfilesByType("category").some((profile) => profile.field === groupField)) {
    return "分组维度必须是系统识别出的分类字段。";
  }
  if (dateField && !getProfilesByType("date").some((profile) => profile.field === dateField)) {
    return "时间字段必须是系统识别出的日期字段，或保持不选择。";
  }
  return "";
}

function buildV2GroupedStats(metricField, groupField) {
  return groupNumericFieldByCategory(metricField, groupField);
}

function renderV2SummaryCards(metricField, groupField, aggregateMethod, dateField, grouped) {
  const validRecordCount = grouped.reduce((total, item) => total + item.count, 0);
  const trendPointCount = dateField ? buildMetricDateTrend(dateField, metricField, aggregateMethod).labels.length : 0;
  const cards = [
    ["分析指标", metricField],
    ["分组维度", groupField],
    ["聚合方式", getAggregateMethodLabel(aggregateMethod)],
    [dateField ? "时间点数量" : "有效记录数", dateField ? formatInteger(trendPointCount) : formatInteger(validRecordCount)]
  ];

  dom.v2SummaryCards.innerHTML = cards.map(([label, value]) => `
    <div class="analysis-summary-card">
      <div class="analysis-summary-label">${escapeHtml(label)}</div>
      <div class="analysis-summary-value" title="${escapeHtml(value)}">${escapeHtml(value)}</div>
    </div>
  `).join("");
}

function renderV2AggregationTables(metricField, groupField, grouped) {
  const bySum = [...grouped].sort((a, b) => compareNumbersDescending(a.sum, b.sum));
  const byAvg = [...grouped].sort((a, b) => compareNumbersDescending(a.avg, b.avg));
  dom.v2SumTable.innerHTML = buildV2AggregationTableHtml(groupField, metricField, bySum, "sum");
  dom.v2AvgTable.innerHTML = buildV2AggregationTableHtml(groupField, metricField, byAvg, "avg");
}

function renderV2PrimaryTable(metricField, groupField, grouped, aggregateMethod) {
  const sorted = [...grouped].sort((a, b) => (
    compareNumbersDescending(a[aggregateMethod], b[aggregateMethod])
  ));
  const aggregateLabel = getAggregateMethodLabel(aggregateMethod);
  dom.v2PrimaryTableTitle.textContent = `${metricField} 按 ${groupField} 分组（${aggregateLabel}）`;
  dom.v2PrimaryTable.innerHTML = buildV2AggregationTableHtml(groupField, metricField, sorted, aggregateMethod);
}

function buildV2AggregationTableHtml(groupField, metricField, rows, valueKey) {
  const valueLabel = { sum: "求和值", avg: "平均值", median: "中位数" }[valueKey] || "汇总值";
  return `
    <thead>
      <tr>
        <th>${escapeHtml(groupField)}</th>
        <th>记录数</th>
        <th>${escapeHtml(metricField)} ${valueLabel}</th>
      </tr>
    </thead>
    <tbody>
      ${rows.map((row) => `
        <tr>
          <td>${escapeHtml(row.name)}</td>
          <td>${formatInteger(row.count)}</td>
          <td>${formatFieldNumber(metricField, row[valueKey])}</td>
        </tr>
      `).join("")}
    </tbody>
  `;
}

function renderV2TopGroupChart(metricField, groupField, grouped, aggregateMethod = "sum") {
  const canvas = document.getElementById("v2TopGroupChart");
  if (!window.Chart) return drawEmptyChart(canvas, "图表组件加载失败，请检查网络连接");
  const aggregateLabel = getAggregateMethodLabel(aggregateMethod);
  const topGroups = [...grouped]
    .sort((a, b) => compareNumbersDescending(a[aggregateMethod], b[aggregateMethod]))
    .slice(0, 10);
  dom.v2TopGroupTitle.textContent = `${metricField} 按 ${groupField} 分组 Top 10（${aggregateLabel}）`;
  canvas.setAttribute("aria-label", dom.v2TopGroupTitle.textContent);
  state.charts.v2TopGroupChart = new Chart(canvas, {
    type: "bar",
    data: {
      labels: topGroups.map((item) => item.name),
      datasets: [{
        label: `${metricField} 按 ${groupField} 分组（${aggregateLabel}）`,
        data: topGroups.map((item) => item[aggregateMethod]),
        backgroundColor: CHART_THEME.primary
      }]
    },
    options: chartOptions(numericAxisLabel(metricField, aggregateLabel), metricField)
  });
}

function renderV2TrendChart(metricField, dateField, aggregateMethod = "sum") {
  const canvas = document.getElementById("v2SelectedTrendChart");
  if (!window.Chart) return drawEmptyChart(canvas, "图表组件加载失败，请检查网络连接");
  const trend = buildMetricDateTrend(dateField, metricField, aggregateMethod);
  if (!trend.labels.length) {
    return drawEmptyChart(canvas, "所选时间字段没有可用的趋势数据");
  }

  const timeUnit = getDateGranularity(dateField) === "year"
    ? "年份"
    : trend.labels.some((label) => /^\d{4}-\d{2}$/.test(label)) ? "月份" : "日期";
  const aggregateLabel = getAggregateMethodLabel(aggregateMethod);
  dom.v2TrendTitle.textContent = `${metricField} 按${timeUnit}趋势（${aggregateLabel}）`;
  canvas.setAttribute("aria-label", dom.v2TrendTitle.textContent);
  state.charts.v2SelectedTrendChart = new Chart(canvas, {
    type: "line",
    data: {
      labels: trend.labels,
      datasets: [{
        label: dom.v2TrendTitle.textContent,
        data: trend.values,
        borderColor: CHART_THEME.secondary,
        backgroundColor: CHART_THEME.secondaryFill,
        tension: 0.25,
        fill: true
      }]
    },
    options: chartOptions(numericAxisLabel(metricField, aggregateLabel), metricField)
  });
}

function buildMetricDateTrend(dateField, metricField, aggregateMethod = "sum") {
  return buildTrendFromRows(dateField, (row) => toNumber(row[metricField]), aggregateMethod);
}

function getAggregateMethodLabel(method) {
  return { sum: "求和", avg: "平均值", median: "中位数" }[method] || "求和";
}

function getCurrentTemplate() {
  return TEMPLATE_DEFINITIONS[dom.scenarioTemplate.value] || TEMPLATE_DEFINITIONS.generic;
}

function renderTemplateMappingForm() {
  const template = getCurrentTemplate();
  dom.templateMappingForm.innerHTML = template.fields.map((field) => `
    <label>
      ${escapeHtml(field.label)}${field.required ? " *" : ""}
      <select id="templateField_${escapeHtml(field.key)}" data-template-field="${escapeHtml(field.key)}">
        ${buildTemplateFieldOptions(field)}
      </select>
    </label>
  `).join("");
}

function buildTemplateFieldOptions(fieldConfig) {
  const placeholder = fieldConfig.required ? "请选择字段" : "不选择";
  const options = getActiveProfiles()
    .filter((profile) => fieldConfig.expected === "any" || profile.typeKey === fieldConfig.expected)
    .filter((profile) => fieldConfig.expected !== "numeric" || fieldHasNumericData(profile.field))
    .filter((profile) => fieldConfig.expected !== "date" || fieldHasDateData(profile.field))
    .map((profile) => {
      return `<option value="${escapeHtml(profile.field)}">${escapeHtml(profile.field)}（${escapeHtml(profile.type)}）</option>`;
    }).join("");
  return `<option value="">${placeholder}</option>${options}`;
}

function resetTemplateResults(message) {
  clearTemplateCharts();
  dom.templateResults.classList.add("hidden");
  dom.templatePrompt.className = "analysis-prompt";
  dom.templatePrompt.textContent = message;
}

function clearTemplateCharts() {
  destroyChart("templateMainChart");
  destroyChart("templateSecondaryChart");
  destroyChart("templateTrendChart");
}

function getTemplateMappings(template) {
  return template.fields.reduce((mappings, field) => {
    const select = document.getElementById(`templateField_${field.key}`);
    mappings[field.key] = select ? select.value : "";
    return mappings;
  }, {});
}

function runTemplateAnalysis() {
  const templateKey = dom.scenarioTemplate.value;
  const template = getCurrentTemplate();
  const mappings = getTemplateMappings(template);
  const validationMessage = validateTemplateMappings(template, mappings);

  clearTemplateCharts();
  if (validationMessage) {
    dom.templateResults.classList.add("hidden");
    dom.templatePrompt.className = "analysis-prompt warning";
    dom.templatePrompt.textContent = validationMessage;
    return;
  }

  const result = {
    generic: analyzeGenericTemplate,
    sales: analyzeSalesTemplate,
    score: analyzeScoreTemplate,
    usedGoods: analyzeUsedGoodsTemplate,
    survey: analyzeSurveyTemplate,
    behavior: analyzeBehaviorTemplate
  }[templateKey](mappings);

  renderTemplateResult(result);
}

function validateTemplateMappings(template, mappings) {
  if (!getAnalysisRows().length) {
    return "当前筛选条件下没有可分析的记录，请调整或清除筛选条件。";
  }
  for (const field of template.fields) {
    const selected = mappings[field.key];
    if (field.required && !selected) {
      return `请为「${field.label.replace("（可选）", "")}」选择字段。`;
    }
    if (!selected) continue;
    const profile = getFieldProfile(selected);
    if (!profile || profile.typeKey === "ignore") {
      return `字段「${selected}」已被忽略或不再可用，请重新选择字段。`;
    }
    if (field.expected !== "any" && profile.typeKey !== field.expected) {
      return `字段「${selected}」当前配置为「${profile.type}」，与「${field.label}」要求的类型不匹配。`;
    }
    if (field.expected === "numeric" && !fieldHasNumericData(selected)) {
      return `字段「${selected}」无法解析出有效数值，请为「${field.label}」选择数值字段。`;
    }
    if (field.expected === "date" && !fieldHasDateData(selected)) {
      return `字段「${selected}」无法解析出有效日期，请为「${field.label}」选择日期字段或保持不选择。`;
    }
  }

  if (dom.scenarioTemplate.value === "survey") {
    const selectedCount = Object.values(mappings).filter(Boolean).length;
    if (selectedCount === 0) {
      return "问卷调查分析至少需要选择一个字段。";
    }
  }

  return "";
}

function renderTemplateResult(result) {
  dom.templatePrompt.className = "analysis-prompt";
  dom.templatePrompt.textContent = result.prompt;
  dom.templateResults.classList.remove("hidden");
  dom.templateSummaryCards.innerHTML = result.cards.map(([label, value]) => `
    <div class="analysis-summary-card">
      <div class="analysis-summary-label">${escapeHtml(label)}</div>
      <div class="analysis-summary-value" title="${escapeHtml(value)}">${escapeHtml(value)}</div>
    </div>
  `).join("");
  dom.templateInsights.innerHTML = result.insights.map((text) => `
    <div class="insight-item">${escapeHtml(text)}</div>
  `).join("");
  dom.templatePrimaryTable.innerHTML = buildSimpleTableHtml(result.table.headers, result.table.rows);

  renderTemplateChart("templateMainChart", dom.templateMainChartTitle, result.mainChart);
  renderTemplateChart("templateSecondaryChart", dom.templateSecondaryChartTitle, result.secondaryChart);
  if (result.trendChart) {
    dom.templateTrendCard.classList.remove("hidden");
    renderTemplateChart("templateTrendChart", dom.templateTrendChartTitle, result.trendChart);
  } else {
    dom.templateTrendCard.classList.add("hidden");
  }
}

function renderTemplateChart(chartId, titleElement, chartConfig) {
  const canvas = document.getElementById(chartId);
  titleElement.textContent = chartConfig.title;
  if (!window.Chart) {
    drawEmptyChart(canvas, "图表组件加载失败，请检查网络连接");
    return;
  }
  if (!chartConfig.labels.length) {
    drawEmptyChart(canvas, "没有足够数据生成该图表");
    return;
  }

  state.charts[chartId] = new Chart(canvas, {
    type: chartConfig.type || "bar",
    data: {
      labels: chartConfig.labels,
      datasets: [{
        label: chartConfig.label,
        data: chartConfig.values,
        borderColor: chartConfig.color || CHART_THEME.primary,
        backgroundColor: chartConfig.background || chartConfig.color || CHART_THEME.primary,
        tension: 0.25,
        fill: chartConfig.type === "line"
      }]
    },
    options: chartOptions(
      numericAxisLabel(chartConfig.valueField, chartConfig.axisLabel || "值"),
      chartConfig.valueField
    )
  });
}

function analyzeGenericTemplate(mappings) {
  const grouped = groupNumericFieldByCategory(mappings.metric, mappings.group);
  const values = getNumericValues(mappings.metric);
  const stats = summarizeNumbers(values);
  const topBySum = grouped.find((item) => Number.isFinite(item.sum));
  const topByAvg = [...grouped]
    .sort((a, b) => compareNumbersDescending(a.avg, b.avg))[0];
  const trend = mappings.date ? buildTrendFromRows(mappings.date, (row) => toNumber(row[mappings.metric]), "sum") : null;
  const insights = [
    `本次通用分析使用「${mappings.metric}」作为数值指标、「${mappings.group}」作为分组维度，有效数值记录为 ${formatInteger(values.length)} 条。`,
    `「${mappings.metric}」的平均值为 ${formatFieldNumber(mappings.metric, stats.mean)}，中位数为 ${formatFieldNumber(mappings.metric, stats.median)}，最小值为 ${formatFieldNumber(mappings.metric, stats.min)}，最大值为 ${formatFieldNumber(mappings.metric, stats.max)}。`,
    topBySum ? `按「${mappings.group}」分组求和最高的是「${topBySum.name}」，求和值为 ${formatFieldNumber(mappings.metric, topBySum.sum)}。` : "所选字段没有形成可比较的分组结果。",
    topByAvg ? `按「${mappings.group}」分组平均值最高的是「${topByAvg.name}」，平均值为 ${formatFieldNumber(mappings.metric, topByAvg.avg)}。该结果只描述数据差异，不说明原因。` : "所选字段没有形成平均值比较结果。"
  ];
  if (trend) insights.push(buildTrendFactText(mappings.metric, mappings.date, trend, "求和值"));

  return {
    prompt: "已生成通用字段映射分析。",
    cards: [
      ["模板", "通用数据分析"],
      ["指标字段", mappings.metric],
      ["分组字段", mappings.group],
      ["有效记录", formatInteger(values.length)]
    ],
    insights,
    table: numericGroupTable(mappings.group, mappings.metric, grouped),
    mainChart: barChartConfig(`按 ${mappings.group} 的 ${mappings.metric} 求和 Top 10`, grouped.slice(0, 10).map((item) => item.name), grouped.slice(0, 10).map((item) => item.sum), `${mappings.metric} 求和`, "求和值", CHART_THEME.primary, mappings.metric),
    secondaryChart: histogramChartConfig(`${mappings.metric} 分布`, values, "记录数", mappings.metric),
    trendChart: trend ? lineChartConfig(`${mappings.metric} 时间趋势`, trend.labels, trend.values, `${mappings.metric} 求和`, "求和值", mappings.metric) : null
  };
}

function analyzeSalesTemplate(mappings) {
  const metricRows = buildSalesMetricRows(mappings);
  const categoryRows = metricRows
    ? groupMetricRowsByField(metricRows, mappings.category)
    : frequencyRows(mappings.category);
  const regionRows = metricRows
    ? groupMetricRowsByField(metricRows, mappings.region)
    : frequencyRows(mappings.region);
  const metricLabel = metricRows ? metricRows.label : "记录数";
  const metricFormatField = metricRows?.formatField || "";
  const categoryTop = metricRows
    ? categoryRows.find((item) => Number.isFinite(item.sum))
    : categoryRows[0];
  const regionTop = metricRows
    ? regionRows.find((item) => Number.isFinite(item.sum))
    : regionRows[0];
  const trend = mappings.date
    ? buildTrendFromRows(mappings.date, metricRows ? (row) => getSalesMetricValue(row, mappings) : null, metricRows ? "sum" : "count")
    : null;

  const insights = [
    `销售模板使用「${mappings.category}」和「${mappings.region}」进行分组，共分析 ${formatInteger(getAnalysisRows().length)} 行记录。`,
    metricRows
      ? `销售数值基于「${metricLabel}」计算，有效数值记录为 ${formatInteger(metricRows.length)} 条。`
      : "未选择可计算的销售数值字段，因此仅按记录数做分组统计。",
    categoryTop ? `按品类字段「${mappings.category}」比较，${metricRows ? "求和" : "记录数"}最高的分组是「${categoryTop.name}」，值为 ${metricRows ? formatFieldNumber(metricFormatField, categoryTop.sum) : formatNumber(categoryTop.count)}。` : "品类字段没有可统计的有效分组。",
    regionTop ? `按地区字段「${mappings.region}」比较，${metricRows ? "求和" : "记录数"}最高的分组是「${regionTop.name}」，值为 ${metricRows ? formatFieldNumber(metricFormatField, regionTop.sum) : formatNumber(regionTop.count)}。该结果只反映 CSV 中的数据事实。` : "地区字段没有可统计的有效分组。"
  ];
  if (trend) insights.push(buildTrendFactText(metricLabel, mappings.date, trend, metricRows ? "求和值" : "记录数"));

  return {
    prompt: "已生成销售数据模板分析。",
    cards: [
      ["模板", "销售数据分析"],
      ["销售指标", metricLabel],
      ["品类字段", mappings.category],
      ["地区字段", mappings.region]
    ],
    insights,
    table: metricRows
      ? numericGroupTable(mappings.category, metricLabel, categoryRows, metricFormatField)
      : frequencyTable(mappings.category, categoryRows),
    mainChart: barChartConfig(`按 ${mappings.category} 的 ${metricLabel} Top 10`, topLabels(categoryRows), topValues(categoryRows, metricRows ? "sum" : "count"), metricRows ? `${metricLabel} 求和` : "记录数", metricRows ? "求和值" : "记录数", CHART_THEME.primary, metricFormatField),
    secondaryChart: barChartConfig(`按 ${mappings.region} 的 ${metricLabel} Top 10`, topLabels(regionRows), topValues(regionRows, metricRows ? "sum" : "count"), metricRows ? `${metricLabel} 求和` : "记录数", metricRows ? "求和值" : "记录数", CHART_THEME.secondary, metricFormatField),
    trendChart: trend ? lineChartConfig(`${metricLabel} 时间趋势`, trend.labels, trend.values, metricRows ? `${metricLabel} 求和` : "记录数", metricRows ? "求和值" : "记录数", metricFormatField) : null
  };
}

function analyzeScoreTemplate(mappings) {
  const values = getNumericValues(mappings.score);
  const stats = summarizeNumbers(values);
  const dimensionResults = [...new Set([mappings.class, mappings.subject].filter(Boolean))]
    .map((field) => ({ field, rows: groupNumericFieldByCategory(mappings.score, field) }));
  const primaryDimension = dimensionResults[0];
  const secondaryDimension = dimensionResults[1];
  const trend = mappings.examDate ? buildTrendFromRows(mappings.examDate, (row) => toNumber(row[mappings.score]), "avg") : null;
  const insights = [
    `成绩模板使用「${mappings.score}」作为成绩字段，有效成绩记录为 ${formatInteger(values.length)} 条。`,
    `成绩平均值为 ${formatFieldNumber(mappings.score, stats.mean)}，中位数为 ${formatFieldNumber(mappings.score, stats.median)}，最低值为 ${formatFieldNumber(mappings.score, stats.min)}，最高值为 ${formatFieldNumber(mappings.score, stats.max)}。`
  ];
  if (!dimensionResults.length) {
    insights.push("未选择班级或科目字段，因此未生成分组平均值比较。");
  } else {
    dimensionResults.forEach(({ field, rows }) => {
      const sorted = [...rows].sort((a, b) => b.avg - a.avg);
      const topAvg = sorted[0];
      const lowAvg = sorted[sorted.length - 1];
      insights.push(topAvg && lowAvg
        ? `按「${field}」分组时，平均值最高的是「${topAvg.name}」(${formatFieldNumber(mappings.score, topAvg.avg)})，平均值最低的是「${lowAvg.name}」(${formatFieldNumber(mappings.score, lowAvg.avg)})。这里只描述字段结果，不评价教学或个体原因。`
        : `字段「${field}」没有形成可比较的成绩分组。`);
    });
  }
  if (trend) insights.push(buildTrendFactText(mappings.score, mappings.examDate, trend, "平均值"));

  return {
    prompt: "已生成成绩数据模板分析。",
    cards: [
      ["模板", "成绩数据分析"],
      ["成绩字段", mappings.score],
      ["有效成绩", formatInteger(values.length)],
      ["平均值", formatFieldNumber(mappings.score, stats.mean)]
    ],
    insights,
    table: dimensionResults.length ? multiDimensionNumericTable(mappings.score, dimensionResults) : simpleStatsTable(mappings.score, stats, values.length),
    mainChart: primaryDimension
      ? barChartConfig(`按 ${primaryDimension.field} 的成绩平均值 Top 10`, topLabels(primaryDimension.rows, "avg"), topValues(primaryDimension.rows, "avg"), "成绩平均值", "平均值", CHART_THEME.primary, mappings.score)
      : histogramChartConfig(`${mappings.score} 分布`, values, "记录数", mappings.score),
    secondaryChart: secondaryDimension
      ? barChartConfig(`按 ${secondaryDimension.field} 的成绩平均值 Top 10`, topLabels(secondaryDimension.rows, "avg"), topValues(secondaryDimension.rows, "avg"), "成绩平均值", "平均值", CHART_THEME.secondary, mappings.score)
      : histogramChartConfig(`${mappings.score} 分布`, values, "记录数", mappings.score),
    trendChart: trend ? lineChartConfig(`${mappings.score} 时间趋势`, trend.labels, trend.values, "成绩平均值", "平均值", mappings.score) : null
  };
}

function analyzeUsedGoodsTemplate(mappings) {
  const values = getNumericValues(mappings.price);
  const stats = summarizeNumbers(values);
  const dimensionResults = [...new Set(["item", "platform", "city", "condition"].map((key) => mappings[key]).filter(Boolean))]
    .map((field) => ({ field, rows: groupNumericFieldByCategory(mappings.price, field) }));
  const primaryDimension = dimensionResults[0];
  const secondaryDimension = dimensionResults[1];
  const trend = mappings.date ? buildTrendFromRows(mappings.date, (row) => toNumber(row[mappings.price]), "avg") : null;
  const insights = [
    `二手商品价格模板使用「${mappings.price}」作为价格字段，有效价格记录为 ${formatInteger(values.length)} 条。`,
    `价格平均值为 ${formatFieldNumber(mappings.price, stats.mean)}，中位数为 ${formatFieldNumber(mappings.price, stats.median)}，最低值为 ${formatFieldNumber(mappings.price, stats.min)}，最高值为 ${formatFieldNumber(mappings.price, stats.max)}。`
  ];
  if (!dimensionResults.length) {
    insights.push("未选择商品、平台、城市或成色字段，因此仅生成价格整体统计。");
  } else {
    dimensionResults.forEach(({ field, rows }) => {
      const topAvg = [...rows].sort((a, b) => b.avg - a.avg)[0];
      insights.push(topAvg
        ? `按「${field}」比较，平均价格最高的分组是「${topAvg.name}」，平均值为 ${formatFieldNumber(mappings.price, topAvg.avg)}。该结果只说明样本中的价格差异，不代表市场整体水平。`
        : `字段「${field}」没有形成可比较的价格分组。`);
    });
  }
  if (trend) insights.push(buildTrendFactText(mappings.price, mappings.date, trend, "平均值"));

  return {
    prompt: "已生成二手商品价格模板分析。",
    cards: [
      ["模板", "二手商品价格分析"],
      ["价格字段", mappings.price],
      ["比较维度", dimensionResults.length ? formatInteger(dimensionResults.length) : "未选择"],
      ["价格中位数", formatFieldNumber(mappings.price, stats.median)]
    ],
    insights,
    table: dimensionResults.length ? multiDimensionNumericTable(mappings.price, dimensionResults) : simpleStatsTable(mappings.price, stats, values.length),
    mainChart: primaryDimension
      ? barChartConfig(`按 ${primaryDimension.field} 的平均价格 Top 10`, topLabels(primaryDimension.rows, "avg"), topValues(primaryDimension.rows, "avg"), "平均价格", "平均值", CHART_THEME.primary, mappings.price)
      : histogramChartConfig(`${mappings.price} 分布`, values, "记录数", mappings.price),
    secondaryChart: secondaryDimension
      ? barChartConfig(`按 ${secondaryDimension.field} 的平均价格 Top 10`, topLabels(secondaryDimension.rows, "avg"), topValues(secondaryDimension.rows, "avg"), "平均价格", "平均值", CHART_THEME.secondary, mappings.price)
      : histogramChartConfig(`${mappings.price} 分布`, values, "记录数", mappings.price),
    trendChart: trend ? lineChartConfig(`${mappings.price} 时间趋势`, trend.labels, trend.values, "平均价格", "平均值", mappings.price) : null
  };
}

function analyzeSurveyTemplate(mappings) {
  const selectedFields = Object.entries(mappings).filter(([, value]) => value);
  const numericSummaries = [];
  ["satisfaction", "age"].forEach((key) => {
    if (!mappings[key]) return;
    const values = getNumericValues(mappings[key]);
    if (values.length) numericSummaries.push({ key, field: mappings[key], stats: summarizeNumbers(values), count: values.length });
  });
  const categoricalFields = ["gender", "city", "choice"].map((key) => mappings[key]).filter(Boolean);
  const primaryCategory = categoricalFields[0] || "";
  const primaryRows = primaryCategory ? frequencyRows(primaryCategory) : [];
  const chartValues = numericSummaries[0] ? getNumericValues(numericSummaries[0].field) : [];
  const insights = [
    `问卷模板共选择 ${formatInteger(selectedFields.length)} 个字段，当前分析记录数为 ${formatInteger(getAnalysisRows().length)}。`,
    numericSummaries[0]
      ? `数值字段「${numericSummaries[0].field}」的平均值为 ${formatFieldNumber(numericSummaries[0].field, numericSummaries[0].stats.mean)}，中位数为 ${formatFieldNumber(numericSummaries[0].field, numericSummaries[0].stats.median)}。`
      : "未选择可解析的数值字段，因此未生成满意度或年龄的数值统计。",
    primaryRows[0]
      ? `分类字段「${primaryCategory}」中出现最多的类别是「${primaryRows[0].name}」，占比 ${formatPercent(primaryRows[0].ratio)}。该结果只描述样本分布。`
      : "未选择分类字段，因此未生成选择题或人口属性频数统计。"
  ];

  return {
    prompt: "已生成问卷调查模板分析。",
    cards: [
      ["模板", "问卷调查分析"],
      ["已选字段", formatInteger(selectedFields.length)],
      ["记录数", formatInteger(getAnalysisRows().length)],
      ["主分类字段", primaryCategory || "未选择"]
    ],
    insights,
    table: surveyTable(numericSummaries, categoricalFields),
    mainChart: primaryCategory
      ? barChartConfig(`${primaryCategory} Top 10`, topLabels(primaryRows), topValues(primaryRows, "count"), "频数", "频数", CHART_THEME.primary)
      : histogramChartConfig(numericSummaries[0]?.field || "数值字段", chartValues, "记录数", numericSummaries[0]?.field || ""),
    secondaryChart: numericSummaries[0]
      ? histogramChartConfig(`${numericSummaries[0].field} 分布`, chartValues, "记录数", numericSummaries[0].field)
      : barChartConfig("分类字段 Top 10", topLabels(primaryRows), topValues(primaryRows, "count"), "频数", "频数", CHART_THEME.secondary),
    trendChart: null
  };
}

function analyzeBehaviorTemplate(mappings) {
  const eventRows = frequencyRows(mappings.event);
  const userCounts = countNonMissingValues(mappings.userId);
  const uniqueUsers = new Set(getAnalysisRows().map((row) => displayValue(row[mappings.userId])).filter(Boolean)).size;
  const avgEventsPerUser = uniqueUsers ? userCounts / uniqueUsers : 0;
  const secondaryResults = [...new Set([mappings.channel, mappings.device].filter(Boolean))]
    .map((field) => ({ field, rows: frequencyRows(field) }));
  const primarySecondary = secondaryResults[0];
  const secondaryRows = primarySecondary ? primarySecondary.rows : topUserRows(mappings.userId);
  const trend = mappings.date ? buildTrendFromRows(mappings.date, null, "count") : null;
  const insights = [
    `用户行为模板使用「${mappings.userId}」作为用户 ID、「${mappings.event}」作为事件字段，有效事件记录为 ${formatInteger(userCounts)} 条。`,
    `唯一用户数为 ${formatInteger(uniqueUsers)}，平均每个用户对应 ${formatNumber(avgEventsPerUser)} 条事件记录。`,
    eventRows[0] ? `事件字段中出现最多的是「${eventRows[0].name}」，出现 ${formatInteger(eventRows[0].count)} 次，占比 ${formatPercent(eventRows[0].ratio)}。该结果不推断用户动机。` : "事件字段没有可统计的有效类别。"
  ];
  secondaryResults.forEach(({ field, rows }) => {
    if (rows[0]) insights.push(`辅助维度「${field}」中记录数最高的是「${rows[0].name}」，出现 ${formatInteger(rows[0].count)} 次。`);
  });
  if (trend) insights.push(buildTrendFactText("事件记录数", mappings.date, trend, "记录数"));

  return {
    prompt: "已生成用户行为模板分析。",
    cards: [
      ["模板", "用户行为分析"],
      ["唯一用户", formatInteger(uniqueUsers)],
      ["事件记录", formatInteger(userCounts)],
      ["人均事件", formatNumber(avgEventsPerUser)]
    ],
    insights,
    table: multiFieldFrequencyTable([{ field: mappings.event, rows: eventRows }, ...secondaryResults]),
    mainChart: barChartConfig(`${mappings.event} Top 10`, topLabels(eventRows), topValues(eventRows, "count"), "事件数", "记录数", CHART_THEME.primary),
    secondaryChart: barChartConfig(`${primarySecondary?.field || mappings.userId} Top 10`, topLabels(secondaryRows), topValues(secondaryRows, "count"), "记录数", "记录数", CHART_THEME.secondary),
    trendChart: trend ? lineChartConfig("事件记录时间趋势", trend.labels, trend.values, "记录数", "记录数") : null
  };
}

function buildSalesMetricRows(mappings) {
  let label = "";
  let formatField = "";
  if (mappings.salesAmount) {
    label = mappings.salesAmount;
    formatField = mappings.salesAmount;
  }
  else if (mappings.unitPrice && mappings.quantity) label = `${mappings.unitPrice} × ${mappings.quantity}`;
  else if (mappings.unitPrice) {
    label = mappings.unitPrice;
    formatField = mappings.unitPrice;
  }
  else if (mappings.quantity) {
    label = mappings.quantity;
    formatField = mappings.quantity;
  }
  else return null;

  const rows = getAnalysisRows()
    .map((row) => ({ row, value: getSalesMetricValue(row, mappings) }))
    .filter((item) => item.value !== null);
  rows.label = label;
  rows.formatField = formatField;
  return rows;
}

function getSalesMetricValue(row, mappings) {
  if (mappings.salesAmount) return toNumber(row[mappings.salesAmount]);
  if (mappings.unitPrice && mappings.quantity) {
    const price = toNumber(row[mappings.unitPrice]);
    const quantity = toNumber(row[mappings.quantity]);
    if (price === null || quantity === null) return null;
    const product = price * quantity;
    return Number.isFinite(product) ? product : null;
  }
  if (mappings.unitPrice) return toNumber(row[mappings.unitPrice]);
  if (mappings.quantity) return toNumber(row[mappings.quantity]);
  return null;
}

function groupNumericFieldByCategory(metricField, groupField) {
  const groups = new Map();
  getAnalysisRows().forEach((row) => {
    const value = toNumber(row[metricField]);
    if (value === null) return;
    const groupName = displayValue(row[groupField]) || "未填写";
    if (!groups.has(groupName)) groups.set(groupName, []);
    groups.get(groupName).push(value);
  });
  return summarizeGroupedValues(groups);
}

function groupMetricRowsByField(metricRows, groupField) {
  const groups = new Map();
  metricRows.forEach((item) => {
    const groupName = displayValue(item.row[groupField]) || "未填写";
    if (!groups.has(groupName)) groups.set(groupName, []);
    groups.get(groupName).push(item.value);
  });
  return summarizeGroupedValues(groups);
}

function summarizeGroupedValues(groups) {
  return Array.from(groups.entries())
    .map(([name, values]) => {
      const { min, max } = numericExtents(values);
      return {
        name,
        count: values.length,
        sum: sumNumbers(values),
        avg: mean(values),
        median: median(values),
        min,
        max
      };
    })
    .sort((a, b) => compareNumbersDescending(a.sum, b.sum));
}

function buildTrendFromRows(dateField, valueGetter, aggregate) {
  const items = [];
  getAnalysisRows().forEach((row) => {
    const date = parseFieldDate(dateField, row[dateField]);
    if (!date) return;
    if (aggregate === "count") {
      items.push({ date, value: 1 });
      return;
    }
    const value = valueGetter(row);
    if (value !== null) items.push({ date, value });
  });

  if (!items.length) return { labels: [], values: [] };
  items.sort((a, b) => a.date - b.date);
  const spanDays = (items[items.length - 1].date - items[0].date) / 86400000;
  const groupByMonth = spanDays > 120;
  const groups = new Map();
  items.forEach(({ date, value }) => {
    const key = getDateTrendKey(dateField, date, groupByMonth);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(value);
  });

  const labels = Array.from(groups.keys()).sort();
  const values = labels.map((label) => {
    const bucket = groups.get(label);
    if (aggregate === "avg") return bucket.length ? mean(bucket) : 0;
    if (aggregate === "median") return bucket.length ? median(bucket) : 0;
    return sumNumbers(bucket);
  });
  return { labels, values };
}

function buildTrendFactText(metricLabel, dateField, trend, valueLabel) {
  if (!trend.labels.length) return `已选择日期字段「${dateField}」，但没有足够的有效日期记录生成趋势。`;
  if (trend.labels.length === 1) return `日期字段「${dateField}」只有 1 个有效时间点，无法判断变化方向。`;
  const first = trend.values[0];
  const last = trend.values[trend.values.length - 1];
  if (!Number.isFinite(first) || !Number.isFinite(last)) {
    return `按日期字段「${dateField}」观察时，聚合结果超出浏览器可表示的有限数值范围，无法可靠判断变化方向。`;
  }
  const direction = last > first ? "上升" : last < first ? "下降" : "基本持平";
  return `按日期字段「${dateField}」观察，「${metricLabel}」${valueLabel}从 ${trend.labels[0]} 到 ${trend.labels[trend.labels.length - 1]} 呈${direction}变化；这只是时间序列数值变化，不说明原因。`;
}

function summarizeNumbers(values) {
  if (!values.length) {
    return { mean: 0, median: 0, min: 0, max: 0, std: 0 };
  }
  const { min, max } = numericExtents(values);
  return {
    mean: mean(values),
    median: median(values),
    min,
    max,
    std: standardDeviation(values)
  };
}

function numericGroupTable(groupField, metricField, rows, formatField = metricField) {
  return {
    headers: [groupField, "记录数", `${metricField} 求和`, `${metricField} 平均值`, "最小值", "最大值"],
    rows: rows.map((row) => [
      row.name,
      formatInteger(row.count),
      formatFieldNumber(formatField, row.sum),
      formatFieldNumber(formatField, row.avg),
      formatFieldNumber(formatField, row.min),
      formatFieldNumber(formatField, row.max)
    ])
  };
}

function multiDimensionNumericTable(metricField, dimensionResults) {
  return {
    headers: ["维度字段", "分组值", "记录数", `${metricField} 求和`, `${metricField} 平均值`, "最小值", "最大值"],
    rows: dimensionResults.flatMap(({ field, rows }) => rows.map((row) => [
      field,
      row.name,
      formatInteger(row.count),
      formatFieldNumber(metricField, row.sum),
      formatFieldNumber(metricField, row.avg),
      formatFieldNumber(metricField, row.min),
      formatFieldNumber(metricField, row.max)
    ]))
  };
}

function frequencyTable(field, rows) {
  return {
    headers: [field, "记录数", "占比"],
    rows: rows.map((row) => [row.name, formatInteger(row.count), formatPercent(row.ratio)])
  };
}

function multiFieldFrequencyTable(fieldResults) {
  return {
    headers: ["字段", "类别", "记录数", "占比"],
    rows: fieldResults.flatMap(({ field, rows }) => rows.map((row) => [
      field,
      row.name,
      formatInteger(row.count),
      formatPercent(row.ratio)
    ]))
  };
}

function simpleStatsTable(field, stats, count) {
  return {
    headers: ["字段", "有效数量", "平均值", "中位数", "最小值", "最大值", "标准差"],
    rows: [[field, formatInteger(count), formatFieldNumber(field, stats.mean), formatFieldNumber(field, stats.median), formatFieldNumber(field, stats.min), formatFieldNumber(field, stats.max), formatFieldNumber(field, stats.std)]]
  };
}

function surveyTable(numericSummaries, categoricalFields) {
  const rows = [];
  numericSummaries.forEach((item) => {
    rows.push([item.field, "数值字段", `有效 ${formatInteger(item.count)} 条，平均值 ${formatFieldNumber(item.field, item.stats.mean)}，中位数 ${formatFieldNumber(item.field, item.stats.median)}`]);
  });
  categoricalFields.forEach((field) => {
    const top = frequencyRows(field)[0];
    rows.push([field, "分类字段", top ? `Top 类别「${top.name}」，占比 ${formatPercent(top.ratio)}` : "没有有效类别"]);
  });
  return {
    headers: ["字段", "类型", "结果摘要"],
    rows
  };
}

function buildSimpleTableHtml(headers, rows) {
  return `
    <thead>
      <tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}</tr>
    </thead>
    <tbody>
      ${rows.map((row) => `
        <tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>
      `).join("")}
    </tbody>
  `;
}

function frequencyRows(field, limit = 50) {
  const counts = countValues(field);
  const total = Array.from(counts.values()).reduce((sum, count) => sum + count, 0);
  return Array.from(counts.entries())
    .map(([name, count]) => ({ name, count, ratio: total ? count / total : 0 }))
    .sort((a, b) => b.count - a.count)
    .slice(0, limit);
}

function topUserRows(field) {
  return frequencyRows(field, 10);
}

function topLabels(rows, valueKey = "sum") {
  return [...rows]
    .sort((a, b) => compareNumbersDescending(a[valueKey], b[valueKey]))
    .slice(0, 10)
    .map((item) => item.name);
}

function topValues(rows, valueKey) {
  return [...rows]
    .sort((a, b) => compareNumbersDescending(a[valueKey], b[valueKey]))
    .slice(0, 10)
    .map((item) => Number.isFinite(item[valueKey]) ? item[valueKey] : null);
}

function barChartConfig(title, labels, values, label, axisLabel, color, valueField = "") {
  return { type: "bar", title, labels, values, label, axisLabel, color, valueField };
}

function lineChartConfig(title, labels, values, label, axisLabel, valueField = "") {
  return {
    type: "line",
    title,
    labels,
    values,
    label,
    axisLabel,
    valueField,
    color: CHART_THEME.secondary,
    background: CHART_THEME.secondaryFill
  };
}

function histogramChartConfig(title, values, axisLabel, sourceField = "") {
  const bins = buildHistogram(values, 10, sourceField);
  return {
    type: "bar",
    title,
    labels: bins.map((bin) => bin.label),
    values: bins.map((bin) => bin.count),
    label: title,
    axisLabel,
    color: CHART_THEME.warning
  };
}

function fieldHasNumericData(field, rows = getAnalysisRows()) {
  return getNumericValues(field, rows).length > 0;
}

function fieldHasDateData(field, rows = getAnalysisRows()) {
  return rows.some((row) => Boolean(parseFieldDate(field, row[field])));
}

function countNonMissingValues(field, rows = getAnalysisRows()) {
  return rows.filter((row) => !isMissing(row[field])).length;
}

function buildTrendInsight(dateField) {
  const numericField = getProfilesByType("numeric")
    .find((profile) => fieldHasNumericData(profile.field))?.field;
  const trend = buildDateTrend(dateField, numericField);
  if (!trend.labels.length) {
    return `日期字段「${dateField}」没有可安全转换的有效日期，因此未生成趋势结论。`;
  }
  if (trend.values.length < 2) {
    return `日期字段「${dateField}」可用于趋势分析，但有效时间点不足以判断变化方向。`;
  }
  const first = trend.values[0];
  const last = trend.values[trend.values.length - 1];
  const direction = last > first ? "上升" : last < first ? "下降" : "基本持平";
  if (first === 0 && last !== 0) {
    return `基于日期字段「${dateField}」观察，${trend.label} 从 ${trend.labels[0]} 到 ${trend.labels[trend.labels.length - 1]} 呈${direction}变化；由于首期值为 0，百分比变化无法计算，绝对变化为 ${formatFieldNumber(numericField || "", Math.abs(last - first))}。`;
  }
  const change = first === 0 ? 0 : (last - first) / Math.abs(first);
  return `基于日期字段「${dateField}」观察，${trend.label} 从 ${trend.labels[0]} 到 ${trend.labels[trend.labels.length - 1]} 呈${direction}变化，变化幅度约 ${formatPercent(Math.abs(change))}。`;
}

function buildDateTrend(dateField, numericField) {
  const items = [];
  getAnalysisRows().forEach((row) => {
    const date = parseFieldDate(dateField, row[dateField]);
    if (!date) return;
    const value = numericField ? toNumber(row[numericField]) : null;
    if (numericField && value === null) return;
    items.push({ date, value });
  });

  if (!items.length) {
    return { labels: [], values: [], label: "趋势", axisLabel: "值" };
  }

  items.sort((a, b) => a.date - b.date);
  const spanDays = (items[items.length - 1].date - items[0].date) / 86400000;
  const groupByMonth = spanDays > 120;
  const groups = new Map();

  items.forEach((item) => {
    const key = getDateTrendKey(dateField, item.date, groupByMonth);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(item.value);
  });

  const labels = Array.from(groups.keys()).sort();
  const hasMetric = Boolean(numericField);
  const values = labels.map((label) => {
    const bucket = groups.get(label);
    if (!hasMetric) return bucket.length;
    const valid = bucket.filter((value) => value !== null);
    return valid.length ? mean(valid) : 0;
  });

  return {
    labels,
    values,
    label: hasMetric
      ? `${numericField}${getDateGranularity(dateField) === "year" ? "年度" : " "}平均值趋势`
      : `${getDateGranularity(dateField) === "year" ? "年度" : ""}记录数趋势`,
    axisLabel: hasMetric ? "平均值" : "记录数"
  };
}

function aggregateByCategory(metricField, groupField, method) {
  const groups = new Map();
  getAnalysisRows().forEach((row) => {
    const group = displayValue(row[groupField]) || "未填写";
    const value = toNumber(row[metricField]);
    if (value === null) return;
    if (!groups.has(group)) groups.set(group, []);
    groups.get(group).push(value);
  });

  return Array.from(groups.entries())
    .map(([name, values]) => {
      let value = 0;
      if (method === "avg") value = mean(values);
      else if (method === "median") value = median(values);
      else value = sumNumbers(values);
      return { name, value, count: values.length };
    })
    .sort((a, b) => compareNumbersDescending(a.value, b.value));
}

function buildHistogram(values, binCount, field = "") {
  if (!values.length) return [];
  const { min, max } = numericExtents(values);
  if (min === max) return [{ label: formatFieldNumber(field, min), count: values.length }];
  const scale = Math.max(Math.abs(min), Math.abs(max));
  const normalizedMin = min / scale;
  const normalizedMax = max / scale;
  const normalizedWidth = (normalizedMax - normalizedMin) / binCount;
  if (!Number.isFinite(normalizedWidth) || normalizedWidth <= 0) {
    return [{ label: `${formatFieldNumber(field, min)} - ${formatFieldNumber(field, max)}`, count: values.length }];
  }
  const bins = Array.from({ length: binCount }, (_, index) => {
    const startRatio = index / binCount;
    const endRatio = (index + 1) / binCount;
    const start = min * (1 - startRatio) + max * startRatio;
    const end = index === binCount - 1
      ? max
      : min * (1 - endRatio) + max * endRatio;
    return {
      start,
      end,
      label: `${formatFieldNumber(field, start)} - ${formatFieldNumber(field, end)}`,
      count: 0
    };
  });
  values.forEach((value) => {
    const normalizedValue = value / scale;
    const rawIndex = Math.floor((normalizedValue - normalizedMin) / normalizedWidth);
    const index = Math.max(0, Math.min(rawIndex, binCount - 1));
    bins[index].count += 1;
  });
  return bins;
}

function chartOptions(axisLabel, valueField = "") {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { display: false },
      tooltip: {
        mode: "index",
        intersect: false,
        displayColors: false,
        backgroundColor: "rgba(22, 24, 40, 0.94)",
        titleColor: "#ffffff",
        bodyColor: "#e5e7ef",
        padding: 11,
        cornerRadius: 8,
        ...(valueField ? {
          callbacks: {
            label(context) {
              const datasetLabel = context.dataset?.label ? `${context.dataset.label}: ` : "";
              const value = typeof context.parsed?.y === "number" ? context.parsed.y : Number(context.raw);
              return `${datasetLabel}${formatFieldNumber(valueField, value)}`;
            }
          }
        } : {})
      }
    },
    scales: {
      x: {
        ticks: { maxRotation: 0, minRotation: 0, color: CHART_THEME.text, autoSkip: true, maxTicksLimit: 9 },
        grid: { display: false },
        border: { display: false }
      },
      y: {
        beginAtZero: true,
        title: { display: true, text: axisLabel, color: CHART_THEME.text },
        ticks: {
          color: CHART_THEME.text,
          ...(valueField ? { callback: (value) => formatFieldNumber(valueField, Number(value)) } : {})
        },
        grid: { color: CHART_THEME.grid },
        border: { display: false }
      }
    }
  };
}

function drawEmptyChart(canvas, text) {
  const context = canvas.getContext("2d");
  context.clearRect(0, 0, canvas.width, canvas.height);
  context.save();
  context.font = '13px Inter, "Segoe UI", "Microsoft YaHei", sans-serif';
  context.fillStyle = CHART_THEME.text;
  context.textAlign = "center";
  context.fillText(text, canvas.width / 2, canvas.height / 2);
  context.restore();
}

function resetAnalysisState() {
  Object.keys(state.charts).forEach((id) => destroyChart(id));
  state.charts = {};
  state.rows = [];
  state.filteredRows = null;
  state.filterSourceRows = null;
  state.fields = [];
  state.profiles = [];
  state.numericStats = [];
  state.qualityNumericStats = [];
  state.categoryStats = [];
  state.duplicateRows = 0;
  state.totalMissing = 0;
  state.parseWarnings = [];
  state.fieldTypeDrafts = Object.create(null);
  state.pendingExcel = null;
  state.sourceFileName = "";
  state.sourceSheetName = "";
  state.sourceType = "";
  state.analysisCompletedAt = null;
  state.insights = [];
  state.customAnalysis = null;
  state.filters = createEmptyFilterConfig();
  state.categoryFilterOptions = [];
  state.categoryFilterDraftValues = new Set();

  document.body.classList.remove("has-results");
  setDataExportStatus("");
  setReportExportStatus("");

  dom.sourceName.textContent = "";
  dom.sourceName.removeAttribute("title");
  dom.reportTitle.textContent = "数据分析工作区";
  dom.overviewCards.innerHTML = "";
  dom.insights.innerHTML = "";
  dom.previewTable.innerHTML = "";
  dom.schemaTable.innerHTML = "";
  dom.fieldConfigMessage.textContent = "";
  dom.fieldConfigMessage.className = "field-config-message";
  dom.applyFieldConfig.disabled = true;
  dom.restoreAutoFieldTypes.disabled = true;
  dom.categoryFilterField.innerHTML = "";
  dom.categoryFilterSearch.value = "";
  dom.categoryFilterValues.innerHTML = "";
  dom.categoryFilterControls.classList.add("hidden");
  dom.categoryFilterEmpty.classList.remove("hidden");
  dom.categoryFilterSelectionCount.textContent = "未选择字段";
  dom.numericFilterField.innerHTML = "";
  dom.numericFilterMin.value = "";
  dom.numericFilterMax.value = "";
  dom.numericFilterRange.textContent = "请选择字段";
  dom.dateFilterField.innerHTML = "";
  dom.dateFilterStart.value = "";
  dom.dateFilterEnd.value = "";
  dom.dateFilterRange.textContent = "请选择字段";
  dom.filterOriginalCount.textContent = "0";
  dom.filterCurrentCount.textContent = "0";
  dom.filterExcludedCount.textContent = "0";
  setFilterMessage("尚未应用筛选，当前分析使用全部记录。", "");
  dom.qualitySummary.innerHTML = "";
  dom.qualityTable.innerHTML = "";
  dom.outlierTable.innerHTML = "";
  dom.numericStatsTable.innerHTML = "";
  dom.categoryStats.innerHTML = "";
  dom.excelWorkbookName.textContent = "";
  dom.excelSheetSelect.innerHTML = "";
  dom.sheetSelectionMessage.textContent = "";
  dom.sheetSelectionMessage.className = "sheet-selection-message";
  dom.sheetSelectionPanel.classList.add("hidden");
  dom.v2FieldPrompt.textContent = "";
  dom.v2FieldPrompt.className = "analysis-prompt";
  dom.v2SummaryCards.innerHTML = "";
  dom.v2ResultSummary.textContent = "";
  dom.v2PrimaryTable.innerHTML = "";
  dom.v2SumTable.innerHTML = "";
  dom.v2AvgTable.innerHTML = "";
  dom.templatePrompt.textContent = "";
  dom.templatePrompt.className = "analysis-prompt";
  dom.templateSummaryCards.innerHTML = "";
  dom.templateInsights.innerHTML = "";
  dom.templatePrimaryTable.innerHTML = "";
  dom.v2AnalysisResults.classList.add("hidden");
  dom.templateResults.classList.add("hidden");
  dom.templateMappingForm.innerHTML = "";
  document.querySelectorAll("#customSection details, #insightsSection details").forEach((details) => {
    details.open = false;
  });
  document.querySelectorAll(".desktop-results-nav a").forEach((link, index) => {
    link.classList.toggle("active", index === 0);
    if (index === 0) link.setAttribute("aria-current", "location");
    else link.removeAttribute("aria-current");
  });
  dom.mobileSectionNav.value = "overviewSection";
  dom.results.classList.add("hidden");
  updateGlobalFilterStatus();
  updateExportButtons();
}

function destroyChart(id) {
  if (state.charts[id]) {
    state.charts[id].destroy();
    state.charts[id] = null;
  }
}

function countDuplicateRows(rows, fields) {
  const seen = new Set();
  let duplicates = 0;
  rows.forEach((row) => {
    const key = buildExactRowKey(row, fields);
    if (seen.has(key)) duplicates += 1;
    else seen.add(key);
  });
  return duplicates;
}

function deduplicateRows(rows, fields) {
  const seen = new Set();
  return rows.filter((row) => {
    const key = buildExactRowKey(row, fields);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function buildExactRowKey(row, fields) {
  return JSON.stringify(fields.map((field) => {
    if (!Object.prototype.hasOwnProperty.call(row, field)) return ["absent"];
    const value = row[field];
    if (value === undefined) return ["undefined"];
    if (value === null) return ["null"];
    return [typeof value, String(value)];
  }));
}

function countValues(field, rows = getAnalysisRows()) {
  const counts = new Map();
  rows.forEach((row) => {
    if (isMissing(row[field])) return;
    const value = displayValue(row[field]);
    counts.set(value, (counts.get(value) || 0) + 1);
  });
  return counts;
}

function getNumericValues(field, rows = getAnalysisRows()) {
  return rows
    .map((row) => toNumber(row[field]))
    .filter((value) => value !== null);
}

function getProfilesByType(typeKey) {
  return state.profiles.filter((profile) => profile.typeKey === typeKey);
}

function getActiveProfiles() {
  return state.profiles.filter((profile) => profile.typeKey !== "ignore");
}

function getFieldProfile(field) {
  return state.profiles.find((profile) => profile.field === field);
}

function parseFieldDate(field, value) {
  return toDate(value, getFieldProfile(field)?.dateStrategy || {});
}

function getDateGranularity(field) {
  return getFieldProfile(field)?.dateStrategy?.granularity || "day";
}

function getDateTrendKey(field, date, groupByMonth = false) {
  if (getDateGranularity(field) === "year") {
    return String(date.getUTCFullYear());
  }
  if (groupByMonth) {
    return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}`;
  }
  return date.toISOString().slice(0, 10);
}

function setSelectOptions(select, values, placeholder = "", allowEmpty = false) {
  const placeholderOption = placeholder
    ? `<option value="">${escapeHtml(placeholder)}</option>`
    : "";
  const valueOptions = values.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`).join("");
  select.innerHTML = values.length || allowEmpty
    ? `${placeholderOption}${valueOptions}`
    : `<option value="">无可用字段</option>`;
  select.disabled = values.length === 0 && !allowEmpty;
}

function missingRateLevel() {
  const denominator = state.rows.length * state.fields.length;
  const rate = denominator ? state.totalMissing / denominator : 0;
  if (rate >= 0.2) return "danger";
  if (rate > 0) return "warning";
  return "success";
}

function emptyTable(message) {
  return `
    <thead><tr><th>提示</th></tr></thead>
    <tbody><tr><td>${escapeHtml(message)}</td></tr></tbody>
  `;
}

function normalizeFieldName(field) {
  return String(field ?? "").replace(/^\uFEFF/, "").trim();
}

function normalizeParsedRow(row, fieldPairs) {
  return fieldPairs.reduce((normalized, pair) => {
    normalized[pair.clean] = Object.prototype.hasOwnProperty.call(row, pair.raw)
      ? row[pair.raw]
      : row[pair.clean];
    return normalized;
  }, Object.create(null));
}

function isMissing(value) {
  return value === null || value === undefined || String(value).trim() === "";
}

function normalizeValue(value) {
  return isMissing(value) ? "" : String(value).trim();
}

function displayValue(value) {
  return isMissing(value) ? "" : String(value).trim();
}

function isIdFieldName(field) {
  const name = String(field || "").trim();
  return /(^|[_\-\s])(id|uuid|guid|code|key|number|no)([_\-\s]|$)/i.test(name)
    || /[a-z0-9](?:Id|ID|Uuid|UUID|Guid|GUID|Code|CODE|Key|KEY|Number|NUMBER|No|NO)$/.test(name)
    || /^(?:user|customer|order|product|account|session|event|device|student|employee|member|record)(?:id|uuid|guid|code|key|number|no)$/i.test(name)
    || /编号|编码|代码|订单号|账号|工号|主键|用户ID|事件ID|会话ID/i.test(name)
    || /(?:学号|学籍号|考生号|准考证号|身份证号|身份证号码|证件号|手机号|手机号码|流水号|序列号|设备号|会员号|合同号|票据号|批次号|档案号)(?:$|[_\-\s（(])/i.test(name);
}

function isYearFieldName(field) {
  const name = String(field || "").trim();
  return /(年份|年度|学年|财年|会计年度|统计年度|统计年|报告年|出生年|入学年|毕业年)$/i.test(name)
    || /(^|[_\-\s])(year|fiscal_year|calendar_year|school_year|yr|fy)([_\-\s]|$)/i.test(name);
}

function buildDateStrategy(field, values) {
  let dmyEvidence = 0;
  let mdyEvidence = 0;
  let yearOnlyCount = 0;

  values.forEach((value) => {
    if (isDateObject(value)) return;
    const raw = String(value).trim();
    const yearText = raw.match(/^(\d{4})\s*年$/);
    if (isPlausibleYear(yearText ? yearText[1] : raw)) yearOnlyCount += 1;
    const match = matchDayOrMonthFirstDate(raw);
    if (!match) return;
    const first = Number(match[1]);
    const second = Number(match[3]);
    const year = Number(match[4]);
    if (first > 12 && second <= 12 && buildUtcDate(year, second, first)) dmyEvidence += 1;
    if (second > 12 && first <= 12 && buildUtcDate(year, first, second)) mdyEvidence += 1;
  });

  let order = "none";
  if (dmyEvidence && mdyEvidence) order = "conflict";
  else if (dmyEvidence) order = "dmy";
  else if (mdyEvidence) order = "mdy";

  const allowYearOnly = isYearFieldName(field)
    && values.length > 0
    && yearOnlyCount / values.length >= 0.8;

  return {
    order,
    allowYearOnly,
    granularity: allowYearOnly ? "year" : "day",
    dmyEvidence,
    mdyEvidence
  };
}

function isPlausibleYear(value) {
  if (!/^\d{4}$/.test(value)) return false;
  const year = Number(value);
  return year >= 1900 && year <= 2200;
}

function isDateObject(value) {
  return Object.prototype.toString.call(value) === "[object Date]"
    && typeof value.getTime === "function";
}

function matchYearFirstDate(raw) {
  return raw.match(new RegExp(`^(\\d{4})([-/.])(\\d{1,2})\\2(\\d{1,2})${DATE_TIME_SUFFIX_PATTERN}$`));
}

function matchDayOrMonthFirstDate(raw) {
  return raw.match(new RegExp(`^(\\d{1,2})([-/.])(\\d{1,2})\\2(\\d{4})${DATE_TIME_SUFFIX_PATTERN}$`));
}

function matchChineseDate(raw) {
  return raw.match(new RegExp(`^(\\d{4})\\s*年\\s*(\\d{1,2})\\s*月\\s*(\\d{1,2})\\s*日?${DATE_TIME_SUFFIX_PATTERN}$`));
}

function detectNumericFormat(values) {
  const convertibleValues = values.filter((value) => toNumber(value) !== null);
  if (!convertibleValues.length) return "number";
  return convertibleValues.every(hasPercentToken) ? "percent" : "number";
}

function hasPercentToken(value) {
  const raw = String(value).trim();
  const unsigned = /^\(.*\)$/.test(raw) ? raw.slice(1, -1).trim() : raw;
  return /%$/.test(unsigned);
}

function toNumber(value) {
  if (isMissing(value)) return null;
  const raw = String(value).trim();
  if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}/.test(raw)) return null;
  const isAccountingNegative = /^\(.*\)$/.test(raw);
  let numericText = (isAccountingNegative ? raw.slice(1, -1) : raw).trim();
  const isPercent = /%$/.test(numericText);
  if (isPercent) numericText = numericText.slice(0, -1).trim();

  const currencyMatches = numericText.match(/[$￥¥€£]/g) || [];
  if (currencyMatches.length > 1) return null;
  if (currencyMatches.length === 1) {
    const currencyIndex = numericText.search(/[$￥¥€£]/);
    const beforeCurrency = numericText.slice(0, currencyIndex).trim();
    const afterCurrency = numericText.slice(currencyIndex + 1).trim();
    const validPrefix = beforeCurrency === "" || /^[+-]$/.test(beforeCurrency);
    const validSuffix = afterCurrency === "";
    if (!validPrefix && !validSuffix) return null;
    numericText = `${beforeCurrency}${afterCurrency}`;
  }

  const plainNumberPattern = /^[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?$/;
  const commaGroupedPattern = /^[-+]?\d{1,3}(?:,\d{3})+(?:\.\d+)?(?:[eE][-+]?\d+)?$/;
  const spaceGroupedPattern = /^[-+]?\d{1,3}(?:[ \u00a0]\d{3})+(?:\.\d+)?(?:[eE][-+]?\d+)?$/;
  let cleaned = numericText;
  if (commaGroupedPattern.test(numericText)) cleaned = numericText.replaceAll(",", "");
  else if (spaceGroupedPattern.test(numericText)) cleaned = numericText.replace(/[ \u00a0]/g, "");
  else if (!plainNumberPattern.test(numericText)) return null;

  let parsed = Number(cleaned);
  if (!Number.isFinite(parsed)) return null;
  if (isAccountingNegative) parsed = -Math.abs(parsed);
  if (isPercent) parsed /= 100;
  return parsed;
}

function toDate(value, strategy = {}) {
  if (isMissing(value)) return null;
  if (isDateObject(value)) {
    if (Number.isNaN(value.getTime())) return null;
    const localMidnight = value.getHours() === 0
      && value.getMinutes() === 0
      && value.getSeconds() === 0
      && value.getMilliseconds() === 0;
    const utcMidnight = value.getUTCHours() === 0
      && value.getUTCMinutes() === 0
      && value.getUTCSeconds() === 0
      && value.getUTCMilliseconds() === 0;
    const useUtcParts = !localMidnight && utcMidnight;
    return buildUtcDate(
      useUtcParts ? value.getUTCFullYear() : value.getFullYear(),
      (useUtcParts ? value.getUTCMonth() : value.getMonth()) + 1,
      useUtcParts ? value.getUTCDate() : value.getDate()
    );
  }

  const raw = String(value).trim();
  const yearText = raw.match(/^(\d{4})\s*年$/);
  const yearOnlyValue = yearText ? yearText[1] : raw;
  if (strategy.allowYearOnly && isPlausibleYear(yearOnlyValue)) {
    return buildUtcDate(Number(yearOnlyValue), 1, 1);
  }
  if (/^[-+]?\d+(\.\d+)?$/.test(raw)) return null;
  const chineseDate = matchChineseDate(raw);
  if (chineseDate) {
    return buildUtcDate(Number(chineseDate[1]), Number(chineseDate[2]), Number(chineseDate[3]));
  }

  const yearFirst = matchYearFirstDate(raw);
  if (yearFirst) {
    return buildUtcDate(Number(yearFirst[1]), Number(yearFirst[3]), Number(yearFirst[4]));
  }

  const dayOrMonthFirst = matchDayOrMonthFirstDate(raw);
  if (dayOrMonthFirst) {
    const first = Number(dayOrMonthFirst[1]);
    const second = Number(dayOrMonthFirst[3]);
    const year = Number(dayOrMonthFirst[4]);
    if (first > 12) return buildUtcDate(year, second, first);
    if (second > 12) return buildUtcDate(year, first, second);
    if (first === second) return buildUtcDate(year, first, second);
    if (strategy.order === "dmy") return buildUtcDate(year, second, first);
    if (strategy.order === "mdy") return buildUtcDate(year, first, second);
    return null;
  }

  // Parse English calendar dates explicitly: Date(string) applies the host
  // timezone and can silently roll an invalid day into the following month.
  const englishDate = raw.replace(/^(?:mon(?:day)?|tue(?:sday)?|wed(?:nesday)?|thu(?:rsday)?|fri(?:day)?|sat(?:urday)?|sun(?:day)?),?\s+/i, "");
  const monthFirst = englishDate.match(/^([a-z]+)\.?[\s/-]+(\d{1,2})(?:,\s*|[\s/-]+)(\d{4})(?:\s+(.+))?$/i);
  const dayFirst = englishDate.match(/^(\d{1,2})[\s/-]+([a-z]+)\.?(?:,\s*|[\s/-]+)(\d{4})(?:\s+(.+))?$/i);
  const englishYearFirst = englishDate.match(/^(\d{4})[\s/-]+([a-z]+)\.?[\s/-]+(\d{1,2})(?:\s+(.+))?$/i);
  if (!monthFirst && !dayFirst && !englishYearFirst) return null;
  const monthName = (monthFirst ? monthFirst[1] : (dayFirst || englishYearFirst)[2]).toLowerCase();
  const months = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"];
  const month = months.findIndex((name) => monthName === name || monthName === name.slice(0, 3) || (name === "september" && monthName === "sept")) + 1;
  const year = Number(englishYearFirst ? englishYearFirst[1] : (monthFirst || dayFirst)[3]);
  const day = Number(monthFirst ? monthFirst[2] : dayFirst ? dayFirst[1] : englishYearFirst[3]);
  const calendarDate = buildUtcDate(year, month, day);
  if (!calendarDate) return null;

  // Accept familiar time/zone suffixes, but never convert the written date
  // through their offset. Validate them rather than ignoring arbitrary text.
  const suffix = (monthFirst || dayFirst || englishYearFirst)[4] || "";
  const time = suffix.match(/^(\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?(?:\s*(am|pm))?(?:\s*(\S.*))?$/i);
  let timezone = suffix;
  if (time) {
    const hour = Number(time[1]);
    if (hour > 23 || Number(time[2]) > 59 || Number(time[3] || 0) > 59) return null;
    if (time[5] && (hour < 1 || hour > 12)) return null;
    timezone = time[6] || "";
  }
  if (timezone && !/^(?:(?:GMT|UTC)(?:[+-](?:[01]\d|2[0-3]):?[0-5]\d)?|Z|[+-](?:[01]\d|2[0-3]):?[0-5]\d|[ECMP][SD]T)(?:\s+\([a-z][a-z .+-]*\))?$/i.test(timezone)) return null;
  return calendarDate;
}

function buildUtcDate(year, month, day) {
  if (year < 1000 || month < 1 || month > 12 || day < 1 || day > 31) return null;
  const parsed = new Date(Date.UTC(year, month - 1, day));
  if (parsed.getUTCFullYear() !== year || parsed.getUTCMonth() !== month - 1 || parsed.getUTCDate() !== day) {
    return null;
  }
  return parsed;
}

function mean(values) {
  if (!values.length) return 0;
  const scale = values.reduce((max, value) => Math.max(max, Math.abs(value)), 0);
  if (scale === 0) return 0;
  const normalizedTotal = values.reduce((sum, value) => sum + value / scale, 0);
  return (normalizedTotal / values.length) * scale;
}

function sumNumbers(values) {
  if (!values.length) return 0;
  const scale = values.reduce((max, value) => Math.max(max, Math.abs(value)), 0);
  if (scale === 0) return 0;
  const normalizedTotal = values.reduce((sum, value) => sum + value / scale, 0);
  const result = normalizedTotal * scale;
  return Number.isFinite(result) ? result : null;
}

function compareNumbersDescending(left, right) {
  const leftIsFinite = Number.isFinite(left);
  const rightIsFinite = Number.isFinite(right);
  if (leftIsFinite && rightIsFinite) return right - left;
  if (leftIsFinite) return -1;
  if (rightIsFinite) return 1;
  return 0;
}

function median(values) {
  if (!values.length) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  const middle = Math.floor(sorted.length / 2);
  if (sorted.length % 2) return sorted[middle];
  const lower = sorted[middle - 1];
  const upper = sorted[middle];
  const sum = lower + upper;
  return Number.isFinite(sum) ? sum / 2 : lower / 2 + upper / 2;
}

function standardDeviation(values, avg = mean(values)) {
  if (!values.length) return 0;
  const scale = values.reduce((max, value) => Math.max(max, Math.abs(value)), Math.abs(avg));
  if (scale === 0) return 0;
  const normalizedAverage = avg / scale;
  const variance = values.reduce((sum, value) => {
    const difference = value / scale - normalizedAverage;
    return sum + difference * difference;
  }, 0) / values.length;
  return Math.sqrt(variance) * scale;
}

function quantile(sortedValues, q) {
  if (!sortedValues.length) return 0;
  const position = (sortedValues.length - 1) * q;
  const base = Math.floor(position);
  const rest = position - base;
  if (sortedValues[base + 1] !== undefined) {
    return sortedValues[base] * (1 - rest) + sortedValues[base + 1] * rest;
  }
  return sortedValues[base];
}

function formatNumber(value) {
  if (!Number.isFinite(value)) return "-";
  return new Intl.NumberFormat("zh-CN", {
    maximumFractionDigits: 2
  }).format(value);
}

function formatInteger(value) {
  return new Intl.NumberFormat("zh-CN", {
    maximumFractionDigits: 0
  }).format(value || 0);
}

function formatImportLimitSize(bytes) {
  return `${Math.round(bytes / (1024 * 1024))} MB`;
}

function formatPercent(value) {
  if (!Number.isFinite(value)) return "0%";
  return `${(value * 100).toFixed(1)}%`;
}

function numericExtents(values) {
  if (!values.length) return { min: 0, max: 0 };
  let min = Infinity;
  let max = -Infinity;
  values.forEach((value) => {
    if (value < min) min = value;
    if (value > max) max = value;
  });
  return { min, max };
}

function getNumericFormat(field) {
  return getFieldProfile(field)?.numericFormat || "number";
}

function formatFieldNumber(field, value) {
  if (!Number.isFinite(value)) return "-";
  if (getNumericFormat(field) === "percent") {
    return new Intl.NumberFormat("zh-CN", {
      style: "percent",
      maximumFractionDigits: 2
    }).format(value);
  }
  return formatNumber(value);
}

function numericAxisLabel(field, fallbackLabel) {
  return getNumericFormat(field) === "percent"
    ? `${fallbackLabel}（百分比）`
    : fallbackLabel;
}

const REPORT_CHART_TITLES = Object.freeze({
  numericDistributionChart: "数值字段分布",
  categoryTopChart: "分类字段 Top 10",
  dateTrendChart: "日期字段趋势",
  customAnalysisChart: "自定义分组图",
  v2TopGroupChart: "自定义分析 Top 10 分组",
  v2SelectedTrendChart: "自定义分析时间趋势",
  templateMainChart: "场景模板主图",
  templateSecondaryChart: "场景模板辅助图",
  templateTrendChart: "场景模板时间趋势"
});

const REPORT_CHART_HEADING_IDS = Object.freeze({
  numericDistributionChart: "numericChartTitle",
  categoryTopChart: "categoryChartTitle",
  dateTrendChart: "dateChartTitle",
  customAnalysisChart: "customAnalysisChartTitle",
  v2TopGroupChart: "v2TopGroupTitle",
  v2SelectedTrendChart: "v2TrendTitle",
  templateMainChart: "templateMainChartTitle",
  templateSecondaryChart: "templateSecondaryChartTitle",
  templateTrendChart: "templateTrendChartTitle"
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
  updateWorkflowSteps();
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
  if (!window.Papa || typeof window.Papa.unparse !== "function") {
    throw new Error("CSV 导出组件未加载，请检查网络连接后刷新页面重试。");
  }
  const sourceFields = [...fields];
  const orderedFields = sourceFields.map(sanitizeCsvExportValue);
  const data = rows.map((row) => sourceFields.map((field) => (
    Object.prototype.hasOwnProperty.call(row, field) && row[field] !== null && row[field] !== undefined
      ? sanitizeCsvExportValue(row[field])
      : ""
  )));
  return window.Papa.unparse(
    { fields: orderedFields, data },
    { header: true, newline: "\r\n", skipEmptyLines: false }
  );
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
    sampleValues: [...profile.sampleValues]
  }));
  const numericStats = state.numericStats.map((stat) => ({
    field: stat.field,
    count: stat.count,
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
    top: category.top.slice(0, 10).map((item) => ({
      name: item.name,
      count: item.count,
      ratio: formatPercent(item.ratio)
    }))
  }));

  return {
    version: "v2.1.2",
    fileName: getReportSourceFileName(),
    sheetName: state.sourceSheetName || "",
    sourceType: state.sourceType || (state.sourceSheetName ? "excel" : "csv"),
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
      }))
    },
    numericStats,
    categoryStats,
    dateTrends: buildReportDateTrends(),
    customAnalysis: buildReportCustomAnalysis(),
    insights: [...state.insights]
  };
}

function getReportSourceFileName() {
  if (state.sourceFileName) return state.sourceFileName;
  const sourceTitle = dom.sourceName?.title || dom.sourceName?.textContent || "";
  return sourceTitle.replace(/^已载入\s*·\s+/, "").split(" · ")[0] || "data-report";
}

function buildReportDateTrends() {
  const metricField = state.numericStats[0]?.field || "";
  return getProfilesByType("date").map((profile) => {
    const trend = buildDateTrend(profile.field, metricField);
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
    const valueLabel = metricField ? `${metricField}平均值` : "记录数";
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
  let current = state.customAnalysis;
  if (!current) {
    const metricField = dom.metricField?.value || "";
    const groupField = dom.groupField?.value || "";
    const metricIsValid = getProfilesByType("numeric").some((profile) => profile.field === metricField);
    const groupIsValid = getProfilesByType("category").some((profile) => profile.field === groupField);
    if (metricIsValid && groupIsValid) {
      const grouped = groupNumericFieldByCategory(metricField, groupField);
      if (grouped.length) {
        current = { metricField, groupField, aggregateMethod: dom.aggregateMethod?.value || "sum", dateField: "", grouped, trend: null };
      }
    }
  }

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
    version: reportData.version || "v2.1.2",
    fileName: reportData.fileName || "data-report",
    sheetName: reportData.sheetName || "",
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
    fieldTypes: (reportData.fieldTypes || []).map((field) => ({
      ...field,
      inferredType: field.inferredType || field.type || "—",
      appliedType: field.appliedType || field.type || "—",
      nonMissingCount: field.nonMissingCount ?? 0,
      uniqueCount: field.uniqueCount ?? 0,
      missingCount: field.missingCount ?? 0,
      conversionFailureCount: field.conversionFailureCount ?? 0,
      sampleValues: Array.isArray(field.sampleValues) ? field.sampleValues : []
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
      outliers: Array.isArray(quality.outliers) ? quality.outliers : []
    },
    numericStats: Array.isArray(reportData.numericStats) ? reportData.numericStats : [],
    categoryStats: Array.isArray(reportData.categoryStats) ? reportData.categoryStats : [],
    dateTrends: Array.isArray(reportData.dateTrends) ? reportData.dateTrends : [],
    customAnalysis: reportData.customAnalysis || {
      completed: false,
      message: "当前没有可导出的自定义分组分析结果。"
    },
    insights: Array.isArray(reportData.insights) ? reportData.insights : []
  };
}

function buildHtmlReport(reportData, chartImages = []) {
  reportData = normalizeReportSnapshot(reportData);
  const safeCharts = chartImages.filter((chart) => isSafeChartDataUrl(chart.dataUrl));
  const sheetMeta = reportData.sheetName
    ? `<div><dt>工作表名称</dt><dd>${escapeHtml(reportData.sheetName)}</dd></div>`
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
      <p><strong>指标字段：</strong>${escapeHtml(reportData.customAnalysis.metricField)}　
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
        field.conversionFailureCount,
        field.sampleValues.join("、") || "—"
      ]),
      "无字段类型结果。"
    )}
  </section>
  <section>
    <h2>数据质量报告</h2>
    <p>缺失值总数：<strong>${escapeHtml(reportData.quality.totalMissing)}</strong>；重复行：<strong>${escapeHtml(reportData.quality.duplicateRows)}</strong>；解析警告：<strong>${escapeHtml(reportData.quality.parseWarningCount)}</strong>。</p>
    <h3>各字段缺失情况</h3>
    ${buildHtmlReportTable(
      ["字段", "类型", "缺失值", "缺失率"],
      reportData.quality.missingByField.map((item) => [item.field, item.type, item.missingCount, item.missingRate]),
      "无字段质量结果。"
    )}
    <h3>数值异常值</h3>
    ${buildHtmlReportTable(
      ["字段", "异常值数量", "下界", "上界"],
      reportData.quality.outliers.map((item) => [item.field, item.outlierCount, item.lowerBound, item.upperBound]),
      "未识别到数值字段，暂无可用数值统计。"
    )}
  </section>
  <section>
    <h2>数值字段描述性统计</h2>
    ${buildHtmlReportTable(
      ["字段", "有效数量", "平均值", "中位数", "最小值", "最大值", "标准差"],
      reportData.numericStats.map((item) => [item.field, item.count, item.mean, item.median, item.min, item.max, item.standardDeviation]),
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
  const lines = [
    "# 数据分析报告",
    "",
    `- 文件名称：${escapeMarkdownText(reportData.fileName)}`,
    ...(reportData.sheetName ? [`- 工作表名称：${escapeMarkdownText(reportData.sheetName)}`] : []),
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
          field.conversionFailureCount,
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
    "### 数值异常值",
    "",
    reportData.quality.outliers.length
      ? buildMarkdownTable(
        ["字段", "异常值数量", "下界", "上界"],
        reportData.quality.outliers.map((item) => [item.field, item.outlierCount, item.lowerBound, item.upperBound])
      )
      : "未识别到数值字段，暂无可用数值统计。",
    "",
    "## 数值字段描述性统计",
    "",
    reportData.numericStats.length
      ? buildMarkdownTable(
        ["字段", "有效数量", "平均值", "中位数", "最小值", "最大值", "标准差"],
        reportData.numericStats.map((item) => [item.field, item.count, item.mean, item.median, item.min, item.max, item.standardDeviation])
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
      `- 指标字段：${escapeMarkdownText(reportData.customAnalysis.metricField)}`,
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
    .replace(/\r?\n/g, "<br>")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/([*_`#[\]])/g, "\\$1");
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

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function setStatus(message, tone = "") {
  dom.statusPanel.className = `status-panel${tone ? ` ${tone}` : ""}`;
  dom.statusPanel.textContent = message;
  dom.statusPanel.setAttribute("role", tone === "error" ? "alert" : "status");
  dom.statusPanel.setAttribute("aria-busy", String(!tone));
  setImportBusy(!tone);
}

function showError(message, importId = state.activeImportId) {
  if (!isCurrentImport(importId)) return;
  resetAnalysisState();
  setImportBusy(false);
  dom.statusPanel.className = "status-panel error";
  dom.statusPanel.textContent = message;
  dom.statusPanel.setAttribute("role", "alert");
  dom.statusPanel.setAttribute("aria-busy", "false");
  dom.results.classList.add("hidden");
}

function setReportExportStatus(message, tone = "") {
  if (!dom.reportExportStatus) return;
  dom.reportExportStatus.className = `inline-status${tone ? ` ${tone}` : ""}`;
  dom.reportExportStatus.textContent = message;
}

function setDataExportStatus(message, tone = "") {
  if (!dom.dataExportStatus) return;
  dom.dataExportStatus.className = `inline-status${tone ? ` ${tone}` : ""}`;
  dom.dataExportStatus.textContent = message;
}
