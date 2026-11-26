// 应用状态与界面：state 和 dom 映射、导入 / 筛选 / 字段配置三条刷新路径，以及概览（含分页的数据预览）、字段、质量（含问题行查看区）、统计、图表（含图表下的分类 / 趋势结论）和自定义分析的渲染，还有写法统一的表的检查表显示和作废。

const state = {
  rows: [],
  // v3.1.0: rows left out of the analysis, as indexes into state.rows, each kind sorted.
  // exclusionSourceRows is the state.rows they belong to; includedRows the rows counted in the
  // analysis (state.rows itself while nothing is left out). Every commit replaces them.
  excludedRows: { suspectedTotal: [], duplicateRows: [] },
  exclusionSourceRows: null,
  includedRows: null,
  // findRepeatedRowIndexes() of state.rows, made the first time it is needed (never on import),
  // kept until the next import; repeatedRowSource is the state.rows it was made from.
  repeatedRowIndexes: null,
  repeatedRowSource: null,
  // Suspected totals a rebuild brought back into the analysis in this import (indexes), so the
  // next rebuild that flags one again can say so once.
  restoredTotalRows: new Set(),
  // What the last setRowExclusion() did, for its message: { kind, on, rows, changed, before,
  // after }, or null.
  lastExclusion: null,
  filteredRows: null,
  filterSourceRows: null,
  fields: [],
  profiles: [],
  numericStats: [],
  qualityNumericStats: [],
  categoryStats: [],
  duplicateRows: 0,
  // v3.4.0: the groups of exactly repeated rows countDuplicateRows() found on import (every row
  // of each group, the kept one too); only imports replace it.
  duplicateRowGroups: [],
  // v3.4.0: buildCheckupSummary(), made on commit, rebuild (field configuration, placeholder
  // confirmations) and after rows are left out or brought back; filters do not change it.
  checkupSummary: null,
  // buildIdDuplicateSummary() of the original rows; imports and field
  // configuration rebuild it, filters do not.
  idDuplicates: [],
  // Field -> placeholder keys confirmed as missing. Every commit replaces it with
  // an empty object; filters never change it.
  placeholderConfirmations: Object.create(null),
  // buildPlaceholderSummary() of the original rows; imports and field
  // configuration rebuild it, filters do not.
  placeholders: [],
  // buildNumericLookingSummary() of the profiles; imports, field configuration
  // and placeholder confirmations rebuild it, filters do not.
  numericLooking: [],
  // findTotalLabelRows() of the original rows; only imports replace it.
  totalLabelRows: [],
  // buildSuspectedTotalSummary() of the original rows; imports, field configuration and
  // placeholder confirmations rebuild it, filters do not.
  suspectedTotals: [],
  // v3.3.0: the labelled rows the totals check ended with (checkSuspectedTotals()), without
  // the label groups that mostly do not add up; written with suspectedTotals, for "另有 N 行".
  suspectedTotalLabelCount: 0,
  // Row object -> { field: "#DIV/0!" }, set only by commitTabularData() from an Excel
  // import; those cells read as "" and count as missing.
  sourceFormulaErrors: new WeakMap(),
  // buildFormulaErrorSummary() of the imported rows; only imports replace it.
  formulaErrors: [],
  // Row object -> { field: true } for Excel date cells, which read as YYYY-MM-DD text; set
  // only by commitTabularData() from an Excel import.
  sourceDateCells: new WeakMap(),
  // buildDateStyleSummary(), buildConstantFieldSummary() and buildLookAlikeSummary(); imports,
  // field configuration and placeholder confirmations rebuild them, filters do not.
  dateStyles: [],
  constantFields: [],
  lookAlikes: [],
  // createFormatScans(): what those summaries read from the rows; only imports replace it.
  formatScans: createFormatScans(),
  // Field -> typeKey applied by the user that differs from the automatic type.
  // Placeholder confirmations keep these types; "恢复自动识别" clears them.
  manualFieldTypes: Object.create(null),
  totalMissing: 0,
  parseWarnings: [],
  numberFormatHints: Object.create(null),
  // null, or { headerRowNumber, skippedRows: [{ rowNumber, preview }] } when the
  // import skipped title rows above the header. Only imports replace it.
  headerDetection: null,
  // Row object -> file row (Excel row, or the CSV line where the record starts),
  // set only by commitTabularData(); read it through getSourceRowNumber().
  sourceRowNumbers: new WeakMap(),
  // Data preview page (1-based); renderPreview() clamps it. Imports, filters and
  // field configuration go back to page 1.
  previewPage: 1,
  // listProblemCategories() as of the last renderQuality(): the 查看 buttons carry an
  // index into it.
  problemCategories: [],
  // The open problem-row viewer, or null: { category: { kind, field, key }, page,
  // list } with list from listProblemRows(), kept while it is open. Imports close it;
  // field configuration and placeholder confirmations list it again; filters never.
  problemView: null,
  // The 查看 button that opened the viewer, for the focus when it closes.
  problemViewOpener: null,
  fieldTypeDrafts: Object.create(null),
  pendingExcel: null,
  // v3.2.1: { bytes, fileName, sheetNames } of the Excel file with more than one worksheet under
  // analysis, for 换一个工作表; null for a CSV, a single-sheet file and after every reset.
  keptExcel: null,
  activeFileReader: null,
  activeImportId: 0,
  charts: {},
  sourceFileName: "",
  sourceSheetName: "",
  sourceType: "",
  analysisCompletedAt: null,
  chartInsights: { category: "", trend: "" },
  // v3.2.0: the figures behind chartInsights.trend (computeTrendInsight()), or null.
  trendInsightData: null,
  customAnalysis: null,
  // buildStandardizedExport() as of the last 检查标准化改动, or null. Imports, field
  // configuration and placeholder confirmations clear it; filters never.
  standardizedExport: null,
  // v3.0.0: the import's message { message, tone, hasNote, summary }, set by every commit and
  // put back by returnToAnalysis(); a new import (resetAnalysisState()) clears it.
  importStatus: null,
  // v3.0.0: the module shown when 换一个文件 left the workspace, for returnToAnalysis().
  returnModule: null,
  // v3.0.0 F2: the field table's search and type filter (more than 20 fields). Applying the
  // field configuration keeps them; imports and resets clear them. Nothing is stored.
  schemaFilter: { query: "", kind: "all" },
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
  statusNote: document.getElementById("statusNote"),
  statusNoteSummary: document.getElementById("statusNoteSummary"),
  statusNoteToggle: document.getElementById("statusNoteToggle"),
  sheetSelectionPanel: document.getElementById("sheetSelectionPanel"),
  excelWorkbookName: document.getElementById("excelWorkbookName"),
  excelSheetSelect: document.getElementById("excelSheetSelect"),
  confirmSheetSelection: document.getElementById("confirmSheetSelection"),
  sheetSelectionMessage: document.getElementById("sheetSelectionMessage"),
  results: document.getElementById("results"),
  returnToAnalysis: document.getElementById("returnToAnalysis"),
  reportTitle: document.getElementById("reportTitle"),
  sourceName: document.getElementById("sourceName"),
  topbarFileName: document.getElementById("topbarFileName"),
  sidebarToggle: document.getElementById("sidebarToggle"),
  appSidebar: document.getElementById("appSidebar"),
  sidebarClose: document.getElementById("sidebarClose"),
  sidebarBackdrop: document.getElementById("sidebarBackdrop"),
  sidebarFileSize: document.getElementById("sidebarFileSize"),
  sidebarFieldsToggle: document.getElementById("sidebarFieldsToggle"),
  sidebarFieldCount: document.getElementById("sidebarFieldCount"),
  sidebarFieldList: document.getElementById("sidebarFieldList"),
  sidebarFilterSummary: document.getElementById("sidebarFilterSummary"),
  sidebarHeaderRow: document.getElementById("sidebarHeaderRow"),
  sidebarHeaderValue: document.getElementById("sidebarHeaderValue"),
  sidebarPlaceholderRow: document.getElementById("sidebarPlaceholderRow"),
  sidebarPlaceholderValue: document.getElementById("sidebarPlaceholderValue"),
  globalFilterStatus: document.getElementById("globalFilterStatus"),
  globalFilterSummary: document.getElementById("globalFilterSummary"),
  globalClearFilters: document.getElementById("globalClearFilters"),
  overviewCards: document.getElementById("overviewCards"),
  checkupPanel: document.getElementById("checkupPanel"),
  checkupTotal: document.getElementById("checkupTotal"),
  checkupExportReport: document.getElementById("checkupExportReport"),
  checkupBar: document.getElementById("checkupBar"),
  checkupLegend: document.getElementById("checkupLegend"),
  previewSize: document.getElementById("previewSize"),
  categoryChartInsight: document.getElementById("categoryChartInsight"),
  dateChartInsight: document.getElementById("dateChartInsight"),
  previewTable: document.getElementById("previewTable"),
  previewPageInfo: document.getElementById("previewPageInfo"),
  previewFirstPage: document.getElementById("previewFirstPage"),
  previewPrevPage: document.getElementById("previewPrevPage"),
  previewNextPage: document.getElementById("previewNextPage"),
  previewLastPage: document.getElementById("previewLastPage"),
  schemaTable: document.getElementById("schemaTable"),
  schemaFilter: document.getElementById("schemaFilter"),
  schemaSearch: document.getElementById("schemaSearch"),
  schemaKinds: document.getElementById("schemaKinds"),
  schemaFilterEmpty: document.getElementById("schemaFilterEmpty"),
  fieldConfigMessage: document.getElementById("fieldConfigMessage"),
  fieldConfigToolbar: document.getElementById("fieldConfigToolbar"),
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
  filterNotCountedItem: document.getElementById("filterNotCountedItem"),
  filterNotCountedCount: document.getElementById("filterNotCountedCount"),
  qualityScopeNote: document.getElementById("qualityScopeNote"),
  duplicateExclusion: document.getElementById("duplicateExclusion"),
  duplicateExclusionText: document.getElementById("duplicateExclusionText"),
  duplicateExclusionAction: document.getElementById("duplicateExclusionAction"),
  duplicateExclusionMessage: document.getElementById("duplicateExclusionMessage"),
  suspectedTotalExcludedNote: document.getElementById("suspectedTotalExcludedNote"),
  suspectedTotalMessage: document.getElementById("suspectedTotalMessage"),
  suspectedTotalTitle: document.getElementById("suspectedTotalTitle"),
  statisticsContainsNote: document.getElementById("statisticsContainsNote"),
  chartsContainsNote: document.getElementById("chartsContainsNote"),
  customContainsNote: document.getElementById("customContainsNote"),
  filterMessage: document.getElementById("filterMessage"),
  applyFilters: document.getElementById("applyFilters"),
  clearAllFilters: document.getElementById("clearAllFilters"),
  qualitySection: document.getElementById("qualitySection"),
  qualitySummary: document.getElementById("qualitySummary"),
  parseWarningSummary: document.getElementById("parseWarningSummary"),
  problemRowsMessage: document.getElementById("problemRowsMessage"),
  problemRowsSection: document.getElementById("problemRowsSection"),
  problemRowsTitle: document.getElementById("problemRowsTitle"),
  problemRowsClose: document.getElementById("problemRowsClose"),
  problemRowsExport: document.getElementById("problemRowsExport"),
  problemRowsExportStatus: document.getElementById("problemRowsExportStatus"),
  problemRowsCategory: document.getElementById("problemRowsCategory"),
  problemRowsSummary: document.getElementById("problemRowsSummary"),
  problemRowsDraftNote: document.getElementById("problemRowsDraftNote"),
  problemRowsTable: document.getElementById("problemRowsTable"),
  problemRowsFirstPage: document.getElementById("problemRowsFirstPage"),
  problemRowsPrevPage: document.getElementById("problemRowsPrevPage"),
  problemRowsNextPage: document.getElementById("problemRowsNextPage"),
  problemRowsLastPage: document.getElementById("problemRowsLastPage"),
  qualityTable: document.getElementById("qualityTable"),
  outlierTable: document.getElementById("outlierTable"),
  suspectedTotalSection: document.getElementById("suspectedTotalSection"),
  suspectedTotalDetails: document.getElementById("suspectedTotalDetails"),
  formulaErrorSection: document.getElementById("formulaErrorSection"),
  formulaErrorDetails: document.getElementById("formulaErrorDetails"),
  numberConversionSection: document.getElementById("numberConversionSection"),
  numberConversionDetails: document.getElementById("numberConversionDetails"),
  numericLookingSection: document.getElementById("numericLookingSection"),
  numericLookingDetails: document.getElementById("numericLookingDetails"),
  dateStyleSection: document.getElementById("dateStyleSection"),
  dateStyleDetails: document.getElementById("dateStyleDetails"),
  constantFieldSection: document.getElementById("constantFieldSection"),
  constantFieldDetails: document.getElementById("constantFieldDetails"),
  lookAlikeSection: document.getElementById("lookAlikeSection"),
  lookAlikeDetails: document.getElementById("lookAlikeDetails"),
  placeholderDetails: document.getElementById("placeholderDetails"),
  placeholderMessage: document.getElementById("placeholderMessage"),
  idDuplicateDetails: document.getElementById("idDuplicateDetails"),
  numericStatsTable: document.getElementById("numericStatsTable"),
  categoryStats: document.getElementById("categoryStats"),
  numericChartField: document.getElementById("numericChartField"),
  categoryChartField: document.getElementById("categoryChartField"),
  dateChartField: document.getElementById("dateChartField"),
  dateChartMetric: document.getElementById("dateChartMetric"),
  dateChartAggregate: document.getElementById("dateChartAggregate"),
  numericChartTitle: document.getElementById("numericChartTitle"),
  categoryChartTitle: document.getElementById("categoryChartTitle"),
  dateChartTitle: document.getElementById("dateChartTitle"),
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
  dataExportOriginalCount: document.getElementById("dataExportOriginalCount"),
  dataExportFilteredCount: document.getElementById("dataExportFilteredCount"),
  dataExportFilteredBefore: document.getElementById("dataExportFilteredBefore"),
  dataExportFilteredAfter: document.getElementById("dataExportFilteredAfter"),
  dataExportOriginalBefore: document.getElementById("dataExportOriginalBefore"),
  dataExportOriginalAfter: document.getElementById("dataExportOriginalAfter"),
  dataExportExcludedNote: document.getElementById("dataExportExcludedNote"),
  exportFilteredCsv: document.getElementById("exportFilteredCsv"),
  dataExportStatus: document.getElementById("dataExportStatus"),
  checkStandardizedExport: document.getElementById("checkStandardizedExport"),
  standardizedExportGate: document.getElementById("standardizedExportGate"),
  standardizedExportPanel: document.getElementById("standardizedExportPanel"),
  standardizedExportSummary: document.getElementById("standardizedExportSummary"),
  standardizedExportTable: document.getElementById("standardizedExportTable"),
  exportStandardizedCsv: document.getElementById("exportStandardizedCsv"),
  exportStandardizedChanges: document.getElementById("exportStandardizedChanges"),
  standardizedExportStatus: document.getElementById("standardizedExportStatus"),
  exportHtmlReport: document.getElementById("exportHtmlReport"),
  reportExportStatus: document.getElementById("reportExportStatus"),
  // v3.2.0: the language switch, in the top bar of the upload page and at the sidebar's foot.
  languageSwitch: document.getElementById("languageSwitch"),
  sidebarLanguageSwitch: document.getElementById("sidebarLanguageSwitch"),
  // v3.2.1 换一个工作表: one button and list in the sidebar and one on the upload page (the same
  // markup, ids prefixed by the place).
  sidebarSheetToggle: document.getElementById("sidebarSheetToggle"),
  sidebarSheetPanel: document.getElementById("sidebarSheetPanel"),
  sidebarSheetList: document.getElementById("sidebarSheetList"),
  sidebarSheetMessage: document.getElementById("sidebarSheetMessage"),
  sidebarSheetLine: document.getElementById("sidebarSheetLine"),
  uploadSheetToggle: document.getElementById("uploadSheetToggle"),
  uploadSheetPanel: document.getElementById("uploadSheetPanel"),
  uploadSheetList: document.getElementById("uploadSheetList"),
  uploadSheetMessage: document.getElementById("uploadSheetMessage")
};

// A category field with more distinct values than this is too scattered for
// the automatic category conclusion.
const INSIGHT_CATEGORY_MAX_DISTINCT = 20;

const CHART_THEME = {
  primary: "#0068e0",
  primaryFill: "rgba(0, 104, 224, 0.12)",
  secondary: "#4792e9",
  secondaryFill: "rgba(71, 146, 233, 0.14)",
  warning: "#b54708",
  text: "#636366",
  grid: "rgba(28, 28, 30, 0.08)"
};

function setImportBusy(isBusy) {
  document.body.classList.toggle("is-loading", isBusy);
  dom.fileInput.disabled = isBusy;
  dom.loadSample.disabled = isBusy;
  dom.confirmSheetSelection.disabled = isBusy;
  dom.dropZone.setAttribute("aria-busy", String(isBusy));
  syncLanguageSwitches();
}

// ---- Interface language (v3.2.0) ----
// index.html keeps the Chinese. An element's data-i18n names the I18N_EN key of its text (its own
// text nodes only, so child elements such as a count stay where they are), data-i18n-attr
// ("aria-label=key title=key") those of its attributes. The Chinese is read from the page the
// first time an element is switched, so switching back writes it back exactly.
const staticChineseText = new WeakMap();

function readI18nAttributePairs(element) {
  return String(element.getAttribute("data-i18n-attr") || "").split(/\s+/).filter(Boolean).map((pair) => pair.split("="));
}

function rememberStaticText(element) {
  if (staticChineseText.has(element)) return staticChineseText.get(element);
  const key = element.getAttribute("data-i18n");
  const textNode = key ? Array.from(element.childNodes || []).find((node) => node.nodeType === 3 && node.nodeValue.trim()) : null;
  const chinese = {
    textNode,
    text: textNode ? textNode.nodeValue : null,
    attributes: Object.fromEntries(readI18nAttributePairs(element).map(([name]) => [name, element.getAttribute(name)]))
  };
  staticChineseText.set(element, chinese);
  return chinese;
}

// v3.2.0 commit 9: script.js reads the Chinese of the page at startup in both languages, before any
// script writes into it. A first switch after an import then never takes a script's text (a
// count, “收起”) for the page's own.
function rememberStaticLanguage() {
  if (typeof document.querySelectorAll !== "function") return;
  document.querySelectorAll("[data-i18n], [data-i18n-attr]").forEach(rememberStaticText);
}

function applyStaticLanguage() {
  if (typeof document.querySelectorAll !== "function") return;
  document.querySelectorAll("[data-i18n], [data-i18n-attr]").forEach((element) => {
    const key = element.getAttribute("data-i18n");
    const chinese = rememberStaticText(element);
    // A text node a script has replaced is no longer in the element: writing it changes nothing
    // on the page, and the script's text stays (commit 11 writes that again).
    if (chinese.textNode) {
      const english = currentLanguage === "en" ? I18N_EN[key] : undefined;
      // Keep the white space around the words, so the Chinese comes back byte for byte.
      chinese.textNode.nodeValue = english === undefined
        ? chinese.text
        : chinese.text.replace(chinese.text.trim(), () => english);
    }
    readI18nAttributePairs(element).forEach(([name, attributeKey]) => {
      // Only an attribute that still holds the page's own text (a chart's aria-label is the
      // script's once it is drawn).
      const current = element.getAttribute(name);
      if (current !== chinese.attributes[name] && current !== I18N_EN[attributeKey]) return;
      const english = currentLanguage === "en" ? I18N_EN[attributeKey] : undefined;
      element.setAttribute(name, english === undefined ? chinese.attributes[name] : english);
    });
  });
}

// Sets the language, <html lang>, the static text and the two switches. The analysis on the page
// is not drawn again here.
function setInterfaceLanguage(language) {
  currentLanguage = language === "en" ? "en" : "zh";
  if (document.documentElement) document.documentElement.lang = currentLanguage === "en" ? "en" : "zh-CN";
  applyStaticLanguage();
  [dom.languageSwitch, dom.sidebarLanguageSwitch].forEach((button) => {
    button.textContent = languageSwitchLabel();
    button.setAttribute("lang", currentLanguage === "en" ? "zh-CN" : "en");
  });
  // The navigation's words change width with the language.
  updateNavFit();
  // v3.2.0 commit 11: the analysis on the page, drawn again in the new language.
  rerenderForLanguage();
}

// v3.2.0 commit 11: everything the scripts wrote, written again in the current language from the
// state, nothing computed again (契约 5 of the plan): the filters, the rows left out, the module,
// the selects of the charts and the custom analysis, its result, the problem-row viewer's category
// and page, the preview page, unapplied column drafts and the standardized check all stay. Order
// as the refresh paths. Charts are destroyed and drawn again (契约 11).
function rerenderForLanguage() {
  // The messages on screen now: a render below may clear one (opening the problem rows again
  // clears its export status), so they are all written back at the end.
  const messages = new Map(oneOffMessages);
  rewriteOneOffMessages();
  rewriteStatusMessage();
  rewriteLanguageTexts();
  if (!state.rows.length || !state.analysisCompletedAt) return;
  const selects = [dom.numericChartField, dom.categoryChartField, dom.dateChartField, dom.dateChartMetric, dom.dateChartAggregate, dom.v2MetricField, dom.v2GroupField, dom.v2AggregateMethod, dom.v2DateField];
  const chosen = selects.map((select) => select.value);
  const problemPage = state.problemView ? state.problemView.page : null;

  renderOverview();
  renderFilterCounts();
  renderPreview();
  renderSchema();
  renderSidebar();
  rerenderFilterPanelForLanguage();
  renderQuality();
  // renderQuality() lists the open category again (refreshProblemView()) and goes to page 1.
  if (state.problemView && problemPage !== null) {
    state.problemView.page = problemPage;
    renderProblemView();
  }
  renderNumericStats();
  renderCategoryStats();
  populateControls({ keepChartSelections: true });
  selects.forEach((select, index) => { select.value = chosen[index]; });
  renderCharts();
  renderCategoryInsight();
  state.chartInsights.trend = state.trendInsightData ? formatTrendInsight(state.trendInsightData) : "";
  setChartInsight(dom.dateChartInsight, state.chartInsights.trend);
  if (state.customAnalysis && !dom.v2AnalysisResults.classList.contains("hidden")) drawV2Analysis();
  if (state.standardizedExport) renderStandardizedExport();
  updateExportButtons();
  // The messages the render functions above cleared or wrote over, back in the new language.
  messages.forEach(({ build, className, tone }, element) => writeOneOffMessage(element, className, build, tone));
}

// The filter panel's own texts: the first options of the three selects, the values' labels (the
// empty one is “（空值）” / “(Blank)”) and their list; the drafts and the chosen fields stay.
function rerenderFilterPanelForLanguage() {
  // The first option is written again in the select's HTML (setSelectOptions() wrote it first),
  // and the chosen field is put back (a new innerHTML resets it).
  [[dom.categoryFilterField, FILTER_NONE_LABELS.category], [dom.numericFilterField, FILTER_NONE_LABELS.numeric], [dom.dateFilterField, FILTER_NONE_LABELS.date]].forEach(([select, label]) => {
    const value = select.value;
    select.innerHTML = String(select.innerHTML).replace(/^<option value="">[^<]*<\/option>/, () => `<option value="">${escapeHtml(Lpick(label))}</option>`);
    select.value = value;
  });
  state.categoryFilterOptions.forEach((option) => {
    if (!option.value) option.label = Lpick(BLANK_GROUP_LABEL);
  });
  renderCategoryFilterValues();
}

// v3.2.0 commit 11: texts of fixed elements written from the state or a past action (the filter
// ranges, the category filter's hint), kept as functions so a switch writes them again.
const languageTexts = new Map();

function writeLanguageText(element, text) {
  element.textContent = typeof text === "function" ? text() : text;
  if (typeof text === "function") languageTexts.set(element, text);
  else languageTexts.delete(element);
}

function rewriteLanguageTexts() {
  languageTexts.forEach((build, element) => { element.textContent = build(); });
}

// v3.2.0 commit 9 (the review's 甲): when the eight module links do not fit beside the brand and
// 换一个文件, the navigation takes a row of its own (body.nav-wrap, style.css) instead of hiding
// its last items. Measured on the real widths, so no breakpoint depends on the words: take the
// class off, measure, put it back only if the links overflow. script.js calls this on every size
// change of the top bar (ResizeObserver) and of the window.
function updateNavFit() {
  const nav = document.getElementById("moduleNav");
  if (!nav || !document.body?.classList) return;
  document.body.classList.remove("nav-wrap");
  if (!document.body.classList.contains("has-results")) return;
  if (nav.scrollWidth > nav.clientWidth + 1) document.body.classList.add("nav-wrap");
}

// v3.2.0 commit 11: the switch's confirmation for screen readers (“Switched to English.”), in a
// role="status" element made at the first switch, at the end of <main> (always shown). A page
// never switched has no such element, as before; the sentence is written a moment after the
// element is there (so it is read out) and taken away after a few seconds.
let languageAnnouncer = null;
const languageAnnounceTimers = { write: null, clear: null };
const LANGUAGE_ANNOUNCE_DELAY_MS = 50;
const LANGUAGE_ANNOUNCE_CLEAR_MS = 5000;

function announceLanguageSwitch() {
  if (!languageAnnouncer) {
    languageAnnouncer = document.createElement("p");
    languageAnnouncer.className = "visually-hidden";
    languageAnnouncer.setAttribute("role", "status");
    languageAnnouncer.setAttribute("aria-live", "polite");
    document.getElementById("mainContent")?.appendChild(languageAnnouncer);
  }
  const element = languageAnnouncer;
  const sentence = L("界面已切换为中文。", "Switched to English.");
  clearTimeout(languageAnnounceTimers.write);
  clearTimeout(languageAnnounceTimers.clear);
  element.textContent = "";
  languageAnnounceTimers.write = setTimeout(() => {
    element.textContent = sentence;
    languageAnnounceTimers.clear = setTimeout(() => { element.textContent = ""; }, LANGUAGE_ANNOUNCE_CLEAR_MS);
    languageAnnounceTimers.clear?.unref?.();
  }, LANGUAGE_ANNOUNCE_DELAY_MS);
  languageAnnounceTimers.write?.unref?.();
}

// The switch: the other language, written into the address as ?lang= (the #module stays).
function switchInterfaceLanguage() {
  if (isLanguageSwitchLocked()) return;
  setInterfaceLanguage(currentLanguage === "en" ? "zh" : "en");
  announceLanguageSwitch();
  if (window.location && window.history && typeof URL === "function") {
    const url = new URL(window.location.href);
    url.searchParams.set("lang", currentLanguage);
    window.history.replaceState(null, "", `${url.pathname}${url.search}${url.hash}`);
  }
}

// No switching while a file is being read or a worksheet waits to be chosen. The worksheet list
// is shown (importer.js) right after setImportBusy(false), so script.js also calls
// syncLanguageSwitches() when #sheetSelectionPanel shows or hides.
function isLanguageSwitchLocked() {
  return document.body.classList.contains("is-loading") || Boolean(state.pendingExcel);
}

function syncLanguageSwitches() {
  const locked = isLanguageSwitchLocked();
  [dom.languageSwitch, dom.sidebarLanguageSwitch].forEach((button) => {
    if (button) button.disabled = locked;
  });
}

// An import opens the first module at the top of the page with the focus on its title
// (v3.0.0); the address stays as it was.
function showResultsWorkspace() {
  document.body.classList.add("has-results");
  setImportBusy(false);
  syncSidebarToggle();
  window.requestAnimationFrame(() => showModule("overviewSection", { hash: false }));
}

// The sidebar (v3.0.0): the file, its fields and the current state, for display only.
const SIDEBAR_TYPE_MARKS = Object.freeze({ id: Lp("#", "#"), date: Lp("日", "D"), category: Lp("Aa", "Aa"), numeric: Lp("12", "12"), ignore: Lp("—", "—") });

function renderSidebar() {
  dom.sidebarFileSize.textContent = L(`${formatInteger(state.rows.length)} 行 × ${formatInteger(state.fields.length)} 列`, `${enCount(state.rows.length, "row", "rows")} × ${enCount(state.fields.length, "column", "columns")}`);
  dom.sidebarFieldCount.textContent = formatInteger(state.fields.length);
  dom.sidebarFieldList.innerHTML = state.profiles.map((profile) => {
    const unit = profile.typeKey === "numeric" ? getFieldNumberUnit(profile.field) : "";
    return `<li class="sidebar-field${profile.typeKey === "ignore" ? " is-ignored" : ""}" title="${escapeHtml(profile.field)}"><span class="sidebar-type">${SIDEBAR_TYPE_MARKS[profile.typeKey] ? Lpick(SIDEBAR_TYPE_MARKS[profile.typeKey]) : ""}</span><span class="sidebar-field-name">${escapeHtml(profile.field)}</span>${unit ? `<span class="sidebar-unit">${escapeHtml(unit)}</span>` : ""}</li>`;
  }).join("");
  const headerRowNumber = state.headerDetection?.headerRowNumber;
  const showsHeader = Boolean(headerRowNumber) && headerRowNumber !== 1;
  dom.sidebarHeaderRow.classList.toggle("hidden", !showsHeader);
  dom.sidebarHeaderValue.textContent = showsHeader ? L(`第 ${formatInteger(headerRowNumber)} 行`, `row ${formatInteger(headerRowNumber)}`) : "";
  const confirmedCount = Object.values(state.placeholderConfirmations).reduce((sum, keys) => sum + keys.length, 0);
  dom.sidebarPlaceholderRow.classList.toggle("hidden", confirmedCount === 0);
  dom.sidebarPlaceholderValue.textContent = confirmedCount ? L(`${formatInteger(confirmedCount)} 处`, formatInteger(confirmedCount)) : "";
  renderSheetSwitches();
}

// ---- 换一个工作表 (v3.2.1) ----

// The sidebar's worksheet line and both lists, from state.keptExcel (nothing shows without it).
// Items carry their index in the file, never the name. Whether a list is open and its message
// stay as they are, so a language switch (which draws the sidebar again) keeps them. The upload
// page's button shows only with 回到「…」的分析 (showUploadPage() and returnToAnalysis() call this);
// without it that list is closed and its message (a failed switch's error) gone.
function renderSheetSwitches() {
  const kept = state.keptExcel;
  dom.sidebarSheetLine.textContent = kept ? formatSheetLine(state.sourceSheetName, kept.sheetNames.length) : "";
  dom.sidebarSheetLine.classList.toggle("hidden", !kept);
  SHEET_SWITCH_PLACES.forEach((place) => {
    const { toggle, list } = sheetSwitchElements(place);
    const shown = Boolean(kept) && (place === "sidebar" || !dom.returnToAnalysis.classList.contains("hidden"));
    toggle.classList.toggle("hidden", !shown);
    if (!shown) {
      setSheetSwitchOpen(place, false);
      writeSheetSwitchMessage(place, "");
    }
    if (!kept) {
      list.innerHTML = "";
      return;
    }
    list.innerHTML = kept.sheetNames.map((name, index) => (name === state.sourceSheetName
      ? `<li class="sheet-switch-current" aria-current="true"><span class="sheet-switch-name">${escapeHtml(name)}</span><span class="sheet-switch-now">${escapeHtml(Lpick(SHEET_SWITCH_OPEN_NOW))}</span></li>`
      : `<li><button class="sheet-switch-item" type="button" data-sheet-index="${index}"><span class="sheet-switch-name">${escapeHtml(name)}</span></button></li>`)).join("");
  });
}

function toggleSheetSwitch(place) {
  setSheetSwitchOpen(place, sheetSwitchElements(place).toggle.getAttribute("aria-expanded") !== "true");
}

// A click on another sheet of a list (one delegated listener per list, script.js). Returns the
// switch's promise (or undefined) so a test can wait for it.
function handleSheetSwitchClick(event, place) {
  const button = event.target.closest?.("[data-sheet-index]");
  if (!button || !state.keptExcel) return undefined;
  const sheetName = state.keptExcel.sheetNames[Number(button.dataset.sheetIndex)];
  return sheetName === undefined ? undefined : switchExcelSheet(sheetName, place);
}

// Escape closes an open list and gives the focus back to its button. It stops there, so the
// phone and narrow sidebar (which closes on Escape) closes only on the next one.
function handleSheetSwitchKeydown(event, place) {
  const { toggle } = sheetSwitchElements(place);
  if (event.key !== "Escape" || toggle.getAttribute("aria-expanded") !== "true") return;
  event.preventDefault();
  event.stopPropagation();
  setSheetSwitchOpen(place, false);
  toggle.focus();
}
// place: "sidebar" or "upload", the list a switch was started from.
const SHEET_SWITCH_PLACES = Object.freeze(["sidebar", "upload"]);

function sheetSwitchElements(place) {
  return {
    toggle: dom[`${place}SheetToggle`],
    panel: dom[`${place}SheetPanel`],
    list: dom[`${place}SheetList`],
    message: dom[`${place}SheetMessage`]
  };
}

function setSheetSwitchOpen(place, open) {
  const { toggle, panel } = sheetSwitchElements(place);
  toggle.setAttribute("aria-expanded", String(open));
  panel.classList.toggle("hidden", !open);
}

// While a sheet is read the list says it is busy. Its buttons stay enabled (disabling the focused
// one would drop the focus); switchExcelSheet() ignores clicks while the page is loading.
function setSheetSwitchBusy(place, busy) {
  sheetSwitchElements(place).list.setAttribute("aria-busy", String(busy));
}

// The reading message and a switch's error, under the list; a one-off message, so a language
// switch writes it again.
function writeSheetSwitchMessage(place, message, tone = "") {
  writeOneOffMessage(sheetSwitchElements(place).message, "sheet-switch-message", message, tone);
}

// After a failed switch: the item that was clicked, or the list's button.
function focusSheetSwitchItem(place, sheetName) {
  const { toggle, list } = sheetSwitchElements(place);
  const index = state.keptExcel ? state.keptExcel.sheetNames.indexOf(sheetName) : -1;
  const item = index >= 0 ? list.querySelector?.(`[data-sheet-index="${index}"]`) : null;
  (item || toggle).focus();
}

function clearSheetSwitches() {
  dom.sidebarSheetLine.textContent = "";
  dom.sidebarSheetLine.classList.add("hidden");
  SHEET_SWITCH_PLACES.forEach((place) => {
    const { toggle, list } = sheetSwitchElements(place);
    setSheetSwitchOpen(place, false);
    toggle.classList.add("hidden");
    list.innerHTML = "";
    list.removeAttribute("aria-busy");
    writeSheetSwitchMessage(place, "");
  });
}

// Open by default from 1100 px, closed below; once the user opens or closes it, body carries
// sidebar-open / sidebar-closed until the next reset (nothing is stored). Below 1100 px it
// covers the content: Escape and the backdrop close it and the focus stays inside it.
const SIDEBAR_WIDE_QUERY = "(min-width: 1100px)";

function isSidebarWide() {
  return typeof window.matchMedia === "function" ? window.matchMedia(SIDEBAR_WIDE_QUERY).matches : true;
}

function isSidebarOpen() {
  const classes = document.body.classList;
  if (classes.contains("sidebar-open")) return true;
  if (classes.contains("sidebar-closed")) return false;
  return isSidebarWide();
}

function syncSidebarToggle() {
  const open = isSidebarOpen();
  dom.sidebarToggle.setAttribute("aria-expanded", String(open));
  dom.sidebarBackdrop.hidden = !(open && !isSidebarWide());
}

function openSidebar() {
  document.body.classList.add("sidebar-open");
  document.body.classList.remove("sidebar-closed");
  syncSidebarToggle();
  if (!isSidebarWide()) dom.sidebarClose.focus();
}

function closeSidebar({ returnFocus = true } = {}) {
  const overlay = !isSidebarWide();
  document.body.classList.add("sidebar-closed");
  document.body.classList.remove("sidebar-open");
  syncSidebarToggle();
  if (overlay && returnFocus) dom.sidebarToggle.focus();
}

function toggleSidebar() {
  if (isSidebarOpen()) closeSidebar();
  else openSidebar();
}

// Escape closes an open overlay; Tab and Shift+Tab stay inside it.
function handleSidebarKeydown(event) {
  if (!isSidebarOpen() || isSidebarWide()) return;
  if (event.key === "Escape") {
    event.preventDefault();
    closeSidebar();
    return;
  }
  if (event.key !== "Tab") return;
  const focusable = Array.from(dom.appSidebar.querySelectorAll("button, a[href]")).filter((element) => element.offsetParent !== null);
  if (!focusable.length) return;
  const first = focusable[0];
  const last = focusable[focusable.length - 1];
  if (event.shiftKey && document.activeElement === first) {
    event.preventDefault();
    last.focus();
  } else if (!event.shiftKey && document.activeElement === last) {
    event.preventDefault();
    first.focus();
  }
}

function toggleSidebarFields() {
  const expanded = dom.sidebarFieldsToggle.getAttribute("aria-expanded") !== "false";
  dom.sidebarFieldsToggle.setAttribute("aria-expanded", String(!expanded));
  dom.sidebarFieldList.classList.toggle("hidden", expanded);
}

// The format-consistency summaries: on import, field configuration and placeholder
// confirmations, never on a filter.
function refreshFormatConsistency() {
  state.dateStyles = buildDateStyleSummary(state.rows, state.profiles, (row) => state.sourceDateCells.get(row), getSourceRowNumber, state.formatScans);
  state.constantFields = buildConstantFieldSummary(state.profiles);
  state.lookAlikes = buildLookAlikeSummary(state.rows, state.profiles, state.formatScans);
}

// headerDetection is { headerRowNumber, skippedRows: [{ rowNumber, cells }] } from the
// importer when title rows were skipped above the header, otherwise null.
// rowNumbers[i] is the file row of rows[i]; null when the import could not tell.
// formulaErrors[i] is the { field: error text } of rows[i] or null; only Excel imports pass it.
// dateCells[i] is the { field: true } of the date cells of rows[i] or null; only Excel imports pass it.
function commitTabularData(fields, rows, sourceName, importId, parseWarnings = [], sheetName = "", numberFormatHints = null, headerDetection = null, rowNumbers = null, formulaErrors = null, dateCells = null) {
  if (!isCurrentImport(importId)) return;
  if (!fields.length || !rows.length) {
    showError(() => L("文件中没有可分析的数据。请确认数据包含表头和至少一行数据。", "The file has no data to analyze. Make sure it has a header row and at least one row of data."), importId);
    return;
  }
  if (fields.length > IMPORT_LIMITS.maxColumns || rows.length > IMPORT_LIMITS.maxRows) {
    showError(() => L(`数据超过 ${formatInteger(IMPORT_LIMITS.maxRows)} 行或 ${formatInteger(IMPORT_LIMITS.maxColumns)} 列限制，请拆分后重试。`, `The data is over the limit of ${formatInteger(IMPORT_LIMITS.maxRows)} rows or ${formatInteger(IMPORT_LIMITS.maxColumns)} columns; split it and try again.`), importId);
    return;
  }

  state.rows = rows;
  state.excludedRows = { suspectedTotal: [], duplicateRows: [] };
  state.exclusionSourceRows = rows;
  state.includedRows = rows;
  state.repeatedRowIndexes = null;
  state.repeatedRowSource = null;
  state.restoredTotalRows = new Set();
  state.lastExclusion = null;
  setExclusionMessage("suspectedTotal", "");
  setExclusionMessage("duplicateRows", "");
  state.filteredRows = rows.slice();
  state.filterSourceRows = state.rows;
  state.fields = fields;
  state.parseWarnings = parseWarnings;
  state.pendingExcel = null;
  state.sourceFileName = sourceName;
  state.sourceSheetName = sheetName;
  state.sourceType = sheetName ? "excel" : "csv";
  state.analysisCompletedAt = new Date();
  state.chartInsights = { category: "", trend: "" };
  state.trendInsightData = null;
  state.customAnalysis = null;
  state.filters = createEmptyFilterConfig();
  state.categoryFilterOptions = [];
  state.categoryFilterDraftValues = new Set();
  state.previewPage = 1;
  state.problemView = null;
  state.problemViewOpener = null;
  state.schemaFilter = { query: "", kind: "all" };
  // Source-file hints are replaced on every commit; only Excel imports pass them.
  state.numberFormatHints = Object.assign(Object.create(null), numberFormatHints);
  state.headerDetection = buildHeaderDetectionState(headerDetection);
  state.sourceRowNumbers = new WeakMap();
  if (Array.isArray(rowNumbers) && rowNumbers.length === rows.length) {
    rows.forEach((row, index) => state.sourceRowNumbers.set(row, rowNumbers[index]));
  }
  state.sourceFormulaErrors = new WeakMap();
  const hasFormulaErrors = Array.isArray(formulaErrors) && formulaErrors.length === rows.length;
  if (hasFormulaErrors) {
    rows.forEach((row, index) => {
      if (formulaErrors[index]) state.sourceFormulaErrors.set(row, formulaErrors[index]);
    });
  }
  state.formatScans = createFormatScans();
  state.sourceDateCells = new WeakMap();
  if (Array.isArray(dateCells) && dateCells.length === rows.length) {
    rows.forEach((row, index) => {
      if (dateCells[index]) state.sourceDateCells.set(row, dateCells[index]);
    });
  }
  state.placeholderConfirmations = Object.create(null);
  state.manualFieldTypes = Object.create(null);
  state.profiles = fields.map((field) => buildColumnProfile(field, rows, state.numberFormatHints[field]));
  state.fieldTypeDrafts = Object.fromEntries(
    state.profiles.map((profile) => [profile.field, profile.inferredTypeKey])
  );
  const duplicateRowGroups = [];
  state.duplicateRows = countDuplicateRows(rows, fields, duplicateRowGroups);
  state.duplicateRowGroups = duplicateRowGroups;
  state.idDuplicates = buildIdDuplicateSummary(rows, fields, state.profiles, getSourceRowNumber, state.placeholderConfirmations);
  state.placeholders = buildPlaceholderSummary(rows, fields, state.profiles, state.placeholderConfirmations, getSourceRowNumber);
  state.numericLooking = buildNumericLookingSummary(state.profiles);
  state.totalLabelRows = findTotalLabelRows(rows, fields);
  const totalCheck = checkSuspectedTotals(rows, fields, state.profiles, getSourceRowNumber, state.totalLabelRows);
  state.suspectedTotals = totalCheck.totals;
  state.suspectedTotalLabelCount = totalCheck.labelRowCount;
  state.formulaErrors = hasFormulaErrors ? buildFormulaErrorSummary(rows, fields, (row) => state.sourceFormulaErrors.get(row), getSourceRowNumber) : [];
  refreshFormatConsistency();
  state.totalMissing = state.profiles.reduce((sum, profile) => sum + profile.missingCount, 0);
  // No filter yet: the analysis rows are the original rows, so one calculation serves both.
  // The two share one array; nothing may modify it in place.
  state.qualityNumericStats = buildNumericStats(state.rows);
  state.numericStats = state.qualityNumericStats;
  state.categoryStats = buildCategoryStats(state.filteredRows);
  state.checkupSummary = buildCheckupSummary();

  const sourceDisplayName = sheetName ? `${sourceName} · ${sheetName}` : sourceName;
  // The sidebar's file name (v3.0.0). report.js names the report after state.sourceFileName and
  // reads this title only when that is empty. v3.2.1: a file with more than one worksheet shows the
  // file alone here; the worksheet has its own line below (renderSheetSwitches()).
  const sidebarSourceName = state.keptExcel ? sourceName : sourceDisplayName;
  dom.sourceName.textContent = sidebarSourceName;
  dom.sourceName.title = sidebarSourceName;
  dom.topbarFileName.textContent = sourceDisplayName;
  dom.topbarFileName.title = sourceDisplayName;
  dom.sheetSelectionPanel.classList.add("hidden");
  dom.results.classList.remove("hidden");
  // v3.2.0: the figures of the import message, not its sentence (buildImportStatus()). Warnings
  // without a line count in the total only.
  state.importStatus = {
    sourceDisplayName,
    warningCount: state.parseWarnings.length,
    affectedRows: new Set(state.parseWarnings.map((warning) => warning.row).filter((row) => row !== null)).size,
    headerDetection: state.headerDetection,
    suspectedTotals: state.suspectedTotals,
    formulaErrors: state.formulaErrors
  };
  showImportStatus(state.importStatus);
  state.returnModule = null;

  renderOverview();
  renderPreview();
  renderSchema();
  renderSidebar();
  initializeFilterControls();
  renderQuality();
  renderNumericStats();
  renderCategoryStats();
  populateControls();
  renderCharts();
  renderChartInsights();
  resetV2AnalysisState();
  clearStandardizedExport();
  updateExportButtons();
  showResultsWorkspace();
}

// state.headerDetection from the importer's { headerRowNumber, skippedRows: [{ rowNumber,
// cells }] }, or null: each skipped row keeps a short preview of its cells.
function buildHeaderDetectionState(headerDetection) {
  return headerDetection
    ? {
      headerRowNumber: headerDetection.headerRowNumber,
      skippedRows: headerDetection.skippedRows.map((row) => ({
        rowNumber: row.rowNumber,
        preview: formatSkippedRowPreview(row.cells)
      }))
    }
    : null;
}

// File row of a row object from the current import, or null for rows that did not
// come through commitTabularData() or an import without row numbers.
function getSourceRowNumber(row) {
  return state.sourceRowNumbers.get(row) ?? null;
}

// Drops the last 检查标准化改动 (v2.10.0) and its panel: imports, field configuration,
// placeholder confirmations and resets call it; filters never do. The buttons follow
// updateExportButtons().
function clearStandardizedExport() {
  state.standardizedExport = null;
  dom.standardizedExportPanel.classList.add("hidden");
  dom.standardizedExportSummary.textContent = "";
  dom.standardizedExportTable.innerHTML = "";
  setStandardizedExportStatus("");
}

const STANDARDIZE_COUNT_KEYS = Object.freeze(["date", "number", "placeholder", "trim", "kept"]);

// The check table: one row per field with a change, a kept cell, a note or a year reading;
// ignored fields only in the sentence above it; then the total.
function renderStandardizedExport() {
  const result = state.standardizedExport;
  dom.standardizedExportPanel.classList.toggle("hidden", !result);
  if (!result) return;
  const listed = result.counts.filter((count) => count.kind !== "ignored"
    && (count.kind === "year" || count.orderHint || STANDARDIZE_COUNT_KEYS.some((key) => count[key])));
  const ignoredCount = result.counts.filter((count) => count.kind === "ignored").length;
  const { totals } = result;
  dom.standardizedExportSummary.textContent = formatStandardizedSummary({
    changed: totals.date + totals.number + totals.placeholder + totals.trim,
    kept: totals.kept,
    quietCount: result.counts.length - listed.length - ignoredCount,
    ignoredCount
  });
  const cell = (value) => `<td>${value ? formatInteger(value) : "—"}</td>`;
  const rows = listed.map((count) => {
    const notes = formatStandardizeFieldNotes(count).map((note) => `<span class="standardize-note">${escapeHtml(note)}</span>`).join("");
    return `<tr><td>${escapeHtml(count.field)}${notes}</td><td>${escapeHtml(Object.prototype.hasOwnProperty.call(FIELD_TYPE_LABELS, count.typeKey) ? fieldTypeLabel(count.typeKey) : "")}</td>${cell(count.date)}${cell(count.number)}${cell(count.placeholder)}${cell(count.trim)}<td>${escapeHtml(formatStandardizeKept(count))}</td></tr>`;
  }).join("");
  const total = `<tr class="standardize-total"><td>${L("合计", "Total")}</td><td></td>${STANDARDIZE_COUNT_KEYS.map((key) => cell(totals[key])).join("")}</tr>`;
  dom.standardizedExportTable.innerHTML = listed.length
    ? `<thead><tr>${L(["字段", "类型", "日期改写", "数字改写", "清空空值写法", "去首尾空格", "保持原值"], ["Column", "Type", "Dates rewritten", "Numbers rewritten", "Blank-like cleared", "Spaces trimmed", "Kept as is"]).map((title) => `<th>${title}</th>`).join("")}</tr></thead><tbody>${rows}${total}</tbody>`
    : "";
}

function renderOverview() {
  const analysisRows = getAnalysisRows();
  const activeProfiles = getActiveProfiles();
  const numericCount = getProfilesByType("numeric").length;
  const categoryCount = getProfilesByType("category").length;
  const activeKeys = activeProfiles.map((profile) => getConfirmedPlaceholderKeys(profile.field));
  const activeMissing = analysisRows.reduce((total, row) => (
    total + activeProfiles.filter((profile, index) => isMissingOrConfirmed(row[profile.field], activeKeys[index])).length
  ), 0);
  // v3.1.0: rows left out narrow the analysis too.
  const narrowed = analysisRows.length !== state.rows.length;
  const rowValue = narrowed
    ? `${formatInteger(state.rows.length)} → ${formatInteger(analysisRows.length)}`
    : state.rows.length;
  const cards = [
    [narrowed ? L("总记录 → 当前分析", "All rows → In analysis") : L("总记录数", "Total rows"), rowValue],
    [L("字段数", "Columns"), state.fields.length],
    [L("数值字段数", "Number columns"), numericCount],
    [L("分类字段数", "Text columns"), categoryCount],
    [L("空值概览", "Blanks"), activeMissing]
  ];

  // v3.0.0: one line of text, same labels and numbers as the old cards. v3.4.0: hidden under the
  // checkup (index.html), still written; the next version removes it.
  dom.overviewCards.innerHTML = `<p class="metric-line">${cards.map(([label, value]) => (
    `<span class="metric-item"><span class="metric-label">${escapeHtml(label)}</span> <strong class="metric-value">${typeof value === "number" ? formatInteger(value) : escapeHtml(value)}</strong></span>`
  )).join("")}</p>`;
  dom.previewSize.textContent = formatPreviewSize(state.rows.length, state.fields.length);
  renderCheckup();
}

// v3.4.0: the checkup's big number, its bar and the legend, drawn from state.checkupSummary (made on
// commit, rebuild and rows left out; a filter only draws it again). The bar has one segment per kind
// of the rows of that kind alone, one for the rows of two or more kinds, then the rows with no
// problem; each at least 6 px wide (style.css), so widths follow the rows only roughly.
function renderCheckup() {
  const summary = state.checkupSummary;
  dom.checkupPanel.classList.toggle("hidden", !summary);
  if (!summary) {
    dom.checkupTotal.innerHTML = "";
    dom.checkupBar.innerHTML = "";
    dom.checkupLegend.innerHTML = "";
    return;
  }
  dom.checkupTotal.innerHTML = summary.rowCount
    ? `<span class="checkup-total-number">${formatInteger(summary.rowCount)}</span> <span class="checkup-total-label">${escapeHtml(formatCheckupTotalLabel(summary.rowCount))}</span>`
    : `<span class="checkup-total-label">${escapeHtml(formatCheckupTotalLabel(0))}</span>`;
  dom.checkupBar.classList.toggle("hidden", !summary.rowCount);
  dom.checkupBar.innerHTML = summary.rowCount
    ? summary.segments.filter((segment) => segment.rows > 0).map((segment) => {
      const label = escapeHtml(formatCheckupSegment(segment.kind, segment.rows));
      return `<span class="checkup-segment" data-kind="${segment.kind}" role="img" title="${label}" aria-label="${label}" style="flex-grow: ${segment.rows}"></span>`;
    }).join("")
    : "";
  dom.checkupLegend.innerHTML = formatCheckupLegend(summary).map(([label, rows, multi]) => (
    `<span class="checkup-legend-item${multi ? " is-multi" : ""}">${escapeHtml(label)} <b>${escapeHtml(formatCheckupRows(rows))}</b></span>`
  )).join("");
}

// v3.0.0: a quality subsection's note shows its first line; 展开 / 收起 shows the rest. The text
// is all in the page either way.
function toggleSectionNote(event) {
  const button = event.target.closest(".note-toggle");
  if (!button) return;
  const expand = button.getAttribute("aria-expanded") !== "true";
  button.previousElementSibling.classList.toggle("is-expanded", expand);
  button.setAttribute("aria-expanded", String(expand));
  button.textContent = expand ? L("收起", "Collapse") : L("展开", "Expand");
}

// v3.0.0 dirty data: table cells are capped and cut (style.css); hovering one that is cut shows
// its full text. Added on hover, so the tables' HTML stays as it was.
function showTruncatedCellTitle(event) {
  const cell = event.target.closest(".table-wrap td, .table-wrap th");
  if (!cell || cell.title || cell.scrollWidth <= cell.clientWidth) return;
  cell.title = cell.textContent;
}

// A category field's values are concentrated enough to summarise when there are at most
// INSIGHT_CATEGORY_MAX_DISTINCT distinct values in the current rows and the top category
// repeats. The category conclusion and the category chart's default field both use it.
function isConcentratedCategory(item) {
  return Boolean(item.top.length && item.distinctCount <= INSIGHT_CATEGORY_MAX_DISTINCT && item.top[0].count >= 2);
}

// The first category field, in field order, that isConcentratedCategory().
function findConcentratedCategory(categoryStats) {
  return categoryStats.find(isConcentratedCategory);
}

// The sentence under the category chart for the chart's selected field, "" when there is
// nothing to say (no field, no rows). The field is looked up in state.categoryStats, which
// is built from the current analysis rows.
function buildCategoryInsight(field) {
  if (!field || !getAnalysisRows().length) return "";
  const item = state.categoryStats.find((entry) => entry.field === field);
  if (!item) return "";
  if (isConcentratedCategory(item)) {
    const top = item.top[0];
    return L(`分类字段「${field}」中占比最高的类别是「${top.name}」，出现 ${formatInteger(top.count)} 次，占比 ${formatPercent(top.ratio)}。`, `In “${field}”, the most common value is “${top.name}”: ${enCount(top.count, "time", "times")}, ${formatPercent(top.ratio)}.`);
  }
  if (!item.top.length) return L(`「${field}」当前没有可统计的取值，不做概括。`, `“${field}” has no values to count right now, so there is no summary.`);
  return L(`「${field}」取值比较分散（超过 ${INSIGHT_CATEGORY_MAX_DISTINCT} 个不同值，或没有重复出现的值），不做概括。`, `“${field}” is too scattered to summarize (more than ${INSIGHT_CATEGORY_MAX_DISTINCT} distinct values, or no value appears twice).`);
}

function renderCategoryInsight() {
  state.chartInsights.category = buildCategoryInsight(dom.categoryChartField.value);
  setChartInsight(dom.categoryChartInsight, state.chartInsights.category);
}

function renderTrendInsight() {
  const field = dom.dateChartField.value;
  // v3.2.0: the figures are kept until the next refresh; a language switch only writes the
  // sentence again (formatTrendInsight()).
  state.trendInsightData = field && getAnalysisRows().length ? computeTrendInsight(field) : null;
  state.chartInsights.trend = state.trendInsightData ? formatTrendInsight(state.trendInsightData) : "";
  setChartInsight(dom.dateChartInsight, state.chartInsights.trend);
}

// Both conclusions follow the chart selects, so every path that rebuilds the selects or
// the analysis rows calls this after populateControls() / renderCharts(); a change of one
// of the two selects only redoes its own sentence (script.js).
function renderChartInsights() {
  renderCategoryInsight();
  renderTrendInsight();
}

function setChartInsight(element, text) {
  element.textContent = text;
  element.classList.toggle("hidden", !text);
}

const ROW_TABLE_PAGE_SIZE = 50;

// One page of rowCount rows: page clamped to 1..pageCount (anything else reads as 1),
// and the slice bounds of that page.
function getRowPage(rowCount, page) {
  const pageCount = Math.max(1, Math.ceil(rowCount / ROW_TABLE_PAGE_SIZE));
  const wanted = Math.floor(Number(page));
  const current = Number.isFinite(wanted) ? Math.min(Math.max(wanted, 1), pageCount) : 1;
  return { page: current, pageCount, start: (current - 1) * ROW_TABLE_PAGE_SIZE, end: current * ROW_TABLE_PAGE_SIZE };
}

// The page a pager button leads to: first, prev, next or last.
function movePage(action, page, pageCount) {
  if (action === "first") return 1;
  if (action === "prev") return Math.max(1, page - 1);
  if (action === "next") return Math.min(pageCount, page + 1);
  if (action === "last") return pageCount;
  return page;
}

// Pager buttons: first and previous disabled on the first page, next and last on the
// last one. buttons: [first, prev, next, last].
function setPagerButtons(buttons, page, pageCount) {
  buttons.forEach((button, index) => {
    button.disabled = index < 2 ? page <= 1 : page >= pageCount;
  });
}

// Table HTML of one page of rows: the file row ("—" when the row has none), an
// optional note column, then every field as displayed and escaped. entries:
// [{ row, line, note, highlightField }]; a null row (a parse warning without an
// imported row) leaves the field cells empty, and the highlightField cell gets
// class "cell-problem". Pure: no DOM, no state.
function buildRowTableHtml(fields, entries, { withNote = false } = {}) {
  const cell = (entry, field) => {
    if (!entry.row) return "<td></td>";
    const value = escapeHtml(displayValue(entry.row[field]));
    return entry.highlightField !== undefined && entry.highlightField !== null && field === entry.highlightField
      ? `<td class="cell-problem">${value}</td>`
      : `<td>${value}</td>`;
  };
  return `
    <thead>
      <tr><th>${L("文件行号", "File row")}</th>${withNote ? `<th>${L("说明", "Note")}</th>` : ""}${fields.map((field) => `<th>${escapeHtml(field)}</th>`).join("")}</tr>
    </thead>
    <tbody>
      ${entries.map((entry) => `<tr><td>${entry.line === null || entry.line === undefined ? "—" : escapeHtml(entry.line)}</td>${withNote ? `<td>${escapeHtml(entry.note || "")}</td>` : ""}${fields.map((field) => cell(entry, field)).join("")}</tr>`).join("")}
    </tbody>
  `;
}

// One page of the analysis rows (they follow the filters), state.previewPage clamped.
function renderPreview() {
  const rows = getAnalysisRows();
  const { page, pageCount, start, end } = getRowPage(rows.length, state.previewPage);
  state.previewPage = page;
  dom.previewPageInfo.textContent = formatPreviewPageInfo(page, pageCount, rows.length, countExcludedRows(), hasActiveFilters());
  setPagerButtons([dom.previewFirstPage, dom.previewPrevPage, dom.previewNextPage, dom.previewLastPage], page, pageCount);
  if (!rows.length) {
    dom.previewTable.classList.add("is-empty-table");
    dom.previewTable.innerHTML = emptyTable(L("当前筛选条件下没有可预览的数据。", "No rows to preview under the current filters."));
    return;
  }
  dom.previewTable.classList.remove("is-empty-table");
  dom.previewTable.innerHTML = buildRowTableHtml(state.fields, rows.slice(start, end).map((row) => ({ row, line: getSourceRowNumber(row) })));
}

function goToPreviewPage(action) {
  const { page, pageCount } = getRowPage(getAnalysisRows().length, state.previewPage);
  state.previewPage = movePage(action, page, pageCount);
  renderPreview();
}

// v3.0.0 F1: the numbers of the old 查看详情 are columns now. A select whose draft differs
// from the applied type carries is-changed; 认不出 follows the draft (updateSchemaDraftRow()).
function renderSchema() {
  dom.schemaTable.innerHTML = `
    <thead>
      <tr>
        <th>${L("字段名", "Column")}</th>
        <th>${L("自动类型", "Detected type")}</th>
        <th>${L("确认类型", "Type")}</th>
        <th class="num">${L("非空值", "Non-blank")}</th>
        <th class="num">${L("唯一值", "Distinct")}</th>
        <th class="num">${L("认不出", "Unrecognized")}</th>
        <th>${L("示例值", "Examples")}</th>
      </tr>
    </thead>
    <tbody>
      ${state.profiles.map((profile) => {
        const draftTypeKey = getDraftTypeKey(profile);
        const failureCount = countFieldConversionFailures(profile.field, draftTypeKey);
        const requiresConversion = draftTypeKey === "numeric" || draftTypeKey === "date";
        const exampleText = profile.sampleValues.join(Lpick(LIST_SEPARATOR)) || "—";
        const failureLabel = formatSchemaFailureLabel(profile.field, failureCount, requiresConversion);
        return `
          <tr>
            <td><span class="field-name" title="${escapeHtml(profile.field)}">${escapeHtml(profile.field)}</span></td>
            <td><span class="type-badge ${profile.inferredTypeKey}">${escapeHtml(fieldTypeLabel(profile.inferredTypeKey))}</span></td>
            <td>
              <select class="schema-type-select${draftTypeKey !== profile.typeKey ? " is-changed" : ""}" data-field-type="${escapeHtml(profile.field)}" aria-label="${escapeHtml(L(`调整 ${profile.field} 的字段类型`, `Change the type of ${profile.field}`))}">
                ${Object.entries(FIELD_TYPE_LABELS).map(([typeKey, label]) => `
                  <option value="${typeKey}"${draftTypeKey === typeKey ? " selected" : ""}>${escapeHtml(Lpick(label))}</option>
                `).join("")}
              </select>
            </td>
            <td class="num">${formatInteger(profile.nonMissingCount)}</td>
            <td class="num">${formatInteger(profile.uniqueCount)}</td>
            <td class="num"><span class="conversion-failure-count${failureCount ? " warning" : ""}" title="${formatSchemaFailureTitle(requiresConversion)}" aria-label="${escapeHtml(failureLabel)}">${requiresConversion ? formatInteger(failureCount) : "—"}</span></td>
            <td><span class="schema-example" title="${escapeHtml(exampleText)}">${escapeHtml(exampleText)}</span></td>
          </tr>
        `;
      }).join("")}
    </tbody>
  `;
  renderSchemaFilter();
  updateFieldConfigActions();
  if (!dom.fieldConfigMessage.textContent) {
    setFieldConfigMessage(() => L("自动识别已完成。可修改类型后统一应用。", "Types detected. You can change them and apply them all at once."), "");
  }
}

// v3.0.0 F2: with more than 20 fields, a search by name and filters by the selects' current
// types. They only hide rows (tr.hidden): the selects stay in the page, so the one in use keeps
// the focus, and applying, restoring and the draft check still cover every field.
const SCHEMA_FILTER_MIN_FIELDS = 21;
const SCHEMA_KINDS = Object.freeze([
  ["all", Lp("全部", "All")], ["numeric", Lp("数值", "Number")], ["category", Lp("分类", "Text")], ["date", Lp("日期", "Date")], ["id", Lp("ID", "ID")], ["ignore", Lp("忽略", "Ignored")],
  ["failing", Lp("有认不出的", "With unrecognized")], ["changed", Lp("改过的", "Changed")]
]);

function isSchemaFilterOn() {
  return state.profiles.length >= SCHEMA_FILTER_MIN_FIELDS;
}

function isSchemaKind(profile, kind) {
  const draftTypeKey = getDraftTypeKey(profile);
  if (kind === "all") return true;
  if (kind === "failing") return countFieldConversionFailures(profile.field, draftTypeKey) > 0;
  if (kind === "changed") return draftTypeKey !== profile.typeKey;
  return draftTypeKey === kind;
}

function countSchemaKinds() {
  return Object.fromEntries(SCHEMA_KINDS.map(([kind]) => [kind, state.profiles.filter((profile) => isSchemaKind(profile, kind)).length]));
}

// Field names the table shows: name contains the search (edge spaces trimmed, case ignored)
// and the field is of the chosen type.
function listSchemaFieldsShown() {
  const query = state.schemaFilter.query.trim().toLocaleLowerCase();
  return state.profiles
    .filter((profile) => (!query || profile.field.toLocaleLowerCase().includes(query)) && isSchemaKind(profile, state.schemaFilter.kind))
    .map((profile) => profile.field);
}

// The type buttons with their counts. A count of 0 hides its button, except 全部 and the chosen
// one: that stays until the user picks another type or searches again.
function renderSchemaKinds() {
  const counts = countSchemaKinds();
  dom.schemaKinds.innerHTML = SCHEMA_KINDS
    .filter(([kind]) => kind === "all" || counts[kind] > 0 || kind === state.schemaFilter.kind)
    .map(([kind, label]) => `<button type="button" class="schema-kind" data-schema-kind="${kind}" aria-pressed="${kind === state.schemaFilter.kind}">${escapeHtml(Lpick(label))} <span class="schema-kind-count">${formatInteger(counts[kind])}</span></button>`)
    .join("");
}

function applySchemaFilter() {
  const on = isSchemaFilterOn();
  const shown = new Set(on ? listSchemaFieldsShown() : state.profiles.map((profile) => profile.field));
  dom.schemaTable.querySelectorAll("select[data-field-type]").forEach((select) => {
    const row = select.closest("tr");
    if (row) row.hidden = !shown.has(select.dataset.fieldType);
  });
  dom.schemaFilterEmpty.classList.toggle("hidden", !on || shown.size > 0);
}

// After each renderSchema(): show or hide the toolbar, fall back to 全部 when the chosen type has
// no field left (after applying), and filter the new rows.
function renderSchemaFilter() {
  const on = isSchemaFilterOn();
  dom.schemaFilter.classList.toggle("hidden", !on);
  if (!on) {
    dom.schemaKinds.innerHTML = "";
    applySchemaFilter();
    return;
  }
  if (state.schemaFilter.kind !== "all" && countSchemaKinds()[state.schemaFilter.kind] === 0) state.schemaFilter.kind = "all";
  dom.schemaSearch.placeholder = formatSchemaSearchPlaceholder(state.profiles.length);
  dom.schemaSearch.value = state.schemaFilter.query;
  renderSchemaKinds();
  applySchemaFilter();
}

function handleSchemaSearch() {
  state.schemaFilter.query = dom.schemaSearch.value;
  if (state.schemaFilter.kind !== "all" && countSchemaKinds()[state.schemaFilter.kind] === 0) state.schemaFilter.kind = "all";
  renderSchemaKinds();
  applySchemaFilter();
}

function selectSchemaKind(kind, { focus = false } = {}) {
  if (!SCHEMA_KINDS.some(([key]) => key === kind)) return;
  state.schemaFilter.kind = kind;
  renderSchemaKinds();
  applySchemaFilter();
  // The buttons were drawn again: give the focus to the new one.
  if (focus) dom.schemaKinds.querySelector?.(`[data-schema-kind="${kind}"]`)?.focus();
}

function handleSchemaKindClick(event) {
  const button = event.target.closest("[data-schema-kind]");
  if (button) selectSchemaKind(button.dataset.schemaKind, { focus: true });
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
  // F2: the counts follow the selects, but the rows are not filtered again, so the select in
  // use stays on screen with its focus.
  if (isSchemaFilterOn()) renderSchemaKinds();
  updateFieldConfigActions();
  if (!hasFieldConfigDraftChanges()) {
    setFieldConfigMessage(() => L("当前选择与已应用的字段配置一致。", "The selection matches the applied column settings."), "");
    return;
  }
  const totalFailures = state.profiles.reduce((sum, item) => (
    sum + countFieldConversionFailures(item.field, getDraftTypeKey(item))
  ), 0);
  setFieldConfigMessage(() => (totalFailures
    ? L(
      `字段配置有未应用的修改；当前配置预计有 ${formatInteger(totalFailures)} 个非空值认不出，应用后会从对应统计中跳过。`,
      `Column settings have changes that are not applied; with these settings ${enCount(totalFailures, "non-blank value", "non-blank values")} will not be recognized and will be left out of the stats once applied.`
    )
    : L("字段配置有未应用的修改。", "Column settings have changes that are not applied.")), "warning");
}

// The 认不出 cell of the field table: its aria-label and its title (renderSchema(), updateSchemaDraftRow()).
function formatSchemaFailureLabel(field, failureCount, requiresConversion) {
  return requiresConversion
    ? L(`${field} 认不出 ${formatInteger(failureCount)} 个`, `${field}: ${enCount(failureCount, "value", "values")} not recognized`)
    : L(`${field} 当前类型无需转换`, `${field}: no conversion for this type`);
}

function formatSchemaFailureTitle(requiresConversion) {
  return requiresConversion ? L("非空值中认不出的数量", "Non-blank values not recognized") : L("该类型不需要转换", "This type needs no conversion");
}

function updateSchemaDraftRow(field, typeKey) {
  const select = Array.from(dom.schemaTable.querySelectorAll("select[data-field-type]"))
    .find((item) => item.dataset.fieldType === field);
  if (!select) return;
  select.value = typeKey;
  const profile = state.profiles.find((item) => item.field === field);
  select.classList.toggle("is-changed", Boolean(profile) && typeKey !== profile.typeKey);

  const failureElement = select.closest("tr")?.querySelector(".conversion-failure-count");
  if (!failureElement) return;
  const failureCount = countFieldConversionFailures(field, typeKey);
  const requiresConversion = typeKey === "numeric" || typeKey === "date";
  failureElement.textContent = requiresConversion ? formatInteger(failureCount) : "—";
  failureElement.classList.toggle("warning", failureCount > 0);
  failureElement.title = formatSchemaFailureTitle(requiresConversion);
  failureElement.setAttribute("aria-label", formatSchemaFailureLabel(field, failureCount, requiresConversion));
}

function applyFieldConfiguration() {
  if (!state.rows.length) return;
  // Only fields whose type changes now are judged again; the others keep their
  // record, which a placeholder confirmation may have made equal to the automatic type.
  const manualFieldTypes = Object.assign(Object.create(null), state.manualFieldTypes);
  state.profiles.forEach((profile) => {
    const typeKey = getDraftTypeKey(profile);
    if (typeKey !== profile.typeKey) {
      if (typeKey !== profile.inferredTypeKey) manualFieldTypes[profile.field] = typeKey;
      else delete manualFieldTypes[profile.field];
    }
    profile.typeKey = typeKey;
    profile.type = FIELD_TYPE_LABELS[typeKey][0];
    profile.conversionFailureCount = countFieldConversionFailures(profile.field, typeKey);
    profile.conversionFailureSamples = getFieldConversionFailureSamples(profile.field, typeKey);
  });
  state.manualFieldTypes = manualFieldTypes;
  state.fieldTypeDrafts = Object.fromEntries(
    state.profiles.map((profile) => [profile.field, profile.typeKey])
  );

  const failureCount = state.profiles.reduce((sum, profile) => sum + profile.conversionFailureCount, 0);
  rebuildAnalysisFromProfiles(() => (failureCount
    ? L(
      `字段配置已应用；${formatInteger(failureCount)} 个非空值认不出，已从对应统计和图表中跳过。`,
      `Column settings applied; ${enCount(failureCount, "non-blank value is", "non-blank values are")} not recognized and left out of the stats and charts.`
    )
    : L("字段配置已应用，统计、图表和分析结论已重新计算。", "Column settings applied; stats, charts and summaries are recalculated.")));
}

function restoreAutomaticFieldTypes() {
  if (!state.rows.length) return;
  state.profiles.forEach((profile) => {
    profile.typeKey = profile.inferredTypeKey;
    profile.type = profile.inferredType;
    profile.conversionFailureCount = countFieldConversionFailures(profile.field, profile.inferredTypeKey);
    profile.conversionFailureSamples = getFieldConversionFailureSamples(profile.field, profile.inferredTypeKey);
  });
  state.fieldTypeDrafts = Object.fromEntries(
    state.profiles.map((profile) => [profile.field, profile.inferredTypeKey])
  );
  state.manualFieldTypes = Object.create(null);
  rebuildAnalysisFromProfiles(() => L("已恢复自动识别，统计、图表和分析结论已重新计算。", "Detected types restored; stats, charts and summaries are recalculated."));
}

function hasFieldConfigDraftChanges() {
  return state.profiles.some((profile) => getDraftTypeKey(profile) !== profile.typeKey);
}

function hasAutomaticTypeDifferences() {
  if (Object.keys(state.manualFieldTypes).length) return true;
  return state.profiles.some((profile) => (
    profile.typeKey !== profile.inferredTypeKey
      || getDraftTypeKey(profile) !== profile.inferredTypeKey
  ));
}

function updateFieldConfigActions() {
  dom.applyFieldConfig.disabled = !hasFieldConfigDraftChanges();
  // v3.0.0: a light blue banner while drafts wait; otherwise a grey line under the table.
  dom.fieldConfigToolbar.classList.toggle("has-draft", hasFieldConfigDraftChanges());
  dom.restoreAutoFieldTypes.disabled = !hasAutomaticTypeDifferences();
  // The problem rows follow the applied configuration; say so while drafts wait.
  dom.problemRowsDraftNote.classList.toggle("hidden", !state.problemView || !hasFieldConfigDraftChanges());
  updateExportButtons();
}

// v3.2.0: a one-off message (the outcome of a click) keeps the function that wrote it, so a
// language switch can write it again in the other language (rewriteOneOffMessages()). message
// is a text or a function returning one; an empty text forgets the element.
const oneOffMessages = new Map();

function writeOneOffMessage(element, className, message, tone = "") {
  const build = typeof message === "function" ? message : () => message;
  element.className = `${className}${tone ? ` ${tone}` : ""}`;
  element.textContent = build();
  if (element.textContent) oneOffMessages.set(element, { build, className, tone });
  else oneOffMessages.delete(element);
}

function rewriteOneOffMessages() {
  oneOffMessages.forEach(({ build, className, tone }, element) => writeOneOffMessage(element, className, build, tone));
}

function setFieldConfigMessage(message, tone = "") {
  writeOneOffMessage(dom.fieldConfigMessage, "field-config-message", message, tone);
}

// actionLabel: what was refused, [中文, English] (an Lp pair), or a text.
function ensureFieldConfigurationApplied(actionLabel) {
  if (!hasFieldConfigDraftChanges()) return true;
  const action = (index) => (Array.isArray(actionLabel) ? actionLabel[index] : actionLabel);
  setFieldConfigMessage(() => L(
    `字段配置有未应用的修改。请先点击“应用字段配置”，再${action(0)}。`,
    `Column settings have changes that are not applied. Click “Apply column settings” first, then ${action(1)}.`
  ), "warning");
  return false;
}

// What ensureFieldConfigurationApplied() refused, in both languages.
const DRAFT_ACTION_PLACEHOLDERS = Lp("处理像是空值的写法", "handle values that may mean blank");
const DRAFT_ACTION_FILTERS = Lp("应用筛选", "apply the filters");

function setPlaceholderMessage(message, tone = "") {
  writeOneOffMessage(dom.placeholderMessage, "placeholder-message", message, tone);
}

// Pending common keys of fields that are not ignored: what "全部确认常见写法" covers.
function getPendingCommonPlaceholders() {
  return state.placeholders.filter((entry) => entry.tier === "common" && !entry.confirmed && !entry.fieldIgnored);
}

// Profile figures if keys were confirmed for field on top of its confirmed keys.
function previewPlaceholderConfirmation(field, keys) {
  const currentKeys = getConfirmedPlaceholderKeys(field);
  const nextKeys = [...currentKeys, ...keys.filter((key) => !currentKeys.includes(key))];
  return estimateConfirmedProfile(getFieldProfile(field), state.rows, currentKeys, nextKeys, state.numberFormatHints[field], state.manualFieldTypes[field]);
}

// Fields whose type would change if every pending common key were confirmed.
function previewAllCommonPlaceholders() {
  const keysByField = new Map();
  getPendingCommonPlaceholders().forEach((entry) => keysByField.set(entry.field, [...(keysByField.get(entry.field) || []), entry.key]));
  return [...keysByField].flatMap(([field, keys]) => {
    const previousTypeKey = getFieldProfile(field).typeKey;
    const { typeKey } = previewPlaceholderConfirmation(field, keys);
    return typeKey === previousTypeKey ? [] : [{ field, typeKey, previousTypeKey }];
  });
}

// v3.1.0: the figures a placeholder message compares, before and after a confirmation or an
// undo: the field's missing count, values it cannot read, type, and whether its ID duplicates
// list the key.
function getPlaceholderFieldFigures(field, key) {
  const profile = getFieldProfile(field);
  return {
    missing: profile.missingCount,
    failures: profile.conversionFailureCount || 0,
    typeKey: profile.typeKey,
    listed: state.idDuplicates.some((item) => item.field === field && item.values.some((value) => placeholderKey(value.value) === key))
  };
}

function confirmPlaceholder(field, key) {
  const entry = state.placeholders.find((item) => item.field === field && item.key === key && !item.confirmed && !item.fieldIgnored);
  const keysByField = Object.create(null);
  if (!entry) return updatePlaceholderConfirmations(keysByField, "");
  keysByField[field] = [...getConfirmedPlaceholderKeys(field), key];
  const before = getPlaceholderFieldFigures(field, key);
  return updatePlaceholderConfirmations(keysByField, () => formatPlaceholderConfirmMessage({
    field, key, count: entry.count, before, after: getPlaceholderFieldFigures(field, key),
    rowCount: getIncludedRows().length, rowsLeftOut: countExcludedRows() > 0
  }));
}

function undoPlaceholderConfirmation(field, key) {
  const keys = getConfirmedPlaceholderKeys(field);
  const keysByField = Object.create(null);
  if (!keys.includes(key)) return updatePlaceholderConfirmations(keysByField, "");
  keysByField[field] = keys.filter((item) => item !== key);
  const entry = state.placeholders.find((item) => item.field === field && item.key === key);
  const before = getPlaceholderFieldFigures(field, key);
  return updatePlaceholderConfirmations(keysByField, () => formatPlaceholderUndoMessage({
    field, key, count: entry ? entry.count : 0, before, after: getPlaceholderFieldFigures(field, key),
    rowCount: getIncludedRows().length, rowsLeftOut: countExcludedRows() > 0,
    ignored: before.typeKey === "ignore"
  }));
}

function confirmAllCommonPlaceholders() {
  const entries = getPendingCommonPlaceholders();
  const keysByField = Object.create(null);
  entries.forEach((entry) => {
    keysByField[entry.field] ??= [...getConfirmedPlaceholderKeys(entry.field)];
    keysByField[entry.field].push(entry.key);
  });
  const typesBefore = Object.fromEntries(entries.map((entry) => [entry.field, getFieldProfile(entry.field).typeKey]));
  return updatePlaceholderConfirmations(keysByField, () => {
    // Fields whose type changed first, as the confirm-all note names them first.
    const changed = (entry) => getFieldProfile(entry.field).typeKey !== typesBefore[entry.field];
    const ordered = [...entries.filter(changed), ...entries.filter((entry) => !changed(entry))];
    const typeChanges = [...new Set(ordered.filter(changed).map((entry) => entry.field))]
      .map((field) => ({ field, typeKey: getFieldProfile(field).typeKey, previousTypeKey: typesBefore[field] }));
    return formatAllPlaceholdersMessage(ordered, typeChanges, getIncludedRows().length, countExcludedRows() > 0);
  });
}

// The fourth refresh path. keysByField replaces the confirmed keys of each listed
// field; refused while the field configuration has unapplied drafts. message is a text, or
// a function that writes it once the analysis is rebuilt (v3.1.0: it compares the counts).
function updatePlaceholderConfirmations(keysByField, message) {
  if (!state.rows.length) return false;
  if (!ensureFieldConfigurationApplied(DRAFT_ACTION_PLACEHOLDERS)) {
    setPlaceholderMessage(() => L(
      "字段配置有未应用的修改。请先点击“应用字段配置”，再处理像是空值的写法。",
      "Column settings have changes that are not applied. Click “Apply column settings” first, then handle values that may mean blank."
    ), "warning");
    return false;
  }
  const fields = Object.keys(keysByField).filter((field) => getFieldProfile(field));
  if (!fields.length) return false;
  const confirmations = Object.assign(Object.create(null), state.placeholderConfirmations);
  fields.forEach((field) => {
    if (keysByField[field].length) confirmations[field] = keysByField[field];
    else delete confirmations[field];
  });
  state.placeholderConfirmations = confirmations;
  fields.forEach(rebuildConfirmedFieldProfile);
  state.totalMissing = state.profiles.reduce((sum, profile) => sum + profile.missingCount, 0);
  rebuildAnalysisFromProfiles(message, setPlaceholderMessage);
  dom.placeholderMessage.focus();
  return true;
}

// Re-profiles a field with its confirmed keys. A manual type stays; any other field
// takes the new automatic type. Its draft follows, so nothing is left unapplied.
function rebuildConfirmedFieldProfile(field) {
  const index = state.profiles.findIndex((profile) => profile.field === field);
  const profile = buildColumnProfile(field, state.rows, state.numberFormatHints[field], getConfirmedPlaceholderKeys(field));
  const manualTypeKey = state.manualFieldTypes[field];
  if (manualTypeKey) {
    profile.typeKey = manualTypeKey;
    profile.type = FIELD_TYPE_LABELS[manualTypeKey][0];
    profile.conversionFailureCount = profile.conversionFailuresByType[manualTypeKey] || 0;
    profile.conversionFailureSamples = [...(profile.conversionFailureSamplesByType[manualTypeKey] || [])];
  }
  state.profiles[index] = profile;
  state.fieldTypeDrafts[field] = profile.typeKey;
}

function createEmptyFilterConfig() {
  return {
    category: null,
    numeric: null,
    date: null
  };
}

// v3.1.0: three layers. state.rows (every row of the file) -> getIncludedRows() (the rows
// counted in the analysis) -> getAnalysisRows() (the filtered rows). Filters start from the
// counted rows. Rows left out of another data set never count (exclusionSourceRows).
function getIncludedRows() {
  return state.exclusionSourceRows === state.rows && Array.isArray(state.includedRows)
    ? state.includedRows
    : state.rows;
}

// Every left-out row of the current rows, both kinds together (a row of both counts once).
function getExcludedIndexSet() {
  if (state.exclusionSourceRows !== state.rows) return new Set();
  return new Set([...state.excludedRows.suspectedTotal, ...state.excludedRows.duplicateRows]);
}

function countExcludedRows() {
  return state.rows.length - getIncludedRows().length;
}

// state.includedRows from state.excludedRows: state.rows itself (the same array) when nothing
// is left out, so the default path keeps every array of v3.0.1.
function rebuildIncludedRows() {
  const excluded = getExcludedIndexSet();
  state.exclusionSourceRows = state.rows;
  state.includedRows = excluded.size ? state.rows.filter((row, index) => !excluded.has(index)) : state.rows;
}

// A filter result counts only while it was made from the current counted rows.
function getAnalysisRows() {
  const included = getIncludedRows();
  return state.filterSourceRows === included && Array.isArray(state.filteredRows)
    ? state.filteredRows
    : included;
}

function hasActiveFilters(filters = state.filters) {
  return Boolean(filters?.category || filters?.numeric || filters?.date);
}

function resetFilterState() {
  state.filteredRows = getIncludedRows().slice();
  state.filterSourceRows = getIncludedRows();
  state.filters = createEmptyFilterConfig();
  state.categoryFilterOptions = [];
  state.categoryFilterDraftValues = new Set();
}

// The first option of the three filter selects (v3.2.0 commit 11: also written on a switch).
const FILTER_NONE_LABELS = Object.freeze({
  category: Lp("不筛选分类字段", "Don’t filter by text"),
  numeric: Lp("不筛选数值字段", "Don’t filter by number"),
  date: Lp("不筛选日期字段", "Don’t filter by date")
});

function initializeFilterControls() {
  const categoryFields = getProfilesByType("category").map((profile) => profile.field);
  const numericFields = getProfilesByType("numeric")
    .filter((profile) => fieldHasNumericData(profile.field, getIncludedRows()))
    .map((profile) => profile.field);
  const dateFields = getProfilesByType("date")
    .filter((profile) => fieldHasDateData(profile.field, getIncludedRows()))
    .map((profile) => profile.field);

  setSelectOptions(dom.categoryFilterField, categoryFields, Lpick(FILTER_NONE_LABELS.category), true);
  setSelectOptions(dom.numericFilterField, numericFields, Lpick(FILTER_NONE_LABELS.numeric), true);
  setSelectOptions(dom.dateFilterField, dateFields, Lpick(FILTER_NONE_LABELS.date), true);
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
  const included = getIncludedRows().length;
  const excluded = countExcludedRows();
  setFilterMessage(() => formatNoFilterMessage(included, excluded), "");
}

function getCategoryFilterOptions(field) {
  if (!field) return [];
  // Confirmed placeholders join the empty value, as in filterRowsByConfig().
  const confirmedKeys = getConfirmedPlaceholderKeys(field);
  const counts = new Map();
  getIncludedRows().forEach((row) => {
    const value = analysisValue(row[field], confirmedKeys);
    counts.set(value, (counts.get(value) || 0) + 1);
  });
  return Array.from(counts.entries())
    .map(([value, count]) => ({
      value,
      count,
      label: value || Lpick(BLANK_GROUP_LABEL)
    }))
    // v3.2.0 (R3): the order is the Chinese one in both languages (the empty group sorts as
    // BLANK_GROUP_KEY, （空值）); only the label shown follows the language.
    .sort((a, b) => (a.value || BLANK_GROUP_KEY).localeCompare(b.value || BLANK_GROUP_KEY, "zh-CN", { numeric: true }));
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
  writeLanguageText(dom.categoryFilterEmpty, () => dom.categoryFilterField.disabled
    ? L("没有可用的分类字段。", "No text columns available.")
    : L("请选择分类字段。", "Choose a text column."));
  renderCategoryFilterValues();
}

function renderCategoryFilterValues() {
  const query = dom.categoryFilterSearch.value.trim().toLocaleLowerCase("zh-CN");
  const visibleOptions = state.categoryFilterOptions.filter((option) => (
    !query || option.label.toLocaleLowerCase("zh-CN").includes(query)
  ));
  const selectedCount = state.categoryFilterDraftValues.size;
  dom.categoryFilterSelectionCount.textContent = dom.categoryFilterField.value
    ? formatFilterSelectionCount(selectedCount)
    : L("未选择字段", "No column chosen");

  if (!visibleOptions.length) {
    dom.categoryFilterValues.innerHTML = `<p class="filter-value-empty">${L("没有匹配的类别。", "No matching values.")}</p>`;
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

// "已选 3 / 6" under the category values.
function formatFilterSelectionCount(selectedCount) {
  const total = formatInteger(state.categoryFilterOptions.length);
  return L(`已选 ${formatInteger(selectedCount)} / ${total}`, `${formatInteger(selectedCount)} / ${total} selected`);
}

// The range hint of a number or date filter, or what to do without one.
function formatFilterRange(from, to) {
  return L(`${from} 至 ${to}`, `${from} to ${to}`);
}

function formatFilterRangeEmpty(disabled) {
  return disabled ? L("没有可用字段", "No columns available") : L("请选择字段", "Choose a column");
}

function updateCategoryFilterDraftSelection(event) {
  const checkbox = event.target.closest("input[data-category-filter-index]");
  if (!checkbox) return;
  const option = state.categoryFilterOptions[Number(checkbox.dataset.categoryFilterIndex)];
  if (!option) return;
  if (checkbox.checked) state.categoryFilterDraftValues.add(option.value);
  else state.categoryFilterDraftValues.delete(option.value);
  dom.categoryFilterSelectionCount.textContent = formatFilterSelectionCount(state.categoryFilterDraftValues.size);
}

function setAllCategoryFilterValues(selected) {
  state.categoryFilterDraftValues = selected
    ? new Set(state.categoryFilterOptions.map((option) => option.value))
    : new Set();
  renderCategoryFilterValues();
}

function updateNumericFilterField() {
  const field = dom.numericFilterField.value;
  const values = field ? getNumericValues(field, getIncludedRows()) : [];
  const enabled = values.length > 0;
  dom.numericFilterMin.disabled = !enabled;
  dom.numericFilterMax.disabled = !enabled;
  if (!enabled) {
    dom.numericFilterMin.value = "";
    dom.numericFilterMax.value = "";
    writeLanguageText(dom.numericFilterRange, () => formatFilterRangeEmpty(dom.numericFilterField.disabled));
    return;
  }
  const { min, max } = numericExtents(values);
  dom.numericFilterMin.value = String(min);
  dom.numericFilterMax.value = String(max);
  writeLanguageText(dom.numericFilterRange, () => formatFilterRange(formatFieldNumber(field, min), formatFieldNumber(field, max)));
}

function updateDateFilterField() {
  const field = dom.dateFilterField.value;
  const dates = field
    ? getIncludedRows().map((row) => parseFieldDate(field, row[field])).filter(Boolean).sort((a, b) => a - b)
    : [];
  const enabled = dates.length > 0;
  dom.dateFilterStart.disabled = !enabled;
  dom.dateFilterEnd.disabled = !enabled;
  if (!enabled) {
    dom.dateFilterStart.value = "";
    dom.dateFilterEnd.value = "";
    writeLanguageText(dom.dateFilterRange, () => formatFilterRangeEmpty(dom.dateFilterField.disabled));
    return;
  }
  const start = dates[0].toISOString().slice(0, 10);
  const end = dates[dates.length - 1].toISOString().slice(0, 10);
  dom.dateFilterStart.value = start;
  dom.dateFilterEnd.value = end;
  writeLanguageText(dom.dateFilterRange, () => formatFilterRange(start, end));
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
      return { error: Lp("请填写数值筛选的最小值和最大值。", "Fill in the minimum and maximum of the number filter."), filters };
    }
    const min = Number(dom.numericFilterMin.value);
    const max = Number(dom.numericFilterMax.value);
    if (!Number.isFinite(min) || !Number.isFinite(max)) {
      return { error: Lp("数值筛选范围必须是有效数字。", "The number filter range must be valid numbers."), filters };
    }
    if (min > max) {
      return { error: Lp("数值筛选的最小值不能大于最大值。", "The minimum of the number filter cannot be greater than the maximum."), filters };
    }
    filters.numeric = { field: numericField, min, max };
  }

  const dateField = dom.dateFilterField.value;
  if (dateField) {
    const start = dom.dateFilterStart.value;
    const end = dom.dateFilterEnd.value;
    if (!start || !end) {
      return { error: Lp("请填写日期筛选的开始日期和结束日期。", "Fill in the start and end dates of the date filter."), filters };
    }
    if (start > end) {
      return { error: Lp("日期筛选的开始日期不能晚于结束日期。", "The start date of the date filter cannot be after the end date."), filters };
    }
    filters.date = { field: dateField, start, end };
  }

  return { error: "", filters };
}

function filterRowsByConfig(rows, filters) {
  const categoryValues = filters.category ? new Set(filters.category.values) : null;
  const categoryKeys = filters.category ? getConfirmedPlaceholderKeys(filters.category.field) : [];
  const startTime = filters.date ? Date.parse(`${filters.date.start}T00:00:00Z`) : null;
  const endTime = filters.date ? Date.parse(`${filters.date.end}T23:59:59.999Z`) : null;
  // The same parser as getNumericValues(), which sets the range in the filter panel.
  const parseNumber = filters.numeric ? getFieldNumberParser(filters.numeric.field) : null;

  return rows.filter((row) => {
    if (filters.category) {
      const value = analysisValue(row[filters.category.field], categoryKeys);
      if (!categoryValues.has(value)) return false;
    }
    if (filters.numeric) {
      const value = parseNumber(row[filters.numeric.field]);
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
  if (!ensureFieldConfigurationApplied(DRAFT_ACTION_FILTERS)) {
    setFilterMessage(() => L("字段配置有未应用的修改，请先应用字段配置。", "Column settings have changes that are not applied; apply them first."), "warning");
    return;
  }
  const { error, filters } = readFilterConfig();
  if (error) {
    setFilterMessage(() => Lpick(error), "warning");
    return;
  }

  state.filters = filters;
  state.filteredRows = filterRowsByConfig(getIncludedRows(), filters);
  state.filterSourceRows = getIncludedRows();
  refreshFilteredAnalysis();

  const activeCount = [filters.category, filters.numeric, filters.date].filter(Boolean).length;
  const filteredOutCount = getIncludedRows().length - state.filteredRows.length;
  if (!activeCount) {
    const included = getIncludedRows().length;
    const excluded = countExcludedRows();
    setFilterMessage(() => L(`未选择筛选字段，当前分析继续使用${formatAllIncludedRows(included, excluded)}。`, `No filter column chosen; the analysis still uses ${formatAllIncludedRows(included, excluded)}.`), "");
  } else {
    const keptCount = state.filteredRows.length;
    setFilterMessage(() => formatFiltersApplied(activeCount, keptCount, filteredOutCount), keptCount ? "success" : "warning");
  }
}

function clearAllFilters() {
  if (!state.rows.length) return;
  resetFilterState();
  initializeFilterControls();
  refreshFilteredAnalysis();
  const included = getIncludedRows().length;
  const excluded = countExcludedRows();
  setFilterMessage(() => L(`已清除全部筛选，当前分析恢复使用${formatAllIncludedRows(included, excluded)}。`, `All filters cleared; the analysis uses ${formatAllIncludedRows(included, excluded)} again.`), "success");
}

function renderFilterCounts() {
  const currentCount = getAnalysisRows().length;
  dom.filterOriginalCount.textContent = formatInteger(state.rows.length);
  dom.filterCurrentCount.textContent = formatInteger(currentCount);
  dom.filterExcludedCount.textContent = formatInteger(Math.max(0, getIncludedRows().length - currentCount));
  const excludedCount = countExcludedRows();
  dom.filterNotCountedItem.classList.toggle("hidden", excludedCount === 0);
  dom.filterNotCountedCount.textContent = formatInteger(excludedCount);
  updateGlobalFilterStatus();
  renderAnalysisContainsNotes();
}

// v3.1.0: the rows of the current analysis (the filtered rows) that are still suspected totals,
// and, with a clean ID column, still extra repeated rows. Without filters or left-out rows the
// repeated rows are state.duplicateRows, so an import builds no list; otherwise the list is
// made once and kept until the next import (listRepeatedRowIndexes()).
function countAnalysisContains() {
  const excluded = getExcludedIndexSet();
  const filtered = hasActiveFilters();
  let analysisRows = null;
  const stillCounted = (indexes) => {
    if (!filtered) return indexes.filter((index) => !excluded.has(index)).length;
    analysisRows = analysisRows || new Set(getAnalysisRows());
    return indexes.filter((index) => analysisRows.has(state.rows[index])).length;
  };
  const totals = state.suspectedTotals.length ? stillCounted(state.suspectedTotals.map((item) => item.index)) : 0;
  let duplicates = 0;
  if (state.duplicateRows && hasCleanIdField()) {
    duplicates = !filtered && !excluded.size ? state.duplicateRows : stillCounted(listRepeatedRowIndexes());
  }
  return { totals, duplicates };
}

function renderAnalysisContainsNotes() {
  const { totals, duplicates } = state.rows.length ? countAnalysisContains() : { totals: 0, duplicates: 0 };
  const text = formatAnalysisContainsNote(totals, duplicates);
  const html = text
    ? `${escapeHtml(text)} <button type="button" class="text-button" data-quality-target="${totals ? "suspectedTotal" : "duplicateRows"}" aria-label="${escapeHtml(formatGoToQualityAriaLabel(totals, duplicates))}">${Lpick(GO_TO_QUALITY_LABEL)}</button>`
    : "";
  [dom.statisticsContainsNote, dom.chartsContainsNote, dom.customContainsNote].forEach((element) => {
    element.innerHTML = html;
    element.classList.toggle("hidden", !text);
  });
}

// 去数据质量处理: the quality module, then its subsection at the top (as 查看 does) and the
// focus on the subsection title; showModule() itself scrolls up and focuses the module title,
// so this comes after it. The repeated rows have no title: their button takes the focus.
function handleAnalysisNoteClick(event) {
  const button = event.target.closest("button[data-quality-target]");
  if (!button) return;
  showModule("qualitySection");
  const totals = button.dataset.qualityTarget === "suspectedTotal";
  (totals ? dom.suspectedTotalSection : dom.duplicateExclusion).scrollIntoView({ block: "start" });
  const target = totals ? dom.suspectedTotalTitle : dom.duplicateExclusionAction.querySelector?.("button[data-exclusion-kind]");
  target?.focus?.({ preventScroll: true });
}

function updateGlobalFilterStatus() {
  if (!dom.globalFilterSummary) return;
  const activeCount = [state.filters.category, state.filters.numeric, state.filters.date].filter(Boolean).length;
  dom.globalFilterSummary.textContent = formatAnalysisScope(activeCount, getAnalysisRows().length, getIncludedRows().length, countExcludedRows());
  dom.globalFilterStatus.classList.toggle("active", activeCount > 0);
  // v3.4.0: on the checkup only while a filter applies or rows are left out.
  dom.globalFilterStatus.classList.toggle("hidden", activeCount === 0 && countExcludedRows() === 0);
  dom.globalClearFilters.classList.toggle("hidden", activeCount === 0);
  // v3.0.0: the same sentence by the statistics, chart and custom-analysis titles (their
  // elements carry no id) and in the sidebar.
  const summary = dom.globalFilterSummary.textContent;
  document.querySelectorAll(".filter-scope").forEach((scope) => scope.classList.toggle("active", activeCount > 0));
  document.querySelectorAll(".filter-scope-text").forEach((text) => { text.textContent = summary; });
  document.querySelectorAll(".filter-scope-clear").forEach((button) => button.classList.toggle("hidden", activeCount === 0));
  dom.sidebarFilterSummary.textContent = summary;
}

function setFilterMessage(message, tone = "") {
  writeOneOffMessage(dom.filterMessage, "filter-message", message, tone);
}

function refreshFilteredAnalysis() {
  const rerunV2Analysis = Boolean(
    state.customAnalysis && !dom.v2AnalysisResults.classList.contains("hidden")
  );
  setDataExportStatus("");
  state.analysisCompletedAt = null;
  updateExportButtons();
  const analysisRows = getAnalysisRows();
  // A filter only removes rows and keeps their order (filterRowsByConfig() is rows.filter()),
  // so as many rows as the original means the same rows in the same order: the full
  // statistics apply. They are then shared with state.qualityNumericStats; nothing may
  // modify them in place.
  state.numericStats = analysisRows.length === state.rows.length
    ? state.qualityNumericStats
    : buildNumericStats(analysisRows);
  state.categoryStats = buildCategoryStats(analysisRows);
  state.previewPage = 1;
  renderOverview();
  renderFilterCounts();
  renderPreview();
  renderNumericStats();
  renderCategoryStats();
  renderCharts();
  renderChartInsights();
  if (rerunV2Analysis) renderV2Analysis();
  else resetV2AnalysisState();
  state.analysisCompletedAt = new Date();
  updateExportButtons();
}

// setMessage shows the outcome: the field configuration area by default, the
// placeholder subsection for placeholder confirmations. fieldConfigMessage is a text or a
// function called once everything is rebuilt. "已清除现有筛选条件。" only when a filter was
// applied before (v3.1.0).
function rebuildAnalysisFromProfiles(fieldConfigMessage = () => L("字段配置已更新。", "Column settings updated."), setMessage = setFieldConfigMessage) {
  const hadFilters = hasActiveFilters();
  state.idDuplicates = buildIdDuplicateSummary(state.rows, state.fields, state.profiles, getSourceRowNumber, state.placeholderConfirmations);
  state.placeholders = buildPlaceholderSummary(state.rows, state.fields, state.profiles, state.placeholderConfirmations, getSourceRowNumber);
  state.numericLooking = buildNumericLookingSummary(state.profiles);
  const totalCheck = checkSuspectedTotals(state.rows, state.fields, state.profiles, getSourceRowNumber, state.totalLabelRows);
  state.suspectedTotals = totalCheck.totals;
  state.suspectedTotalLabelCount = totalCheck.labelRowCount;
  const exclusionChanges = intersectExcludedTotals();
  refreshFormatConsistency();
  clearStandardizedExport();
  state.analysisCompletedAt = null;
  updateExportButtons();
  resetFilterState();
  setDataExportStatus("");
  // The filter was just reset, so the analysis rows are the counted rows: while nothing is left
  // out they are the original rows and one calculation serves both, shared as in
  // commitTabularData(); otherwise the analysis gets its own (v3.1.0).
  state.qualityNumericStats = buildNumericStats(state.rows);
  state.numericStats = getIncludedRows() === state.rows ? state.qualityNumericStats : buildNumericStats(getIncludedRows());
  state.categoryStats = buildCategoryStats(getAnalysisRows());
  state.checkupSummary = buildCheckupSummary();
  state.previewPage = 1;
  renderOverview();
  renderPreview();
  renderSchema();
  renderSidebar();
  initializeFilterControls();
  renderQuality();
  renderNumericStats();
  renderCategoryStats();
  populateControls({ keepChartSelections: true });
  renderCharts();
  renderChartInsights();
  resetV2AnalysisState();
  // The message is built now, in both languages, now that everything is rebuilt: it compares
  // counts taken at this click, which later changes must not touch. A language switch only
  // picks the other text.
  const buildMessage = () => {
    const message = typeof fieldConfigMessage === "function" ? fieldConfigMessage() : fieldConfigMessage;
    const notes = [
      hadFilters ? L("已清除现有筛选条件。", "The filters in use were cleared.") : "",
      formatExclusionRestored(exclusionChanges.restored),
      formatTotalsFlaggedAgain(exclusionChanges.flaggedAgain)
    ].filter(Boolean);
    return [message, ...notes].join(" ");
  };
  setExclusionMessage("suspectedTotal", "");
  setExclusionMessage("duplicateRows", "");
  const built = {};
  const shown = currentLanguage;
  try {
    for (const language of ["zh", "en"]) {
      currentLanguage = language;
      built[language] = buildMessage();
    }
  } finally {
    currentLanguage = shown;
  }
  setMessage(() => built[currentLanguage], "success");
  state.analysisCompletedAt = new Date();
  updateFieldConfigActions();
}

// ---- v3.1.0: rows left out of the analysis ----

// The repeated rows past the first of each group, made when first needed and kept until the
// next import (they depend on the rows and fields only).
function listRepeatedRowIndexes() {
  if (!state.duplicateRows) return [];
  if (state.repeatedRowSource !== state.rows) {
    state.repeatedRowIndexes = findRepeatedRowIndexes(state.rows, state.fields);
    state.repeatedRowSource = state.rows;
  }
  return state.repeatedRowIndexes;
}

// The rows a kind's button leaves out now: every suspected total, or every repeated row.
function listExclusionCandidates(kind) {
  if (kind === "suspectedTotal") return state.suspectedTotals.map((item) => item.index);
  if (kind === "duplicateRows") return listRepeatedRowIndexes();
  return [];
}

// The fifth trigger ("不计入分析" / "撤销"): the filter path, never the rebuild. Records the rows
// marked now; the filters and the custom analysis stay. Refused while drafts wait. Returns
// whether anything changed; state.lastExclusion says what for the messages.
function setRowExclusion(kind, on) {
  if (!state.rows.length || !Object.prototype.hasOwnProperty.call(state.excludedRows, kind)) return false;
  if (hasFieldConfigDraftChanges()) {
    setExclusionMessage(kind, () => Lpick(EXCLUSION_DRAFT_REFUSAL), "warning");
    return false;
  }
  const current = state.exclusionSourceRows === state.rows ? state.excludedRows[kind] : [];
  const indexes = on ? listExclusionCandidates(kind) : current;
  if (!indexes.length) return false;
  const excludedBefore = getExcludedIndexSet();
  const countedBefore = getIncludedRows().length;
  state.excludedRows = { ...state.excludedRows, [kind]: on ? [...indexes] : [] };
  // Left out again by the user: no longer a total the rebuild brought back.
  if (kind === "suspectedTotal" && on) indexes.forEach((index) => state.restoredTotalRows.delete(index));
  rebuildIncludedRows();
  const excludedAfter = getExcludedIndexSet();
  state.lastExclusion = {
    kind,
    on,
    rows: indexes.length,
    // Only the rows this click left out or brought back (a row of both kinds counts once).
    changed: on ? indexes.filter((index) => !excludedBefore.has(index)).length : indexes.filter((index) => !excludedAfter.has(index)).length,
    before: countedBefore,
    after: getIncludedRows().length
  };
  refreshAfterExclusion();
  return true;
}

// The applied filters again from the new counted rows, the filter panel, then the analysis.
function refreshAfterExclusion() {
  state.checkupSummary = buildCheckupSummary();
  const included = getIncludedRows();
  state.filteredRows = hasActiveFilters() ? filterRowsByConfig(included, state.filters) : included.slice();
  state.filterSourceRows = included;
  refreshFilterPanelForIncludedRows();
  refreshFilteredAnalysis();
  const activeCount = [state.filters.category, state.filters.numeric, state.filters.date].filter(Boolean).length;
  const keptCount = state.filteredRows.length;
  const excludedCount = countExcludedRows();
  setFilterMessage(() => (activeCount
    ? formatFiltersApplied(activeCount, keptCount, included.length - keptCount)
    : formatNoFilterMessage(included.length, excludedCount)), activeCount && !keptCount ? "warning" : "");
  renderExclusionControls();
  if (state.problemView) renderProblemView();
  renderSidebar();
}

// The filter panel after the counted rows changed. The category options and the range hints
// follow them. A field without an applied filter also takes the new range in its inputs; an
// applied one keeps the user's numbers. Drafts of the category options stay where they still
// exist (all of them when all were chosen).
function refreshFilterPanelForIncludedRows() {
  const field = dom.categoryFilterField.value;
  if (field) {
    const wasAll = state.categoryFilterDraftValues.size === state.categoryFilterOptions.length;
    state.categoryFilterOptions = getCategoryFilterOptions(field);
    const values = state.categoryFilterOptions.map((option) => option.value);
    state.categoryFilterDraftValues = wasAll ? new Set(values) : new Set(values.filter((value) => state.categoryFilterDraftValues.has(value)));
    renderCategoryFilterValues();
  }
  const numericField = dom.numericFilterField.value;
  const numbers = numericField ? getNumericValues(numericField, getIncludedRows()) : [];
  if (numbers.length) {
    const { min, max } = numericExtents(numbers);
    writeLanguageText(dom.numericFilterRange, () => formatFilterRange(formatFieldNumber(numericField, min), formatFieldNumber(numericField, max)));
    if (state.filters.numeric?.field !== numericField) {
      dom.numericFilterMin.value = String(min);
      dom.numericFilterMax.value = String(max);
    }
  }
  const dateField = dom.dateFilterField.value;
  const dates = dateField ? getIncludedRows().map((row) => parseFieldDate(dateField, row[dateField])).filter(Boolean).sort((a, b) => a - b) : [];
  if (dates.length) {
    const start = dates[0].toISOString().slice(0, 10);
    const end = dates[dates.length - 1].toISOString().slice(0, 10);
    writeLanguageText(dom.dateFilterRange, () => formatFilterRange(start, end));
    if (state.filters.date?.field !== dateField) {
      dom.dateFilterStart.value = start;
      dom.dateFilterEnd.value = end;
    }
  }
}

// Rebuild path (field configuration, placeholders): the left-out totals still suspected stay
// out; the others come back (remembered in restoredTotalRows); a total that came back earlier
// and is flagged again is named once. New suspected totals are not left out. Repeated rows
// never change here: buildExactRowKey() reads every field and no rebuild changes a cell.
function intersectExcludedTotals() {
  if (state.exclusionSourceRows !== state.rows) return { restored: 0, flaggedAgain: [] };
  const current = new Set(state.suspectedTotals.map((item) => item.index));
  const kept = state.excludedRows.suspectedTotal.filter((index) => current.has(index));
  const restored = state.excludedRows.suspectedTotal.filter((index) => !current.has(index));
  const flaggedAgain = [...state.restoredTotalRows].filter((index) => current.has(index)).sort((a, b) => a - b);
  flaggedAgain.forEach((index) => state.restoredTotalRows.delete(index));
  restored.forEach((index) => state.restoredTotalRows.add(index));
  state.excludedRows = { ...state.excludedRows, suspectedTotal: kept };
  rebuildIncludedRows();
  return { restored: restored.length, flaggedAgain: flaggedAgain.map((index) => getSourceRowNumber(state.rows[index])) };
}

function renderQuality() {
  const qualityNumericStats = getQualityNumericStats();
  const missingClass = state.totalMissing === 0 ? "success" : missingRateLevel();
  const duplicateClass = state.duplicateRows === 0 ? "success" : "warning";
  const conversionFailureTotal = state.profiles.reduce((sum, profile) => sum + (profile.conversionFailureCount || 0), 0);
  const conversionFailureClass = conversionFailureTotal === 0 ? "success" : "warning";
  const listedIdFieldCount = countListedIdDuplicateFields(state.idDuplicates);
  const pendingPlaceholderFieldCount = countPendingPlaceholderFields(state.placeholders);
  // First, so every 查看 button below can carry its category's index.
  state.problemCategories = listProblemCategories();
  dom.qualitySummary.innerHTML = `
    <span class="quality-pill ${missingClass}">${L(`空值总数：${formatInteger(state.totalMissing)}`, `Blanks: ${formatInteger(state.totalMissing)}`)}</span>
    <span class="quality-pill ${duplicateClass}">${L(`重复行：${formatInteger(state.duplicateRows)}`, `Duplicate rows: ${formatInteger(state.duplicateRows)}`)}</span>
    <span class="quality-pill ${conversionFailureClass}">${L(`认不出的数字或日期：${formatInteger(conversionFailureTotal)} 个值`, `Numbers or dates not recognized: ${enCount(conversionFailureTotal, "value", "values")}`)}</span>
    <span class="quality-pill">${formatQualityFieldPill(Lp("偏大或偏小的值", "Unusually high or low values"), qualityNumericStats.filter((item) => item.outlierCount > 0).length)}</span>
    <span class="quality-pill ${listedIdFieldCount ? "warning" : "success"}">${formatQualityFieldPill(Lp("编号重复", "Duplicate IDs"), listedIdFieldCount)}</span>
    <span class="quality-pill ${pendingPlaceholderFieldCount ? "warning" : "success"}">${formatQualityFieldPill(Lp("像是空值的写法", "Values that may mean blank"), pendingPlaceholderFieldCount)}</span>
    ${problemViewButton({ kind: "duplicateRows" }, L("查看重复行", "View duplicate rows"))}
  `;
  renderParseWarningSummary();
  renderExclusionControls();
  renderFormulaErrorDetails();
  // Before the outlier table, which returns early when there is no numeric field.
  refreshProblemView();
  renderNumberNoteSection(dom.numberConversionSection, dom.numberConversionDetails, buildNumberConversionSummary(state.profiles).map(formatNumberConversion));
  renderNumberNoteSection(dom.numericLookingSection, dom.numericLookingDetails, state.numericLooking.map(formatNumericLookingHint));
  renderFormatConsistencySections();
  renderIdDuplicateDetails();
  renderPlaceholderDetails();

  dom.qualityTable.innerHTML = `
    <thead>
      <tr>
        <th>${L("字段", "Column")}</th>
        <th>${L("类型", "Type")}</th>
        <th class="num">${L("空值数量", "Blanks")}</th>
        <th class="num">${L("空值占比", "Blank share")}</th>
        <th class="num">${L("认不出", "Unrecognized")}</th>
        <th>${L("查看", "View")}</th>
      </tr>
    </thead>
    <tbody>
      ${state.profiles.map((profile) => `
        <tr>
          <td>${escapeHtml(profile.field)}</td>
          <td>${escapeHtml(fieldTypeLabel(profile.typeKey))}</td>
          <td class="num">${formatInteger(profile.missingCount)}</td>
          <td class="num">${formatPercent(profile.missingRate)}</td>
          <td class="num${profile.conversionFailureCount ? " cell-warning" : ""}">${escapeHtml(formatConversionFailures(profile.conversionFailureCount, profile.conversionFailureSamples, formatInteger(profile.conversionFailureCount)))}</td>
          <td>${problemViewCell([
            problemViewButton({ kind: "missing", field: profile.field }, L("空值", "Blanks")),
            problemViewButton({ kind: "conversion", field: profile.field }, L("认不出", "Unrecognized"))
          ])}</td>
        </tr>
      `).join("")}
    </tbody>
  `;

  if (!qualityNumericStats.length) {
    dom.outlierTable.classList.add("is-empty-table");
    dom.outlierTable.innerHTML = emptyTable(L("未识别到数值字段。", "No number columns found."));
    return;
  }

  dom.outlierTable.classList.remove("is-empty-table");
  dom.outlierTable.innerHTML = `
    <thead>
      <tr>
        <th>${L("字段", "Column")}</th>
        <th class="num">${L("偏大或偏小的个数", "High or low values")}</th>
        <th class="num">${L("下界", "Lower bound")}</th>
        <th class="num">${L("上界", "Upper bound")}</th>
        <th>${L("查看", "View")}</th>
      </tr>
    </thead>
    <tbody>
      ${qualityNumericStats.map((stat) => `
        <tr>
          <td>${escapeHtml(formatFieldLabel(stat.field))}</td>
          <td class="num">${formatInteger(stat.outlierCount)}</td>
          <td class="num">${formatFieldNumber(stat.field, stat.lowerBound)}</td>
          <td class="num">${formatFieldNumber(stat.field, stat.upperBound)}</td>
          <td>${problemViewCell([problemViewButton({ kind: "outlier", field: stat.field }, Lpick(VIEW_LABEL))])}</td>
        </tr>
      `).join("")}
    </tbody>
  `;
}

const VIEW_LABEL = Lp("查看", "View");

// "偏大或偏小的值：2 个字段" and the other pills that count columns.
function formatQualityFieldPill(name, count) {
  return L(`${name[0]}：${formatInteger(count)} 个字段`, `${name[1]}: ${enCount(count, "column", "columns")}`);
}

// A 查看 button for a category of state.problemCategories, or "" when the category
// has nothing to show. It carries only the category's index (no field name).
function problemViewButton(category, text) {
  const index = state.problemCategories.findIndex((item) => isSameProblemCategory(item, category));
  if (index < 0) return "";
  return `<button type="button" class="secondary-button problem-view-button" data-problem-index="${index}" aria-label="${escapeHtml(formatProblemViewLabel(state.problemCategories[index]))}">${escapeHtml(text)}</button>`;
}

// The content of a 查看 cell: its buttons, or a short note so the cell is never empty.
function problemViewCell(buttons) {
  const html = buttons.filter(Boolean).join("");
  return html || `<span class="problem-view-none">${L("无", "None")}</span>`;
}

function isSameProblemCategory(a, b) {
  return a.kind === b.kind && a.field === b.field && a.key === b.key;
}

// "CSV 解析警告：N 条，涉及 M 行" with a 查看 button, under the pills; hidden without warnings.
function renderParseWarningSummary() {
  const count = state.parseWarnings.length;
  dom.parseWarningSummary.classList.toggle("hidden", !count);
  if (!count) {
    dom.parseWarningSummary.innerHTML = "";
    return;
  }
  const rowCount = new Set(state.parseWarnings.map((warning) => warning.row).filter((row) => row !== null)).size;
  dom.parseWarningSummary.innerHTML = `${escapeHtml(L(`读取时有疑问：${formatInteger(count)} 处${rowCount ? `，涉及 ${formatInteger(rowCount)} 行` : ""}`, `Reading issues: ${formatInteger(count)}${rowCount ? ` in ${enCount(rowCount, "row", "rows")}` : ""}`))} ${problemViewButton({ kind: "parseWarnings" }, Lpick(VIEW_LABEL))}`;
}

// "疑似合计 / 小计行" and "公式错误单元格": lines of ready HTML, then footer, the whole
// subsection hidden when there is none. They come before the tables they explain and
// before the early return of renderQuality() (formula errors may be all a table has).
function renderRowNoteSection(section, details, lines, footer = "") {
  section.classList.toggle("hidden", !lines.length);
  details.innerHTML = lines.length
    ? `<ul class="number-note-list">${lines.map((line) => `<li>${line}</li>`).join("")}</ul>${footer}`
    : "";
}

// At most SUSPECTED_TOTAL_LIST_LIMIT lines and one 查看 button: they are one category.
// v3.1.0: the static note gives way to the left-out note (CSS hides it on .is-excluded); the
// button sits next to 查看; the totals that only equal the row above and the labelled rows that
// do not add up are named under it.
// What the two sentences under the suspected totals need (page and report): the totals whose
// every match covers only one row (R1), the labelled rows, and whether any numeric column was
// checked at all. null without suspected totals.
function getSuspectedTotalExtras() {
  if (!state.suspectedTotals.length) return null;
  return {
    singleRow: state.suspectedTotals.filter((item) => item.matches.every((match) => match.rowCount === 1)).length,
    // v3.3.0: the labelled rows still checked, not every labelled row (dropped groups are values).
    labelRows: state.suspectedTotalLabelCount,
    checked: state.profiles.some((profile) => profile.typeKey === "numeric")
  };
}

function renderSuspectedTotalDetails() {
  const overflow = formatSuspectedTotalOverflow(state.suspectedTotals.length);
  const button = problemViewButton({ kind: "suspectedTotal" }, Lpick(VIEW_LABEL));
  const excludedCount = state.exclusionSourceRows === state.rows ? state.excludedRows.suspectedTotal.length : 0;
  dom.suspectedTotalSection.classList.toggle("is-excluded", excludedCount > 0);
  dom.suspectedTotalExcludedNote.classList.toggle("hidden", excludedCount === 0);
  dom.suspectedTotalExcludedNote.textContent = excludedCount ? formatSuspectedTotalExcludedNote(excludedCount) : "";
  const exclusionButton = state.suspectedTotals.length ? exclusionToggleButton("suspectedTotal", excludedCount > 0) : "";
  const notes = formatSuspectedTotalExtras(getSuspectedTotalExtras(), state.suspectedTotals.length);
  renderRowNoteSection(
    dom.suspectedTotalSection,
    dom.suspectedTotalDetails,
    state.suspectedTotals.slice(0, SUSPECTED_TOTAL_LIST_LIMIT).map((item) => escapeHtml(formatSuspectedTotal(item))),
    `<p class="row-note-footer">${overflow ? `${escapeHtml(L(`${overflow}全部行请点「查看」。`, `${overflow} Click “View” for all of them.`))} ` : ""}${button}${exclusionButton}</p>${notes.map((note) => `<p class="row-note-footer">${escapeHtml(note)}</p>`).join("")}`
  );
}

// "不计入分析" or "撤销" for a kind; it carries no data-problem-index or data-placeholder-index.
function exclusionToggleButton(kind, excluded) {
  const label = Lpick(EXCLUSION_KIND_LABELS[kind]);
  const ariaLabel = excluded
    ? L(`撤销：让${label}重新计入分析`, `Undo: count the ${label} in the analysis again`)
    : L(`把${label}设为不计入分析`, `Exclude the ${label} from analysis`);
  return `<button type="button" class="secondary-button row-exclusion-button" data-exclusion-kind="${kind}" data-exclusion-on="${excluded ? "false" : "true"}" aria-label="${escapeHtml(ariaLabel)}">${excluded ? L("撤销", "Undo") : L("不计入分析", "Exclude from analysis")}</button>`;
}

// An applied ID field that is not mostly repeated (state.idDuplicates lists the applied ID fields).
function hasCleanIdField() {
  return state.idDuplicates.some((item) => !item.mostlyRepeated);
}

// Missing values in the left-out rows, counted as state.totalMissing counts them: every field
// (ignored ones too), with the placeholder keys each profile was built with.
function countExcludedMissing() {
  let count = 0;
  getExcludedIndexSet().forEach((index) => {
    const row = state.rows[index];
    state.profiles.forEach((profile) => {
      if (isMissingOrConfirmed(row[profile.field], profile.confirmedKeys)) count += 1;
    });
  });
  return count;
}

// The quality page's parts that follow the rows left out; its counts never change.
function renderExclusionControls() {
  const excludedCount = countExcludedRows();
  const note = formatQualityScopeNote(state.rows.length, excludedCount, excludedCount ? countExcludedMissing() : 0);
  dom.qualityScopeNote.textContent = note;
  dom.qualityScopeNote.classList.toggle("hidden", !note);
  renderSuspectedTotalDetails();
  const duplicateExcluded = state.exclusionSourceRows === state.rows ? state.excludedRows.duplicateRows.length : 0;
  dom.duplicateExclusion.classList.toggle("hidden", !state.duplicateRows);
  dom.duplicateExclusionText.textContent = !state.duplicateRows ? ""
    : duplicateExcluded ? formatDuplicateExcludedNote(duplicateExcluded)
      : hasCleanIdField() ? formatDuplicateIncludedNote(state.duplicateRows) : Lpick(NO_ID_DUPLICATE_NOTE);
  dom.duplicateExclusionAction.innerHTML = state.duplicateRows ? exclusionToggleButton("duplicateRows", duplicateExcluded > 0) : "";
}

function getExclusionMessageElement(kind) {
  return kind === "duplicateRows" ? dom.duplicateExclusionMessage : dom.suspectedTotalMessage;
}

function setExclusionMessage(kind, message, tone = "") {
  writeOneOffMessage(getExclusionMessageElement(kind), "exclusion-message", message, tone);
}

// #problemRowsMessage has no class of its own.
function setProblemRowsMessage(message) {
  writeOneOffMessage(dom.problemRowsMessage, "", message);
}

// Delegated click of #qualitySection for 不计入分析 / 撤销, bound in script.js. The message goes
// to the section that was clicked; the focus to the button that took the clicked one's place.
function handleExclusionClick(event) {
  const button = event.target.closest("button[data-exclusion-kind]");
  if (!button) return;
  const kind = button.dataset.exclusionKind;
  if (!setRowExclusion(kind, button.dataset.exclusionOn === "true")) return;
  const exclusion = state.lastExclusion;
  setExclusionMessage(kind, () => formatExclusionMessage(exclusion), "success");
  const container = kind === "duplicateRows" ? dom.duplicateExclusionAction : dom.suspectedTotalDetails;
  container.querySelector?.("button[data-exclusion-kind]")?.focus?.();
}

// One line and one 查看 button per field: each field is a category.
function renderFormulaErrorDetails() {
  renderRowNoteSection(dom.formulaErrorSection, dom.formulaErrorDetails, state.formulaErrors.map((item) => {
    const text = formatFormulaError({ ...item, fieldIgnored: getFieldProfile(item.field)?.typeKey === "ignore" });
    return `${escapeHtml(text)} ${problemViewButton({ kind: "formulaError", field: item.field }, Lpick(VIEW_LABEL))}`;
  }));
}

// Delegated click handler of #qualitySection for the 查看 buttons, bound in script.js.
function handleProblemViewClick(event) {
  const button = event.target.closest("button[data-problem-index]");
  if (!button) return;
  const category = state.problemCategories[Number(button.dataset.problemIndex)];
  if (!category || !openProblemView(category)) return;
  state.problemViewOpener = button;
  // Only 查看 scrolls: the section goes to the top (under the bar, via scroll-padding-top);
  // focusing the title alone would only bring it into view, at the bottom of the screen.
  dom.problemRowsSection.scrollIntoView({ block: "start" });
  dom.problemRowsTitle.focus({ preventScroll: true });
}

// Opens the viewer on a category of state.problemCategories at page 1. Returns false,
// leaving the viewer as it was, when the category has nothing to show.
function openProblemView(category) {
  const match = state.problemCategories.find((item) => isSameProblemCategory(item, category));
  const list = match ? listProblemRows(match) : null;
  if (!list) return false;
  state.problemView = { category: { kind: match.kind, field: match.field, key: match.key }, page: 1, list };
  setProblemRowsMessage("");
  setProblemRowsExportStatus("");
  renderProblemView();
  return true;
}

function selectProblemCategory(value) {
  const category = state.problemCategories[Number(value)];
  if (category) openProblemView(category);
}

function goToProblemRowsPage(action) {
  const view = state.problemView;
  if (!view) return;
  const { page, pageCount } = getRowPage(view.list.items.length, view.page);
  view.page = movePage(action, page, pageCount);
  renderProblemView();
}

// Kinds whose rows have one cell to point at: the field of the category.
const HIGHLIGHTED_PROBLEM_KINDS = new Set(["missing", "conversion", "outlier", "placeholder", "idDuplicate", "formulaError", "dateStyle", "lookAlike"]);

// The open category's current page: category list, summary, table and pager.
function renderProblemView() {
  const view = state.problemView;
  const index = state.problemCategories.findIndex((item) => isSameProblemCategory(item, view.category));
  const category = state.problemCategories[index];
  const { page, pageCount, start, end } = getRowPage(view.list.items.length, view.page);
  view.page = page;
  dom.problemRowsSection.classList.remove("hidden");
  dom.problemRowsCategory.innerHTML = state.problemCategories.map((item, itemIndex) => (
    `<option value="${itemIndex}"${itemIndex === index ? " selected" : ""}>${escapeHtml(formatProblemCategoryLabel(item))}</option>`
  )).join("");
  dom.problemRowsCategory.value = String(index);
  dom.problemRowsSummary.textContent = formatProblemRowsSummary(category, view.list, page, pageCount);
  dom.problemRowsDraftNote.classList.toggle("hidden", !hasFieldConfigDraftChanges());
  const highlightField = HIGHLIGHTED_PROBLEM_KINDS.has(category.kind) ? category.field : null;
  const excluded = getExcludedIndexSet();
  dom.problemRowsTable.innerHTML = buildRowTableHtml(state.fields, view.list.items.slice(start, end).map((item) => ({
    row: item.index === null ? null : state.rows[item.index],
    line: item.line,
    note: markExcludedNote(item, excluded),
    highlightField: item.highlightField ?? highlightField
  })), { withNote: true });
  setPagerButtons([dom.problemRowsFirstPage, dom.problemRowsPrevPage, dom.problemRowsNextPage, dom.problemRowsLastPage], page, pageCount);
}

// The 关闭 button: hides the viewer and gives the focus back to the 查看 button that
// opened it, or, when a re-render replaced that button, to the new button of the
// same category; with neither, the focus stays.
function closeProblemView() {
  const category = state.problemView?.category;
  const opener = state.problemViewOpener;
  hideProblemView();
  const target = opener?.isConnected ? opener : findProblemViewButton(category);
  if (target) target.focus();
}

function findProblemViewButton(category) {
  if (!category || typeof dom.qualitySection.querySelector !== "function") return null;
  const index = state.problemCategories.findIndex((item) => isSameProblemCategory(item, category));
  return index < 0 ? null : dom.qualitySection.querySelector(`button[data-problem-index="${index}"]`);
}

// v3.1.0: "（不计入分析）" after the note of a left-out row, in the viewer and in 导出这一类.
// Called once per listed row: a pair read with Lpick(), no L().
const EXCLUDED_NOTE_SUFFIX = Lp("（不计入分析）", " (excluded from analysis)");

function markExcludedNote(item, excluded) {
  return item.index !== null && excluded.has(item.index) ? `${item.note || ""}${Lpick(EXCLUDED_NOTE_SUFFIX)}` : item.note;
}

function hideProblemView() {
  state.problemView = null;
  state.problemViewOpener = null;
  dom.problemRowsSection.classList.add("hidden");
  dom.problemRowsCategory.innerHTML = "";
  dom.problemRowsSummary.textContent = "";
  dom.problemRowsTable.innerHTML = "";
  dom.problemRowsDraftNote.classList.add("hidden");
  setProblemRowsMessage("");
  setProblemRowsExportStatus("");
  setPagerButtons([dom.problemRowsFirstPage, dom.problemRowsPrevPage, dom.problemRowsNextPage, dom.problemRowsLastPage], 1, 1);
}

// After field configuration or placeholder confirmations (renderQuality()): lists the
// open category again from page 1, or closes the viewer when it has no rows now.
function refreshProblemView() {
  const view = state.problemView;
  if (!view) return;
  if (!openProblemView(view.category)) {
    hideProblemView();
    setProblemRowsMessage(() => L("这一类已经没有问题行。", "This category has no problem rows any more."));
  }
}

// "日期写法混用" and "写法细节不同的值": one line and one 查看 button per field. "常量列" has no
// button: every row looks the same. They come before the early return of renderQuality().
function renderFormatConsistencySections() {
  renderRowNoteSection(dom.dateStyleSection, dom.dateStyleDetails, state.dateStyles.map((item) => (
    `${escapeHtml(formatDateStyleLine(item))} ${problemViewButton({ kind: "dateStyle", field: item.field }, Lpick(VIEW_LABEL))}`
  )));
  renderNumberNoteSection(dom.constantFieldSection, dom.constantFieldDetails, state.constantFields.map(formatConstantField));
  renderRowNoteSection(dom.lookAlikeSection, dom.lookAlikeDetails, state.lookAlikes.map((item) => (
    `${escapeHtml(formatLookAlikeLine(item))} ${problemViewButton({ kind: "lookAlike", field: item.field }, Lpick(VIEW_LABEL))}`
  )));
}

// "按单位读取的数字" and "疑似数值列": one line per field, the whole subsection
// hidden when there is none.
function renderNumberNoteSection(section, details, lines) {
  section.classList.toggle("hidden", !lines.length);
  details.innerHTML = lines.length
    ? `<ul class="number-note-list">${lines.map((line) => `<li>${escapeHtml(line)}</li>`).join("")}</ul>`
    : "";
}

// One table per ID field with listed duplicates, a hint per mostly repeated field,
// and a closing note naming the ID fields without duplicates.
function renderIdDuplicateDetails() {
  if (!state.idDuplicates.length) {
    dom.idDuplicateDetails.classList.add("is-empty");
    dom.idDuplicateDetails.innerHTML = `<p class="id-duplicate-note">${L("未识别到 ID 字段。", "No ID columns found.")}</p>`;
    return;
  }
  const reported = state.idDuplicates.filter((item) => item.mostlyRepeated || item.values.length);
  const cleanFields = state.idDuplicates.filter((item) => !reported.includes(item)).map((item) => item.field);
  dom.idDuplicateDetails.classList.remove("is-empty");
  dom.idDuplicateDetails.innerHTML = [
    ...reported.map((item) => {
      if (item.mostlyRepeated) {
        return `<p class="id-duplicate-note warning">${escapeHtml(formatMostlyRepeatedIdHint(item))}</p>`;
      }
      const overflow = formatIdDuplicateOverflow(item.duplicateValueCount, item.values.length);
      return `
        <div class="id-duplicate-field">
          <p class="id-duplicate-field-name">${escapeHtml(item.field)} ${problemViewButton({ kind: "idDuplicate", field: item.field }, Lpick(VIEW_LABEL))}</p>
          <div class="table-wrap" tabindex="0" aria-label="${escapeHtml(L(`${item.field} 的重复值表，可横向滚动`, `Duplicate values of ${item.field}, scrolls sideways`))}">
            <table aria-label="${escapeHtml(L(`${item.field} 的重复值`, `Duplicate values of ${item.field}`))}">
              <thead><tr>${L(["重复值", "出现次数", "文件行号", "说明"], ["Duplicate value", "Count", "File rows", "Note"]).map((title) => `<th>${title}</th>`).join("")}</tr></thead>
              <tbody>${item.values.map((value) => `<tr><td>${escapeHtml(value.value)}</td><td>${formatInteger(value.count)}</td><td>${escapeHtml(formatRowNumberList(value.rowNumbers, value.count))}</td><td>${escapeHtml(formatIdenticalRows(value.identicalRowCount, value.identicalGroupCount))}</td></tr>`).join("")}</tbody>
            </table>
          </div>
          ${overflow ? `<p class="id-duplicate-note">${escapeHtml(overflow)}</p>` : ""}
        </div>
      `;
    }),
    cleanFields.length ? `<p class="id-duplicate-note">${escapeHtml(formatCleanIdFieldsNote(cleanFields, reported.length > 0))}</p>` : ""
  ].join("");
}

// One item per field and key (state.placeholders order): pending keys with their
// effect and a confirm button, confirmed keys with an undo button, a whole-field
// line for fields with two or more pending keys, and the confirm-all button.
// Buttons carry indexes into state.placeholders for handlePlaceholderClick().
function renderPlaceholderDetails() {
  if (!state.placeholders.length) {
    dom.placeholderDetails.classList.add("is-empty");
    dom.placeholderDetails.innerHTML = `<p class="placeholder-note">${L("没有发现像是空值的写法。", "No values that may mean blank were found.")}</p>`;
    return;
  }
  const commonCount = getPendingCommonPlaceholders().length;
  const typeChanges = commonCount ? previewAllCommonPlaceholders() : [];
  const items = state.placeholders.map((entry, index) => {
    const textId = `placeholderText${index}`;
    const profile = getFieldProfile(entry.field);
    // Its own data-problem-index: handlePlaceholderClick() only reads data-placeholder-*.
    const viewButton = problemViewButton({ kind: "placeholder", field: entry.field, key: entry.key }, Lpick(VIEW_LABEL));
    let item;
    if (entry.confirmed) {
      item = `<li class="placeholder-item confirmed"><p id="${textId}" class="placeholder-text">${escapeHtml(formatConfirmedPlaceholder(entry))}</p><button type="button" class="secondary-button placeholder-button" data-placeholder-index="${index}" aria-describedby="${textId}">${L("撤销", "Undo")}</button>${viewButton}</li>`;
    } else {
      const change = formatPlaceholderChange(entry.field, profile.missingCount, profile.typeKey, previewPlaceholderConfirmation(entry.field, [entry.key]));
      const notes = formatPlaceholderNotes([entry.key]);
      item = `<li class="placeholder-item"><p id="${textId}" class="placeholder-text">${escapeHtml(formatPendingPlaceholder(entry, change))}</p>${notes ? `<p class="placeholder-note warning">${escapeHtml(notes)}</p>` : ""}<button type="button" class="placeholder-button" data-placeholder-index="${index}" aria-describedby="${textId}">${L("按空值处理", "Treat as blank")}</button>${viewButton}</li>`;
    }
    if (state.placeholders[index + 1]?.field === entry.field) return item;
    const pendingKeys = state.placeholders.filter((other) => other.field === entry.field && !other.confirmed).map((other) => other.key);
    if (pendingKeys.length < 2) return item;
    const change = formatPlaceholderChange(entry.field, profile.missingCount, profile.typeKey, previewPlaceholderConfirmation(entry.field, pendingKeys));
    return `${item}<li class="placeholder-field-summary">${escapeHtml(L(`「${entry.field}」全部按空值处理后：${change}。`, `After treating all of them in “${entry.field}” as blank: ${change}.`))}</li>`;
  });
  const typeChangeNote = typeChanges.length
    ? `<span class="placeholder-note">${escapeHtml(L(`按空值处理后类型会变：${typeChanges.map((change) => `「${change.field}」`).join("")}`, `Treating them as blank changes the type of ${typeChanges.map((change) => `“${change.field}”`).join(", ")}`))}</span>`
    : "";
  dom.placeholderDetails.classList.remove("is-empty");
  dom.placeholderDetails.innerHTML = `
    ${commonCount ? `<div class="placeholder-actions"><button type="button" class="placeholder-button" data-placeholder-all>${L(`常见写法全部按空值处理（${formatInteger(commonCount)} 处）`, `Treat all common ones as blank (${formatInteger(commonCount)})`)}</button>${typeChangeNote}</div>` : ""}
    <ul class="placeholder-list">${items.join("")}</ul>
  `;
}

// Delegated click handler of #placeholderDetails, bound in script.js.
function handlePlaceholderClick(event) {
  const button = event.target.closest("button[data-placeholder-index], button[data-placeholder-all]");
  if (!button) return;
  if (button.dataset.placeholderAll !== undefined) {
    confirmAllCommonPlaceholders();
    return;
  }
  const entry = state.placeholders[Number(button.dataset.placeholderIndex)];
  if (!entry) return;
  if (entry.confirmed) undoPlaceholderConfirmation(entry.field, entry.key);
  else confirmPlaceholder(entry.field, entry.key);
}

function renderNumericStats() {
  if (!state.numericStats.length) {
    dom.numericStatsTable.classList.add("is-empty-table");
    dom.numericStatsTable.innerHTML = emptyTable(L("未识别到数值字段。", "No number columns found."));
    return;
  }

  dom.numericStatsTable.classList.remove("is-empty-table");
  dom.numericStatsTable.innerHTML = `
    <thead>
      <tr>
        <th>${L("字段", "Column")}</th>
        <th class="num">${L("有效数量", "Count")}</th>
        <th class="num">${L("认不出", "Unrecognized")}</th>
        <th class="num">${L("平均值", "Mean")}</th>
        <th class="num">${L("中位数", "Median")}</th>
        <th class="num">${L("最大值", "Max")}</th>
        <th class="num">${L("最小值", "Min")}</th>
        <th class="num">${L("标准差", "Std dev")}</th>
      </tr>
    </thead>
    <tbody>
      ${state.numericStats.map((stat) => `
        <tr>
          <td>${escapeHtml(formatFieldLabel(stat.field))}</td>
          <td class="num">${formatInteger(stat.count)}</td>
          <td class="num${stat.invalidCount ? " cell-warning" : ""}">${formatInteger(stat.invalidCount)}</td>
          <td class="num">${formatFieldNumber(stat.field, stat.mean)}</td>
          <td class="num">${formatFieldNumber(stat.field, stat.median)}</td>
          <td class="num">${formatFieldNumber(stat.field, stat.max)}</td>
          <td class="num">${formatFieldNumber(stat.field, stat.min)}</td>
          <td class="num">${formatFieldNumber(stat.field, stat.std)}</td>
        </tr>
      `).join("")}
    </tbody>
  `;
}

function renderCategoryStats() {
  if (!state.categoryStats.length) {
    dom.categoryStats.innerHTML = `<div class="empty-note">${L("未识别到分类字段。", "No text columns found.")}</div>`;
    return;
  }

  // v3.0.0: a bar as long as the share in each row; the card title carries a tooltip.
  dom.categoryStats.innerHTML = state.categoryStats.map((category) => `
    <div class="frequency-card">
      <h3 title="${escapeHtml(category.field)}">${escapeHtml(category.field)}</h3>
      ${category.top.length ? category.top.map((item) => `
        <div class="frequency-row">
          <span class="category-name" title="${escapeHtml(item.name)}">${escapeHtml(item.name)}</span>
          <span class="frequency-bar" aria-hidden="true"><span style="width: ${(absorbFloatNoise(item.ratio) * 100).toFixed(1)}%"></span></span>
          <strong>${formatInteger(item.count)}</strong>
          <span>${formatPercent(item.ratio)}</span>
        </div>
      `).join("") : `<div class="empty-note">${escapeHtml(Lpick(CATEGORY_STATS_EMPTY_NOTE))}</div>`}
    </div>
  `).join("");
}

// A new import resets the trend chart to record counts and the category chart to the
// first concentrated category field (or the first category field when none is
// concentrated). A field-configuration rebuild keeps the chosen trend metric while it is
// still a numeric field with data, and the chosen category field while it is still a
// category field.
function populateControls({ keepChartSelections = false } = {}) {
  const numericProfiles = getProfilesByType("numeric")
    .filter((profile) => fieldHasNumericData(profile.field));
  const categoryProfiles = getProfilesByType("category");
  const dateProfiles = getProfilesByType("date")
    .filter((profile) => fieldHasDateData(profile.field));
  const numericFields = numericProfiles.map((profile) => profile.field);
  const previousDateChartMetric = dom.dateChartMetric.value;
  const previousDateChartAggregate = dom.dateChartAggregate.value;
  const previousCategoryChartField = dom.categoryChartField.value;
  const categoryFields = categoryProfiles.map((profile) => profile.field);

  setSelectOptions(dom.numericChartField, numericProfiles.map((profile) => profile.field));
  setSelectOptions(dom.categoryChartField, categoryFields);
  dom.categoryChartField.value = keepChartSelections && categoryFields.includes(previousCategoryChartField)
    ? previousCategoryChartField
    : findConcentratedCategory(state.categoryStats)?.field ?? categoryFields[0] ?? "";
  setSelectOptions(dom.dateChartField, dateProfiles.map((profile) => profile.field));
  dom.dateChartField.value = dateProfiles[0]?.field ?? "";
  setSelectOptions(dom.dateChartMetric, numericFields, L("记录数", "Record count"), true);
  dom.dateChartMetric.value = keepChartSelections && numericFields.includes(previousDateChartMetric)
    ? previousDateChartMetric
    : "";
  dom.dateChartAggregate.value = keepChartSelections && previousDateChartAggregate === "sum" ? "sum" : "avg";
  setSelectOptions(dom.v2MetricField, numericProfiles.map((profile) => profile.field), L("请选择数值字段", "Choose a number column"));
  setSelectOptions(dom.v2GroupField, categoryProfiles.map((profile) => profile.field), L("请选择分类字段", "Choose a text column"));
  setSelectOptions(dom.v2DateField, dateProfiles.map((profile) => profile.field), L("不选择时间字段", "No date column"), true);
}

function renderCharts() {
  if (!window.Chart) {
    ["numericDistributionChart", "categoryTopChart", "dateTrendChart"].forEach((id) => {
      drawEmptyChart(document.getElementById(id), Lpick(CHART_LIBRARY_MISSING));
    });
    return;
  }
  renderNumericDistributionChart();
  renderCategoryTopChart();
  renderDateTrendChart();
}

function renderNumericDistributionChart() {
  const field = dom.numericChartField.value;
  const canvas = document.getElementById("numericDistributionChart");
  destroyChart("numericDistributionChart");
  dom.numericChartTitle.textContent = field ? L(`${formatFieldLabel(field)} 数值分布（记录频数）`, `${formatFieldLabel(field)} distribution (record count)`) : L("数值字段分布（记录频数）", "Number column distribution (record count)");
  canvas.setAttribute("aria-label", dom.numericChartTitle.textContent);
  dom.numericChartTitle.title = dom.numericChartTitle.textContent;
  if (!field) return drawEmptyChart(canvas, L("没有可用的数值字段", "No number columns available"));

  const values = getNumericValues(field);
  if (!values.length) return drawEmptyChart(canvas, L("该字段没有认得出的数字", "This column has no recognized numbers"));
  const bins = buildDiscreteValueCounts(values, field) || buildHistogram(values, 10, field);
  state.charts.numericDistributionChart = new Chart(canvas, {
    type: "bar",
    data: {
      labels: bins.map((bin) => bin.label),
      datasets: [{
        label: L(`${formatFieldLabel(field)} 分布（记录频数）`, `${formatFieldLabel(field)} distribution (record count)`),
        data: bins.map((bin) => bin.count),
        backgroundColor: CHART_THEME.primary
      }]
    },
    options: chartOptions(L("记录频数", "Record count"))
  });
}

function renderCategoryTopChart() {
  const field = dom.categoryChartField.value;
  const canvas = document.getElementById("categoryTopChart");
  destroyChart("categoryTopChart");
  dom.categoryChartTitle.textContent = field ? L(`${field} Top 10（记录数）`, `${field} top 10 (record count)`) : L("分类字段 Top 10（记录数）", "Text column top 10 (record count)");
  canvas.setAttribute("aria-label", dom.categoryChartTitle.textContent);
  dom.categoryChartTitle.title = dom.categoryChartTitle.textContent;
  if (!field) return drawEmptyChart(canvas, L("没有可用的分类字段", "No text columns available"));

  const counts = Array.from(countValues(field).entries())
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10);
  if (!counts.length) return drawEmptyChart(canvas, Lpick(CATEGORY_CHART_EMPTY_NOTE));
  state.charts.categoryTopChart = new Chart(canvas, {
    type: "bar",
    data: {
      labels: counts.map(([name]) => name),
      datasets: [{
        label: L(`${field} Top 10（记录数）`, `${field} top 10 (record count)`),
        data: counts.map(([, count]) => count),
        backgroundColor: CHART_THEME.secondary
      }]
    },
    options: chartOptions(L("记录数", "Record count"))
  });
}

function renderDateTrendChart() {
  const field = dom.dateChartField.value;
  const canvas = document.getElementById("dateTrendChart");
  destroyChart("dateTrendChart");
  dom.dateChartTitle.textContent = L("日期字段趋势（记录数）", "Date column trend (record count)");
  canvas.setAttribute("aria-label", dom.dateChartTitle.textContent);
  dom.dateChartTitle.title = dom.dateChartTitle.textContent;
  const { metricField: numericField, aggregate } = getDateChartSelection();
  dom.dateChartAggregate.disabled = !numericField;
  if (!field) return drawEmptyChart(canvas, L("没有可用的日期字段", "No date columns available"));

  const trend = buildDateChartTrend(field, numericField, aggregate);
  if (!trend.labels.length) return drawEmptyChart(canvas, L("该字段没有认得出的日期", "This column has no recognized dates"));
  const timeUnit = getDateGranularity(field) === "year"
    ? L("年份", "year")
    : trend.labels.some((label) => /^\d{4}-\d{2}$/.test(label)) ? L("月份", "month") : L("日期", "date");
  const aggregationLabel = getDateChartAggregateLabel(aggregate);
  const yearly = getDateGranularity(field) === "year";
  dom.dateChartTitle.textContent = L(
    `${numericField ? formatFieldLabel(numericField) : field} 按${timeUnit}趋势（${aggregationLabel}）`,
    `${numericField ? formatFieldLabel(numericField) : field} by ${timeUnit} (${aggregationLabel})`
  );
  canvas.setAttribute("aria-label", dom.dateChartTitle.textContent);
  dom.dateChartTitle.title = dom.dateChartTitle.textContent;
  state.charts.dateTrendChart = new Chart(canvas, {
    type: "line",
    data: {
      labels: trend.labels,
      datasets: [{
        label: numericField
          ? L(`${formatFieldLabel(numericField)}${yearly ? "年度" : " "}${aggregationLabel}趋势`, `${formatFieldLabel(numericField)} ${yearly ? "yearly " : ""}${aggregationLabel} trend`)
          : L(`${yearly ? "年度" : ""}记录数趋势`, `${yearly ? "Yearly record count" : "Record count"} trend`),
        data: trend.values,
        ...trendLineStyle(CHART_THEME.primary)
      }]
    },
    options: chartOptions(numericAxisLabel(numericField, getDateChartAxisLabel(aggregate)), numericField)
  });
}

// Reads the trend chart's own metric and aggregate selects. An empty or no
// longer numeric metric means record counts.
function getDateChartSelection() {
  const selected = dom.dateChartMetric.value;
  const metricField = getProfilesByType("numeric").some((profile) => profile.field === selected) ? selected : "";
  if (!metricField) return { metricField: "", aggregate: "count" };
  return { metricField, aggregate: dom.dateChartAggregate.value === "sum" ? "sum" : "avg" };
}

function getDateChartAggregateLabel(aggregate) {
  return L({ count: "记录数", avg: "平均值", sum: "求和" }[aggregate] || "记录数", { count: "record count", avg: "average", sum: "sum" }[aggregate] || "record count");
}

// The y-axis title: the same Chinese, the English with a capital (the words above are for sentences).
function getDateChartAxisLabel(aggregate) {
  return L({ count: "记录数", avg: "平均值", sum: "求和" }[aggregate] || "记录数", { count: "Record count", avg: "Average", sum: "Sum" }[aggregate] || "Record count");
}

function resetV2AnalysisState() {
  destroyChart("v2TopGroupChart");
  destroyChart("v2SelectedTrendChart");
  state.customAnalysis = null;
  dom.v2AnalysisResults.classList.add("hidden");
  writeOneOffMessage(dom.v2FieldPrompt, "analysis-prompt", () => L("请选择指标字段、分组维度和聚合方式，再点击“生成自定义分析”。时间字段为可选。", "Choose a metric column, a column to group by and an aggregation, then click “Create breakdown”. A date column is optional."));
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
    writeOneOffMessage(dom.v2FieldPrompt, "analysis-prompt", () => getV2ValidationMessage(metricField, groupField, dateField), "warning");
    return;
  }

  const grouped = buildV2GroupedStats(metricField, groupField);
  if (!grouped.length) {
    state.customAnalysis = null;
    dom.v2AnalysisResults.classList.add("hidden");
    writeOneOffMessage(dom.v2FieldPrompt, "analysis-prompt", () => L("所选字段没有可计算的有效记录，请重新选择字段后再分析。", "The chosen columns have no valid rows to compute; choose other columns and try again."), "warning");
    return;
  }

  const metricInvalidCount = countInvalidNumericValues(metricField, getAnalysisRows());
  state.customAnalysis = {
    metricField,
    groupField,
    aggregateMethod,
    dateField,
    metricInvalidCount,
    grouped: grouped.map((item) => ({ ...item })),
    trend: dateField ? buildMetricDateTrend(dateField, metricField, aggregateMethod) : null
  };
  drawV2Analysis(grouped);
}

// v3.2.0 commit 11: the drawing part of renderV2Analysis(), from state.customAnalysis; a language
// switch calls it alone (the groups are not computed again). grouped: the groups to draw.
function drawV2Analysis(grouped = state.customAnalysis.grouped) {
  const { metricField, groupField, aggregateMethod, dateField, metricInvalidCount } = state.customAnalysis;
  const aggregateLabel = getAggregateMethodLabel(aggregateMethod);
  destroyChart("v2TopGroupChart");
  destroyChart("v2SelectedTrendChart");
  writeOneOffMessage(dom.v2FieldPrompt, "analysis-prompt", () => (dateField
    ? L(`已生成分组${getAggregateMethodLabel(aggregateMethod)}汇总和时间趋势。`, `Created the ${getAggregateMethodLabel(aggregateMethod)} by group and its trend over time.`)
    : L(`已生成分组${getAggregateMethodLabel(aggregateMethod)}汇总；未选择时间字段，因此不生成时间趋势图。`, `Created the ${getAggregateMethodLabel(aggregateMethod)} by group; no date column was chosen, so there is no trend chart.`)));
  dom.v2AnalysisResults.classList.remove("hidden");

  renderV2SummaryCards(metricField, groupField, aggregateMethod, dateField, grouped);
  renderV2AggregationTables(metricField, groupField, grouped);
  renderV2PrimaryTable(metricField, groupField, grouped, aggregateMethod);
  renderV2TopGroupChart(metricField, groupField, grouped, aggregateMethod);
  const invalidMetricText = metricInvalidCount
    ? L(
      `指标字段「${metricField}」有 ${formatInteger(metricInvalidCount)} 个值认不出是数字，已从计算中排除。`,
      ` ${enCount(metricInvalidCount, "value", "values")} of the metric column “${metricField}” ${metricInvalidCount === 1 ? "is" : "are"} not recognized as numbers and left out of the calculation.`
    )
    : "";
  dom.v2ResultSummary.textContent = formatCustomAnalysisScope({
    groupField, metricField, aggregateLabel, filtered: hasActiveFilters(), currentCount: getAnalysisRows().length,
    includedCount: getIncludedRows().length, totalCount: state.rows.length, excludedCount: countExcludedRows()
  }) + invalidMetricText;

  if (dateField) {
    dom.v2TrendCard.classList.remove("hidden");
    renderV2TrendChart(metricField, dateField, aggregateMethod);
  } else {
    dom.v2TrendCard.classList.add("hidden");
  }
}

function getV2ValidationMessage(metricField, groupField, dateField) {
  if (!state.rows.length) return L("请先上传或加载 CSV 数据。", "Upload or load a CSV file first.");
  if (!getAnalysisRows().length) return L("当前筛选条件下没有可分析的记录，请调整或清除筛选条件。", "No rows to analyze under the current filters; change or clear the filters.");
  if (!metricField) return L("请选择一个数值字段作为分析指标。", "Choose a number column as the metric.");
  if (!groupField) return L("请选择一个分类字段作为分组维度。", "Choose a text column to group by.");
  if (!getProfilesByType("numeric").some((profile) => profile.field === metricField)) {
    return L("分析指标必须是系统识别出的数值字段。", "The metric must be a column read as numbers.");
  }
  if (!getProfilesByType("category").some((profile) => profile.field === groupField)) {
    return L("分组维度必须是系统识别出的分类字段。", "The column to group by must be a column read as text.");
  }
  if (dateField && !getProfilesByType("date").some((profile) => profile.field === dateField)) {
    return L("时间字段必须是系统识别出的日期字段，或保持不选择。", "The date column must be a column read as dates, or none.");
  }
  return "";
}

function renderV2SummaryCards(metricField, groupField, aggregateMethod, dateField, grouped) {
  const validRecordCount = grouped.reduce((total, item) => total + item.count, 0);
  const trendPointCount = dateField ? buildMetricDateTrend(dateField, metricField, aggregateMethod).labels.length : 0;
  const cards = [
    [L("分析指标", "Metric"), formatFieldLabel(metricField)],
    [L("分组维度", "Group by"), groupField],
    [L("聚合方式", "Aggregation"), getAggregateMethodLabel(aggregateMethod)],
    [dateField ? L("时间点数量", "Time points") : L("有效记录数", "Valid rows"), dateField ? formatInteger(trendPointCount) : formatInteger(validRecordCount)]
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
  dom.v2PrimaryTableTitle.textContent = L(`${formatFieldLabel(metricField)} 按 ${groupField} 分组（${aggregateLabel}）`, `${formatFieldLabel(metricField)} by ${groupField} (${aggregateLabel})`);
  dom.v2PrimaryTableTitle.title = dom.v2PrimaryTableTitle.textContent;
  dom.v2PrimaryTable.innerHTML = buildV2AggregationTableHtml(groupField, metricField, sorted, aggregateMethod);
}

function buildV2AggregationTableHtml(groupField, metricField, rows, valueKey) {
  const valueLabel = L({ sum: "求和值", avg: "平均值", median: "中位数" }[valueKey] || "汇总值", { sum: "sum", avg: "average", median: "median" }[valueKey] || "total");
  return `
    <thead>
      <tr>
        <th>${escapeHtml(groupField)}</th>
        <th>${L("记录数", "Records")}</th>
        <th>${escapeHtml(formatFieldLabel(metricField))} ${valueLabel}</th>
      </tr>
    </thead>
    <tbody>
      ${rows.map((row) => `
        <tr>
          <td>${escapeHtml(displayGroupName(row.name))}</td>
          <td>${formatInteger(row.count)}</td>
          <td>${formatFieldNumber(metricField, row[valueKey])}</td>
        </tr>
      `).join("")}
    </tbody>
  `;
}

function renderV2TopGroupChart(metricField, groupField, grouped, aggregateMethod = "sum") {
  const canvas = document.getElementById("v2TopGroupChart");
  if (!window.Chart) return drawEmptyChart(canvas, Lpick(CHART_LIBRARY_MISSING));
  const aggregateLabel = getAggregateMethodLabel(aggregateMethod);
  const topGroups = [...grouped]
    .sort((a, b) => compareNumbersDescending(a[aggregateMethod], b[aggregateMethod]))
    .slice(0, 10);
  dom.v2TopGroupTitle.textContent = L(`${formatFieldLabel(metricField)} 按 ${groupField} 分组 Top 10（${aggregateLabel}）`, `${formatFieldLabel(metricField)} by ${groupField}, top 10 (${aggregateLabel})`);
  canvas.setAttribute("aria-label", dom.v2TopGroupTitle.textContent);
  dom.v2TopGroupTitle.title = dom.v2TopGroupTitle.textContent;
  state.charts.v2TopGroupChart = new Chart(canvas, {
    type: "bar",
    data: {
      labels: topGroups.map((item) => displayGroupName(item.name)),
      datasets: [{
        label: L(`${formatFieldLabel(metricField)} 按 ${groupField} 分组（${aggregateLabel}）`, `${formatFieldLabel(metricField)} by ${groupField} (${aggregateLabel})`),
        data: topGroups.map((item) => item[aggregateMethod]),
        backgroundColor: CHART_THEME.primary
      }]
    },
    options: chartOptions(numericAxisLabel(metricField, getAggregateMethodAxisLabel(aggregateMethod)), metricField)
  });
}

function renderV2TrendChart(metricField, dateField, aggregateMethod = "sum") {
  const canvas = document.getElementById("v2SelectedTrendChart");
  if (!window.Chart) return drawEmptyChart(canvas, Lpick(CHART_LIBRARY_MISSING));
  const trend = buildMetricDateTrend(dateField, metricField, aggregateMethod);
  if (!trend.labels.length) {
    return drawEmptyChart(canvas, L("所选时间字段没有可用的趋势数据", "The chosen date column has no trend data"));
  }

  const timeUnit = getDateGranularity(dateField) === "year"
    ? L("年份", "year")
    : trend.labels.some((label) => /^\d{4}-\d{2}$/.test(label)) ? L("月份", "month") : L("日期", "date");
  const aggregateLabel = getAggregateMethodLabel(aggregateMethod);
  dom.v2TrendTitle.textContent = L(`${formatFieldLabel(metricField)} 按${timeUnit}趋势（${aggregateLabel}）`, `${formatFieldLabel(metricField)} by ${timeUnit} (${aggregateLabel})`);
  canvas.setAttribute("aria-label", dom.v2TrendTitle.textContent);
  dom.v2TrendTitle.title = dom.v2TrendTitle.textContent;
  state.charts.v2SelectedTrendChart = new Chart(canvas, {
    type: "line",
    data: {
      labels: trend.labels,
      datasets: [{
        label: dom.v2TrendTitle.textContent,
        data: trend.values,
        ...trendLineStyle(CHART_THEME.secondary)
      }]
    },
    options: chartOptions(numericAxisLabel(metricField, getAggregateMethodAxisLabel(aggregateMethod)), metricField)
  });
}

function getAggregateMethodLabel(method) {
  return L({ sum: "求和", avg: "平均值", median: "中位数" }[method] || "求和", { sum: "sum", avg: "average", median: "median" }[method] || "sum");
}

// The y-axis title: the same Chinese, the English with a capital (the words above are for sentences).
function getAggregateMethodAxisLabel(method) {
  return L({ sum: "求和", avg: "平均值", median: "中位数" }[method] || "求和", { sum: "Sum", avg: "Average", median: "Median" }[method] || "Sum");
}

// v3.0.0: a trend is a straight line with white points outlined in the line's colour, no fill.
function trendLineStyle(color) {
  return {
    borderColor: color,
    borderWidth: 2.5,
    tension: 0,
    fill: false,
    pointRadius: 4,
    pointHoverRadius: 5,
    pointBackgroundColor: "#ffffff",
    pointBorderColor: color,
    pointBorderWidth: 2
  };
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
        ticks: { maxRotation: 0, minRotation: 0, color: CHART_THEME.text, autoSkip: true, autoSkipPadding: 16, maxTicksLimit: 9 },
        grid: { display: false },
        border: { display: false }
      },
      y: {
        beginAtZero: true,
        title: { display: true, text: axisLabel, color: CHART_THEME.text },
        ticks: {
          color: CHART_THEME.text,
          maxTicksLimit: 5,
          ...(valueField ? { callback: (value) => formatFieldNumber(valueField, Number(value)) } : {})
        },
        grid: { color: CHART_THEME.grid },
        border: { display: false }
      }
    }
  };
}

// v3.0.0 F3: the sentence is ordinary text in the static <p id="<canvas id>Empty"> over the
// chart frame, no longer painted on the canvas; destroyChart() clears it before a chart is drawn.
// Name and arguments stay: tests wrap this function to collect the sentences.
function drawEmptyChart(canvas, text) {
  const context = typeof canvas.getContext === "function" ? canvas.getContext("2d") : null;
  if (context) context.clearRect(0, 0, canvas.width, canvas.height);
  const empty = document.getElementById(`${canvas.id}Empty`);
  if (!empty) return;
  empty.textContent = text;
  empty.classList.remove("hidden");
}

function clearChartEmpty(id) {
  const empty = document.getElementById(`${id}Empty`);
  if (!empty) return;
  empty.textContent = "";
  empty.classList.add("hidden");
}

function resetAnalysisState() {
  Object.keys(state.charts).forEach((id) => destroyChart(id));
  ["numericDistributionChart", "categoryTopChart", "dateTrendChart", "v2TopGroupChart", "v2SelectedTrendChart"].forEach(clearChartEmpty);
  state.charts = {};
  state.rows = [];
  state.excludedRows = { suspectedTotal: [], duplicateRows: [] };
  state.exclusionSourceRows = null;
  state.includedRows = null;
  state.repeatedRowIndexes = null;
  state.repeatedRowSource = null;
  state.restoredTotalRows = new Set();
  state.lastExclusion = null;
  state.filteredRows = null;
  state.filterSourceRows = null;
  state.fields = [];
  state.profiles = [];
  state.numericStats = [];
  state.qualityNumericStats = [];
  state.categoryStats = [];
  state.duplicateRows = 0;
  state.duplicateRowGroups = [];
  state.checkupSummary = null;
  state.idDuplicates = [];
  state.placeholderConfirmations = Object.create(null);
  state.placeholders = [];
  state.numericLooking = [];
  state.totalLabelRows = [];
  state.suspectedTotals = [];
  state.suspectedTotalLabelCount = 0;
  state.formatScans = createFormatScans();
  state.dateStyles = [];
  state.constantFields = [];
  state.lookAlikes = [];
  state.sourceDateCells = new WeakMap();
  state.sourceFormulaErrors = new WeakMap();
  state.formulaErrors = [];
  state.manualFieldTypes = Object.create(null);
  state.totalMissing = 0;
  state.parseWarnings = [];
  state.numberFormatHints = Object.create(null);
  state.headerDetection = null;
  state.sourceRowNumbers = new WeakMap();
  state.previewPage = 1;
  state.problemCategories = [];
  state.fieldTypeDrafts = Object.create(null);
  state.pendingExcel = null;
  // v3.2.1: the kept bytes go with the analysis (another file, the sample, an import error);
  // switchExcelSheet() holds them across its own reset and gives them back.
  state.keptExcel = null;
  clearSheetSwitches();
  state.sourceFileName = "";
  state.sourceSheetName = "";
  state.sourceType = "";
  state.analysisCompletedAt = null;
  state.chartInsights = { category: "", trend: "" };
  state.trendInsightData = null;
  state.customAnalysis = null;
  state.filters = createEmptyFilterConfig();
  state.categoryFilterOptions = [];
  state.categoryFilterDraftValues = new Set();
  clearStandardizedExport();

  document.body.classList.remove("has-results");
  setDataExportStatus("");
  setReportExportStatus("");

  dom.sourceName.textContent = "";
  dom.sourceName.removeAttribute("title");
  dom.topbarFileName.textContent = "";
  dom.topbarFileName.removeAttribute("title");
  dom.sidebarFileSize.textContent = "";
  dom.sidebarFieldCount.textContent = "";
  dom.sidebarFieldList.innerHTML = "";
  dom.sidebarFieldList.classList.remove("hidden");
  dom.sidebarFieldsToggle.setAttribute("aria-expanded", "true");
  dom.sidebarHeaderRow.classList.add("hidden");
  dom.sidebarHeaderValue.textContent = "";
  dom.sidebarPlaceholderRow.classList.add("hidden");
  dom.sidebarPlaceholderValue.textContent = "";
  document.body.classList.remove("sidebar-open", "sidebar-closed");
  syncSidebarToggle();
  resetStatusNote();
  dom.statusNoteSummary.textContent = "";
  statusMessage = null;
  state.importStatus = null;
  state.returnModule = null;
  dom.returnToAnalysis.classList.add("hidden");
  writeLanguageText(dom.returnToAnalysis, "");
  writeLanguageText(dom.reportTitle, () => L("数据分析工作区", "Analysis workspace"));
  dom.overviewCards.innerHTML = "";
  dom.previewSize.textContent = "";
  renderCheckup();
  setChartInsight(dom.categoryChartInsight, "");
  setChartInsight(dom.dateChartInsight, "");
  dom.previewTable.innerHTML = "";
  dom.previewTable.classList.remove("is-empty-table");
  dom.previewPageInfo.textContent = "";
  setPagerButtons([dom.previewFirstPage, dom.previewPrevPage, dom.previewNextPage, dom.previewLastPage], 1, 1);
  dom.schemaTable.innerHTML = "";
  state.schemaFilter = { query: "", kind: "all" };
  dom.schemaSearch.value = "";
  dom.schemaKinds.innerHTML = "";
  dom.schemaFilter.classList.add("hidden");
  dom.schemaFilterEmpty.classList.add("hidden");
  dom.fieldConfigMessage.textContent = "";
  dom.fieldConfigMessage.className = "field-config-message";
  dom.applyFieldConfig.disabled = true;
  dom.restoreAutoFieldTypes.disabled = true;
  dom.fieldConfigToolbar.classList.remove("has-draft");
  dom.categoryFilterField.innerHTML = "";
  dom.categoryFilterSearch.value = "";
  dom.categoryFilterValues.innerHTML = "";
  dom.categoryFilterControls.classList.add("hidden");
  dom.categoryFilterEmpty.classList.remove("hidden");
  dom.categoryFilterSelectionCount.textContent = L("未选择字段", "No column chosen");
  dom.numericFilterField.innerHTML = "";
  dom.numericFilterMin.value = "";
  dom.numericFilterMax.value = "";
  writeLanguageText(dom.numericFilterRange, () => formatFilterRangeEmpty(false));
  dom.dateFilterField.innerHTML = "";
  dom.dateFilterStart.value = "";
  dom.dateFilterEnd.value = "";
  writeLanguageText(dom.dateFilterRange, () => formatFilterRangeEmpty(false));
  dom.filterOriginalCount.textContent = "0";
  dom.filterCurrentCount.textContent = "0";
  dom.filterExcludedCount.textContent = "0";
  setFilterMessage(() => formatNoFilterMessage(0, 0), "");
  dom.qualitySummary.innerHTML = "";
  dom.parseWarningSummary.innerHTML = "";
  dom.parseWarningSummary.classList.add("hidden");
  renderRowNoteSection(dom.suspectedTotalSection, dom.suspectedTotalDetails, []);
  dom.suspectedTotalSection.classList.remove("is-excluded");
  dom.suspectedTotalExcludedNote.textContent = "";
  dom.suspectedTotalExcludedNote.classList.add("hidden");
  setExclusionMessage("suspectedTotal", "");
  dom.qualityScopeNote.textContent = "";
  dom.qualityScopeNote.classList.add("hidden");
  dom.duplicateExclusion.classList.add("hidden");
  dom.duplicateExclusionText.textContent = "";
  dom.duplicateExclusionAction.innerHTML = "";
  setExclusionMessage("duplicateRows", "");
  dom.filterNotCountedItem.classList.add("hidden");
  dom.filterNotCountedCount.textContent = "0";
  [dom.statisticsContainsNote, dom.chartsContainsNote, dom.customContainsNote].forEach((element) => {
    element.innerHTML = "";
    element.classList.add("hidden");
  });
  renderRowNoteSection(dom.formulaErrorSection, dom.formulaErrorDetails, []);
  renderRowNoteSection(dom.dateStyleSection, dom.dateStyleDetails, []);
  renderNumberNoteSection(dom.constantFieldSection, dom.constantFieldDetails, []);
  renderRowNoteSection(dom.lookAlikeSection, dom.lookAlikeDetails, []);
  hideProblemView();
  dom.qualityTable.innerHTML = "";
  dom.outlierTable.innerHTML = "";
  dom.outlierTable.classList.remove("is-empty-table");
  dom.idDuplicateDetails.innerHTML = "";
  dom.idDuplicateDetails.classList.remove("is-empty");
  dom.placeholderDetails.innerHTML = "";
  dom.placeholderDetails.classList.remove("is-empty");
  setPlaceholderMessage("");
  dom.numericStatsTable.innerHTML = "";
  dom.numericStatsTable.classList.remove("is-empty-table");
  dom.categoryStats.innerHTML = "";
  dom.excelWorkbookName.textContent = "";
  dom.excelSheetSelect.innerHTML = "";
  writeOneOffMessage(dom.sheetSelectionMessage, "sheet-selection-message", "");
  dom.sheetSelectionPanel.classList.add("hidden");
  writeOneOffMessage(dom.v2FieldPrompt, "analysis-prompt", "");
  dom.v2SummaryCards.innerHTML = "";
  dom.v2ResultSummary.textContent = "";
  dom.v2PrimaryTable.innerHTML = "";
  dom.v2SumTable.innerHTML = "";
  dom.v2AvgTable.innerHTML = "";
  dom.v2AnalysisResults.classList.add("hidden");
  document.querySelectorAll("#customSection details").forEach((details) => {
    details.open = false;
  });
  showModule("overviewSection", { quiet: true });
  dom.results.classList.add("hidden");
  updateGlobalFilterStatus();
  updateExportButtons();
}

const CHART_LIBRARY_MISSING = Lp("图表组件加载失败，请刷新页面重试", "The chart library failed to load; refresh the page and try again");

function destroyChart(id) {
  clearChartEmpty(id);
  if (state.charts[id]) {
    state.charts[id].destroy();
    state.charts[id] = null;
  }
}

function setSelectOptions(select, values, placeholder = "", allowEmpty = false) {
  const placeholderOption = placeholder
    ? `<option value="">${escapeHtml(placeholder)}</option>`
    : "";
  const valueOptions = values.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`).join("");
  select.innerHTML = values.length || allowEmpty
    ? `${placeholderOption}${valueOptions}`
    : `<option value="">${L("无可用字段", "No columns available")}</option>`;
  select.disabled = values.length === 0 && !allowEmpty;
}

function emptyTable(message) {
  return `
    <thead><tr><th>${L("提示", "Note")}</th></tr></thead>
    <tbody><tr><td>${escapeHtml(message)}</td></tr></tbody>
  `;
}

// The import message from state.importStatus, in the current language: after the import, back
// from 换一个文件 (script.js) and after a language switch.
function buildImportStatus({ sourceDisplayName, warningCount, affectedRows, headerDetection, suspectedTotals, formulaErrors }) {
  const notes = [formatHeaderDetectionNote(headerDetection), formatSuspectedTotalStatus(suspectedTotals), formatFormulaErrorStatus(formulaErrors)].filter(Boolean);
  const headerNote = notes.join(L("", " "));
  const message = warningCount
    ? L(
      `CSV 已载入，但读取时有 ${formatInteger(warningCount)} 处疑问${affectedRows ? `，涉及 ${formatInteger(affectedRows)} 行` : ""}。请核对数据预览和字段数量。${headerNote}`,
      `The CSV is loaded, but there ${warningCount === 1 ? "is 1 reading issue" : `are ${formatInteger(warningCount)} reading issues`}${affectedRows ? ` in ${enCount(affectedRows, "row", "rows")}` : ""}. Check the data preview and the number of columns.${headerNote ? ` ${headerNote}` : ""}`
    )
    : L(`${sourceDisplayName} 已成功载入，分析结果已生成。${headerNote}`, `${sourceDisplayName} is loaded and the analysis is ready.${headerNote ? ` ${headerNote}` : ""}`);
  return {
    message,
    tone: warningCount ? "warning" : "success",
    hasNote: notes.length > 0,
    // v3.4.0: a reading warning is one note more, so N is what “看说明” opens.
    summary: notes.length || warningCount ? formatStatusNoteSummary(sourceDisplayName, notes.length + (warningCount ? 1 : 0)) : ""
  };
}

function showImportStatus(status) {
  const { tone, hasNote, summary } = buildImportStatus(status);
  setStatus(() => buildImportStatus(status).message, tone, hasNote, () => buildImportStatus(status).summary);
  dom.statusNoteSummary.textContent = summary;
  // v3.4.0: the import message starts folded to its summary, a warning too; #statusPanel keeps its
  // text, classes and role (a screen reader still reads it).
  if (summary) foldStatusNote();
}

// v3.2.0 commit 11: the status bar's message (and the folded summary of an import) as functions,
// so a language switch writes them again (rewriteStatusMessage()); null when the bar is empty.
let statusMessage = null;

function clearStatusMessage() {
  statusMessage = null;
  dom.statusPanel.textContent = "";
}

function rewriteStatusMessage() {
  if (!statusMessage) return;
  dom.statusPanel.textContent = statusMessage.build();
  if (statusMessage.summary) dom.statusNoteSummary.textContent = statusMessage.summary();
  dom.statusNoteToggle.textContent = dom.statusNote.classList.contains("is-collapsed") ? L("看说明", "Show notes") : L("收起", "Collapse");
}

// hasNote keeps a successful status on screen once results show (style.css): its message
// says something the page does not say at the top. message: a text or a function writing it.
function setStatus(message, tone = "", hasNote = false, summary = null) {
  const build = typeof message === "function" ? message : () => message;
  statusMessage = { build, summary };
  dom.statusPanel.className = `status-panel${tone ? ` ${tone}` : ""}${hasNote ? " has-note" : ""}`;
  dom.statusPanel.textContent = build();
  dom.statusPanel.setAttribute("role", tone === "error" ? "alert" : "status");
  dom.statusPanel.setAttribute("aria-busy", String(!tone));
  setImportBusy(!tone);
  // Only a success with notes can fold (v3.0.0); every new message starts unfolded, and
  // showImportStatus() folds the import message (v3.4.0, a warning too). The kind
  // of message is marked on #statusNote: style.css keeps successes to the first module, while
  // warnings and errors (the report export's among them) show on every module.
  resetStatusNote();
  dom.statusNote.classList.toggle("can-collapse", tone === "success" && hasNote);
  dom.statusNote.dataset.tone = tone || "loading";
}

// The fold of the import message (v3.0.0) lives on #statusNote, never on #statusPanel: its
// text, classes and role stay as setStatus() wrote them, and a folded message is only moved
// out of sight (style.css), so screen readers still have it.
function resetStatusNote() {
  dom.statusNote.classList.remove("can-collapse", "is-collapsed");
  dom.statusNoteToggle.setAttribute("aria-expanded", "true");
  dom.statusNoteToggle.textContent = L("收起", "Collapse");
}

function foldStatusNote() {
  dom.statusNote.classList.add("can-collapse", "is-collapsed");
  dom.statusNoteToggle.setAttribute("aria-expanded", "false");
  dom.statusNoteToggle.textContent = L("看说明", "Show notes");
}

function toggleStatusNote() {
  const collapse = !dom.statusNote.classList.contains("is-collapsed");
  dom.statusNote.classList.toggle("is-collapsed", collapse);
  dom.statusNoteToggle.setAttribute("aria-expanded", String(!collapse));
  dom.statusNoteToggle.textContent = collapse ? L("看说明", "Show notes") : L("收起", "Collapse");
}

// message: a text or a function writing it (an error built in both languages is written again
// after a language switch).
function showError(message, importId = state.activeImportId) {
  if (!isCurrentImport(importId)) return;
  resetAnalysisState();
  setImportBusy(false);
  const build = typeof message === "function" ? message : () => message;
  statusMessage = { build, summary: null };
  dom.statusPanel.className = "status-panel error";
  dom.statusPanel.textContent = build();
  dom.statusPanel.setAttribute("role", "alert");
  dom.statusPanel.setAttribute("aria-busy", "false");
  dom.statusNote.dataset.tone = "error";
  dom.results.classList.add("hidden");
}
