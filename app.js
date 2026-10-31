// 应用状态与界面：state 和 dom 映射、导入 / 筛选 / 字段配置三条刷新路径，以及概览（含分页的数据预览）、字段、质量（含问题行查看区）、统计、图表（含图表下的分类 / 趋势结论）和自定义分析的渲染。

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
  activeFileReader: null,
  activeImportId: 0,
  charts: {},
  sourceFileName: "",
  sourceSheetName: "",
  sourceType: "",
  analysisCompletedAt: null,
  chartInsights: { category: "", trend: "" },
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
  categoryChartInsight: document.getElementById("categoryChartInsight"),
  dateChartInsight: document.getElementById("dateChartInsight"),
  previewTable: document.getElementById("previewTable"),
  previewPageInfo: document.getElementById("previewPageInfo"),
  previewFirstPage: document.getElementById("previewFirstPage"),
  previewPrevPage: document.getElementById("previewPrevPage"),
  previewNextPage: document.getElementById("previewNextPage"),
  previewLastPage: document.getElementById("previewLastPage"),
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
  numberConversionSection: document.getElementById("numberConversionSection"),
  numberConversionDetails: document.getElementById("numberConversionDetails"),
  numericLookingSection: document.getElementById("numericLookingSection"),
  numericLookingDetails: document.getElementById("numericLookingDetails"),
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
  dataExportDeduplicatedCount: document.getElementById("dataExportDeduplicatedCount"),
  exportFilteredCsv: document.getElementById("exportFilteredCsv"),
  exportDeduplicatedCsv: document.getElementById("exportDeduplicatedCsv"),
  dataExportStatus: document.getElementById("dataExportStatus"),
  exportHtmlReport: document.getElementById("exportHtmlReport"),
  exportMarkdownReport: document.getElementById("exportMarkdownReport"),
  reportExportStatus: document.getElementById("reportExportStatus")
};

// A category field with more distinct values than this is too scattered for
// the automatic category conclusion.
const INSIGHT_CATEGORY_MAX_DISTINCT = 20;

const CHART_THEME = {
  primary: "#625bde",
  primaryFill: "rgba(98, 91, 222, 0.12)",
  secondary: "#0f9f8f",
  secondaryFill: "rgba(15, 159, 143, 0.12)",
  warning: "#e09a32",
  text: "#697084",
  grid: "rgba(211, 215, 227, 0.72)"
};

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
  window.requestAnimationFrame(() => {
    dom.reportTitle?.focus({ preventScroll: true });
    dom.results.scrollIntoView({ block: "start" });
  });
}

// headerDetection is { headerRowNumber, skippedRows: [{ rowNumber, cells }] } from the
// importer when title rows were skipped above the header, otherwise null.
// rowNumbers[i] is the file row of rows[i]; null when the import could not tell.
function commitTabularData(fields, rows, sourceName, importId, parseWarnings = [], sheetName = "", numberFormatHints = null, headerDetection = null, rowNumbers = null) {
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
  state.chartInsights = { category: "", trend: "" };
  state.customAnalysis = null;
  state.filters = createEmptyFilterConfig();
  state.categoryFilterOptions = [];
  state.categoryFilterDraftValues = new Set();
  state.previewPage = 1;
  state.problemView = null;
  state.problemViewOpener = null;
  // Source-file hints are replaced on every commit; only Excel imports pass them.
  state.numberFormatHints = Object.assign(Object.create(null), numberFormatHints);
  state.headerDetection = buildHeaderDetectionState(headerDetection);
  state.sourceRowNumbers = new WeakMap();
  if (Array.isArray(rowNumbers) && rowNumbers.length === rows.length) {
    rows.forEach((row, index) => state.sourceRowNumbers.set(row, rowNumbers[index]));
  }
  state.placeholderConfirmations = Object.create(null);
  state.manualFieldTypes = Object.create(null);
  state.profiles = fields.map((field) => buildColumnProfile(field, rows, state.numberFormatHints[field]));
  state.fieldTypeDrafts = Object.fromEntries(
    state.profiles.map((profile) => [profile.field, profile.inferredTypeKey])
  );
  state.duplicateRows = countDuplicateRows(rows, fields);
  state.idDuplicates = buildIdDuplicateSummary(rows, fields, state.profiles, getSourceRowNumber, state.placeholderConfirmations);
  state.placeholders = buildPlaceholderSummary(rows, fields, state.profiles, state.placeholderConfirmations, getSourceRowNumber);
  state.numericLooking = buildNumericLookingSummary(state.profiles);
  state.totalMissing = state.profiles.reduce((sum, profile) => sum + profile.missingCount, 0);
  // No filter yet: the analysis rows are the original rows, so one calculation serves both.
  // The two share one array; nothing may modify it in place.
  state.qualityNumericStats = buildNumericStats(state.rows);
  state.numericStats = state.qualityNumericStats;
  state.categoryStats = buildCategoryStats(state.filteredRows);

  const sourceDisplayName = sheetName ? `${sourceName} · ${sheetName}` : sourceName;
  dom.sourceName.textContent = `已载入 · ${sourceDisplayName}`;
  dom.sourceName.title = sourceDisplayName;
  dom.sheetSelectionPanel.classList.add("hidden");
  dom.results.classList.remove("hidden");
  const headerNote = formatHeaderDetectionNote(state.headerDetection);
  if (state.parseWarnings.length) {
    // Warnings without a line count in the total only.
    const affectedRows = new Set(state.parseWarnings.map((warning) => warning.row).filter((row) => row !== null)).size;
    setStatus(`CSV 已载入，但发现 ${formatInteger(state.parseWarnings.length)} 个解析警告${affectedRows ? `，涉及 ${formatInteger(affectedRows)} 行` : ""}。请核对数据预览和字段数量。${headerNote}`, "warning");
  } else {
    setStatus(`${sourceDisplayName} 已成功载入，分析结果已生成。${headerNote}`, "success");
  }

  renderOverview();
  renderPreview();
  renderSchema();
  initializeFilterControls();
  renderQuality();
  renderNumericStats();
  renderCategoryStats();
  populateControls();
  renderCharts();
  renderChartInsights();
  resetV2AnalysisState();
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

function renderOverview() {
  const analysisRows = getAnalysisRows();
  const activeProfiles = getActiveProfiles();
  const numericCount = getProfilesByType("numeric").length;
  const categoryCount = getProfilesByType("category").length;
  const activeKeys = activeProfiles.map((profile) => getConfirmedPlaceholderKeys(profile.field));
  const activeMissing = analysisRows.reduce((total, row) => (
    total + activeProfiles.filter((profile, index) => isMissingOrConfirmed(row[profile.field], activeKeys[index])).length
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
    return `分类字段「${field}」中占比最高的类别是「${top.name}」，出现 ${formatInteger(top.count)} 次，占比 ${formatPercent(top.ratio)}。`;
  }
  if (!item.top.length) return `「${field}」当前没有可统计的取值，不做概括。`;
  return `「${field}」取值比较分散（超过 ${INSIGHT_CATEGORY_MAX_DISTINCT} 个不同值，或没有重复出现的值），不做概括。`;
}

function renderCategoryInsight() {
  state.chartInsights.category = buildCategoryInsight(dom.categoryChartField.value);
  setChartInsight(dom.categoryChartInsight, state.chartInsights.category);
}

function renderTrendInsight() {
  const field = dom.dateChartField.value;
  state.chartInsights.trend = field && getAnalysisRows().length ? buildTrendInsight(field) : "";
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
      <tr><th>文件行号</th>${withNote ? "<th>说明</th>" : ""}${fields.map((field) => `<th>${escapeHtml(field)}</th>`).join("")}</tr>
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
  dom.previewPageInfo.textContent = `第 ${formatInteger(page)} / ${formatInteger(pageCount)} 页，共 ${formatInteger(rows.length)} 行`;
  setPagerButtons([dom.previewFirstPage, dom.previewPrevPage, dom.previewNextPage, dom.previewLastPage], page, pageCount);
  if (!rows.length) {
    dom.previewTable.innerHTML = emptyTable("当前筛选条件下没有可预览的数据。");
    return;
  }
  dom.previewTable.innerHTML = buildRowTableHtml(state.fields, rows.slice(start, end).map((row) => ({ row, line: getSourceRowNumber(row) })));
}

function goToPreviewPage(action) {
  const { page, pageCount } = getRowPage(getAnalysisRows().length, state.previewPage);
  state.previewPage = movePage(action, page, pageCount);
  renderPreview();
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
    profile.type = FIELD_TYPE_LABELS[typeKey];
    profile.conversionFailureCount = countFieldConversionFailures(profile.field, typeKey);
    profile.conversionFailureSamples = getFieldConversionFailureSamples(profile.field, typeKey);
  });
  state.manualFieldTypes = manualFieldTypes;
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
    profile.conversionFailureSamples = getFieldConversionFailureSamples(profile.field, profile.inferredTypeKey);
  });
  state.fieldTypeDrafts = Object.fromEntries(
    state.profiles.map((profile) => [profile.field, profile.inferredTypeKey])
  );
  state.manualFieldTypes = Object.create(null);
  rebuildAnalysisFromProfiles("已恢复自动识别，统计、图表和分析结论已重新计算。");
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
  dom.restoreAutoFieldTypes.disabled = !hasAutomaticTypeDifferences();
  // The problem rows follow the applied configuration; say so while drafts wait.
  dom.problemRowsDraftNote.classList.toggle("hidden", !state.problemView || !hasFieldConfigDraftChanges());
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

function setPlaceholderMessage(message, tone = "") {
  dom.placeholderMessage.className = `placeholder-message${tone ? ` ${tone}` : ""}`;
  dom.placeholderMessage.textContent = message;
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

function confirmPlaceholder(field, key) {
  const pending = state.placeholders.some((entry) => entry.field === field && entry.key === key && !entry.confirmed && !entry.fieldIgnored);
  const keysByField = Object.create(null);
  if (pending) keysByField[field] = [...getConfirmedPlaceholderKeys(field), key];
  return updatePlaceholderConfirmations(keysByField, `已把「${field}」的「${key}」按缺失处理，统计、图表和分析结论已重新计算。`);
}

function undoPlaceholderConfirmation(field, key) {
  const keys = getConfirmedPlaceholderKeys(field);
  const keysByField = Object.create(null);
  if (keys.includes(key)) keysByField[field] = keys.filter((item) => item !== key);
  return updatePlaceholderConfirmations(keysByField, `已撤销「${field}」的「${key}」，统计、图表和分析结论已重新计算。`);
}

function confirmAllCommonPlaceholders() {
  const entries = getPendingCommonPlaceholders();
  const keysByField = Object.create(null);
  entries.forEach((entry) => {
    keysByField[entry.field] ??= [...getConfirmedPlaceholderKeys(entry.field)];
    keysByField[entry.field].push(entry.key);
  });
  return updatePlaceholderConfirmations(keysByField, `已把 ${formatInteger(entries.length)} 处常见写法按缺失处理，统计、图表和分析结论已重新计算。`);
}

// The fourth refresh path. keysByField replaces the confirmed keys of each listed
// field; refused while the field configuration has unapplied drafts.
function updatePlaceholderConfirmations(keysByField, message) {
  if (!state.rows.length) return false;
  if (!ensureFieldConfigurationApplied("处理疑似缺失值")) {
    setPlaceholderMessage("字段配置有未应用的修改。请先点击“应用字段配置”，再处理疑似缺失值。", "warning");
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
    profile.type = FIELD_TYPE_LABELS[manualTypeKey];
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

function getAnalysisRows() {
  return state.filterSourceRows === state.rows && Array.isArray(state.filteredRows)
    ? state.filteredRows
    : state.rows;
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
  // Confirmed placeholders join the empty value, as in filterRowsByConfig().
  const confirmedKeys = getConfirmedPlaceholderKeys(field);
  const counts = new Map();
  state.rows.forEach((row) => {
    const value = analysisValue(row[field], confirmedKeys);
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
// placeholder subsection for placeholder confirmations.
function rebuildAnalysisFromProfiles(fieldConfigMessage = "字段配置已更新。", setMessage = setFieldConfigMessage) {
  state.idDuplicates = buildIdDuplicateSummary(state.rows, state.fields, state.profiles, getSourceRowNumber, state.placeholderConfirmations);
  state.placeholders = buildPlaceholderSummary(state.rows, state.fields, state.profiles, state.placeholderConfirmations, getSourceRowNumber);
  state.numericLooking = buildNumericLookingSummary(state.profiles);
  state.analysisCompletedAt = null;
  updateExportButtons();
  resetFilterState();
  setDataExportStatus("");
  // The filter was just reset, so the analysis rows are the original rows: one calculation
  // serves both, shared as in commitTabularData().
  state.qualityNumericStats = buildNumericStats(state.rows);
  state.numericStats = state.qualityNumericStats;
  state.categoryStats = buildCategoryStats(getAnalysisRows());
  state.previewPage = 1;
  renderOverview();
  renderPreview();
  renderSchema();
  initializeFilterControls();
  renderQuality();
  renderNumericStats();
  renderCategoryStats();
  populateControls({ keepChartSelections: true });
  renderCharts();
  renderChartInsights();
  resetV2AnalysisState();
  setMessage(`${fieldConfigMessage} 已清除现有筛选条件。`, "success");
  state.analysisCompletedAt = new Date();
  updateFieldConfigActions();
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
    <span class="quality-pill ${missingClass}">缺失值总数：${formatInteger(state.totalMissing)}</span>
    <span class="quality-pill ${duplicateClass}">重复行：${formatInteger(state.duplicateRows)}</span>
    <span class="quality-pill ${conversionFailureClass}">格式问题：${formatInteger(conversionFailureTotal)} 个值</span>
    <span class="quality-pill">异常值字段：${formatInteger(qualityNumericStats.filter((item) => item.outlierCount > 0).length)}</span>
    <span class="quality-pill ${listedIdFieldCount ? "warning" : "success"}">ID 重复字段：${formatInteger(listedIdFieldCount)}</span>
    <span class="quality-pill ${pendingPlaceholderFieldCount ? "warning" : "success"}">疑似缺失值：${formatInteger(pendingPlaceholderFieldCount)} 个字段</span>
    ${problemViewButton({ kind: "duplicateRows" }, "查看重复行")}
  `;
  renderParseWarningSummary();
  // Before the outlier table, which returns early when there is no numeric field.
  refreshProblemView();
  renderNumberNoteSection(dom.numberConversionSection, dom.numberConversionDetails, buildNumberConversionSummary(state.profiles).map(formatNumberConversion));
  renderNumberNoteSection(dom.numericLookingSection, dom.numericLookingDetails, state.numericLooking.map(formatNumericLookingHint));
  renderIdDuplicateDetails();
  renderPlaceholderDetails();

  dom.qualityTable.innerHTML = `
    <thead>
      <tr>
        <th>字段</th>
        <th>类型</th>
        <th>缺失值数量</th>
        <th>缺失率</th>
        <th>无法转换</th>
        <th>查看</th>
      </tr>
    </thead>
    <tbody>
      ${state.profiles.map((profile) => `
        <tr>
          <td>${escapeHtml(profile.field)}</td>
          <td>${escapeHtml(profile.type)}</td>
          <td>${formatInteger(profile.missingCount)}</td>
          <td>${formatPercent(profile.missingRate)}</td>
          <td${profile.conversionFailureCount ? ' class="cell-warning"' : ""}>${escapeHtml(formatConversionFailures(profile.conversionFailureCount, profile.conversionFailureSamples, formatInteger(profile.conversionFailureCount)))}</td>
          <td>${problemViewCell([
            problemViewButton({ kind: "missing", field: profile.field }, "缺失"),
            problemViewButton({ kind: "conversion", field: profile.field }, "格式问题")
          ])}</td>
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
        <th>查看</th>
      </tr>
    </thead>
    <tbody>
      ${qualityNumericStats.map((stat) => `
        <tr>
          <td>${escapeHtml(formatFieldLabel(stat.field))}</td>
          <td>${formatInteger(stat.outlierCount)}</td>
          <td>${formatFieldNumber(stat.field, stat.lowerBound)}</td>
          <td>${formatFieldNumber(stat.field, stat.upperBound)}</td>
          <td>${problemViewCell([problemViewButton({ kind: "outlier", field: stat.field }, "查看")])}</td>
        </tr>
      `).join("")}
    </tbody>
  `;
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
  return html || `<span class="problem-view-none">无</span>`;
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
  dom.parseWarningSummary.innerHTML = `${escapeHtml(`CSV 解析警告：${formatInteger(count)} 条${rowCount ? `，涉及 ${formatInteger(rowCount)} 行` : ""}`)} ${problemViewButton({ kind: "parseWarnings" }, "查看")}`;
}

// Delegated click handler of #qualitySection for the 查看 buttons, bound in script.js.
function handleProblemViewClick(event) {
  const button = event.target.closest("button[data-problem-index]");
  if (!button) return;
  const category = state.problemCategories[Number(button.dataset.problemIndex)];
  if (!category || !openProblemView(category)) return;
  state.problemViewOpener = button;
  dom.problemRowsTitle.focus();
}

// Opens the viewer on a category of state.problemCategories at page 1. Returns false,
// leaving the viewer as it was, when the category has nothing to show.
function openProblemView(category) {
  const match = state.problemCategories.find((item) => isSameProblemCategory(item, category));
  const list = match ? listProblemRows(match) : null;
  if (!list) return false;
  state.problemView = { category: { kind: match.kind, field: match.field, key: match.key }, page: 1, list };
  dom.problemRowsMessage.textContent = "";
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
const HIGHLIGHTED_PROBLEM_KINDS = new Set(["missing", "conversion", "outlier", "placeholder", "idDuplicate"]);

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
  dom.problemRowsTable.innerHTML = buildRowTableHtml(state.fields, view.list.items.slice(start, end).map((item) => ({
    row: item.index === null ? null : state.rows[item.index],
    line: item.line,
    note: item.note,
    highlightField
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

function hideProblemView() {
  state.problemView = null;
  state.problemViewOpener = null;
  dom.problemRowsSection.classList.add("hidden");
  dom.problemRowsCategory.innerHTML = "";
  dom.problemRowsSummary.textContent = "";
  dom.problemRowsTable.innerHTML = "";
  dom.problemRowsDraftNote.classList.add("hidden");
  dom.problemRowsMessage.textContent = "";
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
    dom.problemRowsMessage.textContent = "这一类已经没有问题行。";
  }
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
    dom.idDuplicateDetails.innerHTML = `<p class="id-duplicate-note">未识别到 ID 字段。</p>`;
    return;
  }
  const reported = state.idDuplicates.filter((item) => item.mostlyRepeated || item.values.length);
  const cleanFields = state.idDuplicates.filter((item) => !reported.includes(item)).map((item) => item.field);
  dom.idDuplicateDetails.innerHTML = [
    ...reported.map((item) => {
      if (item.mostlyRepeated) {
        return `<p class="id-duplicate-note warning">${escapeHtml(formatMostlyRepeatedIdHint(item))}</p>`;
      }
      const overflow = formatIdDuplicateOverflow(item.duplicateValueCount, item.values.length);
      return `
        <div class="id-duplicate-field">
          <p class="id-duplicate-field-name">${escapeHtml(item.field)} ${problemViewButton({ kind: "idDuplicate", field: item.field }, "查看")}</p>
          <div class="table-wrap" tabindex="0" aria-label="${escapeHtml(item.field)} 的重复值表，可横向滚动">
            <table aria-label="${escapeHtml(item.field)} 的重复值">
              <thead><tr><th>重复值</th><th>出现次数</th><th>文件行号</th><th>说明</th></tr></thead>
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
    dom.placeholderDetails.innerHTML = `<p class="placeholder-note">没有发现疑似缺失值。</p>`;
    return;
  }
  const commonCount = getPendingCommonPlaceholders().length;
  const typeChanges = commonCount ? previewAllCommonPlaceholders() : [];
  const items = state.placeholders.map((entry, index) => {
    const textId = `placeholderText${index}`;
    const profile = getFieldProfile(entry.field);
    // Its own data-problem-index: handlePlaceholderClick() only reads data-placeholder-*.
    const viewButton = problemViewButton({ kind: "placeholder", field: entry.field, key: entry.key }, "查看");
    let item;
    if (entry.confirmed) {
      item = `<li class="placeholder-item confirmed"><p id="${textId}" class="placeholder-text">${escapeHtml(formatConfirmedPlaceholder(entry))}</p><button type="button" class="secondary-button placeholder-button" data-placeholder-index="${index}" aria-describedby="${textId}">撤销</button>${viewButton}</li>`;
    } else {
      const change = formatPlaceholderChange(entry.field, profile.missingCount, profile.typeKey, previewPlaceholderConfirmation(entry.field, [entry.key]));
      const notes = formatPlaceholderNotes([entry.key]);
      item = `<li class="placeholder-item"><p id="${textId}" class="placeholder-text">${escapeHtml(formatPendingPlaceholder(entry, change))}</p>${notes ? `<p class="placeholder-note warning">${escapeHtml(notes)}</p>` : ""}<button type="button" class="placeholder-button" data-placeholder-index="${index}" aria-describedby="${textId}">按缺失处理</button>${viewButton}</li>`;
    }
    if (state.placeholders[index + 1]?.field === entry.field) return item;
    const pendingKeys = state.placeholders.filter((other) => other.field === entry.field && !other.confirmed).map((other) => other.key);
    if (pendingKeys.length < 2) return item;
    const change = formatPlaceholderChange(entry.field, profile.missingCount, profile.typeKey, previewPlaceholderConfirmation(entry.field, pendingKeys));
    return `${item}<li class="placeholder-field-summary">${escapeHtml(`「${entry.field}」全部确认后：${change}。`)}</li>`;
  });
  const typeChangeNote = typeChanges.length
    ? `<span class="placeholder-note">${escapeHtml(`确认后类型会变：${typeChanges.map((change) => `「${change.field}」`).join("")}`)}</span>`
    : "";
  dom.placeholderDetails.innerHTML = `
    ${commonCount ? `<div class="placeholder-actions"><button type="button" class="placeholder-button" data-placeholder-all>全部确认常见写法（${formatInteger(commonCount)} 处）</button>${typeChangeNote}</div>` : ""}
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
    dom.numericStatsTable.innerHTML = emptyTable("未识别到数值字段。");
    return;
  }

  dom.numericStatsTable.innerHTML = `
    <thead>
      <tr>
        <th>字段</th>
        <th>有效数量</th>
        <th>无法转换</th>
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
          <td>${escapeHtml(formatFieldLabel(stat.field))}</td>
          <td>${formatInteger(stat.count)}</td>
          <td${stat.invalidCount ? ' class="cell-warning"' : ""}>${formatInteger(stat.invalidCount)}</td>
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
  setSelectOptions(dom.dateChartMetric, numericFields, "记录数", true);
  dom.dateChartMetric.value = keepChartSelections && numericFields.includes(previousDateChartMetric)
    ? previousDateChartMetric
    : "";
  dom.dateChartAggregate.value = keepChartSelections && previousDateChartAggregate === "sum" ? "sum" : "avg";
  setSelectOptions(dom.v2MetricField, numericProfiles.map((profile) => profile.field), "请选择数值字段");
  setSelectOptions(dom.v2GroupField, categoryProfiles.map((profile) => profile.field), "请选择分类字段");
  setSelectOptions(dom.v2DateField, dateProfiles.map((profile) => profile.field), "不选择时间字段", true);
}

function renderCharts() {
  if (!window.Chart) {
    ["numericDistributionChart", "categoryTopChart", "dateTrendChart"].forEach((id) => {
      drawEmptyChart(document.getElementById(id), "图表组件加载失败，请刷新页面重试");
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
  dom.numericChartTitle.textContent = field ? `${formatFieldLabel(field)} 数值分布（记录频数）` : "数值字段分布（记录频数）";
  canvas.setAttribute("aria-label", dom.numericChartTitle.textContent);
  if (!field) return drawEmptyChart(canvas, "没有可用的数值字段");

  const values = getNumericValues(field);
  if (!values.length) return drawEmptyChart(canvas, "该字段没有可安全转换的有效数值");
  const bins = buildDiscreteValueCounts(values, field) || buildHistogram(values, 10, field);
  state.charts.numericDistributionChart = new Chart(canvas, {
    type: "bar",
    data: {
      labels: bins.map((bin) => bin.label),
      datasets: [{
        label: `${formatFieldLabel(field)} 分布（记录频数）`,
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
  dom.dateChartTitle.textContent = "日期字段趋势（记录数）";
  canvas.setAttribute("aria-label", dom.dateChartTitle.textContent);
  const { metricField: numericField, aggregate } = getDateChartSelection();
  dom.dateChartAggregate.disabled = !numericField;
  if (!field) return drawEmptyChart(canvas, "没有可用的日期字段");

  const trend = buildDateChartTrend(field, numericField, aggregate);
  if (!trend.labels.length) return drawEmptyChart(canvas, "该字段没有可安全转换的有效日期");
  const timeUnit = getDateGranularity(field) === "year"
    ? "年份"
    : trend.labels.some((label) => /^\d{4}-\d{2}$/.test(label)) ? "月份" : "日期";
  const aggregationLabel = getDateChartAggregateLabel(aggregate);
  const yearly = getDateGranularity(field) === "year";
  dom.dateChartTitle.textContent = numericField
    ? `${formatFieldLabel(numericField)} 按${timeUnit}趋势（${aggregationLabel}）`
    : `${field} 按${timeUnit}趋势（${aggregationLabel}）`;
  canvas.setAttribute("aria-label", dom.dateChartTitle.textContent);
  state.charts.dateTrendChart = new Chart(canvas, {
    type: "line",
    data: {
      labels: trend.labels,
      datasets: [{
        label: numericField
          ? `${formatFieldLabel(numericField)}${yearly ? "年度" : " "}${aggregationLabel}趋势`
          : `${yearly ? "年度" : ""}记录数趋势`,
        data: trend.values,
        borderColor: CHART_THEME.primary,
        backgroundColor: CHART_THEME.primaryFill,
        tension: 0.25,
        fill: true
      }]
    },
    options: chartOptions(numericAxisLabel(numericField, aggregationLabel), numericField)
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
  return { count: "记录数", avg: "平均值", sum: "合计" }[aggregate] || "记录数";
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

  renderV2SummaryCards(metricField, groupField, aggregateMethod, dateField, grouped);
  renderV2AggregationTables(metricField, groupField, grouped);
  renderV2PrimaryTable(metricField, groupField, grouped, aggregateMethod);
  renderV2TopGroupChart(metricField, groupField, grouped, aggregateMethod);
  const invalidMetricText = metricInvalidCount
    ? `指标字段「${metricField}」有 ${formatInteger(metricInvalidCount)} 个值无法转换为数字，已从计算中排除。`
    : "";
  dom.v2ResultSummary.textContent = (hasActiveFilters()
    ? `已按 ${groupField} 对 ${metricField} 进行${aggregateLabel}，结果基于当前筛选数据 ${formatInteger(getAnalysisRows().length)} / ${formatInteger(state.rows.length)} 行。`
    : `已按 ${groupField} 对 ${metricField} 进行${aggregateLabel}，结果基于完整数据 ${formatInteger(state.rows.length)} 行。`) + invalidMetricText;

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

function renderV2SummaryCards(metricField, groupField, aggregateMethod, dateField, grouped) {
  const validRecordCount = grouped.reduce((total, item) => total + item.count, 0);
  const trendPointCount = dateField ? buildMetricDateTrend(dateField, metricField, aggregateMethod).labels.length : 0;
  const cards = [
    ["分析指标", formatFieldLabel(metricField)],
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
  dom.v2PrimaryTableTitle.textContent = `${formatFieldLabel(metricField)} 按 ${groupField} 分组（${aggregateLabel}）`;
  dom.v2PrimaryTable.innerHTML = buildV2AggregationTableHtml(groupField, metricField, sorted, aggregateMethod);
}

function buildV2AggregationTableHtml(groupField, metricField, rows, valueKey) {
  const valueLabel = { sum: "求和值", avg: "平均值", median: "中位数" }[valueKey] || "汇总值";
  return `
    <thead>
      <tr>
        <th>${escapeHtml(groupField)}</th>
        <th>记录数</th>
        <th>${escapeHtml(formatFieldLabel(metricField))} ${valueLabel}</th>
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
  if (!window.Chart) return drawEmptyChart(canvas, "图表组件加载失败，请刷新页面重试");
  const aggregateLabel = getAggregateMethodLabel(aggregateMethod);
  const topGroups = [...grouped]
    .sort((a, b) => compareNumbersDescending(a[aggregateMethod], b[aggregateMethod]))
    .slice(0, 10);
  dom.v2TopGroupTitle.textContent = `${formatFieldLabel(metricField)} 按 ${groupField} 分组 Top 10（${aggregateLabel}）`;
  canvas.setAttribute("aria-label", dom.v2TopGroupTitle.textContent);
  state.charts.v2TopGroupChart = new Chart(canvas, {
    type: "bar",
    data: {
      labels: topGroups.map((item) => item.name),
      datasets: [{
        label: `${formatFieldLabel(metricField)} 按 ${groupField} 分组（${aggregateLabel}）`,
        data: topGroups.map((item) => item[aggregateMethod]),
        backgroundColor: CHART_THEME.primary
      }]
    },
    options: chartOptions(numericAxisLabel(metricField, aggregateLabel), metricField)
  });
}

function renderV2TrendChart(metricField, dateField, aggregateMethod = "sum") {
  const canvas = document.getElementById("v2SelectedTrendChart");
  if (!window.Chart) return drawEmptyChart(canvas, "图表组件加载失败，请刷新页面重试");
  const trend = buildMetricDateTrend(dateField, metricField, aggregateMethod);
  if (!trend.labels.length) {
    return drawEmptyChart(canvas, "所选时间字段没有可用的趋势数据");
  }

  const timeUnit = getDateGranularity(dateField) === "year"
    ? "年份"
    : trend.labels.some((label) => /^\d{4}-\d{2}$/.test(label)) ? "月份" : "日期";
  const aggregateLabel = getAggregateMethodLabel(aggregateMethod);
  dom.v2TrendTitle.textContent = `${formatFieldLabel(metricField)} 按${timeUnit}趋势（${aggregateLabel}）`;
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

function getAggregateMethodLabel(method) {
  return { sum: "求和", avg: "平均值", median: "中位数" }[method] || "求和";
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
  state.idDuplicates = [];
  state.placeholderConfirmations = Object.create(null);
  state.placeholders = [];
  state.numericLooking = [];
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
  state.sourceFileName = "";
  state.sourceSheetName = "";
  state.sourceType = "";
  state.analysisCompletedAt = null;
  state.chartInsights = { category: "", trend: "" };
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
  setChartInsight(dom.categoryChartInsight, "");
  setChartInsight(dom.dateChartInsight, "");
  dom.previewTable.innerHTML = "";
  dom.previewPageInfo.textContent = "";
  setPagerButtons([dom.previewFirstPage, dom.previewPrevPage, dom.previewNextPage, dom.previewLastPage], 1, 1);
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
  dom.parseWarningSummary.innerHTML = "";
  dom.parseWarningSummary.classList.add("hidden");
  hideProblemView();
  dom.qualityTable.innerHTML = "";
  dom.outlierTable.innerHTML = "";
  dom.idDuplicateDetails.innerHTML = "";
  dom.placeholderDetails.innerHTML = "";
  setPlaceholderMessage("");
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
  dom.v2AnalysisResults.classList.add("hidden");
  document.querySelectorAll("#customSection details").forEach((details) => {
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

function emptyTable(message) {
  return `
    <thead><tr><th>提示</th></tr></thead>
    <tbody><tr><td>${escapeHtml(message)}</td></tr></tbody>
  `;
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
