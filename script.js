// 启动入口（最后加载）：启动自检、Chart.js 全局样式、页面导航和全部事件绑定。所有加载时立即执行的代码都只放在这里。

// The eight modules in page order (v3.0.0). Declared before the load-time calls below use it.
const MODULE_IDS = ["overviewSection", "schemaSection", "qualitySection", "filterSection", "statisticsSection", "chartsSection", "customSection", "exportSection"];

ensureDomElementsPresent();

// A cached older index.html can lack elements this script expects. Stop here
// with a refresh notice (inline styles, because style.css may be stale too)
// instead of failing later on the first addEventListener.
// The same notice covers app scripts that did not load: an old cached index.html
// loads only script.js, and a file that failed to load leaves its names
// undeclared. typeof is safe on undeclared names, and dom itself lives in app.js.
function ensureDomElementsPresent() {
  const missingScripts = [
    ["parsing.js", typeof toNumber],
    ["profiling.js", typeof buildColumnProfile],
    ["format.js", typeof formatNumber],
    ["stats.js", typeof buildNumericStats],
    ["quality.js", typeof countDuplicateRows],
    ["app.js", typeof commitTabularData],
    ["importer.js", typeof parseDataFile],
    ["report.js", typeof buildReportData]
  ].filter(([, type]) => type !== "function").map(([file]) => file);
  const missingKeys = missingScripts.length
    ? missingScripts
    : Object.keys(dom).filter((key) => !dom[key]);
  if (!missingKeys.length) return;

  const notice = document.createElement("div");
  notice.setAttribute("role", "alert");
  notice.style.cssText = "position:sticky;top:0;z-index:2147483647;display:flex;flex-wrap:wrap;align-items:center;justify-content:center;gap:12px;padding:12px 16px;border-bottom:1px solid #fedf89;background:#fffaeb;color:#b54708;font:600 14px/1.5 system-ui,-apple-system,\"Segoe UI\",\"Microsoft YaHei\",sans-serif;";
  const message = document.createElement("span");
  message.textContent = "页面已更新，请刷新后再使用。";
  const button = document.createElement("button");
  button.type = "button";
  button.textContent = "刷新页面";
  button.style.cssText = "padding:6px 14px;border:1px solid #b54708;border-radius:6px;background:#b54708;color:#fff;font:inherit;cursor:pointer;";
  button.addEventListener("click", () => window.location.reload());
  notice.append(message, button);
  document.body.prepend(notice);
  throw new Error(missingScripts.length
    ? `脚本文件缺失，HTML 与脚本版本不一致：${missingKeys.join(", ")}`
    : `页面元素缺失，HTML 与 script.js 版本不一致：${missingKeys.join(", ")}`);
}

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
    setStatus("正在加载 sample-orders.csv ...");
    const response = await fetch("examples/sample-orders.csv");
    if (!response.ok) throw new Error("示例数据加载失败");
    const text = await response.text();
    if (!isCurrentImport(importId)) return;
    parseCsvText(text, "sample-orders.csv", importId);
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
  dom.dateChartMetric,
  dom.dateChartAggregate
].forEach((control) => control.addEventListener("change", renderCharts));

dom.categoryChartField.addEventListener("change", renderCategoryInsight);

dom.dateChartField.addEventListener("change", renderTrendInsight);

dom.v2AnalyzeButton.addEventListener("click", renderV2Analysis);

dom.globalClearFilters.addEventListener("click", clearAllFilters);

// The other three 清除筛选 by the module titles carry no id (v3.0.0).
document.querySelectorAll(".filter-scope-clear").forEach((button) => button.addEventListener("click", clearAllFilters));

dom.statusNoteToggle.addEventListener("click", toggleStatusNote);

dom.sidebarToggle.addEventListener("click", toggleSidebar);
dom.sidebarClose.addEventListener("click", () => closeSidebar());
dom.sidebarBackdrop.addEventListener("click", () => closeSidebar());
dom.sidebarFieldsToggle.addEventListener("click", toggleSidebarFields);
dom.appSidebar.addEventListener("keydown", handleSidebarKeydown);
if (typeof window.matchMedia === "function") window.matchMedia(SIDEBAR_WIDE_QUERY).addEventListener("change", syncSidebarToggle);

dom.exportFilteredCsv.addEventListener("click", () => exportProcessedCsv("filtered"));

dom.exportDeduplicatedCsv.addEventListener("click", () => exportProcessedCsv("deduplicated"));
dom.checkStandardizedExport.addEventListener("click", checkStandardizedExport);
dom.exportStandardizedCsv.addEventListener("click", () => exportStandardizedData("data"));
dom.exportStandardizedChanges.addEventListener("click", () => exportStandardizedData("changes"));

dom.exportHtmlReport.addEventListener("click", exportHtmlReport);

dom.exportMarkdownReport.addEventListener("click", exportMarkdownReport);

dom.confirmSheetSelection.addEventListener("click", analyzeSelectedExcelSheet);

dom.excelSheetSelect.addEventListener("change", () => {
  dom.sheetSelectionMessage.className = "sheet-selection-message";
  dom.sheetSelectionMessage.textContent = "";
});

dom.previewFirstPage.addEventListener("click", () => goToPreviewPage("first"));

dom.previewPrevPage.addEventListener("click", () => goToPreviewPage("prev"));

dom.previewNextPage.addEventListener("click", () => goToPreviewPage("next"));

dom.previewLastPage.addEventListener("click", () => goToPreviewPage("last"));

dom.schemaTable.addEventListener("change", (event) => {
  const select = event.target.closest("select[data-field-type]");
  if (!select) return;
  stageFieldTypeChange(select.dataset.fieldType, select.value);
});

dom.applyFieldConfig.addEventListener("click", applyFieldConfiguration);

dom.schemaSearch.addEventListener("input", handleSchemaSearch);

dom.schemaKinds.addEventListener("click", handleSchemaKindClick);

dom.restoreAutoFieldTypes.addEventListener("click", restoreAutomaticFieldTypes);

dom.placeholderDetails.addEventListener("click", handlePlaceholderClick);

dom.qualitySection.addEventListener("click", handleProblemViewClick);

dom.problemRowsCategory.addEventListener("change", () => selectProblemCategory(dom.problemRowsCategory.value));

dom.problemRowsClose.addEventListener("click", closeProblemView);

dom.problemRowsExport.addEventListener("click", exportProblemRows);

dom.problemRowsFirstPage.addEventListener("click", () => goToProblemRowsPage("first"));

dom.problemRowsPrevPage.addEventListener("click", () => goToProblemRowsPage("prev"));

dom.problemRowsNextPage.addEventListener("click", () => goToProblemRowsPage("next"));

dom.problemRowsLastPage.addEventListener("click", () => goToProblemRowsPage("last"));

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

// v3.0.0: one module at a time. The others get .module-off (visibility: hidden, laid out in
// place, see style.css) and inert, never display: none or .hidden: a chart drawn while its
// module is hidden must keep its size, and captureCurrentChartImages() skips anything inside
// .hidden. Switching only shows and hides; it recalculates and redraws nothing.
// quiet: only the classes (resets). hash: false keeps the address (imports).
function showModule(sectionId, { quiet = false, hash = true } = {}) {
  if (!MODULE_IDS.includes(sectionId)) return;
  // style.css shows the import message only on the first module.
  document.body.dataset.module = sectionId;
  MODULE_IDS.forEach((id) => {
    const section = document.getElementById(id);
    const isActive = id === sectionId;
    section.classList.toggle("module-off", !isActive);
    section.inert = !isActive;
  });
  document.querySelectorAll("#moduleNav a").forEach((link) => {
    const isActive = link.getAttribute("href") === `#${sectionId}`;
    link.classList.toggle("active", isActive);
    if (isActive) link.setAttribute("aria-current", "page");
    else link.removeAttribute("aria-current");
  });
  if (quiet) return;
  // A narrow navigation scrolls sideways; bring the current link into view.
  document.querySelector?.(`#moduleNav a[href="#${sectionId}"]`)?.scrollIntoView?.({ block: "nearest", inline: "nearest" });
  if (hash) window.history.replaceState(null, "", `#${sectionId}`);
  window.scrollTo({ top: 0, behavior: "instant" });
  document.getElementById(`${sectionId}Title`).focus({ preventScroll: true });
}

function setupResultNavigation() {
  document.querySelectorAll("#moduleNav a").forEach((link) => {
    link.addEventListener("click", (event) => {
      event.preventDefault();
      showModule(link.getAttribute("href").slice(1));
    });
  });
  showModule(MODULE_IDS[0], { quiet: true });
}

function setupWorkspaceNavigation() {
  // 换一个文件: the top bar's, and the copy in the phone sidebar.
  document.querySelectorAll(".nav-new-analysis").forEach((link) => {
    link.addEventListener("click", (event) => {
      event.preventDefault();
      showUploadPage();
    });
  });
  dom.returnToAnalysis.addEventListener("click", returnToAnalysis);
}

// 换一个文件 (v3.0.0): back to the upload page. The analysis stays in memory until a new import
// begins (beginImport() resets it), so the upload page offers the way back to it; the focus goes
// to that entry, or to the file input when there is nothing to go back to.
function showUploadPage() {
  const canReturn = state.rows.length > 0 && Boolean(state.importStatus) && !state.pendingExcel
    && !document.body.classList.contains("is-loading");
  state.returnModule = document.body.dataset.module || MODULE_IDS[0];
  if (!isSidebarWide()) document.body.classList.remove("sidebar-open");
  syncSidebarToggle();
  document.body.classList.remove("has-results");
  dom.results.classList.add("hidden");
  dom.statusPanel.className = "status-panel hidden";
  dom.statusPanel.textContent = "";
  dom.statusPanel.removeAttribute("aria-busy");
  resetStatusNote();
  setImportBusy(false);
  const sourceDisplayName = state.sourceSheetName ? `${state.sourceFileName} · ${state.sourceSheetName}` : state.sourceFileName;
  dom.returnToAnalysis.textContent = canReturn ? formatReturnToAnalysis(sourceDisplayName) : "";
  dom.returnToAnalysis.classList.toggle("hidden", !canReturn);
  window.scrollTo({ top: 0, behavior: "instant" });
  if (canReturn) dom.returnToAnalysis.focus();
  else dom.fileInput.focus();
}

// Back to the analysis 换一个文件 left: the same module and the import's message, shown again
// as they were; nothing is recalculated or redrawn.
function returnToAnalysis() {
  if (!state.rows.length || !state.importStatus) return;
  dom.returnToAnalysis.classList.add("hidden");
  dom.returnToAnalysis.textContent = "";
  document.body.classList.add("has-results");
  dom.results.classList.remove("hidden");
  const { message, tone, hasNote, summary } = state.importStatus;
  setStatus(message, tone, hasNote);
  dom.statusNoteSummary.textContent = summary;
  syncSidebarToggle();
  showModule(state.returnModule || MODULE_IDS[0]);
}
