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
    ["format.js", typeof L],
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

  // v3.2.0: an English page is hidden until this script has set the language; show it again
  // first, so the notice below is not hidden with it. The mark is also the only sign of the
  // language here: format.js (L) may be the file that did not load.
  const english = Boolean(document.documentElement?.hasAttribute?.("data-lang-pending"));
  document.documentElement?.removeAttribute("data-lang-pending");
  const notice = document.createElement("div");
  notice.setAttribute("role", "alert");
  notice.style.cssText = "position:sticky;top:0;z-index:2147483647;display:flex;flex-wrap:wrap;align-items:center;justify-content:center;gap:12px;padding:12px 16px;border-bottom:1px solid #fedf89;background:#fffaeb;color:#b54708;font:600 14px/1.5 system-ui,-apple-system,\"Segoe UI\",\"Microsoft YaHei\",sans-serif;";
  const message = document.createElement("span");
  message.textContent = english ? "The page was updated. Refresh it to continue." : "页面已更新，请刷新后再使用。";
  const button = document.createElement("button");
  button.type = "button";
  button.textContent = english ? "Refresh" : "刷新页面";
  button.style.cssText = "padding:6px 14px;border:1px solid #b54708;border-radius:6px;background:#b54708;color:#fff;font:inherit;cursor:pointer;";
  button.addEventListener("click", () => window.location.reload());
  notice.append(message, button);
  document.body.prepend(notice);
  throw new Error(missingScripts.length
    ? `脚本文件缺失，HTML 与脚本版本不一致：${missingKeys.join(", ")}`
    : `页面元素缺失，HTML 与 script.js 版本不一致：${missingKeys.join(", ")}`);
}

// v3.2.0: the language, before anything is drawn or bound: ?lang= first, then the browser's
// first language. Chinese needs nothing (the page is written in it); the English page has been
// hidden by the inline script in <head> until now.
const INITIAL_LANGUAGE = pickInterfaceLanguage(window.location?.search, readNavigatorLanguages(typeof navigator === "undefined" ? null : navigator));
// The page's own Chinese, read now in both languages (commit 9): later switches write it back.
rememberStaticLanguage();
if (INITIAL_LANGUAGE !== "zh") setInterfaceLanguage(INITIAL_LANGUAGE);
document.documentElement?.removeAttribute("data-lang-pending");

if (window.Chart) {
  Chart.defaults.color = CHART_THEME.text;
  Chart.defaults.font.family = '-apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif';
  Chart.defaults.font.size = 12;
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
  // v3.2.0: the English interface loads the English sample (the same orders and problems).
  const sampleName = L("sample-orders.csv", "sample-orders-en.csv");
  try {
    setStatus(L(`正在加载 ${sampleName} ...`, `Loading ${sampleName} ...`));
    const response = await fetch(`examples/${sampleName}`);
    if (!response.ok) throw languageError(Lp("示例数据加载失败", "The sample data could not be loaded"));
    const text = await response.text();
    if (!isCurrentImport(importId)) return;
    parseCsvText(text, sampleName, importId);
  } catch (error) {
    showError(errorMessageBuilder(error), importId);
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

dom.checkStandardizedExport.addEventListener("click", checkStandardizedExport);
dom.exportStandardizedCsv.addEventListener("click", () => exportStandardizedData("data"));
dom.exportStandardizedChanges.addEventListener("click", () => exportStandardizedData("changes"));

dom.exportHtmlReport.addEventListener("click", exportHtmlReport);
// v3.4.0: the checkup's 导出体检报告 is the same HTML report, with the same gate (exportHtmlReport()).
dom.checkupExportReport.addEventListener("click", () => exportHtmlReport());
// v3.4.0 commit 4: the bar's segments and the cards' actions.
dom.checkupBar.addEventListener("click", handleCheckupClick);
dom.checkupCards.addEventListener("click", handleCheckupClick);

dom.confirmSheetSelection.addEventListener("click", analyzeSelectedExcelSheet);

// v3.2.1 换一个工作表: each list's button, its items (one delegated listener) and Escape.
SHEET_SWITCH_PLACES.forEach((place) => {
  const { toggle, panel, list } = sheetSwitchElements(place);
  toggle.addEventListener("click", () => toggleSheetSwitch(place));
  list.addEventListener("click", (event) => handleSheetSwitchClick(event, place));
  [toggle, panel].forEach((element) => element.addEventListener("keydown", (event) => handleSheetSwitchKeydown(event, place)));
});

[dom.languageSwitch, dom.sidebarLanguageSwitch].forEach((button) => button.addEventListener("click", switchInterfaceLanguage));


// v3.2.0 commit 9: the navigation gets its own row when it does not fit (updateNavFit(), app.js).
window.addEventListener("resize", updateNavFit);
// The navigation too: it appears with an import (display: none → flex) without the top bar
// changing its size, so watching the top bar alone misses a page opened narrow.
if (typeof ResizeObserver === "function" && document.querySelector?.(".app-header")) {
  const navFitObserver = new ResizeObserver(updateNavFit);
  navFitObserver.observe(document.querySelector(".app-header"));
  navFitObserver.observe(document.getElementById("moduleNav"));
}

// The switches are disabled while a worksheet waits to be chosen: follow the list's panel.
if (typeof MutationObserver === "function") {
  new MutationObserver(syncLanguageSwitches).observe(dom.sheetSelectionPanel, { attributes: true, attributeFilter: ["class"] });
}

dom.excelSheetSelect.addEventListener("change", () => {
  writeOneOffMessage(dom.sheetSelectionMessage, "sheet-selection-message", "");
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

dom.qualitySection.addEventListener("click", handleExclusionClick);

[dom.statisticsContainsNote, dom.chartsContainsNote, dom.customContainsNote].forEach((note) => note.addEventListener("click", handleAnalysisNoteClick));

dom.qualitySection.addEventListener("click", toggleSectionNote);

// The notes with a 展开 button (more than one sentence) show their first line until opened.
document.querySelectorAll(".note-toggle").forEach((button) => button.previousElementSibling.classList.add("is-clamped"));

dom.results.addEventListener("mouseover", showTruncatedCellTitle);

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
      writeOneOffMessage(dom.v2FieldPrompt, "analysis-prompt", () => L("字段已更新，点击“生成自定义分析”生成新的结果。", "The columns changed; click “Create breakdown” for new results."));
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
  clearStatusMessage();
  dom.statusPanel.removeAttribute("aria-busy");
  resetStatusNote();
  setImportBusy(false);
  const sourceDisplayName = state.sourceSheetName ? `${state.sourceFileName} · ${state.sourceSheetName}` : state.sourceFileName;
  // v3.2.1 (R8): kept as a function, so a language switch writes it again (rewriteLanguageTexts()).
  writeLanguageText(dom.returnToAnalysis, canReturn ? () => formatReturnToAnalysis(sourceDisplayName) : "");
  dom.returnToAnalysis.classList.toggle("hidden", !canReturn);
  // v3.2.1: 换一个工作表 comes with the way back (a multi-sheet file only).
  renderSheetSwitches();
  window.scrollTo({ top: 0, behavior: "instant" });
  if (canReturn) dom.returnToAnalysis.focus();
  else dom.fileInput.focus();
}

// Back to the analysis 换一个文件 left: the same module and the import's message, shown again
// as they were; nothing is recalculated or redrawn.
function returnToAnalysis() {
  if (!state.rows.length || !state.importStatus) return;
  dom.returnToAnalysis.classList.add("hidden");
  writeLanguageText(dom.returnToAnalysis, "");
  // v3.2.1: and the upload page's 换一个工作表 goes, its list closed and its message cleared.
  renderSheetSwitches();
  document.body.classList.add("has-results");
  dom.results.classList.remove("hidden");
  // v3.2.0: the message is written again from its figures, in the current language.
  showImportStatus(state.importStatus);
  syncSidebarToggle();
  showModule(state.returnModule || MODULE_IDS[0]);
}
