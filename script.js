// 启动入口（最后加载）：启动自检、Chart.js 全局样式、页面导航和全部事件绑定。所有加载时立即执行的代码都只放在这里。

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
    ["templates.js", typeof runTemplateAnalysis],
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
  dom.dateChartMetric,
  dom.dateChartAggregate,
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
