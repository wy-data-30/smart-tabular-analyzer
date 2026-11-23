// 显示格式化：数值、整数、百分比和字段格式的展示文本，带单位的字段标签，表头位置说明，编号重复的行号列表和说明文字，像是空值的写法、按单位读取的数字、像是数字的文本列、疑似合计 / 小计行、公式错误单元格、日期写法、整列只有一个值和“同一个值，写法不一样”的说明文字，CSV 读取时有疑问的中文消息，问题行的按钮、类别和摘要文字，写法统一的表的处理文字、检查表说明和状态行，以及 HTML 转义。
// v3.2.0 起还有界面语言：currentLanguage、L()、Lp() / Lpick()、enCount()、pickInterfaceLanguage()、readNavigatorLanguages()、languageSwitchLabel() 和 index.html 的英文表 I18N_EN。

// ---- Interface language (v3.2.0) ----
// Only setInterfaceLanguage() (app.js) changes it. Every text shown to the user reads it through
// L() or Lpick(); the data, the field names and the detection rules never depend on it.
let currentLanguage = "zh";

// One text in both languages, Chinese first. Both arguments are evaluated, so keep L() off the
// per-value and per-row paths.
function L(zh, en) {
  return currentLanguage === "en" ? en : zh;
}

// A fixed text in both languages kept in a constant, Lp("中文", "English"), read with Lpick().
function Lp(zh, en) {
  return Object.freeze([zh, en]);
}

function Lpick(pair) {
  return currentLanguage === "en" ? pair[1] : pair[0];
}

// "1 row", "2 rows", "1,234 rows": English only, the Chinese text has no plural.
function enCount(count, one, many) {
  return `${formatInteger(count)} ${count === 1 ? one : many}`;
}

// ?lang=zh or ?lang=en first (exactly these, lower case), then the browser's first language:
// zh, zh-CN, zh-TW, zh-HK… are Chinese, anything else English. Without either, as in the
// tests, Chinese. The inline script in index.html's <head> follows the same rule (a test runs
// both on the same cases).
function pickInterfaceLanguage(search, languages) {
  const asked = /(?:^|[?&])lang=(zh|en)(?:&|$)/.exec(String(search || ""))?.[1];
  if (asked) return asked;
  const first = Array.isArray(languages) && languages.length ? String(languages[0] || "") : "";
  if (!first) return "zh";
  return /^zh(?:-|$)/i.test(first) ? "zh" : "en";
}

// navigator.languages, or [navigator.language] when that list is empty; [] without a navigator.
function readNavigatorLanguages(browser) {
  if (!browser) return [];
  if (Array.isArray(browser.languages) && browser.languages.length) return browser.languages;
  return browser.language ? [browser.language] : [];
}

// v3.2.0 commit 11: an error whose message can be written again in the other language. source
// is an Lp() pair or a function that writes the message in the current language.
function languageError(source) {
  const build = typeof source === "function" ? source : () => Lpick(source);
  const error = new Error(build());
  error.buildMessage = build;
  return error;
}

// What the status bar shows for a caught error: a function for a languageError (written again
// after a switch), the plain message for any other error (PapaParse, SheetJS, the browser).
function errorMessageBuilder(error) {
  return typeof error?.buildMessage === "function" ? error.buildMessage : String(error?.message ?? error ?? "");
}

// The message of a caught error in the current language.
function errorText(error) {
  return typeof error?.buildMessage === "function" ? error.buildMessage() : String(error?.message ?? error ?? "");
}

// The switch names the other language, in that language.
const LANGUAGE_NAME_ZH = "中文";
function languageSwitchLabel() {
  return currentLanguage === "en" ? LANGUAGE_NAME_ZH : "English";
}

// English for the static text of index.html: data-i18n="key" (the text) and
// data-i18n-attr="attribute=key …" (attributes). The Chinese stays in index.html.
const I18N_EN = Object.freeze({
  "page.title": "Sheet Checkup | Check CSV and Excel files for data problems",
  "skip": "Skip to main content",
  "sidebar.toggle": "File and columns",
  "brand.label": "Sheet Checkup home",
  "brand.name": "Sheet Checkup",
  "nav.label": "Analysis modules",
  "nav.overview": "Overview",
  "nav.schema": "Columns",
  "nav.quality": "Quality",
  "nav.filter": "Filter",
  "nav.statistics": "Stats",
  "nav.charts": "Charts",
  "nav.custom": "Breakdown",
  "nav.export": "Export",
  "changeFile": "Change file",
  "sidebar.title": "File and columns",
  "sidebar.close": "Close file and columns",
  "sidebar.fields": "Columns",
  "sidebar.status": "Current status",
  "sidebar.filter": "Filter",
  "sidebar.header": "Header",
  "sidebar.placeholders": "Treated as blank",
  "privacy.sidebar": "Your file is processed only in this browser on this computer and is never uploaded.",
  "hero.title": "Upload a file for a checkup",
  "hero.description": "Nothing to install. Your file never leaves this computer.",
  "upload.title": "Drop a file here",
  "upload.limits": "CSV, XLSX or XLS, up to 25 MB, 100,000 rows and 200 columns",
  "upload.choose": "Choose a file",
  "upload.sample": "Use sample data",
  "upload.encoding": "CSV encoding",
  "upload.encodingAuto": "Detect automatically",
  "upload.encodingHint": "If unsure, keep “Detect automatically”",
  "upload.download": "Download sample CSV",
  "details.open": "Collapse",
  "details.closed": "Expand",
  // v3.2.0 commit 9: the rest of the page.
  "status.collapse": "Collapse",
  "sheet.title": "Choose a worksheet to analyze",
  "sheet.label": "Worksheet",
  "sheet.confirm": "Analyze this worksheet",
  "sample.title": "What it checks for",
  "sample.note": "A few rows of the sample data",
  "sample.tableWrap": "A few rows of the sample data, scrolls sideways",
  "sample.row": "#",
  "sample.orderId": "Order ID",
  "sample.orderDate": "Order date",
  "sample.region": "Region",
  "sample.amount": "Amount",
  "sample.flagged": "What gets flagged",
  "sample.south": "South",
  "sample.noteSpace": "An extra space at the end",
  "sample.otherDate": "02/18/2025",
  "sample.southwest": "Southwest",
  "sample.noteDate": "Written in a different date format from the other rows",
  "sample.noteDuplicate": "Same as the row above",
  "sample.northwest": "Northwest",
  "sample.notePlaceholder": "May mean blank; not recognized as a number",
  "sample.total": "Total",
  "sample.noteTotal": "A total row mixed in with the data",
  "workspace.title": "Analysis workspace",
  "scope.label": "In analysis: ",
  "scope.none": "No data loaded yet",
  "scope.clear": "Clear filters",
  "overview.title": "Size and preview",
  "preview.title": "Data preview",
  "preview.tableWrap": "Data preview table, scrolls sideways",
  "pager.first": "First",
  "pager.prev": "Previous",
  "pager.next": "Next",
  "pager.last": "Last",
  "schema.title": "Check column types",
  "schema.aside": "Change a type only if it was detected wrongly; applying recalculates everything",
  "schema.search": "Search columns",
  "schema.kinds": "Filter columns by type",
  "schema.empty": "No matching columns.",
  "schema.tableWrap": "Column settings table, scrolls sideways",
  "schema.table": "Column types",
  "schema.restore": "Restore detected types",
  "schema.apply": "Apply column settings",
  "quality.title": "Blanks, duplicates and other issues",
  "quality.aside": "Always based on the imported data; filters do not change it",
  "quality.totals": "Possible total or subtotal rows",
  "quality.totalsNote": "These rows still count as ordinary data, so their numbers are counted again in sums, averages and the check for unusually high or low values.",
  "quality.formulaNote": "The formulas in these cells returned an error; they are treated as blank (and counted as blanks). Rows that are all errors are not imported and are not listed here.",
  "quality.dateStyleNote": "This tool reads all of these as the same date, but Excel cannot sort or filter dates stored as text, and other systems may read them wrongly.",
  "quality.constantNote": "All non-blank values in these columns are the same. That is not always an error: it may be a fixed column added on export, or values left out or filled in wrongly.",
  "quality.lookAlikeNote": "Leading or trailing spaces: this tool ignores them, but they will not match when exported to other systems or used with VLOOKUP. One value written in several ways: this tool counts them as different values, so value counts, filters and duplicate IDs count them separately. Only invisible characters: the cell looks empty, but this tool does not treat it as blank, so value counts and filters show an extra value you cannot see.",
  "quality.formula": "Cells with formula errors",
  "note.expand": "Expand",
  "quality.missing": "Blanks",
  "quality.missingWrap": "Blanks table, scrolls sideways",
  "quality.missingTable": "Blanks by column",
  "quality.outliers": "Unusually high or low values in number columns",
  "quality.outliersWrap": "Unusually high or low values table, scrolls sideways",
  "problemRows.title": "Problem rows",
  "problemRows.export": "Export these rows",
  "problemRows.close": "Close",
  "problemRows.category": "Category",
  "problemRows.draft": "Column settings have changes that are not applied; these rows follow the applied settings.",
  "problemRows.tableWrap": "Problem rows table, scrolls sideways",
  "quality.units": "Numbers read with a unit",
  "quality.numericLooking": "Text columns that look numeric",
  "quality.dateStyles": "Mixed date formats",
  "quality.constant": "Single-value columns",
  "quality.lookAlike": "Same value, written differently",
  "quality.placeholders": "Values that may mean blank",
  "quality.placeholdersNote": "These values are often used to mean “no data”. Once they are treated as blank, they no longer count as numbers or dates not recognized, and they are left out of the statistics, the top 10, groups, filter options and the duplicate ID check. The original data, the data preview and the exported CSV do not change.",
  "quality.idDuplicates": "Duplicate IDs",
  "filter.title": "Choose the rows to analyze",
  "filter.aside": "Changes only the analysis, not the original data or data quality",
  "filter.counts": "Row counts",
  "filter.original": "Original rows",
  "filter.current": "Rows in analysis",
  "filter.filteredOut": "Rows filtered out",
  "filter.notCounted": "Excluded from analysis",
  "filter.category": "Text column",
  "filter.noColumn": "No column chosen",
  "filter.field": "Column",
  "filter.search": "Search values",
  "filter.searchPlaceholder": "Type to search",
  "filter.selectAll": "Select all",
  "filter.clearSelection": "Clear selection",
  "filter.values": "Values",
  "filter.chooseText": "Choose a text column.",
  "filter.numeric": "Number column",
  "filter.chooseColumn": "Choose a column",
  "filter.min": "Minimum",
  "filter.max": "Maximum",
  "filter.date": "Date column",
  "filter.start": "Start date",
  "filter.end": "End date",
  "filter.none": "No filter is applied; the analysis uses all rows.",
  "filter.clearAll": "Clear all filters",
  "filter.apply": "Apply filters",
  "statistics.title": "Number and text stats",
  "statistics.numeric": "Number columns",
  "statistics.numericWrap": "Number statistics table, scrolls sideways",
  "statistics.numericTable": "Number column statistics",
  "statistics.category": "Value counts of text columns",
  "charts.title": "Distributions and trends",
  "charts.numeric": "Number column distribution (record count)",
  "charts.field": "Column",
  "charts.numericChart": "Number column distribution chart",
  "charts.category": "Text column top 10 (record count)",
  "charts.categoryChart": "Most common values chart",
  "charts.date": "Date column trend (record count)",
  "charts.metric": "Metric",
  "charts.aggregate": "Aggregation",
  "aggregate.avg": "Average",
  "aggregate.sum": "Sum",
  "charts.dateChart": "Date column trend chart",
  "custom.title": "Combine columns as you need",
  "custom.note": "Choose a metric, a column to group by and an aggregation to see the results; a date column is optional.",
  "custom.create": "Create breakdown",
  "custom.createNote": "Creates a summary table, a group chart and a short summary.",
  "custom.metric": "Metric column",
  "custom.group": "Group by",
  "custom.aggregate": "Aggregation",
  "aggregate.median": "Median",
  "custom.date": "Date column (optional)",
  "custom.summaryTable": "Summary table",
  "custom.summaryWrap": "Breakdown summary, scrolls sideways",
  "custom.summary": "Breakdown summary",
  "custom.groups": "Groups",
  "custom.groupsChart": "Breakdown group bar chart",
  "custom.trend": "Trend over time",
  "custom.trendChart": "Breakdown trend chart",
  "custom.details": "See all statistics",
  "custom.sum": "Sum by group",
  "custom.sumTable": "Sum by group",
  "custom.sumWrap": "Sum by group, scrolls sideways",
  "custom.avg": "Average by group",
  "custom.avgTable": "Average by group",
  "custom.avgWrap": "Average by group, scrolls sideways",
  "export.title": "Save the data and the report",
  "export.data": "Export the data",
  "export.dataNote": "The export happens only in this browser. Files are UTF-8 with a BOM and keep the original column order; the original file is not changed or overwritten.",
  "export.analysisRows": "Rows in analysis",
  "export.analysisNote": "Exports the rows the statistics and charts use (the rows after filtering).",
  "export.analysisButton": "Export rows in analysis",
  "export.standardized": "Standardized table",
  "export.standardizedNote": "All original rows: dates written as YYYY-MM-DD; numbers without units, thousands separators or currency signs; values treated as blank emptied; leading and trailing spaces removed from text. Check which cells will change first.",
  "export.check": "Check which cells will change",
  "export.changesWrap": "Standardized changes table, scrolls sideways",
  "export.standardizedButton": "Export the standardized table",
  "export.changesButton": "Export the change list",
  "export.report": "Report",
  "export.reportTitle": "Create the report in this browser",
  "export.reportHint": "The report includes the current charts and can be printed to PDF from the browser. Nothing is uploaded during the export.",
  "export.reportButton": "Export report",
  "footer.github": "GitHub repository"
});

function absorbFloatNoise(value) {
  return Math.abs(value) < 1e-10 ? 0 : value;
}

function stripNegativeZeroSign(text) {
  // "-0", "-0.0", "-0%" and "-0.0%" are zero after rounding; show them unsigned.
  return /^-[0.,]*%?$/.test(text) ? text.slice(1) : text;
}

function formatNumber(value) {
  if (!Number.isFinite(value)) return "-";
  const displayed = absorbFloatNoise(value);
  const magnitude = Math.abs(displayed);
  const options = magnitude > 0 && magnitude < 1
    ? { maximumSignificantDigits: 3 }
    : { maximumFractionDigits: 2 };
  return stripNegativeZeroSign(new Intl.NumberFormat("zh-CN", options).format(displayed));
}

function formatInteger(value) {
  return new Intl.NumberFormat("zh-CN", {
    maximumFractionDigits: 0
  }).format(value || 0);
}

// v3.2.0: the list separator of the language, for lists built on a per-row path.
const LIST_SEPARATOR = Lp("、", ", ");

// The name of an applied type (数值 / 分类 / 日期 / ID / 忽略) in the current language.
function fieldTypeLabel(typeKey) {
  return Lpick(FIELD_TYPE_LABELS[typeKey]);
}

// Plain text such as "3（N/A、待定、约500）"; callers escape it for their output format.
function formatConversionFailures(count, samples = [], countText = String(count ?? 0)) {
  if (!count || !samples.length) return countText;
  return L(`${countText}（${samples.join("、")}）`, `${countText} (${samples.join(", ")})`);
}

function formatPercent(value) {
  if (!Number.isFinite(value)) return "0%";
  return stripNegativeZeroSign(`${(absorbFloatNoise(value) * 100).toFixed(1)}%`);
}

function formatFieldNumber(field, value) {
  if (!Number.isFinite(value)) return "-";
  if (getNumericFormat(field) === "percent") {
    return stripNegativeZeroSign(new Intl.NumberFormat("zh-CN", {
      style: "percent",
      maximumFractionDigits: 2
    }).format(absorbFloatNoise(value)));
  }
  return formatNumber(value);
}

// The unit goes before "（百分比）".
function numericAxisLabel(field, fallbackLabel) {
  const unit = getFieldNumberUnit(field);
  const unitText = unit ? L(`（${unit}）`, ` (${unit})`) : "";
  return getNumericFormat(field) === "percent"
    ? L(`${fallbackLabel}${unitText}（百分比）`, `${fallbackLabel}${unitText} (percent)`)
    : `${fallbackLabel}${unitText}`;
}

// "金额（元）" for a field read with a unit; the name stays when it already ends
// with the unit: …u, …(u), …[u], …【u】, …/u, …_u, …-u or … u, compared after NFKC
// and in lower case. A Latin unit must not follow a letter ("bag" has no kg).
function formatFieldLabelWithUnit(field, unit) {
  if (!unit) return field;
  const name = String(field).normalize("NFKC").trim().toLowerCase();
  const target = unit.toLowerCase();
  const bracketed = [`(${target})`, `[${target}]`, `【${target}】`].some((ending) => name.endsWith(ending));
  const bare = name.endsWith(target) && !(/^[a-z]/.test(target) && /[a-z]/.test(name.charAt(name.length - target.length - 1)));
  return bracketed || bare ? field : L(`${field}（${unit}）`, `${field} (${unit})`);
}

function formatFieldLabel(field) {
  return formatFieldLabelWithUnit(field, getFieldNumberUnit(field));
}

// One line of "按单位读取的数字" for a buildNumberConversionSummary() item.
function formatNumberConversion(item) {
  const parts = [];
  if (item.unitCount) {
    parts.push(L(
      `按"${item.unit}"读取了 ${formatInteger(item.unitCount)} 个值（${item.unitSamples.join("、")}）`,
      `read ${enCount(item.unitCount, "value", "values")} as “${item.unit}” (${item.unitSamples.join(", ")})`
    ));
  }
  if (item.magnitudeCount) {
    parts.push(L(
      `按万/亿换算了 ${formatInteger(item.magnitudeCount)} 个值`,
      `scaled ${enCount(item.magnitudeCount, "value", "values")} written in units of 10,000 or 100,000,000`
    ));
  }
  if (item.fullwidthCount) {
    parts.push(L(
      `${formatInteger(item.fullwidthCount)} 个全角数字已按半角读取`,
      `read ${enCount(item.fullwidthCount, "full-width number", "full-width numbers")} as half-width`
    ));
  }
  if (item.conflictUnits.length) {
    const units = item.conflictUnits.map(({ unit, count }) => L(`${unit}（${formatInteger(count)} 个）`, `${unit} (${formatInteger(count)})`));
    parts.push(L(
      `同时出现 ${units.slice(0, -1).join("、")}和 ${units[units.length - 1]}，没有按单位读取，这些值不能读成数字。`,
      `has both ${units.slice(0, -1).join(", ")} and ${units[units.length - 1]}, so no unit is applied and these values cannot be read as numbers.`
    ));
  }
  return L(`「${item.field}」：${parts.join("；")}`, `“${item.field}”: ${parts.join("; ")}`);
}

// One line of "疑似数值列" for a state.numericLooking item.
function formatNumericLookingHint(item) {
  return L(
    `「${item.field}」：去掉空值和像是空值的写法后，${formatInteger(item.valueCount)} 个值里有 ${formatInteger(item.numericCount)} 个能读成数字（${formatPercent(item.ratio)}），其余的如 ${item.samples.join("、")}。如果这一列应该是数字，请修正这些值，或在字段配置里把类型改成数值。`,
    `“${item.field}”: leaving out blanks and values that may mean blank, ${formatInteger(item.numericCount)} of ${enCount(item.valueCount, "value", "values")} can be read as numbers (${formatPercent(item.ratio)}); the others include ${item.samples.join(", ")}. If this column should hold numbers, correct those values, or change its type to Number in Column settings.`
  );
}

// How many columns a suspected total's line names before "等 N 列".
const TOTAL_MATCHED_FIELD_LIMIT = 5;
// How many row numbers the status bar names before 等.
const STATUS_ROW_NUMBER_LIMIT = 5;

// The notes of "疑似合计 / 小计行" and "公式错误单元格" in reports; index.html has the same
// sentences.
const SUSPECTED_TOTAL_NOTE = Lp(
  "这些行目前仍按普通数据参与统计，其中的数字会被重复计入求和、平均值和偏大或偏小的值的判断。",
  "These rows still count as ordinary data, so their numbers are counted again in sums, averages and the check for unusually high or low values."
);
const FORMULA_ERROR_NOTE = Lp(
  "这些单元格的公式算出了错误，按空值处理（已计入空值）。整行都是错误值的行不会导入，也不在这里列出。",
  "The formulas in these cells returned an error; they are treated as blank (and counted as blanks). Rows that are all errors are not imported and are not listed here."
);

// English rows and their numbers: "row 44", "rows 5, 8, 10", "rows 1, 2, 3, 4, 5 and more".
function formatEnglishRows(lines, more = false) {
  return `${lines.length === 1 && !more ? "row" : "rows"} ${lines.join(", ")}${more ? " and more" : ""}`;
}

function formatTotalMatchedFields(fields) {
  if (fields.length <= TOTAL_MATCHED_FIELD_LIMIT) return fields.join(Lpick(LIST_SEPARATOR));
  const shown = fields.slice(0, TOTAL_MATCHED_FIELD_LIMIT);
  return L(`${shown.join("、")} 等 ${fields.length} 列`, `${shown.join(", ")} and others (${fields.length} columns in all)`);
}

// "第 5–32 行之和（28 行）", "第 2–13 行中的明细之和（9 行）" when other rows lie in between,
// "第 9 行（只有 1 行）", or the range in words for a table without row numbers.
function formatTotalAddends({ scope, firstLine, lastLine, rowCount, contiguous }) {
  const count = formatInteger(rowCount);
  if (firstLine === null || lastLine === null) {
    return L(`${scope === "all" ? "全表" : "上方"} ${count} 行明细之和`, `the sum of ${enCount(rowCount, "detail row", "detail rows")} ${scope === "all" ? "in the whole table" : "above"}`);
  }
  if (rowCount === 1) return L(`第 ${firstLine} 行（只有 1 行）`, `row ${firstLine} (1 row only)`);
  const lines = `${firstLine}–${lastLine}`;
  return contiguous
    ? L(`第 ${lines} 行之和（${count} 行）`, `the sum of rows ${lines} (${count} rows)`)
    : L(`第 ${lines} 行中的明细之和（${count} 行）`, `the sum of the detail rows among rows ${lines} (${count} rows)`);
}

// "销售额、成本等于第 5–32 行之和（28 行）": the 说明 of a suspected total's problem row.
function formatSuspectedTotalMatches(item) {
  return item.matches.map((match) => L(
    `${formatTotalMatchedFields(match.fields)}等于${formatTotalAddends(match)}`,
    `${formatTotalMatchedFields(match.fields)} ${match.fields.length === 1 ? "equals" : "equal"} ${formatTotalAddends(match)}`
  )).join(L("；", "; "));
}

// One line of "疑似合计 / 小计行".
function formatSuspectedTotal(item) {
  return L(
    `${item.line === null ? "有一行" : `第 ${item.line} 行`}「${item.label}」：${formatSuspectedTotalMatches(item)}。`,
    `${item.line === null ? "A row" : `Row ${item.line}`} “${item.label}”: ${formatSuspectedTotalMatches(item)}.`
  );
}

// "共 N 行，只列出前 20 行。" when a list of suspected totals is cut short, or "".
function formatSuspectedTotalOverflow(count) {
  if (count <= SUSPECTED_TOTAL_LIST_LIMIT) return "";
  return L(
    `共 ${formatInteger(count)} 行，只列出前 ${SUSPECTED_TOTAL_LIST_LIMIT} 行。`,
    `${formatInteger(count)} rows in all; only the first ${SUSPECTED_TOTAL_LIST_LIMIT} are listed.`
  );
}

// The status bar's sentence after an import, or "".
function formatSuspectedTotalStatus(items) {
  if (!items.length) return "";
  if (items.length === 1) {
    const [{ line, label }] = items;
    return L(
      `${line === null ? "有一行" : `第 ${line} 行`}「${formatSkippedRowPreview([label])}」疑似合计行，可以在“数据质量”里设为不计入分析。`,
      `${line === null ? "A row" : `Row ${line}`} “${formatSkippedRowPreview([label])}” looks like a total row; you can exclude it from analysis in Data quality.`
    );
  }
  const lines = items.map((item) => item.line);
  const named = !lines.some((line) => line === null);
  const shownLines = lines.slice(0, STATUS_ROW_NUMBER_LIMIT);
  const more = lines.length > STATUS_ROW_NUMBER_LIMIT;
  return L(
    `发现 ${formatInteger(items.length)} 行疑似合计 / 小计${named ? `（第 ${shownLines.join("、")} 行${more ? "等" : ""}）` : ""}，可以在“数据质量”里设为不计入分析。`,
    `Found ${formatInteger(items.length)} possible total or subtotal rows${named ? ` (${formatEnglishRows(shownLines, more)})` : ""}; you can exclude them from analysis in Data quality.`
  );
}

// ---- v3.1.0: rows left out of the analysis, the sentences of the rebuild path ----
// How many file rows a sentence names before 等.
const EXCLUSION_LINE_LIMIT = 5;

// "第 44 行", "第 5、8 行", "第 1、2、3、4、5 行等"; null without row numbers.
function formatExclusionLines(lines) {
  if (!lines.length || lines.some((line) => line === null)) return null;
  const shown = lines.slice(0, EXCLUSION_LINE_LIMIT);
  const more = lines.length > EXCLUSION_LINE_LIMIT;
  return L(`第 ${shown.join("、")} 行${more ? "等" : ""}`, formatEnglishRows(shown, more));
}

// After a rebuild: excluded totals that are no longer suspected totals came back.
function formatExclusionRestored(count) {
  if (!count) return "";
  return L(
    `有 ${formatInteger(count)} 行不再是疑似合计行，已恢复计入分析。`,
    count === 1
      ? "1 row is no longer a possible total row and counts in the analysis again."
      : `${formatInteger(count)} rows are no longer possible total rows and count in the analysis again.`
  );
}

// After a rebuild: totals that came back earlier in this import are suspected totals again
// (they stay counted until the user leaves them out again).
function formatTotalsFlaggedAgain(lines) {
  if (!lines.length) return "";
  const named = formatExclusionLines(lines);
  const subject = named ? `${named.charAt(0).toUpperCase()}${named.slice(1)}` : enCount(lines.length, "row", "rows");
  return L(
    `${named || `有 ${formatInteger(lines.length)} 行`}又被标为疑似合计行，目前仍计入分析，可以再设为不计入分析。`,
    lines.length === 1
      ? `${subject} is marked as a possible total row again; it still counts in the analysis, and you can exclude it again.`
      : `${subject} are marked as possible total rows again; they still count in the analysis, and you can exclude them again.`
  );
}

// ---- v3.1.0: the page around rows left out of the analysis ----
const EXCLUSION_KIND_LABELS = Object.freeze({ suspectedTotal: Lp("疑似合计行", "possible total rows"), duplicateRows: Lp("重复的行", "duplicate rows") });
// English names with a count: [one, many].
const EXCLUSION_KIND_NOUNS = Object.freeze({ suspectedTotal: ["possible total row", "possible total rows"], duplicateRows: ["duplicate row", "duplicate rows"] });
const EXCLUSION_DRAFT_REFUSAL = Lp(
  "字段配置有未应用的修改。请先点击“应用字段配置”，再设置不计入分析。",
  "Column settings have changes that are not applied. Click “Apply column settings” first, then exclude rows from analysis."
);
const NO_ID_DUPLICATE_NOTE = Lp(
  "这张表没有能区分每一行的编号列，内容相同的行也可能是不同的记录，先点“查看重复行”确认。",
  "This table has no ID column that tells each row apart, so rows with the same content may still be different records. Click “View duplicate rows” to check first."
);
// Empty states that also show without a filter (they said 当前筛选 before v3.1.0).
const CATEGORY_STATS_EMPTY_NOTE = Lp("当前分析的行里没有非空类别值。", "The rows in analysis have no non-blank values.");
const CATEGORY_CHART_EMPTY_NOTE = Lp("当前分析的行里没有可统计的分类值", "The rows in analysis have no values to count");

// "（另有 2 行不计入分析）", or "" when nothing is left out.
function formatExcludedTail(excludedCount) {
  if (!excludedCount) return "";
  return L(`（另有 ${formatInteger(excludedCount)} 行不计入分析）`, ` (${formatInteger(excludedCount)} more excluded from analysis)`);
}

// The scope sentence of #globalFilterSummary, the three .filter-scope-text and the sidebar.
// Nothing left out: the v3.0.1 sentence.
function formatAnalysisScope(activeCount, currentCount, includedCount, excludedCount) {
  const tail = formatExcludedTail(excludedCount);
  if (activeCount) {
    return L(
      `已筛选 ${formatInteger(activeCount)} 项 · ${formatInteger(currentCount)} / ${formatInteger(includedCount)} 行${tail}`,
      `${enCount(activeCount, "filter", "filters")} · ${formatInteger(currentCount)} / ${enCount(includedCount, "row", "rows")}${tail}`
    );
  }
  if (excludedCount) return L(`${formatInteger(includedCount)} 行${tail}`, `${enCount(includedCount, "row", "rows")}${tail}`);
  return L(`全部 ${formatInteger(includedCount)} 行`, includedCount === 1 ? "1 row" : `All ${formatInteger(includedCount)} rows`);
}

// "全部记录", or "全部 40 行（不含不计入分析的 2 行）".
function formatAllIncludedRows(includedCount, excludedCount) {
  if (!excludedCount) return L("全部记录", "all rows");
  return L(
    `全部 ${formatInteger(includedCount)} 行（不含不计入分析的 ${formatInteger(excludedCount)} 行）`,
    `all ${enCount(includedCount, "row", "rows")} (not counting ${enCount(excludedCount, "row", "rows")} excluded from analysis)`
  );
}

// v3.2.0 commit 10: the words around an export option's row count: “共 42 行” / “42 rows”. The
// number itself is written into its own <strong> (dataExportFilteredCount, dataExportOriginalCount).
function formatExportRowCount(count) {
  return L({ before: "共 ", after: " 行" }, { before: "", after: count === 1 ? " row" : " rows" });
}

// After the export's static note while rows are left out (v3.1.0).
function formatDataExportExcludedNote(excludedCount) {
  return L(`不含不计入分析的 ${formatInteger(excludedCount)} 行。`, `Not counting ${enCount(excludedCount, "row", "rows")} excluded from analysis.`);
}

// The report's 筛选条件 without a filter while rows are left out; "" otherwise, and the report
// keeps its REPORT_NO_FILTER_TEXT (完整数据).
function formatReportNoFilterText(includedCount, excludedCount) {
  if (!excludedCount) return "";
  return L(
    `未设置筛选，结果基于${formatAllIncludedRows(includedCount, excludedCount)}。`,
    `No filter is set; the results are based on ${formatAllIncludedRows(includedCount, excludedCount)}.`
  );
}

// The filter panel's message while no filter is applied.
function formatNoFilterMessage(includedCount, excludedCount) {
  if (!excludedCount) return L("尚未应用筛选，当前分析使用全部记录。", "No filter is applied; the analysis uses all rows.");
  return L(`当前分析使用${formatAllIncludedRows(includedCount, excludedCount)}。`, `The analysis uses ${formatAllIncludedRows(includedCount, excludedCount)}.`);
}

// After a filter is applied: filteredOutCount counts from the rows counted in the analysis.
function formatFiltersApplied(activeCount, keptCount, filteredOutCount) {
  if (!keptCount) {
    return L(
      `已应用 ${formatInteger(activeCount)} 个筛选条件，但没有记录符合条件。`,
      `Applied ${enCount(activeCount, "filter", "filters")}, but no rows match.`
    );
  }
  return L(
    `已应用 ${formatInteger(activeCount)} 个筛选条件，保留 ${formatInteger(keptCount)} 行，筛除 ${formatInteger(filteredOutCount)} 行。`,
    `Applied ${enCount(activeCount, "filter", "filters")}: kept ${enCount(keptCount, "row", "rows")}, filtered out ${enCount(filteredOutCount, "row", "rows")}.`
  );
}

// The data preview's page line: the rows left out are named only without a filter.
function formatPreviewPageInfo(page, pageCount, rowCount, excludedCount, filtered) {
  const tail = excludedCount && !filtered;
  return L(
    `第 ${formatInteger(page)} / ${formatInteger(pageCount)} 页，共 ${formatInteger(rowCount)} 行${tail ? `（不含不计入分析的 ${formatInteger(excludedCount)} 行）` : ""}`,
    `Page ${formatInteger(page)} of ${formatInteger(pageCount)}, ${enCount(rowCount, "row", "rows")}${tail ? ` (not counting ${enCount(excludedCount, "row", "rows")} excluded from analysis)` : ""}`
  );
}

// The custom analysis's first sentence; nothing left out: the v3.0.1 sentences.
function formatCustomAnalysisScope({ groupField, metricField, aggregateLabel, filtered, currentCount, includedCount, totalCount, excludedCount }) {
  const tail = formatExcludedTail(excludedCount);
  if (filtered) {
    return L(
      `已按 ${groupField} 对 ${metricField} 进行${aggregateLabel}，结果基于当前筛选数据 ${formatInteger(currentCount)} / ${formatInteger(includedCount)} 行${tail}。`,
      `${metricField} by ${groupField} (${aggregateLabel}), based on the filtered rows, ${formatInteger(currentCount)} / ${enCount(includedCount, "row", "rows")}${tail}.`
    );
  }
  if (excludedCount) {
    return L(
      `已按 ${groupField} 对 ${metricField} 进行${aggregateLabel}，结果基于 ${formatInteger(includedCount)} 行${tail}。`,
      `${metricField} by ${groupField} (${aggregateLabel}), based on ${enCount(includedCount, "row", "rows")}${tail}.`
    );
  }
  return L(
    `已按 ${groupField} 对 ${metricField} 进行${aggregateLabel}，结果基于完整数据 ${formatInteger(totalCount)} 行。`,
    `${metricField} by ${groupField} (${aggregateLabel}), based on all the data, ${enCount(totalCount, "row", "rows")}.`
  );
}

// On top of the quality page while rows are left out; missing counts as state.totalMissing does.
function formatQualityScopeNote(totalCount, excludedCount, excludedMissing) {
  if (!excludedCount) return "";
  const total = formatInteger(totalCount);
  return L(
    `质量页检查的是文件里的全部 ${total} 行，包括不计入分析的 ${formatInteger(excludedCount)} 行${excludedMissing ? `（其中有 ${formatInteger(excludedMissing)} 个空值）` : ""}；偏大或偏小的值也按这 ${total} 行算。`,
    `The quality checks cover all ${enCount(totalCount, "row", "rows")} of the file, including ${enCount(excludedCount, "row", "rows")} excluded from analysis${excludedMissing ? ` (with ${enCount(excludedMissing, "blank", "blanks")})` : ""}; unusually high or low values are also judged on these ${total} rows.`
  );
}

// The totals subsection's note once they are left out (SUSPECTED_TOTAL_NOTE before).
function formatSuspectedTotalExcludedNote(count) {
  return L(
    `这 ${formatInteger(count)} 行已不计入分析：统计、图表、自定义分析和筛选不再包含${count > 1 ? "它们" : "它"}。`,
    `${enCount(count, "row is", "rows are")} excluded from analysis: stats, charts, breakdowns and filters no longer include ${count > 1 ? "them" : "it"}.`
  );
}

// Suspected totals whose every match covers one row only (they may be ordinary rows).
function formatSingleRowTotals(count) {
  if (!count) return "";
  return L(
    `其中 ${formatInteger(count)} 行只等于上面一行，可能不是合计行，先点“查看”确认。`,
    count === 1
      ? "1 of them equals only the row above and may not be a total row; click “View” to check first."
      : `${formatInteger(count)} of them equal only the row above and may not be total rows; click “View” to check first.`
  );
}

// Labelled rows that are not suspected totals; only when a numeric column was checked.
function formatUnmatchedTotalLabels(labelCount, totalCount, checked) {
  if (!checked || labelCount <= totalCount) return "";
  const extra = labelCount - totalCount;
  return L(
    `另有 ${formatInteger(extra)} 行也带这类字样，但数字对不上，没有列出。`,
    extra === 1
      ? "1 more row also has such a word, but its numbers do not add up, so it is not listed."
      : `${formatInteger(extra)} more rows also have such words, but their numbers do not add up, so they are not listed.`
  );
}

// The two sentences under the suspected totals, page and report alike: extras is
// { singleRow, labelRows, checked } (getSuspectedTotalExtras()), or null for an old snapshot.
function formatSuspectedTotalExtras(extras, totalCount) {
  if (!extras) return [];
  return [formatSingleRowTotals(extras.singleRow), formatUnmatchedTotalLabels(extras.labelRows, totalCount, extras.checked)].filter(Boolean);
}

// The report's "筛选条件" line: { suspectedTotal: { count, lines }, duplicateRows: { count, lines } }
// (buildReportExclusions()). At most REPORT_EXCLUSION_LINE_LIMIT file rows, then 等; lines is
// null when a row has no file row number, and the brackets are left out.
const REPORT_EXCLUSION_LINE_LIMIT = 20;

function formatReportExclusions(exclusions) {
  if (!exclusions) return "";
  const parts = ["suspectedTotal", "duplicateRows"].map((kind) => {
    const item = exclusions[kind];
    if (!item?.count) return "";
    const listed = Array.isArray(item.lines) && item.lines.length;
    const more = listed && item.count > item.lines.length;
    return L(
      `${Lpick(EXCLUSION_KIND_LABELS[kind])} ${formatInteger(item.count)} 行${listed ? `（第 ${item.lines.join("、")} 行${more ? "等" : ""}）` : ""}`,
      `${enCount(item.count, ...EXCLUSION_KIND_NOUNS[kind])}${listed ? ` (${formatEnglishRows(item.lines, more)})` : ""}`
    );
  }).filter(Boolean);
  if (!parts.length) return "";
  return L(
    `不计入分析：${parts.join("、")}${formatReportExclusionOverlap(exclusions)}。`,
    `Excluded from analysis: ${parts.join(", ")}${formatReportExclusionOverlap(exclusions)}.`
  );
}

// Rows of both kinds are counted in each: then "，共 N 行（第 7 行两类都是）", N the rows left out.
// exclusions.overlap is { count, lines } (lines null without file rows), or null without overlap.
function formatReportExclusionOverlap({ suspectedTotal, duplicateRows, overlap }) {
  if (!overlap?.count) return "";
  const total = (suspectedTotal?.count || 0) + (duplicateRows?.count || 0) - overlap.count;
  const listed = Array.isArray(overlap.lines) && overlap.lines.length;
  const more = listed && overlap.count > overlap.lines.length;
  return L(
    `，共 ${formatInteger(total)} 行${listed ? `（第 ${overlap.lines.join("、")} 行${more ? "等" : ""}两类都是）` : ""}`,
    `, ${enCount(total, "row", "rows")} in all${listed ? ` (${formatEnglishRows(overlap.lines, more)} ${overlap.lines.length === 1 && !more ? "is" : "are"} both)` : ""}`
  );
}

// The repeated-rows line before they are left out, when a clean ID column tells the records apart.
function formatDuplicateIncludedNote(count) {
  return L(
    `重复行里多出来的 ${formatInteger(count)} 行目前仍计入分析。`,
    `${enCount(count, "extra copy", "extra copies")} of duplicate rows still ${count === 1 ? "counts" : "count"} in the analysis.`
  );
}

function formatDuplicateExcludedNote(count) {
  return L(`多出来的 ${formatInteger(count)} 行已不计入分析。`, `${enCount(count, "extra copy is", "extra copies are")} excluded from analysis.`);
}

// The note on top of the statistics, charts and custom analysis: the rows of the current
// analysis that are still suspected totals or extra repeated rows. Empty when there are none.
const GO_TO_QUALITY_LABEL = Lp("去数据质量处理", "Go to Data quality");

function formatAnalysisContainsNote(totalCount, duplicateCount) {
  if (!totalCount && !duplicateCount) return "";
  return L(
    `这些结果包含 ${[totalCount ? `${formatInteger(totalCount)} 行疑似合计行` : "", duplicateCount ? `${formatInteger(duplicateCount)} 行重复的行` : ""].filter(Boolean).join("、")}。`,
    `These results include ${[totalCount ? enCount(totalCount, "possible total row", "possible total rows") : "", duplicateCount ? enCount(duplicateCount, "duplicate row", "duplicate rows") : ""].filter(Boolean).join(" and ")}.`
  );
}

function formatGoToQualityAriaLabel(totalCount, duplicateCount) {
  return L(
    `${GO_TO_QUALITY_LABEL[0]}${[totalCount ? "疑似合计行" : "", duplicateCount ? "重复的行" : ""].filter(Boolean).join("和")}`,
    `${GO_TO_QUALITY_LABEL[1]} for ${[totalCount ? "possible total rows" : "", duplicateCount ? "duplicate rows" : ""].filter(Boolean).join(" and ")}`
  );
}

// The message of a click: { kind, on, rows, changed, before, after } (state.lastExclusion).
// "N 行" counts only the rows this click left out or brought back.
function formatExclusionMessage({ kind, on, rows, changed, before, after }) {
  const otherKind = kind === "duplicateRows" ? "suspectedTotal" : "duplicateRows";
  const label = EXCLUSION_KIND_LABELS[kind][0];
  const other = EXCLUSION_KIND_LABELS[otherKind][0];
  const [one, many] = EXCLUSION_KIND_NOUNS[kind];
  const otherMany = EXCLUSION_KIND_NOUNS[otherKind][1];
  const counted = L(`参与分析的行 ${formatInteger(before)} → ${formatInteger(after)}`, `rows in analysis ${formatInteger(before)} → ${formatInteger(after)}`);
  const still = L(`参与分析的行仍是 ${formatInteger(after)}`, `rows in analysis still ${formatInteger(after)}`);
  if (on) {
    if (changed) return L(`已把 ${formatInteger(changed)} 行${label}设为不计入分析，${counted}。`, `Excluded ${enCount(changed, one, many)} from analysis; ${counted}.`);
    if (kind === "duplicateRows") {
      return L(
        `多出来的 ${formatInteger(rows)} 行已经作为${other}不计入分析，${still}。`,
        `${rows === 1 ? "The 1 extra copy is" : `The ${formatInteger(rows)} extra copies are`} already excluded as ${otherMany}; ${still}.`
      );
    }
    return L(
      `这 ${formatInteger(rows)} 行${label}已经作为${other}不计入分析，${still}。`,
      `${rows === 1 ? `This 1 ${one} is` : `These ${formatInteger(rows)} ${many} are`} already excluded as ${otherMany}; ${still}.`
    );
  }
  if (changed) {
    return L(
      `已撤销：${formatInteger(changed)} 行${label}重新计入分析，${counted}。`,
      `Undone: ${enCount(changed, one, many)} ${changed === 1 ? "counts" : "count"} in the analysis again; ${counted}.`
    );
  }
  return L(
    `已撤销：这 ${formatInteger(rows)} 行仍作为${other}不计入分析，${still}。`,
    `Undone: ${rows === 1 ? "this 1 row is" : `these ${formatInteger(rows)} rows are`} still excluded as ${otherMany}; ${still}.`
  );
}

function formatFormulaErrorCount(items) {
  return items.reduce((sum, item) => sum + item.count, 0);
}

// The status bar's sentence after an import, or "".
function formatFormulaErrorStatus(items) {
  if (!items.length) return "";
  const count = formatFormulaErrorCount(items);
  return L(
    `发现 ${formatInteger(count)} 个公式错误单元格，已按空值处理，详情见“数据质量”。`,
    `Found ${enCount(count, "formula error cell", "formula error cells")}, treated as blank; see Data quality.`
  );
}

// The folded import message (v3.0.0): count is how many of the header, total and formula
// error sentences the message carries.
function formatStatusNoteSummary(sourceName, count) {
  return L(`${sourceName} 已载入，有 ${formatInteger(count)} 条说明`, `${sourceName} loaded, ${enCount(count, "note", "notes")}`);
}

// The field table's search box (v3.0.0 F2), shown for more than 20 fields.
function formatSchemaSearchPlaceholder(fieldCount) {
  return L(`搜索 ${formatInteger(fieldCount)} 个字段`, `Search ${enCount(fieldCount, "column", "columns")}`);
}

// The upload page's way back to the analysis still in memory (v3.0.0).
function formatReturnToAnalysis(sourceName) {
  return L(`回到「${sourceName}」的分析`, `Back to the analysis of “${sourceName}”`);
}

// One line of "公式错误单元格": 「利润率」：3 个（#DIV/0! 2 个、#REF! 1 个），第 3、5、9 行。
function formatFormulaError({ field, count, byError, lines, fieldIgnored = false }) {
  const errors = Object.entries(byError);
  const numbered = lines.length && !lines.some((line) => line === null);
  const more = count > lines.length;
  return L(
    `「${field}」${fieldIgnored ? "（字段已忽略）" : ""}：${formatInteger(count)} 个（${errors.map(([error, errorCount]) => `${error} ${formatInteger(errorCount)} 个`).join("、")}），${numbered ? `第 ${lines.join("、")} 行${more ? "等" : ""}` : "行号：—"}。`,
    `“${field}”${fieldIgnored ? " (column ignored)" : ""}: ${formatInteger(count)} (${errors.map(([error, errorCount]) => `${error} ${formatInteger(errorCount)}`).join(", ")}), ${numbered ? formatEnglishRows(lines, more) : "row: —"}.`
  );
}

// Longest preview of one skipped title row, counted in characters.
const SKIPPED_ROW_PREVIEW_LENGTH = 20;
// How many skipped rows the header note quotes before adding 等.
const SKIPPED_ROW_PREVIEW_COUNT = 2;

function formatSkippedRowPreview(cells) {
  const characters = [...cells.join(" ")];
  return characters.length > SKIPPED_ROW_PREVIEW_LENGTH
    ? `${characters.slice(0, SKIPPED_ROW_PREVIEW_LENGTH).join("")}…`
    : characters.join("");
}

// "表头在第 N 行，已跳过上方 K 行：「…」「…」。" for state.headerDetection, or "".
// K counts every row above the header; only non-empty rows have a preview.
function formatHeaderDetectionNote(headerDetection) {
  if (!headerDetection) return "";
  const { headerRowNumber, skippedRows = [] } = headerDetection;
  const previews = skippedRows.slice(0, SKIPPED_ROW_PREVIEW_COUNT).map((row) => row.preview);
  const more = skippedRows.length > SKIPPED_ROW_PREVIEW_COUNT;
  const skipped = headerRowNumber - 1;
  return L(
    `表头在第 ${headerRowNumber} 行，已跳过上方 ${skipped} 行${previews.length ? `：${previews.map((preview) => `「${preview}」`).join("")}${more ? "等" : ""}` : ""}。`,
    `The header is on row ${headerRowNumber}; ${skipped === 1 ? "1 row above it was" : `${skipped} rows above it were`} skipped${previews.length ? `: ${previews.map((preview) => `“${preview}”`).join(", ")}${more ? " and more" : ""}` : ""}.`
  );
}

// The texts of the kept CSV parse warnings, both languages; one is picked per call, so a file
// with many warnings builds each message once.
const CSV_WARNING_TEXTS = Lp({
  tooFewWithCells: (line, cellCount, headerCount) => `第 ${line} 行只有 ${cellCount} 列（表头有 ${headerCount} 列），缺少的单元格按空值处理。`,
  tooFewLine: (line) => `第 ${line} 行的列数少于表头，缺少的单元格按空值处理。`,
  tooFewCells: (cellCount, headerCount) => `有一条记录的列数少于表头（只有 ${cellCount} 列，表头有 ${headerCount} 列），缺少的单元格按空值处理。`,
  tooFew: () => "有一条记录的列数少于表头，缺少的单元格按空值处理。",
  quotesLine: (line) => `第 ${line} 行有不规范的引号（引号后面紧跟了其他字符），这一行可能被读错。`,
  quotes: () => "有一处不规范的引号（引号后面紧跟了其他字符），所在的记录可能被读错。",
  separators: { ";": ["分号（；）", "分号"], "\t": ["制表符", "制表符"], "|": ["竖线（|）", "竖线"] },
  oneColumn: (shown, name) => `整份文件只读出 1 列，但表头里有${shown}。如果原文件是用${name}分隔的，请另存为逗号分隔的 CSV 后重新上传。`
}, {
  tooFewWithCells: (line, cellCount, headerCount) => `Row ${line} has only ${cellCount} ${cellCount === 1 ? "column" : "columns"} (the header has ${headerCount}); the missing cells are treated as blank.`,
  tooFewLine: (line) => `Row ${line} has fewer columns than the header; the missing cells are treated as blank.`,
  tooFewCells: (cellCount, headerCount) => `A record has fewer columns than the header (only ${cellCount}, the header has ${headerCount}); the missing cells are treated as blank.`,
  tooFew: () => "A record has fewer columns than the header; the missing cells are treated as blank.",
  quotesLine: (line) => `Row ${line} has a quote mark followed by other characters, so this row may be read wrongly.`,
  quotes: () => "A quote mark is followed by other characters, so its record may be read wrongly.",
  separators: { ";": ["semicolons (;)", "semicolons"], "\t": ["tabs", "tabs"], "|": ["vertical bars (|)", "vertical bars"] },
  oneColumn: (shown, name) => `Only 1 column was read from the whole file, but the header has ${shown}. If the file is separated by ${name}, save it as a comma-separated CSV and upload it again.`
});

// The text of a kept CSV parse warning by its PapaParse code, or "" for other
// codes. line is the warning's file line or null; for TooFewFields cellCount and
// headerCount are the cells of the row and of the header (cellCount null when
// unknown); for UndetectableDelimiter separator is the character found in the header
// of a one-column file.
function formatCsvWarningMessage(code, { line = null, cellCount = null, headerCount = null, separator = "" } = {}) {
  const texts = Lpick(CSV_WARNING_TEXTS);
  if (code === "TooFewFields") {
    if (line && cellCount !== null) return texts.tooFewWithCells(line, cellCount, headerCount);
    if (line) return texts.tooFewLine(line);
    if (cellCount !== null) return texts.tooFewCells(cellCount, headerCount);
    return texts.tooFew();
  }
  if (code === "InvalidQuotes") return line ? texts.quotesLine(line) : texts.quotes();
  if (code === "UndetectableDelimiter") {
    const [shown, name] = texts.separators[separator] || [];
    return name ? texts.oneColumn(shown, name) : "";
  }
  return "";
}

// aria-label of a 查看 button for a problem-row category.
function formatProblemViewLabel(category) {
  const field = category.field;
  switch (category.kind) {
    case "missing": return L(`查看「${field}」的空值行`, `View blank rows in “${field}”`);
    case "conversion": return L(`查看「${field}」认不出的值`, `View values not recognized in “${field}”`);
    case "outlier": return L(`查看「${field}」偏大或偏小的值`, `View unusually high or low values in “${field}”`);
    case "duplicateRows": return L("查看重复行", "View duplicate rows");
    case "idDuplicate": return L(`查看「${field}」的重复行`, `View rows with duplicate IDs in “${field}”`);
    case "placeholder": return L(`查看「${field}」中「${category.key}」所在的行`, `View rows with “${category.key}” in “${field}”`);
    case "suspectedTotal": return L("查看疑似合计 / 小计行", "View possible total or subtotal rows");
    case "formulaError": return L(`查看「${field}」的公式错误`, `View formula errors in “${field}”`);
    case "dateStyle": return L(`查看「${field}」不是主写法的行`, `View rows of “${field}” not in the main date format`);
    case "lookAlike": return L(`查看「${field}」写法不一样的行`, `View rows of “${field}” with values written differently`);
    default: return L("查看读取时有疑问的行", "View rows with reading issues");
  }
}

// Option text of a problem-row category, with the quality page's count.
function formatProblemCategoryLabel(category) {
  const count = formatInteger(category.count);
  const rows = enCount(category.count, "row", "rows");
  const values = enCount(category.count, "value", "values");
  const field = category.field;
  switch (category.kind) {
    case "missing": return L(`空值 · ${field}（${count} 行）`, `Blanks · ${field} (${rows})`);
    case "conversion": return L(`认不出的数字或日期 · ${field}（${count} 个值）`, `Numbers or dates not recognized · ${field} (${values})`);
    case "outlier": return L(`偏大或偏小的值 · ${field}（${count} 行）`, `Unusually high or low values · ${field} (${rows})`);
    case "duplicateRows": return L(`整行重复（${count} 行）`, `Duplicate rows (${rows})`);
    case "idDuplicate": return L(`编号重复 · ${field}（${count} 个重复值）`, `Duplicate IDs · ${field} (${enCount(category.count, "duplicate value", "duplicate values")})`);
    case "placeholder": return L(`像是空值的写法 · ${field} · 「${category.key}」（${count} 行）`, `Values that may mean blank · ${field} · “${category.key}” (${rows})`);
    case "suspectedTotal": return L(`疑似合计 / 小计行（${count} 行）`, `Possible total or subtotal rows (${rows})`);
    case "formulaError": return L(`公式错误 · ${field}（${count} 个值）`, `Formula errors · ${field} (${values})`);
    case "dateStyle": return L(`日期写法 · ${field}（${count} 个值）`, `Mixed date formats · ${field} (${values})`);
    case "lookAlike": return L(`写法不一样 · ${field}（${count} 行）`, `Variants · ${field} (${rows})`);
    default: return L(`读取时有疑问（${count} 处）`, `Reading issues (${enCount(category.count, "issue", "issues")})`);
  }
}

// Viewer summary: the listed count in its unit, then "第 X / Y 页". Duplicate rows
// say how the quality count (without each group's kept first row) relates to the list.
function formatProblemRowsSummary(category, list, page, pageCount) {
  const pages = L(`第 ${formatInteger(page)} / ${formatInteger(pageCount)} 页`, `page ${formatInteger(page)} of ${formatInteger(pageCount)}`);
  const listedCount = list.items.length;
  const listed = formatInteger(listedCount);
  if (category.kind === "duplicateRows") {
    return L(
      `重复 ${formatInteger(list.count)} 行（${formatInteger(list.groupCount)} 组，共 ${listed} 行，每组第一行保留），${pages}`,
      `${enCount(list.count, "duplicate row", "duplicate rows")} (${enCount(list.groupCount, "group", "groups")}, ${enCount(listedCount, "row", "rows")} in all, the first row of each group kept), ${pages}`
    );
  }
  if (category.kind === "conversion" || category.kind === "formulaError" || category.kind === "dateStyle") return L(`共 ${listed} 个值，${pages}`, `${enCount(listedCount, "value", "values")} in all, ${pages}`);
  if (category.kind === "parseWarnings") return L(`共 ${listed} 处，${pages}`, `${enCount(listedCount, "issue", "issues")} in all, ${pages}`);
  return L(`共 ${listed} 行，${pages}`, `${enCount(listedCount, "row", "rows")} in all, ${pages}`);
}

// "6、19", plus " 等 15 个" when totalCount is more than the numbers shown. Any row
// without a number shows "—" instead of a partial list.
function formatRowNumberList(rowNumbers, totalCount = rowNumbers.length) {
  if (!rowNumbers.length || rowNumbers.some((rowNumber) => rowNumber === null)) return "—";
  if (totalCount <= rowNumbers.length) return rowNumbers.join(Lpick(LIST_SEPARATOR));
  return L(`${rowNumbers.join("、")} 等 ${totalCount} 个`, `${rowNumbers.join(", ")} and more, ${totalCount} in all`);
}

// The note on one duplicate ID value: how many of its rows repeat a whole row.
function formatIdenticalRows(rowCount, groupCount) {
  if (!rowCount) return "";
  return L(
    `其中 ${rowCount} 行内容完全相同${groupCount > 1 ? `（分 ${groupCount} 组）` : ""}`,
    `${rowCount} of these rows are identical${groupCount > 1 ? ` (in ${groupCount} groups)` : ""}`
  );
}

function formatIdDuplicateOverflow(totalCount, listedCount) {
  if (totalCount <= listedCount) return "";
  return L(`共 ${totalCount} 个重复值，只列出前 ${listedCount} 个`, `${totalCount} duplicate values in all; only the first ${listedCount} are listed`);
}

// Names the ID fields without duplicates; afterOthers when other ID fields were reported first.
function formatCleanIdFieldsNote(fields, afterOthers = false) {
  return L(
    `${afterOthers ? "其余" : ""}类型为 ID 的字段没有编号重复（已检查：${fields.join("、")}）`,
    `${afterOthers ? "The other columns" : "Columns"} of type ID have no duplicate IDs (checked: ${fields.join(", ")})`
  );
}

// Shown instead of the value list for a mostly repeated ID field (plain text).
function formatMostlyRepeatedIdHint({ field, distinctCount, nonMissingCount, uniqueRatio }) {
  return L(
    `「${field}」的值大多重复出现（${formatInteger(distinctCount)} 个不同值，共 ${formatInteger(nonMissingCount)} 行，唯一率 ${formatPercent(uniqueRatio)}），像是同一对象的多条记录，不逐个列出。如果这一列应该每行唯一，请检查数据；不需要检查的话，可以在字段配置里改类型。`,
    `Most values of “${field}” repeat (${enCount(distinctCount, "distinct value", "distinct values")} in ${enCount(nonMissingCount, "row", "rows")}, ${formatPercent(uniqueRatio)} unique); they look like several records of the same things, so they are not listed one by one. If each row should have its own value, check the data; if not, you can change the type in Column settings.`
  );
}

// Shown with every maybe-valid placeholder key; the dash note only for 「-」. The English names
// only the writings this column has (R11): they come from the data, not from a fixed list.
const PLACEHOLDER_MAYBE_NOTE = Lp(
  "「无」「暂无」「-」「/」「NA」也可能是有效值（例如“有/无”选项、北美地区），请确认这一列里它确实表示没有数据。",
  (keys) => {
    // The examples of the Chinese sentence, for the writings found: NA can be North America, the
    // others (无, 暂无, -, /) an answer to a yes/no question.
    const examples = [keys.some((key) => key !== "NA") ? "a yes/no answer" : "", keys.includes("NA") ? "NA for North America" : ""].filter(Boolean).join(", or ");
    return keys.length === 1
      ? `“${keys[0]}” may also be a real value (for example ${examples}); make sure that in this column it means there is no data.`
      : `${keys.map((key) => `“${key}”`).join(", ")} may also be real values (for example ${examples}); make sure that in this column they mean there is no data.`;
  }
);
const PLACEHOLDER_DASH_NOTE = Lp("在会计格式里，「-」通常表示 0。", "In accounting formats, “-” usually means 0.");

// 「N/A」, plus "（原文：Ｎ／Ａ、N/A）" when a raw text differs from the key.
function formatPlaceholderValue({ key, rawValues }) {
  if (!rawValues.some((value) => value !== key)) return L(`「${key}」`, `“${key}”`);
  return L(`「${key}」（原文：${rawValues.join("、")}）`, `“${key}” (written as ${rawValues.join(", ")})`);
}

// "（第 15 行）", "（第 2、…、11 行等）" when more rows hold the key, "（行号：—）" without row numbers.
function formatPlaceholderRows(rowNumbers, count) {
  const list = formatRowNumberList(rowNumbers);
  if (list === "—") return L("（行号：—）", " (row: —)");
  const more = count > rowNumbers.length;
  return L(`（第 ${list} 行${more ? "等" : ""}）`, ` (${formatEnglishRows(rowNumbers, more)})`);
}

// "缺失值 0 → 1，类型仍为数值", or the new type when the confirmation changes it.
function formatPlaceholderChange(field, missingBefore, typeKeyBefore, preview) {
  const same = preview.typeKey === typeKeyBefore;
  return L(
    `空值 ${formatInteger(missingBefore)} → ${formatInteger(preview.missingCount)}，${same ? `类型仍为${fieldTypeLabel(typeKeyBefore)}` : `「${field}」会识别为${fieldTypeLabel(preview.typeKey)}（现在是${fieldTypeLabel(typeKeyBefore)}）`}`,
    `blanks ${formatInteger(missingBefore)} → ${formatInteger(preview.missingCount)}, ${same ? `type still ${fieldTypeLabel(typeKeyBefore)}` : `“${field}” will be read as ${fieldTypeLabel(preview.typeKey)} (now ${fieldTypeLabel(typeKeyBefore)})`}`
  );
}

function formatPendingPlaceholder(entry, change) {
  return L(
    `「${entry.field}」：${formatPlaceholderValue(entry)}${formatInteger(entry.count)} 个${formatPlaceholderRows(entry.rowNumbers, entry.count)}。按空值处理后，${change}。`,
    `“${entry.field}”: ${formatPlaceholderValue(entry)}, ${enCount(entry.count, "cell", "cells")}${formatPlaceholderRows(entry.rowNumbers, entry.count)}. Treated as blank: ${change}.`
  );
}

function formatConfirmedPlaceholder(entry) {
  return L(
    `「${entry.field}」${entry.fieldIgnored ? "（字段已忽略）" : ""}：${formatPlaceholderValue(entry)}已按空值处理（${formatInteger(entry.count)} 个）。`,
    `“${entry.field}”${entry.fieldIgnored ? " (column ignored)" : ""}: ${formatPlaceholderValue(entry)} treated as blank (${enCount(entry.count, "cell", "cells")}).`
  );
}

function formatPlaceholderNotes(keys) {
  const maybe = keys.filter((key) => placeholderTier(key) === "maybe");
  if (!maybe.length) return "";
  const dash = keys.includes("-");
  return L(
    `${PLACEHOLDER_MAYBE_NOTE[0]}${dash ? PLACEHOLDER_DASH_NOTE[0] : ""}`,
    `${PLACEHOLDER_MAYBE_NOTE[1](maybe)}${dash ? ` ${PLACEHOLDER_DASH_NOTE[1]}` : ""}`
  );
}

// v3.1.0: what treating a spelling as empty (or undoing it) changed. before / after are
// { missing, failures, typeKey, listed } of the field: its missing count, values it cannot
// read, applied type and whether the ID duplicates list the key.
function formatPlaceholderCounts(key, before, after, confirmed) {
  const reads = (typeKey) => typeKey === "numeric" || typeKey === "date";
  const parts = [L(`空值 ${formatInteger(before.missing)} → ${formatInteger(after.missing)}`, `blanks ${formatInteger(before.missing)} → ${formatInteger(after.missing)}`)];
  if (reads(before.typeKey) || reads(after.typeKey)) parts.push(L(`认不出 ${formatInteger(before.failures)} → ${formatInteger(after.failures)}`, `not recognized ${formatInteger(before.failures)} → ${formatInteger(after.failures)}`));
  if (before.typeKey === "category" && after.typeKey === "category") {
    parts.push(confirmed
      ? L(`分类里不再有「${key}」这一类`, `the value counts no longer have “${key}”`)
      : L(`分类里重新有「${key}」这一类`, `the value counts have “${key}” again`));
  }
  if (confirmed && before.listed && !after.listed) parts.push(L(`编号重复里不再列出「${key}」`, `duplicate IDs no longer list “${key}”`));
  if (!confirmed && !before.listed && after.listed) parts.push(L(`编号重复里重新列出「${key}」`, `duplicate IDs list “${key}” again`));
  return parts.join(L("，", ", "));
}

// "「金额」现在识别为数值（原来是分类）", without the full stop.
function formatPlaceholderTypeChange(field, typeKeyBefore, typeKeyAfter) {
  return L(
    `「${field}」现在识别为${fieldTypeLabel(typeKeyAfter)}（原来是${fieldTypeLabel(typeKeyBefore)}）`,
    `“${field}” is now read as ${fieldTypeLabel(typeKeyAfter)} (was ${fieldTypeLabel(typeKeyBefore)})`
  );
}

// "行数还是 42", or "参与分析的行还是 40" when rows are left out of the analysis (v3.1.0):
// rowCount is the rows counted in the analysis.
function formatPlaceholderRowCount(rowCount, rowsLeftOut) {
  return L(
    `${rowsLeftOut ? "参与分析的行" : "行数"}还是 ${formatInteger(rowCount)}`,
    rowsLeftOut ? `the rows in analysis are still ${formatInteger(rowCount)}` : `there are still ${enCount(rowCount, "row", "rows")}`
  );
}

// count: the cells the key covers; rowCount: the rows of the analysis, which do not change.
function formatPlaceholderConfirmMessage({ field, key, count, before, after, rowCount, rowsLeftOut = false }) {
  const cells = formatInteger(count);
  const typeChanged = before.typeKey !== after.typeKey;
  const counts = formatPlaceholderCounts(key, before, after, true);
  const rowCountText = formatPlaceholderRowCount(rowCount, rowsLeftOut);
  const typeText = typeChanged ? formatPlaceholderTypeChange(field, before.typeKey, after.typeKey) : "";
  return L(
    `已把「${field}」的 ${cells} 个「${key}」按空值处理：${counts}。只是把这 ${cells} 格当作空，${count === 1 ? "这一行" : "这些行"}没有删，其他字段照常参与分析，所以${rowCountText}。${typeChanged ? `${typeText}。` : after.typeKey === "numeric" ? `平均值等统计不变，因为「${key}」本来就没有算进去。` : after.typeKey === "date" ? `按日期的趋势不变，因为「${key}」本来就没有算进去。` : ""}`,
    `Treated ${cells} “${key}” in “${field}” as blank: ${counts}. ${count === 1 ? "Only this 1 cell is" : `Only these ${cells} cells are`} treated as blank; ${count === 1 ? "its row is" : "their rows are"} not deleted and the other columns still count, so ${rowCountText}.${typeChanged ? ` ${typeText}.` : after.typeKey === "numeric" ? ` Averages and other stats do not change, because “${key}” was never counted in them.` : after.typeKey === "date" ? ` Trends by date do not change, because “${key}” was never counted in them.` : ""}`
  );
}

function formatPlaceholderUndoMessage({ field, key, count, before, after, rowCount, rowsLeftOut = false, ignored }) {
  if (ignored) {
    return L(
      `已撤销「${field}」的「${key}」。「${field}」已忽略，撤销只影响以后改回其他类型时的结果。`,
      `Undone “${key}” in “${field}”. “${field}” is ignored, so this only matters if its type is changed back later.`
    );
  }
  const typeChanged = before.typeKey !== after.typeKey;
  const counts = formatPlaceholderCounts(key, before, after, false);
  const rowCountText = formatPlaceholderRowCount(rowCount, rowsLeftOut);
  const typeText = typeChanged ? formatPlaceholderTypeChange(field, before.typeKey, after.typeKey) : "";
  return L(
    `已撤销「${field}」的「${key}」：${counts}。这 ${formatInteger(count)} 格重新按原来的文字处理，${rowCountText}。${typeChanged ? `${typeText}。` : ""}`,
    `Undone “${key}” in “${field}”: ${counts}. ${count === 1 ? "This 1 cell is" : `These ${formatInteger(count)} cells are`} read as their original text again, and ${rowCountText}.${typeChanged ? ` ${typeText}.` : ""}`
  );
}

// entries: [{ field, key }] with the fields whose type changed first; typeChanges: [{ field,
// typeKey, previousTypeKey }] in the same order.
function formatAllPlaceholdersMessage(entries, typeChanges, rowCount, rowsLeftOut = false) {
  const types = typeChanges.map((change) => formatPlaceholderTypeChange(change.field, change.previousTypeKey, change.typeKey));
  const rowCountText = formatPlaceholderRowCount(rowCount, rowsLeftOut);
  return L(
    `已把 ${formatInteger(entries.length)} 处常见写法按空值处理（${entries.map((entry) => `「${entry.field}」的「${entry.key}」`).join("、")}）。${types.length ? `${types.join("；")}。` : ""}只是把这些格当作空，行都没有删，所以${rowCountText}。`,
    `Treated ${enCount(entries.length, "common one", "common ones")} as blank (${entries.map((entry) => `“${entry.key}” in “${entry.field}”`).join(", ")}). ${types.length ? `${types.join("; ")}. ` : ""}Only these cells are treated as blank and no rows are deleted, so ${rowCountText}.`
  );
}

// The notes of "日期写法混用", "常量列" and "写法细节不同的值" in reports; index.html has the same
// sentences.
const DATE_STYLE_NOTE = Lp(
  "这些写法本工具都能读成同一个日期，但在 Excel 里文本日期不能按日期排序、筛选，导入别的系统时可能读错。",
  "This tool reads all of these as the same date, but Excel cannot sort or filter dates stored as text, and other systems may read them wrongly."
);
const CONSTANT_FIELD_NOTE = Lp(
  "这些列的非空值全都一样。不一定是错误：可能是导出时带上的固定列，也可能是漏填或填错。",
  "All non-blank values in these columns are the same. That is not always an error: it may be a fixed column added on export, or values left out or filled in wrongly."
);
const LOOK_ALIKE_NOTE = Lp(
  "首尾空格：本工具已按去掉空格处理，但导出给别的系统、用 VLOOKUP 时会对不上。同一个值有多种写法：本工具把它们当成不同的值，分类统计、筛选和编号重复都会分开算。只有不可见字符：格子看起来是空的，但本工具不按空值处理，分类统计和筛选里会多出一个看不见的值。",
  "Leading or trailing spaces: this tool ignores them, but they will not match when exported to other systems or used with VLOOKUP. One value written in several ways: this tool counts them as different values, so value counts, filters and duplicate IDs count them separately. Only invisible characters: the cell looks empty, but this tool does not treat it as blank, so value counts and filters show an extra value you cannot see."
);
// Names of the date styles, as a user would recognise them in Excel.
const DATE_STYLE_NAMES = Object.freeze({
  excelCell: Lp("Excel 日期单元格", "Excel date cell"),
  ymdDash: Lp("文本 年-月-日", "text YYYY-MM-DD"),
  ymdSlash: Lp("文本 年/月/日", "text YYYY/MM/DD"),
  ymdDot: Lp("文本 年.月.日", "text YYYY.MM.DD"),
  chinese: Lp("文本 中文年月日", "text with Chinese year/month/day characters"),
  dmy: Lp("文本 日/月/年", "text DD/MM/YYYY"),
  mdy: Lp("文本 月/日/年", "text MM/DD/YYYY"),
  dayMonthSame: Lp("文本 日和月相同（看不出先后）", "text with the same day and month (order unclear)"),
  english: Lp("文本 英文月份", "text with English month names"),
  yearOnly: Lp("只写年份", "year only")
});

const VISIBLE_TEXT_NAMES = Object.freeze({
  "\t": Lp("[制表符]", "[tab]"), "\n": Lp("[换行]", "[line break]"), "\r": Lp("[换行]", "[line break]"), "\u00A0": Lp("[不换行空格]", "[no-break space]"), "\u3000": Lp("[全角空格]", "[full-width space]"),
  "\u00AD": Lp("[软连字符]", "[soft hyphen]"), "\u200E": Lp("[方向符]", "[direction mark]"), "\u200F": Lp("[方向符]", "[direction mark]"),
  "\u200B": Lp("[零宽字符]", "[zero-width character]"), "\u200C": Lp("[零宽字符]", "[zero-width character]"), "\u200D": Lp("[零宽字符]", "[zero-width character]"), "\u2060": Lp("[零宽字符]", "[zero-width character]"), "\uFEFF": Lp("[零宽字符]", "[zero-width character]")
});
const VISIBLE_TEXT_OTHER = Lp("[空白]", "[whitespace]");
const VISIBLE_TEXT_PATTERN = new RegExp(`[\\s${INVISIBLE_CHARACTER_CLASS}]`, "g");

// A text with its whitespace and invisible characters shown: "␣" for a half-width space at
// an end or in a run (a single space inside stays), a name in brackets for the others.
// Plain text; callers escape it for their output format.
function formatVisibleText(text) {
  return String(text)
    .replace(/^ +| +$| {2,}/g, (spaces) => "␣".repeat(spaces.length))
    .replace(VISIBLE_TEXT_PATTERN, (character) => (character === " " ? " " : Lpick(VISIBLE_TEXT_NAMES[character] || VISIBLE_TEXT_OTHER)));
}

function formatDateStylePart({ key, count, sample }) {
  return L(
    `${DATE_STYLE_NAMES[key][0]} ${formatInteger(count)} 个（如 ${formatVisibleText(sample)}）`,
    `${DATE_STYLE_NAMES[key][1]}: ${formatInteger(count)} (e.g. ${formatVisibleText(sample)})`
  );
}

// 「入职日期」：主写法是 Excel 日期单元格 24 个（如 2016-04-01）；另有 6 行是别的写法：文本 年/月/日 2 个（如 2021/3/5）、…。
function formatDateStyleLine({ field, styles, otherCount }) {
  const [main, ...others] = styles;
  return L(
    `「${field}」：主写法是 ${formatDateStylePart(main)}；另有 ${formatInteger(otherCount)} 行是别的写法：${others.map(formatDateStylePart).join("、")}。`,
    `“${field}”: the main format is ${formatDateStylePart(main)}; ${otherCount === 1 ? "1 other row uses another format" : `${formatInteger(otherCount)} other rows use other formats`}: ${others.map(formatDateStylePart).join(", ")}.`
  );
}
// 「备注」：只有一种取值「正常」（6 行），另有 3 行缺失。
function formatConstantField({ field, value, count, missingCount }) {
  return L(
    `「${field}」：只有一种取值「${formatVisibleText(value)}」（${formatInteger(count)} 行）${missingCount ? `，另有 ${formatInteger(missingCount)} 行是空值` : ""}。`,
    `“${field}”: only one value, “${formatVisibleText(value)}” (${enCount(count, "row", "rows")})${missingCount ? `, and ${enCount(missingCount, "row is", "rows are")} blank` : ""}.`
  );
}

const LOOK_ALIKE_REASON_NAMES = Object.freeze({
  compat: Lp("全角半角或兼容写法", "full-width, half-width or compatibility forms"),
  case: Lp("大小写", "upper and lower case"),
  invisible: Lp("不可见字符", "invisible characters"),
  innerSpace: Lp("中间空格多少", "spaces inside")
});

function formatLookAlikeReasons(reasons) {
  return reasons.map((reason) => Lpick(LOOK_ALIKE_REASON_NAMES[reason])).join(Lpick(LIST_SEPARATOR));
}

// 「HD」2 个 / 「ＨＤ」1 个 / 「hd」1 个（全角半角或兼容写法、大小写）
function formatLookAlikeGroup({ variants, variantCount = variants.length, reasons }) {
  const more = variantCount > variants.length;
  return L(
    `${variants.map(({ text, count }) => `「${formatVisibleText(text)}」${formatInteger(count)} 个`).join(" / ")}${more ? ` 等 ${formatInteger(variantCount)} 种` : ""}（${formatLookAlikeReasons(reasons)}）`,
    `${variants.map(({ text, count }) => `“${formatVisibleText(text)}” ${formatInteger(count)}`).join(" / ")}${more ? ` and more, ${formatInteger(variantCount)} ways in all` : ""} (${formatLookAlikeReasons(reasons)})`
  );
}

// 「区域」共 18 行：首尾空格 2 格（如「华东␣」）；只有不可见字符 1 格；3 组写法不同：…；…。
function formatLookAlikeLine({ field, edgeSpaceCount, edgeSample, invisibleOnlyCount = 0, groups, groupCount, rowCount }) {
  const parts = [];
  if (edgeSpaceCount) {
    parts.push(L(
      `首尾空格 ${formatInteger(edgeSpaceCount)} 格（如「${formatVisibleText(edgeSample)}」）`,
      `leading or trailing spaces in ${enCount(edgeSpaceCount, "cell", "cells")} (e.g. “${formatVisibleText(edgeSample)}”)`
    ));
  }
  if (invisibleOnlyCount) {
    parts.push(L(`只有不可见字符 ${formatInteger(invisibleOnlyCount)} 格`, `only invisible characters in ${enCount(invisibleOnlyCount, "cell", "cells")}`));
  }
  if (groupCount) {
    const cut = groupCount > groups.length;
    const listed = groups.map(formatLookAlikeGroup);
    parts.push(L(
      `${formatInteger(groupCount)} 组写法不一样${cut ? `，只列出前 ${groups.length} 组` : ""}：${listed.join("；")}`,
      `${groupCount === 1 ? "1 value is" : `${formatInteger(groupCount)} values are`} written differently${cut ? `, only the first ${groups.length} listed` : ""}: ${listed.join("; ")}`
    ));
  }
  return L(`「${field}」共 ${formatInteger(rowCount)} 行：${parts.join("；")}。`, `“${field}”, ${enCount(rowCount, "row", "rows")}: ${parts.join("; ")}.`);
}

// The 说明 of a problem row (one call per listed row): a template pair, only the chosen one is built.
const DATE_STYLE_ROW_NOTE = Lp(
  (style, main) => `${style}（主写法是 ${main}）`,
  (style, main) => `${style} (the main format is ${main})`
);

function formatDateStyleRowNote(key, mainKey) {
  return Lpick(DATE_STYLE_ROW_NOTE)(Lpick(DATE_STYLE_NAMES[key]), Lpick(DATE_STYLE_NAMES[mainKey]));
}

// The words of a look-alike problem row's 说明 (one call per listed row).
const LOOK_ALIKE_ROW_NOTE_WORDS = Lp({
  group: (position, reasons) => `第 ${position} 组（${reasons}）`,
  invisibleOnly: "只有不可见字符",
  edge: "首尾空格",
  joiner: "，",
  value: (parts, text) => `${parts}：「${text}」`
}, {
  group: (position, reasons) => `group ${position} (${reasons})`,
  invisibleOnly: "only invisible characters",
  edge: "leading or trailing spaces",
  joiner: ", ",
  value: (parts, text) => `${parts}: “${text}”`
});

// The 说明 of a problem row: the cell's own text with its whitespace shown, since the table
// shows the value trimmed.
function formatLookAlikeRowNote(value, group, hasEdge, invisibleOnly) {
  const words = Lpick(LOOK_ALIKE_ROW_NOTE_WORDS);
  const parts = [];
  if (group) parts.push(words.group(group.position + 1, formatLookAlikeReasons(group.reasons)));
  if (invisibleOnly) parts.push(words.invisibleOnly);
  if (hasEdge) parts.push(words.edge);
  return words.value(parts.join(words.joiner), formatVisibleText(value));
}

// The 处理 of a standardized cell (v2.10.0), for the check table's samples and the change list.
const STANDARDIZE_ACTION_LABELS = Object.freeze({
  date: Lp("日期改写为 YYYY-MM-DD", "Date rewritten as YYYY-MM-DD"),
  number: Lp("数字改写为纯数字", "Number rewritten as a plain number"),
  placeholder: Lp("已按空值处理的写法，清空", "Treated as blank, cleared"),
  trim: Lp("去首尾空格", "Leading and trailing spaces removed"),
  keptTime: Lp("带时间，保持原值", "Has a time, kept as is"),
  keptPercent: Lp("百分比，不改写", "Percentage, not rewritten"),
  keptScientific: Lp("科学计数法，不改写", "Scientific notation, not rewritten"),
  keptCurrencyMix: Lp("货币符号不一致，保持原值", "Mixed currency symbols, kept as is"),
  keptPendingPlaceholder: Lp("像是空值的写法，未确认", "May mean blank, not confirmed"),
  keptDateUnreadable: Lp("读不成日期，保持原值", "Not recognized as a date, kept as is"),
  keptDateOrder: Lp("日月顺序无法确定，保持原值", "Day and month order unclear, kept as is"),
  keptNumberUnreadable: Lp("读不成数字，保持原值", "Not recognized as a number, kept as is"),
  keptUnsafe: Lp("无法安全改写，保持原值", "Cannot be rewritten safely, kept as is")
});

function formatStandardizeAction(action) {
  const pair = STANDARDIZE_ACTION_LABELS[action];
  return pair ? Lpick(pair) : "";
}

const STANDARDIZE_ORDER_HINT = Lp(
  "导出后这一列剩下的日或月在前的写法没有了确定顺序的依据，再次导入本工具时可能读不成日期。",
  "After export, the dates left in this column no longer show whether the day or the month comes first, so this tool may not read them as dates when the file is imported again."
);

// Why 检查标准化改动 is unavailable, or "" when it is.
function formatStandardizedExportGate(hasRows, hasDrafts, analysisDone) {
  if (!hasRows) return L("请先上传并解析有效的数据文件。", "Upload and read a valid data file first.");
  if (hasDrafts) return L("字段配置有未应用的修改，请先应用或撤销后再检查。", "Column settings have changes that are not applied; apply or undo them, then check again.");
  if (!analysisDone) return L("分析尚未完成。", "The analysis is not finished yet.");
  return "";
}

// The notes under a field name in the check table: read by year, a unit left out of the
// export, the day-month order a second import may not find (D8).
function formatStandardizeFieldNotes(count) {
  const notes = [];
  if (count.kind === "year") notes.push(L("按年份读取，不改写日期", "Read as years, dates not rewritten"));
  if (count.unit && count.number) notes.push(L(`单位「${count.unit}」不写进导出文件，表头仍是「${count.field}」`, `The unit “${count.unit}” is not written to the exported file; the header stays “${count.field}”`));
  if (count.orderHint) notes.push(Lpick(STANDARDIZE_ORDER_HINT));
  return notes;
}

// 2（如「2024-05-08 10:30」、「待定」）
function formatStandardizeKept(count) {
  if (!count.kept) return "—";
  return L(
    `${formatInteger(count.kept)}（如${count.keptSamples.map((sample) => `「${formatVisibleText(sample)}」`).join("、")}）`,
    `${formatInteger(count.kept)} (e.g. ${count.keptSamples.map((sample) => `“${formatVisibleText(sample)}”`).join(", ")})`
  );
}

// The status after 导出改动清单: all of it, or the first `exported` of `total` (D13).
function formatStandardizedChangesStatus(exported, total, fileName) {
  return exported < total
    ? L(`只导出了前 ${formatInteger(exported)} 条，共 ${formatInteger(total)} 条：${fileName}`, `Exported only the first ${formatInteger(exported)} of ${enCount(total, "change", "changes")}: ${fileName}`)
    : L(`已导出改动清单 ${formatInteger(total)} 条：${fileName}`, `Exported the change list, ${enCount(total, "change", "changes")}: ${fileName}`);
}

// The sentence above the check table.
function formatStandardizedSummary({ changed, kept, quietCount, ignoredCount }) {
  const ignored = ignoredCount
    ? L(`已忽略的 ${formatInteger(ignoredCount)} 个字段原样导出。`, `${enCount(ignoredCount, "ignored column is", "ignored columns are")} exported as is.`)
    : "";
  if (!changed && !kept) {
    return L(`没有需要统一写法的单元格，导出的数据与原始数据相同。${ignored}`, `No cells need standardizing; the exported data is the same as the original.${ignored ? ` ${ignored}` : ""}`);
  }
  const quiet = quietCount ? L(`其余 ${formatInteger(quietCount)} 个字段没有改动。`, `The other ${enCount(quietCount, "column has", "columns have")} no changes.`) : "";
  return L(
    `共改写 ${formatInteger(changed)} 格，保持原值 ${formatInteger(kept)} 格。${quiet}${ignored}`,
    `Rewrote ${enCount(changed, "cell", "cells")}, kept ${enCount(kept, "cell", "cells")} as is.${[quiet, ignored].filter(Boolean).map((text) => ` ${text}`).join("")}`
  );
}
function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
