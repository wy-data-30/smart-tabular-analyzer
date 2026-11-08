// 显示格式化：数值、整数、百分比和字段格式的展示文本，带单位的字段标签，表头位置说明，ID 重复值的行号列表和说明文字，疑似缺失值、按单位读取的数字、疑似数值列、疑似合计 / 小计行、公式错误单元格、日期写法、常量列和写法细节不同的值的说明文字，CSV 解析警告的中文消息，问题行的按钮、类别和摘要文字，以及 HTML 转义。

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

// Plain text such as "3（N/A、待定、约500）"; callers escape it for their output format.
function formatConversionFailures(count, samples = [], countText = String(count ?? 0)) {
  return count && samples.length ? `${countText}（${samples.join("、")}）` : countText;
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
  const unitText = unit ? `（${unit}）` : "";
  return getNumericFormat(field) === "percent"
    ? `${fallbackLabel}${unitText}（百分比）`
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
  return bracketed || bare ? field : `${field}（${unit}）`;
}

function formatFieldLabel(field) {
  return formatFieldLabelWithUnit(field, getFieldNumberUnit(field));
}

// One line of "按单位读取的数字" for a buildNumberConversionSummary() item.
function formatNumberConversion(item) {
  const parts = [];
  if (item.unitCount) parts.push(`按"${item.unit}"读取了 ${formatInteger(item.unitCount)} 个值（${item.unitSamples.join("、")}）`);
  if (item.magnitudeCount) parts.push(`按万/亿换算了 ${formatInteger(item.magnitudeCount)} 个值`);
  if (item.fullwidthCount) parts.push(`${formatInteger(item.fullwidthCount)} 个全角数字已按半角读取`);
  if (item.conflictUnits.length) {
    const units = item.conflictUnits.map(({ unit, count }) => `${unit}（${formatInteger(count)} 个）`);
    parts.push(`同时出现 ${units.slice(0, -1).join("、")}和 ${units[units.length - 1]}，没有按单位读取，这些值不能读成数字。`);
  }
  return `「${item.field}」：${parts.join("；")}`;
}

// One line of "疑似数值列" for a state.numericLooking item.
function formatNumericLookingHint(item) {
  return `「${item.field}」：去掉缺失值和疑似缺失值后，${formatInteger(item.valueCount)} 个值里有 ${formatInteger(item.numericCount)} 个能读成数字（${formatPercent(item.ratio)}），其余的如 ${item.samples.join("、")}。如果这一列应该是数字，请修正这些值，或在字段配置里把类型改成数值。`;
}

// How many columns a suspected total's line names before "等 N 列".
const TOTAL_MATCHED_FIELD_LIMIT = 5;
// How many row numbers the status bar names before 等.
const STATUS_ROW_NUMBER_LIMIT = 5;

// The notes of "疑似合计 / 小计行" and "公式错误单元格" in reports; index.html has the same
// sentences.
const SUSPECTED_TOTAL_NOTE = "这些行目前仍按普通数据参与统计，其中的数字会被重复计入求和、平均值和异常值。";
const FORMULA_ERROR_NOTE = "这些单元格的公式算出了错误，按缺失值处理（已计入缺失值）。整行都是错误值的行不会导入，也不在这里列出。";

function formatTotalMatchedFields(fields) {
  if (fields.length <= TOTAL_MATCHED_FIELD_LIMIT) return fields.join("、");
  return `${fields.slice(0, TOTAL_MATCHED_FIELD_LIMIT).join("、")} 等 ${fields.length} 列`;
}

// "第 5–32 行之和（28 行）", "第 2–13 行中的明细之和（9 行）" when other rows lie in between,
// "第 9 行（只有 1 行）", or the range in words for a table without row numbers.
function formatTotalAddends({ scope, firstLine, lastLine, rowCount, contiguous }) {
  const count = formatInteger(rowCount);
  if (firstLine === null || lastLine === null) return `${scope === "all" ? "全表" : "上方"} ${count} 行明细之和`;
  if (rowCount === 1) return `第 ${firstLine} 行（只有 1 行）`;
  const lines = `第 ${firstLine}–${lastLine} 行`;
  return contiguous ? `${lines}之和（${count} 行）` : `${lines}中的明细之和（${count} 行）`;
}

// "销售额、成本等于第 5–32 行之和（28 行）": the 说明 of a suspected total's problem row.
function formatSuspectedTotalMatches(item) {
  return item.matches.map((match) => `${formatTotalMatchedFields(match.fields)}等于${formatTotalAddends(match)}`).join("；");
}

// One line of "疑似合计 / 小计行".
function formatSuspectedTotal(item) {
  return `${item.line === null ? "有一行" : `第 ${item.line} 行`}「${item.label}」：${formatSuspectedTotalMatches(item)}。`;
}

// "共 N 行，只列出前 20 行。" when a list of suspected totals is cut short, or "".
function formatSuspectedTotalOverflow(count) {
  return count > SUSPECTED_TOTAL_LIST_LIMIT ? `共 ${formatInteger(count)} 行，只列出前 ${SUSPECTED_TOTAL_LIST_LIMIT} 行。` : "";
}

// The status bar's sentence after an import, or "".
function formatSuspectedTotalStatus(items) {
  if (!items.length) return "";
  const ending = "，目前仍计入统计，详情见“数据质量”。";
  if (items.length === 1) {
    const [{ line, label }] = items;
    return `${line === null ? "有一行" : `第 ${line} 行`}「${formatSkippedRowPreview([label])}」疑似合计行${ending}`;
  }
  const lines = items.map((item) => item.line);
  const shown = lines.some((line) => line === null)
    ? ""
    : `（第 ${lines.slice(0, STATUS_ROW_NUMBER_LIMIT).join("、")} 行${lines.length > STATUS_ROW_NUMBER_LIMIT ? "等" : ""}）`;
  return `发现 ${formatInteger(items.length)} 行疑似合计 / 小计${shown}${ending}`;
}

function formatFormulaErrorCount(items) {
  return items.reduce((sum, item) => sum + item.count, 0);
}

// The status bar's sentence after an import, or "".
function formatFormulaErrorStatus(items) {
  return items.length ? `发现 ${formatInteger(formatFormulaErrorCount(items))} 个公式错误单元格，已按缺失值处理，详情见“数据质量”。` : "";
}

// One line of "公式错误单元格": 「利润率」：3 个（#DIV/0! 2 个、#REF! 1 个），第 3、5、9 行。
function formatFormulaError({ field, count, byError, lines, fieldIgnored = false }) {
  const errors = Object.entries(byError).map(([error, errorCount]) => `${error} ${formatInteger(errorCount)} 个`).join("、");
  const rows = !lines.length || lines.some((line) => line === null)
    ? "行号：—"
    : `第 ${lines.join("、")} 行${count > lines.length ? "等" : ""}`;
  return `「${field}」${fieldIgnored ? "（字段已忽略）" : ""}：${formatInteger(count)} 个（${errors}），${rows}。`;
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
  const quoted = skippedRows.slice(0, SKIPPED_ROW_PREVIEW_COUNT).map((row) => `「${row.preview}」`).join("");
  const more = skippedRows.length > SKIPPED_ROW_PREVIEW_COUNT ? "等" : "";
  return `表头在第 ${headerRowNumber} 行，已跳过上方 ${headerRowNumber - 1} 行${quoted ? `：${quoted}${more}` : ""}。`;
}

// Chinese text of a kept CSV parse warning by its PapaParse code, or "" for other
// codes. line is the warning's file line or null; for TooFewFields cellCount and
// headerCount are the cells of the row and of the header (cellCount null when
// unknown); for UndetectableDelimiter separator is the character found in the header
// of a one-column file.
function formatCsvWarningMessage(code, { line = null, cellCount = null, headerCount = null, separator = "" } = {}) {
  if (code === "TooFewFields") {
    const missingCells = "缺少的单元格按空值处理。";
    if (line && cellCount !== null) return `第 ${line} 行只有 ${cellCount} 列（表头有 ${headerCount} 列），${missingCells}`;
    if (line) return `第 ${line} 行的列数少于表头，${missingCells}`;
    if (cellCount !== null) return `有一条记录的列数少于表头（只有 ${cellCount} 列，表头有 ${headerCount} 列），${missingCells}`;
    return `有一条记录的列数少于表头，${missingCells}`;
  }
  if (code === "InvalidQuotes") {
    return line
      ? `第 ${line} 行有不规范的引号（引号后面紧跟了其他字符），这一行可能被读错。`
      : "有一处不规范的引号（引号后面紧跟了其他字符），所在的记录可能被读错。";
  }
  if (code === "UndetectableDelimiter") {
    const [shown, name] = { ";": ["分号（；）", "分号"], "\t": ["制表符", "制表符"], "|": ["竖线（|）", "竖线"] }[separator] || [];
    return name ? `整份文件只读出 1 列，但表头里有${shown}。如果原文件是用${name}分隔的，请另存为逗号分隔的 CSV 后重新上传。` : "";
  }
  return "";
}

// aria-label of a 查看 button for a problem-row category.
function formatProblemViewLabel(category) {
  switch (category.kind) {
    case "missing": return `查看「${category.field}」的缺失行`;
    case "conversion": return `查看「${category.field}」的格式问题`;
    case "outlier": return `查看「${category.field}」的异常值`;
    case "duplicateRows": return "查看重复行";
    case "idDuplicate": return `查看「${category.field}」的重复行`;
    case "placeholder": return `查看「${category.field}」中「${category.key}」所在的行`;
    case "suspectedTotal": return "查看疑似合计 / 小计行";
    case "formulaError": return `查看「${category.field}」的公式错误`;
    case "dateStyle": return `查看「${category.field}」不是主写法的行`;
    case "lookAlike": return `查看「${category.field}」写法不同的行`;
    default: return "查看解析警告";
  }
}

// Option text of a problem-row category, with the quality page's count.
function formatProblemCategoryLabel(category) {
  const count = formatInteger(category.count);
  switch (category.kind) {
    case "missing": return `缺失 · ${category.field}（${count} 行）`;
    case "conversion": return `格式问题 · ${category.field}（${count} 个值）`;
    case "outlier": return `异常值 · ${category.field}（${count} 行）`;
    case "duplicateRows": return `整行重复（${count} 行）`;
    case "idDuplicate": return `ID 重复 · ${category.field}（${count} 个重复值）`;
    case "placeholder": return `疑似缺失值 · ${category.field} · 「${category.key}」（${count} 行）`;
    case "suspectedTotal": return `疑似合计 / 小计行（${count} 行）`;
    case "formulaError": return `公式错误 · ${category.field}（${count} 个值）`;
    case "dateStyle": return `日期写法 · ${category.field}（${count} 个值）`;
    case "lookAlike": return `写法不同 · ${category.field}（${count} 行）`;
    default: return `解析警告（${count} 条）`;
  }
}

// Viewer summary: the listed count in its unit, then "第 X / Y 页". Duplicate rows
// say how the quality count (without each group's kept first row) relates to the list.
function formatProblemRowsSummary(category, list, page, pageCount) {
  const pages = `第 ${formatInteger(page)} / ${formatInteger(pageCount)} 页`;
  const listed = formatInteger(list.items.length);
  if (category.kind === "duplicateRows") return `重复 ${formatInteger(list.count)} 行（${formatInteger(list.groupCount)} 组，共 ${listed} 行，每组第一行保留），${pages}`;
  if (category.kind === "conversion" || category.kind === "formulaError" || category.kind === "dateStyle") return `共 ${listed} 个值，${pages}`;
  if (category.kind === "parseWarnings") return `共 ${listed} 条警告，${pages}`;
  return `共 ${listed} 行，${pages}`;
}

// "6、19", plus " 等 15 个" when totalCount is more than the numbers shown. Any row
// without a number shows "—" instead of a partial list.
function formatRowNumberList(rowNumbers, totalCount = rowNumbers.length) {
  if (!rowNumbers.length || rowNumbers.some((rowNumber) => rowNumber === null)) return "—";
  const list = rowNumbers.join("、");
  return totalCount > rowNumbers.length ? `${list} 等 ${totalCount} 个` : list;
}

// The note on one duplicate ID value: how many of its rows repeat a whole row.
function formatIdenticalRows(rowCount, groupCount) {
  if (!rowCount) return "";
  return `其中 ${rowCount} 行内容完全相同${groupCount > 1 ? `（分 ${groupCount} 组）` : ""}`;
}

function formatIdDuplicateOverflow(totalCount, listedCount) {
  return totalCount > listedCount ? `共 ${totalCount} 个重复值，只列出前 ${listedCount} 个` : "";
}

// Names the ID fields without duplicates; afterOthers when other ID fields were reported first.
function formatCleanIdFieldsNote(fields, afterOthers = false) {
  return `${afterOthers ? "其余 " : ""}ID 字段没有重复值（已检查：${fields.join("、")}）`;
}

// Shown instead of the value list for a mostly repeated ID field (plain text).
function formatMostlyRepeatedIdHint({ field, distinctCount, nonMissingCount, uniqueRatio }) {
  return `「${field}」的值大多重复出现（${formatInteger(distinctCount)} 个不同值，共 ${formatInteger(nonMissingCount)} 行，唯一率 ${formatPercent(uniqueRatio)}），像是同一对象的多条记录，不逐个列出。如果这一列应该每行唯一，请检查数据；不需要检查的话，可以在字段配置里改类型。`;
}

// Shown with every maybe-valid placeholder key; the dash note only for 「-」.
const PLACEHOLDER_MAYBE_NOTE = "「无」「暂无」「-」「/」「NA」也可能是有效值（例如“有/无”选项、北美地区），请确认这一列里它确实表示没有数据。";
const PLACEHOLDER_DASH_NOTE = "在会计格式里，「-」通常表示 0。";

// 「N/A」, plus "（原文：Ｎ／Ａ、N/A）" when a raw text differs from the key.
function formatPlaceholderValue({ key, rawValues }) {
  return rawValues.some((value) => value !== key) ? `「${key}」（原文：${rawValues.join("、")}）` : `「${key}」`;
}

// "（第 15 行）", "（第 2、…、11 行等）" when more rows hold the key, "（行号：—）" without row numbers.
function formatPlaceholderRows(rowNumbers, count) {
  const list = formatRowNumberList(rowNumbers);
  if (list === "—") return "（行号：—）";
  return `（第 ${list} 行${count > rowNumbers.length ? "等" : ""}）`;
}

// "缺失值 0 → 1，类型仍为数值", or the new type when the confirmation changes it.
function formatPlaceholderChange(field, missingBefore, typeKeyBefore, preview) {
  const typeText = preview.typeKey === typeKeyBefore
    ? `类型仍为${FIELD_TYPE_LABELS[typeKeyBefore]}`
    : `「${field}」会识别为${FIELD_TYPE_LABELS[preview.typeKey]}（现在是${FIELD_TYPE_LABELS[typeKeyBefore]}）`;
  return `缺失值 ${formatInteger(missingBefore)} → ${formatInteger(preview.missingCount)}，${typeText}`;
}

function formatPendingPlaceholder(entry, change) {
  return `「${entry.field}」：${formatPlaceholderValue(entry)}${formatInteger(entry.count)} 个${formatPlaceholderRows(entry.rowNumbers, entry.count)}。确认后${change}。`;
}

function formatConfirmedPlaceholder(entry) {
  return `「${entry.field}」${entry.fieldIgnored ? "（字段已忽略）" : ""}：${formatPlaceholderValue(entry)}已按缺失处理（${formatInteger(entry.count)} 个）。`;
}

function formatPlaceholderNotes(keys) {
  return keys.some((key) => placeholderTier(key) === "maybe")
    ? `${PLACEHOLDER_MAYBE_NOTE}${keys.includes("-") ? PLACEHOLDER_DASH_NOTE : ""}`
    : "";
}

// The notes of "日期写法混用", "常量列" and "写法细节不同的值" in reports; index.html has the same
// sentences.
const DATE_STYLE_NOTE = "这些写法本工具都能读成同一个日期，但在 Excel 里文本日期不能按日期排序、筛选，导入别的系统时可能读错。";
const CONSTANT_FIELD_NOTE = "这些列的非空值全都一样。不一定是错误：可能是导出时带上的固定列，也可能是漏填或填错。";
const LOOK_ALIKE_NOTE = "首尾空格：本工具已按去掉空格处理，但导出给别的系统、用 VLOOKUP 时会对不上。同一个值有多种写法：本工具把它们当成不同的值，分类统计、筛选和 ID 重复都会分开算。只有不可见字符：格子看起来是空的，但本工具不按缺失处理，分类统计和筛选里会多出一个看不见的值。";
// Names of the date styles, as a user would recognise them in Excel.
const DATE_STYLE_NAMES = Object.freeze({
  excelCell: "Excel 日期单元格",
  ymdDash: "文本 年-月-日",
  ymdSlash: "文本 年/月/日",
  ymdDot: "文本 年.月.日",
  chinese: "文本 中文年月日",
  dmy: "文本 日/月/年",
  mdy: "文本 月/日/年",
  dayMonthSame: "文本 日和月相同（看不出先后）",
  english: "文本 英文月份",
  yearOnly: "只写年份"
});

const VISIBLE_TEXT_NAMES = Object.freeze({
  "\t": "[制表符]", "\n": "[换行]", "\r": "[换行]", "\u00A0": "[不换行空格]", "\u3000": "[全角空格]",
  "\u00AD": "[软连字符]", "\u200E": "[方向符]", "\u200F": "[方向符]",
  "\u200B": "[零宽字符]", "\u200C": "[零宽字符]", "\u200D": "[零宽字符]", "\u2060": "[零宽字符]", "\uFEFF": "[零宽字符]"
});
const VISIBLE_TEXT_PATTERN = new RegExp(`[\\s${INVISIBLE_CHARACTER_CLASS}]`, "g");

// A text with its whitespace and invisible characters shown: "␣" for a half-width space at
// an end or in a run (a single space inside stays), a name in brackets for the others.
// Plain text; callers escape it for their output format.
function formatVisibleText(text) {
  return String(text)
    .replace(/^ +| +$| {2,}/g, (spaces) => "␣".repeat(spaces.length))
    .replace(VISIBLE_TEXT_PATTERN, (character) => (character === " " ? " " : VISIBLE_TEXT_NAMES[character] || "[空白]"));
}

function formatDateStylePart({ key, count, sample }) {
  return `${DATE_STYLE_NAMES[key]} ${formatInteger(count)} 个（如 ${formatVisibleText(sample)}）`;
}

// 「入职日期」：主写法是 Excel 日期单元格 24 个（如 2016-04-01）；另有 6 行是别的写法：文本 年/月/日 2 个（如 2021/3/5）、…。
function formatDateStyleLine({ field, styles, otherCount }) {
  const [main, ...others] = styles;
  return `「${field}」：主写法是 ${formatDateStylePart(main)}；另有 ${formatInteger(otherCount)} 行是别的写法：${others.map(formatDateStylePart).join("、")}。`;
}
// 「备注」：只有一种取值「正常」（6 行），另有 3 行缺失。
function formatConstantField({ field, value, count, missingCount }) {
  return `「${field}」：只有一种取值「${formatVisibleText(value)}」（${formatInteger(count)} 行）${missingCount ? `，另有 ${formatInteger(missingCount)} 行缺失` : ""}。`;
}

const LOOK_ALIKE_REASON_NAMES = Object.freeze({ compat: "全角半角或兼容写法", case: "大小写", invisible: "不可见字符", innerSpace: "中间空格多少" });

function formatLookAlikeReasons(reasons) {
  return reasons.map((reason) => LOOK_ALIKE_REASON_NAMES[reason]).join("、");
}

// 「HD」2 个 / 「\uFF28\uFF24」1 个 / 「hd」1 个（全角半角或兼容写法、大小写）
function formatLookAlikeGroup({ variants, variantCount = variants.length, reasons }) {
  const texts = variants.map(({ text, count }) => `「${formatVisibleText(text)}」${formatInteger(count)} 个`).join(" / ");
  return `${texts}${variantCount > variants.length ? ` 等 ${formatInteger(variantCount)} 种` : ""}（${formatLookAlikeReasons(reasons)}）`;
}

// 「区域」共 18 行：首尾空格 2 格（如「华东␣」）；只有不可见字符 1 格；3 组写法不同：…；…。
function formatLookAlikeLine({ field, edgeSpaceCount, edgeSample, invisibleOnlyCount = 0, groups, groupCount, rowCount }) {
  const parts = [];
  if (edgeSpaceCount) parts.push(`首尾空格 ${formatInteger(edgeSpaceCount)} 格（如「${formatVisibleText(edgeSample)}」）`);
  if (invisibleOnlyCount) parts.push(`只有不可见字符 ${formatInteger(invisibleOnlyCount)} 格`);
  if (groupCount) {
    const overflow = groupCount > groups.length ? `，只列出前 ${groups.length} 组` : "";
    parts.push(`${formatInteger(groupCount)} 组写法不同${overflow}：${groups.map(formatLookAlikeGroup).join("；")}`);
  }
  return `「${field}」共 ${formatInteger(rowCount)} 行：${parts.join("；")}。`;
}

function formatDateStyleRowNote(key, mainKey) {
  return `${DATE_STYLE_NAMES[key]}（主写法是 ${DATE_STYLE_NAMES[mainKey]}）`;
}

// The 说明 of a problem row: the cell's own text with its whitespace shown, since the table
// shows the value trimmed.
function formatLookAlikeRowNote(value, group, hasEdge, invisibleOnly) {
  const parts = [];
  if (group) parts.push(`第 ${group.position + 1} 组（${formatLookAlikeReasons(group.reasons)}）`);
  if (invisibleOnly) parts.push("只有不可见字符");
  if (hasEdge) parts.push("首尾空格");
  return `${parts.join("，")}：「${formatVisibleText(value)}」`;
}

// The 处理 of a standardized cell (v2.10.0), for the check table's samples and the change list.
const STANDARDIZE_ACTION_LABELS = Object.freeze({
  date: "日期改写为 YYYY-MM-DD",
  number: "数字改写为纯数字",
  placeholder: "已确认的疑似缺失值，清空",
  trim: "去首尾空格",
  keptTime: "带时间，保持原值",
  keptPercent: "百分比，不改写",
  keptScientific: "科学计数法，不改写",
  keptCurrencyMix: "货币符号不一致，保持原值",
  keptPendingPlaceholder: "疑似缺失值，未确认",
  keptDateUnreadable: "读不成日期，保持原值",
  keptDateOrder: "日月顺序无法确定，保持原值",
  keptNumberUnreadable: "读不成数字，保持原值",
  keptUnsafe: "无法安全改写，保持原值"
});

function formatStandardizeAction(action) {
  return STANDARDIZE_ACTION_LABELS[action] || "";
}
function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
