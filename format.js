// 显示格式化：数值、整数、百分比和字段格式的展示文本，表头位置说明，ID 重复值的行号列表和说明文字，疑似缺失值的说明文字，以及 HTML 转义。

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

function numericAxisLabel(field, fallbackLabel) {
  return getNumericFormat(field) === "percent"
    ? `${fallbackLabel}（百分比）`
    : fallbackLabel;
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

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
