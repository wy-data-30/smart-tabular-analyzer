// 单元格值的缺失判断与解析：疑似缺失值的写法清单，把原始文本安全地转换为数值（可按字段的数字策略读取单位、万/亿和全角数字）和日历日期（按书写的日历日，不受时区影响）。

const DATE_TIME_SUFFIX_PATTERN = "(?:[ T](?:[01]?\\d|2[0-3]):[0-5]\\d(?::[0-5]\\d(?:\\.\\d{1,9})?)?(?:Z|[+-](?:[01]\\d|2[0-3])(?::?[0-5]\\d)?)?)?";

function isYearOnlyDateValue(value, strategy = {}) {
  if (!strategy.allowYearOnly || isDateObject(value) || isMissing(value)) return false;
  const raw = String(value).trim();
  const yearText = raw.match(/^(\d{4})\s*年$/);
  return isPlausibleYear(yearText ? yearText[1] : raw);
}

function isMissing(value) {
  // Fast path: trim() removes only U+0009–U+000D, U+0020, U+00A0, U+1680,
  // U+2000–U+200A, U+2028, U+2029, U+202F, U+205F, U+3000 and U+FEFF. A string
  // whose first character is none of these cannot trim to "", so it is not blank.
  if (typeof value === "string" && value.length) {
    const code = value.charCodeAt(0);
    if (code > 32 && code !== 160 && (code < 0x1680 || (code > 0x3000 && code !== 0xfeff))) return false;
  }
  return value === null || value === undefined || String(value).trim() === "";
}

function normalizeValue(value) {
  return isMissing(value) ? "" : String(value).trim();
}

function displayValue(value) {
  return isMissing(value) ? "" : String(value).trim();
}

// Cell texts often used for "no data". They stay ordinary values until the user
// confirms them per field and key. "common" keys can be confirmed together;
// "maybe" keys are also real answers in some columns (有/无, 北美 NA).
const PLACEHOLDER_VALUES = Object.freeze({
  common: Object.freeze(["N/A", "n/a", "#N/A", "NULL", "null", "Null", "None", "none", "NaN", "nan", "NAN", "—", "——", "--"]),
  maybe: Object.freeze(["无", "暂无", "-", "/", "NA"])
});

const PLACEHOLDER_TIER_BY_KEY = new Map([
  ...PLACEHOLDER_VALUES.common.map((key) => [key, "common"]),
  ...PLACEHOLDER_VALUES.maybe.map((key) => [key, "maybe"])
]);

// NFKC never maps a code point to nothing and no key character is a composition
// result, so a cell longer than two UTF-16 units per key character cannot match.
const PLACEHOLDER_MAX_CELL_LENGTH = 2 * Math.max(...[...PLACEHOLDER_TIER_BY_KEY.keys()].map((key) => key.length));

// The list key of a text cell (trimmed, NFKC, trimmed again), or "" when the
// whole cell is not on the list.
function placeholderKey(value) {
  if (typeof value !== "string") return "";
  const trimmed = value.trim();
  if (!trimmed || trimmed.length > PLACEHOLDER_MAX_CELL_LENGTH) return "";
  const key = trimmed.normalize("NFKC").trim();
  return PLACEHOLDER_TIER_BY_KEY.has(key) ? key : "";
}

function placeholderTier(key) {
  return PLACEHOLDER_TIER_BY_KEY.get(key) || "";
}

// Missing for analysis and quality: isMissing(), or a placeholder key confirmed
// for the value's field (confirmedKeys, empty when none).
function isMissingOrConfirmed(value, confirmedKeys) {
  return isMissing(value) || (confirmedKeys.length > 0 && confirmedKeys.includes(placeholderKey(value)));
}

// normalizeValue() / displayValue() for analysis: "" also for a confirmed placeholder.
function analysisValue(value, confirmedKeys) {
  return isMissingOrConfirmed(value, confirmedKeys) ? "" : String(value).trim();
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

function hasPercentToken(value) {
  const raw = String(value).trim();
  const unsigned = /^\(.*\)$/.test(raw) ? raw.slice(1, -1).trim() : raw;
  return /%$/.test(unsigned);
}

// Full-width characters a number strategy reads as half-width. A fixed table, not
// NFKC, so "①" and "²" never become digits.
const FULLWIDTH_NUMBER_CHARS = Object.freeze({
  "０": "0", "１": "1", "２": "2", "３": "3", "４": "4", "５": "5", "６": "6", "７": "7", "８": "8", "９": "9",
  "，": ",", "．": ".", "－": "-", "＋": "+", "％": "%"
});

// Units a number strategy may read, longer spellings first. Latin letters match in
// any case and compare in lower case. No "g": 4G / 5G and 8G / 16G are not grams.
const NUMBER_UNITS = Object.freeze(["千克", "公斤", "小时", "分钟", "元", "克", "斤", "吨", "件", "个", "箱", "台", "人", "次", "秒", "天", "米", "岁", "kg", "cm", "mm", "km"]);

// Sign, a plain or comma-grouped number, 万 / 亿 right after it, then optional
// space and a unit.
const NUMBER_WITH_UNIT_PATTERN = new RegExp(`^([+-]?(?:\\d{1,3}(?:,\\d{3})+(?:\\.\\d+)?|\\d+(?:\\.\\d+)?|\\.\\d+))(万|亿)?(?:\\s*(${NUMBER_UNITS.join("|")}))?$`, "i");

function toHalfwidthNumberText(text) {
  return text.replace(/[０-９，．－＋％]/g, (char) => FULLWIDTH_NUMBER_CHARS[char]);
}

// { numberText, magnitude, unit } for a trimmed, half-width text with 万 / 亿, a
// unit or both; otherwise null.
function matchNumberWithUnit(text) {
  const match = text.match(NUMBER_WITH_UNIT_PATTERN);
  if (!match || (!match[2] && !match[3])) return null;
  return { numberText: match[1], magnitude: match[2] || "", unit: (match[3] || "").toLowerCase() };
}

// A field's number strategy from its non-missing values. Only number-like values
// decide, never their share of the column: one unit among them is adopted, two or
// more spellings adopt none (unitConflict, with counts in units).
function buildNumberStrategy(values) {
  const units = new Map();
  let allowMagnitude = false;
  let fullwidth = false;
  values.forEach((value) => {
    if (typeof value !== "string") return;
    const raw = value.trim();
    const text = toHalfwidthNumberText(raw);
    const match = matchNumberWithUnit(text);
    if (text !== raw && (match || toPlainNumber(text) !== null)) fullwidth = true;
    if (!match) return;
    if (match.unit) units.set(match.unit, (units.get(match.unit) || 0) + 1);
    if (match.magnitude) allowMagnitude = true;
  });
  return {
    unit: units.size === 1 ? [...units.keys()][0] : "",
    unitConflict: units.size > 1,
    units: Object.fromEntries(units),
    allowMagnitude,
    fullwidth
  };
}

function hasNumberStrategySwitch(strategy) {
  return Boolean(strategy && (strategy.unit || strategy.allowMagnitude || strategy.fullwidth));
}

// Without a strategy, or with one that has no switch, the plain parser (v2.4.0).
function toNumber(value, strategy) {
  return hasNumberStrategySwitch(strategy) ? toNumberWithStrategy(value, strategy) : toPlainNumber(value);
}

// A one-argument parser for a strategy: toPlainNumber itself when it has no switch.
function getNumberParser(strategy) {
  return hasNumberStrategySwitch(strategy) ? (value) => toNumberWithStrategy(value, strategy) : toPlainNumber;
}

function toNumberWithStrategy(value, strategy) {
  if (isMissing(value)) return null;
  const raw = String(value).trim();
  const text = strategy.fullwidth ? toHalfwidthNumberText(raw) : raw;
  const plain = toPlainNumber(text);
  if (plain !== null) return plain;
  const match = matchNumberWithUnit(text);
  if (!match || (match.unit && match.unit !== strategy.unit) || (match.magnitude && !strategy.allowMagnitude)) return null;
  if (!match.magnitude) return toPlainNumber(match.numberText);
  return shiftDecimalNumber(match.numberText, match.magnitude === "万" ? 4 : 8);
}

// 0.07万 is 700 exactly: the decimal point moves right as text, no float multiply.
function shiftDecimalNumber(numberText, places) {
  const [, sign, integerPart, fractionPart = ""] = numberText.replaceAll(",", "").match(/^([+-]?)(\d*)(?:\.(\d+))?$/);
  const fraction = fractionPart.padEnd(places, "0");
  const shifted = `${sign}${integerPart}${fraction.slice(0, places)}${fraction.length > places ? `.${fraction.slice(places)}` : ""}`;
  const parsed = Number(shifted);
  return Number.isFinite(parsed) ? parsed : null;
}

function toPlainNumber(value) {
  if (isMissing(value)) return null;
  const raw = String(value).trim();
  // Fast path: up to 15 plain digits always convert (the cap keeps Number() finite).
  if (/^\d{1,15}$/.test(raw)) return Number(raw);
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

  return parseEnglishCalendarDate(raw);
}

function parseEnglishCalendarDate(raw) {
  let text = raw.replace(/^(?:mon(?:day)?|tue(?:sday)?|wed(?:nesday)?|thu(?:rsday)?|fri(?:day)?|sat(?:urday)?|sun(?:day)?),?\s+/i, "");
  let timeBeforeYear = false;

  // Remove one validated clock, either after the date or between Month Day
  // and Year (asctime order). Its value never participates in date conversion.
  const clockStart = text.match(/(^|[\s,])(?=\d{1,2}:)/);
  const timeIndex = clockStart ? clockStart.index + clockStart[0].length : -1;
  if (timeIndex >= 0) {
    const prefix = text.slice(0, timeIndex).trim();
    const prefixParts = prefix.split(/[\s,/-]+/).filter(Boolean);
    if (prefixParts.length !== 2 && prefixParts.length !== 3) return null;
    timeBeforeYear = prefixParts.length === 2;
    const time = text.slice(timeIndex).match(/^(\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?(?:\s*(am|pm))?(?=$|[\s,]|[+-]|Z|GMT|UTC)/i);
    if (!time) return null;
    const hour = Number(time[1]);
    if (hour > 23 || Number(time[2]) > 59 || Number(time[3] || 0) > 59) return null;
    if (time[5] && (hour < 1 || hour > 12)) return null;
    text = `${prefix} ${text.slice(timeIndex + time[0].length).trim()}`.trim();
  }

  // Tokenize the three calendar components separately from the optional zone.
  // Month Day Year, Day Month Year and Year Month Day are unambiguous here.
  const parts = text.match(/^([a-z]+\.?|\d{1,4})[\s,/-]+([a-z]+\.?|\d{1,4})[\s,/-]+(\d{1,4})(?:[\s,]+(.*))?$/i);
  if (!parts) return null;
  const months = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"];
  const monthFirst = /^[a-z]/i.test(parts[1]);
  if (timeBeforeYear && !monthFirst) return null;
  const monthName = (monthFirst ? parts[1] : parts[2]).replace(/\.$/, "").toLowerCase();
  const month = months.findIndex((name) => monthName === name || monthName === name.slice(0, 3) || (name === "september" && monthName === "sept")) + 1;
  const yearFirst = !monthFirst && /^\d{4}$/.test(parts[1]);
  const year = yearFirst ? parts[1] : parts[3];
  const day = monthFirst ? parts[2] : yearFirst ? parts[3] : parts[1];
  if (!/^\d{4}$/.test(year) || !/^\d{1,2}$/.test(day)) return null;
  const calendarDate = buildUtcDate(Number(year), month, Number(day));
  if (!calendarDate) return null;

  // A zone is syntax only: applying its offset would change the written day.
  const timezone = (parts[4] || "").trim();
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
