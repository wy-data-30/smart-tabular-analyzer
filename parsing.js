// 单元格值的缺失判断与解析：把原始文本安全地转换为数值和日历日期（按书写的日历日，不受时区影响）。

const DATE_TIME_SUFFIX_PATTERN = "(?:[ T](?:[01]?\\d|2[0-3]):[0-5]\\d(?::[0-5]\\d(?:\\.\\d{1,9})?)?(?:Z|[+-](?:[01]\\d|2[0-3])(?::?[0-5]\\d)?)?)?";

function isYearOnlyDateValue(value, strategy = {}) {
  if (!strategy.allowYearOnly || isDateObject(value) || isMissing(value)) return false;
  const raw = String(value).trim();
  const yearText = raw.match(/^(\d{4})\s*年$/);
  return isPlausibleYear(yearText ? yearText[1] : raw);
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
