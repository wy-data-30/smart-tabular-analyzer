// 字段识别：根据取值判断字段类型（ID / 日期 / 数值 / 分类），生成字段 profile，并提供按字段读取 profile 的工具。

const FIELD_TYPE_LABELS = {
  numeric: "数值",
  category: "分类",
  date: "日期",
  id: "ID",
  ignore: "忽略"
};

// confirmedKeys: placeholder keys of this field that count as missing.
function buildColumnProfile(field, rows, numberFormatHint = "", confirmedKeys = []) {
  const values = rows.map((row) => row[field]);
  const nonMissingValues = values.filter((value) => !isMissingOrConfirmed(value, confirmedKeys));
  const total = rows.length;
  const nonMissingCount = nonMissingValues.length;
  const uniqueValues = new Set(nonMissingValues.map((value) => normalizeValue(value)));
  const uniqueCount = uniqueValues.size;
  const uniqueRatio = nonMissingCount ? uniqueCount / nonMissingCount : 0;
  const conversionFailureSamplesByType = { numeric: [], date: [] };
  const numberStrategy = buildNumberStrategy(nonMissingValues);
  const parseNumber = getNumberParser(numberStrategy);
  const numberReadings = { unitCount: 0, unitSamples: [], magnitudeCount: 0, fullwidthCount: 0 };
  // Figures for the numeric-looking hint, without any placeholder spelling. They
  // never convert to numbers, so only failed values need placeholderKey().
  const numericLookingSamples = [];
  let placeholderFailureCount = 0;
  const numericValues = [];
  nonMissingValues.forEach((value) => {
    const number = parseNumber(value);
    if (number === null) {
      addConversionFailureSample(conversionFailureSamplesByType.numeric, value);
      if (placeholderKey(value)) placeholderFailureCount += 1;
      else addConversionFailureSample(numericLookingSamples, value);
      return;
    }
    numericValues.push(number);
    if (parseNumber !== toPlainNumber && toPlainNumber(value) === null) addNumberReading(numberReadings, value, numberStrategy);
  });
  const dateStrategy = buildDateStrategy(field, nonMissingValues);
  const dateValues = [];
  nonMissingValues.forEach((value) => {
    const date = toDate(value, dateStrategy);
    if (date) dateValues.push(date);
    else addConversionFailureSample(conversionFailureSamplesByType.date, value);
  });
  const numericRatio = nonMissingCount ? numericValues.length / nonMissingCount : 0;
  const dateRatio = nonMissingCount ? dateValues.length / nonMissingCount : 0;
  const typeKey = inferFieldTypeKey(field, {
    nonMissingCount,
    uniqueCount,
    numericCount: numericValues.length,
    dateCount: dateValues.length
  });
  const type = FIELD_TYPE_LABELS[typeKey];
  const conversionFailuresByType = {
    numeric: nonMissingCount - numericValues.length,
    date: nonMissingCount - dateValues.length
  };
  const detectedNumericFormat = detectNumericFormat(nonMissingValues, numberStrategy);
  const numericFormat = detectedNumericFormat === "number" && numberFormatHint === "percent"
    ? "percent"
    : detectedNumericFormat;
  const sampleValues = [];
  const sampleKeys = new Set();
  nonMissingValues.forEach((value) => {
    const displayed = displayValue(value);
    const key = normalizeValue(value);
    if (sampleValues.length >= 3 || sampleKeys.has(key)) return;
    sampleKeys.add(key);
    sampleValues.push(displayed);
  });

  return {
    field,
    type,
    typeKey,
    inferredType: type,
    inferredTypeKey: typeKey,
    total,
    nonMissingCount,
    missingCount: total - nonMissingCount,
    missingRate: total ? (total - nonMissingCount) / total : 0,
    uniqueCount,
    uniqueRatio,
    numericCount: numericValues.length,
    dateCount: dateValues.length,
    numericRatio,
    dateRatio,
    dateStrategy,
    numberStrategy,
    numberReadings,
    numericLooking: {
      valueCount: nonMissingCount - placeholderFailureCount,
      numericCount: numericValues.length,
      samples: numericLookingSamples
    },
    numericFormat,
    sampleValues,
    conversionFailuresByType,
    conversionFailureSamplesByType,
    conversionFailureCount: conversionFailuresByType[typeKey] || 0,
    conversionFailureSamples: [...(conversionFailureSamplesByType[typeKey] || [])]
  };
}

// The automatic type from a field's counts, shared by buildColumnProfile() and the
// placeholder preview: ID by name, or by at least 12 values that are 98% unique and
// below 70% numbers and dates; then date or numeric at 80%; otherwise category.
function inferFieldTypeKey(field, { nonMissingCount, uniqueCount, numericCount, dateCount }) {
  const uniqueRatio = nonMissingCount ? uniqueCount / nonMissingCount : 0;
  const numericRatio = nonMissingCount ? numericCount / nonMissingCount : 0;
  const dateRatio = nonMissingCount ? dateCount / nonMissingCount : 0;
  if (isIdFieldName(field) || (uniqueRatio >= 0.98 && nonMissingCount >= 12 && numericRatio < 0.7 && dateRatio < 0.7)) return "id";
  if (dateRatio >= 0.8) return "date";
  if (numericRatio >= 0.8) return "numeric";
  return "category";
}

// Counts a value only the number strategy converts: read with the unit (up to three
// distinct samples), scaled by 万 / 亿, or written in full width.
function addNumberReading(readings, value, strategy) {
  const raw = String(value).trim();
  const text = strategy.fullwidth ? toHalfwidthNumberText(raw) : raw;
  if (text !== raw) readings.fullwidthCount += 1;
  const match = matchNumberWithUnit(text);
  if (!match) return;
  if (match.unit) {
    readings.unitCount += 1;
    if (readings.unitSamples.length < 3 && !readings.unitSamples.includes(raw)) readings.unitSamples.push(raw);
  }
  if (match.magnitude) readings.magnitudeCount += 1;
}

// Keeps up to three distinct (trimmed) raw values that failed conversion.
function addConversionFailureSample(samples, value) {
  if (samples.length >= 3) return;
  const sample = displayValue(value);
  if (!samples.includes(sample)) samples.push(sample);
}

function getProfilesByType(typeKey) {
  return state.profiles.filter((profile) => profile.typeKey === typeKey);
}

function getActiveProfiles() {
  return state.profiles.filter((profile) => profile.typeKey !== "ignore");
}

function getFieldProfile(field) {
  return state.profiles.find((profile) => profile.field === field);
}

function parseFieldDate(field, value) {
  return toDate(value, getFieldProfile(field)?.dateStrategy || {});
}

// The number parser of a field, for callers to take once outside their row loops:
// toPlainNumber itself when the field's strategy has no switch.
function getFieldNumberParser(field) {
  return getNumberParser(getFieldProfile(field)?.numberStrategy);
}

function parseFieldNumber(field, value) {
  return getFieldNumberParser(field)(value);
}

function getDateGranularity(field) {
  return getFieldProfile(field)?.dateStrategy?.granularity || "day";
}

function isIdFieldName(field) {
  const name = String(field || "").trim();
  return /^(?:序号|行号|行序号|index|idx|no\.|row no\.?|#)$/i.test(name)
    || /[\u4e00-\u9fff]\s*id$/i.test(name)
    || /(^|[_\-\s])(id|uuid|guid|code|key|number|no)([_\-\s]|$)/i.test(name)
    || /[a-z0-9](?:Id|ID|Uuid|UUID|Guid|GUID|Code|CODE|Key|KEY|Number|NUMBER|No|NO)$/.test(name)
    || /^(?:user|customer|order|product|account|session|event|device|student|employee|member|record)(?:id|uuid|guid|code|key|number|no)$/i.test(name)
    || /编号|编码|代码|订单号|账号|工号|主键|用户ID|事件ID|会话ID/i.test(name)
    || /(?:学号|学籍号|考生号|准考证号|身份证号|身份证号码|证件号|手机号|手机号码|流水号|序列号|设备号|会员号|合同号|票据号|批次号|档案号)(?:$|[_\-\s（(])/i.test(name);
}

function isYearFieldName(field) {
  const name = String(field || "").trim();
  return /(年份|年度|学年|财年|会计年度|统计年度|统计年|报告年|出生年|入学年|毕业年)$/i.test(name)
    || /(^|[_\-\s])(year|fiscal_year|calendar_year|school_year|yr|fy)([_\-\s]|$)/i.test(name);
}

function buildDateStrategy(field, values) {
  let dmyEvidence = 0;
  let mdyEvidence = 0;
  let yearOnlyCount = 0;

  values.forEach((value) => {
    if (isDateObject(value)) return;
    const raw = String(value).trim();
    const yearText = raw.match(/^(\d{4})\s*年$/);
    if (isPlausibleYear(yearText ? yearText[1] : raw)) yearOnlyCount += 1;
    const match = matchDayOrMonthFirstDate(raw);
    if (!match) return;
    const first = Number(match[1]);
    const second = Number(match[3]);
    const year = Number(match[4]);
    if (first > 12 && second <= 12 && buildUtcDate(year, second, first)) dmyEvidence += 1;
    if (second > 12 && first <= 12 && buildUtcDate(year, first, second)) mdyEvidence += 1;
  });

  let order = "none";
  if (dmyEvidence && mdyEvidence) order = "conflict";
  else if (dmyEvidence) order = "dmy";
  else if (mdyEvidence) order = "mdy";

  const allowYearOnly = isYearFieldName(field)
    && values.length > 0
    && yearOnlyCount / values.length >= 0.8;

  return {
    order,
    allowYearOnly,
    granularity: allowYearOnly ? "year" : "day",
    dmyEvidence,
    mdyEvidence
  };
}

// Percent when every convertible value carries %; a full-width strategy reads ％ too.
function detectNumericFormat(values, strategy) {
  const parseNumber = getNumberParser(strategy);
  const convertibleValues = values.filter((value) => parseNumber(value) !== null);
  if (!convertibleValues.length) return "number";
  const percentToken = strategy?.fullwidth ? (value) => hasPercentToken(toHalfwidthNumberText(String(value))) : hasPercentToken;
  return convertibleValues.every(percentToken) ? "percent" : "number";
}

// The unit a field's numbers are read with, or "".
function getFieldNumberUnit(field) {
  return getFieldProfile(field)?.numberStrategy?.unit || "";
}

function getNumericFormat(field) {
  return getFieldProfile(field)?.numericFormat || "number";
}
