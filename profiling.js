// 字段识别：根据取值判断字段类型（ID / 日期 / 数值 / 分类），生成字段 profile，并提供按字段读取 profile 的工具。

const FIELD_TYPE_LABELS = {
  numeric: "数值",
  category: "分类",
  date: "日期",
  id: "ID",
  ignore: "忽略"
};

function buildColumnProfile(field, rows, numberFormatHint = "") {
  const values = rows.map((row) => row[field]);
  const nonMissingValues = values.filter((value) => !isMissing(value));
  const total = rows.length;
  const nonMissingCount = nonMissingValues.length;
  const uniqueValues = new Set(nonMissingValues.map((value) => normalizeValue(value)));
  const uniqueCount = uniqueValues.size;
  const uniqueRatio = nonMissingCount ? uniqueCount / nonMissingCount : 0;
  const conversionFailureSamplesByType = { numeric: [], date: [] };
  const numericValues = [];
  nonMissingValues.forEach((value) => {
    const number = toNumber(value);
    if (number !== null) numericValues.push(number);
    else addConversionFailureSample(conversionFailureSamplesByType.numeric, value);
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
  const idName = isIdFieldName(field);

  let typeKey = "category";

  if (idName || (uniqueRatio >= 0.98 && nonMissingCount >= 12 && numericRatio < 0.7 && dateRatio < 0.7)) {
    typeKey = "id";
  } else if (dateRatio >= 0.8) {
    typeKey = "date";
  } else if (numericRatio >= 0.8) {
    typeKey = "numeric";
  }

  const type = FIELD_TYPE_LABELS[typeKey];
  const conversionFailuresByType = {
    numeric: nonMissingCount - numericValues.length,
    date: nonMissingCount - dateValues.length
  };
  const detectedNumericFormat = detectNumericFormat(nonMissingValues);
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
    numericRatio,
    dateRatio,
    dateStrategy,
    numericFormat,
    sampleValues,
    conversionFailuresByType,
    conversionFailureSamplesByType,
    conversionFailureCount: conversionFailuresByType[typeKey] || 0,
    conversionFailureSamples: [...(conversionFailureSamplesByType[typeKey] || [])]
  };
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

function detectNumericFormat(values) {
  const convertibleValues = values.filter((value) => toNumber(value) !== null);
  if (!convertibleValues.length) return "number";
  return convertibleValues.every(hasPercentToken) ? "percent" : "number";
}

function getNumericFormat(field) {
  return getFieldProfile(field)?.numericFormat || "number";
}
