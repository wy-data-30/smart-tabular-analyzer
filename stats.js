// 统计计算：描述统计、IQR 异常值、分类频数、分组聚合、日期趋势和直方图分箱。

// Non-missing values in `rows` that the field's number parser cannot convert.
function countInvalidNumericValues(field, rows = getAnalysisRows()) {
  const confirmedKeys = getConfirmedPlaceholderKeys(field);
  const parseNumber = getFieldNumberParser(field);
  return rows.reduce((count, row) => (
    count + (!isMissingOrConfirmed(row[field], confirmedKeys) && parseNumber(row[field]) === null ? 1 : 0)
  ), 0);
}

// The returned array and its objects are shared: state.numericStats can be the very same
// array as state.qualityNumericStats (see refreshFilteredAnalysis()), so no caller may
// modify them in place.
function buildNumericStats(rows = getAnalysisRows()) {
  return getProfilesByType("numeric").flatMap((profile) => {
    const confirmedKeys = getConfirmedPlaceholderKeys(profile.field);
    const parseNumber = getNumberParser(profile.numberStrategy);
    const values = [];
    let invalidCount = 0;
    rows.forEach((row) => {
      const value = row[profile.field];
      if (isMissingOrConfirmed(value, confirmedKeys)) return;
      const number = parseNumber(value);
      if (number === null) invalidCount += 1;
      else values.push(number);
    });
    if (!values.length) return [];
    const sorted = [...values].sort((a, b) => a - b);
    const avg = mean(values);
    const min = sorted[0];
    const max = sorted[sorted.length - 1];
    const q1 = quantile(sorted, 0.25);
    const q3 = quantile(sorted, 0.75);
    const iqr = q3 - q1;
    const lowerBound = q1 - 1.5 * iqr;
    const upperBound = q3 + 1.5 * iqr;
    const outliers = values.filter((value) => value < lowerBound || value > upperBound);

    return [{
      field: profile.field,
      count: values.length,
      invalidCount,
      mean: avg,
      median: medianOfSorted(sorted),
      min,
      max,
      std: standardDeviation(values, avg),
      q1,
      q3,
      lowerBound,
      upperBound,
      outlierCount: outliers.length
    }];
  });
}

function buildCategoryStats(rows = getAnalysisRows()) {
  return getProfilesByType("category").map((profile) => {
    const counts = countValues(profile.field, rows);
    const confirmedKeys = getConfirmedPlaceholderKeys(profile.field);
    const nonMissingCount = rows.filter((row) => !isMissingOrConfirmed(row[profile.field], confirmedKeys)).length;
    const top = Array.from(counts.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10)
      .map(([name, count]) => ({
        name,
        count,
        ratio: nonMissingCount ? count / nonMissingCount : 0
      }));

    return {
      field: profile.field,
      total: nonMissingCount,
      distinctCount: counts.size,
      top
    };
  });
}

// Same bucketing as buildDateTrend(): month when the dated span exceeds 120
// days, otherwise day; year-only fields by year.
function buildDateChartTrend(dateField, metricField, aggregate) {
  const parseNumber = getFieldNumberParser(metricField);
  return buildTrendFromRows(dateField, (row) => parseNumber(row[metricField]), metricField ? aggregate : "count");
}

function buildV2GroupedStats(metricField, groupField) {
  return groupNumericFieldByCategory(metricField, groupField);
}

function buildMetricDateTrend(dateField, metricField, aggregateMethod = "sum") {
  const parseNumber = getFieldNumberParser(metricField);
  return buildTrendFromRows(dateField, (row) => parseNumber(row[metricField]), aggregateMethod);
}

function groupNumericFieldByCategory(metricField, groupField) {
  const groupKeys = getConfirmedPlaceholderKeys(groupField);
  const parseNumber = getFieldNumberParser(metricField);
  const groups = new Map();
  getAnalysisRows().forEach((row) => {
    const value = parseNumber(row[metricField]);
    if (value === null) return;
    const groupName = analysisValue(row[groupField], groupKeys) || "未填写";
    if (!groups.has(groupName)) groups.set(groupName, []);
    groups.get(groupName).push(value);
  });
  return summarizeGroupedValues(groups);
}

function groupMetricRowsByField(metricRows, groupField) {
  const groupKeys = getConfirmedPlaceholderKeys(groupField);
  const groups = new Map();
  metricRows.forEach((item) => {
    const groupName = analysisValue(item.row[groupField], groupKeys) || "未填写";
    if (!groups.has(groupName)) groups.set(groupName, []);
    groups.get(groupName).push(item.value);
  });
  return summarizeGroupedValues(groups);
}

function summarizeGroupedValues(groups) {
  return Array.from(groups.entries())
    .map(([name, values]) => {
      const { min, max } = numericExtents(values);
      return {
        name,
        count: values.length,
        sum: sumNumbers(values),
        avg: mean(values),
        median: median(values),
        min,
        max
      };
    })
    .sort((a, b) => compareNumbersDescending(a.sum, b.sum));
}

function buildTrendFromRows(dateField, valueGetter, aggregate) {
  const items = [];
  getAnalysisRows().forEach((row) => {
    const date = parseFieldDate(dateField, row[dateField]);
    if (!date) return;
    if (aggregate === "count") {
      items.push({ date, value: 1 });
      return;
    }
    const value = valueGetter(row);
    if (value !== null) items.push({ date, value });
  });

  if (!items.length) return { labels: [], values: [] };
  items.sort((a, b) => a.date - b.date);
  const spanDays = (items[items.length - 1].date - items[0].date) / 86400000;
  const groupByMonth = spanDays > 120;
  const groups = new Map();
  items.forEach(({ date, value }) => {
    const key = getDateTrendKey(dateField, date, groupByMonth);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(value);
  });

  const labels = Array.from(groups.keys()).sort();
  const values = labels.map((label) => {
    const bucket = groups.get(label);
    if (aggregate === "avg") return bucket.length ? mean(bucket) : 0;
    if (aggregate === "median") return bucket.length ? median(bucket) : 0;
    return sumNumbers(bucket);
  });
  return { labels, values };
}

function summarizeNumbers(values) {
  if (!values.length) {
    return { mean: 0, median: 0, min: 0, max: 0, std: 0 };
  }
  const { min, max } = numericExtents(values);
  return {
    mean: mean(values),
    median: median(values),
    min,
    max,
    std: standardDeviation(values)
  };
}

function frequencyRows(field, limit = 50) {
  const counts = countValues(field);
  const total = Array.from(counts.values()).reduce((sum, count) => sum + count, 0);
  return Array.from(counts.entries())
    .map(([name, count]) => ({ name, count, ratio: total ? count / total : 0 }))
    .sort((a, b) => b.count - a.count)
    .slice(0, limit);
}

function fieldHasNumericData(field, rows = getAnalysisRows()) {
  return getNumericValues(field, rows).length > 0;
}

function fieldHasDateData(field, rows = getAnalysisRows()) {
  return rows.some((row) => Boolean(parseFieldDate(field, row[field])));
}

function countNonMissingValues(field, rows = getAnalysisRows()) {
  const confirmedKeys = getConfirmedPlaceholderKeys(field);
  return rows.filter((row) => !isMissingOrConfirmed(row[field], confirmedKeys)).length;
}

// The trend conclusion always compares record counts; numeric fields and the
// chart's metric selects do not change it.
function buildTrendInsight(dateField) {
  const trend = buildDateTrend(dateField, "");
  if (!trend.labels.length) {
    return `日期字段「${dateField}」没有可安全转换的有效日期，因此未生成趋势结论。`;
  }
  if (trend.values.length < 2) {
    return `日期字段「${dateField}」可用于趋势分析，但有效时间点不足以判断变化方向。`;
  }
  const first = trend.values[0];
  const last = trend.values[trend.values.length - 1];
  const range = `基于日期字段「${dateField}」观察，${trend.label} 从 ${trend.labels[0]} 到 ${trend.labels[trend.labels.length - 1]}`;
  const note = buildPartialPeriodNote(dateField, trend);
  // Every bucket holds at least one record, so the first count is never 0.
  if (last === first) return `${range} 基本持平。${note}`;
  const direction = last > first ? "上升" : "下降";
  return `${range} 呈${direction}变化，变化幅度约 ${formatPercent(Math.abs((last - first) / first))}。${note}`;
}

// Month and year buckets may be incomplete when the latest written date ends
// before that month / year. Day buckets and year-only values ("2025") are not
// judged. Dates are UTC calendar days from parseFieldDate(), compared with the
// same bucket keys as getDateTrendKey(), so the runtime time zone never matters.
function buildPartialPeriodNote(dateField, trend) {
  const lastLabel = trend.labels[trend.labels.length - 1];
  const yearly = getDateGranularity(dateField) === "year";
  if (!yearly && !/^\d{4}-\d{2}$/.test(lastLabel)) return "";
  const strategy = getFieldProfile(dateField)?.dateStrategy || {};
  let lastDate = null;
  getAnalysisRows().forEach((row) => {
    const value = row[dateField];
    const date = parseFieldDate(dateField, value);
    if (!date || (yearly && isYearOnlyDateValue(value, strategy))) return;
    if (!lastDate || date > lastDate) lastDate = date;
  });
  if (!lastDate || getDateTrendKey(dateField, lastDate, !yearly) !== lastLabel) return "";
  const year = lastDate.getUTCFullYear();
  const month = lastDate.getUTCMonth();
  const periodEnd = yearly ? Date.UTC(year, 11, 31) : Date.UTC(year, month + 1, 0);
  if (lastDate.getTime() >= periodEnd) return "";
  const period = yearly ? `${year} 年` : lastLabel;
  return `注意：最后一期（${period}）的数据截至 ${month + 1} 月 ${lastDate.getUTCDate()} 日，可能不完整，比较结果仅供参考。`;
}

function buildDateTrend(dateField, numericField) {
  const parseNumber = getFieldNumberParser(numericField);
  const items = [];
  getAnalysisRows().forEach((row) => {
    const date = parseFieldDate(dateField, row[dateField]);
    if (!date) return;
    const value = numericField ? parseNumber(row[numericField]) : null;
    if (numericField && value === null) return;
    items.push({ date, value });
  });

  if (!items.length) {
    return { labels: [], values: [], label: "趋势", axisLabel: "值" };
  }

  items.sort((a, b) => a.date - b.date);
  const spanDays = (items[items.length - 1].date - items[0].date) / 86400000;
  const groupByMonth = spanDays > 120;
  const groups = new Map();

  items.forEach((item) => {
    const key = getDateTrendKey(dateField, item.date, groupByMonth);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(item.value);
  });

  const labels = Array.from(groups.keys()).sort();
  const hasMetric = Boolean(numericField);
  const values = labels.map((label) => {
    const bucket = groups.get(label);
    if (!hasMetric) return bucket.length;
    const valid = bucket.filter((value) => value !== null);
    return valid.length ? mean(valid) : 0;
  });

  return {
    labels,
    values,
    label: hasMetric
      ? `${numericField}${getDateGranularity(dateField) === "year" ? "年度" : " "}平均值趋势`
      : `${getDateGranularity(dateField) === "year" ? "年度" : ""}记录数趋势`,
    axisLabel: hasMetric ? "平均值" : "记录数"
  };
}

function aggregateByCategory(metricField, groupField, method) {
  const groupKeys = getConfirmedPlaceholderKeys(groupField);
  const parseNumber = getFieldNumberParser(metricField);
  const groups = new Map();
  getAnalysisRows().forEach((row) => {
    const group = analysisValue(row[groupField], groupKeys) || "未填写";
    const value = parseNumber(row[metricField]);
    if (value === null) return;
    if (!groups.has(group)) groups.set(group, []);
    groups.get(group).push(value);
  });

  return Array.from(groups.entries())
    .map(([name, values]) => {
      let value = 0;
      if (method === "avg") value = mean(values);
      else if (method === "median") value = median(values);
      else value = sumNumbers(values);
      return { name, value, count: values.length };
    })
    .sort((a, b) => compareNumbersDescending(a.value, b.value));
}

// Small integer ranges (scores, ratings) get one bar per value from min to max,
// zeros included; anything else returns null so callers fall back to buildHistogram.
function buildDiscreteValueCounts(values, field = "") {
  if (!values.length || getNumericFormat(field) === "percent") return null;
  if (!values.every(Number.isInteger)) return null;
  const { min, max } = numericExtents(values);
  if (!(max - min <= 11)) return null;
  const bins = Array.from({ length: max - min + 1 }, (_, index) => ({
    label: formatFieldNumber(field, min + index),
    count: 0
  }));
  values.forEach((value) => {
    bins[value - min].count += 1;
  });
  return bins;
}

function buildHistogram(values, binCount, field = "") {
  if (!values.length) return [];
  const { min, max } = numericExtents(values);
  if (min === max) return [{ label: formatFieldNumber(field, min), count: values.length }];
  const scale = Math.max(Math.abs(min), Math.abs(max));
  const normalizedMin = min / scale;
  const normalizedMax = max / scale;
  const normalizedWidth = (normalizedMax - normalizedMin) / binCount;
  if (!Number.isFinite(normalizedWidth) || normalizedWidth <= 0) {
    return [{ label: `${formatFieldNumber(field, min)} - ${formatFieldNumber(field, max)}`, count: values.length }];
  }
  const bins = Array.from({ length: binCount }, (_, index) => {
    const startRatio = index / binCount;
    const endRatio = (index + 1) / binCount;
    const start = min * (1 - startRatio) + max * startRatio;
    const end = index === binCount - 1
      ? max
      : min * (1 - endRatio) + max * endRatio;
    return {
      start,
      end,
      label: `${formatFieldNumber(field, start)} - ${formatFieldNumber(field, end)}`,
      count: 0
    };
  });
  values.forEach((value) => {
    const normalizedValue = value / scale;
    const rawIndex = Math.floor((normalizedValue - normalizedMin) / normalizedWidth);
    const index = Math.max(0, Math.min(rawIndex, binCount - 1));
    bins[index].count += 1;
  });
  return bins;
}

function countValues(field, rows = getAnalysisRows()) {
  const confirmedKeys = getConfirmedPlaceholderKeys(field);
  const counts = new Map();
  rows.forEach((row) => {
    if (isMissingOrConfirmed(row[field], confirmedKeys)) return;
    const value = displayValue(row[field]);
    counts.set(value, (counts.get(value) || 0) + 1);
  });
  return counts;
}

function getNumericValues(field, rows = getAnalysisRows()) {
  const parseNumber = getFieldNumberParser(field);
  return rows
    .map((row) => parseNumber(row[field]))
    .filter((value) => value !== null);
}

function getDateTrendKey(field, date, groupByMonth = false) {
  if (getDateGranularity(field) === "year") {
    return String(date.getUTCFullYear());
  }
  if (groupByMonth) {
    return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}`;
  }
  return date.toISOString().slice(0, 10);
}

function mean(values) {
  if (!values.length) return 0;
  const scale = values.reduce((max, value) => Math.max(max, Math.abs(value)), 0);
  if (scale === 0) return 0;
  const normalizedTotal = values.reduce((sum, value) => sum + value / scale, 0);
  return (normalizedTotal / values.length) * scale;
}

function sumNumbers(values) {
  if (!values.length) return 0;
  const scale = values.reduce((max, value) => Math.max(max, Math.abs(value)), 0);
  if (scale === 0) return 0;
  const normalizedTotal = values.reduce((sum, value) => sum + value / scale, 0);
  const result = normalizedTotal * scale;
  return Number.isFinite(result) ? result : null;
}

function compareNumbersDescending(left, right) {
  const leftIsFinite = Number.isFinite(left);
  const rightIsFinite = Number.isFinite(right);
  if (leftIsFinite && rightIsFinite) return right - left;
  if (leftIsFinite) return -1;
  if (rightIsFinite) return 1;
  return 0;
}

function median(values) {
  if (!values.length) return 0;
  return medianOfSorted([...values].sort((a, b) => a - b));
}

// Median of values already in ascending order.
function medianOfSorted(sorted) {
  if (!sorted.length) return 0;
  const middle = Math.floor(sorted.length / 2);
  if (sorted.length % 2) return sorted[middle];
  const lower = sorted[middle - 1];
  const upper = sorted[middle];
  const sum = lower + upper;
  return Number.isFinite(sum) ? sum / 2 : lower / 2 + upper / 2;
}

function standardDeviation(values, avg = mean(values)) {
  if (!values.length) return 0;
  const scale = values.reduce((max, value) => Math.max(max, Math.abs(value)), Math.abs(avg));
  if (scale === 0) return 0;
  const normalizedAverage = avg / scale;
  const variance = values.reduce((sum, value) => {
    const difference = value / scale - normalizedAverage;
    return sum + difference * difference;
  }, 0) / values.length;
  return Math.sqrt(variance) * scale;
}

function quantile(sortedValues, q) {
  if (!sortedValues.length) return 0;
  const position = (sortedValues.length - 1) * q;
  const base = Math.floor(position);
  const rest = position - base;
  if (sortedValues[base + 1] !== undefined) {
    return sortedValues[base] * (1 - rest) + sortedValues[base + 1] * rest;
  }
  return sortedValues[base];
}

function numericExtents(values) {
  if (!values.length) return { min: 0, max: 0 };
  let min = Infinity;
  let max = -Infinity;
  values.forEach((value) => {
    if (value < min) min = value;
    if (value > max) max = value;
  });
  return { min, max };
}
