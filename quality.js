// 数据质量：类型转换失败计数与样例、精确重复行判断与去重、缺失率等级，以及质量页使用的原始数据数值统计。

function countConversionFailuresForValues(values, typeKey) {
  if (typeKey !== "numeric" && typeKey !== "date") return 0;
  const convert = typeKey === "numeric" ? toNumber : toDate;
  return values.reduce((count, value) => (
    count + (!isMissing(value) && convert(value) === null ? 1 : 0)
  ), 0);
}

function countFieldConversionFailures(field, typeKey) {
  if (typeKey !== "numeric" && typeKey !== "date") return 0;
  const profile = getFieldProfile(field);
  if (profile?.conversionFailuresByType) {
    return profile.conversionFailuresByType[typeKey] || 0;
  }
  return countConversionFailuresForValues(
    state.rows.map((row) => row[field]).filter((value) => !isMissing(value)),
    typeKey
  );
}

function getFieldConversionFailureSamples(field, typeKey) {
  if (typeKey !== "numeric" && typeKey !== "date") return [];
  const profile = getFieldProfile(field);
  if (profile?.conversionFailureSamplesByType) {
    return [...(profile.conversionFailureSamplesByType[typeKey] || [])];
  }
  const convert = typeKey === "numeric" ? toNumber : toDate;
  const samples = [];
  state.rows.forEach((row) => {
    if (!isMissing(row[field]) && convert(row[field]) === null) addConversionFailureSample(samples, row[field]);
  });
  return samples;
}

function getQualityNumericStats() {
  return state.filterSourceRows === state.rows
    ? state.qualityNumericStats
    : buildNumericStats(state.rows);
}

function countDuplicateRows(rows, fields) {
  const seen = new Set();
  let duplicates = 0;
  rows.forEach((row) => {
    const key = buildExactRowKey(row, fields);
    if (seen.has(key)) duplicates += 1;
    else seen.add(key);
  });
  return duplicates;
}

function deduplicateRows(rows, fields) {
  const seen = new Set();
  return rows.filter((row) => {
    const key = buildExactRowKey(row, fields);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function buildExactRowKey(row, fields) {
  return JSON.stringify(fields.map((field) => {
    if (!Object.prototype.hasOwnProperty.call(row, field)) return ["absent"];
    const value = row[field];
    if (value === undefined) return ["undefined"];
    if (value === null) return ["null"];
    return [typeof value, String(value)];
  }));
}

function missingRateLevel() {
  const denominator = state.rows.length * state.fields.length;
  const rate = denominator ? state.totalMissing / denominator : 0;
  if (rate >= 0.2) return "danger";
  if (rate > 0) return "warning";
  return "success";
}
