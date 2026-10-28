// 数据质量：类型转换失败计数与样例、精确重复行判断与去重、ID 字段重复值、疑似缺失值、缺失率等级，以及质量页使用的原始数据数值统计。

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

// An ID field with at least this many values and a unique share below this ratio
// looks like several records per object: it gets a hint instead of a value list.
const ID_MOSTLY_REPEATED_MIN_VALUES = 12;
const ID_MOSTLY_REPEATED_UNIQUE_RATIO = 0.5;
// Each ID field lists at most this many duplicate values, each with at most this
// many row numbers.
const ID_DUPLICATE_VALUE_LIMIT = 10;
const ID_DUPLICATE_ROW_NUMBER_LIMIT = 10;

// Duplicate values of every ID field (typeKey "id"), in field order. Values compare
// after normalizeValue() (trimmed, case kept); missing values are skipped. Listed
// values run from most to least frequent, ties by first appearance; the
// identical-row counts (buildExactRowKey over all fields) cover listed values only.
// getRowNumber(row) gives a row's file row or null.
function buildIdDuplicateSummary(rows, fields, profiles, getRowNumber) {
  const rowKeys = new Array(rows.length);
  const rowKeyAt = (index) => (rowKeys[index] ??= buildExactRowKey(rows[index], fields));
  const idFields = fields.filter((field) => profiles.some((profile) => profile.field === field && profile.typeKey === "id"));
  return idFields.map((field) => {
    const indexesByValue = new Map();
    let nonMissingCount = 0;
    rows.forEach((row, index) => {
      const value = normalizeValue(row[field]);
      if (value === "") return;
      nonMissingCount += 1;
      const indexes = indexesByValue.get(value);
      if (indexes) indexes.push(index);
      else indexesByValue.set(value, [index]);
    });
    const duplicates = [...indexesByValue].filter(([, indexes]) => indexes.length > 1)
      .sort(([, a], [, b]) => b.length - a.length || a[0] - b[0]);
    const uniqueRatio = nonMissingCount ? indexesByValue.size / nonMissingCount : 0;
    const mostlyRepeated = nonMissingCount >= ID_MOSTLY_REPEATED_MIN_VALUES && uniqueRatio < ID_MOSTLY_REPEATED_UNIQUE_RATIO;
    const listed = mostlyRepeated ? [] : duplicates.slice(0, ID_DUPLICATE_VALUE_LIMIT);
    return {
      field,
      nonMissingCount,
      distinctCount: indexesByValue.size,
      uniqueRatio,
      mostlyRepeated,
      duplicateValueCount: duplicates.length,
      values: listed.map(([value, indexes]) => {
        const sameRowCounts = new Map();
        indexes.forEach((index) => sameRowCounts.set(rowKeyAt(index), (sameRowCounts.get(rowKeyAt(index)) || 0) + 1));
        const identicalGroups = [...sameRowCounts.values()].filter((count) => count > 1);
        return {
          value,
          count: indexes.length,
          rowNumbers: indexes.slice(0, ID_DUPLICATE_ROW_NUMBER_LIMIT).map((index) => getRowNumber(rows[index])),
          identicalRowCount: identicalGroups.reduce((sum, count) => sum + count, 0),
          identicalGroupCount: identicalGroups.length
        };
      })
    };
  });
}

// ID fields that list duplicate values; mostly repeated fields do not count.
function countListedIdDuplicateFields(idDuplicates) {
  return idDuplicates.filter((item) => item.values.length).length;
}

// Placeholder candidates of the original rows: one entry per field and
// placeholderKey(), in field order, keys by first appearance. Ignored fields add no
// new candidates but keep the keys confirmed before. confirmations maps a field to
// its confirmed keys; getRowNumber(row) gives a row's file row or null.
function buildPlaceholderSummary(rows, fields, profiles, confirmations, getRowNumber) {
  return fields.flatMap((field) => {
    const profile = profiles.find((item) => item.field === field);
    if (!profile) return [];
    const fieldIgnored = profile.typeKey === "ignore";
    const confirmedKeys = Object.prototype.hasOwnProperty.call(confirmations, field) ? confirmations[field] : [];
    if (fieldIgnored && !confirmedKeys.length) return [];
    const entries = new Map();
    rows.forEach((row) => {
      const key = placeholderKey(row[field]);
      if (!key || (fieldIgnored && !confirmedKeys.includes(key))) return;
      let entry = entries.get(key);
      if (!entry) {
        entry = { field, key, tier: placeholderTier(key), count: 0, rawValues: [], rowNumbers: [], confirmed: confirmedKeys.includes(key), fieldIgnored };
        entries.set(key, entry);
      }
      entry.count += 1;
      const rawValue = row[field].trim();
      if (!entry.rawValues.includes(rawValue)) entry.rawValues.push(rawValue);
      if (entry.rowNumbers.length < ID_DUPLICATE_ROW_NUMBER_LIMIT) entry.rowNumbers.push(getRowNumber(row));
    });
    return [...entries.values()];
  });
}

function missingRateLevel() {
  const denominator = state.rows.length * state.fields.length;
  const rate = denominator ? state.totalMissing / denominator : 0;
  if (rate >= 0.2) return "danger";
  if (rate > 0) return "warning";
  return "success";
}
