// 数据质量：类型转换失败计数与样例、精确重复行判断与去重、ID 字段重复值、疑似缺失值、疑似数值列、缺失率等级，以及质量页使用的原始数据数值统计。

function countConversionFailuresForValues(values, typeKey, field) {
  if (typeKey !== "numeric" && typeKey !== "date") return 0;
  const convert = typeKey === "numeric" ? getFieldNumberParser(field) : toDate;
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
  const confirmedKeys = getConfirmedPlaceholderKeys(field);
  return countConversionFailuresForValues(
    state.rows.map((row) => row[field]).filter((value) => !isMissingOrConfirmed(value, confirmedKeys)),
    typeKey,
    field
  );
}

function getFieldConversionFailureSamples(field, typeKey) {
  if (typeKey !== "numeric" && typeKey !== "date") return [];
  const profile = getFieldProfile(field);
  if (profile?.conversionFailureSamplesByType) {
    return [...(profile.conversionFailureSamplesByType[typeKey] || [])];
  }
  const convert = typeKey === "numeric" ? getFieldNumberParser(field) : toDate;
  const confirmedKeys = getConfirmedPlaceholderKeys(field);
  const samples = [];
  state.rows.forEach((row) => {
    if (!isMissingOrConfirmed(row[field], confirmedKeys) && convert(row[field]) === null) addConversionFailureSample(samples, row[field]);
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
// after normalizeValue() (trimmed, case kept); missing values and placeholder keys
// confirmed in confirmations are skipped. Listed values run from most to least
// frequent, ties by first appearance; the identical-row counts (buildExactRowKey
// over all fields) cover listed values only. getRowNumber(row) gives a row's file
// row or null.
function buildIdDuplicateSummary(rows, fields, profiles, getRowNumber, confirmations = {}) {
  const rowKeys = new Array(rows.length);
  const rowKeyAt = (index) => (rowKeys[index] ??= buildExactRowKey(rows[index], fields));
  const idFields = fields.filter((field) => profiles.some((profile) => profile.field === field && profile.typeKey === "id"));
  return idFields.map((field) => {
    const confirmedKeys = getConfirmedPlaceholderKeys(field, confirmations);
    const indexesByValue = new Map();
    let nonMissingCount = 0;
    rows.forEach((row, index) => {
      const value = analysisValue(row[field], confirmedKeys);
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

// Fields with a placeholder key still to confirm; ignored fields do not count.
// Report snapshots list pending entries only, which have neither flag.
function countPendingPlaceholderFields(entries) {
  return new Set(entries.filter((entry) => !entry.confirmed && !entry.fieldIgnored).map((entry) => entry.field)).size;
}

// Placeholder keys confirmed as missing for a field (an empty array when none).
function getConfirmedPlaceholderKeys(field, confirmations = state.placeholderConfirmations) {
  return Object.prototype.hasOwnProperty.call(confirmations, field) ? confirmations[field] : [];
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
    const confirmedKeys = getConfirmedPlaceholderKeys(field, confirmations);
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

// A field's profile figures if nextKeys (currentKeys plus more) were confirmed.
// Placeholders never convert to numbers or dates, so the numeric and date counts
// stay and only the non-missing and unique counts drop. Placeholders give no date
// order evidence, and only a year-named field can allow bare years, whose share
// depends on the non-missing count: such a field is re-profiled when its date
// strategy changes.
function estimateConfirmedProfile(profile, rows, currentKeys, nextKeys, numberFormatHint = "", manualTypeKey = "") {
  const field = profile.field;
  const yearField = isYearFieldName(field);
  const remainingValues = [];
  const removedValues = new Set();
  let removedCount = 0;
  rows.forEach((row) => {
    const value = row[field];
    if (isMissingOrConfirmed(value, currentKeys)) return;
    if (isMissingOrConfirmed(value, nextKeys)) {
      removedCount += 1;
      removedValues.add(normalizeValue(value));
    } else if (yearField) {
      remainingValues.push(value);
    }
  });
  let counts;
  if (!yearField || JSON.stringify(buildDateStrategy(field, remainingValues)) === JSON.stringify(profile.dateStrategy)) {
    counts = {
      nonMissingCount: profile.nonMissingCount - removedCount,
      uniqueCount: profile.uniqueCount - removedValues.size,
      numericCount: profile.numericCount,
      dateCount: profile.dateCount,
      estimated: true
    };
  } else {
    const full = buildColumnProfile(field, rows, numberFormatHint, nextKeys);
    counts = { nonMissingCount: full.nonMissingCount, uniqueCount: full.uniqueCount, numericCount: full.numericCount, dateCount: full.dateCount, estimated: false };
  }
  const inferredTypeKey = inferFieldTypeKey(field, counts);
  const typeKey = manualTypeKey || inferredTypeKey;
  const convertedCount = { numeric: counts.numericCount, date: counts.dateCount }[typeKey];
  return {
    typeKey,
    inferredTypeKey,
    missingCount: rows.length - counts.nonMissingCount,
    uniqueCount: counts.uniqueCount,
    conversionFailureCount: convertedCount === undefined ? 0 : counts.nonMissingCount - convertedCount,
    estimated: counts.estimated
  };
}

// A category field reads as numeric-looking when, without missing values and
// placeholder spellings, at least this many values remain and their numeric share
// is at least the minimum but below the automatic numeric threshold.
const NUMERIC_LOOKING_MIN_VALUES = 5;
const NUMERIC_LOOKING_MIN_RATIO = 0.5;
const NUMERIC_LOOKING_MAX_RATIO = 0.8;

// Numeric-looking fields among the applied category fields, in field order, from
// each profile's numericLooking figures. Only a hint: types stay as they are.
function buildNumericLookingSummary(profiles) {
  return profiles.flatMap((profile) => {
    if (profile.typeKey !== "category" || !profile.numericLooking) return [];
    const { valueCount, numericCount, samples } = profile.numericLooking;
    const ratio = valueCount ? numericCount / valueCount : 0;
    if (valueCount < NUMERIC_LOOKING_MIN_VALUES || ratio < NUMERIC_LOOKING_MIN_RATIO || ratio >= NUMERIC_LOOKING_MAX_RATIO) return [];
    return [{ field: profile.field, valueCount, numericCount, ratio, samples: [...samples] }];
  });
}

function missingRateLevel() {
  const denominator = state.rows.length * state.fields.length;
  const rate = denominator ? state.totalMissing / denominator : 0;
  if (rate >= 0.2) return "danger";
  if (rate > 0) return "warning";
  return "success";
}
