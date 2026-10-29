// 数据质量：类型转换失败计数与样例、精确重复行判断与去重、ID 字段重复值、疑似缺失值、按单位读取的数字、疑似数值列、缺失率等级，以及质量页使用的原始数据数值统计。

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
    const { indexesByValue, nonMissingCount } = groupIdValueIndexes(rows, field, getConfirmedPlaceholderKeys(field, confirmations));
    const duplicates = sortIdDuplicates(indexesByValue);
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

// Row indexes of each non-missing value of an ID field, compared after
// analysisValue() (so confirmed placeholder keys are missing), values in order of
// first appearance; nonMissingCount counts the values grouped.
function groupIdValueIndexes(rows, field, confirmedKeys) {
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
  return { indexesByValue, nonMissingCount };
}

// [value, indexes] of the values seen more than once, most frequent first, ties by
// first appearance: the order of the quality page's duplicate table.
function sortIdDuplicates(indexesByValue) {
  return [...indexesByValue].filter(([, indexes]) => indexes.length > 1)
    .sort(([, a], [, b]) => b.length - a.length || a[0] - b[0]);
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

// Fields whose number strategy has a switch or a unit conflict, ignored fields
// left out, in field order. Conflicting units run from most to fewest values,
// ties by first appearance.
function buildNumberConversionSummary(profiles) {
  return profiles.flatMap((profile) => {
    const strategy = profile.numberStrategy;
    if (profile.typeKey === "ignore" || !strategy || !(hasNumberStrategySwitch(strategy) || strategy.unitConflict)) return [];
    const { unitCount, unitSamples, magnitudeCount, fullwidthCount } = profile.numberReadings;
    const conflictUnits = strategy.unitConflict
      ? Object.entries(strategy.units).map(([unit, count]) => ({ unit, count })).sort((a, b) => b.count - a.count)
      : [];
    if (!unitCount && !magnitudeCount && !fullwidthCount && !conflictUnits.length) return [];
    return [{ field: profile.field, unit: strategy.unit, unitCount, unitSamples: [...unitSamples], magnitudeCount, fullwidthCount, conflictUnits }];
  });
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

// ---- Problem rows: the rows behind each quality-page count ----
// A category is { kind, field, key } with kind one of missing, conversion, outlier,
// duplicateRows, idDuplicate, placeholder, parseWarnings. Every list comes from the
// same checks as its quality count and reads the original rows (state.rows), never
// the filters. Items are { index (into state.rows, or null), line (file row or
// null), note, ... }, in file order; the duplicate kinds go group by group.

function problemRowItem(index, note, extra = {}) {
  return { index, line: getSourceRowNumber(state.rows[index]), note, ...extra };
}

// Missing values of a field, counted as its profile counts missingCount (with the
// placeholder keys the profile was built with). Ignored fields keep theirs.
function listMissingRows(field) {
  const profile = getFieldProfile(field);
  if (!profile) return null;
  const items = [];
  state.rows.forEach((row, index) => {
    if (isMissingOrConfirmed(row[field], profile.confirmedKeys)) items.push(problemRowItem(index, isMissing(row[field]) ? "空值" : "已按缺失处理"));
  });
  return { items, count: profile.missingCount };
}

// Values that fail the applied numeric or date type, by getProfileConversionCheck().
function listConversionFailureRows(field) {
  const profile = getFieldProfile(field);
  const fails = profile && getProfileConversionCheck(profile, profile.typeKey);
  if (!fails) return null;
  const items = [];
  state.rows.forEach((row, index) => {
    const value = row[field];
    if (!isMissingOrConfirmed(value, profile.confirmedKeys) && fails(value)) items.push(problemRowItem(index, `无法转换为${profile.type}`));
  });
  return { items, count: profile.conversionFailureCount };
}

// Values strictly outside the IQR bounds of the quality statistics, read with the
// same number parser; values that fail conversion are never outliers.
function listOutlierRows(field) {
  const stat = getQualityNumericStats().find((item) => item.field === field);
  const profile = getFieldProfile(field);
  if (!stat || !profile) return null;
  const parseNumber = getNumberParser(profile.numberStrategy);
  const above = `高于上界 ${formatFieldNumber(field, stat.upperBound)}`;
  const below = `低于下界 ${formatFieldNumber(field, stat.lowerBound)}`;
  const items = [];
  state.rows.forEach((row, index) => {
    if (isMissingOrConfirmed(row[field], profile.confirmedKeys)) return;
    const number = parseNumber(row[field]);
    if (number === null) return;
    if (number > stat.upperBound) items.push(problemRowItem(index, above));
    else if (number < stat.lowerBound) items.push(problemRowItem(index, below));
  });
  return { items, count: stat.outlierCount };
}

// Exactly repeated rows (buildExactRowKey() over every field), grouped, groups by
// first appearance. The first row of a group is kept by the deduplicated export, so
// the quality count (state.duplicateRows) leaves it out: items = count + groupCount.
function listDuplicateRowGroups() {
  if (!state.duplicateRows) return null;
  const groups = new Map();
  state.rows.forEach((row, index) => {
    const key = buildExactRowKey(row, state.fields);
    const indexes = groups.get(key);
    if (indexes) indexes.push(index);
    else groups.set(key, [index]);
  });
  const items = [];
  let group = 0;
  groups.forEach((indexes) => {
    if (indexes.length < 2) return;
    group += 1;
    indexes.forEach((index, position) => items.push(problemRowItem(index, `第 ${group} 组${position ? "" : "（保留）"}`, { group, kept: !position })));
  });
  return { items, count: state.duplicateRows, groupCount: group };
}

// Every duplicate value of an ID field and every row of it, in the order of the
// quality page's table (sortIdDuplicates()) but without its limits. Mostly repeated
// fields get no list.
function listIdDuplicateRows(field) {
  const summary = state.idDuplicates.find((item) => item.field === field);
  const profile = getFieldProfile(field);
  if (!summary || !profile || summary.mostlyRepeated || !summary.duplicateValueCount) return null;
  const { indexesByValue } = groupIdValueIndexes(state.rows, field, profile.confirmedKeys);
  const items = [];
  sortIdDuplicates(indexesByValue).forEach(([value, indexes], position) => {
    const note = `第 ${position + 1} 组：「${value}」出现 ${indexes.length} 次`;
    indexes.forEach((index) => items.push(problemRowItem(index, note, { group: position + 1, value })));
  });
  return { items, count: items.length, valueCount: summary.duplicateValueCount };
}

// Rows holding one placeholder key of a field, confirmed or not.
function listPlaceholderRows(field, key) {
  const entry = state.placeholders.find((item) => item.field === field && item.key === key);
  if (!entry) return null;
  const note = entry.confirmed ? "已按缺失处理" : "待确认";
  const items = [];
  state.rows.forEach((row, index) => {
    if (placeholderKey(row[field]) === key) items.push(problemRowItem(index, note));
  });
  return { items, count: entry.count };
}

// One item per parse warning, in warning order. A warning whose line is not among
// the imported rows (a dropped blank row) or that has no line gets no row.
function listParseWarningRows() {
  if (!state.parseWarnings.length) return null;
  const indexByLine = new Map();
  state.rows.forEach((row, index) => {
    const line = getSourceRowNumber(row);
    if (line !== null && !indexByLine.has(line)) indexByLine.set(line, index);
  });
  const items = state.parseWarnings.map((warning) => {
    const index = warning.row === null ? null : indexByLine.get(warning.row) ?? null;
    const reason = warning.row === null ? "（没有行号）" : index === null ? "（这一行没有可分析的数据，未导入）" : "";
    return { index, line: warning.row, note: `${warning.message}${reason}` };
  });
  return { items, count: items.length };
}

// Categories with rows to show, in this order: missing, format problems, outliers
// (per field), duplicate rows, ID duplicates (per field), placeholders (per field
// and key), parse warnings. count / unit describe the quality figure for labels:
// row, value, duplicateRow (without each group's kept row), duplicateValue, warning.
function listProblemCategories() {
  return [
    ...state.profiles.filter((profile) => profile.missingCount > 0)
      .map((profile) => ({ kind: "missing", field: profile.field, count: profile.missingCount, unit: "row" })),
    ...state.profiles.filter((profile) => profile.conversionFailureCount > 0)
      .map((profile) => ({ kind: "conversion", field: profile.field, count: profile.conversionFailureCount, unit: "value" })),
    ...getQualityNumericStats().filter((stat) => stat.outlierCount > 0)
      .map((stat) => ({ kind: "outlier", field: stat.field, count: stat.outlierCount, unit: "value" })),
    ...(state.duplicateRows > 0 ? [{ kind: "duplicateRows", count: state.duplicateRows, unit: "duplicateRow" }] : []),
    ...state.idDuplicates.filter((item) => !item.mostlyRepeated && item.duplicateValueCount > 0)
      .map((item) => ({ kind: "idDuplicate", field: item.field, count: item.duplicateValueCount, unit: "duplicateValue" })),
    ...state.placeholders.map((entry) => ({ kind: "placeholder", field: entry.field, key: entry.key, count: entry.count, unit: "row" })),
    ...(state.parseWarnings.length ? [{ kind: "parseWarnings", count: state.parseWarnings.length, unit: "warning" }] : [])
  ];
}

// The rows of one category, or null when it has none to show now.
function listProblemRows(category) {
  switch (category?.kind) {
    case "missing": return listMissingRows(category.field);
    case "conversion": return listConversionFailureRows(category.field);
    case "outlier": return listOutlierRows(category.field);
    case "duplicateRows": return listDuplicateRowGroups();
    case "idDuplicate": return listIdDuplicateRows(category.field);
    case "placeholder": return listPlaceholderRows(category.field, category.key);
    case "parseWarnings": return listParseWarningRows();
    default: return null;
  }
}

function missingRateLevel() {
  const denominator = state.rows.length * state.fields.length;
  const rate = denominator ? state.totalMissing / denominator : 0;
  if (rate >= 0.2) return "danger";
  if (rate > 0) return "warning";
  return "success";
}
