const assert = require("node:assert/strict");
const crypto = require("node:crypto");
const fs = require("node:fs");
const path = require("node:path");
const { createScriptContext } = require("./test-context.cjs");

const { context, elements, source, evaluate } = createScriptContext();

const projectRoot = path.join(__dirname, "..");
const htmlSource = fs.readFileSync(path.join(projectRoot, "index.html"), "utf8");
const missingDomIds = [...new Set(
  [...source.matchAll(/document\.getElementById\("([^"]+)"\)/g)]
    .map((match) => match[1])
    .filter((id) => !new RegExp(`id=["']${id.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}["']`).test(htmlSource))
)];
assert.deepEqual(missingDomIds, []);
assert.doesNotMatch(htmlSource, /xlsx@0\.18\.5|xlsx-0\.18\.5/);
const scriptTags = [...htmlSource.matchAll(/<script\b[^>]*\bsrc="([^"]+)"[^>]*><\/script>/g)]
  .map((match) => ({ tag: match[0], src: match[1] }));
assert.deepEqual(scriptTags.filter(({ src }) => /^(?:https?:)?\/\//i.test(src)).map(({ src }) => src), []);
const vendorScripts = scriptTags.filter(({ src }) => src.startsWith("vendor/"));
assert.deepEqual(vendorScripts.map(({ src }) => src), [
  "vendor/papaparse-5.4.1.min.js",
  "vendor/chart-4.4.1.umd.min.js",
  "vendor/xlsx-0.20.3.full.min.js"
]);
const appScriptIndex = scriptTags.findIndex(({ src }) => src === "script.js");
assert.ok(appScriptIndex >= 0);
vendorScripts.forEach(({ tag, src }) => {
  assert.ok(scriptTags.findIndex((script) => script.src === src) < appScriptIndex, `${src} must load before script.js`);
  const integrity = tag.match(/\bintegrity="(sha384-[A-Za-z0-9+/=]+)"/)?.[1];
  assert.ok(integrity, `${src} is missing integrity`);
  const vendorPath = path.join(projectRoot, src);
  assert.equal(fs.existsSync(vendorPath), true, `missing vendor file: ${src}`);
  const actual = `sha384-${crypto.createHash("sha384").update(fs.readFileSync(vendorPath)).digest("base64")}`;
  assert.equal(actual, integrity, `${src} does not match its integrity`);
});
const licenseDir = path.join(projectRoot, "vendor", "licenses");
const licenseFiles = fs.existsSync(licenseDir) ? fs.readdirSync(licenseDir).sort() : [];
assert.deepEqual(licenseFiles, ["chartjs-LICENSE.md", "papaparse-LICENSE", "sheetjs-LICENSE"]);
assert.equal(fs.existsSync(path.join(projectRoot, ".nojekyll")), true);

const localAssetReferences = [...htmlSource.matchAll(/(?:src|href)="([^"]+)"/g)]
  .map((match) => match[1])
  .filter((reference) => !/^(?:https?:|data:|#|mailto:|javascript:)/i.test(reference));
assert.ok(localAssetReferences.length > 0);
localAssetReferences.forEach((reference) => {
  assert.equal(reference.startsWith("/"), false, `root-relative asset is not Pages-safe: ${reference}`);
  const localPath = reference.split(/[?#]/, 1)[0];
  assert.equal(fs.existsSync(path.join(projectRoot, localPath)), true, `missing local asset: ${reference}`);
});
assert.match(source, /fetch\("examples\/sample-data\.csv"\)/);
const resultSectionOrder = [
  "overviewSection",
  "schemaSection",
  "qualitySection",
  "filterSection",
  "statisticsSection",
  "chartsSection",
  "customSection",
  "insightsSection",
  "exportSection"
].map((id) => htmlSource.indexOf(`id="${id}"`));
assert.equal(resultSectionOrder.every((position, index) => position >= 0 && (!index || position > resultSectionOrder[index - 1])), true);
assert.match(htmlSource, /id="mobileSectionNav"/);
assert.match(htmlSource, /id="globalFilterSummary"/);
assert.match(htmlSource, /id="globalClearFilters"/);
assert.match(htmlSource, /href="#exportSection"/);
assert.doesNotMatch(htmlSource, /href="#(?:dataExportSection|reportSection)"/);
assert.match(htmlSource, /<title>Smart Tabular Analyzer \| 智能表格数据分析工具<\/title>/);
const legacyProductNamePattern = new RegExp(["Smart", "CSV", "Analyzer"].join(" "));
assert.doesNotMatch(htmlSource, legacyProductNamePattern);
assert.doesNotMatch(source, legacyProductNamePattern);

const utf8CsvText = "姓名,城市\n张三,上海";
const utf8CsvBytes = Array.from(new TextEncoder().encode(utf8CsvText));
const gbkCsvBytes = [
  0xd0, 0xd5, 0xc3, 0xfb, 0x2c, 0xb3, 0xc7, 0xca, 0xd0, 0x0a,
  0xd5, 0xc5, 0xc8, 0xfd, 0x2c, 0xc9, 0xcf, 0xba, 0xa3
];
const gb18030OnlyBytes = [0x95, 0x34, 0xb2, 0x35];

assert.equal(
  evaluate(`decodeCsvBuffer(new Uint8Array(${JSON.stringify(utf8CsvBytes)}).buffer, "utf-8").text`),
  utf8CsvText
);
assert.equal(
  evaluate(`decodeCsvBuffer(new Uint8Array(${JSON.stringify(gbkCsvBytes)}).buffer, "gbk").text`),
  utf8CsvText
);
assert.equal(
  evaluate(`decodeCsvBuffer(new Uint8Array(${JSON.stringify(gbkCsvBytes)}).buffer, "gb18030").text`),
  utf8CsvText
);
assert.equal(
  evaluate(`decodeCsvBuffer(new Uint8Array(${JSON.stringify(gbkCsvBytes)}).buffer, "auto").text`),
  utf8CsvText
);
assert.equal(
  evaluate(`decodeCsvBuffer(new Uint8Array(${JSON.stringify(gb18030OnlyBytes)}).buffer, "gb18030").text`),
  "𠮷"
);
assert.equal(
  evaluate(`decodeCsvBuffer(new Uint8Array(${JSON.stringify(gb18030OnlyBytes)}).buffer, "auto").text`),
  "𠮷"
);

assert.equal(evaluate('toNumber("10%")'), 0.1);
assert.equal(evaluate('toNumber("(10%)")'), -0.1);
assert.equal(evaluate('toNumber("￥1,234.50")'), 1234.5);
assert.equal(evaluate('toNumber("(1,250)")'), -1250);
assert.equal(evaluate('toNumber("2026-07-17")'), null);

assert.equal(evaluate('toDate("2026-07-17").toISOString()'), "2026-07-17T00:00:00.000Z");
assert.equal(evaluate('toDate("2026-02-30")'), null);
assert.equal(evaluate('toDate("07/08/2026")'), null);
assert.equal(evaluate('toDate("17/07/2026").toISOString()'), "2026-07-17T00:00:00.000Z");
assert.equal(evaluate('toDate("2026年7月19日").toISOString()'), "2026-07-19T00:00:00.000Z");
assert.equal(evaluate('toDate("2026年2月30日")'), null);
assert.equal(evaluate('toDate("2026-07-17 garbage")'), null);
assert.equal(evaluate('toDate("08/08/2026").toISOString()'), "2026-08-08T00:00:00.000Z");
assert.equal(evaluate('toDate(new Date(2026, 6, 24)).toISOString()'), "2026-07-24T00:00:00.000Z");
assert.equal(evaluate('toDate(new Date(Date.UTC(2026, 6, 24))).toISOString()'), "2026-07-24T00:00:00.000Z");

const studentRows = Array.from({ length: 12 }, (_, index) => ({ 学号: String(20260001 + index) }));
context.__studentRows = studentRows;
assert.equal(evaluate('buildColumnProfile("学号", __studentRows).typeKey'), "id");
assert.equal(evaluate('buildColumnProfile("学号数量", __studentRows.map((row, index) => ({"学号数量": index + 1}))).typeKey'), "numeric");
assert.equal(evaluate('buildColumnProfile("编号", __studentRows.map((row, index) => ({"编号": 100001 + index}))).typeKey'), "id");
assert.equal(evaluate('buildColumnProfile("记录值", __studentRows.map((row, index) => ({"记录值": 100001 + index}))).typeKey'), "numeric");

const yearRows = [
  { 年份: "2023", 销售额: "100" },
  { 年份: "2024", 销售额: "200" },
  { 年份: "2025", 销售额: "300" }
];
context.__yearRows = yearRows;
assert.equal(evaluate('buildColumnProfile("年份", __yearRows).typeKey'), "date");
assert.equal(evaluate('buildColumnProfile("年份", __yearRows).dateStrategy.granularity'), "year");
assert.equal(evaluate('buildColumnProfile("年份", __yearRows).conversionFailuresByType.date'), 0);
assert.equal(evaluate('buildColumnProfile("销售额", __yearRows).typeKey'), "numeric");
assert.equal(evaluate('buildColumnProfile("年龄", [{"年龄":"20"},{"年龄":"21"}]).typeKey'), "numeric");
evaluate('state.rows = __yearRows; state.profiles = [buildColumnProfile("年份", __yearRows), buildColumnProfile("销售额", __yearRows)]');
assert.equal(
  evaluate('JSON.stringify(buildDateTrend("年份", "销售额").labels)'),
  JSON.stringify(["2023", "2024", "2025"])
);
assert.equal(
  evaluate('JSON.stringify(buildMetricDateTrend("年份", "销售额").labels)'),
  JSON.stringify(["2023", "2024", "2025"])
);
assert.equal(
  evaluate('JSON.stringify(buildTrendFromRows("年份", (row) => toNumber(row["销售额"]), "sum").labels)'),
  JSON.stringify(["2023", "2024", "2025"])
);

const mixedDateRows = [
  { 日期: "2025-01-02", 指标: "10" },
  { 日期: "2025/1/3", 指标: "20" },
  { 日期: "13/01/2025", 指标: "30" },
  { 日期: "2025年1月4日", 指标: "40" },
  { 日期: "01/02/2025", 指标: "50" }
];
context.__mixedDateRows = mixedDateRows;
assert.equal(evaluate('buildColumnProfile("日期", __mixedDateRows).dateStrategy.order'), "dmy");
assert.equal(evaluate('buildColumnProfile("日期", __mixedDateRows).dateRatio'), 1);
assert.equal(evaluate('buildColumnProfile("日期", __mixedDateRows).typeKey'), "date");
evaluate('state.rows = __mixedDateRows; state.profiles = [buildColumnProfile("日期", __mixedDateRows), buildColumnProfile("指标", __mixedDateRows)]');
assert.equal(evaluate('parseFieldDate("日期", "01/02/2025").toISOString()'), "2025-02-01T00:00:00.000Z");
assert.equal(evaluate('buildDateTrend("日期", "指标").labels.length'), 5);

const conflictingDateRows = [
  { 日期: "20/07/2026" },
  { 日期: "07/20/2026" },
  { 日期: "07/08/2026" }
];
context.__conflictingDateRows = conflictingDateRows;
assert.equal(evaluate('buildColumnProfile("日期", __conflictingDateRows).dateStrategy.order'), "conflict");
assert.equal(evaluate('buildColumnProfile("日期", __conflictingDateRows).conversionFailuresByType.date'), 1);
assert.equal(evaluate('buildColumnProfile("日期", [{"日期":"31/02/2026"},{"日期":"07/08/2026"}]).dateStrategy.order'), "none");

const formattedNumericRows = [
  { 价格: "￥1,234.50", 转化率: "10%" },
  { 价格: "2,000", 转化率: "25%" },
  { 价格: "980", 转化率: "100%" },
  { 价格: "￥3,500", 转化率: "0%" },
  { 价格: "错误", 转化率: "错误" },
  { 价格: "", 转化率: "" }
];
context.__formattedNumericRows = formattedNumericRows;
assert.equal(evaluate('buildColumnProfile("价格", __formattedNumericRows).typeKey'), "numeric");
assert.equal(evaluate('buildColumnProfile("价格", __formattedNumericRows).numericFormat'), "number");
assert.equal(evaluate('buildColumnProfile("价格", __formattedNumericRows).conversionFailuresByType.numeric'), 1);
assert.equal(evaluate('buildColumnProfile("转化率", __formattedNumericRows).typeKey'), "numeric");
assert.equal(evaluate('buildColumnProfile("转化率", __formattedNumericRows).numericFormat'), "percent");
assert.equal(evaluate('buildColumnProfile("转化率", __formattedNumericRows).conversionFailuresByType.numeric'), 1);
assert.equal(evaluate('buildColumnProfile("混合比例", [{"混合比例":"10%"},{"混合比例":"0.2"}]).numericFormat'), "number");
evaluate('state.rows = __formattedNumericRows; state.profiles = [buildColumnProfile("价格", __formattedNumericRows), buildColumnProfile("转化率", __formattedNumericRows)]');
assert.equal(evaluate('formatFieldNumber("转化率", 0.3375)'), "33.75%");
assert.equal(evaluate('formatFieldNumber("价格", 1234.5)'), "1,234.5");
assert.equal(evaluate('chartOptions("平均值", "转化率").scales.y.ticks.callback(0.1)'), "10%");

context.__largeValues = Array.from({ length: 125_000 }, (_, index) => index - 62_500);
assert.equal(evaluate("summarizeNumbers(__largeValues).min"), -62_500);
assert.equal(evaluate("summarizeNumbers(__largeValues).max"), 62_499);
assert.equal(
  evaluate("buildHistogram(__largeValues, 10).reduce((total, bin) => total + bin.count, 0)"),
  125_000
);
assert.equal(
  evaluate('summarizeGroupedValues(new Map([["全部", __largeValues]]))[0].max'),
  62_499
);

assert.equal(evaluate('countConversionFailuresForValues(["0", "bad", "", null], "numeric")'), 1);
assert.equal(evaluate('countConversionFailuresForValues(["2026-01-01", "not-a-date", " "], "date")'), 1);
assert.equal(evaluate('countConversionFailuresForValues(["任意文本"], "category")'), 0);

const excelData = evaluate(`JSON.stringify(buildExcelTabularData(
  [["", "", ""], ["订单编号", "地区", "销售额"], ["A-001", "华东", "1,200"], ["A-002", "华南", "980"]],
  [["", "", ""], ["订单编号", "地区", "销售额"], ["A-001", "华东", 1200], ["A-002", "华南", 980]],
  "销售数据"
))`);
assert.deepEqual(JSON.parse(excelData), {
  fields: ["订单编号", "地区", "销售额"],
  rows: [
    { 订单编号: "A-001", 地区: "华东", 销售额: "1200" },
    { 订单编号: "A-002", 地区: "华南", 销售额: "980" }
  ]
});

const excelPrecisionResult = JSON.parse(evaluate(`JSON.stringify(buildExcelTabularData(
  [["精确值"], ["1.23"]],
  [["精确值"], [1.23456789]],
  "精度测试"
).rows)`));
assert.deepEqual(excelPrecisionResult, [{ 精确值: "1.23456789" }]);

const excelRawDateResult = JSON.parse(evaluate(`JSON.stringify((() => {
  const data = buildExcelTabularData(
    [["日期"], ["7/17/26"]],
    [["日期"], [new Date(Date.UTC(2026, 6, 17))]],
    "日期表"
  );
  return {
    type: buildColumnProfile("日期", data.rows).typeKey,
    iso: toDate(data.rows[0]["日期"]).toISOString()
  };
})())`));
assert.deepEqual(excelRawDateResult, {
  type: "date",
  iso: "2026-07-17T00:00:00.000Z"
});

const excelProtoResult = JSON.parse(evaluate(`JSON.stringify((() => {
  const data = buildExcelTabularData(
    [["__proto__"], ["excel-value"]],
    [["__proto__"], ["excel-value"]],
    "特殊表头"
  );
  const row = data.rows[0];
  return {
    own: Object.prototype.hasOwnProperty.call(row, "__proto__"),
    value: row["__proto__"]
  };
})())`));
assert.deepEqual(excelProtoResult, { own: true, value: "excel-value" });

function excelError(formattedRows, rawRows, sheetName = "测试表") {
  return evaluate(`(() => {
    try {
      buildExcelTabularData(${JSON.stringify(formattedRows)}, ${JSON.stringify(rawRows)}, ${JSON.stringify(sheetName)});
      return "";
    } catch (error) {
      return error.message;
    }
  })()`);
}

assert.match(excelError([], []), /为空/);
assert.match(excelError([["字段A", "字段B"]], [["字段A", "字段B"]]), /只有表头/);
assert.match(excelError([["张三", "20"]], [["张三", 20]]), /未识别到有效文本表头/);
assert.match(excelError([["字段A", ""], ["值A", "值B"]], [["字段A", ""], ["值A", "值B"]]), /缺少表头/);
assert.match(excelError([["字段A", "字段A"], ["值A", "值B"]], [["字段A", "字段A"], ["值A", "值B"]]), /重复表头/);

assert.doesNotThrow(() => evaluate('validateExcelFileSignature(new Uint8Array([0x50,0x4b,0x03,0x04]), "valid.xlsx")'));
assert.doesNotThrow(() => evaluate('validateExcelFileSignature(new Uint8Array([0xd0,0xcf,0x11,0xe0,0xa1,0xb1,0x1a,0xe1]), "valid.xls")'));
assert.match(
  evaluate(`(() => {
    try {
      validateExcelFileSignature(new Uint8Array([0x61,0x62,0x63]), "fake.xlsx");
      return "";
    } catch (error) {
      return error.message;
    }
  })()`),
  /格式与扩展名不匹配/
);

assert.equal(
  evaluate('buildColumnProfile("用户ID", [{"用户ID":"u1"},{"用户ID":"u1"},{"用户ID":"u2"}]).typeKey'),
  "id"
);
for (const field of ["paid", "valid", "grid", "PAID", "VALID", "GRID"]) {
  assert.equal(evaluate(`isIdFieldName(${JSON.stringify(field)})`), false);
}
for (const field of ["id", "user_id", "userId", "orderID", "orderid", "customerCode"]) {
  assert.equal(evaluate(`isIdFieldName(${JSON.stringify(field)})`), true);
}

const encodedProtoHeader = evaluate('encodeCsvHeaderForParser("__proto__", 0)');
const encodedNormalHeader = evaluate('encodeCsvHeaderForParser("普通字段", 1)');
context.__encodedProtoHeader = encodedProtoHeader;
context.__encodedNormalHeader = encodedNormalHeader;
context.__safeProtoCsvRow = {
  [encodedProtoHeader]: "csv-value",
  [encodedNormalHeader]: "normal-value"
};
const csvProtoResult = JSON.parse(evaluate(`JSON.stringify((() => {
  const importId = beginImport();
  handleParsedData({
    errors: [],
    meta: { fields: [__encodedProtoHeader, __encodedNormalHeader] },
    data: [__safeProtoCsvRow]
  }, "special-header.csv", importId);
  const row = state.rows[0];
  return {
    fields: state.fields,
    own: Object.prototype.hasOwnProperty.call(row, "__proto__"),
    value: row["__proto__"],
    normal: row["普通字段"]
  };
})())`));
assert.deepEqual(csvProtoResult, {
  fields: ["__proto__", "普通字段"],
  own: true,
  value: "csv-value",
  normal: "normal-value"
});

const firstImport = evaluate("beginImport()");
const secondImport = evaluate("beginImport()");
assert.equal(evaluate(`isCurrentImport(${firstImport})`), false);
assert.equal(evaluate(`isCurrentImport(${secondImport})`), true);

context.FileReader = { LOADING: 1 };
context.__activeReader = {
  readyState: 1,
  aborted: false,
  abort() { this.aborted = true; }
};
evaluate("state.activeFileReader = __activeReader");
evaluate("beginImport()");
assert.equal(context.__activeReader.aborted, true);
assert.equal(evaluate("state.activeFileReader"), null);

context.Papa = {
  WORKERS_SUPPORTED: true,
  parse(_text, config) { context.__workerConfig = config; }
};
context.window.Papa = context.Papa;
evaluate('parseCsvText("字段\\n值", "worker.csv", beginImport())');
assert.equal(context.__workerConfig.worker, false);

const oneColumnImport = evaluate("beginImport()");
evaluate(`handleParsedData({
  errors: [{type: "Delimiter", code: "UndetectableDelimiter", message: "bad delimiter"}],
  meta: {fields: ["字段"]},
  data: [{"字段": "1"}]
}, "one-column.csv", ${oneColumnImport})`);
assert.equal(evaluate("state.rows.length"), 1);
assert.equal(evaluate("state.fields.length"), 1);
assert.equal(evaluate("state.parseWarnings.length"), 0);
assert.notEqual(elements.get("statusPanel").className, "status-panel error");

const malformedImport = evaluate("beginImport()");
evaluate(`handleParsedData({
  errors: [
    {type: "Delimiter", code: "UndetectableDelimiter", message: "bad delimiter"},
    {type: "Quotes", code: "MissingQuotes", message: "unterminated quote"}
  ],
  meta: {fields: ["字段"]},
  data: [{"字段": "1"}]
}, "broken.csv", ${malformedImport})`);
assert.equal(evaluate("state.rows.length"), 0);
assert.equal(elements.get("statusPanel").className, "status-panel error");

const extraFieldImport = evaluate("beginImport()");
evaluate(`handleParsedData({
  errors: [{type: "FieldMismatch", code: "TooManyFields", message: "extra value", row: 0}],
  meta: {fields: ["字段"]},
  data: [{"字段": "保留", "__parsed_extra": ["不应丢弃"]}]
}, "extra-field.csv", ${extraFieldImport})`);
assert.equal(evaluate("state.rows.length"), 0);
assert.equal(elements.get("statusPanel").className, "status-panel error");
assert.match(elements.get("statusPanel").textContent, /数据列多于表头/);

const duplicateHeaderImport = evaluate("beginImport()");
evaluate(`handleParsedData({
  errors: [],
  meta: {fields: ["字段", " 字段 "]},
  data: [{"字段": "A", " 字段 ": "B"}]
}, "duplicate-header.csv", ${duplicateHeaderImport})`);
assert.equal(evaluate("state.rows.length"), 0);
assert.match(elements.get("statusPanel").textContent, /重复字段/);

const blankHeaderImport = evaluate("beginImport()");
evaluate(`handleParsedData({
  errors: [],
  meta: {fields: ["", "字段"]},
  data: [{"": "A", "字段": "B"}]
}, "blank-header.csv", ${blankHeaderImport})`);
assert.equal(evaluate("state.rows.length"), 0);
assert.match(elements.get("statusPanel").textContent, /缺少表头/);

const emptyTrailingColumnImport = evaluate("beginImport()");
evaluate(`handleParsedData({
  errors: [],
  meta: {fields: ["字段", ""]},
  data: [{"字段": "A", "": ""}]
}, "empty-trailing-column.csv", ${emptyTrailingColumnImport})`);
assert.equal(evaluate("state.rows.length"), 1);
assert.equal(evaluate("state.fields.length"), 1);

const warningImport = evaluate("beginImport()");
evaluate(`handleParsedData({
  errors: [{type: "FieldMismatch", code: "TooFewFields", message: "row mismatch", row: 0}],
  meta: {fields: ["用户ID", "数值"]},
  data: [{"用户ID": "u1", "数值": "10"}, {"用户ID": "u2", "数值": "20"}]
}, "warning.csv", ${warningImport})`);
assert.equal(evaluate("state.rows.length"), 2);
assert.equal(evaluate("state.parseWarnings.length"), 1);
assert.equal(elements.get("statusPanel").className, "status-panel warning");

const configImport = evaluate("beginImport()");
evaluate(`handleParsedData({
  errors: [],
  meta: {fields: ["订单号", "金额", "下单日期", "渠道", "备注"]},
  data: [
    {"订单号":"A001","金额":"100","下单日期":"2026-01-01","渠道":"线上","备注":"foo"},
    {"订单号":"A002","金额":"bad","下单日期":"not-a-date","渠道":"门店","备注":""},
    {"订单号":"A003","金额":"300","下单日期":"2026-01-03","渠道":"线上","备注":"bar"},
    {"订单号":"A004","金额":"","下单日期":"","渠道":"门店","备注":"baz"}
  ]
}, "field-config.csv", ${configImport})`);

assert.equal(
  evaluate('JSON.stringify(state.profiles.find((item) => item.field === "渠道").sampleValues)'),
  JSON.stringify(["线上", "门店"])
);
assert.equal(evaluate('state.profiles.find((item) => item.field === "金额").typeKey'), "category");
assert.equal(evaluate("state.numericStats.length"), 0);

evaluate('stageFieldTypeChange("金额", "numeric")');
evaluate('stageFieldTypeChange("下单日期", "date")');
evaluate('stageFieldTypeChange("备注", "ignore")');
assert.equal(evaluate('state.profiles.find((item) => item.field === "金额").typeKey'), "category");
assert.equal(evaluate('state.fieldTypeDrafts["金额"]'), "numeric");
assert.equal(evaluate('countFieldConversionFailures("金额", "numeric")'), 1);
assert.equal(evaluate('countFieldConversionFailures("下单日期", "date")'), 1);
assert.equal(evaluate("hasFieldConfigDraftChanges()"), true);

const originalRows = evaluate("JSON.stringify(state.rows)");
evaluate("applyFieldConfiguration()");
assert.equal(evaluate('state.profiles.find((item) => item.field === "金额").typeKey'), "numeric");
assert.equal(evaluate('state.profiles.find((item) => item.field === "金额").conversionFailureCount'), 1);
assert.equal(evaluate('state.profiles.find((item) => item.field === "下单日期").typeKey'), "date");
assert.equal(evaluate('state.profiles.find((item) => item.field === "下单日期").conversionFailureCount'), 1);
assert.equal(evaluate('state.profiles.find((item) => item.field === "备注").typeKey'), "ignore");
assert.equal(evaluate('state.numericStats.find((item) => item.field === "金额").count'), 2);
assert.equal(evaluate('state.numericStats.find((item) => item.field === "金额").mean'), 200);
assert.equal(
  evaluate('JSON.stringify(buildDateTrend("下单日期", "金额"))'),
  JSON.stringify({
    labels: ["2026-01-01", "2026-01-03"],
    values: [100, 300],
    label: "金额 平均值趋势",
    axisLabel: "平均值"
  })
);
assert.equal(evaluate('state.categoryStats.some((item) => item.field === "备注")'), false);
assert.equal(evaluate('getActiveProfiles().some((item) => item.field === "备注")'), false);
assert.equal(evaluate('buildTemplateFieldOptions({required:true, expected:"numeric"}).includes("金额")'), true);
assert.equal(evaluate('buildTemplateFieldOptions({required:true, expected:"any"}).includes("备注")'), false);
assert.equal(evaluate("JSON.stringify(state.rows)"), originalRows);

evaluate("restoreAutomaticFieldTypes()");
assert.equal(evaluate('state.profiles.find((item) => item.field === "金额").typeKey'), "category");
assert.equal(evaluate('state.profiles.find((item) => item.field === "下单日期").typeKey'), "category");
assert.equal(evaluate('state.profiles.find((item) => item.field === "备注").typeKey'), "category");
assert.equal(evaluate("state.numericStats.length"), 0);

evaluate('stageFieldTypeChange("渠道", "numeric")');
evaluate("applyFieldConfiguration()");
assert.equal(evaluate('state.profiles.find((item) => item.field === "渠道").conversionFailureCount'), 4);
assert.equal(evaluate("state.numericStats.length"), 0);
assert.equal(evaluate('buildTemplateFieldOptions({required:true, expected:"numeric"}).includes("渠道")'), false);
assert.equal(evaluate('JSON.stringify(buildDateTrend("下单日期", "渠道").labels)'), "[]");
evaluate("restoreAutomaticFieldTypes()");

// The trend conclusion compares record counts, so a metric that moves from 0
// is irrelevant: one record per day on both days is a flat count.
const flatCountInsight = "基于日期字段「date」观察，记录数趋势 从 2026-01-01 到 2026-01-02 基本持平。";
const zeroBaselineInsight = evaluate(`(() => {
  state.rows = [
    { date: "2026-01-01", metric: "0" },
    { date: "2026-01-02", metric: "10" }
  ];
  state.fields = ["date", "metric"];
  state.profiles = state.fields.map((field) => buildColumnProfile(field, state.rows));
  return buildTrendInsight("date");
})()`);
assert.equal(zeroBaselineInsight, flatCountInsight);

const negativeZeroBaselineInsight = evaluate(`(() => {
  state.rows = [
    { date: "2026-01-01", metric: "0" },
    { date: "2026-01-02", metric: "-10" }
  ];
  state.fields = ["date", "metric"];
  state.profiles = state.fields.map((field) => buildColumnProfile(field, state.rows));
  return buildTrendInsight("date");
})()`);
assert.equal(negativeZeroBaselineInsight, flatCountInsight);

// V2.1 report export: build a serializable report model that is shared by
// both exporters. Keep these tests DOM-independent so empty analytical
// sections can be verified without triggering a browser download.
const reportRows = [
  { 订单号: "A001", 地区: "华东", 销售额: "100", 日期: "2026-01-01", 备注: "" },
  { 订单号: "A002", 地区: "华南", 销售额: "200", 日期: "2026-01-02", 备注: "正常" },
  { 订单号: "A002", 地区: "华南", 销售额: "200", 日期: "2026-01-02", 备注: "正常" }
];
context.__reportRows = reportRows;
evaluate(`(() => {
  state.rows = __reportRows;
  state.fields = ["订单号", "地区", "销售额", "日期", "备注"];
  state.profiles = state.fields.map((field) => buildColumnProfile(field, state.rows));
  const orderProfile = state.profiles.find((profile) => profile.field === "订单号");
  orderProfile.typeKey = "id";
  orderProfile.type = FIELD_TYPE_LABELS.id;
  state.numericStats = buildNumericStats();
  state.categoryStats = buildCategoryStats();
  state.categoryStats.push({
    field: "测试 Top 10",
    total: 12,
    top: Array.from({length: 12}, (_, index) => ({
      name: "类别" + (index + 1),
      count: 12 - index,
      ratio: (12 - index) / 78
    }))
  });
  state.totalMissing = state.profiles.reduce((sum, profile) => sum + profile.missingCount, 0);
  state.duplicateRows = 1;
  state.parseWarnings = [{message: "第 3 行字段数量不一致", row: 3}];
  state.sourceFileName = "销售 数据.xlsx";
  state.sourceSheetName = "华东<明细>";
  state.analysisCompletedAt = new Date(2026, 6, 17, 14, 30, 5);
  state.insights = ["数据集包含 3 行、5 个字段。", "销售额整体上升。"];
  state.customAnalysis = {
    metricField: "销售额",
    groupField: "地区",
    dateField: "日期",
    grouped: [
      {name: "华南", count: 2, sum: 400, avg: 200},
      {name: "华东", count: 1, sum: 100, avg: 100}
    ],
    trend: {
      labels: ["2026-01-01", "2026-01-02"],
      values: [100, 400]
    }
  };
})()`);

const reportData = JSON.parse(evaluate(
  'JSON.stringify(buildReportData(new Date(2026, 6, 17, 14, 30, 5)))'
));
assert.equal(reportData.fileName, "销售 数据.xlsx");
assert.equal(reportData.sheetName, "华东<明细>");
assert.ok(reportData.analysisTime);
assert.equal(reportData.dataScale.rows, 3);
assert.equal(reportData.dataScale.fields, 5);
assert.equal(reportData.fieldTypes.length, 5);
assert.equal(reportData.fieldTypes.find((item) => item.field === "订单号").typeKey, "id");
assert.equal(reportData.fieldTypes.find((item) => item.field === "销售额").typeKey, "numeric");
assert.equal(reportData.quality.totalMissing, 1);
assert.equal(reportData.quality.duplicateRows, 1);
assert.equal(reportData.quality.parseWarnings.length, 1);
assert.equal(reportData.numericStats.find((item) => item.field === "销售额").mean, "166.67");
assert.equal(reportData.categoryStats.find((item) => item.field === "测试 Top 10").top.length, 10);
assert.equal(reportData.dateTrends.find((item) => item.field === "日期").labels.length, 2);
assert.equal(reportData.customAnalysis.metricField, "销售额");
assert.equal(reportData.customAnalysis.rows[0].group, "华南");
assert.ok(Array.isArray(reportData.insights));
assert.ok(reportData.insights.some((item) => item.includes("数据集包含")));

context.__reportData = reportData;
const reportHtml = evaluate(`buildHtmlReport(__reportData, [{
  id: "numericDistributionChart",
  title: "销售额 <分布>",
  dataUrl: "data:image/png;base64,ZmFrZS1jaGFydA=="
}])`);
assert.match(reportHtml, /^<!DOCTYPE html>/i);
assert.match(reportHtml, /<meta[^>]+charset=["']?UTF-8/i);
assert.match(reportHtml, /<style[\s>]/i);
assert.match(reportHtml, /销售 数据\.xlsx/);
assert.match(reportHtml, /华东&lt;明细&gt;/);
assert.match(reportHtml, /字段类型结果/);
assert.match(reportHtml, /数据质量报告/);
assert.match(reportHtml, /数值字段描述性统计/);
assert.match(reportHtml, /分类字段 Top 10/);
assert.match(reportHtml, /日期趋势摘要/);
assert.match(reportHtml, /自定义分组分析结果/);
assert.match(reportHtml, /自动分析结论/);
assert.match(reportHtml, /销售额 &lt;分布&gt;/);
assert.match(reportHtml, /data:image\/png;base64,ZmFrZS1jaGFydA==/);
assert.match(reportHtml, /Smart Tabular Analyzer/);
assert.match(reportHtml, /https:\/\/github\.com\/wy-data-30\/smart-tabular-analyzer/);
assert.doesNotMatch(reportHtml, legacyProductNamePattern);
assert.doesNotMatch(reportHtml, /<(?:script|link)[^>]+(?:src|href)=["']https?:/i);

const reportMarkdown = evaluate("buildMarkdownReport(__reportData)");
assert.match(reportMarkdown, /^# .+/);
assert.match(reportMarkdown, /销售 数据\.xlsx/);
assert.match(reportMarkdown, /华东&lt;明细&gt;/);
assert.match(reportMarkdown, /## 字段类型结果/);
assert.match(reportMarkdown, /## 数据质量报告/);
assert.match(reportMarkdown, /## 数值字段描述性统计/);
assert.match(reportMarkdown, /## 分类字段 Top 10/);
assert.match(reportMarkdown, /## 日期趋势摘要/);
assert.match(reportMarkdown, /## 自定义分组分析结果/);
assert.match(reportMarkdown, /## 自动分析结论/);
assert.match(reportMarkdown, /\| 销售额 \|/);
assert.match(reportMarkdown, /Smart Tabular Analyzer/);
assert.match(reportMarkdown, /https:\/\/github\.com\/wy-data-30\/smart-tabular-analyzer/);
assert.doesNotMatch(reportMarkdown, legacyProductNamePattern);

assert.match(
  evaluate('buildReportFileName("html", new Date(2026, 6, 17, 14, 30, 5))'),
  /^销售 数据(?:\.xlsx)?[-_].*20260717.*143005.*\.html$/
);
evaluate('state.sourceFileName = "客户:销售?.xlsx"');
const markdownFileName = evaluate('buildReportFileName("md", new Date(2026, 6, 17, 14, 30, 5))');
assert.match(markdownFileName, /20260717.*143005.*\.md$/);
assert.doesNotMatch(markdownFileName, /[<>:"/\\|?*]/);
evaluate('state.sourceFileName = "report\\u202Egnp.exe.csv"');
const bidiSafeFileName = evaluate('buildDataExportFileName("filtered", new Date(2026, 6, 17))');
assert.doesNotMatch(bidiSafeFileName, /[\u202A-\u202E\u2066-\u2069]/);
evaluate('state.sourceFileName = "销售 数据.xlsx"');

assert.equal(evaluate("canExportReport()"), true);
evaluate("updateExportButtons()");
assert.equal(elements.get("exportHtmlReport").disabled, false);
assert.equal(elements.get("exportMarkdownReport").disabled, false);

evaluate("state.analysisCompletedAt = null");
assert.equal(evaluate("canExportReport()"), false);
evaluate("updateExportButtons()");
assert.equal(elements.get("exportHtmlReport").disabled, true);
assert.equal(elements.get("exportMarkdownReport").disabled, true);
evaluate("state.analysisCompletedAt = new Date(2026, 6, 17, 14, 30, 5); state.rows = []");
assert.equal(evaluate("canExportReport()"), false);

// A date/category-only file still gets a date summary even though descriptive
// numeric statistics and metric-based charts are unavailable.
context.__noNumericRows = [
  { 日期: "2026-01-01", 说明: "开始" },
  { 日期: "2026-01-02", 说明: "继续" }
];
evaluate(`(() => {
  state.rows = __noNumericRows;
  state.fields = ["日期", "说明"];
  state.profiles = state.fields.map((field) => buildColumnProfile(field, state.rows));
  state.numericStats = buildNumericStats();
  state.categoryStats = buildCategoryStats();
  state.totalMissing = 0;
  state.duplicateRows = 0;
  state.parseWarnings = [];
  state.sourceFileName = "无数值字段.csv";
  state.sourceSheetName = "";
  state.sourceType = "csv";
  state.analysisCompletedAt = new Date(2026, 6, 17, 15, 0, 0);
  state.insights = ["未识别到可用于统计的数值字段。"];
  state.customAnalysis = null;
})()`);
const noNumericReport = JSON.parse(evaluate(
  'JSON.stringify(buildReportData(new Date(2026, 6, 17, 15, 0, 0)))'
));
assert.equal(noNumericReport.numericStats.length, 0);
assert.equal(noNumericReport.dateTrends.length, 1);
assert.deepEqual(noNumericReport.dateTrends[0].labels, ["2026-01-01", "2026-01-02"]);
context.__noNumericReport = noNumericReport;
assert.doesNotThrow(() => evaluate("buildHtmlReport(__noNumericReport, [])"));
assert.doesNotThrow(() => evaluate("buildMarkdownReport(__noNumericReport)"));

// Report renderers must keep working when no optional analytical section has
// data. This covers category-only files, empty trends and no custom analysis.
const emptySectionsReport = {
  ...reportData,
  dataScale: { rows: 1, fields: 1 },
  fieldTypes: [{
    field: "说明",
    typeKey: "category",
    type: "分类",
    nonMissingCount: 1,
    uniqueCount: 1,
    sampleValues: ["仅文本"]
  }],
  quality: {
    ...reportData.quality,
    totalMissing: 0,
    duplicateRows: 0,
    parseWarningCount: 0,
    parseWarnings: [],
    missingByField: [],
    outliers: []
  },
  numericStats: [],
  categoryStats: [],
  dateTrends: [],
  customAnalysis: { completed: false, message: "当前没有可导出的自定义分组分析结果。" },
  insights: ["未识别到可用于统计的数值字段。"]
};
context.__emptySectionsReport = emptySectionsReport;
assert.doesNotThrow(() => evaluate("buildHtmlReport(__emptySectionsReport, [])"));
assert.doesNotThrow(() => evaluate("buildMarkdownReport(__emptySectionsReport)"));
assert.match(evaluate("buildHtmlReport(__emptySectionsReport, [])"), /未识别到数值字段|无可用数值统计/);
assert.match(evaluate("buildMarkdownReport(__emptySectionsReport)"), /未识别到数值字段|无可用数值统计/);
assert.doesNotThrow(() => evaluate('buildHtmlReport({fileName:"空数据.csv", dataScale:{rows:0, fields:0}}, [])'));
assert.doesNotThrow(() => evaluate('buildMarkdownReport({fileName:"空数据.csv", dataScale:{rows:0, fields:0}})'));

console.log("core tests passed");
