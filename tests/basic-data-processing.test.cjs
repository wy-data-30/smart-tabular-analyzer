const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");
const { createScriptContext } = require("./test-context.cjs");

// The real vendored SheetJS builds and reads workbooks at test time, so no
// binary fixture is committed and the production read options are exercised.
const XLSX = require(path.join(__dirname, "..", "vendor", "xlsx-0.20.3.full.min.js"));
const Papa = require(path.join(__dirname, "..", "vendor", "papaparse-5.4.1.min.js"));
const fs = require("node:fs");
const crypto = require("node:crypto");

// v3.2.0: index.html as the tests compare it, without the data-i18n* attributes that carry the
// English keys, so the Chinese text and markup read as before. Tests read index.html only this way.
const DATA_I18N_ATTRIBUTE = /\s+data-i18n[\w-]*="[^"]*"/g;
const stripI18nAttributes = (html) => html.replace(DATA_I18N_ATTRIBUTE, "");
function indexHtmlSource() {
  return stripI18nAttributes(fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8"));
}

function put(context, name, value) {
  context[name] = value;
}

function buildXlsxBuffer(sheets) {
  const workbook = XLSX.utils.book_new();
  sheets.forEach(({ name, rows, formats = {} }) => {
    const worksheet = XLSX.utils.aoa_to_sheet(rows, { cellDates: true });
    Object.entries(formats).forEach(([address, format]) => {
      worksheet[address].z = format;
    });
    XLSX.utils.book_append_sheet(workbook, worksheet, name);
  });
  return XLSX.write(workbook, { type: "array", bookType: "xlsx" });
}

function createExcelContext(options = {}) {
  const scope = createScriptContext(options);
  scope.context.XLSX = XLSX;
  scope.context.window.XLSX = XLSX;
  return scope;
}

function importXlsxSheet(context, evaluate, buffer, sheetName, { newImport = true } = {}) {
  put(context, "__xlsxBuffer", buffer);
  put(context, "__xlsxSheet", sheetName);
  put(context, "__newImport", newImport);
  evaluate(`(() => {
    const importId = __newImport ? beginImport() : state.activeImportId;
    importExcelSheet(readExcelSheet(__xlsxBuffer, __xlsxSheet), "formats.xlsx", __xlsxSheet, importId);
  })()`);
}

function tableCells(elements, id) {
  return [...elements.get(id).innerHTML.matchAll(/<td(?: class="[^"]*")?>([^<]*)<\/td>/g)].map((match) => match[1]);
}

function profileSummary(evaluate) {
  return JSON.parse(evaluate("JSON.stringify(state.profiles.map((profile) => [profile.field, profile.typeKey, profile.numericFormat]))"));
}

test("CSV encoding and parsed rows preserve Chinese headers and values", () => {
  const { context, evaluate } = createScriptContext();
  const csvText = "姓名,城市,成绩\n张三,上海,90";
  const utf8Bytes = Array.from(new TextEncoder().encode(csvText));

  assert.equal(
    evaluate(`decodeCsvBuffer(new Uint8Array(${JSON.stringify(utf8Bytes)}).buffer, "utf-8").text`),
    csvText
  );

  const papaStub = {
    parse(text, options) {
      if (options.header === false) {
        // The second, line-number pass: one synchronous step per line, like PapaParse.
        let cursor = 0;
        text.split("\n").forEach((line) => {
          cursor = Math.min(text.length, cursor + line.length + 1);
          options.step({ data: line.split(","), errors: [], meta: { cursor } });
        });
        return;
      }
      context.__papaOptions = {
        header: options.header,
        skipEmptyLines: options.skipEmptyLines,
        preview: options.preview,
        worker: options.worker,
        hasHeaderTransform: typeof options.transformHeader === "function"
      };
      const [headerLine, ...dataLines] = text.split(/\r?\n/);
      const fields = headerLine.split(",").map(options.transformHeader);
      const data = dataLines.filter(Boolean).map((line) => {
        const values = line.split(",");
        return Object.fromEntries(fields.map((field, index) => [field, values[index] ?? ""]));
      });
      options.complete({ errors: [], meta: { fields }, data });
    }
  };
  context.Papa = papaStub;
  context.window.Papa = papaStub;

  const importId = evaluate("beginImport()");
  put(context, "__importId", importId);
  put(context, "__csvText", `${csvText}\n`);
  evaluate('parseCsvText(__csvText, "中文成绩.csv", __importId)');

  assert.equal(context.__papaOptions.header, true);
  assert.equal(context.__papaOptions.skipEmptyLines, true);
  assert.equal(context.__papaOptions.preview, 100001);
  assert.equal(context.__papaOptions.worker, false);
  assert.equal(context.__papaOptions.hasHeaderTransform, true);
  assert.equal(evaluate("state.rows.length"), 1);
  assert.equal(evaluate('state.fields.join(",")'), "姓名,城市,成绩");
  assert.equal(evaluate('state.rows[0]["姓名"]'), "张三");
  assert.equal(evaluate('state.rows[0]["城市"]'), "上海");
});

test("short CSV rows treat cells under prototype-named headers as missing", () => {
  const { evaluate } = createScriptContext();
  const result = JSON.parse(evaluate(`JSON.stringify((() => {
    const headers = ["name", "constructor", "__proto__", "toString"];
    const keys = headers.map((header, index) => encodeCsvHeaderForParser(header, index));
    // PapaParse header mode builds plain {} rows and omits missing trailing cells.
    const fullRow = {};
    keys.forEach((key, index) => { fullRow[key] = "v" + index; });
    const shortRow = {};
    shortRow[keys[0]] = "short";
    const importId = beginImport();
    handleParsedData({
      errors: [{ type: "FieldMismatch", code: "TooFewFields", message: "Too few fields", row: 1 }],
      meta: { fields: keys },
      data: [fullRow, shortRow]
    }, "short-row.csv", importId);
    const prototypeHeaders = headers.slice(1);
    return {
      fields: state.fields,
      rowCount: state.rows.length,
      shortCellsMissing: prototypeHeaders.map((field) => isMissing(state.rows[1][field])),
      missingCounts: prototypeHeaders.map((field) => getFieldProfile(field).missingCount),
      suspiciousCells: state.rows.flatMap((row) => state.fields
        .filter((field) => typeof row[field] === "function" || String(row[field]) === "[object Object]"))
    };
  })())`));

  assert.deepEqual(result.fields, ["name", "constructor", "__proto__", "toString"]);
  assert.equal(result.rowCount, 2);
  assert.deepEqual(result.shortCellsMissing, [true, true, true]);
  assert.deepEqual(result.missingCounts, [1, 1, 1]);
  assert.deepEqual(result.suspiciousCells, []);
});

test("Excel worksheet arrays become rows with Chinese fields", () => {
  const { evaluate } = createScriptContext();
  const result = JSON.parse(evaluate(`JSON.stringify(buildExcelTabularData(
    [["", "", ""], ["订单编号", "地区", "销售额"], ["A-001", "华东", "1,200"]],
    [["", "", ""], ["订单编号", "地区", "销售额"], ["A-001", "华东", 1200]],
    "销售数据"
  ))`));

  assert.deepEqual(result.fields, ["订单编号", "地区", "销售额"]);
  assert.deepEqual(result.rows, [
    { 订单编号: "A-001", 地区: "华东", 销售额: "1200" }
  ]);

  const emptyMessage = evaluate(`(() => {
    try {
      buildExcelTabularData([], [], "空工作表");
      return "";
    } catch (error) {
      return error.message;
    }
  })()`);
  assert.match(emptyMessage, /为空/);
});

test("Excel percent-formatted numbers are analyzed like CSV percent text", () => {
  const excel = createExcelContext();
  const buffer = buildXlsxBuffer([{
    name: "S1",
    rows: [
      ["毛利率", "精确值", "金额", "下单日期"],
      [0.26, 1.23456789, 1234.5, new Date(Date.UTC(2026, 6, 17))],
      [0.26, 1200, 99.9, new Date(Date.UTC(2026, 6, 18))],
      [0.31, 3.5, 10, new Date(Date.UTC(2026, 6, 19))]
    ],
    formats: { A2: "0%", A3: "0%", A4: "0.00%", C2: "¥#,##0.00", C3: "¥#,##0.00", C4: "¥#,##0.00", D2: "yyyy-mm-dd", D3: "yyyy-mm-dd", D4: "yyyy-mm-dd" }
  }]);
  importXlsxSheet(excel.context, excel.evaluate, buffer, "S1");

  assert.deepEqual(profileSummary(excel.evaluate), [
    ["毛利率", "numeric", "percent"],
    ["精确值", "numeric", "number"],
    ["金额", "numeric", "number"],
    ["下单日期", "date", "number"]
  ]);
  assert.equal(excel.evaluate('state.rows[0]["毛利率"]'), "0.26");
  assert.equal(excel.evaluate('state.rows[0]["精确值"]'), "1.23456789");
  assert.equal(excel.evaluate('state.rows[0]["下单日期"]'), "2026-07-17");

  const csv = createScriptContext();
  put(csv.context, "__rows", ["26%", "26%", "31%"].map((value) => ({ 毛利率: value })));
  csv.evaluate('commitTabularData(["毛利率"], __rows, "rates.csv", beginImport())');

  // One stats row is 字段, 有效数量, 无法转换, 平均值, 中位数, 最大值, 最小值, 标准差.
  const excelRateCells = tableCells(excel.elements, "numericStatsTable").slice(0, 8);
  assert.deepEqual(excelRateCells, tableCells(csv.elements, "numericStatsTable"));
  assert.equal(excelRateCells[4], "26%");
});

test("Excel columns with only a minority of percent cells stay plain numbers", () => {
  const { context, evaluate } = createExcelContext();
  const buffer = buildXlsxBuffer([{
    name: "S1",
    rows: [["少数百分比", "一半百分比"], [0.2, 0.2], [0.3, 0.3], [0.4, 0.4], [0.5, 0.5]],
    formats: { A2: "0%", B2: "0%", B3: "0%", A4: '0"%"', A5: "0\\%" }
  }]);
  importXlsxSheet(context, evaluate, buffer, "S1");

  assert.deepEqual(profileSummary(evaluate), [
    ["少数百分比", "numeric", "number"],
    ["一半百分比", "numeric", "number"]
  ]);
});

test("Excel percent hints do not leak across worksheets or into a later CSV", () => {
  const { context, evaluate } = createExcelContext();
  const buffer = buildXlsxBuffer([
    { name: "A", rows: [["毛利率"], [0.26], [0.3]], formats: { A2: "0%", A3: "0%" } },
    { name: "B", rows: [["毛利率"], [0.26], [0.3]] }
  ]);

  importXlsxSheet(context, evaluate, buffer, "A");
  assert.deepEqual(profileSummary(evaluate), [["毛利率", "numeric", "percent"]]);

  importXlsxSheet(context, evaluate, buffer, "B", { newImport: false });
  assert.deepEqual(profileSummary(evaluate), [["毛利率", "numeric", "number"]]);

  importXlsxSheet(context, evaluate, buffer, "A");
  put(context, "__csvRows", [{ 毛利率: "0.26" }, { 毛利率: "0.3" }]);
  evaluate('commitTabularData(["毛利率"], __csvRows, "rates.csv", beginImport())');
  assert.deepEqual(profileSummary(evaluate), [["毛利率", "numeric", "number"]]);
  assert.equal(evaluate("Object.keys(state.numberFormatHints).length"), 0);

  importXlsxSheet(context, evaluate, buffer, "A");
  evaluate("resetAnalysisState()");
  assert.equal(evaluate("Object.keys(state.numberFormatHints).length"), 0);
});

test("Excel percent hints survive field configuration and restoring automatic types", () => {
  const { context, elements, evaluate } = createExcelContext();
  const buffer = buildXlsxBuffer([{
    name: "S1",
    rows: [["毛利率", "门店"], [0.26, "1"], [0.26, "2"], [0.31, "3"]],
    formats: { A2: "0%", A3: "0%", A4: "0%" }
  }]);
  importXlsxSheet(context, evaluate, buffer, "S1");

  evaluate('stageFieldTypeChange("门店", "category"); applyFieldConfiguration()');
  assert.equal(evaluate('getFieldProfile("毛利率").numericFormat'), "percent");
  assert.equal(tableCells(elements, "numericStatsTable")[4], "26%");

  evaluate("restoreAutomaticFieldTypes()");
  assert.equal(evaluate('getFieldProfile("毛利率").numericFormat'), "percent");
  assert.equal(evaluate('state.numberFormatHints["毛利率"]'), "percent");
  assert.equal(tableCells(elements, "numericStatsTable")[4], "26%");
});

// Builds one "毛利率" column below `leadingBlankRows` empty rows; formatOf(i)
// returns the number format for the i-th data row (0-based).
function sampledRateSheet(dataRowCount, formatOf, leadingBlankRows = 0) {
  const rows = [...Array.from({ length: leadingBlankRows }, () => [null]), ["毛利率"]];
  const formats = {};
  for (let index = 0; index < dataRowCount; index += 1) {
    rows.push([0.2 + (index % 10) / 100]);
    formats[`A${leadingBlankRows + index + 2}`] = formatOf(index);
  }
  return buildXlsxBuffer([{ name: "S1", rows, formats }]);
}

function importedRateFormat(buffer) {
  const { context, evaluate } = createExcelContext();
  importXlsxSheet(context, evaluate, buffer, "S1");
  return evaluate('getFieldProfile("毛利率").numericFormat');
}

test("Excel percent hints sample only the first 1,000 data rows", () => {
  // Cases from the request: both results also hold under a full scan.
  assert.equal(importedRateFormat(sampledRateSheet(1500, (index) => (index < 1000 ? "0%" : "General"))), "percent");
  assert.equal(importedRateFormat(sampledRateSheet(1500, (index) => (index < 1000 ? "General" : "0%"))), "number");

  // A full scan would give the opposite result for these two sheets.
  assert.equal(importedRateFormat(sampledRateSheet(2500, (index) => (index < 1000 ? "0%" : "General"))), "percent");
  assert.equal(importedRateFormat(sampledRateSheet(2500, (index) => (index < 1000 ? "General" : "0%"))), "number");
});

test("Excel percent sampling starts below the detected header row", () => {
  // Data rows 500-999 are percent (501 of the first 1,000), so only a window
  // that starts right below the header has a percent majority.
  const buffer = sampledRateSheet(2000, (index) => (index >= 499 && index < 1000 ? "0%" : "General"), 3);
  assert.equal(importedRateFormat(buffer), "percent");
});

// Records every Chart config so tests can read the labels and data the
// numeric distribution chart was drawn with.
function createChartContext() {
  const scope = createScriptContext();
  const configs = [];
  class FakeChart {
    constructor(canvas, config) { configs.push({ canvas, config }); }
    destroy() {}
  }
  scope.context.Chart = FakeChart;
  scope.context.window.Chart = FakeChart;
  scope.lastChart = (canvasId) => {
    const canvas = scope.context.document.getElementById(canvasId);
    const { config } = configs.filter((entry) => entry.canvas === canvas).pop();
    return { labels: [...config.data.labels], data: [...config.data.datasets[0].data] };
  };
  scope.lastDistribution = () => scope.lastChart("numericDistributionChart");
  return scope;
}

// The bundled examples: sample-orders.csv stays in examples/ (the "use sample data" button
// loads it, v2.11.0), the other six, sample-data.csv among them, live in tests/fixtures/examples/.
// Every reader goes through these two helpers, so a file in either directory is found and listed once.
const EXAMPLE_DIRS = [path.join(__dirname, "..", "examples"), path.join(__dirname, "fixtures", "examples")];

function listExampleCsvFiles() {
  return EXAMPLE_DIRS.flatMap((dir) => fs.readdirSync(dir)).filter((fileName) => /\.csv$/.test(fileName)).sort();
}

function readExampleCsv(fileName) {
  const dir = EXAMPLE_DIRS.find((candidate) => fs.existsSync(path.join(candidate, fileName)));
  assert.ok(dir, `example file not found: ${fileName}`);
  return fs.readFileSync(path.join(dir, fileName), "utf8");
}

function commitExampleCsv(scope, fileName) {
  const text = readExampleCsv(fileName);
  const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });
  put(scope.context, "__exampleFields", parsed.meta.fields);
  put(scope.context, "__exampleRows", parsed.data);
  scope.evaluate(`commitTabularData(__exampleFields, __exampleRows, ${JSON.stringify(fileName)}, beginImport())`);
}

function drawDistribution(scope, field) {
  put(scope.context, "__chartField", field);
  scope.evaluate("dom.numericChartField.value = __chartField; renderNumericDistributionChart()");
  return scope.lastDistribution();
}

test("the distribution chart counts small integer scores per value", () => {
  const scope = createChartContext();
  commitExampleCsv(scope, "sample-survey.csv");

  const chart = drawDistribution(scope, "满意度");
  assert.deepEqual(chart.labels, ["2", "3", "4", "5", "6", "7", "8", "9"]);
  assert.equal(chart.labels.some((label) => label.includes(" - ")), false);
  assert.equal(chart.data.reduce((total, count) => total + count, 0), scope.evaluate('getNumericValues("满意度").length'));
});

test("the distribution chart keeps ten intervals for continuous values", () => {
  const scope = createChartContext();
  commitExampleCsv(scope, "sample-data.csv");

  const chart = drawDistribution(scope, "销售额");
  const expected = JSON.parse(scope.evaluate('JSON.stringify(buildHistogram(getNumericValues("销售额"), 10, "销售额"))'));
  assert.equal(chart.labels.length, 10);
  assert.deepEqual(chart.labels, expected.map((bin) => bin.label));
  assert.deepEqual(chart.data, expected.map((bin) => bin.count));
});

test("the distribution chart re-decides per-value counts after filtering", () => {
  const scope = createChartContext();
  const rows = Array.from({ length: 30 }, (_, index) => ({ 分值: String(index + 1), 组: index < 5 ? "A" : "B" }));
  put(scope.context, "__rows", rows);
  scope.evaluate('commitTabularData(["分值", "组"], __rows, "scores.csv", beginImport())');

  assert.equal(drawDistribution(scope, "分值").labels.length, 10);

  scope.evaluate(`(() => {
    state.filters = { category: { field: "组", values: ["A"] }, numeric: null, date: null };
    state.filteredRows = filterRowsByConfig(state.rows, state.filters);
    state.filterSourceRows = state.rows;
    refreshFilteredAnalysis();
  })()`);
  assert.deepEqual(scope.lastDistribution(), { labels: ["1", "2", "3", "4", "5"], data: [1, 1, 1, 1, 1] });

  scope.evaluate("clearAllFilters()");
  const cleared = scope.lastDistribution();
  assert.equal(cleared.labels.length, 10);
  assert.equal(cleared.labels.every((label) => label.includes(" - ")), true);
});

// 15 amounts, 3 of them not numbers (12 / 15 = 80%, so the column is still
// inferred as numeric). Only the first three rows are 华东 and hold one bad value.
const BAD_AMOUNT_ROWS = ["100", "200", "N/A", "300", "待定", "400", "约500", "600", "700", "800", "900", "1000", "1100", "1200", "1300"]
  .map((amount, index) => ({ 金额: amount, 地区: index < 3 ? "华东" : "华南" }));

function commitRows(scope, fields, rows, fileName = "bad-values.csv") {
  put(scope.context, "__badFields", fields);
  put(scope.context, "__badRows", rows);
  scope.evaluate(`commitTabularData(__badFields, __badRows, ${JSON.stringify(fileName)}, beginImport())`);
}

function cellsWithClass(elements, id) {
  return [...elements.get(id).innerHTML.matchAll(/<td( class="[^"]*")?>([^<]*)<\/td>/g)]
    .map((match) => ({ warning: /warning/.test(match[1] || ""), text: match[2] }));
}

function qualityPill(elements, label) {
  return [...elements.get("qualitySummary").innerHTML.matchAll(/<span class="quality-pill ([^"]*)">([^<]*)<\/span>/g)]
    .map((match) => ({ tone: match[1].trim(), text: match[2] }))
    .find((pill) => pill.text.startsWith(label));
}

function qualityRow(elements, field) {
  const row = [...elements.get("qualityTable").innerHTML.matchAll(/<tr>([\s\S]*?)<\/tr>/g)]
    .map((match) => [...match[1].matchAll(/<td( class="[^"]*")?>([^<]*)<\/td>/g)].map((cell) => ({ warning: /warning/.test(cell[1] || ""), text: cell[2] })))
    .find((cells) => cells[0]?.text === field);
  return row;
}

test("unconvertible values are flagged on the quality page from the original data", () => {
  const scope = createScriptContext();
  commitRows(scope, ["金额", "地区"], BAD_AMOUNT_ROWS);

  assert.equal(scope.evaluate('getFieldProfile("金额").conversionFailureCount'), 3);
  assert.deepEqual(JSON.parse(scope.evaluate('JSON.stringify(getFieldProfile("金额").conversionFailureSamples)')), ["N/A", "待定", "约500"]);
  assert.deepEqual(qualityPill(scope.elements, "认不出的数字或日期"), { tone: "warning", text: "认不出的数字或日期：3 个值" });
  const amountRow = qualityRow(scope.elements, "金额");
  assert.equal(amountRow[amountRow.length - 1].warning, true);
  assert.equal(amountRow[amountRow.length - 1].text, "3（N/A、待定、约500）");
  assert.deepEqual(qualityRow(scope.elements, "地区").slice(-1)[0], { warning: false, text: "0" });

  const clean = createScriptContext();
  commitRows(clean, ["金额"], ["1", "2", "3"].map((amount) => ({ 金额: amount })), "clean.csv");
  assert.deepEqual(qualityPill(clean.elements, "认不出的数字或日期"), { tone: "success", text: "认不出的数字或日期：0 个值" });
});

test("unconvertible dates count as format problems with samples", () => {
  const scope = createScriptContext();
  const dates = ["下周一", "2024-13-45", ...Array.from({ length: 8 }, (_, index) => `2024-01-0${index + 1}`)];
  commitRows(scope, ["日期"], dates.map((value) => ({ 日期: value })), "dates.csv");

  assert.equal(scope.evaluate('getFieldProfile("日期").typeKey'), "date");
  assert.deepEqual(qualityPill(scope.elements, "认不出的数字或日期"), { tone: "warning", text: "认不出的数字或日期：2 个值" });
  assert.equal(qualityRow(scope.elements, "日期").slice(-1)[0].text, "2（下周一、2024-13-45）");
});

test("statistics count unconvertible values for the current rows only", () => {
  const scope = createScriptContext();
  commitRows(scope, ["金额", "地区"], BAD_AMOUNT_ROWS);

  assert.equal(scope.evaluate('state.numericStats.find((item) => item.field === "金额").invalidCount'), 3);
  assert.match(scope.elements.get("numericStatsTable").innerHTML, /<th(?: class="[^"]*")?>有效数量<\/th>\s*<th(?: class="[^"]*")?>认不出<\/th>/);
  assert.deepEqual(cellsWithClass(scope.elements, "numericStatsTable").slice(0, 3), [
    { warning: false, text: "金额" },
    { warning: false, text: "12" },
    { warning: true, text: "3" }
  ]);

  scope.evaluate(`(() => {
    state.filters = { category: { field: "地区", values: ["华东"] }, numeric: null, date: null };
    state.filteredRows = filterRowsByConfig(state.rows, state.filters);
    state.filterSourceRows = state.rows;
    refreshFilteredAnalysis();
  })()`);
  assert.equal(scope.evaluate('state.numericStats.find((item) => item.field === "金额").invalidCount'), 1);
  assert.deepEqual(cellsWithClass(scope.elements, "numericStatsTable").slice(1, 3), [
    { warning: false, text: "2" },
    { warning: true, text: "1" }
  ]);
  assert.deepEqual(qualityPill(scope.elements, "认不出的数字或日期"), { tone: "warning", text: "认不出的数字或日期：3 个值" });
  assert.equal(qualityRow(scope.elements, "金额").slice(-1)[0].text, "3（N/A、待定、约500）");
});

test("field configuration changes update unconvertible counts and samples", () => {
  const scope = createScriptContext();
  commitRows(scope, ["金额", "地区"], BAD_AMOUNT_ROWS);

  scope.evaluate('stageFieldTypeChange("金额", "category"); applyFieldConfiguration()');
  assert.equal(scope.evaluate('getFieldProfile("金额").conversionFailureCount'), 0);
  assert.deepEqual(JSON.parse(scope.evaluate('JSON.stringify(getFieldProfile("金额").conversionFailureSamples)')), []);
  assert.deepEqual(qualityPill(scope.elements, "认不出的数字或日期"), { tone: "success", text: "认不出的数字或日期：0 个值" });

  scope.evaluate('stageFieldTypeChange("金额", "numeric"); applyFieldConfiguration()');
  assert.equal(scope.evaluate('getFieldProfile("金额").conversionFailureCount'), 3);
  assert.deepEqual(JSON.parse(scope.evaluate('JSON.stringify(getFieldProfile("金额").conversionFailureSamples)')), ["N/A", "待定", "约500"]);
  assert.deepEqual(qualityPill(scope.elements, "认不出的数字或日期"), { tone: "warning", text: "认不出的数字或日期：3 个值" });

  scope.evaluate('stageFieldTypeChange("金额", "category"); applyFieldConfiguration(); restoreAutomaticFieldTypes()');
  assert.equal(scope.evaluate('getFieldProfile("金额").conversionFailureCount'), 3);
  assert.deepEqual(JSON.parse(scope.evaluate('JSON.stringify(getFieldProfile("金额").conversionFailureSamples)')), ["N/A", "待定", "约500"]);
});

test("custom analysis summary names excluded unconvertible metric values", () => {
  const run = (rows) => {
    const scope = createScriptContext();
    commitRows(scope, ["金额", "地区"], rows);
    scope.evaluate('dom.v2MetricField.value = "金额"; dom.v2GroupField.value = "地区"; dom.v2AggregateMethod.value = "sum"; dom.v2DateField.value = ""; renderV2Analysis()');
    return scope;
  };

  const bad = run(BAD_AMOUNT_ROWS);
  assert.match(bad.elements.get("v2ResultSummary").textContent, /指标字段「金额」有 3 个值认不出是数字，已从计算中排除。$/);
  assert.equal(bad.evaluate("state.customAnalysis.metricInvalidCount"), 3);

  bad.evaluate(`(() => {
    state.filters = { category: { field: "地区", values: ["华东"] }, numeric: null, date: null };
    state.filteredRows = filterRowsByConfig(state.rows, state.filters);
    state.filterSourceRows = state.rows;
    refreshFilteredAnalysis();
  })()`);
  assert.match(bad.elements.get("v2ResultSummary").textContent, /有 1 个值认不出是数字/);

  const clean = run(BAD_AMOUNT_ROWS.map((row, index) => ({ ...row, 金额: String((index + 1) * 100) })));
  assert.doesNotMatch(clean.elements.get("v2ResultSummary").textContent, /认不出/);
  assert.equal(clean.evaluate("state.customAnalysis.metricInvalidCount"), 0);
});

test("reports carry unconvertible counts and escaped samples", () => {
  const scope = createScriptContext();
  const rows = BAD_AMOUNT_ROWS.map((row, index) => (index === 2 ? { ...row, 金额: "<b>x</b>|y" } : row));
  commitRows(scope, ["金额", "地区"], rows);
  scope.evaluate('dom.v2MetricField.value = "金额"; dom.v2GroupField.value = "地区"; dom.v2AggregateMethod.value = "sum"; dom.v2DateField.value = ""; renderV2Analysis()');

  const report = JSON.parse(scope.evaluate('JSON.stringify(buildReportData(new Date("2026-10-25T00:00:00.000Z")))'));
  assert.equal(report.numericStats.find((item) => item.field === "金额").invalidCount, 3);
  assert.equal(report.customAnalysis.metricInvalidCount, 3);
  assert.deepEqual(report.fieldTypes.find((item) => item.field === "金额").conversionFailureSamples, ["<b>x</b>|y", "待定", "约500"]);
  assert.deepEqual(report.fieldTypes.find((item) => item.field === "地区").conversionFailureSamples, []);

  put(scope.context, "__report", report);
  const html = scope.evaluate("buildHtmlReport(__report, [])");
  assert.match(html.split("数值字段描述性统计")[1], /<th>有效数量<\/th><th>认不出<\/th>/);
  // The whole header row, in order (the Markdown report pinned it until v3.2.0).
  assert.match(html.split("数值字段描述性统计")[1], /<th>字段<\/th><th>有效数量<\/th><th>认不出<\/th><th>平均值<\/th><th>中位数<\/th><th>最小值<\/th><th>最大值<\/th><th>标准差<\/th>/);
  assert.match(html.split("数值字段描述性统计")[1], /<td>金额<\/td><td>12<\/td><td>3<\/td>/);
  assert.match(html, /<td>3（&lt;b&gt;x&lt;\/b&gt;\|y、待定、约500）<\/td>/);
  assert.doesNotMatch(html, /<b>x<\/b>/);


  const legacy = {
    numericStats: [{ field: "金额", count: 2, mean: "1", median: "1", min: "1", max: "1", standardDeviation: "0" }],
    fieldTypes: [{ field: "金额", inferredType: "数值", appliedType: "数值", conversionFailureCount: 1 }],
    customAnalysis: { completed: true, metricField: "金额", groupField: "地区", rows: [] }
  };
  put(scope.context, "__legacy", legacy);
  const normalized = JSON.parse(scope.evaluate("JSON.stringify(normalizeReportSnapshot(__legacy))"));
  assert.equal(normalized.numericStats[0].invalidCount, 0);
  assert.deepEqual(normalized.fieldTypes[0].conversionFailureSamples, []);
  assert.equal(normalized.customAnalysis.metricInvalidCount, 0);
  assert.doesNotThrow(() => scope.evaluate("buildHtmlReport(__legacy, [])"));
});

// sample-data 订单日期 spans Jan-Jun 2025, so trends are grouped by month.
const SAMPLE_MONTHS = ["2025-01", "2025-02", "2025-03", "2025-04", "2025-05", "2025-06"];
// buildDateTrend("订单日期", "销售额") values from the HEAD before this change.
const SAMPLE_SALES_MONTHLY_AVERAGE = [2400, 2665.0000000000005, 4033.333333333333, 3167.5, 4110.833333333334, 11142.5];

function drawDateTrend(scope, { metric = "", aggregate = "avg" } = {}) {
  put(scope.context, "__trendMetric", metric);
  put(scope.context, "__trendAggregate", aggregate);
  scope.evaluate('dom.dateChartField.value = "订单日期"; dom.dateChartMetric.value = __trendMetric; dom.dateChartAggregate.value = __trendAggregate; renderDateTrendChart()');
  return { ...scope.lastChart("dateTrendChart"), title: scope.elements.get("dateChartTitle").textContent };
}

test("the date trend chart defaults to record counts per period", () => {
  const scope = createChartContext();
  commitExampleCsv(scope, "sample-data.csv");

  assert.equal(scope.evaluate("dom.dateChartMetric.value"), "");
  const chart = drawDateTrend(scope);
  assert.equal(chart.title, "订单日期 按月份趋势（记录数）");
  assert.deepEqual(chart.labels, SAMPLE_MONTHS);
  assert.deepEqual(chart.data, [6, 6, 6, 6, 6, 8]);
  assert.equal(scope.elements.get("dateChartAggregate").disabled, true);
});

test("the date trend chart averages or sums the selected metric", () => {
  const scope = createChartContext();
  commitExampleCsv(scope, "sample-data.csv");

  const average = drawDateTrend(scope, { metric: "销售额", aggregate: "avg" });
  assert.equal(average.title, "销售额 按月份趋势（平均值）");
  assert.deepEqual(average.labels, SAMPLE_MONTHS);
  assert.deepEqual(average.data, SAMPLE_SALES_MONTHLY_AVERAGE);
  assert.equal(scope.elements.get("dateChartAggregate").disabled, false);

  const total = drawDateTrend(scope, { metric: "销售额", aggregate: "sum" });
  assert.equal(total.title, "销售额 按月份趋势（求和）");
  assert.deepEqual(total.labels, SAMPLE_MONTHS);
  assert.equal(total.data[0], 14400);
  assert.equal(total.data[5], 89140);
});

test("date trend selections follow the import, filter and field-configuration paths", () => {
  const scope = createChartContext();
  commitExampleCsv(scope, "sample-data.csv");
  drawDateTrend(scope, { metric: "销售额", aggregate: "sum" });

  scope.evaluate(`(() => {
    state.filters = { category: { field: "地区", values: ["华东"] }, numeric: null, date: null };
    state.filteredRows = filterRowsByConfig(state.rows, state.filters);
    state.filterSourceRows = state.rows;
    refreshFilteredAnalysis();
  })()`);
  assert.equal(scope.evaluate("dom.dateChartMetric.value"), "销售额");
  assert.equal(scope.evaluate("dom.dateChartAggregate.value"), "sum");
  assert.equal(scope.elements.get("dateChartTitle").textContent, "销售额 按月份趋势（求和）");

  scope.evaluate('stageFieldTypeChange("数量", "category"); applyFieldConfiguration()');
  assert.equal(scope.evaluate("dom.dateChartMetric.value"), "销售额");
  assert.equal(scope.evaluate("dom.dateChartAggregate.value"), "sum");

  scope.evaluate('stageFieldTypeChange("销售额", "category"); applyFieldConfiguration()');
  assert.equal(scope.evaluate("dom.dateChartMetric.value"), "");
  assert.match(scope.elements.get("dateChartTitle").textContent, /（记录数）$/);

  scope.evaluate('dom.dateChartMetric.value = "成本"; dom.dateChartAggregate.value = "sum"');
  commitExampleCsv(scope, "sample-data.csv");
  assert.equal(scope.evaluate("dom.dateChartMetric.value"), "");
  assert.equal(scope.evaluate("dom.dateChartAggregate.value"), "avg");
});

test("report date trends follow the trend chart selection", () => {
  const scope = createChartContext();
  commitExampleCsv(scope, "sample-data.csv");
  const report = () => JSON.parse(scope.evaluate("JSON.stringify(buildReportDateTrends())"))[0];

  const counts = report();
  assert.equal(counts.metricField, "");
  assert.deepEqual(counts.values, [6, 6, 6, 6, 6, 8]);
  assert.match(counts.summary, /记录数从 2025-01 的 6 变为 2025-06 的 8/);

  scope.evaluate('dom.dateChartMetric.value = "销售额"; dom.dateChartAggregate.value = "avg"');
  const average = report();
  assert.equal(average.metricField, "销售额");
  assert.deepEqual(average.values, SAMPLE_SALES_MONTHLY_AVERAGE);
  assert.match(average.summary, /销售额平均值从 2025-01 的 2,400 变为 2025-06 的 11,142.5/);

  scope.evaluate('dom.dateChartAggregate.value = "sum"');
  const total = report();
  assert.equal(total.values[0], 14400);
  assert.match(total.summary, /销售额求和从 2025-01 的 14,400 变为 2025-06 的 89,140/);
});

function trendInsightFor(rows, dateField = "日期") {
  const scope = createScriptContext();
  commitRows(scope, Object.keys(rows[0]), rows, "trend.csv");
  put(scope.context, "__trendDateField", dateField);
  return { scope, text: scope.evaluate("buildTrendInsight(__trendDateField)") };
}

const datedRows = (dates) => dates.map((date) => ({ 日期: date }));
const PARTIAL_NOTE = /注意：最后一期/;

test("the trend insight compares record counts and flags a partial last month", () => {
  const scope = createChartContext();
  commitExampleCsv(scope, "sample-data.csv");
  assert.equal(
    scope.evaluate('buildTrendInsight("订单日期")'),
    "基于日期字段「订单日期」观察，记录数趋势 从 2025-01 到 2025-06 呈上升变化，变化幅度约 33.3%。注意：最后一期（2025-06）的数据截至 6 月 28 日，可能不完整，比较结果仅供参考。"
  );
});

test("the partial-period note follows month ends, leap days and day buckets", () => {
  const months = ["2025-01-05", "2025-02-05", "2025-03-05", "2025-04-05", "2025-05-05"];
  assert.doesNotMatch(trendInsightFor(datedRows([...months, "2025-06-30"])).text, PARTIAL_NOTE);
  assert.match(trendInsightFor(datedRows([...months, "2025-06-29"])).text, /注意：最后一期（2025-06）的数据截至 6 月 29 日/);

  const leap = ["2023-10-01", "2023-11-01", "2023-12-01", "2024-01-01"];
  assert.match(trendInsightFor(datedRows([...leap, "2024-02-28"])).text, /注意：最后一期（2024-02）的数据截至 2 月 28 日/);
  assert.doesNotMatch(trendInsightFor(datedRows([...leap, "2024-02-29"])).text, PARTIAL_NOTE);

  // A 19-day span is bucketed by day, so no note even though June is incomplete.
  const daily = trendInsightFor(datedRows(["2025-06-01", "2025-06-02", "2025-06-02", "2025-06-20"]));
  assert.match(daily.text, /从 2025-06-01 到 2025-06-20/);
  assert.doesNotMatch(daily.text, PARTIAL_NOTE);
});

test("the partial-period note handles year fields and year-only values", () => {
  const years = (last) => ["2021", "2022", "2023", "2024", last].map((value) => ({ 年份: value }));
  const partial = trendInsightFor(years("2025-06-30"), "年份");
  assert.equal(partial.scope.evaluate('getDateGranularity("年份")'), "year");
  assert.match(partial.text, /注意：最后一期（2025 年）的数据截至 6 月 30 日，可能不完整，比较结果仅供参考。$/);
  assert.doesNotMatch(trendInsightFor(years("2025-12-31"), "年份").text, PARTIAL_NOTE);
  assert.doesNotMatch(trendInsightFor(years("2025"), "年份").text, PARTIAL_NOTE);
});

test("the partial-period note follows the filtered rows", () => {
  const rows = [
    ["2025-01-10", "A"], ["2025-02-10", "A"], ["2025-03-10", "A"], ["2025-04-10", "A"], ["2025-05-31", "A"],
    ["2025-01-12", "C"], ["2025-03-12", "C"], ["2025-05-20", "C"],
    ["2025-01-15", "B"], ["2025-06-15", "B"]
  ].map(([date, region]) => ({ 日期: date, 地区: region }));
  const { scope, text } = trendInsightFor(rows);
  assert.match(text, /注意：最后一期（2025-06）的数据截至 6 月 15 日/);

  const filtered = (region) => {
    put(scope.context, "__trendRegion", region);
    scope.evaluate(`(() => {
      state.filters = { category: { field: "地区", values: [__trendRegion] }, numeric: null, date: null };
      state.filteredRows = filterRowsByConfig(state.rows, state.filters);
      state.filterSourceRows = state.rows;
      refreshFilteredAnalysis();
    })()`);
    return scope.evaluate('buildTrendInsight("日期")');
  };
  assert.doesNotMatch(filtered("A"), PARTIAL_NOTE);
  assert.match(filtered("C"), /注意：最后一期（2025-05）的数据截至 5 月 20 日/);
  assert.equal(scope.evaluate("state.chartInsights.trend"), scope.evaluate('buildTrendInsight("日期")'));
});

test("a flat record-count trend reads 基本持平 without a change ratio", () => {
  const scope = createChartContext();
  commitExampleCsv(scope, "sample-used-products.csv");
  const usedProducts = scope.evaluate('buildTrendInsight("上架日期")');
  assert.equal(usedProducts, "基于日期字段「上架日期」观察，记录数趋势 从 2025-04-01 到 2025-07-27 基本持平。");

  const monthly = trendInsightFor(datedRows(["2025-01-05", "2025-02-05", "2025-03-05", "2025-04-05", "2025-05-05", "2025-06-10"])).text;
  assert.equal(monthly, "基于日期字段「日期」观察，记录数趋势 从 2025-01 到 2025-06 基本持平。注意：最后一期（2025-06）的数据截至 6 月 10 日，可能不完整，比较结果仅供参考。");
  for (const text of [usedProducts, monthly]) {
    assert.doesNotMatch(text, /呈基本持平变化/);
    assert.doesNotMatch(text, /变化幅度/);
  }
});

test("the trend insight ignores numeric fields and the trend chart's metric selects", () => {
  const dates = ["2025-01-05", "2025-02-05", "2025-03-05", "2025-04-05", "2025-06-10"];
  const withAmounts = (amounts) => dates.map((date, index) => ({ 日期: date, 金额: String(amounts[index]) }));
  const { scope, text } = trendInsightFor(withAmounts([1, 2, 3, 4, 5]));
  assert.match(text, /记录数趋势/);
  assert.equal(trendInsightFor(withAmounts([500, 1, 900, 2, 0])).text, text);
  assert.equal(trendInsightFor(datedRows(dates)).text, text);

  scope.evaluate('dom.dateChartMetric.value = "金额"; dom.dateChartAggregate.value = "sum"; renderCharts()');
  assert.equal(scope.evaluate('buildTrendInsight("日期")'), text);
});

function categoryInsight(scope) {
  return JSON.parse(scope.evaluate('JSON.stringify(state.chartInsights.category.startsWith("分类字段「") ? state.chartInsights.category : null)'));
}

test("the category conclusion skips near-unique fields in sample-students", () => {
  const scope = createChartContext();
  commitExampleCsv(scope, "sample-students.csv");
  assert.equal(categoryInsight(scope), "分类字段「班级」中占比最高的类别是「二班」，出现 25 次，占比 21.0%。");
  assert.equal(scope.evaluate('state.categoryStats.find((item) => item.field === "姓名").distinctCount'), 40);
});

test("the category conclusion of sample-data is unchanged", () => {
  const scope = createChartContext();
  commitExampleCsv(scope, "sample-data.csv");
  assert.equal(categoryInsight(scope), "分类字段「地区」中占比最高的类别是「华东」，出现 8 次，占比 21.6%。");
});

test("the category conclusion accepts 20 distinct values but not 21", () => {
  const scope = createScriptContext();
  // 22 rows: 甲 has 21 distinct values (v0 twice), 乙 has 20 (w0 and w19 twice).
  const a = ["v0", ...Array.from({ length: 21 }, (_, index) => `v${index}`)];
  const b = ["w0", ...Array.from({ length: 19 }, (_, index) => `w${index}`), "w19", "w19"];
  const rows = a.map((value, index) => ({ 甲: value, 乙: b[index] }));
  commitRows(scope, ["甲", "乙"], rows, "distinct.csv");
  assert.equal(scope.evaluate('state.categoryStats.find((item) => item.field === "甲").distinctCount'), 21);
  assert.equal(scope.evaluate('state.categoryStats.find((item) => item.field === "乙").distinctCount'), 20);
  assert.equal(categoryInsight(scope), "分类字段「乙」中占比最高的类别是「w0」，出现 2 次，占比 9.1%。");
});

test("the category conclusion needs a top category that appears at least twice", () => {
  const scope = createScriptContext();
  const rows = Array.from({ length: 15 }, (_, index) => ({ 编码值: `K${index + 1}`, 渠道: index % 3 ? "线上" : "门店" }));
  commitRows(scope, ["编码值", "渠道"], rows, "unique.csv");
  // 15 distinct values out of 15 is inferred as ID; make it a category on purpose.
  scope.evaluate('stageFieldTypeChange("编码值", "category"); applyFieldConfiguration()');
  assert.equal(scope.evaluate('state.categoryStats.find((item) => item.field === "编码值").distinctCount'), 15);
  assert.equal(categoryInsight(scope), "分类字段「渠道」中占比最高的类别是「线上」，出现 10 次，占比 66.7%。");
});

test("the category conclusion says so for a scattered field and is hidden without a category field", () => {
  const scattered = createScriptContext();
  const rows = Array.from({ length: 30 }, (_, index) => ({ 甲: `a${index % 25}`, 乙: `b${index % 26}` }));
  commitRows(scattered, ["甲", "乙"], rows, "scattered.csv");
  assert.equal(categoryInsight(scattered), null);
  assert.equal(scattered.evaluate("state.chartInsights.category"), "「甲」取值比较分散（超过 20 个不同值，或没有重复出现的值），不做概括。");

  const noCategory = createScriptContext();
  commitRows(noCategory, ["日期", "金额"], [{ 日期: "2025-01-01", 金额: "1" }, { 日期: "2025-01-02", 金额: "2" }], "plain.csv");
  assert.equal(categoryInsight(noCategory), null);
  assert.equal(noCategory.evaluate("state.chartInsights.category"), "");
  assert.equal(noCategory.elements.get("categoryChartInsight").classList.contains("hidden"), true);
});

test("the category conclusion follows the selected field and is re-read from the filtered rows", () => {
  const scope = createScriptContext();
  const rows = Array.from({ length: 40 }, (_, index) => ({
    款号: index < 20 ? `g${Math.floor(index / 2)}` : `g${index - 10}`,
    区域: index < 20 ? "X" : "Y"
  }));
  commitRows(scope, ["款号", "区域"], rows, "filtered.csv");
  assert.equal(scope.evaluate('state.categoryStats.find((item) => item.field === "款号").distinctCount'), 30);
  assert.equal(categoryInsight(scope), "分类字段「区域」中占比最高的类别是「X」，出现 20 次，占比 50.0%。");

  scope.evaluate(`(() => {
    state.filters = { category: { field: "区域", values: ["X"] }, numeric: null, date: null };
    state.filteredRows = filterRowsByConfig(state.rows, state.filters);
    state.filterSourceRows = state.rows;
    refreshFilteredAnalysis();
  })()`);
  assert.equal(scope.evaluate('state.categoryStats.find((item) => item.field === "款号").distinctCount'), 10);
  // The chart keeps its field (区域) through the filter, and so does the conclusion; picking 款号 re-reads it from the filtered rows.
  assert.equal(categoryInsight(scope), "分类字段「区域」中占比最高的类别是「X」，出现 20 次，占比 100.0%。");
  scope.evaluate('dom.categoryChartField.value = "款号"; renderCategoryInsight()');
  assert.equal(categoryInsight(scope), "分类字段「款号」中占比最高的类别是「g0」，出现 2 次，占比 10.0%。");

  const report = JSON.parse(scope.evaluate('JSON.stringify(buildReportData(new Date("2026-10-25T00:00:00.000Z")))'));
  assert.deepEqual(report.categoryStats.map((item) => [item.field, item.distinctCount]), [["款号", 10], ["区域", 1]]);
  assert.deepEqual(report.insights, [scope.evaluate("state.chartInsights.category")]);
});

const NO_FILTER_SENTENCE = "未设置筛选，结果基于完整数据。";
const QUALITY_SCOPE_SENTENCE = "数据质量部分始终基于原始数据，不受筛选影响。";

function applyFilterState(scope, filters) {
  put(scope.context, "__appliedFilters", filters);
  scope.evaluate(`(() => {
    state.filters = { category: null, numeric: null, date: null, ...__appliedFilters };
    state.filteredRows = filterRowsByConfig(state.rows, state.filters);
    state.filterSourceRows = state.rows;
    refreshFilteredAnalysis();
  })()`);
}

// A literal text inside a RegExp (v3.2.0: the HTML assertions that took over from the Markdown ones).
const regExpText = (text) => String(text).replace(/[.*+?^${}()|[\]\\/]/g, "\\$&");

// Every marker is in the text, each one after the one before (v2.12.0): a bare
// indexOf(a) < indexOf(b) also passes when a is missing.
function assertInOrder(text, markers) {
  const positions = markers.map((marker) => text.indexOf(marker));
  markers.forEach((marker, index) => assert.ok(positions[index] >= 0, `missing: ${marker}`));
  for (let index = 1; index < markers.length; index += 1) {
    assert.ok(positions[index] > positions[index - 1], `${markers[index - 1]} before ${markers[index]}`);
  }
}

function reportOf(scope) {
  const report = JSON.parse(scope.evaluate('JSON.stringify(buildReportData(new Date("2026-10-25T00:00:00.000Z")))'));
  put(scope.context, "__filterReport", report);
  return {
    report,
    html: scope.evaluate("buildHtmlReport(__filterReport, [])")
  };
}

test("reports say when no filter is applied", () => {
  const scope = createChartContext();
  commitExampleCsv(scope, "sample-data.csv");
  const { report, html } = reportOf(scope);
  assert.deepEqual(report.filters, []);
  for (const text of [html]) {
    assert.ok(text.includes("筛选条件"));
    assert.ok(text.includes(NO_FILTER_SENTENCE));
    assert.ok(text.includes(QUALITY_SCOPE_SENTENCE));
  }
  assert.ok(html.indexOf("<h2>数据规模</h2>") < html.indexOf("<h2>筛选条件</h2>"));
  assert.ok(html.indexOf("<h2>筛选条件</h2>") < html.indexOf("<h2>字段类型结果</h2>"));
});

test("reports list applied category, numeric and date filters in order", () => {
  const scope = createChartContext();
  commitExampleCsv(scope, "sample-data.csv");
  applyFilterState(scope, {
    category: { field: "地区", values: ["华东", "华北", "华南"] },
    numeric: { field: "销售额", min: 1000, max: 5000 },
    date: { field: "订单日期", start: "2025-01-01", end: "2025-03-31" }
  });
  const { report, html } = reportOf(scope);
  assert.deepEqual(report.filters, [
    { type: "category", field: "地区", keptCount: 3, totalCount: 6, keptValues: ["华北", "华东", "华南"], description: "地区：保留 3 / 共 6 个类别（华北、华东、华南）" },
    { type: "numeric", field: "销售额", min: 1000, max: 5000, description: "销售额：1,000 至 5,000" },
    { type: "date", field: "订单日期", start: "2025-01-01", end: "2025-03-31", description: "订单日期：2025-01-01 至 2025-03-31" }
  ]);
  const analysisRows = scope.evaluate("getAnalysisRows().length");
  assert.equal(report.dataScale.rows, analysisRows);
  assert.equal(report.dataScale.filteredOutRows, 38 - analysisRows);
  assert.ok(report.dataScale.filteredOutRows > 0);

  assert.ok(html.includes("<th>字段</th><th>条件</th>"));
  assert.ok(html.includes("<td>地区</td><td>地区：保留 3 / 共 6 个类别（华北、华东、华南）</td>"));
  assert.ok(html.includes("<td>销售额</td><td>销售额：1,000 至 5,000</td>"));
  assert.ok(html.includes("<td>订单日期</td><td>订单日期：2025-01-01 至 2025-03-31</td>"));
  assert.ok(!html.includes(NO_FILTER_SENTENCE));
  assert.ok(html.includes(QUALITY_SCOPE_SENTENCE));
});

test("report category filters show at most ten kept values and empty values", () => {
  const scope = createScriptContext();
  const codes = Array.from({ length: 15 }, (_, index) => `c${String(index + 1).padStart(2, "0")}`);
  commitRows(scope, ["类别"], [...codes, ...codes].map((code) => ({ 类别: code })), "many.csv");
  applyFilterState(scope, { category: { field: "类别", values: codes.slice(0, 12) } });
  const [many] = reportOf(scope).report.filters;
  assert.equal(many.keptCount, 12);
  assert.equal(many.totalCount, 15);
  assert.deepEqual(many.keptValues, codes.slice(0, 10));
  assert.equal(many.description, `类别：保留 12 / 共 15 个类别（${codes.slice(0, 10).join("、")}……等 12 个）`);

  const blanks = createScriptContext();
  commitRows(blanks, ["地区"], ["华东", "", "  ", "华南", "华东"].map((region) => ({ 地区: region })), "blank.csv");
  applyFilterState(blanks, { category: { field: "地区", values: ["", "华东"] } });
  const [blank] = reportOf(blanks).report.filters;
  assert.deepEqual(blank.keptValues, ["（空值）", "华东"]);
  assert.equal(blank.description, "地区：保留 2 / 共 3 个类别（（空值）、华东）");
});

test("report numeric filters on percent fields use the applied fraction range", () => {
  const scope = createScriptContext();
  commitRows(scope, ["转化率"], ["10%", "20%", "35%", "40%", "50%"].map((rate) => ({ 转化率: rate })), "rates.csv");
  applyFilterState(scope, { numeric: { field: "转化率", min: 0.2, max: 0.4 } });
  const { report } = reportOf(scope);
  assert.deepEqual(report.filters, [{ type: "numeric", field: "转化率", min: 0.2, max: 0.4, description: "转化率：20% 至 40%" }]);
  assert.deepEqual(JSON.parse(scope.evaluate('JSON.stringify(getAnalysisRows().map((row) => row["转化率"]))')), ["20%", "35%", "40%"]);
});

test("report filters ignore unapplied drafts and clear with the filters", () => {
  const scope = createScriptContext();
  commitRows(scope, ["地区", "金额"], [
    { 地区: "华东", 金额: "100" }, { 地区: "华南", 金额: "200" }, { 地区: "华北", 金额: "300" }
  ], "drafts.csv");
  scope.evaluate('dom.categoryFilterField.value = "地区"; updateCategoryFilterField(); state.categoryFilterDraftValues.delete("华北"); applyFilters()');
  const applied = reportOf(scope).report.filters;
  assert.equal(applied[0].description, "地区：保留 2 / 共 3 个类别（华东、华南）");

  scope.evaluate('state.categoryFilterDraftValues.delete("华南"); dom.numericFilterField.value = "金额"; updateNumericFilterField(); dom.numericFilterMin.value = "150"');
  assert.deepEqual(reportOf(scope).report.filters, applied);

  scope.evaluate("clearAllFilters()");
  assert.deepEqual(reportOf(scope).report.filters, []);
});

test("report filter descriptions are escaped in HTML", () => {
  const scope = createScriptContext();
  commitRows(scope, ["备注"], ["<b>x</b>|y", "第一行\n第二行", "其他"].map((note) => ({ 备注: note })), "escape.csv");
  applyFilterState(scope, { category: { field: "备注", values: ["<b>x</b>|y", "第一行\n第二行"] } });
  const { html } = reportOf(scope);
  // The same values also appear as field samples, so look only at the filter section.
  const htmlSection = html.split("<h2>筛选条件</h2>")[1]?.split("<h2>字段类型结果</h2>")[0] ?? "";
  assert.ok(htmlSection.includes("<td>备注</td><td>备注：保留 2 / 共 3 个类别（&lt;b&gt;x&lt;/b&gt;|y、第一行\n第二行）</td>"));
  assert.ok(!html.includes("<b>x</b>"));
});

test("legacy report snapshots without filters still render", () => {
  const scope = createScriptContext();
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(normalizeReportSnapshot({}).filters)")), []);
  assert.ok(scope.evaluate("buildHtmlReport({}, [])").includes(NO_FILTER_SENTENCE));
});

test("field inference recognizes ID, numeric, date and category columns", () => {
  const { context, evaluate } = createScriptContext();
  const rows = Array.from({ length: 12 }, (_, index) => ({
    学号: String(20260001 + index),
    成绩: String(70 + index),
    考试日期: `2026-07-${String(index + 1).padStart(2, "0")}`,
    班级: index % 2 === 0 ? "一班" : "二班"
  }));
  put(context, "__rows", rows);

  assert.equal(evaluate('buildColumnProfile("学号", __rows).typeKey'), "id");
  assert.equal(evaluate('buildColumnProfile("成绩", __rows).typeKey'), "numeric");
  assert.equal(evaluate('buildColumnProfile("考试日期", __rows).typeKey'), "date");
  assert.equal(evaluate('buildColumnProfile("班级", __rows).typeKey'), "category");
});

test("missing values, duplicate rows and IQR outliers are detected", () => {
  const { context, evaluate } = createScriptContext();
  const rows = [
    { 数值: "1", 类别: "A", 备注: "" },
    { 数值: "2", 类别: "A", 备注: null },
    { 数值: "3", 类别: "B", 备注: "正常" },
    { 数值: "4", 类别: "B", 备注: "正常" },
    { 数值: "100", 类别: "B", 备注: "异常值" },
    { 数值: "4", 类别: "B", 备注: "正常" }
  ];
  put(context, "__rows", rows);
  evaluate(`(() => {
    state.rows = __rows;
    state.fields = ["数值", "类别", "备注"];
    state.profiles = state.fields.map((field) => buildColumnProfile(field, state.rows));
    state.numericStats = buildNumericStats();
    state.categoryStats = buildCategoryStats();
  })()`);

  assert.equal(evaluate('state.profiles.find((item) => item.field === "备注").missingCount'), 2);
  assert.equal(evaluate('countDuplicateRows(state.rows, state.fields)'), 1);
  assert.equal(evaluate('state.numericStats[0].outlierCount'), 1);
  assert.equal(evaluate('state.numericStats[0].max'), 100);
});

test("quality outlier bounds near zero are displayed without negative zero", () => {
  const { context, elements, evaluate } = createScriptContext();
  const rows = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0.002, 0.004, 0.01].map((value) => ({ 比例: String(value) }));
  put(context, "__rows", rows);
  evaluate('commitTabularData(["比例"], __rows, "near-zero.csv", beginImport())');

  assert.equal(evaluate("state.qualityNumericStats[0].lowerBound"), -0.00075);
  const outlierCells = [...elements.get("outlierTable").innerHTML.matchAll(/<td(?: class="[^"]*")?>([^<]*)<\/td>/g)].map((match) => match[1]);
  assert.deepEqual(outlierCells, ["比例", "3", "-0.00075", "0.00125"]);
  assert.equal(outlierCells.includes("-0"), false);

  const outliers = JSON.parse(evaluate('JSON.stringify(buildReportData(new Date("2026-10-25T00:00:00.000Z")).quality.outliers)'));
  assert.deepEqual(outliers, [{ field: "比例", outlierCount: 3, lowerBound: "-0.00075", upperBound: "0.00125" }]);
});

test("descriptive statistics show small averages instead of rounding them to zero", () => {
  const { context, elements, evaluate } = createScriptContext();
  const rows = ["0.001", "0.0012", "0.0014"].map((value) => ({ 误差: value }));
  put(context, "__rows", rows);
  evaluate('commitTabularData(["误差"], __rows, "small-mean.csv", beginImport())');

  const statsCells = [...elements.get("numericStatsTable").innerHTML.matchAll(/<td(?: class="[^"]*")?>([^<]*)<\/td>/g)].map((match) => match[1]);
  // statsCells[2] is the 无法转换 count; the statistics start at index 3.
  assert.equal(statsCells[0], "误差");
  assert.equal(statsCells[2], "0");
  assert.equal(statsCells[3], "0.0012");
  assert.equal(statsCells.slice(3).includes("0"), false);
});

test("descriptive and category statistics return expected values", () => {
  const { context, evaluate } = createScriptContext();
  const rows = [
    { 分数: "10", 等级: "A" },
    { 分数: "20", 等级: "A" },
    { 分数: "30", 等级: "B" },
    { 分数: "40", 等级: "" }
  ];
  put(context, "__rows", rows);
  evaluate(`(() => {
    state.rows = __rows;
    state.fields = ["分数", "等级"];
    state.profiles = state.fields.map((field) => buildColumnProfile(field, state.rows));
    state.numericStats = buildNumericStats();
    state.categoryStats = buildCategoryStats();
  })()`);

  assert.equal(evaluate('state.numericStats[0].mean'), 25);
  assert.equal(evaluate('state.numericStats[0].median'), 25);
  assert.equal(evaluate('state.numericStats[0].min'), 10);
  assert.equal(evaluate('state.numericStats[0].max'), 40);
  assert.equal(evaluate('state.categoryStats[0].top[0].name'), "A");
  assert.equal(evaluate('state.categoryStats[0].top[0].count'), 2);
  assert.equal(evaluate('state.categoryStats[0].top[0].ratio'), 2 / 3);
});

test("empty, single-column and type-missing datasets remain safe", () => {
  const { context, elements, evaluate } = createScriptContext();

  let importId = evaluate("beginImport()");
  put(context, "__importId", importId);
  put(context, "__emptyResult", { errors: [], meta: { fields: [] }, data: [] });
  evaluate('handleParsedData(__emptyResult, "empty.csv", __importId)');
  assert.equal(evaluate("state.rows.length"), 0);
  assert.equal(elements.get("statusPanel").className, "status-panel error");

  importId = evaluate("beginImport()");
  put(context, "__importId", importId);
  put(context, "__singleColumnResult", {
    errors: [],
    meta: { fields: ["状态"] },
    data: [{ 状态: "正常" }, { 状态: "异常" }]
  });
  evaluate('handleParsedData(__singleColumnResult, "single-column.csv", __importId)');
  assert.equal(evaluate("state.fields.length"), 1);
  assert.equal(evaluate("state.rows.length"), 2);
  assert.equal(evaluate("state.numericStats.length"), 0);

  put(context, "__numericOnlyRows", [{ 数量: "1" }, { 数量: "2" }, { 数量: "3" }]);
  evaluate(`(() => {
    state.rows = __numericOnlyRows;
    state.fields = ["数量"];
    state.profiles = [buildColumnProfile("数量", state.rows)];
    state.numericStats = buildNumericStats();
    state.categoryStats = buildCategoryStats();
  })()`);
  assert.equal(evaluate("state.numericStats.length"), 1);
  assert.equal(evaluate("state.categoryStats.length"), 0);
});

test("category, numeric and date filters combine without mutating original rows", () => {
  const { context, evaluate } = createScriptContext();
  const rows = [
    { 地区: "华东", 销售额: "100", 日期: "2026-01-05" },
    { 地区: "华东", 销售额: "220", 日期: "2026-01-10" },
    { 地区: "华南", 销售额: "150", 日期: "2026-01-12" },
    { 地区: "华东", 销售额: "180", 日期: "2026-02-02" },
    { 地区: "华东", 销售额: "", 日期: "2026-01-20" }
  ];
  put(context, "__rows", rows);
  const originalSnapshot = JSON.stringify(rows);
  const filtered = JSON.parse(evaluate(`JSON.stringify((() => {
    state.rows = __rows;
    state.fields = ["地区", "销售额", "日期"];
    state.profiles = state.fields.map((field) => buildColumnProfile(field, state.rows));
    return filterRowsByConfig(state.rows, {
      category: {field: "地区", values: ["华东"]},
      numeric: {field: "销售额", min: 90, max: 200},
      date: {field: "日期", start: "2026-01-01", end: "2026-01-31"}
    });
  })())`));

  assert.deepEqual(filtered, [rows[0]]);
  assert.equal(JSON.stringify(rows), originalSnapshot);
  assert.equal(evaluate("state.rows.length"), 5);
});

test("filtered statistics change while original quality outliers remain stable", () => {
  const { context, elements, evaluate } = createScriptContext();
  const rows = [
    { 数值: "1" },
    { 数值: "2" },
    { 数值: "3" },
    { 数值: "4" },
    { 数值: "100" }
  ];
  put(context, "__rows", rows);
  evaluate(`(() => {
    state.rows = __rows;
    state.fields = ["数值"];
    state.profiles = [buildColumnProfile("数值", state.rows)];
    state.filteredRows = state.rows.slice(0, 4);
    state.filterSourceRows = state.rows;
    state.numericStats = buildNumericStats(getAnalysisRows());
    state.qualityNumericStats = buildNumericStats(state.rows);
    state.categoryStats = buildCategoryStats(getAnalysisRows());
    state.totalMissing = 0;
    state.duplicateRows = 0;
    state.analysisCompletedAt = new Date("2026-07-18T00:00:00.000Z");
  })()`);

  assert.equal(evaluate("state.rows.length"), 5);
  assert.equal(evaluate("getAnalysisRows().length"), 4);
  assert.equal(evaluate("state.numericStats[0].outlierCount"), 0);
  assert.equal(evaluate("getQualityNumericStats()[0].outlierCount"), 1);

  const report = JSON.parse(evaluate(
    'JSON.stringify(buildReportData(new Date("2026-07-18T01:00:00.000Z")))'
  ));
  assert.equal(report.dataScale.rows, 4);
  assert.equal(report.dataScale.originalRows, 5);
  assert.equal(report.dataScale.filteredOutRows, 1);
  assert.equal(report.quality.outliers[0].outlierCount, 1);

  evaluate('rebuildAnalysisFromProfiles("测试字段配置")');
  assert.equal(evaluate("getAnalysisRows().length"), 5);
  assert.match(elements.get("previewTable").innerHTML, />100</);
});

test("processed CSV exports preserve fields, Chinese text, BOM and original rows", async () => {
  const { context, elements, evaluate } = createScriptContext();
  const rows = [
    { 姓名: "张三", 备注: "上海,重点客户" },
    { 姓名: "张三", 备注: "上海,重点客户" },
    { 姓名: "李四", 备注: "包含\"引号\"" }
  ];
  put(context, "__rows", rows);
  const originalSnapshot = JSON.stringify(rows);

  context.window.Papa = {
    unparse(input, options) {
      context.__unparseInput = input;
      context.__unparseOptions = options;
      const escapeCell = (value) => {
        const text = String(value ?? "");
        return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
      };
      return [
        input.fields.map(escapeCell).join(","),
        ...input.data.map((row) => row.map(escapeCell).join(","))
      ].join(options.newline);
    }
  };

  evaluate(`(() => {
    state.rows = __rows;
    state.filteredRows = [state.rows[2]];
    state.filterSourceRows = state.rows;
    state.fields = ["姓名", "备注"];
    state.profiles = state.fields.map((field) => buildColumnProfile(field, state.rows));
    state.duplicateRows = countDuplicateRows(state.rows, state.fields);
    state.sourceFileName = "中文销售.xlsx";
  })()`);

  assert.equal(evaluate('getProcessedExportRows("filtered").length'), 1);
  assert.equal(JSON.stringify(rows), originalSnapshot);

  const csv = evaluate('buildCsvExportContent(getProcessedExportRows("filtered"), state.fields)');
  assert.equal(
    csv,
    '姓名,备注\r\n李四,"包含""引号"""'
  );
  assert.deepEqual(Array.from(context.__unparseInput.fields), ["姓名", "备注"]);
  assert.deepEqual(
    Array.from(context.__unparseInput.data, (row) => Array.from(row)),
    [["李四", '包含"引号"']]
  );
  assert.equal(context.__unparseOptions.newline, "\r\n");

  assert.equal(
    evaluate('buildDataExportFileName("filtered", new Date(2026, 6, 18))'),
    "中文销售_当前分析_20260718.csv"
  );

  const blob = evaluate('createUtf8BomBlob("姓名,城市\\r\\n张三,上海", "text/csv;charset=utf-8")');
  const bytes = new Uint8Array(await blob.arrayBuffer());
  assert.deepEqual(Array.from(bytes.slice(0, 3)), [0xef, 0xbb, 0xbf]);

  evaluate("updateExportButtons()");
  assert.equal(elements.get("dataExportOriginalCount").textContent, "3");
  assert.equal(elements.get("dataExportFilteredCount").textContent, "1");
  assert.equal(elements.get("exportFilteredCsv").disabled, false);

  evaluate("state.filteredRows = []");
  assert.equal(
    evaluate('buildCsvExportContent(getProcessedExportRows("filtered"), state.fields)'),
    "姓名,备注"
  );

  // v3.1.0: without a filter, every counted row in order, the repeated row kept (the
  // deduplicated export is gone): a comma in a cell, quotes, rows joined by CRLF.
  evaluate("state.filteredRows = null; state.filterSourceRows = null");
  assert.equal(evaluate('getProcessedExportRows("filtered").length'), 3);
  assert.equal(
    evaluate('buildCsvExportContent(getProcessedExportRows("filtered"), state.fields)'),
    '姓名,备注\r\n张三,"上海,重点客户"\r\n张三,"上海,重点客户"\r\n李四,"包含""引号"""'
  );
  assert.deepEqual(
    Array.from(context.__unparseInput.data, (row) => Array.from(row)),
    [["张三", "上海,重点客户"], ["张三", "上海,重点客户"], ["李四", '包含"引号"']]
  );
  assert.equal(JSON.stringify(rows), originalSnapshot);
});

// Stage 3 acceptance fixtures (tests/fixtures/stage3/, see known-issues.json)
// imported through parseDataFile() with the real PapaParse and SheetJS.
const STAGE3_FIXTURES = path.join(__dirname, "fixtures", "stage3");

async function importStage3Fixture(fileName, language) {
  const scope = createExcelContext(language === undefined ? {} : { language });
  let papaCalls = 0;
  const papaHeaderOptions = [];
  const countingPapa = Object.assign(Object.create(Papa), {
    parse(...args) {
      papaCalls += 1;
      papaHeaderOptions.push(args[1]?.header);
      return Papa.parse(...args);
    }
  });
  put(scope.context, "Papa", countingPapa);
  scope.context.window.Papa = countingPapa;
  const bytes = fs.readFileSync(path.join(STAGE3_FIXTURES, fileName));
  // The VM has no FileReader, so readFileAsArrayBuffer() uses file.arrayBuffer().
  put(scope.context, "__stage3File", {
    name: fileName,
    size: bytes.length,
    arrayBuffer: async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.length)
  });
  await scope.evaluate("parseDataFile(__stage3File)");
  return { ...scope, papaCalls: () => papaCalls, papaHeaderOptions: () => [...papaHeaderOptions] };
}

function importedSummary(evaluate) {
  return JSON.parse(evaluate(`JSON.stringify({
    rows: state.rows.length,
    fields: state.fields,
    types: state.profiles.map((profile) => profile.typeKey),
    duplicateRows: state.duplicateRows,
    totalMissing: state.totalMissing
  })`));
}

test("a header in the first row of the order fixture imports as before", async () => {
  const scope = await importStage3Fixture("订单表.csv");
  assert.deepEqual(importedSummary(scope.evaluate), {
    rows: 30,
    fields: ["订单号", "下单日期", "客户名称", "大区", "产品", "数量", "金额", "备注"],
    types: ["id", "date", "category", "category", "category", "numeric", "numeric", "category"],
    duplicateRows: 0,
    totalMissing: 26
  });
  assert.equal(scope.papaCalls(), 2);
  assert.equal(scope.papaHeaderOptions()[1], false);
});

test("a header in the first row of the roster fixture imports as before", async () => {
  const scope = await importStage3Fixture("花名册.xlsx");
  assert.deepEqual(importedSummary(scope.evaluate), {
    rows: 30,
    fields: ["工号", "姓名", "部门", "身份证号", "入职日期", "月薪", "手机号"],
    types: ["id", "category", "category", "id", "date", "numeric", "id"],
    duplicateRows: 1,
    totalMissing: 2
  });
});

async function importWorkbookFile(buffer, fileName = "t.xlsx", language) {
  const scope = createExcelContext(language === undefined ? {} : { language });
  const bytes = Buffer.from(buffer);
  put(scope.context, "__workbookFile", {
    name: fileName,
    size: bytes.length,
    arrayBuffer: async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.length)
  });
  await scope.evaluate("parseDataFile(__workbookFile)");
  return scope;
}

const STAGE3_ATTENDANCE_HEADER = ["姓名", "部门", "出勤天数", "迟到次数"];
const STAGE3_ATTENDANCE_ROWS = [["张三", "销售部", 22, 1], ["李四", "技术部", 21, 0], ["王五", "财务部", 20, 2]];

test("the store report imports with the header from row 4 below its title area", async () => {
  const { evaluate } = await importStage3Fixture("门店报表.xlsx");
  assert.deepEqual(JSON.parse(evaluate("JSON.stringify(state.fields)")),
    ["门店编号", "门店名称", "城市", "开业日期", "销售额", "成本", "毛利率", "店长"]);
  assert.equal(evaluate("state.rows.length"), 29);
  assert.equal(evaluate('getFieldProfile("毛利率").numericFormat'), "percent");
});

test("Excel imports skip a title row with a unit note and a school/date row", async () => {
  const unitNote = await importWorkbookFile(buildXlsxBuffer([{
    name: "S1", rows: [["考勤表", "单位：天"], STAGE3_ATTENDANCE_HEADER, ...STAGE3_ATTENDANCE_ROWS]
  }]));
  assert.equal(unitNote.evaluate('state.fields.join(",")'), STAGE3_ATTENDANCE_HEADER.join(","));
  assert.equal(unitNote.evaluate("state.rows.length"), 3);

  const schoolDate = await importWorkbookFile(buildXlsxBuffer([{
    name: "S1", rows: [["期末成绩"], ["学校：一中", "日期：2024-06"], STAGE3_ATTENDANCE_HEADER, ...STAGE3_ATTENDANCE_ROWS]
  }]));
  assert.equal(schoolDate.evaluate('state.fields.join(",")'), STAGE3_ATTENDANCE_HEADER.join(","));
  assert.equal(schoolDate.evaluate("state.rows.length"), 3);
});

test("Excel header row numbers count leading blank rows and the !ref start", () => {
  const { context, evaluate } = createExcelContext();
  const worksheet = XLSX.utils.aoa_to_sheet([["考勤表"], STAGE3_ATTENDANCE_HEADER, ...STAGE3_ATTENDANCE_ROWS], { origin: "B4" });
  worksheet["!ref"] = "B3:E8"; // row 3 is a blank row inside the used range
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "S1");
  put(context, "__offsetBuffer", XLSX.write(workbook, { type: "array", bookType: "xlsx" }));
  const table = JSON.parse(evaluate(`JSON.stringify((() => {
    return readExcelSheet(__offsetBuffer, "S1");
  })())`));
  assert.deepEqual(table.fields, STAGE3_ATTENDANCE_HEADER);
  assert.equal(table.rows.length, 3);
  assert.equal(table.headerDetection.headerRowNumber, 5);
  assert.deepEqual(table.headerDetection.skippedRows, [{ rowNumber: 4, cells: ["考勤表"] }]);
});

// Appended to a header error when the first non-empty row was a skippable title row.
const TITLE_FALLBACK_HINT = "第 1 行看起来像标题。如果表头不在第 1 行，请删除表头上方的标题行后重新上传。";

test("Excel detection falls back when a later data column has no header", async () => {
  // Column E is empty at the header and for the first 250 rows, so only the
  // full check after detection sees its data.
  const rows = [["考勤汇总"], [...STAGE3_ATTENDANCE_HEADER, ""]];
  for (let index = 1; index <= 300; index += 1) {
    rows.push([`员工${index}`, ["销售部", "技术部", "财务部"][index % 3], 20 + (index % 3), index % 4, index > 250 ? "补录" : ""]);
  }
  const scope = await importWorkbookFile(buildXlsxBuffer([{ name: "S1", rows }]));
  assert.equal(scope.evaluate("state.rows.length"), 0);
  assert.equal(scope.elements.get("statusPanel").textContent, `工作表「S1」的第 2 列缺少表头，请补充字段名后重新上传。${TITLE_FALLBACK_HINT}`);
});

test("Excel sheets whose first row is not a title keep their header errors", async () => {
  // A single-cell first row is skippable, so its error also carries the title hint.
  const cases = [
    [[["姓名"], ["学生1", "销售部"], ["学生2", "技术部"], ["学生3", "销售部"], ["学生4", "财务部"], ["学生5", "技术部"]], 2, TITLE_FALLBACK_HINT],
    [[["姓名", "部门"], ["张三", "销售部", new Date(Date.UTC(2024, 0, 5)), 90], ["李四", "技术部", new Date(Date.UTC(2024, 1, 6)), 85], ["王五", "财务部", new Date(Date.UTC(2024, 2, 7)), 78]], 3, ""],
    [[["姓名", "", "出勤天数", "迟到次数"], ...STAGE3_ATTENDANCE_ROWS], 2, ""]
  ];
  for (const [rows, column, hint] of cases) {
    const scope = await importWorkbookFile(buildXlsxBuffer([{ name: "S1", rows }]));
    assert.equal(scope.evaluate("state.rows.length"), 0);
    assert.equal(scope.elements.get("statusPanel").textContent, `工作表「S1」的第 ${column} 列缺少表头，请补充字段名后重新上传。${hint}`);
  }
});

function importCsvTextWithPapa(text) {
  const scope = createExcelContext();
  put(scope.context, "Papa", Papa);
  scope.context.window.Papa = Papa;
  put(scope.context, "__csvSourceText", text);
  scope.evaluate('parseCsvText(__csvSourceText, "t.csv", beginImport())');
  return scope;
}

const csvLines = (rows) => `${rows.map((row) => row.join(",")).join("\r\n")}\r\n`;
const STAGE3_GRADE_HEADER = ["学号", "姓名", "语文", "数学"];
const STAGE3_GRADE_ROWS = [["S01", "张三", "90", "85"], ["S02", "李四", "78", "92"], ["S03", "王五", "88", "70"]];

test("the grade fixtures import with the header from line 2 below their title row", async () => {
  for (const fileName of ["成绩表.csv", "成绩表-标题无逗号.csv"]) {
    const { evaluate } = await importStage3Fixture(fileName);
    assert.deepEqual(JSON.parse(evaluate("JSON.stringify(state.fields)")),
      ["学号", "姓名", "性别", "语文", "数学", "英语", "物理", "总分"], fileName);
    assert.equal(evaluate("state.rows.length"), 30, fileName);
    assert.equal(evaluate('getFieldProfile("数学").typeKey'), "numeric", fileName);
    assert.equal(evaluate('getFieldProfile("物理").typeKey'), "numeric", fileName);
  }
});

test("CSV imports skip a title, unit note and blank row, or a delimiter-only row", () => {
  const titled = importCsvTextWithPapa(csvLines([["期中成绩", "", "", ""], ["单位：分", "", "", ""], ["", "", "", ""], STAGE3_GRADE_HEADER, ...STAGE3_GRADE_ROWS]));
  assert.equal(titled.evaluate('state.fields.join(",")'), STAGE3_GRADE_HEADER.join(","));
  assert.equal(titled.evaluate("state.rows.length"), 3);

  const delimiterRow = importCsvTextWithPapa(csvLines([["", "", "", ""], STAGE3_GRADE_HEADER, ...STAGE3_GRADE_ROWS]));
  assert.equal(delimiterRow.evaluate('state.fields.join(",")'), STAGE3_GRADE_HEADER.join(","));
  assert.equal(delimiterRow.evaluate("state.rows.length"), 3);
});

test("a semicolon CSV below one title row is re-detected after skipping the title", () => {
  const scope = importCsvTextWithPapa("成绩单\r\n姓名;部门;分数\r\n张三;销售部;90\r\n李四;技术部;85\r\n王五;财务部;78\r\n");
  assert.equal(scope.evaluate('state.fields.join("|")'), "姓名|部门|分数");
  assert.equal(scope.evaluate('state.rows[2]["分数"]'), "78");
});

test("CSV header row numbers count physical lines above the header", () => {
  const { context, evaluate } = createExcelContext();
  put(context, "Papa", Papa);
  context.window.Papa = Papa;
  put(context, "__titledText", csvLines([["期中成绩", "", "", ""], ["单位：分", "", "", ""], ["", "", "", ""], STAGE3_GRADE_HEADER, ...STAGE3_GRADE_ROWS]));
  const found = JSON.parse(evaluate("JSON.stringify(findCsvHeaderBelowTitleRows(__titledText))"));
  assert.equal(found.headerDetection.headerRowNumber, 4);
  assert.deepEqual(found.headerDetection.skippedRows, [
    { rowNumber: 1, cells: ["期中成绩"] },
    { rowNumber: 2, cells: ["单位：分"] }
  ]);
  assert.equal(found.lineOffset, 3);
  assert.equal(context.__titledText.slice(found.offset).split("\r\n")[0], STAGE3_GRADE_HEADER.join(","));
});

const STAGE3_EIGHT_GRADE_ROWS = Array.from({ length: 8 }, (_, index) => [`S0${index + 1}`, `学生${index + 1}`, "80", "90"]);

test("a CSV title row above a data row with an extra cell falls back to the first error", () => {
  // The extra cell fails the candidate header's column check, so detection falls back.
  const rows = STAGE3_EIGHT_GRADE_ROWS.map((row) => [...row]);
  rows[6].push("多出");
  const scope = importCsvTextWithPapa(csvLines([["期中成绩"], STAGE3_GRADE_HEADER, ...rows]));
  assert.equal(scope.evaluate("state.rows.length"), 0);
  assert.equal(scope.elements.get("statusPanel").textContent, `CSV 第 2 行的数据列多于表头，继续分析会丢失额外单元格。请补全表头或修正该行后重试。${TITLE_FALLBACK_HINT}`);
});

test("CSV parse warning row numbers include the skipped title lines", () => {
  const rows = STAGE3_EIGHT_GRADE_ROWS.map((row) => [...row]);
  rows[6].pop();
  const scope = importCsvTextWithPapa(csvLines([["期中成绩"], STAGE3_GRADE_HEADER, ...rows]));
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(state.parseWarnings.map((warning) => warning.row))")), [9]);
});

test("CSV files whose first row is not a title keep their header errors", () => {
  const missingHeader = "CSV 存在缺少表头的数据列。请补全每一列的字段名后重试。";
  // A single-cell first row is skippable, so its error also carries the title hint.
  const cases = [
    [[["姓名", ""], ["学生1", "61"], ["学生2", "75"], ["学生3", "82"], ["学生4", "90"], ["学生5", "68"]], `${missingHeader}${TITLE_FALLBACK_HINT}`],
    [[["姓名", ""], ["学生1", "销售部"], ["学生2", "技术部"], ["学生3", "销售部"], ["学生4", "财务部"], ["学生5", "技术部"]], `${missingHeader}${TITLE_FALLBACK_HINT}`],
    [[["姓名", "性别", "", ""], ["张三", "男", "90", "85"], ["李四", "女", "78", "92"], ["王五", "男", "88", "70"]], missingHeader],
    [[["姓名", "工号", "", ""], ["张三", "E01", "销售部", "上海"], ["李四", "E02", "技术部", "北京"], ["王五", "E03", "财务部", "广州"]], missingHeader],
    [[["姓名", "", "出勤天数", "迟到次数"], ["张三", "销售部", "22", "1"], ["李四", "技术部", "21", "0"]], missingHeader],
    [[["城市"], ["上海"], ["北京"], ["广州", "待核实"], ["深圳"], ["杭州"]], `CSV 第 4 行的数据列多于表头，继续分析会丢失额外单元格。请补全表头或修正该行后重试。${TITLE_FALLBACK_HINT}`]
  ];
  for (const [rows, message] of cases) {
    const scope = importCsvTextWithPapa(csvLines(rows));
    assert.equal(scope.evaluate("state.rows.length"), 0);
    assert.equal(scope.elements.get("statusPanel").textContent, message);
  }
});

const STORE_HEADER_NOTE = "表头在第 4 行，已跳过上方 3 行：「2024年各门店销售报表」「制表日期：2024-08-01」。";
const GRADE_HEADER_NOTE = "表头在第 2 行，已跳过上方 1 行：「初二(3)班期中成绩」。";

test("the status bar names the header row and the skipped title rows", async () => {
  const store = await importStage3Fixture("门店报表.xlsx");
  assert.equal(store.elements.get("statusPanel").textContent, `门店报表.xlsx · 销售汇总 已成功载入，分析结果已生成。${STORE_HEADER_NOTE}第 33 行「合计」疑似合计行，可以在“数据质量”里设为不计入分析。`);
  assert.deepEqual(JSON.parse(store.evaluate("JSON.stringify(state.headerDetection)")), {
    headerRowNumber: 4,
    skippedRows: [
      { rowNumber: 1, preview: "2024年各门店销售报表" },
      { rowNumber: 2, preview: "制表日期：2024-08-01" }
    ]
  });
  for (const fileName of ["成绩表.csv", "成绩表-标题无逗号.csv"]) {
    const grades = await importStage3Fixture(fileName);
    assert.equal(grades.elements.get("statusPanel").textContent, `${fileName} 已成功载入，分析结果已生成。${GRADE_HEADER_NOTE}`);
  }
});

test("the header note follows the parse-warning status too", () => {
  const rows = STAGE3_EIGHT_GRADE_ROWS.map((row) => [...row]);
  rows[6].pop();
  const scope = importCsvTextWithPapa(csvLines([["期中成绩"], STAGE3_GRADE_HEADER, ...rows]));
  assert.equal(scope.elements.get("statusPanel").textContent,
    "CSV 已载入，但读取时有 1 处疑问，涉及 1 行。请核对数据预览和字段数量。表头在第 2 行，已跳过上方 1 行：「期中成绩」。");
});

test("reports carry the header note and a later first-row import clears it", async () => {
  const scope = await importStage3Fixture("门店报表.xlsx");
  put(scope.context, "__headerNote", STORE_HEADER_NOTE);
  const { report, html } = reportOf(scope);
  assert.deepEqual(report.headerDetection, JSON.parse(scope.evaluate("JSON.stringify(state.headerDetection)")));
  assert.ok(html.includes(`<div><dt>表头位置</dt><dd>${scope.evaluate("escapeHtml(__headerNote)")}</dd></div>`));

  const orders = await importStage3Fixture("订单表.csv");
  assert.equal(orders.evaluate("state.headerDetection"), null);
  assert.doesNotMatch(orders.elements.get("statusPanel").textContent, /表头在第/);
  const after = reportOf(orders);
  assert.equal(after.report.headerDetection, null);
  assert.doesNotMatch(after.html, /表头位置/);

  // The same scope importing again also replaces the previous detection.
  const sameScope = await importStage3Fixture("门店报表.xlsx");
  put(sameScope.context, "__orderRows", [{ 城市: "上海" }]);
  sameScope.evaluate('commitTabularData(["城市"], __orderRows, "later.csv", beginImport())');
  assert.equal(sameScope.evaluate("state.headerDetection"), null);
});

test("filtering and field configuration keep the header detection", async () => {
  const scope = await importStage3Fixture("门店报表.xlsx");
  const before = scope.evaluate("JSON.stringify(state.headerDetection)");
  assert.equal(JSON.parse(before ?? "null")?.headerRowNumber, 4);
  applyFilterState(scope, { category: { field: "城市", values: ["上海"] } });
  assert.equal(scope.evaluate("getAnalysisRows().length") < 29, true);
  assert.equal(scope.evaluate("JSON.stringify(state.headerDetection)"), before);
  scope.evaluate('stageFieldTypeChange("城市", "id"); applyFieldConfiguration()');
  assert.equal(scope.evaluate('getFieldProfile("城市").typeKey'), "id");
  assert.equal(scope.evaluate("JSON.stringify(state.headerDetection)"), before);
});

test("legacy report snapshots without header detection still render", () => {
  const scope = createScriptContext();
  assert.equal(scope.evaluate("normalizeReportSnapshot({}).headerDetection"), null);
  assert.doesNotMatch(scope.evaluate("buildHtmlReport({}, [])"), /表头位置/);
});

test("header errors add the title hint only when skipping was tried", async () => {
  const numericTitle = importCsvTextWithPapa(csvLines([["2024", "", "", ""], STAGE3_GRADE_HEADER, ...STAGE3_GRADE_ROWS]));
  assert.equal(numericTitle.elements.get("statusPanel").textContent, "CSV 存在缺少表头的数据列。请补全每一列的字段名后重试。");

  const vetoed = importCsvTextWithPapa(csvLines([["成绩单", ""], ["姓名", ""], ["学生1", "61"], ["学生2", "75"], ["学生3", "82"], ["学生4", "90"], ["学生5", "68"]]));
  assert.equal(vetoed.elements.get("statusPanel").textContent, `CSV 存在缺少表头的数据列。请补全每一列的字段名后重试。${TITLE_FALLBACK_HINT}`);

  const excel = await importWorkbookFile(buildXlsxBuffer([{
    name: "S1", rows: [["成绩单"], ["姓名"], ["学生1", 61], ["学生2", 75], ["学生3", 82], ["学生4", 90], ["学生5", 68]]
  }]));
  assert.equal(excel.elements.get("statusPanel").textContent, `工作表「S1」的第 2 列缺少表头，请补充字段名后重新上传。${TITLE_FALLBACK_HINT}`);
});

test("the title hint appears in the sheet selection panel for multi-sheet workbooks", async () => {
  const scope = await importWorkbookFile(buildXlsxBuffer([
    { name: "说明", rows: [["本工作簿为虚构数据"]] },
    { name: "数据", rows: [["成绩单"], ["姓名"], ["学生1", 61], ["学生2", 75], ["学生3", 82], ["学生4", 90], ["学生5", 68]] }
  ]));
  scope.elements.get("excelSheetSelect").value = "数据";
  // v3.2.1: the sheet is read after the list says so.
  const reading = scope.evaluate("analyzeSelectedExcelSheet()");
  assert.equal(scope.elements.get("sheetSelectionMessage").textContent, "正在读取并分析工作表「数据」...");
  await reading;
  assert.equal(scope.elements.get("sheetSelectionMessage").textContent,
    `工作表「数据」的第 2 列缺少表头，请补充字段名后重新上传。${TITLE_FALLBACK_HINT}`);
});

// File line numbers of data rows: the row a user sees in Excel, or the physical
// line (CR, LF and CRLF each end one) where a CSV record starts.
function sourceRowNumbers(evaluate) {
  return JSON.parse(evaluate("JSON.stringify(state.rows.map((row) => getSourceRowNumber(row)))"));
}

const lineRange = (first, last) => Array.from({ length: last - first + 1 }, (_, index) => first + index);

async function importCsvFile(text, fileName = "lines.csv") {
  const scope = createExcelContext();
  let papaCalls = 0;
  const papaOptions = [];
  const countingPapa = Object.assign(Object.create(Papa), {
    parse(...args) {
      papaCalls += 1;
      papaOptions.push({ header: args[1]?.header, preview: args[1]?.preview });
      return Papa.parse(...args);
    }
  });
  put(scope.context, "Papa", countingPapa);
  scope.context.window.Papa = countingPapa;
  const bytes = Buffer.from(text, "utf8");
  put(scope.context, "__csvFile", {
    name: fileName,
    size: bytes.length,
    arrayBuffer: async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.length)
  });
  await scope.evaluate("parseDataFile(__csvFile)");
  return { ...scope, papaCalls: () => papaCalls, papaOptions: () => papaOptions.map((options) => ({ ...options })) };
}

test("Excel row numbers follow the !ref start and skip blank rows inside the data", async () => {
  const worksheet = XLSX.utils.aoa_to_sheet([["编号", "名称"], ["A1", "x"], ["A2", "y"], ["", ""], ["A1", "z"], ["A3", "w"]], { origin: "B5" });
  worksheet["!ref"] = "B3:C10"; // rows 3-4 are blank inside the used range, row 8 is a blank data row
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "S1");
  const buffer = XLSX.write(workbook, { type: "array", bookType: "xlsx" });

  const { context, evaluate } = createExcelContext();
  put(context, "__offsetBuffer", buffer);
  const table = JSON.parse(evaluate("JSON.stringify(readExcelSheet(__offsetBuffer, \"S1\"))"));
  assert.deepEqual(table.rowNumbers, [6, 7, 9, 10]);

  const imported = await importWorkbookFile(buffer);
  assert.deepEqual(sourceRowNumbers(imported.evaluate), [6, 7, 9, 10]);
});

test("stage3 Excel fixtures map every row to its worksheet row", async () => {
  const store = await importStage3Fixture("门店报表.xlsx");
  assert.deepEqual(sourceRowNumbers(store.evaluate), lineRange(5, 33));
  const roster = await importStage3Fixture("花名册.xlsx");
  assert.deepEqual(sourceRowNumbers(roster.evaluate), lineRange(2, 31));
});

test("stage3 CSV fixtures map every row to its file line", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  assert.deepEqual(sourceRowNumbers(orders.evaluate), lineRange(2, 31));
  for (const fileName of ["成绩表.csv", "成绩表-标题无逗号.csv"]) {
    const grades = await importStage3Fixture(fileName);
    assert.deepEqual(sourceRowNumbers(grades.evaluate), lineRange(3, 32), fileName);
  }
});

// Title row, blank line under the header, a quoted line break, a delimiter-only row
// and a whitespace-only row; the data rows start on lines 4, 5, 9 and 11.
const LINE_NUMBER_CSV_LINES = [
  "期中成绩",
  "学号,姓名,备注",
  "",
  "S01,张三,",
  "S02,李四,\"第一行",
  "第二行\"",
  ",,",
  "   ",
  "S01,张三,",
  "",
  "S03,王五,ok"
];

test("CSV row numbers are the same for LF, CR-only and CRLF with a BOM", async () => {
  for (const [label, linebreak, bom] of [["LF", "\n", ""], ["CR", "\r", ""], ["CRLF+BOM", "\r\n", "\uFEFF"]]) {
    const scope = await importCsvFile(`${bom}${LINE_NUMBER_CSV_LINES.join(linebreak)}${linebreak}`);
    assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(state.rows.map((row) => row[\"学号\"]))")), ["S01", "S02", "S01", "S03"], label);
    assert.deepEqual(sourceRowNumbers(scope.evaluate), [4, 5, 9, 11], label);
  }
});

test("a quoted line break inside a CSV header moves the first data row to line 3", async () => {
  const scope = await importCsvFile("\"a\nb\",c\n1,2\n3,4\n");
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(state.fields)")), ["a\nb", "c"]);
  assert.deepEqual(sourceRowNumbers(scope.evaluate), [3, 4]);
});

test("CSV row numbers are all null when the line pass does not match the parsed rows", () => {
  const { context, evaluate } = createScriptContext();
  const papaStub = {
    parse(text, options) {
      if (options.header === false) {
        // One record fewer than the main parse returned.
        options.step({ data: ["姓名"], errors: [], meta: { cursor: 3 } });
        options.step({ data: ["张三"], errors: [], meta: { cursor: 6 } });
        return;
      }
      const fields = ["姓名"].map(options.transformHeader);
      options.complete({ errors: [], meta: { fields, delimiter: ",", linebreak: "\n" }, data: [{ [fields[0]]: "张三" }, { [fields[0]]: "李四" }] });
    }
  };
  context.Papa = papaStub;
  context.window.Papa = papaStub;
  evaluate('parseCsvText("姓名\\n张三\\n李四\\n", "mismatch.csv", beginImport())');
  assert.equal(evaluate("state.rows.length"), 2);
  assert.deepEqual(sourceRowNumbers(evaluate), [null, null]);
});

test("rows that did not come through an import have no row number", () => {
  const { context, evaluate } = createScriptContext();
  put(context, "__plainRows", [{ 城市: "上海" }, { 城市: "北京" }]);
  evaluate('commitTabularData(["城市"], __plainRows, "plain.csv", beginImport())');
  assert.deepEqual(sourceRowNumbers(evaluate), [null, null]);

  put(context, "__numberedRows", [{ 城市: "上海" }, { 城市: "北京" }]);
  evaluate('commitTabularData(["城市"], __numberedRows, "numbered.csv", beginImport(), [], "", null, null, [7, 9])');
  assert.deepEqual(sourceRowNumbers(evaluate), [7, 9]);

  evaluate('state.rows = [{ 城市: "广州" }, __numberedRows[0]]');
  assert.deepEqual(sourceRowNumbers(evaluate), [null, 7]);
  evaluate("resetAnalysisState()");
  assert.equal(evaluate("getSourceRowNumber(__numberedRows[0])"), null);
});

// Guards recorded on the code before row numbers existed: these files must report
// exactly the same errors and parse warnings, and an over-limit file is parsed once.
test("CSV edge files keep their errors and parse warnings", () => {
  const cases = [
    ["a,b\r\n", "status-panel error", "CSV 中没有可分析的数据。请确认文件包含表头和至少一行数据。", []],
    ["", "status-panel error", "CSV 中没有可分析的数据。请确认文件包含表头和至少一行数据。", []],
    ["a,b\r\n1\r\n", "status-panel warning", "CSV 已载入，但读取时有 1 处疑问，涉及 1 行。请核对数据预览和字段数量。", [
      { row: 2, code: "TooFewFields", message: "第 2 行只有 1 列（表头有 2 列），缺少的单元格按空值处理。" }
    ]],
    ["a,b\r\n1,2,3\r\n", "status-panel error", "CSV 第 2 行的数据列多于表头，继续分析会丢失额外单元格。请补全表头或修正该行后重试。", []]
  ];
  for (const [text, className, status, warnings] of cases) {
    const scope = importCsvTextWithPapa(text);
    const panel = scope.elements.get("statusPanel");
    assert.equal(panel.className, className, JSON.stringify(text));
    assert.equal(panel.textContent, status, JSON.stringify(text));
    // v3.2.0: stored as code and figures; the text is written by formatParseWarning().
    assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(state.parseWarnings.map((warning) => ({ row: warning.row, code: warning.code, message: formatParseWarning(warning) })))")), warnings, JSON.stringify(text));
  }
});

// Exact file lines for CSV warnings and "CSV 第 N 行" errors: FieldMismatch rows index
// the data records, Quotes rows every record read (header and blank records too).
const lineWarnings = (scope) => JSON.parse(scope.evaluate(
  'JSON.stringify(state.parseWarnings.filter((warning) => warning.code !== "UndetectableDelimiter").map((warning) => [warning.code, warning.row]))'
));
const EXTRA_CELL_ERROR = "的数据列多于表头，继续分析会丢失额外单元格。请补全表头或修正该行后重试。";
const exactGradeRow = (index) => `S${String(index).padStart(3, "0")},学生${index},80,90`;

test("CSV warnings name the exact line after blank lines, quoted line breaks and any line break", async () => {
  const cases = [
    ["blank line under the header", "a,b\n1,2\n\n3\n4,5\n", [["TooFewFields", 4]]],
    ["quoted line break", "a,b\n\"x\ny\",2\n3\n4,5\n", [["TooFewFields", 4]]],
    ["blank line under the header, CRLF", "a,b\r\n1,2\r\n\r\n3\r\n4,5\r\n", [["TooFewFields", 4]]],
    ["quoted line break, lone CR", "a,b\r\"x\ry\",2\r3\r4,5\r", [["TooFewFields", 4]]],
    ["whitespace-only row after a blank line (dropped from the rows)", "a,b\n1,2\n\n \n3,4\n", [["TooFewFields", 4]]],
    ["malformed quote", "a,b\n1,2\n\"x\"y\",3\n4,5\n", [["InvalidQuotes", 3]]],
    ["malformed quote after a blank line", "a,b\n1,2\n\n\"x\"y\",3\n4,5\n", [["InvalidQuotes", 4]]],
    ["malformed quote on the first data line", "a,b\n\"x\"y\",3\n4,5\n", [["InvalidQuotes", 2]]],
    ["title row, blank line under the header, short row (retry)", `期中成绩\n学号,姓名,语文,数学\n\n${[1, 2, 3, 4, 5].map(exactGradeRow).join("\n")}\nS006,学生6,80\n${exactGradeRow(7)}\n${exactGradeRow(8)}\n`, [["TooFewFields", 9]]]
  ];
  for (const [label, text, expected] of cases) {
    const scope = await importCsvFile(text);
    assert.deepEqual(lineWarnings(scope), expected, label);
  }
});

test("CSV errors name the exact line of an extra cell or an unclosed quote", async () => {
  const cases = [
    ["extra cell after a blank line", "a,b\n1,2\n\n3,4,5\n6,7\n", `CSV 第 4 行${EXTRA_CELL_ERROR}`],
    ["extra cell after a quoted line break", "a,b\n\"x\ny\",2\n3,4,5\n", `CSV 第 4 行${EXTRA_CELL_ERROR}`],
    ["extra cell after a blank line, CRLF", "a,b\r\n1,2\r\n\r\n3,4,5\r\n6,7\r\n", `CSV 第 4 行${EXTRA_CELL_ERROR}`],
    ["title row, blank line, then the header", `期中成绩\n\n学号,姓名,语文,数学\n${[1, 2, 3, 4, 5, 6].map(exactGradeRow).join("\n")}\nS007,学生7,80,90,多出\n`, `CSV 第 3 行${EXTRA_CELL_ERROR}${TITLE_FALLBACK_HINT}`],
    ["unclosed quote", "a,b\n1,2\n\"x,3\n4,5\n", "CSV 第 3 行开始的引号没有闭合，之后的内容无法正确分列。请检查该行的引号后重试。"],
    ["unclosed quote after a blank line", "a,b\n\n1,2\n\"x,3\n", "CSV 第 4 行开始的引号没有闭合，之后的内容无法正确分列。请检查该行的引号后重试。"]
  ];
  for (const [label, text, status] of cases) {
    const scope = await importCsvFile(text);
    assert.equal(scope.evaluate("state.rows.length"), 0, label);
    assert.equal(scope.elements.get("statusPanel").textContent, status, label);
  }
});

test("an extra cell found only after skipping title rows shows the re-parse's own line and header note", async () => {
  // Header detection reads 210 rows, so the extra cell on line 253 is found by the re-parse.
  const rows = Array.from({ length: 250 }, (_, index) => exactGradeRow(index + 1));
  const scope = await importCsvFile(`期中成绩\n学号,姓名,语文,数学\n${rows.join("\n")}\nS251,学生251,80,90,多出\n`);
  assert.equal(scope.evaluate("state.rows.length"), 0);
  assert.equal(scope.elements.get("statusPanel").textContent, `表头在第 2 行，已跳过上方 1 行：「期中成绩」。CSV 第 253 行${EXTRA_CELL_ERROR}`);
});

test("every CSV warning is counted, with thousands separators in the status", async () => {
  const shortRows = (count) => `a,b\r\n${Array.from({ length: 10 }, (_, index) => `${index},${index}`).join("\r\n")}\r\n${Array.from({ length: count }, (_, index) => String(100 + index)).join("\r\n")}\r\n`;
  const scope = await importCsvFile(shortRows(25));
  assert.equal(scope.elements.get("statusPanel").textContent, "CSV 已载入，但读取时有 25 处疑问，涉及 25 行。请核对数据预览和字段数量。");
  assert.equal(scope.evaluate("state.parseWarnings.length"), 25);
  assert.equal(scope.evaluate("state.parseWarnings[24].row"), 36);
  assert.equal(reportOf(scope).report.quality.parseWarningCount, 25);

  const many = await importCsvFile(shortRows(1000));
  assert.equal(many.elements.get("statusPanel").textContent, "CSV 已载入，但读取时有 1,000 处疑问，涉及 1,000 行。请核对数据预览和字段数量。");
});

test("CSV warnings get no line when the line pass does not match the parsed rows", () => {
  const { context, elements, evaluate } = createScriptContext();
  const papaStub = {
    parse(text, options) {
      if (options.header === false) {
        // One record fewer than the main parse returned.
        options.step({ data: ["姓名", "城市"], errors: [], meta: { cursor: 6 } }, { abort() {} });
        options.step({ data: ["张三"], errors: [], meta: { cursor: 9 } }, { abort() {} });
        return;
      }
      const fields = ["姓名", "城市"].map(options.transformHeader);
      options.complete({
        errors: [{ type: "FieldMismatch", code: "TooFewFields", message: "Too few fields", row: 1 }],
        meta: { fields, delimiter: ",", linebreak: "\n" },
        data: [{ [fields[0]]: "张三", [fields[1]]: "上海" }, { [fields[0]]: "李四" }]
      });
    }
  };
  context.Papa = papaStub;
  context.window.Papa = papaStub;
  evaluate('parseCsvText("姓名,城市\\n张三,上海\\n李四\\n", "mismatch.csv", beginImport())');
  assert.equal(evaluate("state.rows.length"), 2);
  assert.deepEqual(JSON.parse(evaluate("JSON.stringify(state.parseWarnings.map((warning) => warning.row))")), [null]);
  assert.equal(elements.get("statusPanel").textContent, "CSV 已载入，但读取时有 1 处疑问。请核对数据预览和字段数量。");
});

test("guard: imports that skip title rows or need no error line call Papa.parse as often as before", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  assert.equal(orders.papaCalls(), 2);
  // 成绩表-标题无逗号.csv fails its first parse with an extra cell, but the retry takes over.
  for (const fileName of ["成绩表.csv", "成绩表-标题无逗号.csv"]) {
    const grades = await importStage3Fixture(fileName);
    assert.equal(grades.evaluate("state.rows.length"), 30, fileName);
    assert.equal(grades.papaCalls(), 4, fileName);
  }
});

test("a shown extra-cell error adds one line pass after the main parse and header detection", async () => {
  const extra = await importCsvFile("a,b\r\n1,2,3\r\n");
  assert.equal(extra.elements.get("statusPanel").textContent, `CSV 第 2 行${EXTRA_CELL_ERROR}`);
  assert.deepEqual(extra.papaOptions().map((options) => options.header), [true, false, false]);
});

// v3.2.0: the text of a warning is written by formatParseWarning() from its code and figures.
const codedWarnings = (scope) => JSON.parse(scope.evaluate("JSON.stringify(state.parseWarnings.map((warning) => [warning.row, warning.code, formatParseWarning(warning)]))"));

test("imported CSV warnings carry Chinese messages and keep their codes", async () => {
  const short = await importCsvFile("a,b,c\n1,2,3\n4\n5,6,7\n");
  assert.deepEqual(codedWarnings(short), [[3, "TooFewFields", "第 3 行只有 1 列（表头有 3 列），缺少的单元格按空值处理。"]]);
  const quote = await importCsvFile("a,b\n1,2\n\"x\"y\",3\n4,5\n");
  assert.deepEqual(codedWarnings(quote), [[3, "InvalidQuotes", "第 3 行有不规范的引号（引号后面紧跟了其他字符），这一行可能被读错。"]]);

  // Without the source text there is no line; the cell count still comes from the row.
  const { evaluate } = createScriptContext();
  evaluate(`handleParsedData({
    errors: [{ type: "FieldMismatch", code: "TooFewFields", message: "Too few fields", row: 1 }],
    meta: { fields: ["a", "b", "c"] },
    data: [{ a: "1", b: "2", c: "3" }, { a: "4" }]
  }, "no-source.csv", beginImport())`);
  assert.deepEqual(JSON.parse(evaluate("JSON.stringify(state.parseWarnings.map((warning) => ({ row: warning.row, code: warning.code, message: formatParseWarning(warning) })))")), [
    { row: null, code: "TooFewFields", message: "有一条记录的列数少于表头（只有 1 列，表头有 3 列），缺少的单元格按空值处理。" }
  ]);
});

test("a one-column CSV whose header holds a semicolon, tab or bar gets one Chinese delimiter warning", async () => {
  for (const [separator, name, short] of [[";", "分号（；）", "分号"], ["\t", "制表符", "制表符"], ["|", "竖线（|）", "竖线"]]) {
    const scope = await importCsvFile(["a", "b"].join(separator) + `\n1${separator}2\n3\n5${separator}6\n`);
    assert.equal(scope.evaluate("state.fields.length"), 1, name);
    assert.deepEqual(codedWarnings(scope), [
      [null, "UndetectableDelimiter", `整份文件只读出 1 列，但表头里有${name}。如果原文件是用${short}分隔的，请另存为逗号分隔的 CSV 后重新上传。`]
    ], name);
    assert.equal(scope.elements.get("statusPanel").textContent, "CSV 已载入，但读取时有 1 处疑问。请核对数据预览和字段数量。", name);
  }
});

test("guard: a plain one-column CSV still gets no delimiter warning", async () => {
  const scope = await importCsvFile("字段\n1\n2\n");
  assert.deepEqual(codedWarnings(scope), []);
  assert.equal(scope.elements.get("statusPanel").className, "status-panel success");
});

test("a comma CSV with several columns no longer lists the delimiter guess as a warning", async () => {
  const scope = await importCsvFile("a,b\n1,2\n3\n5,6\n");
  assert.deepEqual(codedWarnings(scope).map(([row, code]) => [row, code]), [[3, "TooFewFields"]]);
});

// ID duplicate detection on the original rows (state.idDuplicates).
function idDuplicates(evaluate) {
  return JSON.parse(evaluate("JSON.stringify(state.idDuplicates)"));
}

const listedValues = (item) => item.values.map(({ value, count, rowNumbers, identicalRowCount }) => [value, count, rowNumbers, identicalRowCount]);

test("the order fixture reports DD1005 twice on lines 6 and 19", async () => {
  const { evaluate } = await importStage3Fixture("订单表.csv");
  assert.deepEqual(idDuplicates(evaluate), [{
    field: "订单号",
    nonMissingCount: 30,
    distinctCount: 29,
    uniqueRatio: 29 / 30,
    mostlyRepeated: false,
    duplicateValueCount: 1,
    values: [{ value: "DD1005", count: 2, rowNumbers: [6, 19], identicalRowCount: 0, identicalGroupCount: 0 }]
  }]);
});

test("the roster fixture reports one value three times in each ID field", async () => {
  const { evaluate } = await importStage3Fixture("花名册.xlsx");
  const summary = idDuplicates(evaluate);
  assert.deepEqual(summary.map((item) => item.field), ["工号", "身份证号", "手机号"]);
  assert.deepEqual(summary.map(listedValues), [
    [["E1012", 3, [13, 14, 25], 2]],
    [["990101199005160123", 3, [13, 14, 25], 2]],
    [["13800010012", 3, [13, 14, 25], 2]]
  ]);
  assert.deepEqual(summary.map((item) => item.values[0].identicalGroupCount), [1, 1, 1]);
});

test("the grade and store fixtures have no ID duplicates", async () => {
  for (const [fileName, fields] of [["成绩表.csv", ["学号", "姓名"]], ["成绩表-标题无逗号.csv", ["学号", "姓名"]], ["门店报表.xlsx", ["门店编号", "门店名称"]]]) {
    const summary = idDuplicates((await importStage3Fixture(fileName)).evaluate);
    assert.deepEqual(summary.map((item) => item.field), fields, fileName);
    assert.ok(summary.every((item) => !item.mostlyRepeated && item.duplicateValueCount === 0 && !item.values.length), fileName);
  }
});

test("user and session IDs in the behavior example are mostly repeated, event IDs are listed", async () => {
  const text = readExampleCsv("sample-user-behavior.csv");
  const summary = idDuplicates((await importCsvFile(text, "sample-user-behavior.csv")).evaluate);
  assert.deepEqual(summary.map((item) => [item.field, item.mostlyRepeated, item.values.length]), [
    ["事件ID", false, 2],
    ["用户ID", true, 0],
    ["会话ID", true, 0]
  ]);
  assert.deepEqual(listedValues(summary[0]), [["EV-00026", 2, [27, 120], 2], ["EV-00065", 2, [66, 121], 2]]);
  assert.deepEqual(summary.slice(1).map((item) => [item.distinctCount, item.nonMissingCount]), [[46, 119], [30, 120]]);
});

test("an Excel number and the same text are one ID value", async () => {
  const scope = await importWorkbookFile(buildXlsxBuffer([{
    name: "S1", rows: [["工号", "部门"], [1012, "销售"], ["1012", "销售"], [1013, "技术"], ["001012", "销售"]]
  }]));
  assert.deepEqual(listedValues(idDuplicates(scope.evaluate)[0]), [["1012", 2, [2, 3], 2]]);
});

test("long ID duplicate lists stop at 10 values and 10 row numbers", async () => {
  // 40 unique IDs keep the unique share at 52 / 77, so the field is not mostly repeated.
  const ids = [
    ...Array.from({ length: 15 }, () => "S000"),
    ...Array.from({ length: 11 }, (_, index) => [`S${index + 1}`, `S${index + 1}`]).flat(),
    ...Array.from({ length: 40 }, (_, index) => `S${index + 100}`)
  ];
  const scope = await importCsvFile(`学号\r\n${ids.join("\r\n")}\r\n`);
  const [summary] = idDuplicates(scope.evaluate);
  put(scope.context, "__longSummary", summary);
  assert.equal(scope.evaluate("formatRowNumberList(__longSummary.values[0].rowNumbers, __longSummary.values[0].count)"), "2、3、4、5、6、7、8、9、10、11 等 15 个");
  assert.equal(scope.evaluate("formatIdDuplicateOverflow(__longSummary.duplicateValueCount, __longSummary.values.length)"), "共 12 个重复值，只列出前 10 个");
});

// 工号 repeats once; 城市 has three values over 12 rows.
const ID_CONFIG_CSV = ["工号,城市,金额", ...Array.from({ length: 12 }, (_, index) => `${index === 11 ? "E01" : `E${String(index + 1).padStart(2, "0")}`},${["上海", "北京", "广州"][index % 3]},${100 + index}`)].join("\r\n");

test("ID duplicates follow field configuration but not filters", async () => {
  const scope = await importCsvFile(`${ID_CONFIG_CSV}\r\n`);
  const fieldsWithResults = () => idDuplicates(scope.evaluate).map((item) => item.field);
  assert.deepEqual(fieldsWithResults(), ["工号"]);
  assert.deepEqual(listedValues(idDuplicates(scope.evaluate)[0]), [["E01", 2, [2, 13], 0]]);

  scope.evaluate('stageFieldTypeChange("城市", "id"); applyFieldConfiguration()');
  assert.deepEqual(fieldsWithResults(), ["工号", "城市"]);
  assert.equal(idDuplicates(scope.evaluate)[1].mostlyRepeated, true);

  scope.evaluate('stageFieldTypeChange("工号", "category"); applyFieldConfiguration()');
  assert.deepEqual(fieldsWithResults(), ["城市"]);

  scope.evaluate("restoreAutomaticFieldTypes()");
  assert.deepEqual(fieldsWithResults(), ["工号"]);

  const before = scope.evaluate("JSON.stringify(state.idDuplicates)");
  applyFilterState(scope, { category: { field: "城市", values: ["上海"] } });
  assert.equal(scope.evaluate("getAnalysisRows().length"), 4);
  assert.equal(scope.evaluate("JSON.stringify(state.idDuplicates)"), before);
});

test("a later import without ID fields clears the previous ID duplicates", async () => {
  const scope = await importStage3Fixture("订单表.csv");
  assert.equal(idDuplicates(scope.evaluate).length, 1);
  put(scope.context, "__cityRows", [{ 城市: "上海" }, { 城市: "上海" }]);
  scope.evaluate('commitTabularData(["城市"], __cityRows, "cities.csv", beginImport())');
  assert.deepEqual(idDuplicates(scope.evaluate), []);
  scope.evaluate("beginImport()");
  assert.deepEqual(idDuplicates(scope.evaluate), []);
});

test("a CSV over the row limit still fails before any line-number pass", async () => {
  const text = ["值", ...Array.from({ length: 100001 }, (_, index) => String(index))].join("\r\n");
  const scope = await importCsvFile(`${text}\r\n`, "over.csv");
  assert.equal(scope.evaluate("state.rows.length"), 0);
  assert.equal(scope.elements.get("statusPanel").textContent, "CSV 超过 100,000 行数据限制，请拆分文件后重试。");
  assert.equal(scope.papaCalls(), 1);
});

// The row preview (maxRows + 1) counts blank records too, so blank lines used to
// leave a truncated parse under the limit and drop the rows after it silently.
const blankLineRows = (count) => Array.from({ length: count }, (_, index) => `R${index},${index}\r\n\r\n`).join("");

test("a CSV whose blank lines fill the row preview imports every row", async () => {
  const scope = await importCsvFile(`编号,值\r\n${blankLineRows(60000)}`, "blank-lines.csv");
  assert.equal(scope.evaluate("state.rows.length"), 60000);
  assert.equal(scope.evaluate('state.rows[59999]["编号"]'), "R59999");
  assert.equal(scope.evaluate("getSourceRowNumber(state.rows[0])"), 2);
  assert.equal(scope.evaluate("getSourceRowNumber(state.rows[59999])"), 120000);
  assert.equal(scope.evaluate("state.rows.every((row) => getSourceRowNumber(row) !== null)"), true);
  assert.equal(scope.elements.get("statusPanel").className, "status-panel success");
  // Guard: the first parse keeps its row preview; the whole text is parsed again only after the count.
  assert.deepEqual(scope.papaOptions(), [
    { header: true, preview: 100001 },
    { header: false, preview: undefined },
    { header: true, preview: undefined },
    { header: false, preview: undefined }
  ]);
});

test("a CSV over the row limit fails even when blank lines fill the row preview", async () => {
  const text = ["值", "", "", "", "", "", ...Array.from({ length: 100001 }, (_, index) => String(index))].join("\r\n");
  const scope = await importCsvFile(`${text}\r\n`, "over-with-blanks.csv");
  assert.equal(scope.evaluate("state.rows.length"), 0);
  assert.equal(scope.elements.get("statusPanel").textContent, "CSV 超过 100,000 行数据限制，请拆分文件后重试。");
  assert.equal(scope.papaCalls(), 2);
});

test("guard: a CSV with exactly 100,000 rows and trailing blank lines imports all of them", async () => {
  const text = ["值", ...Array.from({ length: 100000 }, (_, index) => String(index)), "", "", ""].join("\r\n");
  const scope = await importCsvFile(`${text}\r\n`, "exact-with-blanks.csv");
  assert.equal(scope.evaluate("state.rows.length"), 100000);
  assert.equal(scope.evaluate("getSourceRowNumber(state.rows[99999])"), 100001);
});

test("a titled CSV whose blank lines fill the row preview imports every row after the retry", async () => {
  const scope = await importCsvFile(`成绩单\r\n编号,值\r\n${blankLineRows(60000)}`, "titled-blank-lines.csv");
  assert.equal(scope.evaluate('state.fields.join(",")'), "编号,值");
  assert.equal(scope.evaluate("state.rows.length"), 60000);
  assert.equal(scope.evaluate("getSourceRowNumber(state.rows[0])"), 3);
  assert.equal(scope.evaluate("getSourceRowNumber(state.rows[59999])"), 120001);
  assert.match(scope.elements.get("statusPanel").textContent, /表头在第 2 行，已跳过上方 1 行：「成绩单」。$/);
});

// ID duplicates on the quality page (#idDuplicateDetails) and in reports.
const idDetails = (scope) => scope.elements.get("idDuplicateDetails").innerHTML;
const behaviorCsv = () => readExampleCsv("sample-user-behavior.csv");

test("the quality page has a static ID duplicate container in the dom map", () => {
  const html = indexHtmlSource();
  const quality = html.slice(html.indexOf('id="qualitySection"'), html.indexOf('id="filterSection"'));
  assert.match(quality, /<h3 class="subsection-title">编号重复<\/h3>\s*<div id="idDuplicateDetails"/);
  const { evaluate } = createScriptContext();
  assert.equal(evaluate('Object.keys(dom).includes("idDuplicateDetails")'), true);
});

test("the ID duplicate pill counts listed fields after the existing pills", async () => {
  const pillTexts = (scope) => [...scope.elements.get("qualitySummary").innerHTML.matchAll(/<span class="quality-pill[^"]*">([^<]*)<\/span>/g)].map((match) => match[1]);
  const orders = await importStage3Fixture("订单表.csv");
  assert.deepEqual(pillTexts(orders).map((text) => text.split("：")[0]), ["空值总数", "重复行", "认不出的数字或日期", "偏大或偏小的值", "编号重复", "像是空值的写法"]);
  assert.deepEqual(qualityPill(orders.elements, "编号重复"), { tone: "warning", text: "编号重复：1 个字段" });
  const grades = await importStage3Fixture("成绩表.csv");
  assert.deepEqual(qualityPill(grades.elements, "编号重复"), { tone: "success", text: "编号重复：0 个字段" });
  // 用户ID and 会话ID are mostly repeated, so only 事件ID counts.
  const behavior = await importCsvFile(behaviorCsv(), "sample-user-behavior.csv");
  assert.deepEqual(qualityPill(behavior.elements, "编号重复"), { tone: "warning", text: "编号重复：1 个字段" });
});

test("the quality page lists each ID field's duplicate values in its own table", async () => {
  const roster = await importStage3Fixture("花名册.xlsx");
  const html = idDetails(roster);
  ["工号", "身份证号", "手机号"].forEach((field) => assert.ok(html.includes(`<table aria-label="${field} 的重复值">`), field));
  assert.match(html, /<th>重复值<\/th><th>出现次数<\/th><th>文件行号<\/th><th>说明<\/th>/);
  assert.ok(html.includes("<td>E1012</td><td>3</td><td>13、14、25</td><td>其中 2 行内容完全相同</td>"));

  const behavior = await importCsvFile(behaviorCsv(), "sample-user-behavior.csv");
  put(behavior.context, "__hintItem", idDuplicates(behavior.evaluate)[1]);
  assert.ok(idDetails(behavior).includes(behavior.evaluate("escapeHtml(formatMostlyRepeatedIdHint(__hintItem))")));
  assert.ok(!idDetails(behavior).includes('aria-label="用户ID 的重复值"'));

  const grades = await importStage3Fixture("成绩表.csv");
  assert.ok(idDetails(grades).includes("类型为 ID 的字段没有编号重复（已检查：学号、姓名）"));

  const none = createScriptContext();
  commitRows(none, ["城市"], [{ 城市: "上海" }, { 城市: "上海" }], "cities.csv");
  assert.ok(idDetails(none).includes("未识别到 ID 字段。"));

  const mixed = createScriptContext();
  commitRows(mixed, ["学号", "订单号"], [{ 学号: "S1", 订单号: "D1" }, { 学号: "S2", 订单号: "D1" }], "mixed.csv");
  assert.ok(idDetails(mixed).includes('<table aria-label="订单号 的重复值">'));
  assert.ok(idDetails(mixed).includes("其余类型为 ID 的字段没有编号重复（已检查：学号）"));
  assert.ok(idDetails(mixed).includes("<td>D1</td><td>2</td><td>—</td><td></td>"));
});

test("ID duplicate details show without numeric fields, escape data and note long lists", async () => {
  const textOnly = createScriptContext();
  commitRows(textOnly, ["<i>编号</i>", "姓名"], [{ "<i>编号</i>": "<b>x</b>", 姓名: "甲" }, { "<i>编号</i>": "<b>x</b>", 姓名: "乙" }], "text-only.csv");
  assert.match(textOnly.elements.get("outlierTable").innerHTML, /未识别到数值字段/);
  const html = idDetails(textOnly);
  assert.ok(html.includes("<td>&lt;b&gt;x&lt;/b&gt;</td><td>2</td>"));
  assert.ok(html.includes("&lt;i&gt;编号&lt;/i&gt;"));
  assert.doesNotMatch(html, /<b>x<\/b>|<i>编号/);

  const ids = [
    ...Array.from({ length: 15 }, () => "S000"),
    ...Array.from({ length: 11 }, (_, index) => [`S${index + 1}`, `S${index + 1}`]).flat(),
    ...Array.from({ length: 40 }, (_, index) => `S${index + 100}`)
  ];
  const long = await importCsvFile(`学号\r\n${ids.join("\r\n")}\r\n`);
  assert.ok(idDetails(long).includes("<td>S000</td><td>15</td><td>2、3、4、5、6、7、8、9、10、11 等 15 个</td>"));
  assert.ok(idDetails(long).includes("共 12 个重复值，只列出前 10 个"));

  long.evaluate("beginImport()");
  assert.equal(idDetails(long), "");
});

test("reports add an ID duplicate section after the outliers and count it in the summary", async () => {
  const roster = await importStage3Fixture("花名册.xlsx");
  const { report, html } = reportOf(roster);
  assert.deepEqual(report.quality.idDuplicates, idDuplicates(roster.evaluate));
  assert.ok(html.includes("；读取时有疑问：<strong>0</strong> 处；编号重复：<strong>3</strong> 个字段；日期写法混用：<strong>1</strong> 个字段。</p>"));
  const htmlSection = html.slice(html.indexOf("<h3>编号重复</h3>"), html.indexOf("数值字段描述性统计"));
  assertInOrder(html, ["<h3>偏大或偏小的值</h3>", "<h3>编号重复</h3>"]);
  assert.match(htmlSection, /<th>字段<\/th><th>重复值<\/th><th>出现次数<\/th><th>文件行号<\/th><th>说明<\/th>/);
  assert.ok(htmlSection.includes("<td>工号</td><td>E1012</td><td>3</td><td>13、14、25</td><td>其中 2 行内容完全相同</td>"));
  assertInOrder(html, ["<h3>编号重复</h3>", "<h2>数值字段描述性统计</h2>"]);

});

test("report ID duplicates show mostly repeated fields, overflow, clean fields and no ID fields", async () => {
  const behavior = await importCsvFile(behaviorCsv(), "sample-user-behavior.csv");
  const behaviorReport = reportOf(behavior);
  put(behavior.context, "__hintItem", idDuplicates(behavior.evaluate)[1]);
  const hint = behavior.evaluate("formatMostlyRepeatedIdHint(__hintItem)");
  assert.ok(behaviorReport.html.includes(`<td>用户ID</td><td>—</td><td>—</td><td>—</td><td>${behavior.evaluate("escapeHtml(formatMostlyRepeatedIdHint(__hintItem))")}</td>`));
  assert.ok(behaviorReport.html.includes("<td>事件ID</td><td>EV-00026</td><td>2</td><td>27、120</td><td>其中 2 行内容完全相同</td>"));
  assert.ok(behaviorReport.html.includes("；编号重复：<strong>1</strong> 个字段。"));

  const ids = [
    ...Array.from({ length: 15 }, () => "S000"),
    ...Array.from({ length: 11 }, (_, index) => [`S${index + 1}`, `S${index + 1}`]).flat(),
    ...Array.from({ length: 40 }, (_, index) => `S${index + 100}`)
  ];
  const long = reportOf(await importCsvFile(`学号\r\n${ids.join("\r\n")}\r\n`));
  assert.ok(long.html.includes("<p>「学号」共 12 个重复值，只列出前 10 个</p>"));

  const grades = reportOf(await importStage3Fixture("成绩表.csv"));
  const gradeSection = grades.html.slice(grades.html.indexOf("<h3>编号重复</h3>"), grades.html.indexOf("数值字段描述性统计"));
  assert.ok(gradeSection.includes("<p>类型为 ID 的字段没有编号重复（已检查：学号、姓名）</p>"));
  assert.doesNotMatch(gradeSection, /<table>/);
  assert.ok(grades.html.includes("；编号重复：<strong>0</strong> 个字段。"));

  const none = createScriptContext();
  commitRows(none, ["城市"], [{ 城市: "上海" }, { 城市: "北京" }], "cities.csv");
  const noneReport = reportOf(none);
  assert.deepEqual(noneReport.report.quality.idDuplicates, []);
  assert.ok(noneReport.html.includes("<p>未识别到 ID 字段。</p>"));
  assert.match(noneReport.html, /<h3>编号重复<\/h3>\s*<p>未识别到 ID 字段。<\/p>/);
});

test("report ID duplicate values are escaped for HTML", () => {
  const scope = createScriptContext();
  const value = "<b>x</b>|y\nz";
  commitRows(scope, ["编号", "姓名"], [{ 编号: value, 姓名: "甲" }, { 编号: value, 姓名: "乙" }], "escape.csv");
  const { html } = reportOf(scope);
  assert.ok(html.includes("<td>编号</td><td>&lt;b&gt;x&lt;/b&gt;|y\nz</td><td>2</td><td>—</td><td></td>"));
  assert.doesNotMatch(html, /<b>x<\/b>/);
});

test("legacy report snapshots without ID duplicates leave the section out", () => {
  const scope = createScriptContext();
  assert.equal(scope.evaluate("normalizeReportSnapshot({}).quality.idDuplicates"), null);
  const legacy = { quality: { totalMissing: 1, duplicateRows: 0, parseWarningCount: 0, outliers: [{ field: "金额", outlierCount: 1, lowerBound: "1", upperBound: "9" }] } };
  put(scope.context, "__legacyQuality", legacy);
  const html = scope.evaluate("buildHtmlReport(__legacyQuality, [])");
  assert.doesNotMatch(html, /<h3>编号重复<\/h3>|；编号重复：/);
  assert.ok(html.includes("；读取时有疑问：<strong>0</strong> 处。</p>"));
  assert.ok(html.includes("<td>金额</td><td>1</td><td>1</td><td>9</td>"));
});

// Placeholder missing values (stage 3 item 3): candidates from the original rows.
const placeholderEntries = (evaluate) => JSON.parse(evaluate("JSON.stringify(state.placeholders.map(({ field, key, tier, count, rowNumbers }) => [field, key, tier, count, rowNumbers]))"));

test("the order fixture lists its five placeholder cells with tiers and file rows", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  // known-issues.json O6-O10.
  assert.deepEqual(placeholderEntries(orders.evaluate), [
    ["下单日期", "—", "common", 1, [17]],
    ["数量", "—", "common", 1, [12]],
    ["金额", "N/A", "common", 1, [15]],
    ["备注", "N/A", "common", 1, [29]],
    ["备注", "—", "common", 1, [30]]
  ]);
  assert.equal(orders.evaluate("JSON.stringify(state.placeholderConfirmations)"), "{}");
});

test("the grade fixtures and the bundled examples have no placeholder candidates, except the two the sample button's file holds", async () => {
  for (const fileName of ["成绩表.csv", "成绩表-标题无逗号.csv"]) {
    const scope = await importStage3Fixture(fileName);
    assert.equal(scope.evaluate("state.rows.length"), 30, fileName);
    assert.deepEqual(placeholderEntries(scope.evaluate), [], fileName);
  }
  const examples = listExampleCsvFiles();
  assert.equal(examples.length, 8);
  for (const fileName of examples) {
    const scope = await importCsvFile(readExampleCsv(fileName), fileName);
    assert.ok(scope.evaluate("state.rows.length") > 0, fileName);
    const fieldsAndKeys = placeholderEntries(scope.evaluate).map(([field, key]) => [field, key]);
    const expected = { "sample-orders.csv": [["金额", "N/A"], ["销售代表", "\u2014"]], "sample-orders-en.csv": [["Amount", "N/A"], ["Sales rep", "\u2014"]] }[fileName] || [];
    assert.deepEqual(fieldsAndKeys, expected, fileName);
  }
});

test("ignored fields are not checked for placeholders", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  orders.evaluate('stageFieldTypeChange("备注", "ignore"); applyFieldConfiguration()');
  assert.deepEqual(placeholderEntries(orders.evaluate).map(([field, key]) => [field, key]), [["下单日期", "—"], ["数量", "—"], ["金额", "N/A"]]);
  orders.evaluate("restoreAutomaticFieldTypes()");
  assert.equal(placeholderEntries(orders.evaluate).length, 5);
});

test("placeholder candidates and confirmations are cleared on reset and replaced on every import", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  orders.evaluate('state.placeholderConfirmations = { 金额: ["N/A"] }');
  commitRows(orders, ["城市"], [{ 城市: "上海" }, { 城市: "N/A" }], "cities.csv");
  assert.equal(orders.evaluate("JSON.stringify(state.placeholderConfirmations)"), "{}");
  assert.deepEqual(placeholderEntries(orders.evaluate), [["城市", "N/A", "common", 1, [null]]]);

  orders.evaluate("resetAnalysisState()");
  assert.equal(orders.evaluate("JSON.stringify(state.placeholderConfirmations)"), "{}");
  assert.deepEqual(placeholderEntries(orders.evaluate), []);
});

// Confirming placeholders (stage 3 item 3): the field's profile, analysis and
// quality treat the confirmed key as missing; the original rows never change.
const placeholderConfirmations = (evaluate) => JSON.parse(evaluate("JSON.stringify(state.placeholderConfirmations)"));
const manualFieldTypes = (evaluate) => JSON.parse(evaluate("JSON.stringify(state.manualFieldTypes)"));
const fieldFigures = (evaluate, field) => JSON.parse(evaluate(`JSON.stringify((({ typeKey, inferredTypeKey, missingCount, uniqueCount, conversionFailureCount }) => ({ typeKey, inferredTypeKey, missingCount, uniqueCount, conversionFailureCount }))(getFieldProfile(${JSON.stringify(field)})))`));
const placeholderMessage = (scope) => scope.elements.get("placeholderMessage").textContent;

function commitColumns(columns, fileName = "placeholders.csv") {
  const scope = createScriptContext();
  const fields = Object.keys(columns);
  const rows = columns[fields[0]].map((_, index) => Object.fromEntries(fields.map((field) => [field, columns[field][index]])));
  commitRows(scope, fields, rows, fileName);
  return scope;
}

const AMOUNT_WITH_NA = ["100", "200", "N/A", "约500", "300", "400"];
const YEARS_WITH_DASHES = ["2020", "2021", "—", "—", "2022", "2023", "—"];

test("confirming a placeholder re-profiles its field: 4 of 5 amounts make it numeric", () => {
  const scope = commitColumns({ 金额: AMOUNT_WITH_NA });
  assert.deepEqual(fieldFigures(scope.evaluate, "金额"), { typeKey: "category", inferredTypeKey: "category", missingCount: 0, uniqueCount: 6, conversionFailureCount: 0 });
  assert.equal(scope.evaluate('confirmPlaceholder("金额", "N/A")'), true);
  assert.deepEqual(fieldFigures(scope.evaluate, "金额"), { typeKey: "numeric", inferredTypeKey: "numeric", missingCount: 1, uniqueCount: 5, conversionFailureCount: 1 });
  assert.deepEqual(JSON.parse(scope.evaluate('JSON.stringify(getFieldProfile("金额").conversionFailureSamples)')), ["约500"]);
  assert.deepEqual(placeholderConfirmations(scope.evaluate), { 金额: ["N/A"] });
  assert.equal(scope.evaluate('state.fieldTypeDrafts["金额"]'), "numeric");
  assert.equal(scope.evaluate("hasFieldConfigDraftChanges()"), false);
  assert.equal(scope.evaluate("state.totalMissing"), 1);
  assert.equal(scope.evaluate("state.placeholders[0].confirmed"), true);
  assert.equal(placeholderMessage(scope), "已把「金额」的 1 个「N/A」按空值处理：空值 0 → 1，认不出 0 → 1。只是把这 1 格当作空，这一行没有删，其他字段照常参与分析，所以行数还是 6。「金额」现在识别为数值（原来是分类）。");
});

test("the order fixture goes from 26 to 27 missing values with one confirmation and to 31 with all five", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  const missingByField = () => JSON.parse(orders.evaluate("JSON.stringify(Object.fromEntries(state.profiles.map((profile) => [profile.field, profile.missingCount])))"));
  const types = () => JSON.parse(orders.evaluate("JSON.stringify(state.profiles.map((profile) => profile.typeKey))"));
  const orderTypes = ["id", "date", "category", "category", "category", "numeric", "numeric", "category"];
  assert.deepEqual(missingByField(), { 订单号: 0, 下单日期: 0, 客户名称: 0, 大区: 0, 产品: 0, 数量: 0, 金额: 0, 备注: 26 });

  orders.evaluate('confirmPlaceholder("金额", "N/A")');
  assert.equal(orders.evaluate("state.totalMissing"), 27);
  assert.deepEqual(fieldFigures(orders.evaluate, "金额"), { typeKey: "numeric", inferredTypeKey: "numeric", missingCount: 1, uniqueCount: 29, conversionFailureCount: 0 });

  orders.evaluate("confirmAllCommonPlaceholders()");
  assert.equal(orders.evaluate("state.totalMissing"), 31);
  assert.deepEqual(missingByField(), { 订单号: 0, 下单日期: 1, 客户名称: 0, 大区: 0, 产品: 0, 数量: 1, 金额: 1, 备注: 28 });
  assert.deepEqual(types(), orderTypes);
  assert.equal(orders.evaluate('getFieldProfile("下单日期").conversionFailureCount'), 0);
  assert.equal(orders.evaluate('getFieldProfile("数量").conversionFailureCount'), 0);
  assert.equal(orders.evaluate("state.placeholders.every((entry) => entry.confirmed)"), true);
  assert.deepEqual(placeholderConfirmations(orders.evaluate), { 下单日期: ["—"], 数量: ["—"], 金额: ["N/A"], 备注: ["N/A", "—"] });
});

test("confirming all common placeholders skips maybe-valid keys and ignored fields and names type changes first", () => {
  const scope = commitColumns({
    金额: ["100", "200", "N/A", "约500", "300", "400", "500"],
    年份: YEARS_WITH_DASHES,
    是否会员: ["有", "无", "有", "无", "有", "有", "有"],
    说明: ["N/A", "甲", "乙", "丙", "丁", "戊", "己"]
  });
  scope.evaluate('stageFieldTypeChange("说明", "ignore"); applyFieldConfiguration()');
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(previewAllCommonPlaceholders())")), [
    { field: "金额", typeKey: "numeric", previousTypeKey: "category" },
    { field: "年份", typeKey: "date", previousTypeKey: "category" }
  ]);
  scope.evaluate("confirmAllCommonPlaceholders()");
  assert.deepEqual(placeholderConfirmations(scope.evaluate), { 金额: ["N/A"], 年份: ["—"] });
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(state.profiles.map((profile) => profile.typeKey))")), ["numeric", "date", "category", "ignore"]);
  assert.equal(placeholderMessage(scope), "已把 2 处常见写法按空值处理（「金额」的「N/A」、「年份」的「—」）。「金额」现在识别为数值（原来是分类）；「年份」现在识别为日期（原来是分类）。只是把这些格当作空，行都没有删，所以行数还是 7。");
});

test("confirmed placeholders leave Top 10, filter options, filter matches and groups as empty values", () => {
  const scope = commitColumns({ 城市: ["上海", "N/A", "北京", "N/A", "上海", ""], 金额: ["10", "20", "30", "40", "50", "60"] });
  const { evaluate } = scope;
  const cityStats = () => JSON.parse(evaluate('JSON.stringify(buildCategoryStats().find((item) => item.field === "城市"))'));
  assert.deepEqual(cityStats().top.map((item) => item.name), ["上海", "N/A", "北京"]);

  evaluate('confirmPlaceholder("城市", "N/A")');
  const stats = cityStats();
  assert.deepEqual(stats.top.map((item) => [item.name, item.count]), [["上海", 2], ["北京", 1]]);
  assert.equal(stats.total, 3);
  assert.deepEqual(JSON.parse(evaluate('JSON.stringify([...countValues("城市")])')), [["上海", 2], ["北京", 1]]);
  assert.deepEqual(
    JSON.parse(evaluate('JSON.stringify(Object.fromEntries(getCategoryFilterOptions("城市").map((option) => [option.label, option.count])))')),
    { "（空值）": 3, 上海: 2, 北京: 1 }
  );
  assert.deepEqual(
    JSON.parse(evaluate('JSON.stringify(filterRowsByConfig(state.rows, { category: { field: "城市", values: [""] }, numeric: null, date: null }).map((row) => row["金额"]))')),
    ["20", "40", "60"]
  );
  assert.deepEqual(
    JSON.parse(evaluate('JSON.stringify(groupNumericFieldByCategory("金额", "城市").map((group) => [group.name, group.count, group.sum]))')),
    [["（空值）", 3, 120], ["上海", 2, 60], ["北京", 1, 30]]
  );
});

test("a confirmed placeholder is no longer an ID duplicate, and undoing lists it again", () => {
  const scope = commitColumns({ 工号: ["E1", "E2", "N/A", "E3", "N/A"] });
  const listed = () => JSON.parse(scope.evaluate("JSON.stringify(state.idDuplicates[0].values.map((value) => [value.value, value.count]))"));
  assert.deepEqual(listed(), [["N/A", 2]]);
  scope.evaluate('confirmPlaceholder("工号", "N/A")');
  assert.deepEqual(listed(), []);
  assert.equal(scope.evaluate('getFieldProfile("工号").typeKey'), "id");
  scope.evaluate('undoPlaceholderConfirmation("工号", "N/A")');
  assert.deepEqual(listed(), [["N/A", 2]]);
});

test("confirming dashes in a year column changes its date strategy and type", () => {
  const scope = commitColumns({ 年份: YEARS_WITH_DASHES });
  assert.equal(scope.evaluate('getFieldProfile("年份").typeKey'), "category");
  const preview = JSON.parse(scope.evaluate('JSON.stringify(previewPlaceholderConfirmation("年份", ["—"]))'));
  assert.equal(preview.typeKey, "date");
  assert.equal(preview.estimated, false);
  scope.evaluate('confirmPlaceholder("年份", "—")');
  assert.deepEqual(JSON.parse(scope.evaluate('JSON.stringify(getFieldProfile("年份").dateStrategy)')), { order: "none", allowYearOnly: true, granularity: "year", dmyEvidence: 0, mdyEvidence: 0 });
  assert.deepEqual(fieldFigures(scope.evaluate, "年份"), { typeKey: "date", inferredTypeKey: "date", missingCount: 3, uniqueCount: 4, conversionFailureCount: 0 });
});

test("undoing confirmations restores profiles, missing totals, ID duplicates and candidates exactly", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  const snapshot = () => orders.evaluate("JSON.stringify({ profiles: state.profiles, totalMissing: state.totalMissing, idDuplicates: state.idDuplicates, placeholders: state.placeholders, drafts: state.fieldTypeDrafts, confirmations: state.placeholderConfirmations })");
  const before = snapshot();
  orders.evaluate('confirmPlaceholder("金额", "N/A"); confirmPlaceholder("备注", "—"); confirmPlaceholder("备注", "N/A")');
  assert.notEqual(snapshot(), before);
  orders.evaluate('undoPlaceholderConfirmation("备注", "N/A"); undoPlaceholderConfirmation("金额", "N/A"); undoPlaceholderConfirmation("备注", "—")');
  assert.equal(snapshot(), before);
  assert.equal(placeholderMessage(orders), "已撤销「备注」的「—」：空值 27 → 26，分类里重新有「—」这一类。这 1 格重新按原来的文字处理，行数还是 30。");
});

test("manual field types survive confirmations; a type equal to the automatic one is not manual", () => {
  // 9a: inferred numeric, manually category.
  const counts = commitColumns({ 数量: ["1", "2", "N/A", "3", "4", "5"] });
  counts.evaluate('stageFieldTypeChange("数量", "category"); applyFieldConfiguration()');
  assert.deepEqual(manualFieldTypes(counts.evaluate), { 数量: "category" });
  counts.evaluate('confirmPlaceholder("数量", "N/A")');
  assert.deepEqual(fieldFigures(counts.evaluate, "数量"), { typeKey: "category", inferredTypeKey: "numeric", missingCount: 1, uniqueCount: 5, conversionFailureCount: 0 });
  assert.equal(counts.evaluate('state.fieldTypeDrafts["数量"]'), "category");
  assert.equal(counts.evaluate("hasFieldConfigDraftChanges()"), false);

  // 9b: choosing the automatic type is not a manual type, so the field follows the new inference.
  const same = commitColumns({ 金额: AMOUNT_WITH_NA });
  same.evaluate('stageFieldTypeChange("金额", "category"); applyFieldConfiguration()');
  assert.deepEqual(manualFieldTypes(same.evaluate), {});
  same.evaluate('confirmPlaceholder("金额", "N/A")');
  assert.equal(same.evaluate('getFieldProfile("金额").typeKey'), "numeric");

  // 9c: a manual type that the confirmation makes automatic stays manual through the undo.
  const manual = commitColumns({ 金额: AMOUNT_WITH_NA });
  manual.evaluate('stageFieldTypeChange("金额", "numeric"); applyFieldConfiguration()');
  assert.deepEqual(manualFieldTypes(manual.evaluate), { 金额: "numeric" });
  manual.evaluate('confirmPlaceholder("金额", "N/A")');
  assert.deepEqual([manual.evaluate('getFieldProfile("金额").typeKey'), manual.evaluate('getFieldProfile("金额").inferredTypeKey')], ["numeric", "numeric"]);
  assert.equal(manual.evaluate("hasAutomaticTypeDifferences()"), true);
  manual.evaluate('undoPlaceholderConfirmation("金额", "N/A")');
  assert.deepEqual([manual.evaluate('getFieldProfile("金额").typeKey'), manual.evaluate('getFieldProfile("金额").inferredTypeKey')], ["numeric", "category"]);
  manual.evaluate("restoreAutomaticFieldTypes()");
  assert.deepEqual(manualFieldTypes(manual.evaluate), {});
  assert.equal(manual.evaluate('getFieldProfile("金额").typeKey'), "category");
});

test("placeholder actions are refused while the field configuration has unapplied drafts", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  orders.evaluate('stageFieldTypeChange("数量", "category")');
  const snapshot = () => orders.evaluate("JSON.stringify({ profiles: state.profiles, totalMissing: state.totalMissing, placeholders: state.placeholders, confirmations: state.placeholderConfirmations, drafts: state.fieldTypeDrafts })");
  const before = snapshot();
  const refusal = "字段配置有未应用的修改。请先点击“应用字段配置”，再处理像是空值的写法。";
  assert.equal(orders.evaluate('confirmPlaceholder("金额", "N/A")'), false);
  assert.equal(orders.evaluate("confirmAllCommonPlaceholders()"), false);
  assert.equal(orders.evaluate('undoPlaceholderConfirmation("金额", "N/A")'), false);
  assert.equal(snapshot(), before);
  assert.equal(placeholderMessage(orders), refusal);
  assert.equal(orders.elements.get("fieldConfigMessage").textContent, refusal);
});

test("restoring automatic types keeps confirmations; ignored fields keep and can undo theirs", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  orders.evaluate('confirmPlaceholder("金额", "N/A"); confirmPlaceholder("备注", "N/A")');
  orders.evaluate('stageFieldTypeChange("数量", "category"); stageFieldTypeChange("备注", "ignore"); applyFieldConfiguration()');
  assert.deepEqual(manualFieldTypes(orders.evaluate), { 数量: "category", 备注: "ignore" });
  // The ignored field lists its confirmed key only; its other placeholder is not checked.
  assert.deepEqual(placeholderEntries(orders.evaluate).filter(([field]) => field === "备注"), [["备注", "N/A", "common", 1, [29]]]);
  assert.equal(orders.evaluate('state.placeholders.find((entry) => entry.field === "备注").fieldIgnored'), true);
  assert.equal(orders.evaluate('getFieldProfile("备注").missingCount'), 27);

  orders.evaluate('undoPlaceholderConfirmation("备注", "N/A")');
  assert.deepEqual(placeholderEntries(orders.evaluate).filter(([field]) => field === "备注"), []);
  assert.equal(orders.evaluate('getFieldProfile("备注").typeKey'), "ignore");
  assert.equal(orders.evaluate('getFieldProfile("备注").missingCount'), 26);

  orders.evaluate("restoreAutomaticFieldTypes()");
  assert.deepEqual(placeholderConfirmations(orders.evaluate), { 金额: ["N/A"] });
  assert.equal(orders.evaluate("state.totalMissing"), 27);
  assert.equal(fieldFigures(orders.evaluate, "金额").missingCount, 1);
  assert.equal(orders.evaluate('getFieldProfile("数量").typeKey'), "numeric");
});

test("filters keep confirmations, a confirmation clears filters, and a new import clears both records", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  orders.evaluate('confirmPlaceholder("金额", "N/A")');
  orders.evaluate(`(() => {
    state.filters = { category: { field: "大区", values: ["华东"] }, numeric: null, date: null };
    state.filteredRows = filterRowsByConfig(state.rows, state.filters);
    refreshFilteredAnalysis();
  })()`);
  assert.ok(orders.evaluate("getAnalysisRows().length") < 30);
  assert.deepEqual(placeholderConfirmations(orders.evaluate), { 金额: ["N/A"] });
  assert.equal(orders.evaluate("state.totalMissing"), 27);

  orders.evaluate('confirmPlaceholder("备注", "N/A")');
  assert.equal(orders.evaluate("hasActiveFilters()"), false);
  assert.equal(orders.evaluate("getAnalysisRows().length"), 30);
  assert.match(placeholderMessage(orders), /已清除现有筛选条件。$/);

  orders.evaluate('stageFieldTypeChange("数量", "category"); applyFieldConfiguration()');
  commitRows(orders, ["城市"], [{ 城市: "上海" }, { 城市: "N/A" }], "cities.csv");
  assert.deepEqual(placeholderConfirmations(orders.evaluate), {});
  assert.deepEqual(manualFieldTypes(orders.evaluate), {});
  assert.equal(orders.evaluate("hasAutomaticTypeDifferences()"), false);
});

// Previews are arithmetic on the current profile; they must equal a full re-profile.
function comparePlaceholderPreviews(scope) {
  return JSON.parse(scope.evaluate(`JSON.stringify((() => {
    const result = { checked: 0, estimated: 0, mismatches: [] };
    const fields = [...new Set(state.placeholders.filter((entry) => !entry.confirmed && !entry.fieldIgnored).map((entry) => entry.field))];
    fields.forEach((field) => {
      const pending = state.placeholders.filter((entry) => entry.field === field && !entry.confirmed).map((entry) => entry.key);
      const keySets = pending.length > 1 ? [...pending.map((key) => [key]), pending] : [pending];
      keySets.forEach((keys) => {
        const preview = previewPlaceholderConfirmation(field, keys);
        const full = buildColumnProfile(field, state.rows, state.numberFormatHints[field], [...getConfirmedPlaceholderKeys(field), ...keys]);
        const typeKey = state.manualFieldTypes[field] || full.inferredTypeKey;
        const expected = { typeKey, missingCount: full.missingCount, uniqueCount: full.uniqueCount, conversionFailureCount: full.conversionFailuresByType[typeKey] || 0 };
        const actual = { typeKey: preview.typeKey, missingCount: preview.missingCount, uniqueCount: preview.uniqueCount, conversionFailureCount: preview.conversionFailureCount };
        result.checked += 1;
        if (preview.estimated) result.estimated += 1;
        if (JSON.stringify(actual) !== JSON.stringify(expected)) result.mismatches.push({ field, keys, actual, expected });
      });
    });
    return result;
  })())`));
}

test("placeholder previews equal a full re-profile for single keys and whole fields", async () => {
  const cases = [
    { 金额: AMOUNT_WITH_NA },
    { 是否会员: ["有", "无", "有", "无", "有", "有"] },
    { 金额: ["1,200", "-", " - ", "3,400", "560", "(200)"] },
    { 地区: ["NA", "EU", "APAC", "NA", "EU", "LATAM"] },
    { 姓名: ["张三", "Nan", "李四", "王五"] },
    { 备注: ["Ｎ／Ａ", "－－", "好", "N/A", " N/A "] },
    { 折扣: ["N/A", "N/A", "N/A", "N/A", "N/A"] },
    { 工号: ["E1", "E2", "N/A", "E3", "N/A"] },
    { 日期: ["2024-01-01", "2024-01-02", "—", "2024-01-04", "2024-01-05"] },
    { 年份: YEARS_WITH_DASHES },
    { 姓名: ["A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10", "A11", "N/A", "N/A", "A12"] },
    { 金额: ["100", "200", "300", "400", "500", "600", "700", "N/A", "—", "约500"] }
  ];
  let checked = 0;
  let estimated = 0;
  for (const columns of cases) {
    const result = comparePlaceholderPreviews(commitColumns(columns));
    assert.deepEqual(result.mismatches, [], JSON.stringify(columns));
    checked += result.checked;
    estimated += result.estimated;
  }
  // 9a: a manual type wins over the inferred one in the preview too.
  const manual = commitColumns({ 数量: ["1", "2", "N/A", "3", "4", "5"] });
  manual.evaluate('stageFieldTypeChange("数量", "category"); applyFieldConfiguration()');
  assert.deepEqual(comparePlaceholderPreviews(manual).mismatches, []);
  // Previews start from what is already confirmed: one of two keys tips the amounts to numeric.
  const combined = commitColumns({ 金额: ["100", "200", "300", "400", "500", "600", "700", "N/A", "—", "约500"] });
  assert.equal(combined.evaluate('previewPlaceholderConfirmation("金额", ["—"]).typeKey'), "category");
  combined.evaluate('confirmPlaceholder("金额", "N/A")');
  assert.equal(combined.evaluate('previewPlaceholderConfirmation("金额", ["—"]).typeKey'), "numeric");
  assert.deepEqual(comparePlaceholderPreviews(combined).mismatches, []);

  const orders = await importStage3Fixture("订单表.csv");
  const orderResult = comparePlaceholderPreviews(orders);
  assert.deepEqual(orderResult.mismatches, []);
  // Four fields: 备注 has two keys, so its whole-field preview is checked as well.
  assert.equal(orderResult.checked, 6);
  // Every key set but the year column's took the arithmetic path.
  assert.ok(checked >= 15);
  assert.equal(checked - estimated, 1);
});

test("confirmations never change the original rows, the data preview, CSV exports or duplicate rows", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  const { evaluate } = orders;
  evaluate("globalThis.__rowsBefore = state.rows");
  const snapshot = () => evaluate(`JSON.stringify({
    rows: state.rows,
    duplicateRows: state.duplicateRows,
    csv: buildCsvExportContent(state.rows, state.fields),
    filtered: buildCsvExportContent(getProcessedExportRows("filtered"), state.fields)
  })`);
  const before = snapshot();
  evaluate("confirmAllCommonPlaceholders()");
  assert.equal(evaluate("state.placeholders.every((entry) => entry.confirmed)"), true);
  assert.equal(snapshot(), before);
  assert.equal(evaluate("state.rows === __rowsBefore"), true);
  // The preview shows the first ten analysis rows; the last ones hold the 备注 placeholders.
  evaluate("state.filteredRows = state.rows.slice(27); renderPreview()");
  const preview = orders.elements.get("previewTable").innerHTML;
  assert.ok(preview.includes("<td>N/A</td>"));
  assert.ok(preview.includes("<td>—</td>"));
});

test("after confirming every order placeholder, no analysis output shows N/A or a placeholder dash", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  const charts = [];
  orders.context.Chart = orders.context.window.Chart = class {
    constructor(canvas, config) { charts.push(config); }
    destroy() {}
  };
  const renderEverything = () => {
    charts.length = 0;
    orders.evaluate(`
      dom.categoryFilterField.value = "备注"; updateCategoryFilterField();
      dom.categoryChartField.value = "备注";
      dom.v2MetricField.value = "金额"; dom.v2GroupField.value = "备注"; dom.v2AggregateMethod.value = "sum"; dom.v2DateField.value = "下单日期";
      renderCharts(); renderV2Analysis();
    `);
    const skipped = new Set(["previewTable", "placeholderDetails", "placeholderMessage"]);
    // The only dash the app writes itself here: the conversion column of types without conversion.
    const html = [...orders.elements].filter(([id]) => !skipped.has(id))
      .map(([id, element]) => `${id}:${element.innerHTML}${element.textContent}`).join("\n")
      .replace(/当前类型无需转换">—<\/span>/g, "");
    return `${html}\n${JSON.stringify(charts.map((config) => config.data.labels))}`;
  };
  const before = renderEverything();
  assert.ok(before.includes("N/A") && before.includes("—"));
  orders.evaluate("confirmAllCommonPlaceholders()");
  const after = renderEverything();
  assert.equal(after.includes("N/A"), false, after.split("\n").filter((line) => line.includes("N/A")).join("\n").slice(0, 600));
  assert.equal(after.includes("—"), false, after.split("\n").filter((line) => line.includes("—")).join("\n").slice(0, 600));
});

test("applying the configuration for another field keeps a manual type set before a confirmation", () => {
  const scope = commitColumns({ 金额: AMOUNT_WITH_NA, 备注: ["好", "差", "好", "差", "好", "中"] });
  scope.evaluate('stageFieldTypeChange("金额", "numeric"); applyFieldConfiguration()');
  scope.evaluate('confirmPlaceholder("金额", "N/A")');
  // The confirmation makes numeric automatic for 金额; applying 备注 must not drop its manual record.
  scope.evaluate('stageFieldTypeChange("备注", "ignore"); applyFieldConfiguration()');
  assert.deepEqual(manualFieldTypes(scope.evaluate), { 金额: "numeric", 备注: "ignore" });
  scope.evaluate('undoPlaceholderConfirmation("金额", "N/A")');
  assert.deepEqual([scope.evaluate('getFieldProfile("金额").typeKey'), scope.evaluate('getFieldProfile("金额").inferredTypeKey')], ["numeric", "category"]);
  assert.deepEqual(manualFieldTypes(scope.evaluate), { 金额: "numeric", 备注: "ignore" });
});

// Placeholder candidates on the quality page (#placeholderDetails) and in reports.
const placeholderDetails = (scope) => scope.elements.get("placeholderDetails").innerHTML;
const PLACEHOLDER_MAYBE_NOTE_TEXT = "「无」「暂无」「-」「/」「NA」也可能是有效值（例如“有/无”选项、北美地区），请确认这一列里它确实表示没有数据。";

test("the quality page has a static placeholder subsection before the ID duplicates", () => {
  const html = indexHtmlSource();
  const quality = html.slice(html.indexOf('id="qualitySection"'), html.indexOf('id="filterSection"'));
  const placeholderSection = quality.slice(quality.indexOf('<h3 class="subsection-title">像是空值的写法</h3>'), quality.indexOf('<h3 class="subsection-title">编号重复</h3>'));
  assertInOrder(quality, ['<h3 class="subsection-title">像是空值的写法</h3>', '<h3 class="subsection-title">编号重复</h3>']);
  assert.ok(placeholderSection.includes("这些写法常被用来表示“没有数据”。按空值处理后：不再算作认不出的数字或日期，也不参与统计、Top 10、分组、筛选选项和编号重复检查。原始数据、数据预览和导出的 CSV 不变。"));
  // The message sits outside the container that renderPlaceholderDetails() rewrites.
  assert.match(placeholderSection, /<div id="placeholderDetails" class="placeholder-details"><\/div>/);
  assert.match(placeholderSection, /<p id="placeholderMessage" class="placeholder-message" role="status" tabindex="-1"><\/p>/);
  const { evaluate } = createScriptContext();
  assert.equal(evaluate('["placeholderDetails", "placeholderMessage"].every((id) => Object.keys(dom).includes(id))'), true);
  assert.match(fs.readFileSync(path.join(__dirname, "..", "script.js"), "utf8"), /dom\.placeholderDetails\.addEventListener\("click", handlePlaceholderClick\);/);
});

test("the placeholder pill counts fields with unconfirmed keys after the ID duplicate pill", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  assert.deepEqual(qualityPill(orders.elements, "像是空值的写法"), { tone: "warning", text: "像是空值的写法：4 个字段" });
  orders.evaluate('confirmPlaceholder("金额", "N/A")');
  assert.deepEqual(qualityPill(orders.elements, "像是空值的写法"), { tone: "warning", text: "像是空值的写法：3 个字段" });
  orders.evaluate("confirmAllCommonPlaceholders()");
  assert.deepEqual(qualityPill(orders.elements, "像是空值的写法"), { tone: "success", text: "像是空值的写法：0 个字段" });
  const grades = await importStage3Fixture("成绩表.csv");
  assert.deepEqual(qualityPill(grades.elements, "像是空值的写法"), { tone: "success", text: "像是空值的写法：0 个字段" });
});

test("placeholder details show each key with its rows and effect, one button each and a whole-field line", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  const html = placeholderDetails(orders);
  for (const text of [
    "「下单日期」：「—」1 个（第 17 行）。按空值处理后，空值 0 → 1，类型仍为日期。",
    "「数量」：「—」1 个（第 12 行）。按空值处理后，空值 0 → 1，类型仍为数值。",
    "「金额」：「N/A」1 个（第 15 行）。按空值处理后，空值 0 → 1，类型仍为数值。",
    "「备注」：「N/A」1 个（第 29 行）。按空值处理后，空值 26 → 27，类型仍为分类。",
    "「备注」：「—」1 个（第 30 行）。按空值处理后，空值 26 → 27，类型仍为分类。",
    "「备注」全部按空值处理后：空值 26 → 28，类型仍为分类。"
  ]) assert.ok(html.includes(text), text);
  assert.deepEqual([...html.matchAll(/data-placeholder-index="(\d+)"[^>]*>按空值处理</g)].map((match) => match[1]), ["0", "1", "2", "3", "4"]);
  assert.ok(html.includes("常见写法全部按空值处理（5 处）"));
  assert.doesNotMatch(html, /按空值处理后类型会变/);
  assert.doesNotMatch(html, /也可能是有效值/);

  const grades = await importStage3Fixture("成绩表.csv");
  assert.ok(placeholderDetails(grades).includes("没有发现像是空值的写法。"));
});

test("placeholder details name raw texts, maybe-valid keys, type changes and missing row numbers", () => {
  const amount = commitColumns({ 金额: ["100", "200", "N/A", "约500", "300", "400", "500"], 年份: YEARS_WITH_DASHES });
  const amountHtml = placeholderDetails(amount);
  assert.ok(amountHtml.includes("「金额」：「N/A」1 个（行号：—）。按空值处理后，空值 0 → 1，「金额」会识别为数值（现在是分类）。"));
  assert.ok(amountHtml.includes("「年份」：「—」3 个（行号：—）。按空值处理后，空值 0 → 3，「年份」会识别为日期（现在是分类）。"));
  assert.ok(amountHtml.includes("常见写法全部按空值处理（2 处）"));
  assert.ok(amountHtml.includes("按空值处理后类型会变：「金额」「年份」"));

  const raw = commitColumns({ 备注: ["Ｎ／Ａ", "－－", "好", "N/A", " N/A "] });
  const rawHtml = placeholderDetails(raw);
  assert.ok(rawHtml.includes("「备注」：「N/A」（原文：Ｎ／Ａ、N/A）3 个（行号：—）。按空值处理后，空值 0 → 3，类型仍为分类。"));
  assert.ok(rawHtml.includes("「备注」：「--」（原文：－－）1 个（行号：—）。按空值处理后，空值 0 → 1，类型仍为分类。"));
  assert.ok(rawHtml.includes("「备注」全部按空值处理后：空值 0 → 4，类型仍为分类。"));

  const maybe = commitColumns({ 是否会员: ["有", "无", "有", "无", "有", "有"], 金额: ["1,200", "-", " - ", "3,400", "560", "(200)"] });
  const maybeHtml = placeholderDetails(maybe);
  assert.ok(maybeHtml.includes("「是否会员」：「无」2 个（行号：—）。按空值处理后，空值 0 → 2，类型仍为分类。"));
  assert.equal(maybeHtml.split(PLACEHOLDER_MAYBE_NOTE_TEXT).length - 1, 2);
  assert.ok(maybeHtml.includes(`${PLACEHOLDER_MAYBE_NOTE_TEXT}在会计格式里，「-」通常表示 0。`));
  assert.equal(maybeHtml.split("在会计格式里").length - 1, 1);
  // No common key, so no confirm-all button.
  assert.doesNotMatch(maybeHtml, /常见写法全部按空值处理/);

  const escaped = commitColumns({ "<b>x</b>": ["N/A", "1", "2"] });
  assert.ok(placeholderDetails(escaped).includes("「&lt;b&gt;x&lt;/b&gt;」：「N/A」1 个"));
  assert.doesNotMatch(placeholderDetails(escaped), /<b>x<\/b>/);
});

test("confirmed keys show an undo button, ignored fields are marked, and more than ten rows end with 等", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  orders.evaluate('confirmPlaceholder("金额", "N/A"); confirmPlaceholder("备注", "N/A")');
  orders.evaluate('stageFieldTypeChange("备注", "ignore"); applyFieldConfiguration()');
  const html = placeholderDetails(orders);
  assert.ok(html.includes("「金额」：「N/A」已按空值处理（1 个）。"));
  assert.ok(html.includes("「备注」（字段已忽略）：「N/A」已按空值处理（1 个）。"));
  assert.deepEqual([...html.matchAll(/data-placeholder-index="(\d+)"[^>]*>撤销</g)].map((match) => match[1]), ["2", "3"]);
  assert.ok(html.includes("常见写法全部按空值处理（2 处）"));

  const many = await importCsvFile(`编号,备注\r\n${Array.from({ length: 12 }, (_, index) => `A${index},N/A`).join("\r\n")}\r\nA99,好\r\n`);
  assert.ok(placeholderDetails(many).includes("「备注」：「N/A」12 个（第 2、3、4、5、6、7、8、9、10、11 行等）。"));
});

test("placeholder buttons confirm, undo and confirm all through one delegated click handler", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  const click = (dataset) => {
    put(orders.context, "__placeholderButton", dataset ? { dataset } : null);
    orders.evaluate("handlePlaceholderClick({ target: { closest: () => __placeholderButton } })");
  };
  click(null);
  assert.deepEqual(placeholderConfirmations(orders.evaluate), {});
  click({ placeholderIndex: "2" });
  assert.deepEqual(placeholderConfirmations(orders.evaluate), { 金额: ["N/A"] });
  click({ placeholderIndex: "2" });
  assert.deepEqual(placeholderConfirmations(orders.evaluate), {});
  click({ placeholderAll: "" });
  assert.equal(orders.evaluate("state.placeholders.every((entry) => entry.confirmed)"), true);
});

test("reports list confirmed and pending placeholders after the missing values and count pending fields", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  const before = reportOf(orders);
  assert.deepEqual(before.report.quality.placeholders.confirmed, []);
  assert.deepEqual(before.report.quality.placeholders.pending[2], { field: "金额", key: "N/A", tier: "common", rawValues: ["N/A"], count: 1, rowNumbers: [15], effect: "按空值处理后，空值 0 → 1，类型仍为数值" });
  assert.equal(before.report.quality.placeholders.pending.length, 5);
  assert.ok(before.html.includes("；编号重复：<strong>1</strong> 个字段；像是空值的写法：<strong>4</strong> 个字段。</p>"));
  const htmlSection = before.html.slice(before.html.indexOf("<h3>像是空值的写法</h3>"), before.html.indexOf("<h3>偏大或偏小的值</h3>"));
  assertInOrder(before.html, ["<h3>各字段空值情况</h3>", "<h3>像是空值的写法</h3>"]);
  assert.match(htmlSection, /<th>字段<\/th><th>写法<\/th><th>档位<\/th><th>出现次数<\/th><th>文件行号<\/th><th>按空值处理后的影响<\/th>/);
  assert.ok(htmlSection.includes("<td>金额</td><td>「N/A」</td><td>常见</td><td>1</td><td>15</td><td>按空值处理后，空值 0 → 1，类型仍为数值</td>"));
  assertInOrder(before.html, ["<h3>像是空值的写法</h3>", "<h3>偏大或偏小的值</h3>"]);

  orders.evaluate('confirmPlaceholder("金额", "N/A"); confirmPlaceholder("备注", "N/A")');
  orders.evaluate('stageFieldTypeChange("备注", "ignore"); applyFieldConfiguration()');
  const after = reportOf(orders);
  assert.deepEqual(after.report.quality.placeholders.confirmed, [
    { field: "金额", key: "N/A", rawValues: ["N/A"], count: 1, fieldIgnored: false },
    { field: "备注", key: "N/A", rawValues: ["N/A"], count: 1, fieldIgnored: true }
  ]);
  assert.ok(after.html.includes("；像是空值的写法：<strong>2</strong> 个字段。</p>"));
  assert.ok(after.html.includes("<td>金额</td><td>「N/A」</td><td>1</td><td>已按空值处理</td>"));
  assert.ok(after.html.includes("<td>备注</td><td>「N/A」</td><td>1</td><td>已按空值处理（字段已忽略）</td>"));
});

test("report placeholder summaries appear only with candidates, and legacy snapshots leave the section out", async () => {
  for (const scope of [await importStage3Fixture("花名册.xlsx"), await importStage3Fixture("成绩表.csv"), await importCsvFile(behaviorCsv(), "sample-user-behavior.csv")]) {
    const { report, html } = reportOf(scope);
    assert.deepEqual(report.quality.placeholders, { confirmed: [], pending: [] });
    assert.doesNotMatch(html, /像是空值的写法：/);
    assert.ok(html.includes("<p>没有发现像是空值的写法。</p>"));
  }
  const scope = createScriptContext();
  assert.equal(scope.evaluate("normalizeReportSnapshot({}).quality.placeholders"), null);
  assert.equal(scope.evaluate("normalizeReportSnapshot({ quality: { placeholders: { confirmed: [] } } }).quality.placeholders"), null);
  put(scope.context, "__legacyPlaceholders", { quality: { totalMissing: 1, duplicateRows: 0, parseWarningCount: 0 } });
  assert.doesNotMatch(scope.evaluate("buildHtmlReport(__legacyPlaceholders, [])"), /<h3>像是空值的写法<\/h3>|；像是空值的写法：/);
});

test("report placeholder sections escape field names", () => {
  const scope = commitColumns({ "<b>备注</b>|x": ["N/A", "好", "-", "好"], "-前缀": ["—", "甲", "乙", "丙"] });
  const { html } = reportOf(scope);
  const section = html.slice(html.indexOf("<h3>像是空值的写法</h3>"), html.indexOf("<h3>偏大或偏小的值</h3>"));
  assert.ok(section.includes("<td>&lt;b&gt;备注&lt;/b&gt;|x</td><td>「N/A」</td>"));
  assert.ok(section.includes("<td>-前缀</td><td>「—」</td>"));
  assert.doesNotMatch(html, /<b>备注<\/b>/);
  assert.ok(section.includes(`<p>说明：${PLACEHOLDER_MAYBE_NOTE_TEXT}在会计格式里，「-」通常表示 0。</p>`));
});

// v2.5.0 number strategies: each field reads units, 万/亿 and full-width digits
// through its own profile.numberStrategy.
const numberProfile = (evaluate, field) => JSON.parse(evaluate(`JSON.stringify((({ typeKey, numberStrategy, numberReadings, conversionFailureCount, conversionFailureSamples }) => ({ typeKey, numberStrategy, numberReadings, conversionFailureCount, conversionFailureSamples }))(getFieldProfile(${JSON.stringify(field)})))`));

test("the order fixture reads 2200元 and 3500元 as yuan and keeps only N/A as a format problem", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  assert.deepEqual(numberProfile(orders.evaluate, "金额"), {
    typeKey: "numeric",
    numberStrategy: { unit: "元", unitConflict: false, units: { 元: 2 }, allowMagnitude: false, fullwidth: false },
    numberReadings: { unitCount: 2, unitSamples: ["2200元", "3500元"], magnitudeCount: 0, fullwidthCount: 0 },
    conversionFailureCount: 1,
    conversionFailureSamples: ["N/A"]
  });
  assert.deepEqual(JSON.parse(orders.evaluate('JSON.stringify(["2200元", "3500元", "￥15,007", "￥8,250", "N/A"].map((value) => parseFieldNumber("金额", value)))')), [2200, 3500, 15007, 8250, null]);
  assert.deepEqual(qualityPill(orders.elements, "认不出的数字或日期"), { tone: "warning", text: "认不出的数字或日期：3 个值" });
  const amountStats = JSON.parse(orders.evaluate('JSON.stringify(state.numericStats.find((item) => item.field === "金额"))'));
  assert.deepEqual(amountStats, {
    field: "金额", count: 29, invalidCount: 1, mean: 3293.7241379310344, median: 1700, min: 704, max: 15007, std: 3305.713026497504,
    q1: 1050, q3: 4600, lowerBound: -4275, upperBound: 9925, outlierCount: 1
  });
  assert.deepEqual(JSON.parse(orders.evaluate('JSON.stringify(state.qualityNumericStats.find((item) => item.field === "金额"))')), amountStats);
});

test("the other stage3 fixtures keep their types and format problem totals", async () => {
  const expected = {
    "成绩表.csv": [["id", "id", "category", "numeric", "numeric", "numeric", "numeric", "numeric"], 3],
    "成绩表-标题无逗号.csv": [["id", "id", "category", "numeric", "numeric", "numeric", "numeric", "numeric"], 3],
    "花名册.xlsx": [["id", "category", "category", "id", "date", "numeric", "id"], 2],
    "门店报表.xlsx": [["id", "id", "category", "date", "numeric", "numeric", "numeric", "category"], 0]
  };
  for (const [fileName, [types, formatProblems]] of Object.entries(expected)) {
    const scope = await importStage3Fixture(fileName);
    assert.deepEqual(importedSummary(scope.evaluate).types, types, fileName);
    assert.equal(qualityPill(scope.elements, "认不出的数字或日期").text, `认不出的数字或日期：${formatProblems} 个值`, fileName);
    assert.equal(scope.evaluate("state.profiles.some((profile) => hasNumberStrategySwitch(profile.numberStrategy))"), false, fileName);
  }
});

test("15 distinct yuan amounts are numeric, while 4G / 5G and 8G-style sizes keep their types", () => {
  const yuan = commitColumns({ 金额: Array.from({ length: 15 }, (_, index) => `${1000 + index * 37}元`) });
  assert.equal(fieldFigures(yuan.evaluate, "金额").typeKey, "numeric");
  const network = commitColumns({ 网络: ["4G", "5G", "5G", "4G", "5G", "3G", "5G", "4G"] });
  assert.equal(fieldFigures(network.evaluate, "网络").typeKey, "category");
  const memory = commitColumns({ 内存: Array.from({ length: 13 }, (_, index) => `${(index + 1) * 4}G`) });
  assert.equal(fieldFigures(memory.evaluate, "内存").typeKey, "id");
});

test("万 amounts are exact integers, so small ranges still get one bar per value", () => {
  const scope = createChartContext();
  commitRows(scope, ["件数"], ["0.0001万", "0.0003万", "0.0005万", "0.0007万", "0.001万", "0.0002万", "0.0007万"].map((value) => ({ 件数: value })), "magnitude.csv");
  assert.equal(scope.evaluate('getFieldProfile("件数").typeKey'), "numeric");
  assert.deepEqual(drawDistribution(scope, "件数"), {
    labels: ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"],
    data: [1, 1, 1, 0, 1, 0, 2, 0, 0, 1]
  });
  put(scope.context, "__magnitudes", ["0.07万", "0.3万", "1万", "2万", "0.5万", "1.15亿"]);
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(__magnitudes.map((value) => toNumber(value, buildNumberStrategy(__magnitudes))))")), [700, 3000, 10000, 20000, 5000, 115000000]);
});

test("placeholder previews equal a full re-profile on unit columns and confirmations keep each number strategy", () => {
  const scope = commitColumns({
    金额: ["2200元", "N/A", "1.2万元", "—", "３５００元", "-", "800元", "n/a", "NULL", "约500"],
    重量: ["5kg", "500千克", "N/A", "2kg", "—", "1kg", "-", "3kg", "无", "4kg"],
    数量: ["１２", "—", "３４", "N/A", "５", "6", "－", "--", "7", "8"]
  });
  const strategies = () => JSON.parse(scope.evaluate("JSON.stringify(state.profiles.map((profile) => profile.numberStrategy))"));
  const before = strategies();
  assert.deepEqual(before, [
    { unit: "元", unitConflict: false, units: { 元: 4 }, allowMagnitude: true, fullwidth: true },
    { unit: "", unitConflict: true, units: { kg: 5, 千克: 1 }, allowMagnitude: false, fullwidth: false },
    { unit: "", unitConflict: false, units: {}, allowMagnitude: false, fullwidth: true }
  ]);
  const result = comparePlaceholderPreviews(scope);
  assert.ok(result.checked >= 10);
  assert.deepEqual(result.mismatches, []);
  scope.evaluate("confirmAllCommonPlaceholders()");
  assert.deepEqual(strategies(), before);
  scope.evaluate('confirmPlaceholder("金额", "-"); confirmPlaceholder("重量", "无")');
  assert.deepEqual(strategies(), before);
  scope.evaluate('undoPlaceholderConfirmation("金额", "N/A"); undoPlaceholderConfirmation("数量", "—")');
  assert.deepEqual(strategies(), before);
  assert.deepEqual(comparePlaceholderPreviews(scope).mismatches, []);
});

test("the numeric filter range and the filter itself read units the same way", async () => {
  const scope = commitColumns({ 金额: ["1200", "800", "3500元", "2200元", "950"], 编号: ["a", "b", "c", "d", "e"] });
  scope.evaluate('dom.numericFilterField.value = "金额"; updateNumericFilterField()');
  assert.equal(scope.elements.get("numericFilterRange").textContent, "800 至 3,500");
  applyFilterState(scope, { numeric: { field: "金额", min: 2000, max: 3500 } });
  assert.deepEqual(JSON.parse(scope.evaluate('JSON.stringify(getAnalysisRows().map((row) => row["金额"]))')), ["3500元", "2200元"]);

  const orders = await importStage3Fixture("订单表.csv");
  applyFilterState(orders, { numeric: { field: "金额", min: 3000, max: 3500 } });
  assert.deepEqual(JSON.parse(orders.evaluate('JSON.stringify(getAnalysisRows().map((row) => [getSourceRowNumber(row), row["金额"]]))')), [[21, "3500元"]]);
});

test("manual type changes on a unit column use its strategy only while numeric", () => {
  const scope = commitColumns({ 金额: ["2200元", "3500元", "约500", "N/A", "100元", "800元", "待定"] });
  const snapshot = () => JSON.parse(scope.evaluate(`JSON.stringify((() => {
    const profile = getFieldProfile("金额");
    return [profile.typeKey, profile.conversionFailureCount, profile.conversionFailureSamples, state.numericStats.map((item) => [item.count, item.invalidCount, item.mean])];
  })())`));
  assert.deepEqual(snapshot(), ["category", 0, [], []]);
  scope.evaluate('stageFieldTypeChange("金额", "numeric"); applyFieldConfiguration()');
  assert.deepEqual(snapshot(), ["numeric", 3, ["约500", "N/A", "待定"], [[4, 3, 1650]]]);
  scope.evaluate('confirmPlaceholder("金额", "N/A")');
  assert.deepEqual(snapshot(), ["numeric", 2, ["约500", "待定"], [[4, 2, 1650]]]);
  scope.evaluate('undoPlaceholderConfirmation("金额", "N/A")');
  assert.deepEqual(snapshot(), ["numeric", 3, ["约500", "N/A", "待定"], [[4, 3, 1650]]]);
  scope.evaluate('stageFieldTypeChange("金额", "id"); applyFieldConfiguration()');
  assert.deepEqual(snapshot(), ["id", 0, [], []]);
  assert.equal(scope.evaluate("state.idDuplicates[0].duplicateValueCount"), 0);
  scope.evaluate('stageFieldTypeChange("金额", "date"); applyFieldConfiguration()');
  assert.deepEqual(snapshot(), ["date", 7, ["2200元", "3500元", "约500"], []]);
  scope.evaluate("restoreAutomaticFieldTypes()");
  assert.deepEqual(snapshot(), ["category", 0, [], []]);
});

test("header detection keeps reading 5kg as text while the analysis reads it as a number", async () => {
  const headerless = await importCsvFile("库存表\r\n苹果,5kg\r\n香蕉,6kg\r\n梨,7kg\r\n桃,8kg\r\n", "headerless.csv");
  assert.deepEqual(JSON.parse(headerless.evaluate("JSON.stringify([state.headerDetection.headerRowNumber, state.fields, state.rows.length])")), [2, ["苹果", "5kg"], 3]);
  assert.equal(headerless.evaluate('getFieldProfile("5kg").typeKey'), "numeric");
  const titled = await importCsvFile("库存表\r\n水果,重量\r\n苹果,5kg\r\n香蕉,6kg\r\n梨,7kg\r\n", "titled.csv");
  assert.deepEqual(JSON.parse(titled.evaluate("JSON.stringify([state.headerDetection.headerRowNumber, state.fields, state.rows.length])")), [2, ["水果", "重量"], 3]);
  assert.equal(titled.evaluate('getFieldProfile("重量").typeKey'), "numeric");
});

test("CSV export still neutralizes -2200元 in a yuan column", () => {
  const scope = commitColumns({ 金额: ["-2200元", "1200元", "800元"] });
  scope.context.window.Papa = Papa;
  assert.equal(scope.evaluate('getFieldProfile("金额").typeKey'), "numeric");
  assert.equal(scope.evaluate("buildCsvExportContent(state.rows, state.fields)"), "金额\r\n'-2200元\r\n1200元\r\n800元");
});

test("the main analysis sums yuan amounts per region and per day", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  const result = JSON.parse(orders.evaluate(`JSON.stringify((() => {
    const regions = Object.fromEntries(groupNumericFieldByCategory("金额", "大区").map((group) => [group.name, group.sum]));
    const trend = buildMetricDateTrend("下单日期", "金额", "sum");
    const at = (label) => trend.values[trend.labels.indexOf(label)];
    return { regions, trend: [at("2024-05-08"), at("2024-05-20")] };
  })())`));
  assert.equal(result.regions["华南"], 6025);
  assert.deepEqual(result.trend, [3560, 4204]);
});

// Numeric-looking category fields: placeholder spellings (confirmed or not) and
// missing values are left out of the count, the share and the samples.
const numericLooking = (evaluate) => JSON.parse(evaluate("JSON.stringify(state.numericLooking)"));

test("the reported amounts with N/A and 约500 are not numeric-looking: without N/A, 4 of 5 is the numeric threshold", () => {
  const scope = commitColumns({ 金额: AMOUNT_WITH_NA });
  assert.equal(fieldFigures(scope.evaluate, "金额").typeKey, "category");
  assert.deepEqual(JSON.parse(scope.evaluate('JSON.stringify(getFieldProfile("金额").numericLooking)')), { valueCount: 5, numericCount: 4, samples: ["约500"] });
  assert.deepEqual(numericLooking(scope.evaluate), []);
});

test("floors 1, 2, 3, B1, B2 are numeric-looking (an accepted false hint)", () => {
  const scope = commitColumns({ 楼层: ["1", "2", "3", "B1", "B2"] });
  assert.deepEqual(numericLooking(scope.evaluate), [{ field: "楼层", valueCount: 5, numericCount: 3, ratio: 0.6, samples: ["B1", "B2"] }]);
});

test("a year column with dashes is not numeric-looking", () => {
  const scope = commitColumns({ 年份: YEARS_WITH_DASHES });
  assert.equal(fieldFigures(scope.evaluate, "年份").typeKey, "category");
  assert.deepEqual(numericLooking(scope.evaluate), []);
});

test("numeric-looking fields list up to three unconvertible samples and never a placeholder spelling", () => {
  const scope = commitColumns({
    评分: ["1", "N/A", "2", "5", "—", "4", "非常愿意", "一般", "3", "-", "", "Ｎ／Ａ"],
    金额: ["1200", "800", "面议", "1500", "面议", "900", "700", "面议", "650", "1100", "N/A", "—"],
    报价: ["1200", "待定", "800", "待定", "1500", "待定", "NULL", "", "无", "--", "", "n/a"],
    价格: ["1200元", "800", "约500", "待定", "1500元", "900", "N/A", "—", "", "", "", ""]
  });
  assert.deepEqual(numericLooking(scope.evaluate), [
    { field: "评分", valueCount: 7, numericCount: 5, ratio: 5 / 7, samples: ["非常愿意", "一般"] },
    { field: "金额", valueCount: 10, numericCount: 7, ratio: 0.7, samples: ["面议"] },
    { field: "报价", valueCount: 6, numericCount: 3, ratio: 0.5, samples: ["待定"] },
    { field: "价格", valueCount: 6, numericCount: 4, ratio: 4 / 6, samples: ["约500", "待定"] }
  ]);
  const placeholderSamples = JSON.parse(scope.evaluate("JSON.stringify(state.numericLooking.flatMap((item) => item.samples).filter((sample) => placeholderKey(sample)))"));
  assert.deepEqual(placeholderSamples, []);
});

test("a field with only four values is not numeric-looking", () => {
  const scope = commitColumns({ 金额: ["1200", "800", "约500", "1500"] });
  assert.equal(fieldFigures(scope.evaluate, "金额").typeKey, "category");
  assert.deepEqual(numericLooking(scope.evaluate), []);
});

test("numeric-looking hints follow applied types and confirmations, and filters keep them", () => {
  const scope = commitColumns({
    楼层: ["1", "2", "3", "B1", "B2", "1", "2", "3", "4", "B1", "1", "2"],
    房号: ["101", "102", "103", "104", "105", "106", "107", "A1", "B2", "C3", "D4", "N/A"]
  });
  const fields = () => numericLooking(scope.evaluate).map((item) => item.field);
  // 房号 starts as an ID: 12 distinct values, 7 of 12 numeric.
  assert.equal(fieldFigures(scope.evaluate, "房号").typeKey, "id");
  assert.deepEqual(fields(), ["楼层"]);

  scope.evaluate('stageFieldTypeChange("楼层", "numeric"); applyFieldConfiguration()');
  assert.deepEqual(fields(), []);
  scope.evaluate('stageFieldTypeChange("楼层", "ignore"); applyFieldConfiguration()');
  assert.deepEqual(fields(), []);
  scope.evaluate("restoreAutomaticFieldTypes()");
  assert.deepEqual(fields(), ["楼层"]);

  // Confirming N/A leaves 11 values, so 房号 is no longer an ID but a category.
  scope.evaluate('confirmPlaceholder("房号", "N/A")');
  assert.equal(fieldFigures(scope.evaluate, "房号").typeKey, "category");
  assert.deepEqual(numericLooking(scope.evaluate)[1], { field: "房号", valueCount: 11, numericCount: 7, ratio: 7 / 11, samples: ["A1", "B2", "C3"] });
  scope.evaluate('undoPlaceholderConfirmation("房号", "N/A")');
  assert.deepEqual(fields(), ["楼层"]);

  const before = scope.evaluate("state.numericLooking");
  applyFilterState(scope, { category: { field: "楼层", values: ["1"] } });
  assert.equal(scope.evaluate("state.numericLooking"), before);
  assert.deepEqual(fields(), ["楼层"]);
});

test("numeric-looking hints are cleared on reset and replaced on every import", () => {
  const scope = commitColumns({ 楼层: ["1", "2", "3", "B1", "B2"] });
  assert.equal(numericLooking(scope.evaluate).length, 1);
  commitRows(scope, ["金额"], ["1", "2", "3", "4", "5"].map((value) => ({ 金额: value })), "next.csv");
  assert.deepEqual(numericLooking(scope.evaluate), []);
  commitRows(scope, ["楼层"], ["1", "2", "3", "B1", "B2"].map((value) => ({ 楼层: value })), "again.csv");
  scope.evaluate("resetAnalysisState()");
  assert.deepEqual(numericLooking(scope.evaluate), []);
});

// Units in labels, the two number subsections of the quality page and reports.
const UNIT_SECTION_COLUMNS = {
  重量: ["5kg", "500千克", "2kg", "1kg", "3kg", "4kg"],
  楼层: ["1", "2", "3", "B1", "B2", "4"],
  销量: ["1.2万", "8000", "３５００", "9500", "1万", "6000"]
};
const WEIGHT_CONFLICT_TEXT = "「重量」：同时出现 kg（5 个）和 千克（1 个），没有按单位读取，这些值不能读成数字。";
const SALES_CONVERSION_TEXT = "「销量」：按万/亿换算了 2 个值；1 个全角数字已按半角读取";
const FLOOR_HINT_TEXT = "「楼层」：去掉空值和像是空值的写法后，6 个值里有 4 个能读成数字（66.7%），其余的如 B1、B2。如果这一列应该是数字，请修正这些值，或在字段配置里把类型改成数值。";
const ORDER_AMOUNT_CONVERSION_TEXT = "「金额」：按\"元\"读取了 2 个值（2200元、3500元）";

function withFakeCharts(scope) {
  const charts = [];
  scope.context.Chart = scope.context.window.Chart = class {
    constructor(canvas, config) { charts.push(config); }
    destroy() {}
  };
  return charts;
}

test("field labels add the unit unless the name already ends with it", () => {
  const { context, evaluate } = createScriptContext();
  context.__labelCases = [
    ["金额", "元", "金额（元）"], ["金额（元）", "元", "金额（元）"], ["金额(元)", "元", "金额(元)"], ["金额/元", "元", "金额/元"],
    ["金额【元】", "元", "金额【元】"], ["金额[元]", "元", "金额[元]"], ["金额 元", "元", "金额 元"], ["金额-元", "元", "金额-元"], ["金额元", "元", "金额元"],
    ["weight_kg", "kg", "weight_kg"], ["重量kg", "kg", "重量kg"], ["重量KG", "kg", "重量KG"], ["重量（ＫＧ）", "kg", "重量（ＫＧ）"],
    ["bag", "kg", "bag（kg）"], ["backlog", "kg", "backlog（kg）"], ["元旦活动数", "元", "元旦活动数（元）"], ["数量", "", "数量"]
  ];
  assert.deepEqual(
    JSON.parse(evaluate("JSON.stringify(__labelCases.map(([field, unit]) => formatFieldLabelWithUnit(field, unit)))")),
    context.__labelCases.map(([, , label]) => label)
  );
});

test("statistics, charts and custom analysis label the yuan field, while selects keep the field name", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  const charts = withFakeCharts(orders);
  orders.evaluate(`
    populateControls();
    dom.numericChartField.value = "金额"; dom.dateChartField.value = "下单日期"; dom.dateChartMetric.value = "金额";
    renderCharts();
    dom.v2MetricField.value = "金额"; dom.v2GroupField.value = "大区"; dom.v2AggregateMethod.value = "sum"; dom.v2DateField.value = "下单日期";
    renderV2Analysis();
  `);
  const text = (id) => orders.elements.get(id).textContent;
  const html = (id) => orders.elements.get(id).innerHTML;
  const statsCells = tableCells(orders.elements, "numericStatsTable");
  assert.deepEqual(statsCells.slice(statsCells.indexOf("金额（元）"), statsCells.indexOf("金额（元）") + 2), ["金额（元）", "29"]);
  assert.ok(statsCells.includes("数量"));
  const outlierCells = tableCells(orders.elements, "outlierTable");
  assert.deepEqual(outlierCells.slice(outlierCells.indexOf("金额（元）"), outlierCells.indexOf("金额（元）") + 2), ["金额（元）", "1"]);
  assert.equal(text("numericChartTitle"), "金额（元） 数值分布（记录频数）");
  assert.equal(text("dateChartTitle"), "金额（元） 按日期趋势（平均值）");
  assert.equal(text("v2PrimaryTableTitle"), "金额（元） 按 大区 分组（求和）");
  assert.equal(text("v2TopGroupTitle"), "金额（元） 按 大区 分组 Top 10（求和）");
  assert.equal(text("v2TrendTitle"), "金额（元） 按日期趋势（求和）");
  assert.ok(html("v2PrimaryTable").includes("<th>金额（元） 求和值</th>"));
  assert.ok(html("v2SummaryCards").includes('title="金额（元）">金额（元）</div>'));
  const axisTitles = charts.map((config) => [config.data.datasets[0].label, config.options.scales.y.title.text]);
  assert.deepEqual(axisTitles, [
    ["金额（元） 分布（记录频数）", "记录频数"],
    ["客户名称 Top 10（记录数）", "记录数"],
    ["金额（元） 平均值趋势", "平均值（元）"],
    ["金额（元） 按 大区 分组（求和）", "求和（元）"],
    ["金额（元） 按日期趋势（求和）", "求和（元）"]
  ]);
  for (const id of ["numericChartField", "dateChartMetric", "v2MetricField", "numericFilterField"]) {
    assert.ok(html(id).includes('<option value="金额">金额</option>'), id);
  }
  assert.ok(html("qualityTable").includes("<td>金额</td>"));
});

test("the quality page lists numbers read by unit and numeric-looking fields, and hides empty subsections", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  const section = (scope, id) => scope.elements.get(id).classList.contains("hidden");
  assert.equal(section(orders, "numberConversionSection"), false);
  assert.equal(orders.elements.get("numberConversionDetails").innerHTML, `<ul class="number-note-list"><li>${orders.evaluate(`escapeHtml(${JSON.stringify(ORDER_AMOUNT_CONVERSION_TEXT)})`)}</li></ul>`);
  assert.equal(section(orders, "numericLookingSection"), true);

  const scope = commitColumns(UNIT_SECTION_COLUMNS);
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(buildNumberConversionSummary(state.profiles).map(formatNumberConversion))")), [WEIGHT_CONFLICT_TEXT, SALES_CONVERSION_TEXT]);
  assert.equal(scope.elements.get("numberConversionDetails").innerHTML, `<ul class="number-note-list"><li>${WEIGHT_CONFLICT_TEXT}</li><li>${SALES_CONVERSION_TEXT}</li></ul>`);
  assert.equal(section(scope, "numericLookingSection"), false);
  assert.equal(scope.elements.get("numericLookingDetails").innerHTML, `<ul class="number-note-list"><li>${FLOOR_HINT_TEXT}</li></ul>`);

  scope.evaluate('stageFieldTypeChange("销量", "ignore"); stageFieldTypeChange("重量", "ignore"); stageFieldTypeChange("楼层", "numeric"); applyFieldConfiguration()');
  assert.equal(section(scope, "numberConversionSection"), true);
  assert.equal(section(scope, "numericLookingSection"), true);
});

test("number conversion texts order conflicting units by count, then first appearance", () => {
  const { context, evaluate } = createScriptContext();
  context.__items = [
    { field: "长度", unit: "", unitCount: 0, unitSamples: [], magnitudeCount: 0, fullwidthCount: 2, conflictUnits: [{ unit: "mm", count: 3 }, { unit: "cm", count: 2 }, { unit: "km", count: 2 }] },
    { field: "金额", unit: "元", unitCount: 5, unitSamples: ["1元", "2元", "3元"], magnitudeCount: 1, fullwidthCount: 1, conflictUnits: [] }
  ];
  assert.deepEqual(JSON.parse(evaluate("JSON.stringify(__items.map(formatNumberConversion))")), [
    "「长度」：2 个全角数字已按半角读取；同时出现 mm（3 个）、cm（2 个）和 km（2 个），没有按单位读取，这些值不能读成数字。",
    "「金额」：按\"元\"读取了 5 个值（1元、2元、3元）；按万/亿换算了 1 个值；1 个全角数字已按半角读取"
  ]);
  const scope = commitColumns({ 长度: ["3cm", "5mm", "7km", "2mm", "4cm", "8km", "6mm"] });
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(buildNumberConversionSummary(state.profiles)[0].conflictUnits)")), [
    { unit: "mm", count: 3 }, { unit: "cm", count: 2 }, { unit: "km", count: 2 }
  ]);
});

test("the quality page has static number subsections between the missing value check and the placeholders", () => {
  const html = indexHtmlSource();
  const quality = html.slice(html.indexOf('id="qualitySection"'), html.indexOf('id="filterSection"'));
  const order = ['id="qualityTable"', 'id="numberConversionSection"', 'id="numericLookingSection"', '<h3 class="subsection-title">像是空值的写法</h3>'].map((marker) => quality.indexOf(marker));
  assert.ok(order.every((index, position) => index > 0 && (position === 0 || index > order[position - 1])), JSON.stringify(order));
  assert.match(quality, /<div id="numberConversionSection" class="module-subsection hidden">\s*<h3 class="subsection-title">按单位读取的数字<\/h3>\s*<div id="numberConversionDetails" class="number-note-details"><\/div>/);
  assert.match(quality, /<div id="numericLookingSection" class="module-subsection hidden">\s*<h3 class="subsection-title">像是数字的文本列<\/h3>\s*<div id="numericLookingDetails" class="number-note-details"><\/div>/);
  const { evaluate } = createScriptContext();
  assert.equal(evaluate('["numberConversionSection", "numberConversionDetails", "numericLookingSection", "numericLookingDetails"].every((id) => Object.keys(dom).includes(id))'), true);
});

test("the quality page renders the number subsections even without numeric fields", () => {
  const scope = commitColumns({ 楼层: ["1", "2", "3", "B1", "B2", "4"], 重量: ["5kg", "500千克", "2kg", "1kg", "3kg", "4kg"] });
  assert.equal(scope.evaluate('getProfilesByType("numeric").length'), 0);
  assert.ok(scope.elements.get("numericLookingDetails").innerHTML.includes("「楼层」"));
  assert.ok(scope.elements.get("numberConversionDetails").innerHTML.includes("「重量」"));
});

test("reports carry number units, conversions and numeric-looking fields like the quality page", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  orders.evaluate('dom.dateChartMetric.value = "金额"; dom.dateChartAggregate.value = "avg"; dom.v2MetricField.value = "金额"; dom.v2GroupField.value = "大区"; dom.v2AggregateMethod.value = "sum"; dom.v2DateField.value = ""; renderV2Analysis()');
  const { report, html } = reportOf(orders);
  assert.deepEqual(report.fieldTypes.map((field) => field.numberUnit), ["", "", "", "", "", "", "元", ""]);
  assert.deepEqual(report.quality.numberConversions, [
    { field: "金额", unit: "元", unitCount: 2, unitSamples: ["2200元", "3500元"], magnitudeCount: 0, fullwidthCount: 0, conflictUnits: [] }
  ]);
  assert.deepEqual(report.quality.numericLooking, []);
  assert.ok(report.dateTrends[0].summary.startsWith("日期字段「下单日期」包含 20 个时间点；金额（元）平均值从 2024-05-06 的"), report.dateTrends[0].summary);
  assert.ok(html.includes("；像是空值的写法：<strong>4</strong> 个字段。</p>"));
  assert.doesNotMatch(html, /按单位读取：<strong>/);
  assert.doesNotMatch(html, /像是数字的文本列/);
  assert.ok(html.includes(`<h3>按单位读取的数字</h3>\n    <ul class="summary-list"><li>${orders.evaluate(`escapeHtml(${JSON.stringify(ORDER_AMOUNT_CONVERSION_TEXT)})`)}</li></ul>`));
  assert.ok(html.includes("<td>金额（元）</td><td>29</td><td>1</td>"));
  assert.ok(html.includes("<td>金额（元）</td><td>1</td><td>-4,275</td><td>9,925</td>"));
  assert.ok(html.includes("<strong>指标字段：</strong>金额（元）"));

  const scope = commitColumns(UNIT_SECTION_COLUMNS);
  const constructed = reportOf(scope);
  assert.deepEqual(constructed.report.quality.numericLooking, [{ field: "楼层", valueCount: 6, numericCount: 4, ratio: 4 / 6, samples: ["B1", "B2"] }]);
  assert.ok(constructed.html.includes("；编号重复：<strong>0</strong> 个字段；像是数字的文本列：<strong>1</strong> 个字段。</p>"));
  assert.doesNotMatch(constructed.html, /按单位读取：<strong>/);
  assert.ok(constructed.html.includes(`<h3>按单位读取的数字</h3>\n    <ul class="summary-list"><li>${WEIGHT_CONFLICT_TEXT}</li><li>${SALES_CONVERSION_TEXT}</li></ul>`));
  assert.ok(constructed.html.includes(`<h3>像是数字的文本列</h3>\n    <ul class="summary-list"><li>${FLOOR_HINT_TEXT}</li></ul>`));
});

test("legacy report snapshots without number units or hints leave both sections and counts out", () => {
  const scope = createScriptContext();
  const normalized = JSON.parse(scope.evaluate('JSON.stringify(normalizeReportSnapshot({ fieldTypes: [{ field: "金额" }] }))'));
  assert.equal(normalized.fieldTypes[0].numberUnit, "");
  assert.equal(normalized.quality.numberConversions, null);
  assert.equal(normalized.quality.numericLooking, null);
  const legacy = {
    fieldTypes: [{ field: "金额" }],
    quality: { totalMissing: 0, duplicateRows: 0, parseWarningCount: 0, outliers: [{ field: "金额", outlierCount: 1, lowerBound: "1", upperBound: "9" }] },
    numericStats: [{ field: "金额", count: 3, invalidCount: 0, mean: "2", median: "2", min: "1", max: "3", standardDeviation: "1" }]
  };
  put(scope.context, "__legacyNumbers", legacy);
  const html = scope.evaluate("buildHtmlReport(__legacyNumbers, [])");
  assert.doesNotMatch(html, /按单位读取|像是数字的文本列/);
  assert.ok(html.includes("<td>金额</td><td>3</td>"));
  // Present but empty lists: no section and no count either.
  put(scope.context, "__emptyNumbers", { ...legacy, quality: { ...legacy.quality, numberConversions: [], numericLooking: [] } });
  assert.doesNotMatch(scope.evaluate("buildHtmlReport(__emptyNumbers, [])"), /按单位读取|像是数字的文本列/);
});

test("number sections and unit labels escape field names in HTML", () => {
  const field = "<b>重\n量</b>|x";
  const scope = commitColumns({ [field]: ["5kg", "6kg", "7kg", "8kg", "9kg"], "-楼层": ["1", "2", "3", "B1", "B2"] });
  const { html } = reportOf(scope);
  assert.ok(html.includes("<td>&lt;b&gt;重\n量&lt;/b&gt;|x（kg）</td><td>5</td>"));
  assert.ok(html.includes(`<li>${scope.evaluate(`escapeHtml(${JSON.stringify(`「${field}」：按"kg"读取了 5 个值（5kg、6kg、7kg）`)})`)}</li>`));
  assert.ok(scope.elements.get("numberConversionDetails").innerHTML.includes("「&lt;b&gt;重\n量&lt;/b&gt;|x」"));
  assert.ok(scope.elements.get("numericLookingDetails").innerHTML.includes("「-楼层」"));
});

// v2.6.0: the checks behind the quality counts, shared with the problem-row lists.
const STAGE3_IMPORT_FILES = fs.readdirSync(STAGE3_FIXTURES).filter((fileName) => /\.(csv|xlsx)$/.test(fileName)).sort();
const EXAMPLE_CSV_FILES = listExampleCsvFiles();

async function importEveryQualityFile() {
  const scopes = [];
  for (const fileName of STAGE3_IMPORT_FILES) scopes.push([fileName, await importStage3Fixture(fileName)]);
  for (const fileName of EXAMPLE_CSV_FILES) {
    scopes.push([fileName, await importCsvFile(readExampleCsv(fileName), fileName)]);
  }
  return scopes;
}

// [field, type, failures counted by the shared check, profile count] per field and type.
const conversionCheckCounts = (scope) => JSON.parse(scope.evaluate(`JSON.stringify(state.profiles.flatMap((profile) => ["numeric", "date"].map((typeKey) => {
  const fails = getProfileConversionCheck(profile, typeKey);
  const count = state.rows.filter((row) => !isMissingOrConfirmed(row[profile.field], profile.confirmedKeys) && fails(row[profile.field])).length;
  return [profile.field, typeKey, count, profile.conversionFailuresByType[typeKey]];
})))`));

test("guard: the shared conversion check counts every field's failures as its profile does, for both types", async () => {
  const scopes = await importEveryQualityFile();
  assert.equal(scopes.length, 18);
  const orders = scopes.find(([fileName]) => fileName === "订单表.csv")[1];
  orders.evaluate("confirmAllCommonPlaceholders()");
  scopes.push(["订单表.csv（全部确认后）", orders]);
  for (const [fileName, scope] of scopes) {
    const counts = conversionCheckCounts(scope);
    assert.ok(counts.length >= 2, fileName);
    for (const [field, typeKey, count, expected] of counts) assert.equal(count, expected, `${fileName} ${field} ${typeKey}`);
  }
  assert.equal(orders.evaluate("getProfileConversionCheck(state.profiles[0], \"category\")"), null);
});

test("guard: profiles keep the placeholder keys they counted as missing", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  const keys = () => JSON.parse(orders.evaluate("JSON.stringify(state.profiles.map((profile) => [profile.field, profile.confirmedKeys, getConfirmedPlaceholderKeys(profile.field)]))"));
  assert.ok(keys().every(([, kept, current]) => kept.length === 0 && current.length === 0));
  orders.evaluate("confirmAllCommonPlaceholders()");
  for (const [field, kept, current] of keys()) assert.deepEqual(kept, current, field);
  assert.deepEqual(keys().find(([field]) => field === "备注")[1], ["N/A", "—"]);
  orders.evaluate('undoPlaceholderConfirmation("备注", "N/A")');
  assert.deepEqual(keys().find(([field]) => field === "备注")[1], ["—"]);
  // A copy: changing the confirmations later never changes a built profile.
  orders.evaluate('state.placeholderConfirmations["备注"].push("x")');
  assert.deepEqual(keys().find(([field]) => field === "备注")[1], ["—"]);
});

test("guard: ID grouping lists each value's rows in first-appearance order", () => {
  const { evaluate } = createScriptContext();
  const groups = JSON.parse(evaluate(`JSON.stringify((() => {
    const rows = ["B", " A", "", "B", "N/A", "A ", "C", "N/A", "B"].map((value) => ({ 编号: value }));
    const plain = groupIdValueIndexes(rows, "编号", []);
    const confirmed = groupIdValueIndexes(rows, "编号", ["N/A"]);
    return { plain: [...plain.indexesByValue], plainCount: plain.nonMissingCount, confirmed: [...confirmed.indexesByValue], sorted: sortIdDuplicates(plain.indexesByValue) };
  })())`));
  assert.deepEqual(groups.plain, [["B", [0, 3, 8]], ["A", [1, 5]], ["N/A", [4, 7]], ["C", [6]]]);
  assert.equal(groups.plainCount, 8);
  assert.deepEqual(groups.confirmed, [["B", [0, 3, 8]], ["A", [1, 5]], ["C", [6]]]);
  assert.deepEqual(groups.sorted, [["B", [0, 3, 8]], ["A", [1, 5]], ["N/A", [4, 7]]]);

  // The most frequent value is not the first one here, so the order tells the rule apart:
  // B (3 times) first, then A and C (2 times each) by first appearance.
  const ordered = JSON.parse(evaluate(`JSON.stringify((() => {
    const rows = ["A", "B", "A", "B", "C", "B", "C"].map((value) => ({ 编号: value }));
    const { indexesByValue } = groupIdValueIndexes(rows, "编号", []);
    const summary = buildIdDuplicateSummary(rows, ["编号"], [{ field: "编号", typeKey: "id" }], () => null);
    return { sorted: sortIdDuplicates(indexesByValue).map(([value]) => value), summary: summary[0].values.map((item) => item.value), mostlyRepeated: summary[0].mostlyRepeated };
  })())`));
  assert.deepEqual(ordered.sorted, ["B", "A", "C"]);
  assert.equal(ordered.mostlyRepeated, false);
  assert.deepEqual(ordered.summary, ["B", "A", "C"]);
});

// v2.6.0: the rows behind each quality count (listProblemCategories / listProblemRows).
// Checks every category against the figure the quality page shows, and the order:
// file order, or by group for the two duplicate kinds.
const problemRowCheck = (scope) => JSON.parse(scope.evaluate(`JSON.stringify((() => {
  const mismatches = [];
  const names = [];
  const ascending = (list) => list.every((value, index) => index === 0 || value > list[index - 1]);
  listProblemCategories().forEach((category) => {
    const name = [category.kind, category.field, category.key].filter((part) => part !== undefined).join(" / ");
    const result = listProblemRows(category);
    if (!result) {
      mismatches.push(name + ": no list");
      return;
    }
    names.push(name + " = " + result.items.length);
    const items = result.items;
    let expected;
    if (category.kind === "missing") expected = getFieldProfile(category.field).missingCount;
    if (category.kind === "conversion") expected = getFieldProfile(category.field).conversionFailureCount;
    if (category.kind === "outlier") expected = getQualityNumericStats().find((stat) => stat.field === category.field).outlierCount;
    if (category.kind === "placeholder") expected = state.placeholders.find((entry) => entry.field === category.field && entry.key === category.key).count;
    if (category.kind === "parseWarnings") expected = state.parseWarnings.length;
    if (category.kind === "suspectedTotal") expected = state.suspectedTotals.length;
    if (category.kind === "formulaError") expected = state.formulaErrors.find((item) => item.field === category.field).count;
    if (category.kind === "dateStyle") expected = state.dateStyles.find((item) => item.field === category.field).otherCount;
    if (category.kind === "lookAlike") expected = state.lookAlikes.find((item) => item.field === category.field).rowCount;
    if (category.kind === "duplicateRows") {
      // The quality count leaves out the first row of each group, which is kept.
      expected = state.duplicateRows + result.groupCount;
      if (items.filter((item) => item.kept).length !== result.groupCount) mismatches.push(name + ": kept rows");
    }
    if (category.kind === "idDuplicate") {
      const summary = state.idDuplicates.find((item) => item.field === category.field);
      const groups = new Map();
      items.forEach((item) => groups.set(item.value, [...(groups.get(item.value) || []), item]));
      if (groups.size !== summary.duplicateValueCount) mismatches.push(name + ": " + groups.size + " values, summary " + summary.duplicateValueCount);
      summary.values.forEach((value, position) => {
        const rows = groups.get(value.value) || [];
        if ([...groups.keys()][position] !== value.value) mismatches.push(name + ": value order at " + position);
        if (rows.length !== value.count) mismatches.push(name + ": " + value.value + " has " + rows.length + " rows, summary " + value.count);
        if (JSON.stringify(rows.slice(0, 10).map((item) => item.line)) !== JSON.stringify(value.rowNumbers)) mismatches.push(name + ": " + value.value + " row numbers");
      });
      expected = [...groups.values()].reduce((sum, rows) => sum + rows.length, 0);
    }
    if (items.length !== expected) mismatches.push(name + ": listed " + items.length + ", quality page " + expected);
    if (category.kind === "duplicateRows" || category.kind === "idDuplicate") {
      const firsts = [];
      items.forEach((item, index) => {
        if (index === 0 || item.group !== items[index - 1].group) firsts.push(item.index);
        else if (item.index <= items[index - 1].index) mismatches.push(name + ": order inside group " + item.group);
      });
      if (category.kind === "duplicateRows" && !ascending(firsts)) mismatches.push(name + ": groups not by first appearance");
    } else if (category.kind === "lookAlike") {
      items.forEach((item, index) => {
        if (index > 0 && item.group === items[index - 1].group && item.index <= items[index - 1].index) mismatches.push(name + ": order inside group " + item.group);
      });
    } else if (category.kind !== "parseWarnings" && !ascending(items.map((item) => item.index))) {
      mismatches.push(name + ": not in file order");
    }
    if (items.some((item) => item.index !== null && item.line !== getSourceRowNumber(state.rows[item.index]))) mismatches.push(name + ": line numbers");
  });
  return { names, mismatches };
})())`));

test("every problem-row list matches its quality count on the stage3 fixtures and examples", async () => {
  const scopes = await importEveryQualityFile();
  assert.equal(scopes.length, 18);
  for (const [fileName, scope] of scopes) {
    const { names, mismatches } = problemRowCheck(scope);
    assert.deepEqual(mismatches, [], fileName);
    assert.ok(names.length > 0, fileName);
  }
  const names = (fileName) => problemRowCheck(scopes.find(([name]) => name === fileName)[1]).names;
  assert.deepEqual(names("订单表.csv"), [
    "missing / 备注 = 26",
    "conversion / 下单日期 = 1", "conversion / 数量 = 1", "conversion / 金额 = 1",
    "outlier / 数量 = 1", "outlier / 金额 = 1",
    "idDuplicate / 订单号 = 2",
    "placeholder / 下单日期 / — = 1", "placeholder / 数量 / — = 1", "placeholder / 金额 / N/A = 1",
    "placeholder / 备注 / N/A = 1", "placeholder / 备注 / — = 1"
  ]);
  assert.deepEqual(names("花名册.xlsx"), [
    "missing / 身份证号 = 2", "conversion / 月薪 = 2", "dateStyle / 入职日期 = 6", "outlier / 月薪 = 3", "duplicateRows = 2",
    "idDuplicate / 工号 = 3", "idDuplicate / 身份证号 = 3", "idDuplicate / 手机号 = 3"
  ]);
  // 用户ID and 会话ID are mostly repeated: no entry for them.
  assert.deepEqual(names("sample-user-behavior.csv").filter((name) => name.startsWith("idDuplicate")), ["idDuplicate / 事件ID = 4"]);
});

test("problem-row lists follow confirmations, manual types, ignored fields and undo", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  const check = () => {
    const result = problemRowCheck(orders);
    assert.deepEqual(result.mismatches, []);
    return result.names;
  };
  orders.evaluate("confirmAllCommonPlaceholders()");
  assert.deepEqual(check().filter((name) => !name.startsWith("placeholder")), [
    "missing / 下单日期 = 1", "missing / 数量 = 1", "missing / 金额 = 1", "missing / 备注 = 28",
    "outlier / 数量 = 1", "outlier / 金额 = 1", "idDuplicate / 订单号 = 2"
  ]);
  const confirmedNotes = JSON.parse(orders.evaluate('JSON.stringify(listProblemRows({ kind: "placeholder", field: "金额", key: "N/A" }).items.map((item) => [item.line, item.note]))'));
  assert.deepEqual(confirmedNotes, [[15, "已按空值处理"]]);

  // 客户名称 becomes a mostly repeated ID field; 备注 is ignored but keeps its missing values.
  orders.evaluate('stageFieldTypeChange("金额", "category"); stageFieldTypeChange("客户名称", "id"); stageFieldTypeChange("备注", "ignore"); applyFieldConfiguration()');
  const ignored = check();
  assert.ok(ignored.includes("missing / 备注 = 28"));
  assert.ok(!ignored.some((name) => name.startsWith("outlier / 金额") || name.startsWith("idDuplicate / 客户名称")));
  assert.ok(ignored.includes("placeholder / 备注 / N/A = 1"));

  orders.evaluate('stageFieldTypeChange("金额", "numeric"); stageFieldTypeChange("数量", "date"); applyFieldConfiguration()');
  assert.ok(check().includes("conversion / 数量 = 29"));
  orders.evaluate('undoPlaceholderConfirmation("金额", "N/A")');
  const undone = check();
  assert.ok(undone.includes("conversion / 金额 = 1"));
  assert.ok(!undone.includes("missing / 金额 = 1"));
  assert.equal(orders.evaluate('state.manualFieldTypes["金额"]'), undefined);
  assert.equal(orders.evaluate('getFieldProfile("金额").typeKey'), "numeric");
});

test("ID duplicate rows follow the quality table's order, with every value and row", async () => {
  const ordered = await importCsvFile(`编号,值\r\n${["A", "B", "A", "B", "C", "B", "C"].map((id, index) => `${id},${index}`).join("\r\n")}\r\n`);
  const items = JSON.parse(ordered.evaluate('JSON.stringify(listProblemRows({ kind: "idDuplicate", field: "编号" }).items.map((item) => [item.value, item.line, item.note]))'));
  assert.deepEqual(items, [
    ["B", 3, "第 1 组：「B」出现 3 次"], ["B", 5, "第 1 组：「B」出现 3 次"], ["B", 7, "第 1 组：「B」出现 3 次"],
    ["A", 2, "第 2 组：「A」出现 2 次"], ["A", 4, "第 2 组：「A」出现 2 次"],
    ["C", 6, "第 3 组：「C」出现 2 次"], ["C", 8, "第 3 组：「C」出现 2 次"]
  ]);

  // More than 10 duplicate values: the summary lists 10, the problem rows every one.
  const ids = [
    ...Array.from({ length: 15 }, () => "S000"),
    ...Array.from({ length: 11 }, (_, index) => [`S${index + 1}`, `S${index + 1}`]).flat(),
    ...Array.from({ length: 40 }, (_, index) => `S${index + 100}`)
  ];
  const long = await importCsvFile(`学号\r\n${ids.join("\r\n")}\r\n`);
  assert.deepEqual(problemRowCheck(long).mismatches, []);
  const result = JSON.parse(long.evaluate('JSON.stringify(listProblemRows({ kind: "idDuplicate", field: "学号" }))'));
  assert.equal(result.items.length, 37);
  assert.equal(new Set(result.items.map((item) => item.value)).size, 12);
  assert.equal(result.items.filter((item) => item.value === "S000").length, 15);
  assert.equal(long.evaluate('state.idDuplicates[0].values.length'), 10);
});

test("duplicate-row lists mark the kept first row of each group", async () => {
  const scope = await importCsvFile("城市,值\r\n上海,1\r\n北京,2\r\n上海,1\r\n北京,2\r\n上海,1\r\n广州,3\r\n");
  const result = JSON.parse(scope.evaluate('JSON.stringify(listProblemRows({ kind: "duplicateRows" }))'));
  assert.equal(result.groupCount, 2);
  assert.equal(scope.evaluate("state.duplicateRows"), 3);
  assert.deepEqual(result.items.map((item) => [item.line, item.group, item.kept, item.note]), [
    [2, 1, true, "第 1 组（保留）"], [4, 1, false, "第 1 组"], [6, 1, false, "第 1 组"],
    [3, 2, true, "第 2 组（保留）"], [5, 2, false, "第 2 组"]
  ]);
});

test("missing, conversion and outlier lists explain each row", async () => {
  const scope = await importCsvFile("金额,备注\r\n10,a\r\n12,N/A\r\n,b\r\n11,c\r\n待定,d\r\n13,e\r\n9,f\r\n10,g\r\n11,h\r\n12,i\r\n500,j\r\n");
  scope.evaluate('confirmPlaceholder("备注", "N/A")');
  const notes = (category) => JSON.parse(scope.evaluate(`JSON.stringify(listProblemRows(${JSON.stringify(category)}).items.map((item) => [item.line, item.note]))`));
  assert.deepEqual(notes({ kind: "missing", field: "金额" }), [[4, "空值"]]);
  assert.deepEqual(notes({ kind: "missing", field: "备注" }), [[3, "已按空值处理"]]);
  assert.deepEqual(notes({ kind: "conversion", field: "金额" }), [[6, "认不出是数字"]]);
  const bounds = JSON.parse(scope.evaluate('JSON.stringify(getQualityNumericStats().find((stat) => stat.field === "金额"))'));
  assert.deepEqual(notes({ kind: "outlier", field: "金额" }), [[12, `高于上界 ${scope.evaluate(`formatFieldNumber("金额", ${bounds.upperBound})`)}`]]);
  assert.equal(scope.evaluate('listProblemRows({ kind: "conversion", field: "备注" })'), null);
  assert.equal(scope.evaluate('listProblemRows({ kind: "missing", field: "不存在" })'), null);
});

test("parse-warning lists show rows that were imported and say when there is none", async () => {
  const scope = await importCsvFile("a,b\n1,2\n\n \n3\n4,5\n");
  const items = JSON.parse(scope.evaluate('JSON.stringify(listProblemRows({ kind: "parseWarnings" }).items)'));
  assert.deepEqual(items.map(({ index, line }) => [index, line]), [[null, 4], [1, 5]]);
  assert.equal(items[0].note, "第 4 行只有 1 列（表头有 2 列），缺少的单元格按空值处理。（这一行没有可分析的数据，未导入）");
  assert.equal(items[1].note, "第 5 行只有 1 列（表头有 2 列），缺少的单元格按空值处理。");
  assert.deepEqual(problemRowCheck(scope).mismatches, []);

  const oneColumn = await importCsvFile("a;b\n1;2\n3\n5;6\n");
  const [warning] = JSON.parse(oneColumn.evaluate('JSON.stringify(listProblemRows({ kind: "parseWarnings" }).items)'));
  assert.equal(warning.index, null);
  assert.match(warning.note, /（没有行号）$/);

  const clean = await importStage3Fixture("订单表.csv");
  assert.equal(clean.evaluate('listProblemRows({ kind: "parseWarnings" })'), null);
  assert.equal(clean.evaluate('listProblemCategories().some((category) => category.kind === "parseWarnings")'), false);
});

// v2.6.0: the data preview pages through the analysis rows, 50 at a time, with the
// file row of each row in front.
const PREVIEW_CITIES = ["上海", "北京", "广州"];
const previewCsv = () => `编号,城市\r\n${Array.from({ length: 120 }, (_, index) => `R${index + 1},${PREVIEW_CITIES[index % 3]}`).join("\r\n")}\r\n`;
// [file row, 编号] of each preview row, and the pager.
const previewState = (scope) => ({
  rows: [...scope.elements.get("previewTable").innerHTML.matchAll(/<tr><td>([^<]*)<\/td><td>([^<]*)<\/td>/g)].map((match) => [match[1], match[2]]),
  info: scope.elements.get("previewPageInfo").textContent,
  disabled: ["previewFirstPage", "previewPrevPage", "previewNextPage", "previewLastPage"].map((id) => scope.elements.get(id).disabled)
});

test("the data preview shows 50 rows a page with file rows and moves between pages", async () => {
  const scope = await importCsvFile(previewCsv(), "preview.csv");
  let shown = previewState(scope);
  assert.equal(shown.rows.length, 50);
  assert.deepEqual(shown.rows[0], ["2", "R1"]);
  assert.deepEqual(shown.rows[49], ["51", "R50"]);
  assert.equal(shown.info, "第 1 / 3 页，共 120 行");
  assert.deepEqual(shown.disabled, [true, true, false, false]);
  assert.match(scope.elements.get("previewTable").innerHTML, /<th>文件行号<\/th><th>编号<\/th><th>城市<\/th>/);

  scope.evaluate('goToPreviewPage("next")');
  shown = previewState(scope);
  assert.deepEqual([shown.rows[0], shown.rows[49]], [["52", "R51"], ["101", "R100"]]);
  assert.equal(shown.info, "第 2 / 3 页，共 120 行");
  assert.deepEqual(shown.disabled, [false, false, false, false]);

  scope.evaluate('goToPreviewPage("last")');
  shown = previewState(scope);
  assert.equal(shown.rows.length, 20);
  assert.deepEqual(shown.rows[19], ["121", "R120"]);
  assert.equal(shown.info, "第 3 / 3 页，共 120 行");
  assert.deepEqual(shown.disabled, [false, false, true, true]);
  scope.evaluate('goToPreviewPage("next")');
  assert.equal(previewState(scope).info, "第 3 / 3 页，共 120 行");

  scope.evaluate('goToPreviewPage("prev")');
  assert.equal(previewState(scope).info, "第 2 / 3 页，共 120 行");
  scope.evaluate('goToPreviewPage("first")');
  assert.equal(previewState(scope).info, "第 1 / 3 页，共 120 行");
});

test("the data preview follows filters from page 1 and keeps the original file rows", async () => {
  const scope = await importCsvFile(previewCsv(), "preview.csv");
  // From the last page to a filter that still has two pages: page 1, not a clamped page 2.
  scope.evaluate('goToPreviewPage("last")');
  assert.equal(previewState(scope).info, "第 3 / 3 页，共 120 行");
  applyFilterState(scope, { category: { field: "城市", values: ["北京", "上海"] } });
  const twoPages = previewState(scope);
  assert.equal(twoPages.info, "第 1 / 2 页，共 80 行");
  assert.deepEqual(twoPages.rows[0], ["2", "R1"]);
  assert.deepEqual(twoPages.disabled, [true, true, false, false]);

  scope.evaluate('goToPreviewPage("last")');
  applyFilterState(scope, { category: { field: "城市", values: ["北京"] } });
  const shown = previewState(scope);
  assert.equal(shown.info, "第 1 / 1 页，共 40 行");
  assert.deepEqual(shown.rows.slice(0, 2), [["3", "R2"], ["6", "R5"]]);
  assert.deepEqual(shown.rows[39], ["120", "R119"]);
  assert.deepEqual(shown.disabled, [true, true, true, true]);

  applyFilterState(scope, { category: { field: "城市", values: [] } });
  assert.match(scope.elements.get("previewTable").innerHTML, /当前筛选条件下没有可预览的数据。/);
  assert.equal(previewState(scope).info, "第 1 / 1 页，共 0 行");

  // A field configuration rebuild clears the filters and goes back to page 1 too.
  applyFilterState(scope, { category: null });
  scope.evaluate('goToPreviewPage("next")');
  scope.evaluate('stageFieldTypeChange("城市", "ignore"); applyFieldConfiguration()');
  assert.equal(previewState(scope).info, "第 1 / 3 页，共 120 行");
});

test("renderPreview() clamps the page and works without a page in the state", () => {
  const { context, elements, evaluate } = createScriptContext();
  put(context, "__rows", Array.from({ length: 60 }, (_, index) => ({ 值: String(index) })));
  evaluate('commitTabularData(["值"], __rows, "plain.csv", beginImport())');
  for (const [page, info] of [[99, "第 2 / 2 页，共 60 行"], [0, "第 1 / 2 页，共 60 行"], [-3, "第 1 / 2 页，共 60 行"], [undefined, "第 1 / 2 页，共 60 行"], ["x", "第 1 / 2 页，共 60 行"]]) {
    put(context, "__page", page);
    evaluate("state.previewPage = __page; renderPreview()");
    assert.equal(elements.get("previewPageInfo").textContent, info, String(page));
  }
  // Rows that did not come through an import have no file row: "—".
  assert.match(elements.get("previewTable").innerHTML, /<tr><td>—<\/td><td>0<\/td><\/tr>/);
  evaluate("beginImport()");
  assert.equal(evaluate("state.previewPage"), 1);
});

test("the preview heading, pager and its buttons are static in index.html and bound in script.js", () => {
  const html = indexHtmlSource();
  const overview = html.slice(html.indexOf('id="overviewSection"'), html.indexOf('id="schemaSection"'));
  assert.match(overview, /<h3 class="subsection-title">数据预览<\/h3>/);
  assert.match(overview, /<table id="previewTable" aria-label="数据预览"><\/table>/);
  assert.doesNotMatch(overview, /前 10 行/);
  const ids = ["previewPageInfo", "previewFirstPage", "previewPrevPage", "previewNextPage", "previewLastPage"];
  ids.forEach((id) => assert.match(overview, new RegExp(`id="${id}"`), id));
  const { evaluate } = createScriptContext();
  assert.equal(evaluate(`${JSON.stringify(ids)}.every((id) => Object.keys(dom).includes(id))`), true);
  const script = fs.readFileSync(path.join(__dirname, "..", "script.js"), "utf8");
  for (const [id, action] of [["previewFirstPage", "first"], ["previewPrevPage", "prev"], ["previewNextPage", "next"], ["previewLastPage", "last"]]) {
    assert.ok(script.includes(`dom.${id}.addEventListener("click", () => goToPreviewPage("${action}"));`), id);
  }
});

// v2.6.0: the problem-row viewer on the quality page.
const PROBLEM_ENTRY_CONTAINERS = ["qualitySummary", "parseWarningSummary", "qualityTable", "outlierTable", "idDuplicateDetails", "placeholderDetails", "dateStyleDetails", "lookAlikeDetails"];
// Every "查看" button: its container, data-problem-index, aria-label and other data attributes.
const problemButtons = (scope) => PROBLEM_ENTRY_CONTAINERS.flatMap((id) => [...scope.elements.get(id).innerHTML.matchAll(/<button\b[^>]*\bdata-problem-index="(\d+)"[^>]*>/g)]
  .map((match) => ({
    container: id,
    index: Number(match[1]),
    label: (match[0].match(/aria-label="([^"]*)"/) || [])[1],
    dataAttributes: [...match[0].matchAll(/\b(data-[a-z-]+)=/g)].map((attribute) => attribute[1])
  })));
const clickProblemButton = (scope, index) => {
  put(scope.context, "__problemButton", index === null ? null : { dataset: { problemIndex: String(index) } });
  scope.evaluate("handleProblemViewClick({ target: { closest: () => __problemButton } })");
};
const openCategory = (scope) => JSON.parse(scope.evaluate("JSON.stringify(state.problemView && state.problemView.category)"));
const categoryAt = (scope, index) => JSON.parse(scope.evaluate(`JSON.stringify(state.problemCategories[${index}])`));
const viewerState = (scope) => ({
  hidden: scope.elements.get("problemRowsSection").classList.contains("hidden"),
  summary: scope.elements.get("problemRowsSummary").textContent,
  message: scope.elements.get("problemRowsMessage").textContent,
  disabled: ["problemRowsFirstPage", "problemRowsPrevPage", "problemRowsNextPage", "problemRowsLastPage"].map((id) => scope.elements.get(id).disabled),
  lines: [...scope.elements.get("problemRowsTable").innerHTML.matchAll(/<tr><td>([^<]*)<\/td><td>/g)].map((match) => match[1])
});
const sameCategory = (category) => ({ kind: category.kind, field: category.field, key: category.key });

test("every 查看 button opens its own category and carries only its index", async () => {
  const scopes = [await importStage3Fixture("订单表.csv"), await importStage3Fixture("花名册.xlsx"), await importCsvFile("a,b\n1,2\n3\n4,5\n", "warnings.csv")];
  for (const scope of scopes) {
    const buttons = problemButtons(scope);
    const categoryCount = scope.evaluate("state.problemCategories.length");
    assert.ok(categoryCount > 0);
    // Each category has a button, and every button names what it shows.
    assert.deepEqual([...new Set(buttons.map((button) => button.index))].sort((a, b) => a - b), Array.from({ length: categoryCount }, (_, index) => index));
    for (const button of buttons) {
      assert.deepEqual(button.dataAttributes, ["data-problem-index"], button.label);
      assert.match(button.label, /^查看/);
      clickProblemButton(scope, button.index);
      assert.deepEqual(sameCategory(openCategory(scope)), sameCategory(categoryAt(scope, button.index)), button.label);
      assert.equal(scope.elements.get("problemRowsCategory").value, String(button.index));
      assert.equal(viewerState(scope).hidden, false);
    }
  }
  const [orders, roster, warnings] = scopes;
  const label = (scope, container) => problemButtons(scope).filter((button) => button.container === container).map((button) => button.label);
  assert.deepEqual(label(orders, "qualityTable").slice(0, 2), ["查看「下单日期」认不出的值", "查看「数量」认不出的值"]);
  assert.ok(label(orders, "qualityTable").includes("查看「备注」的空值行"));
  assert.deepEqual(label(orders, "outlierTable"), ["查看「数量」偏大或偏小的值", "查看「金额」偏大或偏小的值"]);
  assert.deepEqual(label(orders, "idDuplicateDetails"), ["查看「订单号」的重复行"]);
  assert.ok(label(orders, "placeholderDetails").includes("查看「金额」中「N/A」所在的行"));
  assert.deepEqual(label(roster, "qualitySummary"), ["查看重复行"]);
  assert.deepEqual(label(warnings, "parseWarningSummary"), ["查看读取时有疑问的行"]);
  assert.equal(warnings.elements.get("parseWarningSummary").classList.contains("hidden"), false);
  assert.match(warnings.elements.get("parseWarningSummary").innerHTML, /读取时有疑问：1 处，涉及 1 行/);
  assert.equal(orders.elements.get("parseWarningSummary").classList.contains("hidden"), true);

  // A click outside a button, or on a stale index, changes nothing.
  clickProblemButton(orders, null);
  clickProblemButton(orders, 99);
  assert.equal(viewerState(orders).hidden, false);
});

test("the 查看 columns keep a cell element in every row and leave the count cells as they were", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  for (const id of ["qualityTable", "outlierTable"]) {
    const html = orders.elements.get(id).innerHTML;
    assert.match(html, /<th>查看<\/th>/, id);
    const rows = [...html.matchAll(/<tr>([\s\S]*?)<\/tr>/g)].map((match) => match[1]).filter((row) => row.includes("<td"));
    for (const row of rows) {
      const lastCell = [...row.matchAll(/<td[^>]*>([\s\S]*?)<\/td>/g)].pop()[1];
      assert.match(lastCell, /^<(button|span)\b/, `${id}: ${row.slice(0, 60)}`);
    }
  }
  assert.deepEqual(qualityRow(orders.elements, "金额").slice(-1)[0].text, "1（N/A）");
});

test("the viewer lists rows with their file rows, notes and highlighted cells, 50 a page", async () => {
  const csv = `编号,备注\r\n${Array.from({ length: 120 }, (_, index) => `R${index + 1},${index % 2 ? "" : `说明${index}`}`).join("\r\n")}\r\n`;
  const scope = await importCsvFile(csv, "notes.csv");
  scope.evaluate('openProblemView({ kind: "missing", field: "备注" })');
  let view = viewerState(scope);
  assert.equal(view.summary, "共 60 行，第 1 / 2 页");
  assert.deepEqual(view.disabled, [true, true, false, false]);
  assert.equal(view.lines.length, 50);
  assert.deepEqual(view.lines.slice(0, 2), ["3", "5"]);
  const html = scope.elements.get("problemRowsTable").innerHTML;
  assert.match(html, /<th>文件行号<\/th><th>说明<\/th><th>编号<\/th><th>备注<\/th>/);
  assert.match(html, /<tr><td>3<\/td><td>空值<\/td><td>R2<\/td><td class="cell-problem"><\/td><\/tr>/);

  scope.evaluate('goToProblemRowsPage("next")');
  view = viewerState(scope);
  assert.equal(view.summary, "共 60 行，第 2 / 2 页");
  assert.deepEqual(view.disabled, [false, false, true, true]);
  assert.equal(view.lines.length, 10);
  assert.equal(view.lines[9], "121");
  scope.evaluate('goToProblemRowsPage("first")');
  assert.equal(viewerState(scope).summary, "共 60 行，第 1 / 2 页");

  // Filters do not touch the viewer: it stays on the original rows and its page.
  scope.evaluate('goToProblemRowsPage("last")');
  applyFilterState(scope, { category: { field: "编号", values: ["R1"] } });
  assert.equal(viewerState(scope).summary, "共 60 行，第 2 / 2 页");
  assert.equal(openCategory(scope).kind, "missing");
});

test("viewer summaries name the unit of each kind of count", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  orders.evaluate('openProblemView({ kind: "missing", field: "备注" })');
  assert.equal(viewerState(orders).summary, "共 26 行，第 1 / 1 页");
  orders.evaluate('openProblemView({ kind: "conversion", field: "金额" })');
  assert.equal(viewerState(orders).summary, "共 1 个值，第 1 / 1 页");
  const roster = await importStage3Fixture("花名册.xlsx");
  roster.evaluate('openProblemView({ kind: "duplicateRows" })');
  assert.equal(viewerState(roster).summary, "重复 1 行（1 组，共 2 行，每组第一行保留），第 1 / 1 页");
  assert.doesNotMatch(roster.elements.get("problemRowsTable").innerHTML, /cell-problem/);
  const warnings = await importCsvFile("a,b\n1,2\n\n \n3\n4,5\n", "warnings.csv");
  warnings.evaluate('openProblemView({ kind: "parseWarnings" })');
  assert.equal(viewerState(warnings).summary, "共 2 处，第 1 / 1 页");
  assert.deepEqual(viewerState(warnings).lines, ["4", "5"]);
  // The dropped whitespace row has no cells; the imported one shows its values.
  assert.match(warnings.elements.get("problemRowsTable").innerHTML, /<tr><td>4<\/td><td>[^<]*（这一行没有可分析的数据，未导入）<\/td><td><\/td><td><\/td><\/tr>/);
  assert.match(warnings.elements.get("problemRowsTable").innerHTML, /<tr><td>5<\/td><td>[^<]*<\/td><td>3<\/td><td><\/td><\/tr>/);
});

test("the category list switches categories and labels each count", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  orders.evaluate('openProblemView({ kind: "missing", field: "备注" })');
  const options = [...orders.elements.get("problemRowsCategory").innerHTML.matchAll(/<option value="(\d+)"( selected)?>([^<]*)<\/option>/g)].map((match) => [match[1], Boolean(match[2]), match[3]]);
  assert.deepEqual(options.slice(0, 7), [
    ["0", true, "空值 · 备注（26 行）"],
    ["1", false, "认不出的数字或日期 · 下单日期（1 个值）"],
    ["2", false, "认不出的数字或日期 · 数量（1 个值）"],
    ["3", false, "认不出的数字或日期 · 金额（1 个值）"],
    ["4", false, "偏大或偏小的值 · 数量（1 行）"],
    ["5", false, "偏大或偏小的值 · 金额（1 行）"],
    ["6", false, "编号重复 · 订单号（1 个重复值）"]
  ]);
  assert.equal(options[9][2], "像是空值的写法 · 金额 · 「N/A」（1 行）");
  orders.evaluate('selectProblemCategory("6")');
  assert.deepEqual(sameCategory(openCategory(orders)), { kind: "idDuplicate", field: "订单号", key: undefined });
  assert.equal(viewerState(orders).summary, "共 2 行，第 1 / 1 页");
  const roster = await importStage3Fixture("花名册.xlsx");
  roster.evaluate('openProblemView({ kind: "duplicateRows" })');
  assert.match(roster.elements.get("problemRowsCategory").innerHTML, />整行重复（1 行）</);
  const warnings = await importCsvFile("a,b\n1,2\n3\n4,5\n", "warnings.csv");
  warnings.evaluate('openProblemView({ kind: "parseWarnings" })');
  assert.match(warnings.elements.get("problemRowsCategory").innerHTML, />读取时有疑问（1 处）</);
});

test("opening the viewer moves focus to its title, and closing hides it", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  let focused = 0;
  orders.elements.get("problemRowsTitle").focus = () => { focused += 1; };
  clickProblemButton(orders, 0);
  assert.equal(focused, 1);
  orders.evaluate("closeProblemView()");
  assert.equal(viewerState(orders).hidden, true);
  assert.equal(orders.evaluate("state.problemView"), null);
  assert.equal(orders.elements.get("problemRowsTable").innerHTML, "");
  assert.equal(orders.elements.get("problemRowsCategory").innerHTML, "");
});

// v3.0.0: 查看 scrolls the viewer to the top of the screen (the title alone only scrolled
// into view, landing at the bottom); switching category, paging and re-listing never scroll.
test("v3.0.0: 查看 scrolls the problem rows section to the top; category, paging and re-listing do not", async () => {
  const csv = `金额,备注\r\n${Array.from({ length: 120 }, (_, index) => `${index},${index % 2 ? "" : index % 4 ? "N/A" : `说明${index}`}`).join("\r\n")}\r\n`;
  const scope = await importCsvFile(csv, "scroll.csv");
  const scrolls = [];
  const focuses = [];
  scope.elements.get("problemRowsSection").scrollIntoView = (...args) => { scrolls.push(args); };
  scope.elements.get("problemRowsTitle").focus = (...args) => { focuses.push(args); };
  const missingIndex = scope.evaluate('state.problemCategories.findIndex((item) => item.kind === "missing" && item.field === "备注")');
  const placeholderIndex = scope.evaluate('state.problemCategories.findIndex((item) => item.kind === "placeholder")');
  assert.ok(missingIndex >= 0 && placeholderIndex >= 0);

  // A button whose category is gone opens nothing and scrolls nothing.
  clickProblemButton(scope, 999);
  assert.deepEqual(scrolls, []);
  assert.deepEqual(focuses, []);

  clickProblemButton(scope, missingIndex);
  assert.equal(viewerState(scope).hidden, false);
  assert.deepEqual(JSON.parse(JSON.stringify(scrolls)), [[{ block: "start" }]]);
  assert.deepEqual(JSON.parse(JSON.stringify(focuses)), [[{ preventScroll: true }]]);

  scope.evaluate(`selectProblemCategory("${placeholderIndex}")`);
  assert.equal(openCategory(scope).kind, "placeholder");
  scope.evaluate(`selectProblemCategory("${missingIndex}")`);
  scope.evaluate('goToProblemRowsPage("next")');
  assert.equal(viewerState(scope).summary, "共 60 行，第 2 / 2 页");
  scope.evaluate('goToProblemRowsPage("first")');
  scope.evaluate("refreshProblemView()");
  scope.evaluate('confirmPlaceholder("备注", "N/A")');
  assert.equal(viewerState(scope).summary, "共 90 行，第 1 / 2 页");
  assert.equal(scrolls.length, 1);
  assert.equal(focuses.length, 1);
});

test("the viewer finds its category again after field configuration, confirmation and undo", async () => {
  const csv = `金额,备注\r\n${Array.from({ length: 120 }, (_, index) => `${index},${index % 2 ? "" : index % 4 ? "N/A" : `说明${index}`}`).join("\r\n")}\r\n`;
  const scope = await importCsvFile(csv, "refind.csv");
  scope.evaluate('openProblemView({ kind: "missing", field: "备注" })');
  scope.evaluate('goToProblemRowsPage("next")');
  assert.equal(viewerState(scope).summary, "共 60 行，第 2 / 2 页");

  // Confirming N/A keeps the category with more rows, back on page 1.
  scope.evaluate('confirmPlaceholder("备注", "N/A")');
  assert.equal(viewerState(scope).summary, "共 90 行，第 1 / 2 页");
  assert.equal(viewerState(scope).hidden, false);
  scope.evaluate('goToProblemRowsPage("next")');
  scope.evaluate('undoPlaceholderConfirmation("备注", "N/A")');
  assert.equal(viewerState(scope).summary, "共 60 行，第 1 / 2 页");

  // Ignoring the field keeps its missing rows; a field configuration rebuild goes back to page 1.
  scope.evaluate('goToProblemRowsPage("next")');
  scope.evaluate('stageFieldTypeChange("备注", "ignore")');
  assert.equal(scope.elements.get("problemRowsDraftNote").classList.contains("hidden"), false);
  scope.evaluate("applyFieldConfiguration()");
  assert.equal(viewerState(scope).summary, "共 60 行，第 1 / 2 页");
  assert.equal(scope.elements.get("problemRowsDraftNote").classList.contains("hidden"), true);

  // A category that is gone closes the viewer with a note.
  const orders = await importStage3Fixture("订单表.csv");
  orders.evaluate('openProblemView({ kind: "conversion", field: "金额" })');
  orders.evaluate('confirmPlaceholder("金额", "N/A")');
  assert.equal(viewerState(orders).hidden, true);
  assert.equal(viewerState(orders).message, "这一类已经没有问题行。");
  orders.evaluate('openProblemView({ kind: "outlier", field: "金额" })');
  assert.equal(viewerState(orders).message, "");
  orders.evaluate('stageFieldTypeChange("金额", "category"); applyFieldConfiguration()');
  assert.equal(viewerState(orders).hidden, true);
  assert.equal(viewerState(orders).message, "这一类已经没有问题行。");
});

test("a new import closes the viewer and clears its state", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  orders.evaluate('openProblemView({ kind: "missing", field: "备注" })');
  put(orders.context, "__cityRows", [{ 城市: "上海" }, { 城市: "" }]);
  orders.evaluate('commitTabularData(["城市"], __cityRows, "cities.csv", beginImport())');
  assert.equal(orders.evaluate("state.problemView"), null);
  assert.equal(viewerState(orders).hidden, true);
  assert.equal(orders.elements.get("problemRowsTable").innerHTML, "");
  orders.evaluate("beginImport()");
  assert.equal(orders.evaluate("state.problemCategories.length"), 0);
});

test("the viewer escapes field names, values and labels", async () => {
  const field = '<b>备&注"</b>';
  const scope = await importCsvFile(`编号,"${field.replace(/"/g, '""')}",其他\r\nR1,<i>x</i>,a\r\nR2,,<i>y</i>&'q'\r\nR3,'&',c\r\n`, "escape.csv");
  scope.evaluate(`openProblemView({ kind: "missing", field: ${JSON.stringify(field)} })`);
  const html = ["problemRowsTable", "problemRowsCategory", "qualityTable"].map((id) => scope.elements.get(id).innerHTML).join("\n");
  assert.doesNotMatch(html, /<b>|<i>/);
  assert.ok(html.includes("&lt;b&gt;备&amp;注&quot;&lt;/b&gt;"));
  assert.ok(scope.elements.get("problemRowsTable").innerHTML.includes("<td>&lt;i&gt;y&lt;/i&gt;&amp;&#039;q&#039;</td>"));
  assert.ok(scope.elements.get("qualityTable").innerHTML.includes('aria-label="查看「&lt;b&gt;备&amp;注&quot;&lt;/b&gt;」的空值行"'));
  // A category with nothing to show (编号 has no duplicates) leaves the open view as it is.
  scope.evaluate('openProblemView({ kind: "idDuplicate", field: "编号" })');
  assert.equal(openCategory(scope).kind, "missing");
});

test("the problem-row viewer is static in index.html, in the dom map and bound in script.js", () => {
  const html = indexHtmlSource();
  const quality = html.slice(html.indexOf('id="qualitySection"'), html.indexOf('id="filterSection"'));
  const order = ['id="qualitySummary"', 'id="parseWarningSummary"', 'id="outlierTable"', 'id="problemRowsSection"', 'id="numberConversionSection"'].map((marker) => quality.indexOf(marker));
  assert.ok(order.every((index, position) => index > 0 && (position === 0 || index > order[position - 1])), JSON.stringify(order));
  assert.match(quality, /<p id="parseWarningSummary" class="[^"]*\bhidden\b[^"]*"/);
  assert.match(quality, /<div id="problemRowsSection" class="[^"]*\bhidden\b[^"]*"/);
  assert.match(quality, /<h3 id="problemRowsTitle" class="subsection-title" tabindex="-1">/);
  assert.match(quality, /<p id="problemRowsSummary" class="[^"]*" aria-live="polite"><\/p>/);
  const ids = ["qualitySection", "parseWarningSummary", "problemRowsSection", "problemRowsTitle", "problemRowsCategory", "problemRowsSummary", "problemRowsDraftNote", "problemRowsTable",
    "problemRowsFirstPage", "problemRowsPrevPage", "problemRowsNextPage", "problemRowsLastPage", "problemRowsClose", "problemRowsMessage"];
  ids.forEach((id) => assert.match(html, new RegExp(`id="${id}"`), id));
  const { evaluate } = createScriptContext();
  assert.equal(evaluate(`${JSON.stringify(ids)}.every((id) => Object.keys(dom).includes(id))`), true);
  const script = fs.readFileSync(path.join(__dirname, "..", "script.js"), "utf8");
  assert.ok(script.includes('dom.qualitySection.addEventListener("click", handleProblemViewClick);'));
  assert.ok(script.includes('dom.problemRowsCategory.addEventListener("change", () => selectProblemCategory(dom.problemRowsCategory.value));'));
  assert.ok(script.includes('dom.problemRowsClose.addEventListener("click", closeProblemView);'));
  for (const [id, action] of [["problemRowsFirstPage", "first"], ["problemRowsPrevPage", "prev"], ["problemRowsNextPage", "next"], ["problemRowsLastPage", "last"]]) {
    assert.ok(script.includes(`dom.${id}.addEventListener("click", () => goToProblemRowsPage("${action}"));`), id);
  }
  // The placeholder buttons keep their own handler, untouched.
  assert.ok(script.includes('dom.placeholderDetails.addEventListener("click", handlePlaceholderClick);'));
});

test("without numeric fields the viewer still lists its category again or closes", async () => {
  // renderQuality() returns early when there is no numeric field: the viewer must refresh before that.
  const csv = `城市,备注\r\n${Array.from({ length: 120 }, (_, index) => `${PREVIEW_CITIES[index % 3]},${index % 2 ? "" : index % 4 ? "N/A" : `说明${index}`}`).join("\r\n")}\r\n`;
  const scope = await importCsvFile(csv, "categories.csv");
  assert.match(scope.elements.get("outlierTable").innerHTML, /未识别到数值字段/);
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(state.profiles.map((profile) => profile.typeKey))")), ["category", "category"]);
  scope.evaluate('openProblemView({ kind: "missing", field: "备注" })');
  scope.evaluate('goToProblemRowsPage("next")');
  assert.equal(viewerState(scope).summary, "共 60 行，第 2 / 2 页");

  scope.evaluate('confirmPlaceholder("备注", "N/A")');
  assert.equal(viewerState(scope).summary, "共 90 行，第 1 / 2 页");
  scope.evaluate('goToProblemRowsPage("next")');
  scope.evaluate('undoPlaceholderConfirmation("备注", "N/A")');
  assert.equal(viewerState(scope).summary, "共 60 行，第 1 / 2 页");

  // Ignoring 备注 drops its pending N/A entry, so that category closes with the note.
  scope.evaluate('openProblemView({ kind: "placeholder", field: "备注", key: "N/A" })');
  assert.equal(viewerState(scope).summary, "共 30 行，第 1 / 1 页");
  scope.evaluate('stageFieldTypeChange("备注", "ignore"); applyFieldConfiguration()');
  assert.equal(viewerState(scope).hidden, true);
  assert.equal(viewerState(scope).message, "这一类已经没有问题行。");
});

test("closing the viewer gives focus back to the 查看 button that opened it", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  const document = orders.context.document;
  const fakeButton = (index, connected = true) => ({
    dataset: { problemIndex: String(index) },
    isConnected: connected,
    focus() { document.activeElement = this; }
  });
  document.activeElement = null;
  const opener = fakeButton(0);
  put(orders.context, "__opener", opener);
  orders.evaluate("handleProblemViewClick({ target: { closest: () => __opener } })");
  orders.evaluate("closeProblemView()");
  assert.equal(orders.evaluate("document.activeElement === __opener"), true);

  // The opener was rendered away (a confirmation re-renders the quality page): the new
  // button of the same category takes the focus.
  const stale = fakeButton(0);
  put(orders.context, "__opener", stale);
  orders.evaluate("handleProblemViewClick({ target: { closest: () => __opener } })");
  stale.isConnected = false;
  const fresh = fakeButton(0);
  const selectors = [];
  orders.elements.get("qualitySection").querySelector = (selector) => {
    selectors.push(selector);
    return selector === 'button[data-problem-index="0"]' ? fresh : null;
  };
  orders.evaluate("closeProblemView()");
  assert.equal(document.activeElement, fresh);
  assert.deepEqual(selectors, ['button[data-problem-index="0"]']);

  // No button for the category any more: the focus stays where it was.
  document.activeElement = null;
  const gone = fakeButton(0);
  put(orders.context, "__opener", gone);
  orders.evaluate("handleProblemViewClick({ target: { closest: () => __opener } })");
  gone.isConnected = false;
  orders.elements.get("qualitySection").querySelector = () => null;
  orders.evaluate("closeProblemView()");
  assert.equal(document.activeElement, null);
  assert.equal(viewerState(orders).hidden, true);
});

// v2.6.0: "导出这一类" writes the whole open category as CSV through the existing
// download path (BOM, formula neutralizing, file name sanitizing).
function captureDownloads(scope) {
  const downloads = [];
  scope.context.URL = {
    createObjectURL: (blob) => {
      downloads.push({ blob });
      return "blob:problem-rows";
    },
    revokeObjectURL() {}
  };
  scope.context.document.createElement = () => ({
    click() { downloads[downloads.length - 1].fileName = this.download; },
    remove() {}
  });
  return async () => Promise.all(downloads.map(async ({ blob, fileName }) => {
    const bytes = new Uint8Array(await blob.arrayBuffer());
    return { fileName, type: blob.type, bom: Array.from(bytes.slice(0, 3)), text: new TextDecoder("utf-8").decode(bytes.slice(3)) };
  }));
}

const EXPORT_CSV = "编号,金额,备注,说明,文件行号,=x\r\nR1,100,=1+1,s1,f1,a\r\nR2,,=cmd,s2,f2,b\r\nR3,-2200元,,s3,f3,c\r\nR4,-42,,s4,f4,d\r\n";

test("exporting a problem category writes its rows with file rows, notes and every field", async () => {
  const scope = await importCsvFile(EXPORT_CSV, "导出.csv");
  const downloads = captureDownloads(scope);
  scope.evaluate('openProblemView({ kind: "missing", field: "备注" })');
  scope.evaluate("exportProblemRows()");
  scope.evaluate('openProblemView({ kind: "missing", field: "金额" })');
  scope.evaluate("exportProblemRows()");
  const [remarks, amounts] = await downloads();
  const day = scope.evaluate("formatReportFileDate(new Date())");
  assert.deepEqual(remarks.bom, [0xef, 0xbb, 0xbf]);
  assert.equal(remarks.type, "text/csv;charset=utf-8");
  assert.equal(remarks.fileName, `导出_问题行_空值_备注_${day}.csv`);
  // Extra columns that clash with field names get （2）; formulas, headers included, are
  // neutralized, and -2200元 keeps its apostrophe while -42 stays a number.
  assert.equal(remarks.text, [
    "文件行号（2）,说明（2）,编号,金额,备注,说明,文件行号,'=x",
    "4,空值,R3,'-2200元,,s3,f3,c",
    "5,空值,R4,-42,,s4,f4,d"
  ].join("\r\n"));
  assert.equal(amounts.text.split("\r\n")[1], "3,空值,R2,,'=cmd,s2,f2,b");
  assert.equal(scope.elements.get("problemRowsExportStatus").textContent, `已导出 1 条：导出_问题行_空值_金额_${day}.csv`);
  assert.match(scope.elements.get("problemRowsExportStatus").className, /success/);
});

test("problem-row exports cover the whole category, not the current page", async () => {
  const csv = `编号,备注\r\n${Array.from({ length: 120 }, (_, index) => `R${index + 1},${index % 2 ? "" : `说明${index}`}`).join("\r\n")}\r\n`;
  const scope = await importCsvFile(csv, "notes.csv");
  const downloads = captureDownloads(scope);
  scope.evaluate('openProblemView({ kind: "missing", field: "备注" })');
  scope.evaluate('goToProblemRowsPage("last")');
  scope.evaluate("exportProblemRows()");
  const [download] = await downloads();
  const lines = download.text.split("\r\n");
  assert.equal(lines.length, 61);
  assert.equal(lines[1], "3,空值,R2,");
  assert.equal(lines[60], "121,空值,R120,");
  assert.equal(scope.elements.get("problemRowsExportStatus").textContent.startsWith("已导出 60 条："), true);
});

test("parse warnings without an imported row export with empty fields", async () => {
  const scope = await importCsvFile("a,b\n1,2\n\n \n3\n4,5\n", "warnings.csv");
  const downloads = captureDownloads(scope);
  scope.evaluate('openProblemView({ kind: "parseWarnings" })');
  scope.evaluate("exportProblemRows()");
  const [download] = await downloads();
  assert.deepEqual(download.text.split("\r\n"), [
    "文件行号,说明,a,b",
    "4,第 4 行只有 1 列（表头有 2 列），缺少的单元格按空值处理。（这一行没有可分析的数据，未导入）,,",
    "5,第 5 行只有 1 列（表头有 2 列），缺少的单元格按空值处理。,3,"
  ]);
  assert.equal(download.fileName, `warnings_问题行_读取有疑问_${scope.evaluate("formatReportFileDate(new Date())")}.csv`);
});

test("problem-row export names keep the date after sanitizing and shortening the rest", () => {
  const { evaluate } = createScriptContext();
  evaluate('state.sourceFileName = "订单表.csv"');
  const name = (category) => evaluate(`buildProblemRowsFileName(${JSON.stringify(category)}, new Date(2026, 9, 29))`);
  assert.equal(name({ kind: "placeholder", field: "金额", key: "N/A" }), "订单表_问题行_像是空值的写法_金额_N-A_20261029.csv");
  assert.equal(name({ kind: "duplicateRows" }), "订单表_问题行_整行重复_20261029.csv");
  assert.equal(name({ kind: "idDuplicate", field: "订单号" }), "订单表_问题行_编号重复_订单号_20261029.csv");
  assert.equal(name({ kind: "conversion", field: "a:b" }), "订单表_问题行_认不出_a-b_20261029.csv");
  const long = name({ kind: "outlier", field: "长".repeat(100) });
  assert.ok(long.endsWith("_20261029.csv"), long);
  assert.equal(long, `${`订单表_问题行_偏大或偏小_${"长".repeat(100)}`.slice(0, 80)}_20261029.csv`);
});

test("problem-row export headers number the extra columns past every clashing field name", () => {
  const { evaluate } = createScriptContext();
  const headers = (fields) => JSON.parse(evaluate(`JSON.stringify(buildProblemRowsCsvHeaders(${JSON.stringify(fields)}))`));
  assert.deepEqual(headers(["a", "b"]), ["文件行号", "说明", "a", "b"]);
  assert.deepEqual(headers(["文件行号", "文件行号（2）", "说明"]), ["文件行号（3）", "说明（2）", "文件行号", "文件行号（2）", "说明"]);
});

test("guard: the data exports keep their exact output through the shared CSV builder", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  const expected = (rows) => orders.evaluate(`window.Papa.unparse({ fields: state.fields.map(sanitizeCsvExportValue), data: ${rows}.map((row) => state.fields.map((field) => (Object.prototype.hasOwnProperty.call(row, field) && row[field] !== null && row[field] !== undefined ? sanitizeCsvExportValue(row[field]) : ""))) }, { header: true, newline: "\\r\\n", skipEmptyLines: false })`);
  assert.equal(orders.evaluate('buildCsvExportContent(getProcessedExportRows("filtered"), state.fields)'), expected('getProcessedExportRows("filtered")'));
  // v3.1.0: with rows left out, the 40 counted rows, still byte for byte the old way.
  const sample = await sampleWithExclusions();
  assert.equal(sample.evaluate('getProcessedExportRows("filtered").length'), 40);
  const sampleExpected = sample.evaluate(`window.Papa.unparse({ fields: state.fields.map(sanitizeCsvExportValue), data: getProcessedExportRows("filtered").map((row) => state.fields.map((field) => (Object.prototype.hasOwnProperty.call(row, field) && row[field] !== null && row[field] !== undefined ? sanitizeCsvExportValue(row[field]) : ""))) }, { header: true, newline: "\\r\\n", skipEmptyLines: false })`);
  assert.equal(sample.evaluate('buildCsvExportContent(getProcessedExportRows("filtered"), state.fields)'), sampleExpected);
});

test("the export button needs an open category and its status clears with the viewer", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  const downloads = captureDownloads(orders);
  orders.evaluate("exportProblemRows()");
  assert.equal(orders.elements.get("problemRowsExportStatus").textContent, "请先打开一类问题行。");
  assert.deepEqual(await downloads(), []);
  orders.evaluate('openProblemView({ kind: "missing", field: "备注" })');
  orders.evaluate("exportProblemRows()");
  assert.match(orders.elements.get("problemRowsExportStatus").textContent, /^已导出 26 条：/);
  orders.evaluate('openProblemView({ kind: "conversion", field: "金额" })');
  assert.equal(orders.elements.get("problemRowsExportStatus").textContent, "");
  orders.evaluate("exportProblemRows(); closeProblemView()");
  assert.equal(orders.elements.get("problemRowsExportStatus").textContent, "");

  const html = indexHtmlSource();
  const section = html.slice(html.indexOf('id="problemRowsSection"'), html.indexOf('id="numberConversionSection"'));
  assert.match(section, /<button id="problemRowsExport" class="secondary-button" type="button">导出这一类<\/button>/);
  assert.match(section, /<p id="problemRowsExportStatus" class="inline-status" aria-live="polite"><\/p>/);
  const { evaluate } = createScriptContext();
  assert.equal(evaluate('["problemRowsExport", "problemRowsExportStatus"].every((id) => Object.keys(dom).includes(id))'), true);
  assert.ok(fs.readFileSync(path.join(__dirname, "..", "script.js"), "utf8").includes('dom.problemRowsExport.addEventListener("click", exportProblemRows);'));
});

// v2.6.1: numericStats and qualityNumericStats share one array while no row is filtered out.

const STATS_REUSE_COLUMNS = {
  类别: ["甲", "乙", "甲", "乙", "甲", "乙"],
  金额: ["10", "1000", "30", "40", "50", "20"],
  数量: ["1", "2", "3", "4", "5", "6"]
};

function statsJson(scope, expression) {
  return JSON.parse(scope.evaluate(`JSON.stringify(${expression}) ?? "null"`));
}

function filterStatsByCategory(scope, field, dropValue) {
  put(scope.context, "__field", field);
  put(scope.context, "__drop", dropValue);
  scope.evaluate("dom.categoryFilterField.value = __field; updateCategoryFilterField(); state.categoryFilterDraftValues.delete(__drop); applyFilters();");
}

function emptyStatsFilterPanel(scope) {
  scope.evaluate('dom.categoryFilterField.value = ""; dom.numericFilterField.value = ""; dom.dateFilterField.value = ""; updateCategoryFilterField(); updateNumericFilterField(); updateDateFilterField();');
}

// Wraps global functions of the scripts with call counters while action() runs.
function countStatsCalls(scope, names, action) {
  const counts = Object.fromEntries(names.map((name) => [name, 0]));
  names.forEach((name) => {
    const original = scope.context[name];
    scope.context[name] = function counted(...args) { counts[name] += 1; return original.apply(this, args); };
  });
  action();
  return counts;
}

test("a filter gets its own statistics, clearing it brings the full ones back, and the quality statistics never change", () => {
  const scope = commitColumns(STATS_REUSE_COLUMNS);
  const quality = () => JSON.stringify(statsJson(scope, "state.qualityNumericStats"));
  const qualityAtImport = quality();
  assert.deepEqual(statsJson(scope, "state.numericStats"), statsJson(scope, "state.qualityNumericStats"));
  assert.equal(statsJson(scope, "state.numericStats").find((item) => item.field === "金额").max, 1000);

  // 乙 holds the 1000: with only 甲 left the maximum must drop to 50.
  filterStatsByCategory(scope, "类别", "乙");
  const filtered = statsJson(scope, "state.numericStats").find((item) => item.field === "金额");
  assert.deepEqual([filtered.count, filtered.max, filtered.median], [3, 50, 30]);
  assert.equal(statsJson(scope, "state.qualityNumericStats").find((item) => item.field === "金额").max, 1000);
  assert.equal(quality(), qualityAtImport);

  scope.evaluate("clearAllFilters()");
  assert.deepEqual(statsJson(scope, "state.numericStats"), statsJson(scope, "buildNumericStats(state.rows)"));
  assert.equal(quality(), qualityAtImport);

  filterStatsByCategory(scope, "类别", "乙");
  emptyStatsFilterPanel(scope);
  scope.evaluate("applyFilters()");
  assert.equal(scope.evaluate("hasActiveFilters()"), false);
  assert.deepEqual(statsJson(scope, "state.numericStats"), statsJson(scope, "buildNumericStats(state.rows)"));
  assert.equal(quality(), qualityAtImport);
});

test("a filter that keeps every row still shows the full statistics", () => {
  const scope = commitColumns(STATS_REUSE_COLUMNS);
  scope.evaluate('dom.numericFilterField.value = "金额"; updateNumericFilterField(); applyFilters();');
  assert.equal(scope.evaluate("hasActiveFilters()"), true);
  assert.equal(scope.evaluate("getAnalysisRows().length"), 6);
  assert.deepEqual(statsJson(scope, "state.numericStats"), statsJson(scope, "buildNumericStats(state.rows)"));
});

test("field configuration rebuilds both statistics, also while a filter is applied", () => {
  const scope = commitColumns(STATS_REUSE_COLUMNS);
  filterStatsByCategory(scope, "类别", "乙");
  scope.evaluate('stageFieldTypeChange("金额", "ignore"); applyFieldConfiguration()');
  assert.deepEqual(statsJson(scope, "state.numericStats.map((item) => item.field)"), ["数量"]);
  assert.deepEqual(statsJson(scope, "state.qualityNumericStats.map((item) => item.field)"), ["数量"]);
  assert.deepEqual(statsJson(scope, "state.numericStats"), statsJson(scope, "buildNumericStats(state.rows)"));
  scope.evaluate("restoreAutomaticFieldTypes()");
  assert.deepEqual(statsJson(scope, "state.numericStats.map((item) => item.field)"), ["金额", "数量"]);
  assert.deepEqual(statsJson(scope, "state.numericStats"), statsJson(scope, "state.qualityNumericStats"));
});

// Guards for the speed-up itself: they pin how often the statistics are built, not what they say.
test("importing and clearing a filter build the numeric statistics as few times as needed", () => {
  const scope = createScriptContext();
  const fields = Object.keys(STATS_REUSE_COLUMNS);
  const rows = STATS_REUSE_COLUMNS.类别.map((_, index) => Object.fromEntries(fields.map((field) => [field, STATS_REUSE_COLUMNS[field][index]])));
  put(scope.context, "__fields", fields);
  put(scope.context, "__rows", rows);
  const commit = countStatsCalls(scope, ["buildNumericStats"], () => scope.evaluate('commitTabularData(__fields, __rows, "count.csv", beginImport())'));
  assert.equal(commit.buildNumericStats, 1);
  const filter = countStatsCalls(scope, ["buildNumericStats"], () => filterStatsByCategory(scope, "类别", "乙"));
  assert.equal(filter.buildNumericStats, 1);
  const clear = countStatsCalls(scope, ["buildNumericStats"], () => scope.evaluate("clearAllFilters()"));
  assert.equal(clear.buildNumericStats, 0);
  const rebuild = countStatsCalls(scope, ["buildNumericStats"], () => scope.evaluate('stageFieldTypeChange("数量", "ignore"); applyFieldConfiguration()'));
  assert.equal(rebuild.buildNumericStats, 1);
});

test("the statistics table takes the median from its sorted values instead of sorting again", () => {
  const scope = commitColumns(STATS_REUSE_COLUMNS);
  const calls = countStatsCalls(scope, ["median"], () => scope.evaluate("buildNumericStats(state.rows)"));
  assert.equal(calls.median, 0);
});

// v2.6.1: the category chart starts on the first concentrated category field.

function categoryChartValue(scope) {
  return scope.elements.get("categoryChartField").value;
}

// Field types the test data really gets, so a field that turns out to be an ID cannot hide a mistake.
function chartTestTypes(scope) {
  return statsJson(scope, "Object.fromEntries(state.profiles.map((profile) => [profile.field, profile.typeKey]))");
}

function recommitColumns(scope, columns) {
  const fields = Object.keys(columns);
  put(scope.context, "__fields", fields);
  put(scope.context, "__rows", columns[fields[0]].map((_, index) => Object.fromEntries(fields.map((field) => [field, columns[field][index]]))));
  scope.evaluate('commitTabularData(__fields, __rows, "second.csv", beginImport())');
}

test("the category chart starts on the first concentrated category field, and the conclusion under it is about that field", () => {
  const scope = commitColumns({
    姓名: ["u1", "u2", "u3", "u4", "u5", "u6", "u7", "u8", "u9", "u10"],
    部门: ["A", "A", "B", "B", "A", "A", "B", "A", "B", "A"]
  });
  assert.deepEqual(chartTestTypes(scope), { 姓名: "category", 部门: "category" });
  assert.equal(categoryChartValue(scope), "部门");
  assert.match(scope.evaluate("state.chartInsights.category"), /分类字段「部门」中占比最高/);
});

test("the default needs a top category that repeats: exactly twice is enough, once is not", () => {
  const scope = commitColumns({
    甲项: ["u1", "u2", "u3", "u4", "u5", "u6", "u7", "u8", "u9", "u10"],
    成对: ["a", "a", "b", "b", "c", "c", "d", "d", "e", "e"]
  });
  assert.deepEqual(chartTestTypes(scope), { 甲项: "category", 成对: "category" });
  assert.equal(categoryChartValue(scope), "成对");
});

test("the default allows 20 distinct values and not 21", () => {
  const twice = (distinct) => Array.from({ length: 42 }, (_, index) => `v${index % distinct}`);
  const scope = commitColumns({ 区域二十一: twice(21), 区域二十: twice(20) });
  assert.deepEqual(chartTestTypes(scope), { 区域二十一: "category", 区域二十: "category" });
  assert.deepEqual(statsJson(scope, "state.categoryStats.map((item) => [item.field, item.distinctCount])"), [["区域二十一", 21], ["区域二十", 20]]);
  assert.equal(categoryChartValue(scope), "区域二十");
});

test("without a concentrated category field the first category field is kept, and without any it is empty", () => {
  const scattered = commitColumns({
    甲: ["u1", "u2", "u3", "u4", "u5", "u6", "u7", "u8", "u9", "u10"],
    乙: ["w1", "w2", "w3", "w4", "w5", "w6", "w7", "w8", "w9", "w10"]
  });
  assert.deepEqual(chartTestTypes(scattered), { 甲: "category", 乙: "category" });
  assert.equal(categoryChartValue(scattered), "甲");
  const numericOnly = commitColumns({ 数量: ["1", "2", "3", "4", "5", "6"] });
  assert.deepEqual(chartTestTypes(numericOnly), { 数量: "numeric" });
  assert.equal(categoryChartValue(numericOnly), "");
});

test("a field-configuration rebuild keeps a chosen category field that is still a category; a new import starts over", () => {
  const columns = {
    甲: ["u1", "u2", "u3", "u4", "u5", "u6", "u7", "u8", "u9", "u10"],
    乙: ["a", "a", "b", "b", "a", "a", "b", "a", "b", "a"],
    丙: ["x", "x", "y", "y", "y", "x", "x", "y", "y", "x"],
    数量: ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"]
  };
  const scope = commitColumns(columns);
  assert.deepEqual(chartTestTypes(scope), { 甲: "category", 乙: "category", 丙: "category", 数量: "numeric" });
  assert.equal(categoryChartValue(scope), "乙");
  scope.evaluate('dom.categoryChartField.value = "丙"');
  scope.evaluate('stageFieldTypeChange("数量", "ignore"); applyFieldConfiguration()');
  assert.equal(categoryChartValue(scope), "丙");
  // The chosen field leaves the category fields: back to the rule.
  scope.evaluate('stageFieldTypeChange("丙", "ignore"); applyFieldConfiguration()');
  assert.equal(categoryChartValue(scope), "乙");
  // A new import never inherits the old choice.
  scope.evaluate('dom.categoryChartField.value = "甲"');
  recommitColumns(scope, columns);
  assert.equal(categoryChartValue(scope), "乙");
});

// ---------------------------------------------------------------------------------------
// v2.7.0: removals and the two chart conclusions.
// ---------------------------------------------------------------------------------------
const REPO_ROOT = path.join(__dirname, "..");

test("v2.7.0 examples: examples/ keeps only the file the sample button loads, the other six are test fixtures (v2.11.0)", () => {
  assert.deepEqual(fs.readdirSync(path.join(REPO_ROOT, "examples")).sort(), ["sample-orders-en.csv", "sample-orders.csv"]);
  assert.deepEqual(fs.readdirSync(path.join(REPO_ROOT, "tests", "fixtures", "examples")).sort(), [
    "sample-data.csv", "sample-sales.csv", "sample-students.csv", "sample-survey.csv", "sample-used-products.csv", "sample-user-behavior.csv"
  ]);
  assert.deepEqual(["examples", path.join("tests", "fixtures", "examples")].flatMap((dir) => fs.readdirSync(path.join(REPO_ROOT, dir))).sort(), [
    "sample-data.csv", "sample-orders-en.csv", "sample-orders.csv", "sample-sales.csv", "sample-students.csv", "sample-survey.csv", "sample-used-products.csv", "sample-user-behavior.csv"
  ]);
});

const repoFile = (name) => fs.readFileSync(path.join(REPO_ROOT, name), "utf8");
const domKeys = (scope) => JSON.parse(scope.evaluate("JSON.stringify(Object.keys(dom))"));
const typesOf = (scope, names) => JSON.parse(scope.evaluate("JSON.stringify(" + JSON.stringify(names) + ".map((name) => typeof globalThis[name]))"));

test("v2.7.0 workflow: the workflow status bar is gone from the page, the styles and the scripts", () => {
  const scope = createScriptContext();
  assert.doesNotMatch(indexHtmlSource(), /workflow-status|workflowStep/);
  assert.doesNotMatch(repoFile("style.css"), /workflow-|\.step-(?:number|copy|state)/);
  assert.deepEqual(typesOf(scope, ["updateWorkflowSteps"]), ["undefined"]);
  assert.deepEqual(domKeys(scope).filter((key) => key.startsWith("workflowStep")), []);
});

test("v2.7.0 quick grouping: the advanced settings block, its chart and its report fallback are gone", () => {
  const scope = createScriptContext();
  assert.doesNotMatch(indexHtmlSource(), /id="(?:metricField|groupField|aggregateMethod|customAnalysisChart|customAnalysisChartTitle)"|高级设置|快速分组/);
  assert.deepEqual(typesOf(scope, ["renderCustomAnalysisChart"]), ["undefined"]);
  assert.deepEqual(domKeys(scope).filter((key) => ["metricField", "groupField", "aggregateMethod", "customAnalysisChartTitle"].includes(key)), []);
  assert.equal(scope.evaluate('"customAnalysisChart" in REPORT_CHART_TITLES || "customAnalysisChart" in REPORT_CHART_HEADING_IDS'), false);
  assert.doesNotMatch(repoFile("script.js"), /dom\.(?:metricField|groupField|aggregateMethod)/);
});

test("v2.7.0 quick grouping: the report has a custom-analysis section only after the custom analysis was generated", () => {
  const scope = createChartContext();
  commitExampleCsv(scope, "sample-data.csv");
  // Before v2.7.0 the selects of the quick-grouping block (picked by the browser) built the report's
  // section on their own; this line makes the fake DOM do what the browser does.
  scope.evaluate('if (typeof dom.metricField !== "undefined") { dom.metricField.value = "销售额"; dom.groupField.value = "地区"; dom.aggregateMethod.value = "sum"; }');
  assert.equal(scope.evaluate("buildReportData().customAnalysis.completed"), false);
  scope.evaluate('dom.v2MetricField.value = "销售额"; dom.v2GroupField.value = "地区"; dom.v2AggregateMethod.value = "sum"; dom.v2DateField.value = ""; renderV2Analysis()');
  assert.equal(scope.evaluate("buildReportData().customAnalysis.completed"), true);
});

test("v2.7.0 templates: templates.js, its markup, styles, startup check and report entries are gone", () => {
  const scope = createScriptContext();
  const html = indexHtmlSource();
  assert.equal(fs.existsSync(path.join(REPO_ROOT, "templates.js")), false);
  assert.doesNotMatch(html, /templates\.js|scenarioTemplate|使用分析模板|id="template[A-Z]/);
  assert.doesNotMatch(repoFile("style.css"), /\.template-/);
  assert.deepEqual(typesOf(scope, ["runTemplateAnalysis", "renderTemplateMappingForm", "resetTemplateResults", "buildTemplateFieldOptions"]), ["undefined", "undefined", "undefined", "undefined"]);
  assert.deepEqual(domKeys(scope).filter((key) => /template/i.test(key)), []);
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(Object.keys(REPORT_CHART_TITLES))")), [
    "numericDistributionChart", "categoryTopChart", "dateTrendChart", "v2TopGroupChart", "v2SelectedTrendChart"
  ]);
  // The startup check names one representative function per remaining app script.
  assert.deepEqual([...repoFile("script.js").matchAll(/\["([a-z]+\.js)", typeof /g)].map((match) => match[1]), [
    "parsing.js", "profiling.js", "format.js", "stats.js", "quality.js", "app.js", "importer.js", "report.js"
  ]);
});

test("v2.7.0 conclusions: the findings section and its navigation entries are gone, the conclusions sit in the chart cards", () => {
  const html = indexHtmlSource();
  assert.doesNotMatch(html, /insightsSection|id="insights"|主要发现|分析结论<\/a>/);
  // v3.0.0: one navigation of exactly 8 links, in module order and without numbers; the phone
  // select is gone.
  const navBlocks = [...html.matchAll(/<nav\b[^>]*>([\s\S]*?)<\/nav>/g)].map((match) => match[1]).filter((block) => /href="#\w+Section"/.test(block));
  assert.equal(navBlocks.length, 1);
  const navHrefs = [...navBlocks[0].matchAll(/<a\b[^>]*href="#(\w+)"[^>]*>([^<]*)<\/a>/g)].map((match) => match[1]);
  assert.deepEqual(navHrefs, ["overviewSection", "schemaSection", "qualitySection", "filterSection", "statisticsSection", "chartsSection", "customSection", "exportSection"]);
  assert.equal((navBlocks[0].match(/<a\b/g) || []).length, 8);
  assert.doesNotMatch(navBlocks[0], /<span>\d\d<\/span>/);
  assert.doesNotMatch(html, /id="mobileSectionNav"/);
  const chartCard = (canvasId) => html.slice(html.lastIndexOf('<div class="chart-card', html.indexOf('id="' + canvasId + '"')), html.indexOf('id="' + canvasId + '"') + 400);
  assert.match(chartCard("categoryTopChart"), /id="categoryChartInsight" class="chart-insight hidden"/);
  assert.match(chartCard("dateTrendChart"), /id="dateChartInsight" class="chart-insight hidden"/);
  assert.match(repoFile("script.js"), /dom\.categoryChartField\.addEventListener\("change", renderCategoryInsight\)/);
  assert.match(repoFile("script.js"), /dom\.dateChartField\.addEventListener\("change", renderTrendInsight\)/);
  assert.deepEqual(typesOf(createScriptContext(), ["renderInsights", "renderInsightGroups"]), ["undefined", "undefined"]);
});

const CONCLUSION_ROWS = Array.from({ length: 30 }, (_, index) => ({
  部门: ["A", "A", "A", "B", "B", "C"][index % 6],
  品名: "k" + (index % 25),
  签约日期: "2025-" + String(1 + (index % 6)).padStart(2, "0") + "-" + String(5 + (index % 3)).padStart(2, "0"),
  回访日期: index < 12 ? "2025-0" + (1 + (index % 4)) + "-10" : ""
}));
const conclusions = (scope) => JSON.parse(scope.evaluate("JSON.stringify(state.chartInsights)"));
function filterBy(scope, field, values) {
  put(scope.context, "__conclusionFilter", { field, values });
  scope.evaluate("state.filters = { category: __conclusionFilter, numeric: null, date: null }; state.filteredRows = filterRowsByConfig(state.rows, state.filters); state.filterSourceRows = state.rows; refreshFilteredAnalysis();");
}
const conclusionShown = (scope, id) => !scope.elements.get(id).classList.contains("hidden");

test("v2.7.0 conclusions: each sentence follows its own chart select, a field change redoes only its sentence", () => {
  const scope = createScriptContext();
  commitRows(scope, ["部门", "品名", "签约日期", "回访日期"], CONCLUSION_ROWS, "conclusions.csv");
  assert.equal(scope.evaluate("dom.categoryChartField.value"), "部门");
  assert.equal(scope.evaluate("dom.dateChartField.value"), "签约日期");
  const start = conclusions(scope);
  assert.equal(start.category, "分类字段「部门」中占比最高的类别是「A」，出现 15 次，占比 50.0%。");
  assert.equal(start.trend, scope.evaluate('buildTrendInsight("签约日期")'));
  assert.equal(scope.elements.get("categoryChartInsight").textContent, start.category);
  assert.equal(scope.elements.get("dateChartInsight").textContent, start.trend);
  assert.equal(conclusionShown(scope, "categoryChartInsight") && conclusionShown(scope, "dateChartInsight"), true);

  scope.evaluate('dom.categoryChartField.value = "品名"; renderCategoryInsight()');
  assert.deepEqual(conclusions(scope), { category: "「品名」取值比较分散（超过 20 个不同值，或没有重复出现的值），不做概括。", trend: start.trend });
  assert.equal(scope.elements.get("categoryChartInsight").textContent, conclusions(scope).category);

  scope.evaluate('dom.dateChartField.value = "回访日期"; renderTrendInsight()');
  assert.equal(conclusions(scope).trend, scope.evaluate('buildTrendInsight("回访日期")'));
  assert.notEqual(conclusions(scope).trend, start.trend);
  assert.match(conclusions(scope).category, /^「品名」取值比较分散/);

  // The metric and aggregate selects of the trend chart do not move the sentence.
  scope.evaluate('dom.dateChartMetric.value = ""; dom.dateChartAggregate.value = "sum"; renderCharts()');
  assert.equal(conclusions(scope).trend, scope.evaluate('buildTrendInsight("回访日期")'));
});

test("v2.7.0 conclusions: import, filter, field configuration and placeholder confirmation all refresh them", () => {
  const scope = createScriptContext();
  const rows = CONCLUSION_ROWS.map((row, index) => ({ ...row, 部门: index < 14 ? "N/A" : ["A", "A", "B", "C"][index % 4] }));
  commitRows(scope, ["部门", "品名", "签约日期", "回访日期"], rows, "conclusions.csv");
  // import: N/A is still a category value until confirmed.
  assert.match(conclusions(scope).category, /^分类字段「部门」中占比最高的类别是「N\/A」，出现 14 次/);
  // placeholder confirmation (the 4th trigger): N/A drops out, the top category changes.
  scope.evaluate('confirmPlaceholder("部门", "N/A")');
  assert.match(conclusions(scope).category, /^分类字段「部门」中占比最高的类别是「A」/);
  // filter: only the rows of 品名 k0..k4; the select keeps 部门 and the sentence follows the filtered rows.
  filterBy(scope, "品名", ["k0", "k1", "k2", "k3", "k4"]);
  const filtered = scope.evaluate("getAnalysisRows().length");
  assert.ok(filtered > 0 && filtered < 30);
  assert.equal(scope.evaluate("dom.categoryChartField.value"), "部门");
  assert.equal(conclusions(scope).category, scope.evaluate("buildCategoryInsight(\"部门\")"));
  assert.equal(conclusions(scope).trend, scope.evaluate('buildTrendInsight("签约日期")'));
  // a filter that keeps nothing: no sentences, both hidden.
  filterBy(scope, "品名", ["不存在"]);
  assert.deepEqual(conclusions(scope), { category: "", trend: "" });
  assert.equal(conclusionShown(scope, "categoryChartInsight") || conclusionShown(scope, "dateChartInsight"), false);
  // field configuration (the 3rd path) clears the filter and the sentences come back; ignoring the
  // selected category field moves the select, and the sentence, to the next category field.
  scope.evaluate('stageFieldTypeChange("部门", "ignore"); applyFieldConfiguration()');
  assert.equal(scope.evaluate("dom.categoryChartField.value"), "品名");
  assert.equal(conclusions(scope).category, "「品名」取值比较分散（超过 20 个不同值，或没有重复出现的值），不做概括。");
  // every field ignored: nothing to say.
  scope.evaluate(["品名", "签约日期", "回访日期"].map((field) => 'stageFieldTypeChange("' + field + '", "ignore");').join(" ") + " applyFieldConfiguration()");
  assert.deepEqual(conclusions(scope), { category: "", trend: "" });
  assert.equal(conclusionShown(scope, "categoryChartInsight") || conclusionShown(scope, "dateChartInsight"), false);
});

test("v2.7.0 conclusions: a field with no values and a new import reset the sentences", () => {
  const scope = createScriptContext();
  commitRows(scope, ["备注", "金额"], Array.from({ length: 6 }, (_, index) => ({ 备注: "", 金额: String(index + 1) })), "empty-category.csv");
  assert.equal(scope.evaluate("dom.categoryChartField.value"), "备注");
  assert.equal(conclusions(scope).category, "「备注」当前没有可统计的取值，不做概括。");
  scope.evaluate("resetAnalysisState()");
  assert.deepEqual(conclusions(scope), { category: "", trend: "" });
  assert.equal(scope.elements.get("categoryChartInsight").textContent, "");
  assert.equal(conclusionShown(scope, "categoryChartInsight"), false);
});

test("v2.7.0 conclusions: they do not depend on Chart.js being loaded", () => {
  const withoutChart = createScriptContext();
  assert.equal(withoutChart.evaluate("window.Chart"), null);
  commitRows(withoutChart, ["部门", "品名", "签约日期", "回访日期"], CONCLUSION_ROWS, "conclusions.csv");
  const withChart = createChartContext();
  commitRows(withChart, ["部门", "品名", "签约日期", "回访日期"], CONCLUSION_ROWS, "conclusions.csv");
  assert.ok(conclusions(withoutChart).category && conclusions(withoutChart).trend);
  assert.deepEqual(conclusions(withoutChart), conclusions(withChart));
});

test("v2.7.0 conclusions: the report shows the sentences on the page, 暂无 when there are none, and old snapshots as they were", () => {
  const scope = createScriptContext();
  commitRows(scope, ["部门", "品名", "签约日期", "回访日期"], CONCLUSION_ROWS, "conclusions.csv");
  const { report, html } = reportOf(scope);
  assert.deepEqual(report.insights, [conclusions(scope).category, conclusions(scope).trend]);
  assert.ok(html.includes("<h2>自动分析结论</h2>"));
  assert.ok(html.includes("<li>" + conclusions(scope).category + "</li>"));
  // Both sentences, the category one first and the trend one right after it.
  assert.ok(html.includes('<ol class="insights"><li>' + conclusions(scope).category + "</li><li>" + conclusions(scope).trend + "</li></ol>"));
  // one sentence only
  scope.evaluate('dom.dateChartField.value = ""; renderTrendInsight()');
  assert.deepEqual(reportOf(scope).report.insights, [conclusions(scope).category]);
  // none
  scope.evaluate('stageFieldTypeChange("部门", "ignore"); stageFieldTypeChange("品名", "ignore"); stageFieldTypeChange("签约日期", "ignore"); stageFieldTypeChange("回访日期", "ignore"); applyFieldConfiguration()');
  const none = reportOf(scope);
  assert.deepEqual(none.report.insights, []);
  assert.ok(none.html.includes("暂无自动分析结论。"));
  // an old snapshot with the six v2.6.1 sentences is printed as it is; one without the field prints 暂无
  const old = ["数据集包含 30 行，当前分析使用全部记录。", "当前有 4 个字段参与分析。", "缺失值 0 个。", "数值字段「金额」的平均值为 1。", "分类字段「部门」中占比最高的类别是「A」。", "基于日期字段「签约日期」观察。"];
  put(scope.context, "__oldInsights", old);
  const normalized = JSON.parse(scope.evaluate("JSON.stringify(normalizeReportSnapshot({ insights: __oldInsights }).insights)"));
  assert.deepEqual(normalized, old);
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(normalizeReportSnapshot({}).insights)")), []);
});

test("v2.7.0 statistics helpers: the five functions only the removed features used are gone", () => {
  const scope = createScriptContext();
  assert.deepEqual(
    typesOf(scope, ["groupMetricRowsByField", "frequencyRows", "countNonMissingValues", "aggregateByCategory", "summarizeNumbers"]),
    ["undefined", "undefined", "undefined", "undefined", "undefined"]
  );
});

// v2.8.0: suspected total rows and Excel formula errors. Hints only: every existing result stays.
const knownIssues = JSON.parse(fs.readFileSync(path.join(STAGE3_FIXTURES, "known-issues.json"), "utf8"));

// A workbook from an array of rows where { e: code } is an error cell. origin moves the table away from A1.
function buildErrorWorkbook(aoa, { origin = "A1", bookType = "xlsx" } = {}) {
  const start = XLSX.utils.decode_cell(origin);
  const worksheet = XLSX.utils.aoa_to_sheet(aoa.map((row) => row.map((value) => (value && typeof value === "object" ? null : value))), { origin });
  aoa.forEach((row, r) => row.forEach((value, c) => {
    if (value && typeof value === "object") worksheet[XLSX.utils.encode_cell({ r: r + start.r, c: c + start.c })] = { t: "e", v: value.e };
  }));
  worksheet["!ref"] = XLSX.utils.encode_range({ s: start, e: { r: start.r + aoa.length - 1, c: start.c + Math.max(...aoa.map((row) => row.length)) - 1 } });
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "S");
  return XLSX.write(workbook, { type: "array", bookType });
}

const formulaErrors = (scope) => JSON.parse(scope.evaluate("JSON.stringify(state.formulaErrors)"));

test("v2.8.0 errors: the fixture's cells by field, types and missing counts match known-issues.json", async () => {
  const scope = await importStage3Fixture("公式错误.xlsx");
  const expected = knownIssues["公式错误.xlsx"];
  assert.deepEqual(formulaErrors(scope), expected.formulaErrors);
  assert.deepEqual(importedSummary(scope.evaluate).types, expected.expectedTypes.map((item) => item.expected));
  assert.equal(scope.evaluate("state.totalMissing"), expected.expectedTotals.totalMissing);
  // Rows 7 and 12 hold nothing but errors and column G has no header: known limits, not listed.
  assert.equal(scope.evaluate("state.rows.map(getSourceRowNumber).join()"), "2,3,4,5,6,8,9,10,11");
  assert.equal(scope.evaluate('state.fields.join("|")'), "产品|单价|数量|金额|利润率|备注");
  // The summary does not follow the field configuration.
  const before = scope.evaluate("JSON.stringify(state.formulaErrors)");
  scope.evaluate('stageFieldTypeChange("利润率", "ignore"); applyFieldConfiguration()');
  assert.equal(scope.evaluate("JSON.stringify(state.formulaErrors)"), before);
  scope.evaluate("restoreAutomaticFieldTypes()");
  assert.equal(scope.evaluate("JSON.stringify(state.formulaErrors)"), before);
});

test("v2.8.0 errors: a sheet with error cells imports exactly like the same sheet with blank cells", async () => {
  const withErrors = [["名", "额", "说明"], ["A", 10, "好"], ["B", { e: 7 }, { e: 42 }], ["C", 30, { e: 0 }], ["D", { e: 15 }, "好"]];
  const blank = withErrors.map((row) => row.map((value) => (value && typeof value === "object" ? null : value)));
  const summarize = (scope) => scope.evaluate("JSON.stringify({ rows: state.rows, profiles: state.profiles, stats: state.qualityNumericStats, missing: state.totalMissing, categories: listProblemCategories().filter((item) => item.kind !== \"formulaError\") })");
  const a = await importWorkbookFile(buildErrorWorkbook(withErrors));
  const b = await importWorkbookFile(buildErrorWorkbook(blank));
  assert.equal(summarize(a), summarize(b));
  assert.deepEqual(formulaErrors(a).map((item) => [item.field, item.count]), [["额", 2], ["说明", 2]]);
  assert.deepEqual(formulaErrors(b), []);
});

test("v2.8.0 errors: !ref away from A1, cells outside the table, rows and columns that drop out", async () => {
  // A8 and B8 sit left of the table: a reader that forgets the column offset reads them for
  // the first column, and one that forgets the row offset reads other rows' cells.
  const scope = await importWorkbookFile((() => {
    const worksheet = XLSX.utils.aoa_to_sheet([["名", "额", "备注"], ["A", 10, "好"], ["B", null, null], ["C", 30, "好"]], { origin: "C5" });
    worksheet.D7 = { t: "e", v: 7 };
    worksheet.E7 = { t: "e", v: 42 };
    worksheet.A8 = { t: "e", v: 15 };
    worksheet.B8 = { t: "e", v: 15 };
    worksheet["!ref"] = "C5:E8";
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "S");
    return XLSX.write(workbook, { type: "array", bookType: "xlsx" });
  })());
  assert.equal(scope.evaluate("state.rows.map(getSourceRowNumber).join()"), "6,7,8");
  assert.deepEqual(formulaErrors(scope), [
    { field: "额", count: 1, byError: { "#DIV/0!": 1 }, lines: [7] },
    { field: "备注", count: 1, byError: { "#N/A": 1 }, lines: [7] }
  ]);
  // Known limits: whole rows of errors and error-only columns without a header are not listed.
  const dropped = await importWorkbookFile(buildErrorWorkbook([["名", "额", null], ["A", 1, { e: 7 }], [{ e: 7 }, { e: 42 }], ["B", 2], [{ e: 7 }, { e: 7 }]]));
  assert.equal(dropped.evaluate('state.fields.join("|")'), "名|额");
  assert.equal(dropped.evaluate("state.rows.length"), 2);
  assert.deepEqual(formulaErrors(dropped), []);
  // At most ten row numbers per field; the count stays whole.
  const many = await importWorkbookFile(buildErrorWorkbook([["名", "额"], ["A", 1], ...Array.from({ length: 12 }, (_, index) => [`N${index}`, { e: 7 }])]));
  assert.deepEqual(formulaErrors(many), [{ field: "额", count: 12, byError: { "#DIV/0!": 12 }, lines: [3, 4, 5, 6, 7, 8, 9, 10, 11, 12] }]);
});

test("v2.8.0 errors: the text comes from cell.w of error cells only, then the code table, never #WTF?", async () => {
  const { evaluate } = createExcelContext();
  assert.deepEqual(JSON.parse(evaluate(`JSON.stringify([
    readExcelErrorText({ t: "e", v: 7, w: "#DIV/0!" }),
    readExcelErrorText({ t: "e", w: "#SPILL!" }),
    readExcelErrorText({ t: "e", v: 42, w: "2024-01-01" }),
    readExcelErrorText({ t: "e", v: 0 }),
    readExcelErrorText({ t: "e", v: 99 }),
    readExcelErrorText({ t: "e", w: "not-an-error" }),
    readExcelErrorText({ t: "e", v: 255, w: "#WTF?" }),
    readExcelErrorText({ t: "e", w: "#WTF?" })
  ])`)), ["#DIV/0!", "#SPILL!", "#N/A", "#NULL!", "#错误", "#错误", "#错误", "#错误"]);
  // SheetJS's own placeholder reaches cell.w through a real file too.
  const placeholder = await importWorkbookFile(buildErrorWorkbook([["名", "额"], ["A", 1], ["B", { e: 255 }], ["C", 2]]));
  assert.deepEqual(formulaErrors(placeholder), [{ field: "额", count: 1, byError: { "#错误": 1 }, lines: [3] }]);
  // Empty-looking cells of other types are not errors, whatever their w says.
  const scope = createExcelContext();
  put(scope.context, "__sheet", { "!ref": "A1:B3", A1: { t: "s", v: "名" }, B1: { t: "s", v: "额" }, A2: { t: "s", v: "x" }, B2: { t: "s", v: "", w: "#DIV/0!" }, A3: { t: "s", v: "y" }, B3: { t: "z", w: "#N/A" } });
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify((() => { const read = createExcelCellErrorReader(__sheet); return [read(1, 1), read(2, 1), read(1, 0), read(9, 9)]; })())")), ["", "", "", ""]);
  // The reader is asked only about cells that read as empty.
  const asked = JSON.parse(scope.evaluate(`JSON.stringify((() => {
    const seen = [];
    buildExcelTabularData([["名", "额"], ["x", 1], ["y", ""]], [["名", "额"], ["x", 1], ["y", ""]], "S", null, 1, (row, column) => { seen.push([row, column]); return ""; });
    return seen;
  })())`));
  assert.deepEqual(asked, [[2, 1]]);
});

test("v2.8.0 errors: xls files read the same, CSV text stays text", async () => {
  const xls = await importWorkbookFile(buildErrorWorkbook([["名", "额"], ["A", 1], ["B", { e: 7 }], ["C", { e: 0 }], ["D", 2]], { bookType: "xls" }), "t.xls");
  assert.deepEqual(formulaErrors(xls), [{ field: "额", count: 2, byError: { "#DIV/0!": 1, "#NULL!": 1 }, lines: [3, 4] }]);
  // A format problem and a placeholder, no formula error.
  const csv = await importCsvFile("名,额,说明\r\nA,1,#N/A\r\nB,2,好\r\nC,3,好\r\nD,4,好\r\nE,#DIV/0!,好\r\n");
  assert.deepEqual(formulaErrors(csv), []);
  assert.equal(csv.evaluate('getFieldProfile("额").conversionFailureCount'), 1);
  assert.equal(csv.evaluate("state.placeholders.map((entry) => entry.key).join()"), "#N/A");
});

test("v2.8.0 state: a later commit and resetAnalysisState() clear the formula errors", async () => {
  const scope = await importStage3Fixture("公式错误.xlsx");
  assert.equal(scope.evaluate("state.sourceFormulaErrors.has(state.rows[1])"), true);
  assert.equal(scope.evaluate("state.sourceFormulaErrors.has(state.rows[0])"), false);
  put(scope.context, "__oldRow", scope.evaluate("state.rows[1]"));
  put(scope.context, "__cityRows", [{ 城市: "上海" }]);
  // The same import id, so resetAnalysisState() in beginImport() cannot hide a commit that keeps them.
  scope.evaluate('commitTabularData(["城市"], __cityRows, "later.csv", state.activeImportId)');
  assert.equal(scope.evaluate("state.sourceFormulaErrors.has(__oldRow)"), false);
  assert.equal(scope.evaluate("state.formulaErrors.length"), 0);
  // A list of another length is not aligned with the rows: no error is kept.
  put(scope.context, "__twoRows", [{ 城市: "上海" }, { 城市: "北京" }]);
  scope.evaluate('commitTabularData(["城市"], __twoRows, "short.csv", state.activeImportId, [], "", null, null, null, [{ 城市: "#N/A" }])');
  assert.equal(scope.evaluate("state.sourceFormulaErrors.has(__twoRows[0])"), false);
  assert.equal(scope.evaluate("state.formulaErrors.length"), 0);

  const again = await importStage3Fixture("公式错误.xlsx");
  put(again.context, "__keptRow", again.evaluate("state.rows[1]"));
  again.evaluate("resetAnalysisState()");
  assert.equal(again.evaluate("state.sourceFormulaErrors.has(__keptRow)"), false);
  assert.equal(again.evaluate("state.formulaErrors.length"), 0);
});

const suspectedTotals = (scope) => JSON.parse(scope.evaluate("JSON.stringify(state.suspectedTotals)"));
const totalLines = (scope) => suspectedTotals(scope).map((item) => item.line);
const totalsCsv = (rows) => `${rows.map((row) => row.join(",")).join("\r\n")}\r\n`;

test("v2.8.0 labels: whole cells only, with a short Chinese prefix, notes and colons removed", () => {
  const { context, evaluate } = createScriptContext();
  const label = (value) => {
    put(context, "__value", value);
    return evaluate("totalRowLabel(__value)");
  };
  const matches = {
    "合计": "合计", "合\u3000计": "合计", "合计：": "合计", "合计:": "合计", "合计（元）": "合计", "合计(元)": "合计", "合计（元）：": "合计", "合计：（元）": "合计",
    "上海 小计": "上海小计", "华东区合计": "华东区合计", "总计": "总计", "总和": "总和", "汇总": "汇总", "汇\u3000总": "汇总", "总 和": "总和", "累计": "累计", "本月累计": "本月累计",
    "TOTAL": "TOTAL", "Grand total:": "Grandtotal", "Total (USD)": "Total", "Sub-total": "Sub-total", "Sub total": "Subtotal", "sum": "sum", "\uFF54\uFF4F\uFF54\uFF41\uFF4C": "total", "\u017Fum": "sum", "T o t a l": "Total",
    [`${"一".repeat(12)}合计`]: `${"一".repeat(12)}合计`
  };
  for (const [value, expected] of Object.entries(matches)) assert.equal(label(value), expected, JSON.stringify(value));
  const others = [`${"一".repeat(13)}合计`, "合计金额", "小计说明", "Total Sales", "SUMMARY", "统计", "共计", "合共", "", "   ", "\u3000",
    "（合计）", "合计（元）（含税）", "合计[元]", "合计【元】", "Sub\u2010total", "合\u200B计"];
  for (const value of others) assert.equal(label(value), "", JSON.stringify(value));
  for (const expression of ["12", 'new Date("2024-01-01")', "null", "undefined", "true"]) {
    assert.equal(evaluate(`totalRowLabel(${expression})`), "", expression);
  }
});

test("v2.8.0 labels: a prefix that is only a status word is no label, a longer one still is", async () => {
  const { context, evaluate } = createScriptContext();
  const label = (value) => {
    put(context, "__value", value);
    return evaluate("totalRowLabel(__value)");
  };
  for (const value of ["已汇总", "未汇总", "不累计", "无合计", "無合计", "非合计", "待汇总", "是否汇总", "尚未汇总", "暂未汇总", "暂不汇总", "暫未匯總", "暫不匯總", "已匯總", "未合計", "已 汇总", "已\u3000匯總"]) {
    assert.equal(label(value), "", value);
  }
  // The whole prefix must be the status word: these only begin or end with one.
  for (const value of ["南非小计", "东非合计", "上海小计", "无锡合计", "未名湖小计", "不丹合计", "已汇总合计"]) {
    assert.equal(label(value), value, value);
  }
  // A status column must neither be flagged nor split the details of the real total.
  const scope = await importCsvFile(totalsCsv([["名", "状态", "数量"], ["A", "已汇总", 2], ["B", "处理中", 1], ["C", "已汇总", 1], ["D", "处理中", 2], ["E", "处理中", 3], ["合计", "", 9]]));
  assert.deepEqual(suspectedTotals(scope).map((item) => [item.line, item.label, item.matches[0].rowCount]), [[7, "合计", 5]]);
});

test("v2.8.0 labels: traditional keywords count, with a prefix and full-width spaces", async () => {
  const { context, evaluate } = createScriptContext();
  const label = (value) => {
    put(context, "__value", value);
    return evaluate("totalRowLabel(__value)");
  };
  for (const value of ["合計", "總計", "小計", "總和", "匯總", "彙總", "累計", "台北小計"]) assert.equal(label(value), value, value);
  assert.equal(label("合\u3000計"), "合計");
  assert.equal(label("匯 總"), "匯總");
  assert.equal(label("合計金額"), "");
  assert.deepEqual(totalLines(await importCsvFile(totalsCsv([["名", "额"], ["甲", 10], ["乙", 20], ["台北小計", 30], ["丙", 5], ["總計", 35]]))), [4, 6]);
});

test("v2.8.0 labels: the prefilter lets through every keyword and every code point that normalizes into a keyword character", () => {
  const { evaluate } = createScriptContext();
  const prefilter = evaluate("TOTAL_LABEL_PREFILTER");
  for (const keyword of ["合计", "总计", "小计", "总和", "汇总", "累计", "合計", "總計", "小計", "總和", "匯總", "彙總", "累計", "total", "TOTAL", "sum", "SUM"]) {
    assert.ok(prefilter.test(keyword), keyword);
  }
  const keywordCharacters = new Set("合计总小和汇累計總匯彙totalsumgrandbTOTALSUMGRANDB");
  const missed = [];
  let checked = 0;
  for (let code = 0x80; code <= 0x10ffff; code += 1) {
    if (code >= 0xd800 && code <= 0xdfff) continue;
    const character = String.fromCodePoint(code);
    const normalized = character.normalize("NFKC");
    if (normalized === character || ![...normalized].some((piece) => keywordCharacters.has(piece))) continue;
    checked += 1;
    if (!prefilter.test(character)) missed.push(code.toString(16));
  }
  assert.ok(checked > 500);
  assert.deepEqual(missed, []);
});

test("v2.8.0 performance: a text file without labels never normalizes a plain cell and no numbers are read", async () => {
  const scope = await importCsvFile(totalsCsv([["名", "额"], ["甲", "1"], ["乙", "2"], ["丙", "3"]]));
  const calls = JSON.parse(scope.evaluate(`JSON.stringify((() => {
    const original = String.prototype.normalize;
    let normalized = 0;
    String.prototype.normalize = function (...args) { normalized += 1; return original.apply(this, args); };
    const rows = Array.from({ length: 200 }, (_, index) => ({ 名: "门店" + index, 额: String(index) }));
    const result = buildSuspectedTotalSummary(rows, ["名", "额"], [{ field: "额", typeKey: "numeric", get numberStrategy() { throw new Error("read the number strategy"); }, confirmedKeys: [] }], () => null);
    String.prototype.normalize = original;
    return { normalized, result };
  })())`));
  assert.deepEqual(calls, { normalized: 0, result: [] });
});

test("v2.8.0 performance: a rebuild reuses the labels found on import, without scanning the cells again", async () => {
  const store = await importStage3Fixture("门店报表.xlsx");
  const labelChecks = JSON.parse(store.evaluate(`JSON.stringify((() => {
    const original = totalRowLabel;
    let calls = 0;
    totalRowLabel = (value) => { calls += 1; return original(value); };
    stageFieldTypeChange("销售额", "category");
    applyFieldConfiguration();
    const rebuild = calls;
    findTotalLabelRows(state.rows, state.fields);
    totalRowLabel = original;
    return { rebuild, scanned: calls > rebuild, kept: state.suspectedTotals.length };
  })())`));
  assert.deepEqual(labelChecks, { rebuild: 0, scanned: true, kept: 1 });
});

test("v2.8.0 fixtures: the store report's row 33 and the region table match known-issues.json", async () => {
  const store = await importStage3Fixture("门店报表.xlsx");
  const region = await importStage3Fixture("分区小计表.csv");
  const pick = (scope) => suspectedTotals(scope).map((item) => ({ line: item.line, label: item.label, field: item.field, matches: item.matches }));
  assert.deepEqual(pick(store), knownIssues["门店报表.xlsx"].suspectedTotals);
  assert.deepEqual(pick(region), knownIssues["分区小计表.csv"].suspectedTotals);
  assert.deepEqual(importedSummary(region.evaluate).types, knownIssues["分区小计表.csv"].expectedTypes.map((item) => item.expected));
  assert.equal(region.evaluate("state.headerDetection"), null);
  assert.equal(region.evaluate("state.totalMissing"), knownIssues["分区小计表.csv"].expectedTotals.totalMissing);
});

test("v2.8.0 fixtures: no other fixture or example has a suspected total or a formula error (v2.11.0: the sample button's file has one total)", async () => {
  const scopes = await importEveryQualityFile();
  assert.deepEqual(scopes.filter(([, scope]) => suspectedTotals(scope).length).map(([fileName]) => fileName), ["en-formula-errors.xlsx", "分区小计表.csv", "门店报表.xlsx", "sample-orders-en.csv", "sample-orders.csv"]);
  assert.deepEqual(scopes.filter(([, scope]) => formulaErrors(scope).length).map(([fileName]) => fileName), ["en-formula-errors.xlsx", "公式错误.xlsx"]);
});

test("v2.8.0 totals: the three ranges, each tried in order", async () => {
  const cases = [
    // No detail above the first row: only the whole table can match.
    [[["名", "额"], ["合计", 30], ["A", 10], ["B", 20]], [[2, "all", 3, 4, 2]]],
    [[["名", "额"], ["A", 10], ["合计", 60], ["B", 20], ["C", 30]], [[3, "all", 2, 5, 3]]],
    // The total below subtotals adds up the details only, not the subtotals.
    [[["名", "额"], ["A", 10], ["小计", 10], ["B", 20], ["小计", 20], ["总计", 30]], [[3, "segment", 2, 2, 1], [5, "segment", 4, 4, 1], [6, "above", 2, 4, 2]]],
    // Two totals in a row: the second has no detail since the first, so it matches everything above.
    [[["名", "额"], ["A", 10], ["B", 20], ["小计", 30], ["合计", 30]], [[4, "segment", 2, 3, 2], [5, "above", 2, 3, 2]]],
    // Returns count as negative; a total can be negative.
    [[["名", "额"], ["A", 100], ["退货", -30], ["合计", 70]], [[4, "segment", 2, 3, 2]]],
    [[["名", "额"], ["A", -10], ["B", -20], ["合计", -30]], [[4, "segment", 2, 3, 2]]]
  ];
  for (const [rows, expected] of cases) {
    const scope = await importCsvFile(totalsCsv(rows));
    const actual = suspectedTotals(scope).map((item) => [item.line, item.matches[0].scope, item.matches[0].firstLine, item.matches[0].lastLine, item.matches[0].rowCount]);
    assert.deepEqual(actual, expected, JSON.stringify(rows));
  }
});

test("v2.8.0 totals: columns that match different ranges are grouped by range", async () => {
  // 额 adds up since the subtotal, 量 only over every detail above it.
  const scope = await importCsvFile(totalsCsv([["名", "额", "量"], ["A", 10, 1], ["B", 20, 2], ["小计", 30, 99], ["C", 40, 4], ["D", 50, 5], ["总计", 90, 12]]));
  const last = suspectedTotals(scope).find((item) => item.line === 7);
  assert.deepEqual(last.matches.map((match) => [match.scope, match.fields, match.firstLine, match.lastLine, match.rowCount, match.contiguous]), [
    ["segment", ["额"], 5, 6, 2, true],
    ["above", ["量"], 2, 6, 4, false]
  ]);
});

test("v2.8.0 totals: counterexamples stay silent, the tolerance is 0.5% of the sum", async () => {
  const silent = [
    [["名", "额"], ["A", 10], ["B", 20], ["合计", 31]],
    [["名", "额"], ["合计", 99], ["A", 10], ["B", 20]],
    // Adds up but says nothing.
    [["名", "额"], ["A", 10], ["B", 20], ["C", 30]],
    // A zero sum never matches, not even a zero cell.
    [["名", "额"], ["A", 5], ["退货", -5], ["合计", 0]],
    [["名", "额"], ["A", 100], ["B", 100], ["合计", 201.2]],
    // Only an empty or unreadable cell on the total row.
    [["名", "额"], ["A", 10], ["B", 20], ["C", 30], ["合计", ""]],
    [["名", "额"], ["A", 10], ["B", 20], ["C", 5], ["D", 5], ["E", 5], ["合计", "见附表"]],
    [["名", "额"], ["A", 10], ["B", 20], ["C", 5], ["D", 5], ["E", 5], ["合计", "N/A"]],
    // No numeric field at all.
    [["名", "说明"], ["A", "甲"], ["合计", "乙"]]
  ];
  for (const rows of silent) assert.deepEqual(totalLines(await importCsvFile(totalsCsv(rows))), [], JSON.stringify(rows));
  assert.deepEqual(totalLines(await importCsvFile(totalsCsv([["名", "额"], ["A", 100], ["B", 100], ["合计", 201]]))), [4]);
});

test("v2.8.0 totals: one column is enough, ratio columns and cells that are not numbers stay out", async () => {
  const scope = await importCsvFile(totalsCsv([["名", "额", "量", "比"], ["A", 10, 1, "0.2"], ["B", 20, "", "0.3"], ["C", "待定", 2, "0.4"], ["D", 5, 1, "0.5"], ["E", 5, 1, "0.1"], ["F", 5, 1, "0.2"], ["合计", 45, 6, "0.28"]]));
  const [item] = suspectedTotals(scope);
  assert.equal(item.line, 8);
  assert.deepEqual(item.matches.map((match) => [match.scope, match.fields, match.firstLine, match.lastLine, match.rowCount, match.contiguous]), [["segment", ["额", "量"], 2, 7, 5, false]]);
});

test("v2.8.0 totals: rows whose label did not add up go back among the details for one more check", async () => {
  // Row 3 only says 合计 in its note: the real total below must not get lost.
  const note = await importCsvFile(totalsCsv([["名", "备注", "额"], ["A", "", 10], ["B", "合计", 20], ["C", "", 30], ["总计", "", 60]]));
  assert.deepEqual(suspectedTotals(note).map((item) => [item.line, item.label, item.matches[0].firstLine, item.matches[0].lastLine, item.matches[0].rowCount]), [[5, "总计", 2, 4, 3]]);
  // A wrong subtotal must not take the rows that passed the first check down with it.
  assert.deepEqual(totalLines(await importCsvFile(totalsCsv([["名", "额"], ["A", 10], ["B", 20], ["小计", 35], ["C", 40], ["D", 50], ["小计", 90], ["总计", 120]]))), [7, 8]);
  // Labels on every row: no plain detail, so equal neighbours are not totals.
  assert.deepEqual(totalLines(await importCsvFile(totalsCsv([["城市", "类型", "额"], ["上海", "小计", 100], ["杭州", "小计", 100], ["宁波", "小计", 300]]))), []);
  // Only one more check: the total below the note and two subtotals stays out (known limit).
  assert.deepEqual(totalLines(await importCsvFile(totalsCsv([["名", "备注", "额"], ["a", "合计", 10], ["b", "", 20], ["上海小计", "", 30], ["c", "", 40], ["d", "", 50], ["杭州小计", "", 90], ["总计", "", 120]]))), [4, 7]);
  // The row checked again is no addend of its own "whole table" range.
  const top = await importCsvFile(totalsCsv([["名", "备注", "额"], ["总计", "", 60], ["A", "", 10], ["B", "合计", 20], ["C", "", 30]]));
  assert.deepEqual(suspectedTotals(top).map((item) => [item.line, item.matches[0].scope, item.matches[0].firstLine, item.matches[0].lastLine, item.matches[0].rowCount]), [[2, "all", 3, 5, 3]]);
});

test("v2.8.0 totals: huge values neither match by overflow nor hide a later segment", async () => {
  assert.deepEqual(totalLines(await importCsvFile(totalsCsv([["名", "额"], ["A", "1e308"], ["B", "1e308"], ["合计", 5]]))), []);
  assert.deepEqual(totalLines(await importCsvFile(totalsCsv([["名", "额"], ["A", "1e308"], ["B", "1e308"], ["小计", "1e308"], ["C", 1], ["D", 2], ["小计", 3]]))), [7]);
  const absorbed = await importCsvFile(totalsCsv([["名", "额"], ["A", "100000000000000000"], ["小计", "100000000000000000"], ["B", 1], ["C", 1], ["小计", 2]]));
  assert.deepEqual(totalLines(absorbed), [3, 6]);
});

test("v2.8.0 totals: numbers read with the field's strategy, labels in any field, the first one in field order", async () => {
  const units = await importCsvFile(totalsCsv([["名", "额"], ["A", "1万"], ["B", "2万"], ["合计", "3万"]]));
  assert.deepEqual(totalLines(units), [4]);
  const first = await importCsvFile(totalsCsv([["名", "备注", "额"], ["A", "", 10], ["B", "", 20], ["华东合计", "小计", 30]]));
  assert.deepEqual(suspectedTotals(first).map((item) => [item.field, item.label]), [["名", "华东合计"]]);
  const store = await importStage3Fixture("门店报表.xlsx");
  store.evaluate('stageFieldTypeChange("门店编号", "ignore"); applyFieldConfiguration()');
  assert.deepEqual(totalLines(store), [33]);
  assert.equal(suspectedTotals(store)[0].field, "门店编号");
});

test("v2.8.0 refresh: field configuration and confirmations rebuild the totals, filters do not", async () => {
  const store = await importStage3Fixture("门店报表.xlsx");
  const status = store.elements.get("statusPanel").textContent;
  put(store.context, "__before", store.evaluate("state.suspectedTotals"));
  applyFilterState(store, { category: { field: "城市", values: ["上海"] } });
  assert.equal(store.evaluate("state.suspectedTotals === __before"), true);
  store.evaluate('stageFieldTypeChange("销售额", "category"); applyFieldConfiguration()');
  assert.deepEqual(suspectedTotals(store)[0].matches[0].fields, ["成本"]);
  store.evaluate('stageFieldTypeChange("成本", "category"); applyFieldConfiguration()');
  assert.deepEqual(suspectedTotals(store), []);
  store.evaluate("restoreAutomaticFieldTypes()");
  assert.deepEqual(totalLines(store), [33]);
  assert.equal(store.elements.get("statusPanel").textContent, status);

  const placeholder = await importCsvFile(totalsCsv([["名", "额"], ["A", 10], ["B", "N/A"], ["C", "N/A"], ["D", 20], ["合计", 30]]));
  assert.equal(placeholder.evaluate('getFieldProfile("额").typeKey'), "category");
  assert.deepEqual(totalLines(placeholder), []);
  placeholder.evaluate('confirmPlaceholder("额", "N/A")');
  assert.deepEqual(totalLines(placeholder), [6]);
  placeholder.evaluate('undoPlaceholderConfirmation("额", "N/A")');
  assert.deepEqual(totalLines(placeholder), []);
});

test("v2.8.0 state: a later commit and resetAnalysisState() clear the labels and the totals", async () => {
  const store = await importStage3Fixture("门店报表.xlsx");
  assert.equal(store.evaluate("state.totalLabelRows.length"), 1);
  put(store.context, "__plainRows", [{ 城市: "上海" }]);
  store.evaluate('commitTabularData(["城市"], __plainRows, "later.csv", state.activeImportId)');
  assert.equal(store.evaluate("state.totalLabelRows.length"), 0);
  assert.equal(store.evaluate("state.suspectedTotals.length"), 0);
  const again = await importStage3Fixture("门店报表.xlsx");
  again.evaluate("resetAnalysisState()");
  assert.equal(again.evaluate("state.totalLabelRows.length"), 0);
  assert.equal(again.evaluate("state.suspectedTotals.length"), 0);
});

test("v2.8.0 totals: the check changes nothing, the rows and every count stay as they were", async () => {
  const store = await importStage3Fixture("门店报表.xlsx");
  // Row 33 still counts as data: known-issues.json S5.
  assert.equal(store.evaluate("state.rows.length"), 29);
  assert.equal(store.evaluate("state.totalMissing"), 4);
  assert.equal(store.evaluate('getQualityNumericStats().find((stat) => stat.field === "销售额").count'), 29);
  assert.equal(store.evaluate('getQualityNumericStats().find((stat) => stat.field === "销售额").outlierCount'), 1);
  assert.equal(store.evaluate('state.rows[28]["门店编号"]'), "合计");
  assert.equal(store.evaluate('Object.keys(state.rows[28]).join("|")'), "门店编号|门店名称|城市|开业日期|销售额|成本|毛利率|店长");
});

const sectionItems = (scope, id) => [...scope.elements.get(id).innerHTML.matchAll(/<li>(.*?)<\/li>/gs)].map((match) => match[1]);
const sectionButtons = (scope, id) => [...scope.elements.get(id).innerHTML.matchAll(/data-problem-index="(\d+)"/g)].map((match) => Number(match[1]));
const problemCategories = (scope) => JSON.parse(scope.evaluate("JSON.stringify(state.problemCategories)"));
// Two rows per group: one detail and its subtotal, the subtotals at lines 3, 5, 7, ...
const subtotalsCsv = (count) => `名,额\r\n${Array.from({ length: count }, (_, index) => `A${index},10\r\nA${index}小计,10`).join("\r\n")}\r\n`;

test("v2.8.0 page: both subsections are static, after the warnings and before the tables, with their notes", () => {
  const html = indexHtmlSource();
  const positions = ['id="parseWarningSummary"', '<div id="suspectedTotalSection" class="module-subsection row-note-section hidden">', '<div id="formulaErrorSection" class="module-subsection row-note-section hidden">', '<div class="two-column">'].map((text) => html.indexOf(text));
  assert.equal(positions.every((position) => position >= 0), true);
  assert.deepEqual(positions, [...positions].sort((a, b) => a - b));
  assert.match(html, /<p class="section-note">这些行目前仍按普通数据参与统计，其中的数字会被重复计入求和、平均值和偏大或偏小的值的判断。<\/p>/);
  assert.match(html, /<p class="section-note">这些单元格的公式算出了错误，按空值处理（已计入空值）。整行都是错误值的行不会导入，也不在这里列出。<\/p>/);
});

test("v2.8.0 status: the sentences follow the header note, on the warning branch too, one row or many", async () => {
  const region = await importStage3Fixture("分区小计表.csv");
  assert.equal(region.elements.get("statusPanel").textContent, "分区小计表.csv 已成功载入，分析结果已生成。发现 5 行疑似合计 / 小计（第 5、8、10、14、15 行），可以在“数据质量”里设为不计入分析。");
  const errors = await importStage3Fixture("公式错误.xlsx");
  assert.equal(errors.elements.get("statusPanel").textContent, "公式错误.xlsx · 公式错误 已成功载入，分析结果已生成。发现 7 个公式错误单元格，已按空值处理，详情见“数据质量”。");
  const warned = await importCsvFile("名,额\r\nA,10\r\nB\r\nC,20\r\n合计,30\r\n");
  assert.equal(warned.elements.get("statusPanel").textContent, "CSV 已载入，但读取时有 1 处疑问，涉及 1 行。请核对数据预览和字段数量。第 5 行「合计」疑似合计行，可以在“数据质量”里设为不计入分析。");
  const many = await importCsvFile(subtotalsCsv(7));
  assert.match(many.elements.get("statusPanel").textContent, /发现 7 行疑似合计 \/ 小计（第 3、5、7、9、11 行等），可以在“数据质量”里设为不计入分析/);
  const five = await importCsvFile(subtotalsCsv(5));
  assert.match(five.elements.get("statusPanel").textContent, /发现 5 行疑似合计 \/ 小计（第 3、5、7、9、11 行），可以在“数据质量”里设为不计入分析/);
  const { context, evaluate, elements } = createScriptContext();
  put(context, "__rows", [{ 名: "A", 额: "10" }, { 名: "合计", 额: "10" }]);
  evaluate('commitTabularData(["名", "额"], __rows, "x.csv", beginImport())');
  assert.match(elements.get("statusPanel").textContent, /有一行「合计」疑似合计行/);
  put(context, "__manyRows", [{ 名: "A", 额: "10" }, { 名: "A小计", 额: "10" }, { 名: "B", 额: "20" }, { 名: "B小计", 额: "20" }]);
  evaluate('commitTabularData(["名", "额"], __manyRows, "x.csv", beginImport())');
  assert.match(elements.get("statusPanel").textContent, /发现 2 行疑似合计 \/ 小计，可以在“数据质量”里设为不计入分析/);
  const clean = await importCsvFile("名,额\r\nA,10\r\nB,20\r\n");
  assert.doesNotMatch(clean.elements.get("statusPanel").textContent, /疑似合计|公式错误/);
});

test("v2.8.0 page: the store report has one line and one button, the row opens with its label cell highlighted", async () => {
  const store = await importStage3Fixture("门店报表.xlsx");
  assert.equal(store.elements.get("suspectedTotalSection").classList.contains("hidden"), false);
  const items = sectionItems(store, "suspectedTotalDetails");
  assert.deepEqual(items, ["第 33 行「合计」：销售额、成本等于第 5–32 行之和（28 行）。"]);
  const buttons = sectionButtons(store, "suspectedTotalDetails");
  assert.equal(buttons.length, 1);
  assert.deepEqual(problemCategories(store)[buttons[0]], { kind: "suspectedTotal", count: 1, unit: "row" });
  clickProblemButton(store, buttons[0]);
  assert.equal(store.elements.get("problemRowsSummary").textContent, "共 1 行，第 1 / 1 页");
  const table = store.elements.get("problemRowsTable").innerHTML;
  assert.match(table, /<tr><td>33<\/td><td>销售额、成本等于第 5–32 行之和（28 行）<\/td><td class="cell-problem">合计<\/td>/);
  assert.equal([...table.matchAll(/cell-problem/g)].length, 1);
  assert.equal(store.elements.get("formulaErrorSection").classList.contains("hidden"), true);
});

test("v2.8.0 page: five subtotals, one button; the note names only the columns that matched", async () => {
  const region = await importStage3Fixture("分区小计表.csv");
  const items = sectionItems(region, "suspectedTotalDetails");
  assert.equal(items.length, 5);
  assert.match(items[0], /^第 5 行「上海 小计」：销售额、成本、占比等于第 2–4 行之和（3 行）。$/);
  assert.match(items[4], /^第 15 行「/);
  assert.doesNotMatch(items[4], /占比/);
  const buttons = sectionButtons(region, "suspectedTotalDetails");
  assert.equal(buttons.length, 1);
  assert.doesNotMatch(region.elements.get("suspectedTotalDetails").innerHTML, /data-placeholder-index/);
  assert.match(region.elements.get("suspectedTotalDetails").innerHTML, /aria-label="查看疑似合计 \/ 小计行"/);
  assert.doesNotMatch(region.elements.get("suspectedTotalDetails").innerHTML, /只列出/);
  assert.deepEqual(problemCategories(region).find((item) => item.kind === "suspectedTotal"), { kind: "suspectedTotal", count: 5, unit: "row" });
  assert.equal(region.evaluate('formatProblemCategoryLabel({ kind: "suspectedTotal", count: 5 })'), "疑似合计 / 小计行（5 行）");
  clickProblemButton(region, buttons[0]);
  assert.equal(region.elements.get("problemRowsSummary").textContent, "共 5 行，第 1 / 1 页");
  assert.match(region.elements.get("problemRowsTable").innerHTML, /<td>5<\/td><td>销售额、成本、占比等于第 2–4 行之和（3 行）<\/td><td class="cell-problem">上海 小计<\/td>/);
  const csv = region.evaluate("buildProblemRowsCsvContent(state.problemView.list, state.fields)");
  assert.match(csv, /^文件行号,说明,门店,城市,销售额,成本,占比\r\n5,销售额、成本、占比等于第 2–4 行之和（3 行）,上海 小计,上海,369\.4,255\.9,46%/);
  assert.equal(region.evaluate('buildProblemRowsFileName({ kind: "suspectedTotal" }, new Date(2026, 10, 2))'), "分区小计表_问题行_疑似合计_20261102.csv");
  // No new pill: exactly the six of the pills assertion.
  assert.equal([...region.elements.get("qualitySummary").innerHTML.matchAll(/<span class="quality-pill/g)].length, 6);
  assert.deepEqual(problemRowCheck(region).mismatches, []);
});

test("v2.8.0 page: suspected totals list at most 20 lines, the viewer has them all", async () => {
  const exact = await importCsvFile(subtotalsCsv(20));
  assert.equal(sectionItems(exact, "suspectedTotalDetails").length, 20);
  assert.doesNotMatch(exact.elements.get("suspectedTotalDetails").innerHTML, /只列出/);

  const over = await importCsvFile(subtotalsCsv(21));
  const items = sectionItems(over, "suspectedTotalDetails");
  assert.equal(items.length, 20);
  assert.match(items[19], /^第 41 行「A19小计」/);
  assert.match(over.elements.get("suspectedTotalDetails").innerHTML, /共 21 行，只列出前 20 行。全部行请点「查看」。/);
  const buttons = sectionButtons(over, "suspectedTotalDetails");
  assert.equal(buttons.length, 1);
  clickProblemButton(over, buttons[0]);
  assert.equal(over.elements.get("problemRowsSummary").textContent, "共 21 行，第 1 / 1 页");
  assert.equal(over.evaluate("state.problemView.list.items.length"), 21);
  assert.equal(over.evaluate("SUSPECTED_TOTAL_LIST_LIMIT"), 20);

  const large = await importCsvFile(subtotalsCsv(5000));
  assert.equal(sectionItems(large, "suspectedTotalDetails").length, 20);
  assert.match(large.elements.get("suspectedTotalDetails").innerHTML, /共 5,000 行，只列出前 20 行。/);
});

test("v2.8.0 page: formula errors have one button per field and every note, the empty cell is highlighted", async () => {
  const scope = await importStage3Fixture("公式错误.xlsx");
  assert.equal(scope.elements.get("formulaErrorSection").classList.contains("hidden"), false);
  const items = sectionItems(scope, "formulaErrorDetails");
  assert.equal(items.length, 4);
  assert.match(items.join(), /「利润率」：3 个（#DIV\/0! 2 个、#REF! 1 个），第 3、5、9 行。/);
  const buttons = sectionButtons(scope, "formulaErrorDetails");
  assert.equal(buttons.length, 4);
  const categories = problemCategories(scope);
  assert.deepEqual(buttons.map((index) => [categories[index].kind, categories[index].field]), knownIssues["公式错误.xlsx"].formulaErrors.map((item) => ["formulaError", item.field]));
  assert.deepEqual(buttons.map((index) => categories[index].count), knownIssues["公式错误.xlsx"].formulaErrors.map((item) => item.count));
  assert.equal(scope.evaluate('formatProblemCategoryLabel({ kind: "formulaError", field: "备注", count: 2 })'), "公式错误 · 备注（2 个值）");
  assert.equal(scope.evaluate('formatProblemViewLabel({ kind: "formulaError", field: "备注" })'), "查看「备注」的公式错误");

  const notes = JSON.parse(scope.evaluate('JSON.stringify(listProblemRows({ kind: "missing", field: "备注" }).items.map((item) => [item.line, item.note]))'));
  assert.deepEqual(notes, [[5, "公式错误 #NULL!"], [6, "公式错误 #SPILL!"], [8, "空值"]]);
  assert.deepEqual(JSON.parse(scope.evaluate('JSON.stringify(listProblemRows({ kind: "formulaError", field: "利润率" }).items.map((item) => [item.line, item.note]))')), [[3, "公式错误 #DIV/0!"], [5, "公式错误 #REF!"], [9, "公式错误 #DIV/0!"]]);
  assert.deepEqual(problemRowCheck(scope).mismatches, []);
  clickProblemButton(scope, categories.findIndex((item) => item.kind === "formulaError" && item.field === "备注"));
  assert.equal(scope.elements.get("problemRowsSummary").textContent, "共 2 个值，第 1 / 1 页");
  assert.match(scope.elements.get("problemRowsTable").innerHTML, /<tr><td>5<\/td><td>公式错误 #NULL!<\/td><td>P04<\/td><td>8<\/td><td>6<\/td><td>48<\/td><td><\/td><td class="cell-problem"><\/td><\/tr>/);
  assert.equal(scope.evaluate('buildProblemRowsFileName({ kind: "formulaError", field: "利润率" }, new Date(2026, 10, 2))'), "公式错误_问题行_公式错误_利润率_20261102.csv");

  const before = scope.evaluate("JSON.stringify(state.formulaErrors)");
  scope.evaluate('stageFieldTypeChange("利润率", "ignore"); applyFieldConfiguration()');
  assert.equal(scope.evaluate("JSON.stringify(state.formulaErrors)"), before);
  assert.match(scope.elements.get("formulaErrorDetails").innerHTML, /「利润率」（字段已忽略）：3 个（#DIV\/0! 2 个、#REF! 1 个），第 3、5、9 行。/);
  scope.evaluate("restoreAutomaticFieldTypes()");
  assert.doesNotMatch(scope.elements.get("formulaErrorDetails").innerHTML, /字段已忽略/);
});

test("v2.8.0 page: the sentences for row ranges, many columns and missing row numbers", () => {
  const { evaluate } = createScriptContext();
  const line = (item) => evaluate(`formatSuspectedTotal(${JSON.stringify(item)})`);
  const match = { scope: "segment", fields: ["额"], firstLine: 2, lastLine: 13, rowCount: 12, contiguous: true };
  assert.equal(line({ line: 14, label: "合计", matches: [match] }), "第 14 行「合计」：额等于第 2–13 行之和（12 行）。");
  assert.equal(line({ line: 14, label: "合计", matches: [{ ...match, rowCount: 9, contiguous: false }] }), "第 14 行「合计」：额等于第 2–13 行中的明细之和（9 行）。");
  assert.equal(line({ line: 3, label: "小计", matches: [{ ...match, firstLine: 2, lastLine: 2, rowCount: 1 }] }), "第 3 行「小计」：额等于第 2 行（只有 1 行）。");
  assert.equal(line({ line: null, label: "合计", matches: [{ ...match, firstLine: null, lastLine: null, rowCount: 3 }] }), "有一行「合计」：额等于上方 3 行明细之和。");
  assert.equal(line({ line: null, label: "合计", matches: [{ ...match, scope: "all", firstLine: null, lastLine: null, rowCount: 3 }] }), "有一行「合计」：额等于全表 3 行明细之和。");
  assert.equal(line({ line: 14, label: "合计", matches: [{ ...match, rowCount: 1000, firstLine: 2, lastLine: 1001 }] }), "第 14 行「合计」：额等于第 2–1001 行之和（1,000 行）。");
  assert.equal(line({ line: 14, label: "合计", matches: [{ ...match, fields: ["a", "b", "c", "d", "e"] }, { ...match, scope: "all", fields: ["a", "b", "c", "d", "e", "f"] }] }), "第 14 行「合计」：a、b、c、d、e等于第 2–13 行之和（12 行）；a、b、c、d、e 等 6 列等于第 2–13 行之和（12 行）。");

  const error = (item) => evaluate(`formatFormulaError(${JSON.stringify(item)})`);
  const base = { field: "额", count: 2, byError: { "#N/A": 2 }, lines: [3, 4] };
  assert.equal(error(base), "「额」：2 个（#N/A 2 个），第 3、4 行。");
  assert.equal(error({ ...base, lines: [null, null] }), "「额」：2 个（#N/A 2 个），行号：—。");
  assert.equal(error({ ...base, count: 12, byError: { "#N/A": 12 } }), "「额」：12 个（#N/A 12 个），第 3、4 行等。");
  assert.equal(evaluate('formatFormulaErrorStatus([{ count: 1000 }, { count: 500 }])'), "发现 1,500 个公式错误单元格，已按空值处理，详情见“数据质量”。");
  assert.equal(evaluate("formatFormulaErrorStatus([])"), "");
  assert.equal(evaluate("formatSuspectedTotalStatus([])"), "");
});

test("v2.8.0 problem rows: nine kinds in order, and a confirmed placeholder keeps its own note", () => {
  const { context, evaluate } = createScriptContext();
  const amounts = [10, 11, 12, 10, 11, 12, 10, 11, 12, 1000];
  const rows = amounts.map((amount, index) => ({ 订单ID: `P${index}`, 名: `A${index}`, 额: String(amount), 备注: "好" }));
  rows.push({ ...rows[0] });
  rows.push({ 订单ID: "P20", 名: "B", 额: "abc", 备注: "N/A" });
  rows.push({ 订单ID: "P21", 名: "C", 额: "12", 备注: "" });
  rows.push({ 订单ID: "P22", 名: "D", 额: "13", 备注: "" });
  const sum = rows.reduce((total, row) => total + (Number(row.额) || 0), 0);
  rows.push({ 订单ID: "P23", 名: "合计", 额: String(sum), 备注: "好" });
  put(context, "__rows", rows);
  put(context, "__errors", rows.map((row) => (row.订单ID === "P22" ? { 备注: "#REF!" } : null)));
  put(context, "__warnings", [{ row: 2, code: "TooFewFields", message: "少一列" }]);
  evaluate('commitTabularData(["订单ID", "名", "额", "备注"], __rows, "x.xlsx", beginImport(), __warnings, "", null, null, null, __errors)');
  const kinds = JSON.parse(evaluate("JSON.stringify(state.problemCategories.map((item) => item.kind))"));
  assert.deepEqual(kinds.filter((kind, index) => kind !== kinds[index - 1]), ["missing", "formulaError", "conversion", "outlier", "duplicateRows", "suspectedTotal", "idDuplicate", "placeholder", "parseWarnings"]);
  evaluate('confirmPlaceholder("备注", "N/A")');
  assert.deepEqual(JSON.parse(evaluate('JSON.stringify(listProblemRows({ kind: "missing", field: "备注" }).items.map((item) => item.note))')), ["已按空值处理", "空值", "公式错误 #REF!"]);
});

test("v2.8.0 page: labels and field names are escaped", async () => {
  const scope = await importWorkbookFile(buildErrorWorkbook([["名", "<i>额"], ["A", 1], ["B", { e: 7 }], ["<b>合计", 1]]));
  assert.match(scope.elements.get("suspectedTotalDetails").innerHTML, /「&lt;b&gt;合计」：&lt;i&gt;额等于/);
  assert.doesNotMatch(scope.elements.get("suspectedTotalDetails").innerHTML, /<b>|<i>/);
  assert.match(scope.elements.get("formulaErrorDetails").innerHTML, /「&lt;i&gt;额」：1 个/);
  assert.doesNotMatch(scope.elements.get("formulaErrorDetails").innerHTML, /<i>/);
});

test("v2.8.0 page: formula errors show without numeric fields, the totals section hides when the last numeric field goes", async () => {
  const text = await importWorkbookFile(buildErrorWorkbook([["名", "说明"], ["A", "好"], ["B", { e: 42 }], ["C", "好"]]));
  assert.equal(text.evaluate("state.qualityNumericStats.length"), 0);
  assert.equal(text.elements.get("formulaErrorSection").classList.contains("hidden"), false);
  assert.match(text.elements.get("formulaErrorDetails").innerHTML, /「说明」：1 个（#N\/A 1 个），第 3 行。/);

  const totals = await importCsvFile(totalsCsv([["名", "额"], ["A", 10], ["B", 20], ["合计", 30]]));
  const buttons = sectionButtons(totals, "suspectedTotalDetails");
  assert.equal(buttons.length, 1);
  clickProblemButton(totals, buttons[0]);
  assert.equal(totals.evaluate("state.problemView.category.kind"), "suspectedTotal");
  totals.evaluate('stageFieldTypeChange("额", "category"); applyFieldConfiguration()');
  assert.equal(totals.evaluate("state.qualityNumericStats.length"), 0);
  assert.equal(totals.elements.get("suspectedTotalSection").classList.contains("hidden"), true);
  assert.equal(totals.elements.get("suspectedTotalDetails").innerHTML, "");
  assert.equal(totals.evaluate('listProblemRows({ kind: "suspectedTotal" })'), null);
  assert.equal(totals.evaluate("state.problemView"), null);
  assert.equal(totals.elements.get("problemRowsMessage").textContent, "这一类已经没有问题行。");
  totals.evaluate("restoreAutomaticFieldTypes()");
  assert.equal(totals.elements.get("suspectedTotalSection").classList.contains("hidden"), false);
  assert.equal(sectionItems(totals, "suspectedTotalDetails").length, 1);
});

test("v2.8.0 page: the sections follow field configuration and confirmations, filters and the status bar stay", async () => {
  const store = await importStage3Fixture("门店报表.xlsx");
  const status = store.elements.get("statusPanel").textContent;
  const before = store.elements.get("suspectedTotalDetails").innerHTML;
  applyFilterState(store, { category: { field: "城市", values: ["上海"] } });
  assert.equal(store.elements.get("suspectedTotalDetails").innerHTML, before);
  store.evaluate('stageFieldTypeChange("销售额", "category"); applyFieldConfiguration()');
  assert.deepEqual(sectionItems(store, "suspectedTotalDetails"), ["第 33 行「合计」：成本等于第 5–32 行之和（28 行）。"]);
  assert.equal(store.elements.get("statusPanel").textContent, status);

  const placeholder = await importCsvFile(totalsCsv([["名", "额"], ["A", 10], ["B", "N/A"], ["C", "N/A"], ["D", 20], ["合计", 30]]));
  assert.equal(placeholder.elements.get("suspectedTotalSection").classList.contains("hidden"), true);
  placeholder.evaluate('confirmPlaceholder("额", "N/A")');
  assert.equal(placeholder.elements.get("suspectedTotalSection").classList.contains("hidden"), false);
  placeholder.evaluate('undoPlaceholderConfirmation("额", "N/A")');
  assert.equal(placeholder.elements.get("suspectedTotalSection").classList.contains("hidden"), true);
});

test("v2.8.0 state: a later commit and resetAnalysisState() empty both sections", async () => {
  const scope = await importStage3Fixture("公式错误.xlsx");
  put(scope.context, "__cityRows", [{ 城市: "上海" }]);
  // The same import id, so resetAnalysisState() in beginImport() cannot hide a commit that keeps them.
  scope.evaluate('commitTabularData(["城市"], __cityRows, "later.csv", state.activeImportId)');
  assert.equal(scope.elements.get("formulaErrorSection").classList.contains("hidden"), true);
  assert.equal(scope.elements.get("formulaErrorDetails").innerHTML, "");
  const again = await importStage3Fixture("公式错误.xlsx");
  again.evaluate("resetAnalysisState()");
  assert.equal(again.elements.get("formulaErrorSection").classList.contains("hidden"), true);
  assert.equal(again.elements.get("formulaErrorDetails").innerHTML, "");
  const store = await importStage3Fixture("门店报表.xlsx");
  store.evaluate("resetAnalysisState()");
  assert.equal(store.elements.get("suspectedTotalSection").classList.contains("hidden"), true);
  assert.equal(store.elements.get("suspectedTotalDetails").innerHTML, "");
});

test("the import status keeps has-note when it carries a note: header rows, totals, formula errors", async () => {
  const statusClass = (scope) => scope.elements.get("statusPanel").className;
  const store = await importStage3Fixture("门店报表.xlsx");
  assert.equal(statusClass(store), "status-panel success has-note");
  for (const fileName of ["成绩表.csv", "成绩表-标题无逗号.csv"]) {
    const grades = await importStage3Fixture(fileName);
    assert.match(grades.elements.get("statusPanel").textContent, /已跳过上方 1 行/);
    assert.equal(statusClass(grades), "status-panel success has-note");
  }
  assert.equal(statusClass(await importStage3Fixture("分区小计表.csv")), "status-panel success has-note");
  assert.equal(statusClass(await importStage3Fixture("公式错误.xlsx")), "status-panel success has-note");
  const warned = await importCsvFile("名,额\r\nA,10\r\nB\r\nC,20\r\n合计,30\r\n");
  assert.equal(statusClass(warned), "status-panel warning has-note");

  const sample = await importCsvFile(readExampleCsv("sample-data.csv"), "sample-data.csv");
  assert.equal(sample.elements.get("statusPanel").textContent, "sample-data.csv 已成功载入，分析结果已生成。");
  assert.equal(statusClass(sample), "status-panel success");
  const warnedOnly = await importCsvFile("名,额\r\nA,10\r\nB\r\nC,20\r\n");
  assert.equal(statusClass(warnedOnly), "status-panel warning");

  // Any later message replaces the import's: a plain import and a report message drop the class.
  put(store.context, "__plainRows", [{ 城市: "上海" }]);
  store.evaluate('commitTabularData(["城市"], __plainRows, "later.csv", state.activeImportId)');
  assert.equal(statusClass(store), "status-panel success");
  const again = await importStage3Fixture("门店报表.xlsx");
  again.evaluate('setStatus("HTML 报告已生成：x.html", "success")');
  assert.equal(statusClass(again), "status-panel success");
});

test("only a successful status without a note is hidden once results show", () => {
  const css = fs.readFileSync(path.join(__dirname, "..", "style.css"), "utf8");
  assert.match(css, /body\.has-results #statusPanel\.success:not\(\.has-note\)\s*\{\s*display:\s*none;\s*\}/);
  assert.doesNotMatch(css, /body\.has-results #statusPanel\.success\s*\{/);
});

test("v2.8.0 report: sections and summary counts only when there is something, old snapshots without them", async () => {
  const store = await importStage3Fixture("门店报表.xlsx");
  const { report, html } = reportOf(store);
  assert.equal(report.quality.suspectedTotals.length, 1);
  assert.ok(html.includes("；疑似合计 / 小计行：<strong>1</strong> 行。"));
  assert.ok(html.includes("<h3>疑似合计 / 小计行</h3>"));
  assert.ok(html.includes("<li>第 33 行「合计」：销售额、成本等于第 5–32 行之和（28 行）。</li>"));
  assertInOrder(html, ["<h3>疑似合计 / 小计行</h3>", "<h3>各字段空值情况</h3>"]);
  assert.doesNotMatch(html, /公式错误/);

  const errors = await importStage3Fixture("公式错误.xlsx");
  const errorReport = reportOf(errors);
  assert.ok(errorReport.html.includes("；公式错误：<strong>7</strong> 个单元格。"));
  assert.ok(errorReport.html.includes("<h3>公式错误单元格</h3>"));
  assert.ok(errorReport.html.includes("<li>「利润率」：3 个（#DIV/0! 2 个、#REF! 1 个），第 3、5、9 行。</li>"));
  assert.doesNotMatch(errorReport.html, /疑似合计/);
  errors.evaluate('stageFieldTypeChange("利润率", "ignore"); applyFieldConfiguration()');
  assert.match(reportOf(errors).html, /「利润率」（字段已忽略）：3 个/);

  const clean = reportOf(await importStage3Fixture("订单表.csv"));
  assert.deepEqual([clean.report.quality.suspectedTotals, clean.report.quality.formulaErrors], [[], []]);
  assert.doesNotMatch(clean.html, /疑似合计|公式错误/);

  delete report.quality.suspectedTotals;
  delete report.quality.formulaErrors;
  put(store.context, "__old", report);
  assert.equal(store.evaluate("JSON.stringify([normalizeReportSnapshot(__old).quality.suspectedTotals, normalizeReportSnapshot(__old).quality.formulaErrors])"), "[null,null]");
  assert.doesNotMatch(store.evaluate("buildHtmlReport(__old, [])"), /疑似合计|公式错误/);
});

test("v2.8.0 report: the note matches the page, at most 20 totals, the counts and overflow say how many there are", async () => {
  const pageHtml = indexHtmlSource();
  const exact = await importCsvFile(subtotalsCsv(20));
  const note = exact.evaluate("SUSPECTED_TOTAL_NOTE[0]");
  assert.equal(note, "这些行目前仍按普通数据参与统计，其中的数字会被重复计入求和、平均值和偏大或偏小的值的判断。");
  assert.ok(pageHtml.includes(`<p class="section-note">${note}</p>`));
  const atLimit = reportOf(exact);
  assert.equal(atLimit.report.quality.suspectedTotals.length, 20);
  assert.ok(atLimit.html.includes(`<p>${note}</p>`));
  assert.doesNotMatch(atLimit.html, /只列出/);

  const over = reportOf(await importCsvFile(subtotalsCsv(21)));
  assert.equal(over.report.quality.suspectedTotals.length, 21);
  assert.ok(over.html.includes("；疑似合计 / 小计行：<strong>21</strong> 行"));
  for (const output of [over.html]) {
    assert.ok(output.includes("第 41 行「A19小计」"));
    assert.ok(!output.includes("第 43 行「A20小计」"));
    assert.ok(output.includes("共 21 行，只列出前 20 行。"));
    assert.doesNotMatch(output, /全部行请点/);
  }
  assert.equal([...over.html.matchAll(/<li>第 \d+ 行「A\d+小计」/g)].length, 20);
});

test("v2.8.0 report: the sections ignore filters, and the snapshot shares nothing with state", async () => {
  const store = await importStage3Fixture("门店报表.xlsx");
  applyFilterState(store, { category: { field: "城市", values: ["上海"] } });
  assert.ok(store.evaluate("getAnalysisRows().length < state.rows.length"));
  assert.ok(reportOf(store).html.includes("<li>第 33 行「合计」：销售额、成本等于第 5–32 行之和（28 行）。</li>"));
  store.evaluate('__snapshot = buildReportData(new Date("2026-10-25T00:00:00.000Z"))');
  assert.equal(store.evaluate("__snapshot.quality.suspectedTotals[0] === state.suspectedTotals[0]"), false);
  assert.equal(store.evaluate("__snapshot.quality.suspectedTotals[0].matches === state.suspectedTotals[0].matches"), false);
  assert.equal(store.evaluate("__snapshot.quality.suspectedTotals[0].matches[0] === state.suspectedTotals[0].matches[0]"), false);
  assert.equal(store.evaluate("__snapshot.quality.suspectedTotals[0].matches[0].fields === state.suspectedTotals[0].matches[0].fields"), false);
  const errors = await importStage3Fixture("公式错误.xlsx");
  errors.evaluate('__snapshot = buildReportData(new Date("2026-10-25T00:00:00.000Z"))');
  assert.equal(errors.evaluate("__snapshot.quality.formulaErrors[0] === state.formulaErrors[0]"), false);
  assert.equal(errors.evaluate("__snapshot.quality.formulaErrors[0].byError === state.formulaErrors[0].byError"), false);
  assert.equal(errors.evaluate("__snapshot.quality.formulaErrors[0].lines === state.formulaErrors[0].lines"), false);
});

test("v2.8.0 report: labels and fields are escaped for each output", async () => {
  const scope = await importCsvFile('名,<i>额|x</i>\r\nA,10\r\n"<b>y|\n小计",10\r\n');
  const { html } = reportOf(scope);
  assert.ok(html.includes("<li>第 3 行「&lt;b&gt;y|\n小计」：&lt;i&gt;额|x&lt;/i&gt;等于第 2 行（只有 1 行）。</li>"));
  assert.ok(!html.includes("<b>y"));
  const errors = reportOf(await importWorkbookFile(buildErrorWorkbook([["名", "<i>额"], ["A", 1], ["B", { e: 7 }], ["C", 2]])));
  assert.ok(errors.html.includes("<li>「&lt;i&gt;额」：1 个（#DIV/0! 1 个），第 3 行。</li>"));
});

test("v2.8.0 report: the formula error note matches the page and comes before the list, only with errors", async () => {
  const pageHtml = indexHtmlSource();
  const errors = await importStage3Fixture("公式错误.xlsx");
  const note = errors.evaluate("FORMULA_ERROR_NOTE[0]");
  assert.equal(note, "这些单元格的公式算出了错误，按空值处理（已计入空值）。整行都是错误值的行不会导入，也不在这里列出。");
  assert.ok(pageHtml.includes(`<p class="section-note">${note}</p>`));
  const { html } = reportOf(errors);
  assert.ok(html.includes(`<h3>公式错误单元格</h3>\n    <p>${note}</p>\n    <ul class="summary-list"><li>「数量」`));
  const totalsOnly = reportOf(await importStage3Fixture("门店报表.xlsx"));
  assert.ok(!totalsOnly.html.includes(note));
  const clean = reportOf(await importStage3Fixture("订单表.csv"));
  assert.ok(!clean.html.includes(note));
});

// v2.8.1: SheetJS applies a date number format to text and boolean cells under
// sheet_to_json(raw: true), so the importer gives them their own value back. These sheets
// are built cell by cell, with shared strings as Excel stores them.
const DATE_FORMAT = "yyyy-mm-dd";
// Excel serial numbers: 45418 is 2024-05-06.
const dateCell = (serial) => ({ t: "n", v: serial, z: DATE_FORMAT });
const textCell = (text, format) => (format ? { t: "s", v: text, z: format } : { t: "s", v: text });

function buildCellSheetBuffer(cells, ref, bookType = "xlsx") {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, { "!ref": ref, ...JSON.parse(JSON.stringify(cells)) }, "S");
  return XLSX.write(workbook, { type: "buffer", bookType, bookSST: true });
}

async function importCellSheet(cells, ref, bookType = "xlsx") {
  const scope = createExcelContext();
  const bytes = buildCellSheetBuffer(cells, ref, bookType);
  put(scope.context, "__cellFile", {
    name: bookType === "xlsx" ? "cells.xlsx" : "cells.xls",
    size: bytes.length,
    arrayBuffer: async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.length)
  });
  await scope.evaluate("parseDataFile(__cellFile)");
  return scope;
}

const cellSheetFields = (scope) => JSON.parse(scope.evaluate("JSON.stringify(state.fields)"));
const cellSheetColumn = (scope, field) => JSON.parse(scope.evaluate(`JSON.stringify(state.rows.map((row) => row[${JSON.stringify(field)}]))`));
const cellSheetProfile = (scope, field) => JSON.parse(scope.evaluate(`JSON.stringify(getFieldProfile(${JSON.stringify(field)}))`));

// A whole column under a date format, its header too: six dates and two texts.
const WHOLE_DATE_COLUMN = {
  A1: textCell("姓名"), B1: textCell("入职日期", DATE_FORMAT),
  A2: textCell("甲"), B2: dateCell(45418),
  A3: textCell("乙"), B3: textCell("待定", DATE_FORMAT),
  A4: textCell("丙"), B4: textCell("2021/3/5", DATE_FORMAT),
  A5: textCell("丁"), B5: dateCell(45419),
  A6: textCell("戊"), B6: dateCell(45420),
  A7: textCell("己"), B7: dateCell(45421),
  A8: textCell("庚"), B8: dateCell(45422),
  A9: textCell("辛"), B9: dateCell(45423)
};

for (const bookType of ["xlsx", "biff8"]) {
  test(`v2.8.1 a whole column under a date format imports with its header and texts (${bookType})`, async () => {
    const scope = await importCellSheet(WHOLE_DATE_COLUMN, "A1:B9", bookType);
    assert.deepEqual(cellSheetFields(scope), ["姓名", "入职日期"]);
    assert.deepEqual(cellSheetColumn(scope, "入职日期"), ["2024-05-06", "待定", "2021/3/5", "2024-05-07", "2024-05-08", "2024-05-09", "2024-05-10", "2024-05-11"]);
    const profile = cellSheetProfile(scope, "入职日期");
    assert.equal(profile.typeKey, "date");
    assert.equal(profile.missingCount, 0);
    assert.equal(profile.conversionFailureCount, 1);
    assert.deepEqual(profile.conversionFailureSamples, ["待定"]);
  });
}

test("v2.8.1 empty, blank, boolean and number-like cells under a date format keep their own values", async () => {
  const cells = {
    A1: textCell("名"), B1: textCell("值"),
    A2: textCell("空串"), B2: textCell("", DATE_FORMAT),
    A3: textCell("全空格"), B3: textCell("  ", DATE_FORMAT),
    A4: textCell("公式空串"), B4: { t: "s", v: "", f: 'IF(1=1,"",1)', z: DATE_FORMAT },
    A5: textCell("布尔"), B5: { t: "b", v: true, z: DATE_FORMAT },
    A6: textCell("数字文本"), B6: textCell("45418", DATE_FORMAT),
    A7: textCell("真日期"), B7: dateCell(45418)
  };
  const scope = await importCellSheet(cells, "A1:B7");
  const values = cellSheetColumn(scope, "值");
  assert.deepEqual(values, ["", "  ", "", "true", "45418", "2024-05-06"]);
  assert.equal(cellSheetProfile(scope, "值").missingCount, 3);
  assert.ok(!values.includes("1899-12-31"));
  assert.ok(!values.includes("1900-01-01"));
});

test("v2.8.1 guard: dates stay dates and cells without a date format read as before", async () => {
  const cells = {
    A1: textCell("真日期"), B1: textCell("文本"), C1: textCell("数字"), D1: textCell("小数"), E1: textCell("带时间"),
    A2: dateCell(45418), B2: textCell("待定"), C2: { t: "n", v: 45418 }, D2: { t: "n", v: 3.5, z: "0.00" }, E2: { t: "n", v: 45418.4375, z: "yyyy-mm-dd hh:mm" },
    A3: dateCell(45419), B3: textCell("2021/3/5"), C3: { t: "n", v: 7 }, D3: { t: "n", v: 45418, z: "0.00" }, E3: { t: "n", v: 45419.5, z: "yyyy-mm-dd hh:mm" },
    A4: dateCell(45420), B4: textCell("  "), C4: { t: "b", v: true }, D4: textCell("x", "0.00"), E4: dateCell(45420)
  };
  const scope = await importCellSheet(cells, "A1:E4");
  assert.deepEqual(cellSheetColumn(scope, "真日期"), ["2024-05-06", "2024-05-07", "2024-05-08"]);
  assert.deepEqual(cellSheetColumn(scope, "文本"), ["待定", "2021/3/5", "  "]);
  assert.deepEqual(cellSheetColumn(scope, "数字"), ["45418", "7", "true"]);
  assert.deepEqual(cellSheetColumn(scope, "小数"), ["3.5", "45418", "x"]);
  assert.deepEqual(cellSheetColumn(scope, "带时间"), ["2024-05-06", "2024-05-07", "2024-05-08"]);
});

test("v2.8.1 cells under a date format are found on a sheet that does not start at A1", async () => {
  const cells = {
    B3: textCell("名"), C3: textCell("入职日期", DATE_FORMAT),
    B4: textCell("甲"), C4: dateCell(45418),
    B5: textCell("乙"), C5: textCell("待定", DATE_FORMAT),
    B6: textCell("丙"), C6: { t: "b", v: true, z: DATE_FORMAT }
  };
  const scope = await importCellSheet(cells, "B3:C6");
  assert.deepEqual(cellSheetFields(scope), ["名", "入职日期"]);
  assert.deepEqual(cellSheetColumn(scope, "入职日期"), ["2024-05-06", "待定", "true"]);
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(state.rows.map(getSourceRowNumber))")), [4, 5, 6]);
});

// v2.9.0: date styles, constant columns and look-alike values, only hints. The guard below
// keeps everything v2.8.1 computed; the fixture 写法不一致.csv holds what the checks look for.
test("v2.9.0 fixture: 写法不一致.csv has the bytes, types and totals of known-issues.json", async () => {
  const bytes = fs.readFileSync(path.join(STAGE3_FIXTURES, "写法不一致.csv"));
  assert.deepEqual([...bytes.subarray(0, 3)], [0xef, 0xbb, 0xbf]);
  const text = bytes.toString("utf8");
  assert.equal(text.split("\r\n").length - 1, 25);
  assert.equal(text.split("\n").length - 1, 25);
  // The cells that are the point of the fixture, written as escapes here too.
  assert.ok(text.includes("\u3000华东,"));
  assert.ok(text.includes("华东\u200B,"));
  assert.ok(text.includes("\uFF28\uFF24,"));

  const scope = await importStage3Fixture("写法不一致.csv");
  const expected = knownIssues["写法不一致.csv"];
  const summary = importedSummary(scope.evaluate);
  assert.equal(summary.rows, expected.dataRows);
  assert.deepEqual(summary.fields, expected.expectedTypes.map((item) => item.field));
  assert.deepEqual(summary.types, expected.expectedTypes.map((item) => item.expected));
  assert.equal(summary.duplicateRows, expected.expectedTotals.duplicateRows);
  assert.equal(summary.totalMissing, expected.expectedTotals.totalMissing);
  assert.equal(scope.evaluate("state.profiles.reduce((sum, profile) => sum + profile.conversionFailureCount, 0)"), expected.expectedTotals.formatProblems);
  assert.equal(scope.evaluate("state.headerDetection"), null);
  // ID duplicates compare as before: SKU-002 with a trailing space repeats, sku-001 does not.
  assert.deepEqual(idDuplicates(scope.evaluate).map((item) => [item.field, item.values.map((value) => [value.value, value.count, value.rowNumbers])]), [["商品编号", [["SKU-002", 2, [4, 5]]]]]);
});

// The state v2.8.1 computed, by file: profiles, statistics, quality figures, existing problem rows.
// Regenerate tests/fixtures/v2.8.1-guard.json (see docs/TESTING.md) or remove this guard in the
// version that changes these results on purpose, and say so in its PR.
const V281_GUARD_KINDS = ["missing", "formulaError", "conversion", "outlier", "duplicateRows", "suspectedTotal", "idDuplicate", "placeholder", "parseWarnings"];
const v281GuardDigest = (scope) => crypto.createHash("sha256").update(scope.evaluate(`JSON.stringify({
  rows: state.rows, profiles: state.profiles, numericStats: state.numericStats, qualityNumericStats: state.qualityNumericStats, categoryStats: state.categoryStats,
  duplicateRows: state.duplicateRows, totalMissing: state.totalMissing, idDuplicates: state.idDuplicates, placeholders: state.placeholders, numericLooking: state.numericLooking,
  suspectedTotals: state.suspectedTotals, formulaErrors: state.formulaErrors, dedup: state.rows.length - state.duplicateRows,
  problems: listProblemCategories().filter((category) => ${JSON.stringify(V281_GUARD_KINDS)}.includes(category.kind)).map((category) => [category, listProblemRows(category).items])
})`)).digest("hex").slice(0, 32);

test("v2.9.0 guard: everything v2.8.1 computed stays the same on the stage3 fixtures and examples", async () => {
  const expected = JSON.parse(fs.readFileSync(path.join(__dirname, "fixtures", "v2.8.1-guard.json"), "utf8"));
  const actual = {};
  for (const fileName of STAGE3_IMPORT_FILES) {
    if (["写法不一致.csv", "en-issues.csv", "en-formula-errors.xlsx"].includes(fileName)) continue;
    const scope = await importStage3Fixture(fileName);
    actual[fileName] = v281GuardDigest(scope);
    scope.evaluate("confirmAllCommonPlaceholders()");
    actual[`${fileName}（全部确认后）`] = v281GuardDigest(scope);
  }
  for (const fileName of EXAMPLE_CSV_FILES) {
    // v2.11.0: the new sample is newer than the guard, like 写法不一致.csv above.
    if (["sample-orders.csv", "sample-orders-en.csv"].includes(fileName)) continue;
    const scope = await importCsvFile(readExampleCsv(fileName), fileName);
    actual[fileName] = v281GuardDigest(scope);
    scope.evaluate("confirmAllCommonPlaceholders()");
    actual[`${fileName}（全部确认后）`] = v281GuardDigest(scope);
  }
  assert.deepEqual(actual, expected);
});

test("v2.9.0 Excel: date cells are recorded per row and column, with title rows and a shifted !ref", () => {
  const scope = createExcelContext();
  const worksheet = XLSX.utils.aoa_to_sheet([
    ["2024 报表"],
    [],
    ["名", "日期", "文本日期"],
    ["甲", 45000, "2024-01-01"],
    ["", "", ""],
    ["乙", "2024-01-01", 45001],
    ["丙", 45002, "x"]
  ], { origin: "B3" });
  for (const address of ["C6", "D8", "C9"]) {
    worksheet[address].t = "n";
    worksheet[address].z = DATE_FORMAT;
  }
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "S");
  put(scope.context, "__buffer", XLSX.write(workbook, { type: "buffer", bookType: "xlsx" }));
  scope.evaluate('importExcelSheet(readExcelSheet(__buffer, "S"), "s.xlsx", "S", beginImport())');
  const read = JSON.parse(scope.evaluate("JSON.stringify(state.rows.map((row) => [getSourceRowNumber(row), state.sourceDateCells.get(row) || null, row.日期, row.文本日期]))"));
  assert.deepEqual(read.map(([line, cells]) => [line, cells && Object.keys(cells).sort()]), [[6, ["日期"]], [8, ["文本日期"]], [9, ["日期"]]]);
  // A date cell and a text with the same characters read alike but are told apart.
  assert.equal(read[1][2], "2024-01-01");
  assert.equal(read[1][1].日期, undefined);
});

test("v2.9.0 Excel: only real date cells count, and a file without them adds nothing", async () => {
  const cells = {
    A1: textCell("名"), B1: textCell("值", DATE_FORMAT),
    A2: textCell("真日期"), B2: dateCell(45418),
    A3: textCell("文本"), B3: textCell("待定", DATE_FORMAT),
    A4: textCell("空串"), B4: textCell("", DATE_FORMAT),
    A5: textCell("布尔"), B5: { t: "b", v: true, z: DATE_FORMAT },
    A6: textCell("文本日期"), B6: textCell("2024-05-06", DATE_FORMAT),
    A7: textCell("数字"), B7: { t: "n", v: 45418, z: "0.00" },
    // Known limitation: a time-only cell is a real number cell and reads as this date.
    A8: textCell("只有时间"), B8: { t: "n", v: 0.4375, z: "hh:mm" }
  };
  const scope = await importCellSheet(cells, "A1:B8");
  const read = JSON.parse(scope.evaluate("JSON.stringify(state.rows.map((row) => [row.名, row.值, Boolean(state.sourceDateCells.get(row)?.值)]))"));
  assert.deepEqual(read, [
    ["真日期", "2024-05-06", true],
    ["文本", "待定", false],
    ["空串", "", false],
    ["布尔", "true", false],
    ["文本日期", "2024-05-06", false],
    ["数字", "45418", false],
    ["只有时间", "1899-12-31", true]
  ]);
  const plain = createExcelContext();
  const direct = plain.evaluate('buildExcelTabularData([["名", "日"], ["甲", "x"]], [["名", "日"], ["甲", "x"]])');
  assert.equal(direct.dateCells, undefined);
  // A Date that reads as empty is no date cell.
  const invalid = plain.evaluate('buildExcelTabularData([["名", "日"], ["甲", new Date(NaN)]], [["名", "日"], ["甲", new Date(NaN)]])');
  assert.equal(invalid.rows[0].日, "");
  assert.equal(invalid.dateCells, undefined);
});

test("v2.9.0 state: a commit replaces the date cells and resetAnalysisState() clears them", async () => {
  const roster = await importStage3Fixture("花名册.xlsx");
  put(roster.context, "__keptRow", roster.evaluate("state.rows[0]"));
  assert.equal(roster.evaluate("state.sourceDateCells.has(__keptRow)"), true);
  // Without beginImport(): the commit itself replaces them.
  roster.evaluate('commitTabularData(["城市"], [{ 城市: "上海" }], "later.csv", state.activeImportId)');
  assert.equal(roster.evaluate("state.sourceDateCells.has(__keptRow)"), false);
  // A list that does not match the rows is left out, as for row numbers and formula errors.
  roster.evaluate('commitTabularData(["日期"], [{ 日期: "2024-01-01" }], "short.csv", state.activeImportId, [], "", null, null, null, null, [{ 日期: true }, { 日期: true }])');
  assert.equal(roster.evaluate("state.sourceDateCells.has(state.rows[0])"), false);

  const cleared = await importStage3Fixture("花名册.xlsx");
  put(cleared.context, "__keptRow", cleared.evaluate("state.rows[0]"));
  cleared.evaluate("resetAnalysisState()");
  assert.equal(cleared.evaluate("state.sourceDateCells.has(__keptRow)"), false);
});

const dateStylesOf = (scope) => JSON.parse(scope.evaluate("JSON.stringify(state.dateStyles)"));
const setFieldType = (scope, field, typeKey) => scope.evaluate(`state.fieldTypeDrafts[${JSON.stringify(field)}] = ${JSON.stringify(typeKey)}; applyFieldConfiguration()`);

// [written, order, style].
const DATE_STYLE_CASES = [
  ["2024-05-06", "none", "ymdDash"], ["2024-5-6", "none", "ymdDash"], ["2024/05/06", "none", "ymdSlash"], ["2024/5/6", "none", "ymdSlash"],
  ["2024.5.6", "none", "ymdDot"], ["2024年5月6日", "none", "chinese"], ["2024年5月6", "none", "chinese"], ["2024 年 5 月 6 日", "none", "chinese"],
  ["2024-05-06 10:30", "none", "ymdDash"], ["2024-05-06T10:30:00Z", "none", "ymdDash"], ["06/05/2024", "dmy", "dmy"], ["06/05/2024", "mdy", "mdy"],
  ["06-05-2024", "dmy", "dmy"], ["06.05.2024", "mdy", "mdy"],
  ["06/05/2024", "none", null], ["06/05/2024", "conflict", null], ["05/05/2024", "none", "dayMonthSame"], ["05/05/2024", "conflict", "dayMonthSame"],
  ["05/05/2024", "dmy", "dmy"], ["13/05/2024", "conflict", "dmy"], ["05/13/2024", "conflict", "mdy"], ["13/05/2024", "mdy", "dmy"],
  ["May 6, 2024", "none", "english"], ["6 May 2024", "none", "english"], ["Mon, 6 May 2024", "none", "english"], ["2024-May-06", "none", "english"],
  ["\u00A02024/5/6 ", "none", "ymdSlash"], ["\u30002024年5月6日", "none", "chinese"], ["2024-13-01", "none", null], ["2024/5/6/7", "none", null],
  ["待定", "none", null], ["20240506", "none", null], ["", "none", null], ["N/A", "none", null]
];

test("v2.9.0 dateStyleOf: each writing's style, padding and a time part make no new style", () => {
  const { evaluate, context } = createScriptContext();
  for (const [text, order, style] of DATE_STYLE_CASES) {
    put(context, "__text", text);
    assert.equal(evaluate(`dateStyleOf(__text, { order: ${JSON.stringify(order)} })`), style, `${text} ${order}`);
  }
  assert.equal(evaluate('dateStyleOf("2024", { allowYearOnly: true, order: "none" })'), "yearOnly");
  assert.equal(evaluate('dateStyleOf("2024年", { allowYearOnly: true, order: "none" })'), "yearOnly");
  assert.equal(evaluate('dateStyleOf("2024", { order: "none" })'), null);
});

test("v2.9.0 guard: dateStyleOf() is null exactly when toDate() is null", () => {
  const { evaluate } = createScriptContext();
  const result = JSON.parse(evaluate(`JSON.stringify((() => {
    const days = ["1", "01", "6", "06", "12", "13", "29", "31", "32", "0", "00"];
    const years = ["2024", "1999", "0999", "24", "20240"];
    const times = ["", " 10:30", "T10:30:00Z", " 25:00", " 10:30:00.123+08:00", " 10:30 PM"];
    const months = ["May", "may.", "sept", "Foo", "december"];
    const texts = [];
    for (const year of years) for (const month of days) for (const day of days) for (const time of times) {
      for (const separator of ["-", "/", ".", " ", ""]) {
        texts.push(year + separator + month + separator + day + time, day + separator + month + separator + year + time);
      }
      texts.push(year + "年" + month + "月" + day + "日" + time, year + " 年 " + month + " 月 " + day + time);
    }
    for (const month of months) for (const day of days) for (const year of years) for (const time of times) {
      texts.push(month + " " + day + ", " + year + time, day + " " + month + " " + year + time, "Mon, " + day + " " + month + " " + year + time, year + "-" + month + "-" + day);
    }
    const strategies = [{ order: "none" }, { order: "dmy" }, { order: "mdy" }, { order: "conflict" }, { order: "none", allowYearOnly: true }];
    let checked = 0;
    let readable = 0;
    const mismatches = [];
    const styles = new Set();
    for (const text of [...texts, "2024", "2024年", " 2024 ", "", " ", "N/A", "—"]) {
      for (const strategy of strategies) {
        for (const value of [text, " " + text + "\\u3000"]) {
          const style = dateStyleOf(value, strategy);
          const date = toDate(value, strategy);
          checked += 1;
          if (date) readable += 1;
          if (style) styles.add(style);
          if ((style === null) !== (date === null)) mismatches.push([value, strategy.order, style]);
        }
      }
    }
    return { checked, readable, mismatches: mismatches.slice(0, 10), styles: [...styles].sort() };
  })())`));
  assert.deepEqual(result.mismatches, []);
  assert.ok(result.checked > 100000, String(result.checked));
  assert.ok(result.readable > 5000, String(result.readable));
  assert.deepEqual(result.styles, ["chinese", "dayMonthSame", "dmy", "english", "mdy", "yearOnly", "ymdDash", "ymdDot", "ymdSlash"]);
});

test("v2.9.0 date styles: the fixture, the roster and the sample button's file (v2.11.0), and no other file", async () => {
  const scopes = await importEveryQualityFile();
  const expected = knownIssues["写法不一致.csv"].dateStyles;
  const writing = scopes.find(([name]) => name === "写法不一致.csv")[1];
  assert.deepEqual(dateStylesOf(writing).map(({ field, mainKey, otherCount, styles }) => ({ field, mainKey, otherCount, styles })), expected.map(({ field, mainKey, otherCount, styles }) => ({ field, mainKey, otherCount, styles })));
  const roster = scopes.find(([name]) => name === "花名册.xlsx")[1];
  assert.deepEqual(dateStylesOf(roster), [{
    field: "入职日期", mainKey: "excelCell", otherCount: 6,
    styles: [
      { key: "excelCell", count: 24, sample: "2016-04-01", firstLine: 2 },
      { key: "ymdSlash", count: 2, sample: "2021/3/5", firstLine: 6 },
      { key: "chinese", count: 2, sample: "2021年3月5日", firstLine: 10 },
      { key: "ymdDot", count: 2, sample: "2021.03.05", firstLine: 15 }
    ]
  }]);
  assert.deepEqual(scopes.filter(([, scope]) => dateStylesOf(scope).length).map(([name]) => name), ["写法不一致.csv", "花名册.xlsx", "sample-orders-en.csv", "sample-orders.csv"]);
});

test("v2.9.0 date styles: the most frequent style is the main one, ties go to the first seen", async () => {
  // The first value is not the most frequent one, so the order tells the rule apart.
  const scope = await importCsvFile("日期,序\r\n2024/1/1,1\r\n2024-01-02,2\r\n2024-01-03,3\r\n2024.1.4,4\r\n2024.1.5,5\r\n", "m.csv");
  assert.deepEqual(dateStylesOf(scope)[0].styles.map(({ key, count }) => [key, count]), [["ymdDash", 2], ["ymdDot", 2], ["ymdSlash", 1]]);
  assert.equal(dateStylesOf(scope)[0].mainKey, "ymdDash");
  assert.equal(dateStylesOf(scope)[0].otherCount, 3);
  // Padding and a time part alone are one style.
  const same = await importCsvFile("日期,序\r\n2024-1-1,1\r\n2024-01-02,2\r\n2024-01-03 10:30,3\r\n", "s.csv");
  assert.deepEqual(dateStylesOf(same), []);
  // A value that cannot be read is a format problem, not a style.
  const unread = await importCsvFile("日期,序\r\n2024-01-01,1\r\n2024-01-02,2\r\n2024-01-03,3\r\n2024-01-04,4\r\n待定,5\r\n", "u.csv");
  assert.equal(unread.evaluate("state.profiles[0].typeKey"), "date");
  assert.deepEqual(dateStylesOf(unread), []);
});

test("v2.9.0 a field read by year has no date styles", async () => {
  const scope = await importCsvFile(`年份\r\n${["2015", "2016", "2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024/1/1"].join("\r\n")}\r\n`, "y.csv");
  assert.equal(scope.evaluate("state.profiles[0].typeKey"), "date");
  assert.equal(scope.evaluate("state.profiles[0].dateStrategy.granularity"), "year");
  assert.deepEqual(dateStylesOf(scope), []);
});

test("v2.9.0 date styles follow the field configuration and confirmations, not the filters", async () => {
  const writing = await importStage3Fixture("写法不一致.csv");
  writing.evaluate('__styles = state.dateStyles; state.filters.category = { field: "状态", values: ["正常"] }; state.filteredRows = state.rows.slice(0, 3); refreshFilteredAnalysis()');
  assert.equal(writing.evaluate("__styles === state.dateStyles"), true);
  assert.equal(writing.evaluate("getAnalysisRows().length"), 3);
  setFieldType(writing, "下单日期", "category");
  assert.deepEqual(dateStylesOf(writing), []);
  writing.evaluate("restoreAutomaticFieldTypes()");
  assert.equal(dateStylesOf(writing)[0].otherCount, 6);
  // A confirmation rebuilds them too, with the same result.
  writing.evaluate('__styles = state.dateStyles; confirmPlaceholder("客户", "N/A")');
  assert.equal(writing.evaluate("__styles === state.dateStyles"), false);
  assert.equal(dateStylesOf(writing)[0].otherCount, 6);

  // A text field made a date field is read then, with its own strategy.
  const later = await importCsvFile("文字,序\r\n2024-01-01,1\r\n2024/1/2,2\r\nx,3\r\ny,4\r\nz,5\r\n", "l.csv");
  assert.equal(later.evaluate("state.profiles[0].typeKey"), "category");
  assert.deepEqual(dateStylesOf(later), []);
  setFieldType(later, "文字", "date");
  assert.deepEqual(dateStylesOf(later)[0].styles.map(({ key, count }) => [key, count]), [["ymdDash", 1], ["ymdSlash", 1]]);

  const roster = await importStage3Fixture("花名册.xlsx");
  setFieldType(roster, "入职日期", "category");
  assert.deepEqual(dateStylesOf(roster), []);
  setFieldType(roster, "入职日期", "date");
  assert.equal(dateStylesOf(roster)[0].otherCount, 6);
});

test("v2.9.0 a rebuild reads the date cells only for a field it has not scanned yet", async () => {
  const scope = await importStage3Fixture("写法不一致.csv");
  const count = (action) => scope.evaluate(`(() => {
    const each = forEachDateStyle;
    let calls = 0;
    forEachDateStyle = (...args) => { calls += 1; return each(...args); };
    ${action};
    forEachDateStyle = each;
    return calls;
  })()`);
  assert.equal(count("applyFieldConfiguration()"), 0);
  assert.equal(count('confirmPlaceholder("客户", "N/A")'), 0);
  assert.equal(count('state.fieldTypeDrafts["状态"] = "date"; applyFieldConfiguration()'), 1);
  assert.equal(count('state.fieldTypeDrafts["状态"] = "category"; applyFieldConfiguration()'), 0);
  // Each distinct text is classified once: 3000 rows, 3 texts.
  const small = createScriptContext();
  put(small.context, "__rows", Array.from({ length: 3000 }, (_, index) => ({ 日期: ["2024-01-01", "2024/1/2", "2024.1.3"][index % 3] })));
  assert.equal(small.evaluate(`(() => {
    const style = dateStyleOf;
    let calls = 0;
    dateStyleOf = (...args) => { calls += 1; return style(...args); };
    const profile = buildColumnProfile("日期", __rows);
    buildDateStyleSummary(__rows, [profile], () => undefined, () => null);
    dateStyleOf = style;
    return calls;
  })()`), 3);
});

test("v2.9.0 state: a later commit and resetAnalysisState() replace the date styles and their scans", async () => {
  const scope = await importStage3Fixture("写法不一致.csv");
  // The same field name with other cells, without beginImport().
  scope.evaluate('commitTabularData(["下单日期"], Array.from({ length: 3 }, (_, index) => ({ 下单日期: "2024-01-0" + (index + 1) })), "later.csv", state.activeImportId)');
  assert.deepEqual(dateStylesOf(scope), []);
  const cleared = await importStage3Fixture("写法不一致.csv");
  assert.equal(dateStylesOf(cleared).length, 1);
  cleared.evaluate("resetAnalysisState()");
  assert.deepEqual(dateStylesOf(cleared), []);
  assert.equal(cleared.evaluate("state.formatScans.dateStyles.size"), 0);
});

test("v2.9.0 texts: the date style line and how whitespace is shown", async () => {
  const roster = await importStage3Fixture("花名册.xlsx");
  assert.equal(roster.evaluate("formatDateStyleLine(state.dateStyles[0])"), "「入职日期」：主写法是 Excel 日期单元格 24 个（如 2016-04-01）；另有 6 行是别的写法：文本 年/月/日 2 个（如 2021/3/5）、文本 中文年月日 2 个（如 2021年3月5日）、文本 年.月.日 2 个（如 2021.03.05）。");
  const { evaluate, context } = createScriptContext();
  const shown = (text) => {
    put(context, "__text", text);
    return evaluate("formatVisibleText(__text)");
  };
  assert.equal(shown("华东 "), "华东␣");
  assert.equal(shown("  华东"), "␣␣华东");
  assert.equal(shown("上海  浦东"), "上海␣␣浦东");
  assert.equal(shown("上海 浦东"), "上海 浦东");
  assert.equal(shown("\u3000华东\u00A0"), "[全角空格]华东[不换行空格]");
  assert.equal(shown("华\u200B东\uFEFF"), "华[零宽字符]东[零宽字符]");
  assert.equal(shown("华\u00AD东\u200E"), "华[软连字符]东[方向符]");
  assert.equal(shown("a\tb"), "a[制表符]b");
  assert.equal(shown("\u2003x"), "[空白]x");
});

const constantsOf = (scope) => JSON.parse(scope.evaluate("JSON.stringify(state.constantFields)"));

test("v2.9.0 constant columns: the fixture and the formula-error sheet, and no other file", async () => {
  const scopes = await importEveryQualityFile();
  assert.deepEqual(constantsOf(scopes.find(([name]) => name === "写法不一致.csv")[1]), knownIssues["写法不一致.csv"].constantFields);
  const errors = scopes.find(([name]) => name === "公式错误.xlsx")[1];
  assert.deepEqual(constantsOf(errors), [{ field: "备注", value: "正常", count: 6, missingCount: 3 }]);
  assert.equal(errors.evaluate("formatConstantField(state.constantFields[0])"), "「备注」：只有一种取值「正常」（6 行），另有 3 行是空值。");
  assert.deepEqual(scopes.filter(([, scope]) => constantsOf(scope).length).map(([name]) => name), ["en-issues.csv", "公式错误.xlsx", "写法不一致.csv"]);
  assert.equal(scopes.find(([name]) => name === "写法不一致.csv")[1].evaluate("formatConstantField(state.constantFields[0])"), "「状态」：只有一种取值「正常」（24 行）。");
});

test("v2.9.0 constant columns: at least five values, one distinct value after trimming", async () => {
  const columns = async (cells) => constantsOf(await importCsvFile(`值,序\r\n${cells.map((cell, index) => `${cell},${index + 1}`).join("\r\n")}\r\n`, "k.csv"));
  assert.deepEqual(await columns(["a", "a", "a", "a"]), []);
  assert.deepEqual(await columns(["a", "a", "a", "a", "a"]), [{ field: "值", value: "a", count: 5, missingCount: 0 }]);
  assert.deepEqual(await columns(["a ", " a", "a", "a", "a", ""]), [{ field: "值", value: "a", count: 5, missingCount: 1 }]);
  assert.deepEqual(await columns(["1", "1.0", "1", "1", "1"]), []);
  assert.deepEqual(await columns(["1", "\uFF11", "1", "1", "1"]), []);
  assert.deepEqual(await columns(["a", "A", "a", "a", "a"]), []);
  // Every applied type counts; an all-empty column is none.
  assert.deepEqual(await columns(["", "", "", "", ""]), []);
  const types = await importCsvFile("编号,金额,日期\r\nA1,5,2024-01-01\r\nA1,5,2024-01-01\r\nA1,5,2024-01-01\r\nA1,5,2024-01-01\r\nA1,5,2024-01-01\r\n", "t.csv");
  assert.deepEqual(JSON.parse(types.evaluate("JSON.stringify(state.profiles.map((profile) => profile.typeKey))")), ["id", "numeric", "date"]);
  assert.deepEqual(constantsOf(types).map((item) => item.field), ["编号", "金额", "日期"]);
});

test("v2.9.0 constant columns: a column of one placeholder spelling is left to the placeholder section", async () => {
  const single = async (value) => constantsOf(await importCsvFile(`状态,序\r\n${Array.from({ length: 6 }, (_, index) => `${value},${index + 1}`).join("\r\n")}\r\n`, "p.csv"));
  assert.deepEqual(await single("N/A"), []);
  assert.deepEqual(await single("无"), []);
  assert.deepEqual(await single("Ｎ/\uFF21"), []);
  assert.equal((await single("正常")).length, 1);
  // Still listed as a placeholder, confirmed or not, and no constant column either way.
  const listed = await importCsvFile("状态,序\r\nN/A,1\r\nN/A,2\r\nN/A,3\r\nN/A,4\r\nN/A,5\r\n", "l.csv");
  assert.equal(listed.evaluate("state.placeholders.length"), 1);
  assert.deepEqual(constantsOf(listed), []);
  listed.evaluate('confirmPlaceholder("状态", "N/A")');
  assert.deepEqual(constantsOf(listed), []);
});

test("v2.9.0 constant columns follow the field configuration and confirmations, not the filters", async () => {
  const writing = await importStage3Fixture("写法不一致.csv");
  writing.evaluate('__constants = state.constantFields; state.filters.category = { field: "状态", values: ["正常"] }; state.filteredRows = state.rows.slice(0, 3); refreshFilteredAnalysis()');
  assert.equal(writing.evaluate("__constants === state.constantFields"), true);
  setFieldType(writing, "状态", "ignore");
  assert.deepEqual(constantsOf(writing), []);
  // Any applied type but ignored.
  setFieldType(writing, "状态", "numeric");
  assert.deepEqual(constantsOf(writing).map((item) => item.field), ["状态"]);
  writing.evaluate("restoreAutomaticFieldTypes()");
  assert.deepEqual(constantsOf(writing).map((item) => item.field), ["状态"]);

  // Confirming the other value leaves one, and the confirmed rows count as missing.
  const confirmed = await importCsvFile("状态,序\r\n正常,1\r\n正常,2\r\n正常,3\r\n正常,4\r\n无,5\r\n正常,6\r\n", "c.csv");
  assert.deepEqual(constantsOf(confirmed), []);
  confirmed.evaluate('confirmPlaceholder("状态", "无")');
  assert.deepEqual(constantsOf(confirmed), [{ field: "状态", value: "正常", count: 5, missingCount: 1 }]);
  assert.equal(confirmed.evaluate("formatConstantField(state.constantFields[0])"), "「状态」：只有一种取值「正常」（5 行），另有 1 行是空值。");
  confirmed.evaluate('undoPlaceholderConfirmation("状态", "无")');
  assert.deepEqual(constantsOf(confirmed), []);

  // Confirming below the minimum ends it.
  const below = await importCsvFile("状态,序\r\n正常,1\r\n正常,2\r\n正常,3\r\n正常,4\r\n无,5\r\n", "b.csv");
  below.evaluate('confirmPlaceholder("状态", "无")');
  assert.deepEqual(constantsOf(below), []);
});

test("v2.9.0 constant columns read the profiles only, and a commit or reset replaces them", async () => {
  const { evaluate } = createScriptContext();
  // No rows at all: the figures come from the profile.
  assert.deepEqual(JSON.parse(evaluate('JSON.stringify(buildConstantFieldSummary([{ field: "甲", typeKey: "category", uniqueCount: 1, nonMissingCount: 5, missingCount: 2, sampleValues: ["好"] }]))')), [{ field: "甲", value: "好", count: 5, missingCount: 2 }]);
  const scope = await importStage3Fixture("写法不一致.csv");
  assert.equal(constantsOf(scope).length, 1);
  scope.evaluate('commitTabularData(["城市"], [{ 城市: "上海" }, { 城市: "北京" }], "later.csv", state.activeImportId)');
  assert.deepEqual(constantsOf(scope), []);
  const cleared = await importStage3Fixture("写法不一致.csv");
  cleared.evaluate("resetAnalysisState()");
  assert.deepEqual(constantsOf(cleared), []);
});

const lookAlikesOf = (scope) => JSON.parse(scope.evaluate("JSON.stringify(state.lookAlikes)"));

test("v2.9.0 hasEdgeWhitespace() agrees with trim() on every UTF-16 unit at either end", () => {
  const { evaluate } = createScriptContext();
  const mismatches = JSON.parse(evaluate(`JSON.stringify((() => {
    const bad = [];
    for (let code = 0; code <= 0xffff; code += 1) {
      const character = String.fromCharCode(code);
      for (const text of [character + "x", "x" + character, "x" + character + "x"]) {
        if (hasEdgeWhitespace(text) !== (text.trim() !== text)) bad.push(code);
      }
    }
    return bad.slice(0, 10);
  })())`));
  assert.deepEqual(mismatches, []);
  assert.equal(evaluate('hasEdgeWhitespace("\\u200b华东")'), false);
  assert.equal(evaluate('hasEdgeWhitespace("\\ufeff华东")'), true);
  assert.equal(evaluate('hasEdgeWhitespace("")'), false);
});

test("v2.9.0 lookAlikeKey: what merges and what stays apart", () => {
  const { evaluate, context } = createScriptContext();
  const same = (a, b) => {
    put(context, "__a", a);
    put(context, "__b", b);
    return evaluate("lookAlikeKey(__a) === lookAlikeKey(__b)");
  };
  for (const [a, b] of [
    ["ABC公司", "\uFF21\uFF22\uFF23公司"], ["ABC公司", "abc公司"], ["华东（一）", "华东(一)"], ["华东", "华东\u200B"], ["华东", "\u200B华东"], ["华东", "华\u200B东"],
    ["华东", "华\u200C东"], ["华东", "华\u200D东"], ["华东", "华\u2060东"], ["华东", "华\uFEFF东"], ["华东", "华东\u00AD"], ["华东", "华\u00AD东"], ["华东", "华东\u200E"], ["华东", "华东\u200F"],
    ["上海  浦东", "上海 浦东"], ["上海\u3000浦东", "上海 浦东"], ["上海\t浦东", "上海 浦东"],
    ["Ⅱ", "II"], ["①", "1"], ["ﾃｽﾄ", "テスト"], ["İ", "i̇"], ["K", "k"], ["㎡", "m2"], ["华东 \u200B", "华东"]
  ]) assert.equal(same(a, b), true, `${a} / ${b}`);
  for (const [a, b] of [["上海 浦东", "上海浦东"], ["华东", "华东区"], ["華东", "华东"], ["张 三", "张三"], ["ı", "i"], ["华东", "华东\uFE0F"]]) assert.equal(same(a, b), false, `${a} / ${b}`);
  for (const text of ["\u200B\u200C", "\u00AD", "\u200E", "\u200F", "\u2060"]) {
    put(context, "__text", text);
    assert.equal(evaluate("lookAlikeKey(__text)"), "", text);
  }
});

test("v2.9.0 lookAlikeKey leaves pure ASCII texts out of NFKC", () => {
  const scope = createScriptContext();
  put(scope.context, "__rows", Array.from({ length: 200 }, (_, index) => ({ 编号: `ORDER-LONG-${index}` })));
  const calls = scope.evaluate(`(() => {
    const original = String.prototype.normalize;
    let count = 0;
    String.prototype.normalize = function (...args) { count += 1; return original.apply(this, args); };
    buildLookAlikeSummary(__rows, [{ field: "编号", typeKey: "id" }]);
    String.prototype.normalize = original;
    return count;
  })()`);
  assert.equal(calls, 0);
});

test("v2.9.0 lookAlikeReasons: one writing's difference names only its own reason", () => {
  const { evaluate, context } = createScriptContext();
  const reasons = (texts) => {
    put(context, "__texts", texts);
    return JSON.parse(evaluate("JSON.stringify(lookAlikeReasons(__texts))"));
  };
  assert.deepEqual(reasons(["ABC", "abc"]), ["case"]);
  assert.deepEqual(reasons(["ABC", "\uFF21\uFF22\uFF23"]), ["compat"]);
  assert.deepEqual(reasons(["HD", "\uFF28\uFF24", "hd"]), ["compat", "case"]);
  // Both full-width, one with a zero-width character: not a width difference.
  assert.deepEqual(reasons(["\uFF21\uFF22", "\uFF21\uFF22\u200B"]), ["invisible"]);
  assert.deepEqual(reasons(["华东", "华东\u00AD", "华东\u200E"]), ["invisible"]);
  assert.deepEqual(reasons(["上海 浦东", "上海  浦东"]), ["innerSpace"]);
  assert.deepEqual(reasons(["上海 浦东", "上海\u3000浦东"]), ["compat"]);
  assert.deepEqual(reasons(["华东 \u200B", "华东"]), ["invisible"]);
  // Either step alone unifies the Kelvin sign and K: still some reason.
  assert.deepEqual(reasons(["K", "K"]), ["compat"]);
});

test("v2.9.0 look-alike values: the fixture and the sample button's file (v2.11.0), and no other file", async () => {
  const scopes = await importEveryQualityFile();
  const writing = scopes.find(([name]) => name === "写法不一致.csv")[1];
  const summary = lookAlikesOf(writing);
  assert.deepEqual(summary.map(({ field, edgeSpaceCount, groupCount, rowCount, groups }) => ({
    field, edgeSpaceCount, groupCount, rowCount,
    groups: groups.map(({ variants, reasons, rowCount: count }) => ({ variants: variants.map(({ text, count: n }) => [text, n]), reasons, rowCount: count }))
  })), knownIssues["写法不一致.csv"].lookAlikes);
  assert.ok(summary.every((item) => item.invisibleOnlyCount === 0));
  assert.deepEqual(scopes.filter(([, scope]) => lookAlikesOf(scope).length).map(([name]) => name), ["en-issues.csv", "写法不一致.csv", "sample-orders-en.csv", "sample-orders.csv"]);
  // ID duplicates and look-alikes count apart: SKU-002 with a space is a duplicate value and
  // an edge space, SKU-001 and sku-001 are a look-alike and no duplicate.
  assert.deepEqual(idDuplicates(writing.evaluate)[0].values.map((value) => value.value), ["SKU-002"]);
  assert.equal(summary[0].edgeSpaceCount, 1);
  assert.equal(writing.evaluate("formatLookAlikeLine(state.lookAlikes[1])"), "「区域」共 18 行：首尾空格 2 格（如「华东␣」）；3 组写法不一样：「华东」10 个 / 「华东[零宽字符]」1 个（不可见字符）；「HD」2 个 / 「\uFF28\uFF24」1 个 / 「hd」1 个（全角半角或兼容写法、大小写）；「上海 浦东」2 个 / 「上海␣␣浦东」1 个（中间空格多少）。");
});

test("v2.9.0 look-alike values skip placeholders, numbers and dates", async () => {
  const scope = await importCsvFile("分类,数值,日期\r\n N/A,1,2024-01-01\r\nn/a,1.0,2024-01-02\r\nx , 1,2024-01-03 \r\nx,1,2024-01-04\r\ny,2,2024-01-05\r\n", "l.csv");
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(state.profiles.map((profile) => profile.typeKey))")), ["category", "numeric", "date"]);
  assert.deepEqual(lookAlikesOf(scope).map((item) => [item.field, item.edgeSpaceCount, item.groupCount, item.rowCount]), [["分类", 1, 0, 1]]);
});

test("v2.9.0 cells of nothing but invisible characters are counted per field", async () => {
  const scope = await importCsvFile("备注,序\r\n\u200B,1\r\n\u00AD,2\r\n\u200E,3\r\n甲,4\r\n乙,5\r\n丙,6\r\n", "z.csv");
  // They are not missing values: trim() cannot remove them.
  assert.equal(scope.evaluate('getFieldProfile("备注").missingCount'), 0);
  const [item] = lookAlikesOf(scope);
  assert.deepEqual([item.field, item.invisibleOnlyCount, item.edgeSpaceCount, item.groupCount, item.rowCount], ["备注", 3, 0, 0, 3]);
  assert.equal(scope.evaluate("formatLookAlikeLine(state.lookAlikes[0])"), "「备注」共 3 行：只有不可见字符 3 格。");
  // With other findings, and a cell with both counts in both but once in the rows.
  const mixed = await importCsvFile("备注,序\r\n\u200B ,1\r\n甲 ,2\r\nAB,3\r\nab,4\r\n丙,5\r\n", "m.csv");
  const [both] = lookAlikesOf(mixed);
  assert.deepEqual([both.invisibleOnlyCount, both.edgeSpaceCount, both.groupCount, both.rowCount], [1, 2, 1, 4]);
  assert.match(mixed.evaluate("formatLookAlikeLine(state.lookAlikes[0])"), /共 4 行：首尾空格 2 格.*只有不可见字符 1 格；1 组写法不一样/);
  // Zero stays out of the line.
  const none = await importCsvFile("备注,序\r\nAB,1\r\nab,2\r\n", "n.csv");
  assert.doesNotMatch(none.evaluate("formatLookAlikeLine(state.lookAlikes[0])"), /不可见字符/);
  // Fields that are not category or ID are not read.
  const numeric = await importCsvFile("金额,序\r\n1,1\r\n2,2\r\n3,3\r\n4,4\r\n5,5\r\n6,6\r\n7,7\r\n8,8\r\n9,9\r\n\u200B,10\r\n", "o.csv");
  assert.equal(numeric.evaluate('getFieldProfile("金额").typeKey'), "numeric");
  assert.deepEqual(lookAlikesOf(numeric), []);
});

test("v2.9.0 look-alike lists stop at the limits but count everything", async () => {
  const many = Array.from({ length: 12 }, (_, index) => [`K${index}`, `k${index}`]).flat();
  const overflow = await importCsvFile(`分类\r\n${many.join("\r\n")}\r\n`, "o.csv");
  const [item] = lookAlikesOf(overflow);
  assert.deepEqual([item.groups.length, item.groupCount, item.rowCount], [10, 12, 24]);
  assert.match(overflow.evaluate("formatLookAlikeLine(state.lookAlikes[0])"), /12 组写法不一样，只列出前 10 组：/);
  // Twelve writings of one value: ten shown and how many there are.
  const spellings = ["abcd", "Abcd", "aBcd", "abCd", "abcD", "ABcd", "AbCd", "AbcD", "aBCd", "aBcD", "abCD", "ABCd"];
  const variants = await importCsvFile(`分类,序\r\n${spellings.map((text, index) => `${text},${index}`).join("\r\n")}\r\n`, "v.csv");
  // 12 distinct values of 12 make it an ID field by uniqueness: still read.
  const [group] = lookAlikesOf(variants)[0].groups;
  assert.deepEqual([group.variants.length, group.variantCount, group.rowCount], [10, 12, 12]);
  assert.match(variants.evaluate("formatLookAlikeLine(state.lookAlikes[0])"), / 等 12 种（大小写）/);
  // Groups by rows, most first; writings by count, most first, not by first appearance.
  const sorted = await importCsvFile("分类,序\r\nx,1\r\nA,2\r\nb,3\r\nB,4\r\nB,5\r\nB,6\r\nX,7\r\n", "s.csv");
  assert.deepEqual(lookAlikesOf(sorted)[0].groups.map((entry) => entry.variants.map(({ text, count }) => [text, count])), [[["B", 3], ["b", 1]], [["x", 1], ["X", 1]]]);
});

test("v2.9.0 look-alike work runs once per distinct text, and a rebuild reuses it", async () => {
  const small = createScriptContext();
  put(small.context, "__rows", Array.from({ length: 3000 }, (_, index) => ({ 城市: ["上海", "北京", "abc"][index % 3] })));
  assert.equal(small.evaluate(`(() => {
    const original = lookAlikeKey;
    let count = 0;
    lookAlikeKey = (text) => { count += 1; return original(text); };
    const profile = buildColumnProfile("城市", __rows);
    buildLookAlikeSummary(__rows, [profile]);
    lookAlikeKey = original;
    return count;
  })()`), 3);
  const scope = await importStage3Fixture("写法不一致.csv");
  const count = (action) => scope.evaluate(`(() => {
    const collect = collectFieldWritings;
    let calls = 0;
    collectFieldWritings = (...args) => { calls += 1; return collect(...args); };
    ${action};
    collectFieldWritings = collect;
    return calls;
  })()`);
  assert.equal(count("applyFieldConfiguration()"), 0);
  assert.equal(count('confirmPlaceholder("客户", "N/A")'), 0);
  // 金额 was numeric on import, so this is its first scan.
  assert.equal(count('state.fieldTypeDrafts["金额"] = "category"; applyFieldConfiguration()'), 1);
  assert.equal(count('state.fieldTypeDrafts["金额"] = "numeric"; applyFieldConfiguration()'), 0);
});

test("v2.9.0 look-alike values follow the field configuration, not the filters or confirmations", async () => {
  const writing = await importStage3Fixture("写法不一致.csv");
  const fields = () => lookAlikesOf(writing).map((item) => item.field);
  writing.evaluate('__looks = state.lookAlikes; state.filters.category = { field: "状态", values: ["正常"] }; state.filteredRows = state.rows.slice(0, 3); refreshFilteredAnalysis()');
  assert.equal(writing.evaluate("__looks === state.lookAlikes"), true);
  setFieldType(writing, "区域", "ignore");
  assert.deepEqual(fields(), ["商品编号", "客户"]);
  setFieldType(writing, "区域", "numeric");
  assert.deepEqual(fields(), ["商品编号", "客户"]);
  writing.evaluate("restoreAutomaticFieldTypes()");
  assert.deepEqual(fields(), ["商品编号", "区域", "客户"]);
  // Placeholder spellings are left out, confirmed or not: confirming moves nothing.
  const before = JSON.stringify(lookAlikesOf(writing));
  writing.evaluate('confirmPlaceholder("客户", "N/A"); confirmPlaceholder("客户", "n/a")');
  assert.equal(JSON.stringify(lookAlikesOf(writing)), before);
  writing.evaluate('undoPlaceholderConfirmation("客户", "N/A")');
  assert.equal(JSON.stringify(lookAlikesOf(writing)), before);
  // The roster's date column as a category holds nothing to report.
  const roster = await importStage3Fixture("花名册.xlsx");
  setFieldType(roster, "入职日期", "category");
  assert.deepEqual(lookAlikesOf(roster), []);
});

test("v2.9.0 state: a later commit and resetAnalysisState() replace the look-alike values and scans", async () => {
  const scope = await importStage3Fixture("写法不一致.csv");
  assert.equal(lookAlikesOf(scope).length, 3);
  // The same field name with other cells, without beginImport().
  scope.evaluate('commitTabularData(["区域"], [{ 区域: "甲" }, { 区域: "乙" }], "later.csv", state.activeImportId)');
  assert.deepEqual(lookAlikesOf(scope), []);
  const cleared = await importStage3Fixture("写法不一致.csv");
  cleared.evaluate("resetAnalysisState()");
  assert.deepEqual(lookAlikesOf(cleared), []);
  assert.equal(cleared.evaluate("state.formatScans.writings.size"), 0);
});

const problemNotes = (scope, kind, field) => JSON.parse(scope.evaluate(`JSON.stringify(listProblemRows({ kind: ${JSON.stringify(kind)}, field: ${JSON.stringify(field)} }).items.map((item) => [item.line, item.note]))`));
const problemLinesOf = (scope, kind, field) => problemNotes(scope, kind, field).map(([line]) => line);

test("v2.9.0 problem rows: the rows of the date styles", async () => {
  const roster = await importStage3Fixture("花名册.xlsx");
  assert.deepEqual(problemNotes(roster, "dateStyle", "入职日期"), [
    [6, "文本 年/月/日（主写法是 Excel 日期单元格）"],
    [10, "文本 中文年月日（主写法是 Excel 日期单元格）"],
    [15, "文本 年.月.日（主写法是 Excel 日期单元格）"],
    [19, "文本 年/月/日（主写法是 Excel 日期单元格）"],
    [22, "文本 中文年月日（主写法是 Excel 日期单元格）"],
    [28, "文本 年.月.日（主写法是 Excel 日期单元格）"]
  ]);
  const writing = await importStage3Fixture("写法不一致.csv");
  assert.deepEqual(problemLinesOf(writing, "dateStyle", "下单日期"), knownIssues["写法不一致.csv"].dateStyles[0].otherLines);
});

test("v2.9.0 problem rows: look-alike rows go group by group, then the others", async () => {
  const writing = await importStage3Fixture("写法不一致.csv");
  assert.deepEqual(problemLinesOf(writing, "lookAlike", "商品编号"), [2, 3, 5]);
  assert.deepEqual(problemLinesOf(writing, "lookAlike", "区域"), [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 18, 19, 20, 21, 22, 23, 24]);
  assert.deepEqual(problemLinesOf(writing, "lookAlike", "客户"), [2, 4, 6, 8, 9, 15, 17, 19, 21, 23]);
  assert.deepEqual(problemNotes(writing, "lookAlike", "区域").filter(([line]) => [5, 6, 7, 18, 24].includes(line)), [
    [5, "第 1 组（不可见字符），首尾空格：「华东␣」"],
    [6, "第 1 组（不可见字符），首尾空格：「[全角空格]华东」"],
    [7, "第 1 组（不可见字符）：「华东[零宽字符]」"],
    [18, "第 2 组（全角半角或兼容写法、大小写）：「HD」"],
    [24, "第 3 组（中间空格多少）：「上海␣␣浦东」"]
  ]);
  assert.deepEqual(problemNotes(writing, "lookAlike", "商品编号")[2], [5, "首尾空格：「SKU-002␣」"]);
  // Cells of nothing but invisible characters, alone or with edge whitespace.
  const invisible = await importCsvFile("备注,序\r\n甲,1\r\nAB,2\r\n\u200B,3\r\nab,4\r\n\u00AD ,5\r\n乙 ,6\r\n丙,7\r\n", "z.csv");
  assert.deepEqual(problemNotes(invisible, "lookAlike", "备注"), [
    [3, "第 1 组（大小写）：「AB」"],
    [5, "第 1 组（大小写）：「ab」"],
    [4, "只有不可见字符：「[零宽字符]」"],
    [6, "只有不可见字符，首尾空格：「[软连字符]␣」"],
    [7, "首尾空格：「乙␣」"]
  ]);
  assert.equal(invisible.evaluate("state.lookAlikes[0].rowCount"), 5);
  // A field with only such cells has a category of its own.
  const alone = await importCsvFile("备注,序\r\n\u200B,1\r\n甲,2\r\n", "a.csv");
  assert.deepEqual(problemNotes(alone, "lookAlike", "备注"), [[2, "只有不可见字符：「[零宽字符]」"]]);
  assert.deepEqual(JSON.parse(alone.evaluate('JSON.stringify(state.problemCategories.filter((item) => item.kind === "lookAlike"))')), [{ kind: "lookAlike", field: "备注", count: 1, unit: "row" }]);
});

test("v2.9.0 problem rows: order of the categories, labels, buttons, highlighted cells and file names", async () => {
  const writing = await importStage3Fixture("写法不一致.csv");
  const kinds = JSON.parse(writing.evaluate("JSON.stringify(state.problemCategories.map((item) => item.kind))"));
  assert.deepEqual(kinds.filter((kind, index) => kind !== kinds[index - 1]), ["missing", "conversion", "dateStyle", "idDuplicate", "lookAlike", "placeholder"]);
  const buttons = problemButtons(writing).filter((button) => ["dateStyleDetails", "lookAlikeDetails"].includes(button.container));
  assert.deepEqual(buttons.map((button) => button.label), ["查看「下单日期」不是主写法的行", "查看「商品编号」写法不一样的行", "查看「区域」写法不一样的行", "查看「客户」写法不一样的行"]);
  assert.ok(buttons.every((button) => button.dataAttributes.join() === "data-problem-index"));
  clickProblemButton(writing, buttons[2].index);
  assert.deepEqual(sameCategory(openCategory(writing)), sameCategory(categoryAt(writing, buttons[2].index)));
  assert.equal(openCategory(writing).field, "区域");
  assert.equal(writing.elements.get("problemRowsSummary").textContent, "共 18 行，第 1 / 1 页");
  assert.equal((writing.elements.get("problemRowsTable").innerHTML.match(/<td class="cell-problem">/g) || []).length, 18);
  assert.match(writing.elements.get("problemRowsTable").innerHTML, /<td class="cell-problem">华东<\/td>/);
  clickProblemButton(writing, buttons[0].index);
  assert.equal(writing.elements.get("problemRowsSummary").textContent, "共 6 个值，第 1 / 1 页");
  assert.equal((writing.elements.get("problemRowsTable").innerHTML.match(/<td class="cell-problem">/g) || []).length, 6);
  assert.equal(writing.evaluate('formatProblemCategoryLabel(state.problemCategories.find((item) => item.kind === "dateStyle"))'), "日期写法 · 下单日期（6 个值）");
  assert.equal(writing.evaluate('formatProblemCategoryLabel(state.problemCategories.find((item) => item.kind === "lookAlike" && item.field === "区域"))'), "写法不一样 · 区域（18 行）");
  assert.equal(writing.evaluate('buildProblemRowsFileName({ kind: "dateStyle", field: "下单日期" }, new Date(2026, 10, 5))'), "写法不一致_问题行_日期写法_下单日期_20261105.csv");
  assert.equal(writing.evaluate('buildProblemRowsFileName({ kind: "lookAlike", field: "区域" }, new Date(2026, 10, 5))'), "写法不一致_问题行_写法不一样_区域_20261105.csv");
  // Exported cells keep the original value, whitespace included.
  const content = writing.evaluate('buildProblemRowsCsvContent(listProblemRows({ kind: "lookAlike", field: "区域" }), state.fields)');
  assert.ok(content.includes(',"华东 ",'));
  assert.ok(content.includes(",\u3000华东,"));
  assert.ok(content.includes(",华东\u200B,"));
});

test("v2.9.0 the viewer follows the field configuration, and constant columns have no category", async () => {
  const roster = await importStage3Fixture("花名册.xlsx");
  const [button] = problemButtons(roster).filter((item) => item.container === "dateStyleDetails");
  clickProblemButton(roster, button.index);
  assert.equal(viewerState(roster).hidden, false);
  setFieldType(roster, "入职日期", "category");
  assert.equal(viewerState(roster).hidden, true);
  assert.equal(viewerState(roster).message, "这一类已经没有问题行。");
  const errors = await importStage3Fixture("公式错误.xlsx");
  assert.equal(errors.evaluate('state.problemCategories.some((item) => item.kind === "constantField")'), false);
  assert.doesNotMatch(errors.elements.get("constantFieldDetails").innerHTML, /<button/);
  assert.match(errors.elements.get("constantFieldDetails").innerHTML, /另有 3 行是空值/);
});

test("v2.9.0 the three sections render without a numeric field, escape data and hide when empty", () => {
  const scope = createScriptContext();
  const name = "<i>名</i>";
  const dates = "<s>日期</s>";
  commitRows(scope, [name, dates, "固定"], [
    { [name]: "<b>x</b>", [dates]: "2024-01-01", 固定: "<u>同</u>" },
    { [name]: "<B>X</B>", [dates]: "2024/1/2", 固定: "<u>同</u>" },
    { [name]: "y ", [dates]: "2024-01-03", 固定: "<u>同</u>" },
    { [name]: "z", [dates]: "2024-01-04", 固定: "<u>同</u>" },
    { [name]: "z", [dates]: "2024-01-05", 固定: "<u>同</u>" }
  ], "text-only.csv");
  assert.match(scope.elements.get("dateStyleDetails").innerHTML, /「&lt;s&gt;日期&lt;\/s&gt;」/);
  assert.match(scope.elements.get("outlierTable").innerHTML, /未识别到数值字段/);
  for (const [section, details] of [["dateStyleSection", "dateStyleDetails"], ["constantFieldSection", "constantFieldDetails"], ["lookAlikeSection", "lookAlikeDetails"]]) {
    assert.equal(scope.elements.get(section).classList.contains("hidden"), false, section);
    assert.doesNotMatch(scope.elements.get(details).innerHTML, /<b>|<i>|<u>|<B>|<s>/, details);
  }
  assert.match(scope.elements.get("lookAlikeDetails").innerHTML, /&lt;b&gt;x&lt;\/b&gt;/);
  assert.match(scope.elements.get("lookAlikeDetails").innerHTML, /data-problem-index="\d+"/);
  assert.doesNotMatch(scope.elements.get("lookAlikeDetails").innerHTML + scope.elements.get("dateStyleDetails").innerHTML, /data-placeholder-index/);
  const clean = createScriptContext();
  commitRows(clean, ["城市"], [{ 城市: "上海" }, { 城市: "北京" }], "clean.csv");
  for (const section of ["dateStyleSection", "constantFieldSection", "lookAlikeSection"]) assert.equal(clean.elements.get(section).classList.contains("hidden"), true, section);
  // Rendered again from the current state, and emptied by a reset.
  const scopeHtml = scope.elements.get("dateStyleDetails").innerHTML;
  assert.notEqual(scopeHtml, "");
  scope.evaluate("resetAnalysisState()");
  assert.equal(scope.elements.get("dateStyleDetails").innerHTML, "");
  assert.equal(scope.elements.get("dateStyleSection").classList.contains("hidden"), true);
  assert.equal(scope.elements.get("lookAlikeDetails").innerHTML, "");
  assert.equal(scope.elements.get("constantFieldDetails").innerHTML, "");
});

test("v2.9.0 the sections are static, in place and change neither the pills nor the status", async () => {
  const html = indexHtmlSource();
  const positions = ['id="numericLookingSection"', 'id="dateStyleSection"', 'id="constantFieldSection"', 'id="lookAlikeSection"', 'id="placeholderDetails"'].map((id) => html.indexOf(id));
  assert.ok(positions.every((position, index) => position >= 0 && (index === 0 || position > positions[index - 1])), positions.join());
  const writing = await importStage3Fixture("写法不一致.csv");
  assert.equal((writing.elements.get("qualitySummary").innerHTML.match(/class="quality-pill/g) || []).length, 6);
  assert.equal(writing.elements.get("statusPanel").textContent, "写法不一致.csv 已成功载入，分析结果已生成。");
});

test("v2.9.0 reports: the notes are the page's sentences", () => {
  const html = indexHtmlSource();
  const { evaluate } = createScriptContext();
  for (const name of ["DATE_STYLE_NOTE", "CONSTANT_FIELD_NOTE", "LOOK_ALIKE_NOTE"]) assert.ok(html.includes(`<p class="section-note">${evaluate(`${name}[0]`)}</p>`), name);
  assert.match(evaluate("LOOK_ALIKE_NOTE[0]"), /只有不可见字符：格子看起来是空的/);
});

test("v2.9.0 reports: three sections with content, two summary counts, constant columns not counted", async () => {
  const roster = reportOf(await importStage3Fixture("花名册.xlsx"));
  assert.ok(roster.html.includes("<h3>日期写法混用</h3>\n    <p>这些写法本工具都能读成同一个日期，但在 Excel 里文本日期不能按日期排序、筛选，导入别的系统时可能读错。</p>"));
  assert.ok(roster.html.includes("<li>「入职日期」：主写法是 Excel 日期单元格 24 个（如 2016-04-01）；另有 6 行是别的写法：文本 年/月/日 2 个（如 2021/3/5）、文本 中文年月日 2 个（如 2021年3月5日）、文本 年.月.日 2 个（如 2021.03.05）。</li>"));
  assert.ok(!roster.html.includes("整列只有一个值"));

  const errorsScope = await importStage3Fixture("公式错误.xlsx");
  const errors = reportOf(errorsScope);
  assert.ok(errors.html.includes("<li>「备注」：只有一种取值「正常」（6 行），另有 3 行是空值。</li>"));
  // The note right under the heading, then the list.
  assert.match(errors.html, new RegExp(`<h3>整列只有一个值<\\/h3>\\s*<p>${regExpText(errorsScope.evaluate("CONSTANT_FIELD_NOTE[0]"))}<\\/p>\\s*<ul[^>]*>\\s*<li>「备注」`));
  assert.ok(!errors.html.includes("整列只有一个值：<strong>"));

  const writing = reportOf(await importStage3Fixture("写法不一致.csv"));
  assert.ok(writing.html.includes("；日期写法混用：<strong>1</strong> 个字段；同一个值，写法不一样：<strong>3</strong> 个字段。</p>"));
  // The section shows whitespace and invisible characters by name; Top 10 and the sample values
  // still print cells as they are.
  const section = writing.html.slice(writing.html.indexOf("<h3>同一个值，写法不一样</h3>"), writing.html.indexOf("<h3>像是空值的写法</h3>"));
  assert.ok(section.includes("「华东[零宽字符]」1 个"));
  assert.ok(section.includes("（如「SKU-002␣」）"));
  assert.ok(!section.includes("\u200B"));
  // Order: after the numeric-looking section's place, before the placeholders, in both outputs.
  const order = (text, titles) => titles.map((title) => text.indexOf(title));
  const htmlOrder = order(writing.html, ["<h3>各字段空值情况</h3>", "<h3>日期写法混用</h3>", "<h3>整列只有一个值</h3>", "<h3>同一个值，写法不一样</h3>", "<h3>像是空值的写法</h3>"]);
  assert.ok(htmlOrder.every((position, index) => position >= 0 && (index === 0 || position > htmlOrder[index - 1])), htmlOrder.join());

  const clean = reportOf(await importStage3Fixture("订单表.csv"));
  assert.deepEqual([clean.report.quality.dateStyles, clean.report.quality.constantFields, clean.report.quality.lookAlikes], [[], [], []]);
  for (const title of ["日期写法混用", "整列只有一个值", "同一个值，写法不一样"]) assert.ok(!clean.html.includes(title), title);
});

test("v2.9.0 reports: cells of nothing but invisible characters, escaped field names and values", async () => {
  const field = "<b>备注|*</b>";
  const scope = await importCsvFile(`"${field}",序\r\n\u200B,1\r\n<i>x</i>,2\r\n<I>X</I>,3\r\n`, "r.csv");
  const { report, html } = reportOf(scope);
  assert.equal(report.quality.lookAlikes[0].invisibleOnlyCount, 1);
  assert.ok(html.includes("只有不可见字符 1 格"));
  assert.ok(html.includes("「&lt;b&gt;备注|*&lt;/b&gt;」共 3 行"));
  assert.ok(html.includes("「&lt;i&gt;x&lt;/i&gt;」1 个"));
  assert.ok(html.includes("<h3>同一个值，写法不一样</h3>"));
  assert.doesNotMatch(html.slice(html.indexOf("<h3>同一个值，写法不一样</h3>")), /<i>x|<b>备注/);
});

test("v2.9.0 reports: old snapshots leave the sections out, and the snapshot shares nothing with state", async () => {
  const store = await importStage3Fixture("写法不一致.csv");
  const old = reportOf(store).report;
  delete old.quality.dateStyles;
  delete old.quality.constantFields;
  delete old.quality.lookAlikes;
  put(store.context, "__old", old);
  // null itself: JSON.stringify would write undefined in an array as null too.
  assert.equal(store.evaluate("(() => { const quality = normalizeReportSnapshot(__old).quality; return quality.dateStyles === null && quality.constantFields === null && quality.lookAlikes === null; })()"), true);
  const oldOutput = store.evaluate("buildHtmlReport(__old, [])");
  for (const title of ["日期写法混用", "整列只有一个值", "同一个值，写法不一样"]) assert.ok(!oldOutput.includes(title), title);
  store.evaluate('__snapshot = buildReportData(new Date("2026-10-25T00:00:00.000Z"))');
  const shared = JSON.parse(store.evaluate(`JSON.stringify([
    __snapshot.quality.dateStyles[0] === state.dateStyles[0],
    __snapshot.quality.dateStyles[0].styles === state.dateStyles[0].styles,
    __snapshot.quality.dateStyles[0].styles[0] === state.dateStyles[0].styles[0],
    __snapshot.quality.constantFields[0] === state.constantFields[0],
    __snapshot.quality.lookAlikes[0] === state.lookAlikes[0],
    __snapshot.quality.lookAlikes[0].groups === state.lookAlikes[0].groups,
    __snapshot.quality.lookAlikes[0].groups[0] === state.lookAlikes[0].groups[0],
    __snapshot.quality.lookAlikes[0].groups[0].reasons === state.lookAlikes[0].groups[0].reasons,
    __snapshot.quality.lookAlikes[0].groups[0].variants === state.lookAlikes[0].groups[0].variants,
    __snapshot.quality.lookAlikes[0].groups[0].variants[0] === state.lookAlikes[0].groups[0].variants[0]
  ])`));
  assert.deepEqual(shared, shared.map(() => false));
  // Filters change nothing in these sections.
  const before = JSON.stringify(reportOf(store).report.quality.lookAlikes);
  store.evaluate('state.filters.category = { field: "状态", values: ["正常"] }; state.filteredRows = state.rows.slice(0, 3); refreshFilteredAnalysis()');
  assert.equal(JSON.stringify(reportOf(store).report.quality.lookAlikes), before);
});

test("v2.9.0 Excel: rows with the same date fields share one frozen record without a prototype", async () => {
  const cells = {
    A1: textCell("名"), B1: textCell("入职"), C1: textCell("离职"),
    A2: textCell("甲"), B2: dateCell(45418), C2: textCell("在职"),
    A3: textCell("乙"), B3: dateCell(45419), C3: textCell("在职"),
    A4: textCell("丙"), B4: dateCell(45420), C4: dateCell(45500),
    A5: textCell("丁"), B5: textCell("待定"), C5: dateCell(45501),
    A6: textCell("戊"), B6: dateCell(45421), C6: dateCell(45502)
  };
  const scope = await importCellSheet(cells, "A1:C6");
  const read = JSON.parse(scope.evaluate(`JSON.stringify((() => {
    const records = state.rows.map((row) => state.sourceDateCells.get(row) || null);
    return {
      fields: records.map((record) => record && Object.keys(record)),
      same: [records[0] === records[1], records[2] === records[4], records[0] === records[2], records[2] === records[3]],
      frozen: records.filter(Boolean).every((record) => Object.isFrozen(record)),
      prototypes: records.filter(Boolean).every((record) => Object.getPrototypeOf(record) === null)
    };
  })())`));
  assert.deepEqual(read.fields, [["入职"], ["入职"], ["入职", "离职"], ["离职"], ["入职", "离职"]]);
  assert.deepEqual(read.same, [true, true, false, false]);
  assert.equal(read.frozen, true);
  assert.equal(read.prototypes, true);
});

test("v2.9.0 guard: text columns named like Object members never count as date cells", async () => {
  const cells = {
    A1: textCell("constructor"), B1: textCell("toString"), C1: textCell("hasOwnProperty"), D1: textCell("日期"),
    A2: textCell("甲"), B2: textCell("乙"), C2: textCell("丙"), D2: dateCell(45418),
    A3: textCell("丁"), B3: textCell("戊"), C3: textCell("己"), D3: dateCell(45419)
  };
  const scope = await importCellSheet(cells, "A1:D3");
  const read = JSON.parse(scope.evaluate(`JSON.stringify(state.rows.map((row) => {
    const record = state.sourceDateCells.get(row);
    return ["constructor", "toString", "hasOwnProperty", "日期"].map((field) => Boolean(record && record[field]));
  }))`));
  assert.deepEqual(read, [[false, false, false, true], [false, false, false, true]]);
});

// v2.10.0: the one-cell rules on every imported cell of the fixtures and examples.
test("v2.10.0 guard: the one-cell rules never need the safety net on the fixtures and examples", async () => {
  const files = [...STAGE3_IMPORT_FILES.map((name) => ["stage3", name]), ...EXAMPLE_CSV_FILES.map((name) => ["example", name])];
  let total = 0;
  for (const [kind, name] of files) {
    const scope = kind === "stage3" ? await importStage3Fixture(name) : await importCsvFile(readExampleCsv(name), name);
    for (const step of ["import", "confirmed"]) {
      if (step === "confirmed") scope.evaluate("confirmAllCommonPlaceholders()");
      const result = JSON.parse(scope.evaluate(`JSON.stringify((() => {
        let checked = 0;
        const bad = [];
        state.profiles.forEach((profile) => {
          if (profile.typeKey !== "numeric" && profile.typeKey !== "date") return;
          state.rows.forEach((row) => {
            const value = row[profile.field];
            const cell = profile.typeKey === "numeric"
              ? standardizeNumberCell(value, profile.numberStrategy, profile.confirmedKeys, "")
              : standardizeDateCell(value, profile.dateStrategy, profile.confirmedKeys);
            checked += 1;
            if (cell.action === "keptUnsafe") bad.push([profile.field, value]);
            if (cell.action === "number") {
              const after = toPlainNumber(cell.value);
              const before = toNumber(value, profile.numberStrategy);
              if (!(Object.is(after, before) || (after === 0 && before === 0))) bad.push([profile.field, value, cell.value]);
            }
          });
        });
        return { checked, bad };
      })())`));
      assert.deepEqual(result.bad, [], `${name} ${step}`);
      total += result.checked;
    }
  }
  assert.ok(total > 5000, `cells checked: ${total}`);
});

// v2.10.0: the standardized check. Only computed by 检查标准化改动; nothing else changes.
const STANDARDIZE_CSV = [
  "编号,日期,金额,区域,年份,备注",
  "E01,2024/5/6,100元,华东,2019, a ",
  "E02,2024-05-07,\"2,200元\", 华南,2020,b",
  "E03,2024-05-08 10:30,-50元,华北 ,2021年,c",
  "E04,待定,N/A,华东, 2022,N/A",
  "E05,2024年5月9日, 300 ,华南,2023,e",
  "E06,2024-05-10,400,N/A,—,f",
  "E07,06/06/2024,500元,华东,2019,g",
  "E08,2024/5/11,$12,华南,2020,=1+1",
  "E09,2024/5/12,600,华北,2021,i",
  "E10,2024/5/13,700,华东,2022,j",
  "E11,2024/5/14,800,华南,2023,k",
  "E12,2024/5/15,900,华北,2024,l"
].join("\r\n");

function standardizedCounts(scope) {
  const result = JSON.parse(scope.evaluate("JSON.stringify(state.standardizedExport)"));
  return result && Object.fromEntries(result.counts.map((count) => [count.field, [count.kind, count.date, count.number, count.placeholder, count.trim, count.kept]]));
}

function standardizedActions(scope) {
  return JSON.parse(scope.evaluate("JSON.stringify(state.standardizedExport.cells.actions.map((code) => STANDARDIZE_ACTIONS[code]))"));
}

function standardizedCellRows(scope) {
  return JSON.parse(scope.evaluate("JSON.stringify(buildStandardizedCellRows(state.standardizedExport, state.rows, state.fields))"));
}

function csvFileFor(text, name) {
  const bytes = Buffer.from(text, "utf8");
  return { name, size: bytes.length, arrayBuffer: async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.length) };
}

test("v2.10.0 the constructed table: types, counts per field and kinds, a year field as text", async () => {
  const scope = await importCsvFile(STANDARDIZE_CSV, "标准化.csv");
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(state.profiles.map((p) => [p.field, p.typeKey, p.dateStrategy.granularity, p.numberStrategy.unit]))")), [
    ["编号", "id", "day", ""], ["日期", "date", "day", ""], ["金额", "numeric", "day", "元"], ["区域", "category", "day", ""], ["年份", "date", "year", ""], ["备注", "id", "day", ""]
  ]);
  assert.equal(scope.evaluate("state.standardizedExport"), null);
  scope.evaluate("checkStandardizedExport()");
  assert.deepEqual(standardizedCounts(scope), {
    编号: ["text", 0, 0, 0, 0, 0],
    日期: ["date", 8, 0, 0, 0, 2],
    金额: ["number", 0, 5, 0, 1, 1],
    区域: ["text", 0, 0, 0, 2, 0],
    年份: ["year", 0, 0, 0, 1, 0],
    备注: ["text", 0, 0, 0, 1, 0]
  });
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(state.standardizedExport.totals)")), { date: 8, number: 5, placeholder: 0, trim: 5, kept: 3 });
  // A year field: the year texts are not rewritten as dates, only trimmed.
  const rows = standardizedCellRows(scope);
  assert.deepEqual(rows.map((row) => row[4]), ["2019", "2020", "2021年", "2022", "2023", "—", "2019", "2020", "2021", "2022", "2023", "2024"]);
  assert.deepEqual(rows[2], ["E03", "2024-05-08 10:30", "-50", "华北", "2021年", "c"]);
  assert.deepEqual(standardizedActions(scope).slice(0, 4), ["date", "number", "trim", "number"]);
});

test("v2.10.0 confirmed placeholders become empty in every field but an ignored one", async () => {
  const scope = await importCsvFile(STANDARDIZE_CSV, "标准化.csv");
  scope.evaluate('updatePlaceholderConfirmations({ 金额: ["N/A"], 区域: ["N/A"], 年份: ["—"], 备注: ["N/A"] }, "x")');
  scope.evaluate('state.fieldTypeDrafts["备注"] = "ignore"; applyFieldConfiguration()');
  scope.evaluate("checkStandardizedExport()");
  assert.deepEqual(standardizedCounts(scope), {
    编号: ["text", 0, 0, 0, 0, 0],
    日期: ["date", 8, 0, 0, 0, 2],
    金额: ["number", 0, 5, 1, 1, 0],
    区域: ["text", 0, 0, 1, 2, 0],
    年份: ["year", 0, 0, 1, 1, 0],
    备注: ["ignored", 0, 0, 0, 0, 0]
  });
  const rows = standardizedCellRows(scope);
  assert.deepEqual(rows[0], ["E01", "2024-05-06", "100", "华东", "2019", " a "]);
  assert.deepEqual(rows[3], ["E04", "待定", "", "华东", "2022", "N/A"]);
  assert.deepEqual(rows[5], ["E06", "2024-05-10", "400", "", "", "f"]);
  assert.deepEqual(rows[7], ["E08", "2024-05-11", "12", "华南", "2020", "=1+1"]);
});

test("v2.10.0 currency: one sign is dropped, two or more keep every cell with a sign", async () => {
  const column = (values) => ["价格,门店", ...values.map((value, index) => `"${value}",S${index}`)].join("\r\n");
  const single = await importCsvFile(column(["$12", "13", "$1,400", "15", "16"]), "single.csv");
  const yen = await importCsvFile(column(["¥12", "￥13", "14", "15", "16"]), "yen.csv");
  const mixed = await importCsvFile(column(["$12", "￥13", "1,400", "15", "16"]), "mixed.csv");
  // D15: a percent cell's sign counts too, though the cell itself is never rewritten.
  const percent = await importCsvFile(column(["$50%", "￥12", "13", "14", "15"]), "percent.csv");
  [single, yen, mixed, percent].forEach((scope) => {
    assert.equal(scope.evaluate('getFieldProfile("价格").typeKey'), "numeric");
    scope.evaluate("checkStandardizedExport()");
  });
  assert.deepEqual(standardizedActions(single), ["number", "number"]);
  assert.deepEqual(standardizedActions(yen), ["number", "number"]);
  assert.deepEqual(standardizedActions(mixed), ["keptCurrencyMix", "keptCurrencyMix", "number"]);
  assert.deepEqual(standardizedActions(percent), ["keptPercent", "keptCurrencyMix"]);
  assert.deepEqual(JSON.parse(single.evaluate("JSON.stringify(state.standardizedExport.counts.map((c) => c.currencyMode))")), ["single", ""]);
  assert.deepEqual(JSON.parse(mixed.evaluate("JSON.stringify(state.standardizedExport.counts.map((c) => c.currencyMode))")), ["mixed", ""]);
});

test("v2.10.0 guard: the check never runs on import, filter, field configuration or confirmations", async () => {
  const scope = await importCsvFile(STANDARDIZE_CSV, "标准化.csv");
  scope.evaluate("var __standardizeCalls = 0; const __buildStandardizedExport = buildStandardizedExport; buildStandardizedExport = (...args) => { __standardizeCalls += 1; return __buildStandardizedExport(...args); }");
  const calls = () => scope.evaluate("__standardizeCalls");
  put(scope.context, "__csvAgain", csvFileFor(STANDARDIZE_CSV, "again.csv"));
  await scope.evaluate("parseDataFile(__csvAgain)");
  applyFilterState(scope, { category: { field: "区域", values: ["华东"] } });
  scope.evaluate('state.fieldTypeDrafts["备注"] = "ignore"; applyFieldConfiguration()');
  scope.evaluate('updatePlaceholderConfirmations({ 金额: ["N/A"] }, "x")');
  scope.evaluate("restoreAutomaticFieldTypes()");
  assert.equal(calls(), 0);
  scope.evaluate("checkStandardizedExport()");
  assert.equal(calls(), 1);
});

test("v2.10.0 a filter keeps the check; import, commit, field configuration, confirmations and reset drop it", async () => {
  const scope = await importCsvFile(STANDARDIZE_CSV, "标准化.csv");
  const check = () => {
    scope.evaluate("checkStandardizedExport()");
    assert.notEqual(scope.evaluate("state.standardizedExport"), null);
  };
  const dropped = (label) => assert.equal(scope.evaluate("state.standardizedExport"), null, label);
  check();
  scope.evaluate("var __kept = state.standardizedExport");
  applyFilterState(scope, { category: { field: "区域", values: ["华东"] } });
  assert.equal(scope.evaluate("state.standardizedExport === __kept"), true);
  scope.evaluate("clearAllFilters()");
  assert.equal(scope.evaluate("state.standardizedExport === __kept"), true);
  scope.evaluate('state.fieldTypeDrafts["备注"] = "ignore"; applyFieldConfiguration()');
  dropped("field configuration");
  check();
  scope.evaluate("restoreAutomaticFieldTypes()");
  dropped("restore");
  check();
  scope.evaluate('confirmPlaceholder("金额", "N/A")');
  dropped("confirm");
  check();
  scope.evaluate('undoPlaceholderConfirmation("金额", "N/A")');
  dropped("undo");
  check();
  put(scope.context, "__csvAgain", csvFileFor(STANDARDIZE_CSV, "again.csv"));
  await scope.evaluate("parseDataFile(__csvAgain)");
  dropped("import");
  check();
  // A commit without beginImport() (which resets everything first): the commit itself drops it.
  scope.evaluate('commitTabularData(state.fields, state.rows.map((row) => ({ ...row })), "again.csv", state.activeImportId)');
  dropped("commit");
  check();
  scope.evaluate("resetAnalysisState()");
  dropped("reset");
});

const DMY_EVIDENCE_REWRITTEN = "日期,门店\r\n13/05/2024,S1\r\n06/05/2024 10:30,S2\r\n07/05/2024 10:30,S3\r\n08/05/2024 10:30,S4\r\n09/05/2024 10:30,S5\r\n";

test("v2.10.0 D8: a field whose day-month order rests on rewritten cells gets a hint, and the hint is right", async () => {
  const scope = await importCsvFile(DMY_EVIDENCE_REWRITTEN, "dmy.csv");
  assert.equal(scope.evaluate('getFieldProfile("日期").dateStrategy.order'), "dmy");
  scope.evaluate("checkStandardizedExport()");
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(state.standardizedExport.counts.map((c) => [c.field, c.orderHint]))")), [["日期", true], ["门店", false]]);
  // The hint is true: the exported file imported again has lost its order and its date type.
  const text = scope.evaluate("buildCsvTableContent([...state.fields], buildStandardizedCellRows(state.standardizedExport, state.rows, state.fields))");
  const again = await importCsvFile(text, "dmy_标准化.csv");
  assert.equal(again.evaluate('getFieldProfile("日期").dateStrategy.order'), "none");
  assert.equal(again.evaluate('getFieldProfile("日期").typeKey'), "category");
  // The evidence cell written with a time is kept, so the order survives: no hint.
  const kept = await importCsvFile(DMY_EVIDENCE_REWRITTEN.replace("13/05/2024,S1", "13/05/2024 09:00,S1"), "dmy-kept.csv");
  assert.equal(kept.evaluate('getFieldProfile("日期").dateStrategy.order'), "dmy");
  kept.evaluate("checkStandardizedExport()");
  assert.deepEqual(JSON.parse(kept.evaluate("JSON.stringify(state.standardizedExport.counts.map((c) => c.orderHint))")), [false, false]);
  // A plain dmy or mdy column rewritten in full loses its order evidence but every exported cell
  // still reads: no hint, and a second import keeps the date type.
  for (const [label, text] of [
    ["dmy", "日期,门店\r\n13/05/2024,S1\r\n06/05/2024,S2\r\n07/05/2024,S3\r\n08/05/2024,S4\r\n09/05/2024,S5\r\n"],
    ["mdy", "日期,门店\r\n05/13/2024,S1\r\n05/06/2024,S2\r\n05/07/2024,S3\r\n05/08/2024,S4\r\n05/09/2024,S5\r\n"]
  ]) {
    const plain = await importCsvFile(text, `${label}.csv`);
    assert.equal(plain.evaluate('getFieldProfile("日期").dateStrategy.order'), label);
    plain.evaluate("checkStandardizedExport()");
    assert.deepEqual(JSON.parse(plain.evaluate("JSON.stringify(state.standardizedExport.counts.map((c) => c.orderHint))")), [false, false], label);
    const exported = plain.evaluate("buildCsvTableContent([...state.fields], buildStandardizedCellRows(state.standardizedExport, state.rows, state.fields))");
    const again = await importCsvFile(exported, `${label}_标准化.csv`);
    assert.equal(again.evaluate('getFieldProfile("日期").typeKey'), "date", label);
  }
  // A field without a read order is not looked at.
  const none = await importCsvFile("日期,门店\r\n2024/5/6,S1\r\n2024/5/7,S2\r\n2024/5/8,S3\r\n", "none.csv");
  none.evaluate("checkStandardizedExport()");
  assert.deepEqual(JSON.parse(none.evaluate("JSON.stringify(state.standardizedExport.counts.map((c) => c.orderHint))")), [false, false]);
});

test("v2.10.0 all fixtures and examples: rows and exports unchanged, counts = list, round trip keeps every type", async () => {
  const files = [...STAGE3_IMPORT_FILES.map((name) => ["stage3", name]), ...EXAMPLE_CSV_FILES.map((name) => ["example", name])];
  const totals = {};
  for (const [kind, name] of files) {
    const scope = kind === "stage3" ? await importStage3Fixture(name) : await importCsvFile(readExampleCsv(name), name);
    const before = scope.evaluate("JSON.stringify(state.rows)");
    const exports = () => scope.evaluate('buildCsvExportContent(getProcessedExportRows("filtered"), state.fields)');
    const existing = exports();
    scope.evaluate("checkStandardizedExport()");
    assert.equal(scope.evaluate("JSON.stringify(state.rows)"), before, name);
    assert.equal(exports(), existing, name);
    assert.equal(scope.evaluate("state.standardizedExport.counts.every((c, f) => c.date + c.number + c.placeholder + c.trim + c.kept === state.standardizedExport.cells.fields.filter((x) => x === f).length)"), true, name);
    assert.equal(scope.evaluate("state.standardizedExport.counts.some((c) => c.orderHint)"), false, name);
    totals[name] = scope.evaluate("state.standardizedExport.cellCount");
    const text = scope.evaluate("buildCsvTableContent([...state.fields], buildStandardizedCellRows(state.standardizedExport, state.rows, state.fields))");
    const again = await importCsvFile(text, `${name}.csv`);
    const types = (target) => target.evaluate("JSON.stringify(state.profiles.map((p) => [p.field, p.typeKey]))");
    assert.equal(types(again), types(scope), name);
  }
  assert.deepEqual(totals, {
    "en-formula-errors.xlsx": 0, "en-issues.csv": 20, "公式错误.xlsx": 0, "写法不一致.csv": 13, "分区小计表.csv": 14, "成绩表-标题无逗号.csv": 3, "成绩表.csv": 3, "花名册.xlsx": 8, "订单表.csv": 36, "门店报表.xlsx": 0,
    "sample-data.csv": 0, "sample-orders-en.csv": 12, "sample-orders.csv": 11, "sample-sales.csv": 0, "sample-students.csv": 0, "sample-survey.csv": 0, "sample-used-products.csv": 0, "sample-user-behavior.csv": 119
  });
});

test("v2.10.0 the check needs data, a finished analysis and no unapplied field draft", async () => {
  const empty = createScriptContext();
  empty.evaluate("checkStandardizedExport()");
  assert.equal(empty.evaluate("state.standardizedExport"), null);
  const scope = await importCsvFile(STANDARDIZE_CSV, "标准化.csv");
  scope.evaluate('state.fieldTypeDrafts["备注"] = "ignore"');
  scope.evaluate("checkStandardizedExport()");
  assert.equal(scope.evaluate("state.standardizedExport"), null);
  // An earlier check stays while a draft waits (D11) and is still there once the draft is undone.
  scope.evaluate('state.fieldTypeDrafts["备注"] = "id"; checkStandardizedExport(); var __kept = state.standardizedExport');
  scope.evaluate('state.fieldTypeDrafts["备注"] = "ignore"; checkStandardizedExport()');
  assert.equal(scope.evaluate("state.standardizedExport === __kept"), true);
  scope.evaluate('state.fieldTypeDrafts["备注"] = "id"');
  assert.equal(scope.evaluate("state.standardizedExport === __kept"), true);
  scope.evaluate("state.analysisCompletedAt = null; state.standardizedExport = null; checkStandardizedExport()");
  assert.equal(scope.evaluate("state.standardizedExport"), null);
});

// v2.10.0: the check table on the page, the gates of its three buttons and the static structure.
function standardizedButtons(scope) {
  return ["checkStandardizedExport", "exportStandardizedCsv", "exportStandardizedChanges"].map((id) => scope.elements.get(id).disabled);
}

// The check table as rows of cell texts (tags removed, entities kept).
function standardizedTableRows(scope) {
  return [...scope.elements.get("standardizedExportTable").innerHTML.matchAll(/<tr[^>]*>(.*?)<\/tr>/g)]
    .map((row) => [...row[1].matchAll(/<t[hd][^>]*>(.*?)<\/t[hd]>/g)].map((cell) => cell[1].replace(/<[^>]+>/g, "|").replace(/\|+/g, "|").replace(/^\||\|$/g, "")));
}

const standardizedPanelHidden = (scope) => scope.elements.get("standardizedExportPanel").classList.contains("hidden");

test("v2.10.0 the check table: changed fields, kept samples, notes and the total", async () => {
  const scope = await importCsvFile(STANDARDIZE_CSV, "标准化.csv");
  assert.equal(standardizedPanelHidden(scope), true);
  scope.evaluate("checkStandardizedExport()");
  assert.equal(standardizedPanelHidden(scope), false);
  assert.equal(scope.elements.get("standardizedExportSummary").textContent, "共改写 18 格，保持原值 3 格。其余 1 个字段没有改动。");
  assert.deepEqual(standardizedTableRows(scope), [
    ["字段", "类型", "日期改写", "数字改写", "清空空值写法", "去首尾空格", "保持原值"],
    ["日期", "日期", "8", "—", "—", "—", "2（如「2024-05-08 10:30」、「待定」）"],
    ["金额|单位「元」不写进导出文件，表头仍是「金额」", "数值", "—", "5", "—", "1", "1（如「N/A」）"],
    ["区域", "分类", "—", "—", "—", "2", "—"],
    ["年份|按年份读取，不改写日期", "日期", "—", "—", "—", "1", "—"],
    ["备注", "ID", "—", "—", "—", "1", "—"],
    ["合计", "", "8", "5", "—", "5", "3"]
  ]);
  scope.evaluate('state.fieldTypeDrafts["备注"] = "ignore"; applyFieldConfiguration(); checkStandardizedExport()');
  assert.equal(scope.elements.get("standardizedExportSummary").textContent, "共改写 17 格，保持原值 3 格。其余 1 个字段没有改动。已忽略的 1 个字段原样导出。");
  assert.deepEqual(standardizedTableRows(scope).map((row) => row[0].split("|")[0]), ["字段", "日期", "金额", "区域", "年份", "合计"]);
});

test("v2.10.0 the check table: the D8 hint, whitespace shown, data escaped, nothing to change", async () => {
  const dmy = await importCsvFile(DMY_EVIDENCE_REWRITTEN, "dmy.csv");
  dmy.evaluate("checkStandardizedExport()");
  assert.deepEqual(standardizedTableRows(dmy)[1], ["日期|导出后这一列剩下的日或月在前的写法没有了确定顺序的依据，再次导入本工具时可能读不成日期。", "日期", "1", "—", "—", "—", "4（如「06/05/2024 10:30」、「07/05/2024 10:30」、「08/05/2024 10:30」）"]);
  const escaped = await importCsvFile("<b>名</b>,数量\r\n\" <i>x</i>\",1\r\ny,2\r\nz,3\r\n", "escape.csv");
  escaped.evaluate("checkStandardizedExport()");
  const html = escaped.elements.get("standardizedExportTable").innerHTML;
  assert.ok(html.includes("&lt;b&gt;名&lt;/b&gt;"));
  assert.ok(!html.includes("<b>") && !html.includes("<i>"));
  const kept = await importCsvFile("金额,门店\r\n\" 缺考 \",S1\r\n1,S2\r\n2,S3\r\n3,S4\r\n4,S5\r\n", "kept.csv");
  kept.evaluate("checkStandardizedExport()");
  assert.equal(standardizedTableRows(kept)[1][6], "1（如「␣缺考␣」）");
  const clean = await importCsvFile("名称,数量\r\nA,1\r\nB,2\r\nC,3\r\n", "clean.csv");
  clean.evaluate("checkStandardizedExport()");
  assert.equal(clean.elements.get("standardizedExportSummary").textContent, "没有需要统一写法的单元格，导出的数据与原始数据相同。");
  assert.equal(clean.elements.get("standardizedExportTable").innerHTML, "");
  assert.equal(standardizedPanelHidden(clean), false);
  // D9: no listed cell, so the change list cannot be exported; the data still can.
  assert.deepEqual(standardizedButtons(clean), [false, false, true]);
});

test("v2.10.0 the panel follows the check: a filter keeps it, everything else hides and empties it", async () => {
  const scope = await importCsvFile(STANDARDIZE_CSV, "标准化.csv");
  const shown = (label) => {
    scope.evaluate("checkStandardizedExport()");
    assert.equal(standardizedPanelHidden(scope), false, label);
    assert.notEqual(scope.elements.get("standardizedExportTable").innerHTML, "", label);
    assert.deepEqual(standardizedButtons(scope), [false, false, false], label);
  };
  const hidden = (label) => {
    assert.equal(standardizedPanelHidden(scope), true, label);
    assert.equal(scope.elements.get("standardizedExportTable").innerHTML, "", label);
    assert.equal(scope.elements.get("standardizedExportSummary").textContent, "", label);
    assert.equal(scope.elements.get("standardizedExportStatus").textContent, "", label);
    assert.deepEqual(standardizedButtons(scope).slice(1), [true, true], label);
  };
  shown("first");
  const table = scope.elements.get("standardizedExportTable").innerHTML;
  scope.evaluate('setStandardizedExportStatus("x", "success")');
  applyFilterState(scope, { category: { field: "区域", values: ["华东"] } });
  assert.equal(standardizedPanelHidden(scope), false);
  assert.equal(scope.elements.get("standardizedExportTable").innerHTML, table);
  assert.equal(scope.elements.get("standardizedExportStatus").textContent, "x");
  assert.deepEqual(standardizedButtons(scope), [false, false, false]);
  scope.evaluate('state.fieldTypeDrafts["备注"] = "ignore"; applyFieldConfiguration()');
  hidden("field configuration");
  shown("again");
  scope.evaluate('confirmPlaceholder("金额", "N/A")');
  hidden("confirm");
  shown("again");
  scope.evaluate('commitTabularData(state.fields, state.rows.map((row) => ({ ...row })), "again.csv", state.activeImportId)');
  hidden("commit");
  shown("again");
  scope.evaluate("resetAnalysisState()");
  hidden("reset");
});

test("v2.10.0 gate: no data or an unapplied field draft disables the buttons and says why", async () => {
  const empty = createScriptContext();
  empty.evaluate("updateExportButtons()");
  assert.deepEqual(standardizedButtons(empty), [true, true, true]);
  assert.equal(empty.elements.get("standardizedExportGate").textContent, "请先上传并解析有效的数据文件。");
  const scope = await importCsvFile(STANDARDIZE_CSV, "标准化.csv");
  assert.deepEqual(standardizedButtons(scope), [false, true, true]);
  assert.equal(scope.elements.get("standardizedExportGate").textContent, "");
  scope.evaluate("checkStandardizedExport()");
  assert.deepEqual(standardizedButtons(scope), [false, false, false]);
  // D11: a draft disables all three and says why; the check stays and comes back with the draft undone.
  scope.evaluate('state.fieldTypeDrafts["备注"] = "ignore"; updateFieldConfigActions()');
  assert.deepEqual(standardizedButtons(scope), [true, true, true]);
  assert.equal(scope.elements.get("standardizedExportGate").textContent, "字段配置有未应用的修改，请先应用或撤销后再检查。");
  assert.equal(standardizedPanelHidden(scope), false);
  scope.evaluate('state.fieldTypeDrafts["备注"] = "id"; updateFieldConfigActions()');
  assert.deepEqual(standardizedButtons(scope), [false, false, false]);
  assert.equal(scope.elements.get("standardizedExportGate").textContent, "");
});

test("v2.10.0 the page: a third export option, a static hidden panel, print hides its buttons, one binding", () => {
  const html = indexHtmlSource();
  const options = html.slice(html.indexOf('<div class="data-export-options">'), html.indexOf('<div id="standardizedExportPanel"'));
  assert.equal((options.match(/<div class="data-export-option">/g) || []).length, 2);
  assert.match(options, /<h3>写法统一的表<\/h3>[\s\S]*<p id="standardizedExportGate" class="standardize-gate"><\/p>[\s\S]*<button id="checkStandardizedExport" class="secondary-button" type="button" disabled>检查会改哪些格<\/button>/);
  const panel = html.slice(html.indexOf('<div id="standardizedExportPanel"'), html.indexOf('<p id="dataExportStatus"'));
  assert.match(panel, /^<div id="standardizedExportPanel" class="standardized-export hidden">/);
  ["standardizedExportSummary", "standardizedExportTable", "standardizedExportStatus"].forEach((id) => assert.ok(panel.includes(`id="${id}"`), id));
  assert.match(panel, /<button id="exportStandardizedCsv" type="button" disabled>导出写法统一的表<\/button>/);
  assert.match(panel, /<button id="exportStandardizedChanges" class="secondary-button" type="button" disabled>导出改动清单<\/button>/);
  const print = repoFile("style.css").match(/@media print \{([\s\S]*?)\n\}/)[1];
  assert.match(print, /\.standardized-export-actions,[\s\S]*display: none !important;/);
  const script = repoFile("script.js");
  assert.equal((script.match(/dom\.checkStandardizedExport\.addEventListener\("click", checkStandardizedExport\)/g) || []).length, 1);
});

// v2.10.0: the two standardized downloads, their names and the change-list limit.
const STANDARDIZE_CHANGE_LINES = [
  "文件行号,字段,原值,导出值,处理",
  "2,日期,2024/5/6,2024-05-06,日期改写为 YYYY-MM-DD",
  "2,金额,100元,100,数字改写为纯数字",
  "2,备注,\" a \",a,去首尾空格",
  "3,金额,\"2,200元\",2200,数字改写为纯数字",
  "3,区域,\" 华南\",华南,去首尾空格",
  "4,日期,2024-05-08 10:30,2024-05-08 10:30,带时间，保持原值",
  "4,金额,'-50元,-50,数字改写为纯数字",
  "4,区域,\"华北 \",华北,去首尾空格",
  "5,日期,待定,待定,读不成日期，保持原值",
  "5,金额,N/A,N/A,像是空值的写法，未确认",
  "5,年份,\" 2022\",2022,去首尾空格",
  "6,日期,2024年5月9日,2024-05-09,日期改写为 YYYY-MM-DD",
  "6,金额,\" 300 \",300,去首尾空格",
  "8,日期,06/06/2024,2024-06-06,日期改写为 YYYY-MM-DD",
  "8,金额,500元,500,数字改写为纯数字",
  "9,日期,2024/5/11,2024-05-11,日期改写为 YYYY-MM-DD",
  "9,金额,$12,12,数字改写为纯数字",
  "10,日期,2024/5/12,2024-05-12,日期改写为 YYYY-MM-DD",
  "11,日期,2024/5/13,2024-05-13,日期改写为 YYYY-MM-DD",
  "12,日期,2024/5/14,2024-05-14,日期改写为 YYYY-MM-DD",
  "13,日期,2024/5/15,2024-05-15,日期改写为 YYYY-MM-DD"
];

test("v2.10.0 the two downloads: every row standardized, every changed or kept cell listed, values neutralized", async () => {
  const scope = await importCsvFile(STANDARDIZE_CSV, "标准化.csv");
  const downloads = captureDownloads(scope);
  scope.evaluate("checkStandardizedExport()");
  scope.evaluate("var __standardizeCalls = 0; const __buildStandardizedExport = buildStandardizedExport; buildStandardizedExport = (...args) => { __standardizeCalls += 1; return __buildStandardizedExport(...args); }");
  scope.evaluate('exportStandardizedData("data")');
  const dataStatus = scope.elements.get("standardizedExportStatus").textContent;
  scope.evaluate('exportStandardizedData("changes")');
  const changesStatus = scope.elements.get("standardizedExportStatus").textContent;
  // The exports use the result of the check; nothing is computed again.
  assert.equal(scope.evaluate("__standardizeCalls"), 0);
  const [data, changes] = await downloads();
  assert.match(data.fileName, /^标准化_写法统一_\d{8}\.csv$/);
  assert.match(changes.fileName, /^标准化_写法统一改动清单_\d{8}\.csv$/);
  assert.deepEqual([data.bom, changes.bom], [[0xef, 0xbb, 0xbf], [0xef, 0xbb, 0xbf]]);
  assert.deepEqual([data.type, changes.type], ["text/csv;charset=utf-8", "text/csv;charset=utf-8"]);
  assert.equal(changes.text, STANDARDIZE_CHANGE_LINES.join("\r\n"));
  assert.deepEqual(data.text.split("\r\n"), [
    "编号,日期,金额,区域,年份,备注",
    "E01,2024-05-06,100,华东,2019,a",
    "E02,2024-05-07,2200,华南,2020,b",
    "E03,2024-05-08 10:30,-50,华北,2021年,c",
    "E04,待定,N/A,华东,2022,N/A",
    "E05,2024-05-09,300,华南,2023,e",
    "E06,2024-05-10,400,N/A,—,f",
    "E07,2024-06-06,500,华东,2019,g",
    "E08,2024-05-11,12,华南,2020,'=1+1",
    "E09,2024-05-12,600,华北,2021,i",
    "E10,2024-05-13,700,华东,2022,j",
    "E11,2024-05-14,800,华南,2023,k",
    "E12,2024-05-15,900,华北,2024,l"
  ]);
  assert.equal(dataStatus, `已导出写法统一的表 12 行：${data.fileName}`);
  assert.equal(changesStatus, `已导出改动清单 21 条：${changes.fileName}`);
  // State and rows stay as they were.
  assert.equal(scope.evaluate('state.rows[0]["备注"]'), " a ");
  assert.equal(scope.evaluate("state.standardizedExport.cellCount"), 21);
});

test("v2.10.0 D13: the change list stops at the limit, says so in its name and status, and counts everything", async () => {
  const scope = await importCsvFile(STANDARDIZE_CSV, "标准化.csv");
  scope.evaluate("checkStandardizedExport()");
  assert.equal(scope.evaluate("STANDARDIZED_CHANGE_LIMIT"), 1048575);
  const exportWith = async (limit) => {
    const downloads = captureDownloads(scope);
    scope.evaluate(`exportStandardizedData("changes", ${limit})`);
    const [file] = await downloads();
    return { ...file, status: scope.elements.get("standardizedExportStatus").textContent };
  };
  const over = await exportWith(3);
  assert.match(over.fileName, /^标准化_写法统一改动清单_前3条_\d{8}\.csv$/);
  assert.equal(over.text, STANDARDIZE_CHANGE_LINES.slice(0, 4).join("\r\n"));
  assert.equal(over.status, `只导出了前 3 条，共 21 条：${over.fileName}`);
  const equal = await exportWith(21);
  assert.match(equal.fileName, /^标准化_写法统一改动清单_\d{8}\.csv$/);
  assert.equal(equal.text, STANDARDIZE_CHANGE_LINES.join("\r\n"));
  assert.equal(equal.status, `已导出改动清单 21 条：${equal.fileName}`);
  const under = await exportWith(22);
  assert.equal(under.fileName, equal.fileName);
  assert.equal(under.text, equal.text);
  // The check table still counts every cell.
  assert.equal(scope.elements.get("standardizedExportSummary").textContent, "共改写 18 格，保持原值 3 格。其余 1 个字段没有改动。");
  // A large count reads with thousands separators.
  assert.equal(scope.evaluate('formatStandardizedChangesStatus(1048575, 2000000, "x.csv")'), "只导出了前 1,048,575 条，共 2,000,000 条：x.csv");
  assert.equal(scope.evaluate('formatStandardizedChangesStatus(21, 21, "x.csv")'), "已导出改动清单 21 条：x.csv");
});

test("v2.10.0 file names: the stem is sanitized, the date and the limit stay", async () => {
  const scope = await importCsvFile(STANDARDIZE_CSV, "a:b|c.csv");
  assert.equal(scope.evaluate('buildStandardizedFileName("data", new Date(2026, 10, 7))'), "a-b-c_写法统一_20261107.csv");
  assert.equal(scope.evaluate('buildStandardizedFileName("changes", new Date(2026, 10, 7))'), "a-b-c_写法统一改动清单_20261107.csv");
  assert.equal(scope.evaluate('buildStandardizedFileName("changes", new Date(2026, 10, 7), 1048575)'), "a-b-c_写法统一改动清单_前1048575条_20261107.csv");
  const long = await importCsvFile(STANDARDIZE_CSV, `${"长".repeat(120)}.xlsx`);
  const name = long.evaluate('buildStandardizedFileName("changes", new Date(2026, 10, 7), 3)');
  assert.ok(name.endsWith("_20261107.csv"), name);
});

test("v2.10.0 the exports need a check: before one, after it is dropped, and with no listed cell", async () => {
  const scope = await importCsvFile(STANDARDIZE_CSV, "标准化.csv");
  const downloads = captureDownloads(scope);
  scope.evaluate('exportStandardizedData("data")');
  assert.equal(scope.elements.get("standardizedExportStatus").textContent, "请先点“检查会改哪些格”。");
  scope.evaluate('checkStandardizedExport(); state.fieldTypeDrafts["备注"] = "ignore"; applyFieldConfiguration()');
  scope.evaluate('exportStandardizedData("changes")');
  assert.equal(scope.elements.get("standardizedExportStatus").textContent, "请先点“检查会改哪些格”。");
  const clean = await importCsvFile("名称,数量\r\nA,1\r\nB,2\r\nC,3\r\n", "clean.csv");
  const cleanDownloads = captureDownloads(clean);
  clean.evaluate('checkStandardizedExport(); exportStandardizedData("changes")');
  assert.equal(clean.elements.get("standardizedExportStatus").textContent, "没有需要列出的单元格。");
  clean.evaluate('exportStandardizedData("data")');
  const [cleanData] = await cleanDownloads();
  assert.equal(cleanData.text, "名称,数量\r\nA,1\r\nB,2\r\nC,3");
  assert.deepEqual(await downloads(), []);
});

test("v2.10.0 the export buttons are bound once each", () => {
  const script = repoFile("script.js");
  assert.equal((script.match(/dom\.exportStandardizedCsv\.addEventListener\("click", \(\) => exportStandardizedData\("data"\)\)/g) || []).length, 1);
  assert.equal((script.match(/dom\.exportStandardizedChanges\.addEventListener\("click", \(\) => exportStandardizedData\("changes"\)\)/g) || []).length, 1);
});

test("v2.10.0 all fixtures and examples: the downloads match the check and import back with the same types", async () => {
  const files = [...STAGE3_IMPORT_FILES.map((name) => ["stage3", name]), ...EXAMPLE_CSV_FILES.map((name) => ["example", name])];
  for (const [kind, name] of files) {
    const scope = kind === "stage3" ? await importStage3Fixture(name) : await importCsvFile(readExampleCsv(name), name);
    const downloads = captureDownloads(scope);
    scope.evaluate('checkStandardizedExport(); exportStandardizedData("data")');
    const cellCount = scope.evaluate("state.standardizedExport.cellCount");
    if (cellCount) scope.evaluate('exportStandardizedData("changes")');
    const [data, changes] = await downloads();
    assert.equal(data.text, scope.evaluate("buildCsvTableContent([...state.fields], buildStandardizedCellRows(state.standardizedExport, state.rows, state.fields))"), name);
    if (cellCount) {
      assert.equal(Papa.parse(changes.text, { header: true, skipEmptyLines: true }).data.length, cellCount, name);
      assert.ok(!/[^\r]\n/.test(changes.text), `${name} CRLF`);
    }
    assert.ok(!/[^\r]\n/.test(data.text), `${name} CRLF`);
    const again = await importCsvFile(data.text, `${name}.csv`);
    const types = (target) => target.evaluate("JSON.stringify(state.profiles.map((p) => [p.field, p.typeKey]))");
    assert.equal(types(again), types(scope), name);
  }
});

test("v2.10.0 the downloads ignore the filter and never guess a file row", async () => {
  const scope = await importCsvFile(STANDARDIZE_CSV, "标准化.csv");
  scope.evaluate("checkStandardizedExport()");
  applyFilterState(scope, { category: { field: "区域", values: ["华东"] } });
  assert.equal(scope.evaluate("getAnalysisRows().length"), 4);
  let downloads = captureDownloads(scope);
  scope.evaluate('exportStandardizedData("data"); exportStandardizedData("changes")');
  let [data, changes] = await downloads();
  assert.equal(data.text.split("\r\n").length, 13);
  assert.equal(changes.text, STANDARDIZE_CHANGE_LINES.join("\r\n"));
  // Rows committed without file rows: the first column stays empty.
  scope.evaluate('commitTabularData(state.fields, state.rows.map((row) => ({ ...row })), "无行号.csv", state.activeImportId)');
  assert.equal(scope.evaluate("getSourceRowNumber(state.rows[0])"), null);
  downloads = captureDownloads(scope);
  scope.evaluate('checkStandardizedExport(); exportStandardizedData("changes")');
  [changes] = await downloads();
  const lines = changes.text.split("\r\n");
  assert.equal(lines.length, 22);
  assert.ok(lines.slice(1).every((line) => line.startsWith(",")), lines[1]);
  assert.equal(lines[1], ",日期,2024/5/6,2024-05-06,日期改写为 YYYY-MM-DD");
});

// ---------------------------------------------------------------------------------------
// v2.11.0: the file the sample button loads (examples/sample-orders.csv).
// ---------------------------------------------------------------------------------------
const SAMPLE_ORDERS = "sample-orders.csv";
// Compared after dropping the CR of every CRLF: examples/ has no -text rule, so a Windows
// checkout with core.autocrlf=true has CRLF while the repository and GitHub Pages have LF.
const SAMPLE_ORDERS_LF_SHA256 = "39eec2fe25c23d5ede7411b699c2aef6f1d723f877c29f40d43faa0f60d9b947";
const sampleOrdersLfText = () => fs.readFileSync(path.join(REPO_ROOT, "examples", SAMPLE_ORDERS), "utf8").replace(/\r\n/g, "\n");

// Strips tags so a test can compare what a section says.
const sampleOrdersText = (scope, id) => scope.elements.get(id).innerHTML.replace(/<[^>]+>/g, " ").replace(/&quot;/g, "\"").replace(/\s+/g, " ").trim();
const sampleOrdersHidden = (scope, id) => scope.elements.get(id).classList.contains("hidden");
const sampleOrdersPills = (scope) => [...scope.elements.get("qualitySummary").innerHTML.matchAll(/<span class="quality-pill[^"]*">([^<]*)<\/span>/g)].map((match) => match[1]);
const SAMPLE_ORDERS_HINT_SECTIONS = ["suspectedTotalSection", "formulaErrorSection", "numberConversionSection", "numericLookingSection", "dateStyleSection", "constantFieldSection", "lookAlikeSection"];
const sampleOrdersShownSections = (scope) => SAMPLE_ORDERS_HINT_SECTIONS.filter((id) => !sampleOrdersHidden(scope, id));

// Clicks the real "使用示例数据" handler of script.js. The fake DOM ignores addEventListener, so
// the shared element prototype records the listeners while one context loads, then is restored.
async function clickSampleButton(fetchedText, language) {
  const probe = createScriptContext();
  const proto = Object.getPrototypeOf(probe.context.document.getElementById("loadSample"));
  const original = proto.addEventListener;
  const listeners = new Map();
  proto.addEventListener = function (type, listener) {
    if (!listeners.has(this)) listeners.set(this, []);
    listeners.get(this).push([type, listener]);
  };
  let scope;
  try {
    scope = createExcelContext(language === undefined ? {} : { language });
  } finally {
    proto.addEventListener = original;
  }
  put(scope.context, "Papa", Papa);
  scope.context.window.Papa = Papa;
  const fetched = [];
  const parsedNames = [];
  put(scope.context, "fetch", async (url) => {
    fetched.push(url);
    return { ok: true, text: async () => fetchedText };
  });
  const realParse = scope.evaluate("parseCsvText");
  // Forwards every argument: the title-row retry passes a fourth one (its line offset).
  put(scope.context, "parseCsvText", (...args) => {
    parsedNames.push(args[1]);
    return realParse(...args);
  });
  const clicks = (listeners.get(scope.elements.get("loadSample")) || []).filter(([type]) => type === "click");
  assert.equal(clicks.length, 1);
  await clicks[0][1]();
  return { ...scope, fetched, parsedNames };
}

// What fetch().text() gives: UTF-8 decoded, BOM removed.
const sampleOrdersFetchedText = () => new TextDecoder("utf-8").decode(fs.readFileSync(path.join(REPO_ROOT, "examples", SAMPLE_ORDERS)));

// Everything a quality check shows for one import, for comparing two entries.
const sampleOrdersSnapshot = (scope) => JSON.stringify({
  status: [scope.elements.get("statusPanel").textContent, scope.elements.get("statusPanel").className],
  pills: sampleOrdersPills(scope),
  sections: SAMPLE_ORDERS_HINT_SECTIONS.map((id) => [id, sampleOrdersHidden(scope, id)]),
  details: ["suspectedTotalDetails", "formulaErrorDetails", "numberConversionDetails", "numericLookingDetails", "dateStyleDetails", "constantFieldDetails", "lookAlikeDetails",
    "placeholderDetails", "idDuplicateDetails", "parseWarningSummary", "previewTable"].map((id) => [id, scope.elements.get(id).innerHTML]),
  problems: problemRowCheck(scope),
  lines: JSON.parse(scope.evaluate("JSON.stringify(listProblemCategories().map((category) => listProblemRows(category).items.map((item) => item.line)))")),
  types: JSON.parse(scope.evaluate("JSON.stringify(state.profiles.map((profile) => [profile.field, profile.typeKey]))"))
});

test("v2.11.0 sample: the file holds the planned rows, title and total line", () => {
  const raw = fs.readFileSync(path.join(REPO_ROOT, "examples", SAMPLE_ORDERS));
  assert.deepEqual([...raw.subarray(0, 3)], [0xef, 0xbb, 0xbf]);
  const text = sampleOrdersLfText();
  assert.equal(crypto.createHash("sha256").update(text, "utf8").digest("hex"), SAMPLE_ORDERS_LF_SHA256);
  assert.ok(text.endsWith("\n"));
  const lines = text.slice(1, -1).split("\n");
  assert.equal(lines.length, 44);
  assert.equal(lines[0], "2025年上半年销售订单明细,,,,,,,,");
  assert.equal(lines[1], "订单编号,下单日期,地区,品类,产品,客户类型,数量,金额,销售代表");
  assert.equal(lines[43], "合计,,,,,,800,123440,");
  assert.equal(lines[19], lines[20]);
});

test("v2.11.0 sample: the button fetches the new file and names the import after it, like an upload", async () => {
  const script = repoFile("script.js");
  assert.match(script, /fetch\(`examples\/\$\{sampleName\}`\)/);
  assert.match(script, /const sampleName = L\("sample-orders\.csv", "sample-orders-en\.csv"\);/);
  assert.doesNotMatch(script, /sample-data\.csv/);
  assert.match(indexHtmlSource(), /<a href="examples\/sample-orders\.csv" download>下载示例 CSV<\/a>/);
  const clicked = await clickSampleButton(sampleOrdersFetchedText());
  assert.deepEqual(clicked.fetched, ["examples/sample-orders.csv"]);
  // The handler calls parseCsvText() once; the title-row retry calls it again with the same name.
  assert.deepEqual(clicked.parsedNames, [SAMPLE_ORDERS, SAMPLE_ORDERS]);
  const uploaded = await importCsvFile(readExampleCsv(SAMPLE_ORDERS), SAMPLE_ORDERS);
  assert.equal(sampleOrdersSnapshot(clicked), sampleOrdersSnapshot(uploaded));
});

test("v2.11.0 sample: the status line, the six pills and every hint with its rows", async () => {
  const scope = await importCsvFile(readExampleCsv(SAMPLE_ORDERS), SAMPLE_ORDERS);
  assert.equal(scope.elements.get("statusPanel").className, "status-panel success has-note");
  assert.equal(scope.elements.get("statusPanel").textContent,
    "sample-orders.csv 已成功载入，分析结果已生成。表头在第 2 行，已跳过上方 1 行：「2025年上半年销售订单明细」。第 44 行「合计」疑似合计行，可以在“数据质量”里设为不计入分析。");
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(state.profiles.map((profile) => [profile.field, profile.typeKey]))")), [
    ["订单编号", "id"], ["下单日期", "date"], ["地区", "category"], ["品类", "category"], ["产品", "category"],
    ["客户类型", "category"], ["数量", "numeric"], ["金额", "numeric"], ["销售代表", "category"]
  ]);
  assert.deepEqual(sampleOrdersPills(scope), ["空值总数：6", "重复行：1", "认不出的数字或日期：1 个值", "偏大或偏小的值：2 个字段", "编号重复：1 个字段", "像是空值的写法：2 个字段"]);
  assert.deepEqual(sampleOrdersShownSections(scope), ["suspectedTotalSection", "numberConversionSection", "dateStyleSection", "lookAlikeSection"]);
  for (const id of ["categoryChartInsight", "dateChartInsight"]) assert.equal(sampleOrdersHidden(scope, id), false, id);
  assert.equal(scope.elements.get("categoryChartInsight").textContent, "分类字段「地区」中占比最高的类别是「华东」，出现 10 次，占比 24.4%。");
  assert.equal(scope.elements.get("dateChartInsight").textContent, "基于日期字段「下单日期」观察，记录数趋势 从 2025-01 到 2025-06 呈上升变化，变化幅度约 66.7%。");
  assert.equal(sampleOrdersHidden(scope, "parseWarningSummary"), true);
  assert.equal(sampleOrdersText(scope, "suspectedTotalDetails"), "第 44 行「合计」：数量、金额等于第 3\u201343 行之和（41 行）。 查看 不计入分析");
  assert.equal(sampleOrdersText(scope, "numberConversionDetails"), "「金额」：按\"元\"读取了 3 个值（5400元、900元、1050元）");
  assert.equal(sampleOrdersText(scope, "dateStyleDetails"),
    "「下单日期」：主写法是 文本 年-月-日 36 个（如 2025-01-03）；另有 5 行是别的写法：文本 年/月/日 3 个（如 2025/2/18）、文本 中文年月日 2 个（如 2025年4月16日）。 查看");
  assert.equal(sampleOrdersText(scope, "lookAlikeDetails"), "「地区」共 2 行：首尾空格 2 格（如「华南\u2423」）。 查看");
  assert.deepEqual(placeholderEntries(scope.evaluate), [["金额", "N/A", "common", 1, [29]], ["销售代表", "\u2014", "common", 1, [16]]]);
  assert.deepEqual(idDuplicates(scope.evaluate).map((item) => [item.field, item.values.map((value) => [value.value, value.count, value.rowNumbers])]), [["订单编号", [["SO-1018", 2, [20, 21]]]]]);
  const { names, mismatches } = problemRowCheck(scope);
  assert.deepEqual(mismatches, []);
  assert.deepEqual(names, [
    "missing / 下单日期 = 1", "missing / 地区 = 1", "missing / 品类 = 1", "missing / 产品 = 1", "missing / 客户类型 = 1", "missing / 销售代表 = 1",
    "conversion / 金额 = 1", "dateStyle / 下单日期 = 5", "outlier / 数量 = 1", "outlier / 金额 = 1", "duplicateRows = 2", "suspectedTotal = 1",
    "idDuplicate / 订单编号 = 2", "lookAlike / 地区 = 2", "placeholder / 金额 / N/A = 1", "placeholder / 销售代表 / \u2014 = 1"
  ]);
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(listProblemCategories().map((category) => listProblemRows(category).items.map((item) => item.line)))")), [
    [44], [44], [44], [44], [44], [44], [29], [12, 19, 25, 36, 39], [44], [44], [20, 21], [44], [20, 21], [10, 32], [29], [16]
  ]);
});

test("v2.11.0 sample: confirming every common placeholder moves the two values to missing and keeps the total", async () => {
  const scope = await importCsvFile(readExampleCsv(SAMPLE_ORDERS), SAMPLE_ORDERS);
  scope.evaluate("confirmAllCommonPlaceholders()");
  assert.deepEqual(sampleOrdersPills(scope), ["空值总数：8", "重复行：1", "认不出的数字或日期：0 个值", "偏大或偏小的值：2 个字段", "编号重复：1 个字段", "像是空值的写法：0 个字段"]);
  assert.equal(scope.evaluate("state.suspectedTotals.length"), 1);
  assert.deepEqual(JSON.parse(scope.evaluate('JSON.stringify(["金额", "销售代表"].map((field) => getFieldProfile(field).missingCount))')), [1, 2]);
});

test("v2.11.0 sample: the standardized check counts 11 cells, 12 after confirming", async () => {
  const scope = await importCsvFile(readExampleCsv(SAMPLE_ORDERS), SAMPLE_ORDERS);
  scope.evaluate("checkStandardizedExport()");
  const changed = (counts) => Object.fromEntries(Object.entries(counts).filter(([, [, ...numbers]]) => numbers.some(Boolean)));
  assert.deepEqual(changed(standardizedCounts(scope)), {
    下单日期: ["date", 5, 0, 0, 0, 0], 地区: ["text", 0, 0, 0, 2, 0], 金额: ["number", 0, 3, 0, 0, 1]
  });
  assert.equal(scope.evaluate("state.standardizedExport.cellCount"), 11);
  assert.equal(scope.evaluate("state.standardizedExport.counts.some((count) => count.orderHint)"), false);
  scope.evaluate("confirmAllCommonPlaceholders(); checkStandardizedExport()");
  assert.deepEqual(changed(standardizedCounts(scope)), {
    下单日期: ["date", 5, 0, 0, 0, 0], 地区: ["text", 0, 0, 0, 2, 0], 金额: ["number", 0, 3, 1, 0, 0], 销售代表: ["text", 0, 0, 1, 0, 0]
  });
  assert.equal(scope.evaluate("state.standardizedExport.cellCount"), 12);
});

// The sample without its problems, built here (no clean file is committed): the hints only show
// what was put into the file. Fix values: SO-1026 is 6 earphones at 260 each (SO-1010, SO-1038);
// every other 华东 order is 张敏's.
const SAMPLE_ORDERS_FIXES = {
  title: (lines) => lines.slice(1),
  total: (lines) => lines.filter((line) => !line.startsWith("合计,")),
  duplicate: (lines) => lines.filter((line, index) => !(line.startsWith("SO-1018,") && line === lines[index - 1])),
  na: (lines) => lines.map((line) => (line.startsWith("SO-1026,") ? line.replace(",N/A,", ",1560,") : line)),
  dash: (lines) => lines.map((line) => (line.startsWith("SO-1014,") ? line.replace(/,\u2014$/, ",张敏") : line)),
  unit: (lines) => lines.map((line) => line.replace(/,(\d+)元,/, ",$1,")),
  date: (lines) => lines.map((line) => line.replace(/^(SO-\d+),(\d{4})(?:\/|年)(\d{1,2})(?:\/|月)(\d{1,2})日?,/, (match, id, year, month, day) => `${id},${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")},`)),
  space: (lines) => lines.map((line) => line.replace(/ +,/g, ","))
};

const sampleOrdersLines = () => sampleOrdersLfText().slice(1, -1).split("\n");

function sampleOrdersVariant(keep) {
  let lines = sampleOrdersLines();
  for (const [name, fix] of Object.entries(SAMPLE_ORDERS_FIXES)) if (name !== keep) lines = fix(lines);
  return "\uFEFF" + lines.join("\n") + "\n";
}

test("v2.11.0 sample: without its problems nothing is flagged, and each problem alone flags only itself", async () => {
  const PILLS_ZERO = ["空值总数：0", "重复行：0", "认不出的数字或日期：0 个值", "偏大或偏小的值：0 个字段", "编号重复：0 个字段", "像是空值的写法：0 个字段"];
  const cases = {
    clean: { pills: PILLS_ZERO, sections: [], note: false, cells: 0 },
    title: { pills: PILLS_ZERO, sections: [], note: true, cells: 0 },
    total: { pills: ["空值总数：6", "重复行：0", "认不出的数字或日期：0 个值", "偏大或偏小的值：2 个字段", "编号重复：0 个字段", "像是空值的写法：0 个字段"], sections: ["suspectedTotalSection"], note: true, cells: 0 },
    duplicate: { pills: ["空值总数：0", "重复行：1", "认不出的数字或日期：0 个值", "偏大或偏小的值：0 个字段", "编号重复：1 个字段", "像是空值的写法：0 个字段"], sections: [], note: false, cells: 0 },
    na: { pills: ["空值总数：0", "重复行：0", "认不出的数字或日期：1 个值", "偏大或偏小的值：0 个字段", "编号重复：0 个字段", "像是空值的写法：1 个字段"], sections: [], note: false, cells: 1 },
    dash: { pills: ["空值总数：0", "重复行：0", "认不出的数字或日期：0 个值", "偏大或偏小的值：0 个字段", "编号重复：0 个字段", "像是空值的写法：1 个字段"], sections: [], note: false, cells: 0 },
    unit: { pills: PILLS_ZERO, sections: ["numberConversionSection"], note: false, cells: 3 },
    date: { pills: PILLS_ZERO, sections: ["dateStyleSection"], note: false, cells: 5 },
    space: { pills: PILLS_ZERO, sections: ["lookAlikeSection"], note: false, cells: 2 }
  };
  for (const [name, expected] of Object.entries(cases)) {
    const scope = await importCsvFile(sampleOrdersVariant(name === "clean" ? null : name), SAMPLE_ORDERS);
    assert.deepEqual(sampleOrdersPills(scope), expected.pills, name);
    assert.deepEqual(sampleOrdersShownSections(scope), expected.sections, name);
    assert.equal(scope.elements.get("statusPanel").className.includes("has-note"), expected.note, name);
    assert.equal(scope.evaluate("state.placeholders.length"), ["na", "dash"].includes(name) ? 1 : 0, name);
    assert.equal(scope.evaluate("state.parseWarnings.length"), 0, name);
    scope.evaluate("checkStandardizedExport()");
    assert.equal(scope.evaluate("state.standardizedExport.cellCount"), expected.cells, name);
    assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(state.profiles.map((profile) => profile.typeKey))")),
      ["id", "date", "category", "category", "category", "category", "numeric", "numeric", "category"], name);
  }
  // The builder itself: each fix changes the file on its own (a fix that stops matching would
  // leave its problem in every case above), and the clean version has 40 orders and no title.
  const full = sampleOrdersLines();
  for (const [name, fix] of Object.entries(SAMPLE_ORDERS_FIXES)) assert.notDeepEqual(fix(full), full, name);
  assert.equal(sampleOrdersVariant(null).split("\n").length, 1 + 40 + 1);
});

// v2.12.0: the interface terms were unified. None of the replaced words may come back on the
// page, in reports, in problem-row names, aria labels, file names and notes, or in the 处理
// column of the change list (problem-row notes and change-list actions are exported as text).
const OLD_INTERFACE_WORDS = ["缺失", "未填写", "疑似缺失值", "疑似值", "按缺失处理", "格式问题", "无法转换", "转换失败", "异常值", "数值异常",
  "写法细节不同", "写法不同", "解析警告", "条警告", "ID 字段重复", "ID 重复", "ID重复", "常量列", "疑似数值列", "标准化", "可安全转换"];
// The notes shared by the page and the reports: no space between two Chinese characters.
// (v3.2.0: the constants are [中文, English] pairs; [0] is the Chinese.)
const SHARED_NOTE_NAMES = ["SUSPECTED_TOTAL_NOTE[0]", "FORMULA_ERROR_NOTE[0]", "DATE_STYLE_NOTE[0]", "CONSTANT_FIELD_NOTE[0]", "LOOK_ALIKE_NOTE[0]", "PLACEHOLDER_MAYBE_NOTE[0]", "PLACEHOLDER_DASH_NOTE[0]",
  // v3.1.0: the totals note once they are left out, one row and several.
  "formatSuspectedTotalExcludedNote(1)", "formatSuspectedTotalExcludedNote(5)"];
const HAN_SPACE_HAN = /\p{Script=Han} +\p{Script=Han}/u;

function findOldInterfaceWords(where, text, hits) {
  for (const word of OLD_INTERFACE_WORDS) {
    const at = String(text).indexOf(word);
    if (at >= 0) hits.push(`${where}: 「${word}」 in …${String(text).slice(Math.max(0, at - 20), at + word.length + 20)}…`);
  }
}

function findOldWordsInElements(scope, where, hits, prefix = "") {
  for (const [id, element] of scope.elements) {
    if (prefix && !id.startsWith(prefix)) continue;
    findOldInterfaceWords(`${where} #${id}.innerHTML`, element.innerHTML, hits);
    findOldInterfaceWords(`${where} #${id}.textContent`, element.textContent, hits);
    for (const option of element.options || []) findOldInterfaceWords(`${where} #${id} option`, `${option.textContent ?? ""} ${option.label ?? ""}`, hits);
  }
}

// Import → report → every problem-row category (name, aria label, file name, notes, viewer) →
// 检查会改哪些格 → confirm all → check again with the change list → trend summed by a number.
function findOldWordsInImport(scope, name, hits) {
  findOldWordsInElements(scope, `${name} import`, hits);
  const first = reportOf(scope);
  findOldInterfaceWords(`${name} html report`, first.html, hits);
  const categories = JSON.parse(scope.evaluate("JSON.stringify(state.problemCategories)"));
  categories.forEach((category, index) => {
    put(scope.context, "__oldWordsCategory", category);
    findOldInterfaceWords(`${name} category ${index}`, scope.evaluate("formatProblemCategoryLabel(__oldWordsCategory)"), hits);
    findOldInterfaceWords(`${name} category ${index} aria`, scope.evaluate("formatProblemViewLabel(__oldWordsCategory)"), hits);
    findOldInterfaceWords(`${name} category ${index} file`, scope.evaluate("buildProblemRowsFileName(__oldWordsCategory, new Date(2026, 10, 12))"), hits);
    findOldInterfaceWords(`${name} category ${index} notes`, scope.evaluate("JSON.stringify(listProblemRows(__oldWordsCategory).items.map((item) => item.note))"), hits);
    scope.evaluate("openProblemView(__oldWordsCategory)");
    findOldWordsInElements(scope, `${name} category ${index} view`, hits, "problemRows");
  });
  scope.evaluate("checkStandardizedExport()");
  findOldWordsInElements(scope, `${name} standardized`, hits, "standardized");
  findOldInterfaceWords(`${name} standardized files`, scope.evaluate('buildStandardizedFileName("data", new Date(2026, 10, 12)) + buildStandardizedFileName("changes", new Date(2026, 10, 12), 3)'), hits);
  scope.evaluate("confirmAllCommonPlaceholders()");
  // Checked again: the 处理 of a confirmed placeholder only shows up after a confirmation, and
  // only in the change list.
  scope.evaluate("checkStandardizedExport()");
  findOldWordsInElements(scope, `${name} confirmed`, hits);
  findOldInterfaceWords(`${name} change list`, scope.evaluate("buildStandardizedChangesCsvContent(state.standardizedExport, state.rows, state.fields, getSourceRowNumber)"), hits);
  const confirmed = reportOf(scope);
  findOldInterfaceWords(`${name} confirmed html report`, confirmed.html, hits);

  const numeric = scope.evaluate('getProfilesByType("numeric").map((profile) => profile.field)')[0];
  const dateField = scope.evaluate('getProfilesByType("date").map((profile) => profile.field)')[0];
  if (!numeric || !dateField) return 0;
  class FakeChart { destroy() {} }
  scope.context.Chart = FakeChart;
  scope.context.window.Chart = FakeChart;
  put(scope.context, "__oldWordsMetric", numeric);
  put(scope.context, "__oldWordsDate", dateField);
  scope.evaluate('dom.dateChartField.value = __oldWordsDate; dom.dateChartMetric.value = __oldWordsMetric; dom.dateChartAggregate.value = "sum"; renderDateTrendChart()');
  const title = scope.elements.get("dateChartTitle").textContent;
  findOldInterfaceWords(`${name} trend title`, title, hits);
  if (title && !title.endsWith("（求和）")) hits.push(`${name} trend title: ${title}`);
  const trend = reportOf(scope);
  findOldInterfaceWords(`${name} trend report`, trend.html, hits);
  return title ? 1 : 0;
}

test("v2.12.0: old interface words no longer appear in pages, reports, file names or exported notes", async () => {
  const hits = [];
  const html = indexHtmlSource().replace(/<!--[\s\S]*?-->/g, "");
  findOldInterfaceWords("index.html", html, hits);
  const selectStart = html.indexOf('<select id="dateChartAggregate">');
  assert.ok(selectStart >= 0);
  const aggregateOptions = html.slice(selectStart, html.indexOf("</select>", selectStart));
  if (!aggregateOptions.includes('<option value="sum">求和</option>') || aggregateOptions.includes("合计")) hits.push(`index.html #dateChartAggregate: ${aggregateOptions}`);

  const { evaluate } = createScriptContext();
  for (const name of SHARED_NOTE_NAMES) {
    if (HAN_SPACE_HAN.test(evaluate(name))) hits.push(`${name}: a space between two Chinese characters`);
  }

  let trendTitles = 0;
  for (const fileName of STAGE3_IMPORT_FILES) trendTitles += findOldWordsInImport(await importStage3Fixture(fileName), fileName, hits);
  for (const fileName of EXAMPLE_CSV_FILES) trendTitles += findOldWordsInImport(await importCsvFile(readExampleCsv(fileName), fileName), fileName, hits);
  // No fixture has a parse warning or an empty custom-analysis group: a short row and a blank
  // 备注 cover 读取时有疑问 and the empty group's name.
  const warned = await importCsvFile("名,额,备注\r\nA,1,x\r\nB\r\nC,3,\r\nD,4,x\r\nE,5,y\r\n", "warned.csv");
  assert.equal(warned.evaluate("state.parseWarnings.length"), 1);
  warned.evaluate('dom.v2MetricField.value = "额"; dom.v2GroupField.value = "备注"; dom.v2AggregateMethod.value = "sum"; dom.v2DateField.value = ""; renderV2Analysis()');
  assert.equal(JSON.parse(warned.evaluate("JSON.stringify(state.customAnalysis.grouped.map((item) => item.name))")).length, 3);
  findOldWordsInImport(warned, "warned.csv", hits);
  assert.ok(trendTitles > 0);

  assert.deepEqual(hits.slice(0, 20), [], `${hits.length} hits`);
});

// v2.12.0 (D3): the four "nothing readable" sentences. No fixture reaches them, so a filtered
// range without dates or numbers is built here: 25 rows of 组 A with a date and an amount,
// then 5 rows of 组 B with neither, filtered to 组 B.
test("v2.12.0: charts, insights and the report say 认得出 when a filtered range has no dates or numbers", () => {
  const scope = createChartContext();
  const rows = Array.from({ length: 30 }, (_, index) => (index < 25
    ? { 日期: `2025-01-${String(index + 1).padStart(2, "0")}`, 金额: String(100 + index * 10), 组: "A" }
    : { 日期: "", 金额: "", 组: "B" }));
  put(scope.context, "__emptyRangeRows", rows);
  scope.evaluate('commitTabularData(["日期", "金额", "组"], __emptyRangeRows, "empty-range.csv", beginImport())');
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(state.profiles.map((profile) => profile.typeKey))")), ["date", "numeric", "category"]);

  scope.evaluate(`(() => {
    globalThis.__emptyChartTexts = [];
    globalThis.__originalDrawEmptyChart = drawEmptyChart;
    drawEmptyChart = function (canvas, text) {
      const id = ["dateTrendChart", "numericDistributionChart"].find((name) => document.getElementById(name) === canvas) || "other";
      __emptyChartTexts.push([id, text]);
      return __originalDrawEmptyChart(canvas, text);
    };
  })()`);
  try {
    scope.evaluate(`(() => {
      state.filters = { category: { field: "组", values: ["B"] }, numeric: null, date: null };
      state.filteredRows = filterRowsByConfig(state.rows, state.filters);
      state.filterSourceRows = state.rows;
      refreshFilteredAnalysis();
    })()`);
    assert.equal(scope.evaluate("getAnalysisRows().length"), 5);
    const texts = () => JSON.parse(scope.evaluate("JSON.stringify(__emptyChartTexts)"));
    assert.ok(texts().some(([id, text]) => id === "dateTrendChart" && text === "该字段没有认得出的日期"), JSON.stringify(texts()));
    assert.equal(scope.evaluate("state.chartInsights.trend"), "日期字段「日期」里没有认得出的日期，因此未生成趋势结论。");
    const { html } = reportOf(scope);
    assert.ok(html.includes("日期字段「日期」里没有认得出的日期。"));

    // fieldHasNumericData() drops 金额 from #numericChartField once the filtered range has no
    // numbers (in a browser too), so the page shows 没有可用的数值字段; this sentence is the
    // fallback for a selected field with nothing readable, reached here by setting the value by hand.
    scope.evaluate('__emptyChartTexts.length = 0; dom.numericChartField.value = "金额"; renderNumericDistributionChart()');
    assert.deepEqual(texts(), [["numericDistributionChart", "该字段没有认得出的数字"]]);
  } finally {
    scope.evaluate("drawEmptyChart = __originalDrawEmptyChart");
  }
});

// v3.0.0 guards (commit 0). They pass on v2.12.0: they pin what the new shell must leave as it
// is, so the later commits cannot change it unnoticed.

// A fake Chart whose instances the report can capture (captureCurrentChartImages() keeps only
// png / jpeg / webp data URLs).
function withCapturableCharts(scope) {
  scope.context.Chart = scope.context.window.Chart = class {
    constructor(canvas, config) { this.canvas = canvas; this.data = config.data; }
    destroy() {}
    toBase64Image() { return "data:image/png;base64,iVBORw0KGgo="; }
  };
}
const capturedCharts = (scope) => JSON.parse(scope.evaluate("JSON.stringify(captureCurrentChartImages().map((chart) => [chart.id, chart.title]))"));
const GUARD_CUSTOM_ANALYSIS = 'dom.v2MetricField.value = "金额"; dom.v2GroupField.value = "地区"; dom.v2AggregateMethod.value = "sum"; dom.v2DateField.value = "下单日期"; renderV2Analysis()';
const GUARD_EAST_FILTER = 'state.filters = { category: { field: "地区", values: ["华东"] }, numeric: null, date: null }; state.filteredRows = filterRowsByConfig(state.rows, state.filters); state.filterSourceRows = state.rows; refreshFilteredAnalysis()';

test("v3.0.0 guard: the HTML report takes the same charts and titles after import, a custom analysis and a filter", async () => {
  const orders = await importCsvFile(readExampleCsv("sample-orders.csv"), "sample-orders.csv");
  withCapturableCharts(orders);
  // A browser <select> shows its first option; the fake one stays "" until set (contract 11).
  orders.evaluate('dom.numericChartField.value = "数量"; renderCharts()');
  const three = [
    ["numericDistributionChart", "数量 数值分布（记录频数）"],
    ["categoryTopChart", "地区 Top 10（记录数）"],
    ["dateTrendChart", "下单日期 按月份趋势（记录数）"]
  ];
  assert.deepEqual(capturedCharts(orders), three);
  orders.evaluate(GUARD_CUSTOM_ANALYSIS);
  const five = [...three, ["v2TopGroupChart", "金额（元） 按 地区 分组 Top 10（求和）"], ["v2SelectedTrendChart", "金额（元） 按月份趋势（求和）"]];
  assert.deepEqual(capturedCharts(orders), five);
  // The filter reruns the open custom analysis, so all five charts stay.
  orders.evaluate(GUARD_EAST_FILTER);
  assert.equal(orders.evaluate("getAnalysisRows().length"), 10);
  assert.deepEqual(capturedCharts(orders), five);
});

// Records every drawEmptyChart() call as [canvas id, text]; installed before the commit so the
// sentences drawn while importing count too.
function collectEmptyChartTexts(scope) {
  scope.evaluate(`(() => {
    globalThis.__guardEmptyTexts = [];
    const ids = ["numericDistributionChart", "categoryTopChart", "dateTrendChart", "v2TopGroupChart", "v2SelectedTrendChart"];
    const original = drawEmptyChart;
    drawEmptyChart = function (canvas, text) {
      __guardEmptyTexts.push([ids.find((id) => document.getElementById(id) === canvas) || "other", text]);
      return original(canvas, text);
    };
  })()`);
  return () => JSON.parse(scope.evaluate("JSON.stringify(__guardEmptyTexts.splice(0))"));
}
function commitGuardRows(scope, fields, rows, fileName) {
  put(scope.context, "__guardFields", fields);
  put(scope.context, "__guardRows", rows);
  scope.evaluate(`commitTabularData(__guardFields, __guardRows, ${JSON.stringify(fileName)}, beginImport())`);
}
function filterGuardRows(scope, field, values) {
  put(scope.context, "__guardFilter", { field, values });
  scope.evaluate("state.filters = { category: __guardFilter, numeric: null, date: null }; state.filteredRows = filterRowsByConfig(state.rows, state.filters); state.filterSourceRows = state.rows; refreshFilteredAnalysis()");
}

// v3.0.0 F3 restyles these sentences; every one must still appear, word for word, under its
// trigger. (这一类已经没有问题行。 is pinned by the contract 19 tests.)
test("v3.0.0 guard: every empty-state sentence still appears under its trigger", () => {
  const html = (scope, id) => scope.elements.get(id).innerHTML;
  const text = (scope, id) => scope.elements.get(id).textContent;

  // Only text columns: no numeric, date or ID field, no placeholder.
  const words = createChartContext();
  const wordTexts = collectEmptyChartTexts(words);
  commitGuardRows(words, ["名称", "备注"], Array.from({ length: 20 }, (_, index) => ({ 名称: ["甲", "乙", "丙"][index % 3], 备注: index % 2 ? "x" : "y" })), "words.csv");
  assert.deepEqual(JSON.parse(words.evaluate("JSON.stringify(state.profiles.map((profile) => profile.typeKey))")), ["category", "category"]);
  assert.match(html(words, "outlierTable"), /<td>未识别到数值字段。<\/td>/);
  assert.match(html(words, "numericStatsTable"), /<td>未识别到数值字段。<\/td>/);
  assert.match(html(words, "idDuplicateDetails"), />未识别到 ID 字段。</);
  assert.match(html(words, "placeholderDetails"), />没有发现像是空值的写法。</);
  assert.equal(text(words, "numericFilterRange"), "没有可用字段");
  assert.equal(text(words, "dateFilterRange"), "没有可用字段");
  assert.equal(text(words, "categoryFilterEmpty"), "请选择分类字段。");
  assert.deepEqual(wordTexts(), [["numericDistributionChart", "没有可用的数值字段"], ["dateTrendChart", "没有可用的日期字段"]]);
  words.evaluate('dom.categoryFilterField.value = "名称"; updateCategoryFilterField(); dom.categoryFilterSearch.value = "zzz"; renderCategoryFilterValues()');
  assert.match(html(words, "categoryFilterValues"), />没有匹配的类别。</);

  // Only a numeric column: no category field.
  const amounts = createChartContext();
  const amountTexts = collectEmptyChartTexts(amounts);
  commitGuardRows(amounts, ["金额"], Array.from({ length: 20 }, (_, index) => ({ 金额: String(index * 10 + 1) })), "amounts.csv");
  assert.match(html(amounts, "categoryStats"), />未识别到分类字段。</);
  assert.equal(text(amounts, "categoryFilterEmpty"), "没有可用的分类字段。");
  assert.deepEqual(amountTexts().filter(([id]) => id === "categoryTopChart"), [["categoryTopChart", "没有可用的分类字段"]]);

  // A filtered range where 日期, 金额 and 备注 are all empty, then a range with no rows.
  const range = createChartContext();
  const rangeTexts = collectEmptyChartTexts(range);
  commitGuardRows(range, ["日期", "金额", "组", "备注"], Array.from({ length: 30 }, (_, index) => (index < 25
    ? { 日期: `2025-01-${String(index + 1).padStart(2, "0")}`, 金额: String(100 + index * 10), 组: "A", 备注: index % 2 ? "x" : "y" }
    : { 日期: "", 金额: "", 组: "B", 备注: "" })), "range.csv");
  assert.deepEqual(JSON.parse(range.evaluate("JSON.stringify(state.profiles.map((profile) => profile.typeKey))")), ["date", "numeric", "category", "category"]);
  rangeTexts();
  filterGuardRows(range, "组", ["B"]);
  assert.equal(range.evaluate("getAnalysisRows().length"), 5);
  assert.match(html(range, "categoryStats"), />当前分析的行里没有非空类别值。</);
  range.evaluate('dom.categoryChartField.value = "备注"; renderCategoryTopChart(); dom.numericChartField.value = "金额"; renderNumericDistributionChart()');
  const filtered = rangeTexts();
  assert.ok(filtered.some(([id, sentence]) => id === "dateTrendChart" && sentence === "该字段没有认得出的日期"), JSON.stringify(filtered));
  assert.ok(filtered.some(([id, sentence]) => id === "categoryTopChart" && sentence === "当前分析的行里没有可统计的分类值"), JSON.stringify(filtered));
  assert.ok(filtered.some(([id, sentence]) => id === "numericDistributionChart" && sentence === "该字段没有认得出的数字"), JSON.stringify(filtered));
  filterGuardRows(range, "组", ["C"]);
  assert.equal(range.evaluate("getAnalysisRows().length"), 0);
  assert.match(html(range, "previewTable"), /<td>当前筛选条件下没有可预览的数据。<\/td>/);

  // A filtered range with amounts but no dates: the custom analysis has no trend.
  const trend = createChartContext();
  const trendTexts = collectEmptyChartTexts(trend);
  commitGuardRows(trend, ["日期", "金额", "组"], Array.from({ length: 30 }, (_, index) => ({
    日期: index < 25 ? `2025-01-${String(index + 1).padStart(2, "0")}` : "", 金额: String(100 + index * 10), 组: index < 25 ? "A" : "B"
  })), "trend.csv");
  filterGuardRows(trend, "组", ["B"]);
  trendTexts();
  trend.evaluate('dom.v2MetricField.value = "金额"; dom.v2GroupField.value = "组"; dom.v2AggregateMethod.value = "sum"; dom.v2DateField.value = "日期"; renderV2Analysis()');
  assert.deepEqual(trendTexts(), [["v2SelectedTrendChart", "所选时间字段没有可用的趋势数据"]]);

  // Chart.js missing: every chart says so.
  const noChart = createScriptContext();
  const noChartTexts = collectEmptyChartTexts(noChart);
  commitGuardRows(noChart, ["日期", "金额", "组"], Array.from({ length: 30 }, (_, index) => ({
    日期: `2025-01-${String(index + 1).padStart(2, "0")}`, 金额: String(100 + index * 10), 组: index % 2 ? "A" : "B"
  })), "no-chart.csv");
  noChartTexts();
  noChart.evaluate('renderCharts(); dom.v2MetricField.value = "金额"; dom.v2GroupField.value = "组"; dom.v2AggregateMethod.value = "sum"; dom.v2DateField.value = "日期"; renderV2Analysis()');
  const failed = "图表组件加载失败，请刷新页面重试";
  assert.deepEqual(noChartTexts(), [
    ["numericDistributionChart", failed], ["categoryTopChart", failed], ["dateTrendChart", failed],
    ["v2TopGroupChart", failed], ["v2SelectedTrendChart", failed]
  ]);
});

// The report text of the sample (custom analysis, then filtered to 华东), times and version
// blanked, is the v2.12.0 text with only the name changed: neither the v3.0.0 shell nor the
// v3.0.1 rename touches any other word. v3.0.1 renamed the product and the report title; the
// new hash was taken by replacing just those words in the old text and comparing byte for byte.
// v3.2.2 only changed the repository URL in the footer to sheet-checkup; the new value is the old
// text (e32ff801…21421) with just that one replacement, compared byte for byte.
const GUARD_REPORT_SHA256 = "937a0034ba4f73fe6ecbfae26a7428aba0b367273b27bd740f31ae1db782f5f2";
test("v3.0.1 guard: the sample's report text is the v2.12.0 text with only the name changed", async () => {
  const orders = await importCsvFile(readExampleCsv("sample-orders.csv"), "sample-orders.csv");
  const version = orders.evaluate("buildReportData().version");
  const texts = [];
  const snapshot = () => {
    const report = JSON.parse(orders.evaluate("JSON.stringify(buildReportData())"));
    ["analysisTime", "analysisTimeIso", "exportTime"].forEach((key) => {
      assert.ok(key in report, key);
      report[key] = "T";
    });
    put(orders.context, "__guardReport", report);
    texts.push(orders.evaluate("buildHtmlReport(__guardReport, [])"));
  };
  orders.evaluate(GUARD_CUSTOM_ANALYSIS);
  snapshot();
  orders.evaluate(GUARD_EAST_FILTER);
  snapshot();
  const normalized = texts.join("\n----\n").split(version).join("VERSION");
  assert.ok(normalized.includes("VERSION"));
  assert.equal(crypto.createHash("sha256").update(normalized).digest("hex"), GUARD_REPORT_SHA256);
});

// v3.0.0 commit 1: one module at a time.
const MODULE_SECTION_IDS = ["overviewSection", "schemaSection", "qualitySection", "filterSection", "statisticsSection", "chartsSection", "customSection", "exportSection"];

// The fake page has no navigation links, history or scrolling, and its focus() does nothing:
// this gives showModule() all three and records what it did.
function withModuleNavigation(scope) {
  const record = { urls: [], scrolls: [], focused: [] };
  const links = MODULE_SECTION_IDS.map((id) => {
    const attributes = { href: `#${id}` };
    const classes = new Set();
    return {
      getAttribute: (name) => (name in attributes ? attributes[name] : null),
      setAttribute: (name, value) => { attributes[name] = String(value); },
      removeAttribute: (name) => { delete attributes[name]; },
      classList: {
        toggle(name, force) { if (force) classes.add(name); else classes.delete(name); return classes.has(name); },
        contains: (name) => classes.has(name)
      }
    };
  });
  const document = scope.context.document;
  const originalQuery = document.querySelectorAll;
  document.querySelectorAll = (selector) => (selector === "#moduleNav a" ? links : originalQuery.call(document, selector));
  scope.context.window.history = { replaceState: (state, title, url) => record.urls.push(url) };
  // JSON copy: objects made inside the VM have another realm's prototype.
  scope.context.window.scrollTo = (...args) => record.scrolls.push(JSON.parse(JSON.stringify(args)));
  MODULE_SECTION_IDS.forEach((id) => { document.getElementById(`${id}Title`).focus = () => record.focused.push(id); });
  document.getElementById("reportTitle").focus = () => record.focused.push("reportTitle");
  record.current = () => links.filter((link) => link.getAttribute("aria-current") === "page").map((link) => link.getAttribute("href"));
  record.shown = () => MODULE_SECTION_IDS.filter((id) => !document.getElementById(id).classList.contains("module-off"));
  record.inert = () => MODULE_SECTION_IDS.filter((id) => document.getElementById(id).inert === true);
  record.hiddenClass = () => MODULE_SECTION_IDS.filter((id) => document.getElementById(id).classList.contains("hidden"));
  record.clear = () => { record.urls.length = 0; record.scrolls.length = 0; record.focused.length = 0; };
  return record;
}

test("v3.0.0 modules: the page wraps the eight modules in one stage, each title can take the focus", () => {
  const html = indexHtmlSource();
  const stageStart = html.indexOf('<div id="moduleStage" class="module-stage">');
  assert.ok(stageStart >= 0);
  assertInOrder(html.slice(stageStart), MODULE_SECTION_IDS.map((id) => `<section id="${id}"`));
  MODULE_SECTION_IDS.forEach((id) => {
    const section = html.slice(html.indexOf(`<section id="${id}"`));
    assert.match(section.slice(0, 400), new RegExp(`<h2 id="${id}Title" tabindex="-1">`), id);
  });
  assert.match(html, /<nav id="moduleNav"[^>]*>/);
});

test("v3.0.0 modules: hidden modules keep their layout (visibility, never display: none or .hidden), print shows all", () => {
  const css = repoFile("style.css");
  const stage = css.match(/\n\.module-stage \{([^}]*)\}/);
  assert.ok(stage, ".module-stage rule");
  assert.match(stage[1], /overflow: clip;/);
  const off = css.match(/\n\.module-stage \.module-off \{([^}]*)\}/);
  assert.ok(off, ".module-off rule");
  assert.match(off[1], /visibility: hidden;/);
  assert.match(off[1], /position: absolute;/);
  assert.doesNotMatch(off[1], /display:/);
  const script = repoFile("script.js");
  const start = script.indexOf("function showModule(");
  assert.ok(start >= 0);
  const showModule = script.slice(start, start + script.slice(start).search(/\r?\n\}\r?\n/));
  assert.match(showModule, /module-off/);
  assert.doesNotMatch(showModule, /"hidden"/);
  const print = css.match(/@media print \{([\s\S]*?)\n\}/)[1];
  assert.equal((css.match(/@media print \{/g) || []).length, 1);
  assert.match(print, /\.module-stage \.module-off \{[^}]*position: static !important;[^}]*visibility: visible !important;/);
  assert.match(print, /\.skip-link,/);
  assert.match(print, /\.chart-frame canvas \{[^}]*max-width: 100% !important;/);
  assert.match(print, /\.chart-frame,\s*\.chart-card\.wide \.chart-frame \{\s*height: auto;/);
});

test("v3.0.0 modules: showModule shows one module, marks one link, writes the hash, scrolls up and focuses its title", () => {
  const scope = createScriptContext();
  const record = withModuleNavigation(scope);
  scope.evaluate('showModule("qualitySection")');
  assert.deepEqual(record.shown(), ["qualitySection"]);
  assert.deepEqual(record.inert(), MODULE_SECTION_IDS.filter((id) => id !== "qualitySection"));
  assert.deepEqual(record.hiddenClass(), []);
  assert.deepEqual(record.current(), ["#qualitySection"]);
  assert.deepEqual(record.urls, ["#qualitySection"]);
  assert.deepEqual(record.scrolls, [[{ top: 0, behavior: "instant" }]]);
  assert.deepEqual(record.focused, ["qualitySection"]);
  // Not a module: nothing changes.
  record.clear();
  scope.evaluate('showModule("dropZone")');
  assert.deepEqual(record.shown(), ["qualitySection"]);
  assert.deepEqual([record.urls, record.scrolls, record.focused], [[], [], []]);
});

test("v3.0.0 modules: an import shows the overview, scrolls up and focuses its title; a reset goes back quietly", () => {
  const scope = createChartContext();
  const record = withModuleNavigation(scope);
  scope.evaluate('showModule("exportSection")');
  record.clear();
  commitGuardRows(scope, ["城市", "金额"], Array.from({ length: 12 }, (_, index) => ({ 城市: index % 2 ? "上海" : "北京", 金额: String(index) })), "cities.csv");
  assert.deepEqual(record.shown(), ["overviewSection"]);
  assert.deepEqual(record.current(), ["#overviewSection"]);
  assert.deepEqual(record.focused, ["overviewSection"]);
  assert.deepEqual(record.scrolls, [[{ top: 0, behavior: "instant" }]]);
  scope.evaluate('showModule("chartsSection")');
  record.clear();
  scope.evaluate("resetAnalysisState()");
  assert.deepEqual(record.shown(), ["overviewSection"]);
  assert.deepEqual(record.current(), ["#overviewSection"]);
  assert.deepEqual([record.urls, record.scrolls, record.focused], [[], [], []]);
});

test("v3.0.0 modules: field configuration, placeholder confirmations and filters stay in the current module", async () => {
  const orders = await importCsvFile(readExampleCsv("sample-orders.csv"), "sample-orders.csv");
  const record = withModuleNavigation(orders);
  orders.evaluate('showModule("qualitySection")');
  record.clear();
  orders.evaluate('stageFieldTypeChange("客户类型", "id"); applyFieldConfiguration()');
  assert.equal(orders.evaluate('getFieldProfile("客户类型").typeKey'), "id");
  orders.evaluate("confirmAllCommonPlaceholders()");
  assert.ok(orders.evaluate("Object.keys(state.placeholderConfirmations).length") > 0);
  orders.evaluate(GUARD_EAST_FILTER);
  assert.deepEqual(record.shown(), ["qualitySection"]);
  assert.deepEqual(record.current(), ["#qualitySection"]);
  assert.deepEqual([record.urls, record.scrolls], [[], []]);
  assert.ok(!record.focused.some((id) => MODULE_SECTION_IDS.includes(id)), JSON.stringify(record.focused));
});

// v3.0.0 commit 2: top bar and sidebar.
test("v3.0.0 shell: the top bar holds the toggle, the brand, the navigation and 换一个文件; no section labels", () => {
  const html = indexHtmlSource();
  const header = html.slice(html.indexOf('<header class="app-header">'), html.indexOf("</header>", html.indexOf('<header class="app-header">')));
  assertInOrder(header, ['<button id="sidebarToggle"', "<strong>表格体检</strong>", '<nav id="moduleNav"', '<a class="nav-new-analysis secondary-button" href="#dropZone">换一个文件</a>']);
  assert.match(header, /<button id="sidebarToggle" class="[^"]*" type="button" aria-controls="appSidebar" aria-expanded="false" aria-label="文件和字段">/);
  // v3.0.0 look2: no privacy badge in the top bar. Outside the export module (whose notes say how
  // exporting works) the page states it twice: under the upload title and at the foot of the sidebar.
  assert.doesNotMatch(header, /数据仅在浏览器本地处理/);
  const exportStart = html.indexOf('<section id="exportSection"');
  const outsideExport = html.slice(0, exportStart) + html.slice(html.indexOf("<footer", exportStart));
  assert.doesNotMatch(outsideExport.replace(/<p class="(?:sidebar-privacy|product-description)">[^<]*<\/p>/g, ""), /不会上传|不会离开|仅在浏览器本地|本地处理/);
  assert.deepEqual(outsideExport.match(/<p class="[^"]+">[^<]*(?:不会上传|不会离开)[^<]*<\/p>/g),
    ['<p class="sidebar-privacy">表格只在这台电脑的浏览器里处理，不会上传。</p>', '<p class="product-description">不用安装，表格不会离开这台电脑。</p>']);
  assert.doesNotMatch(html, /class="header-cta"|上传文件<\/a>|上传新文件/);
  assert.doesNotMatch(html, /class="section-label"/);
  assert.doesNotMatch(html, /class="report-header"|class="results-nav"/);
  assert.match(html, /<h2 id="reportTitle" class="visually-hidden" tabindex="-1">数据分析工作区<\/h2>/);
  assert.match(html, /<section id="results" class="results hidden" aria-labelledby="reportTitle">/);
});

test("v3.0.0 shell: the sidebar's static parts, and one 换一个文件 inside it for phones", () => {
  const html = indexHtmlSource();
  const start = html.indexOf('<aside id="appSidebar"');
  assert.ok(start >= 0);
  const sidebar = html.slice(start, html.indexOf("</aside>", start));
  assertInOrder(sidebar, [
    '<h2 id="sidebarTitle" class="sidebar-title">文件和字段</h2>',
    '<button id="sidebarClose"',
    '<p id="sourceName" class="source-name"></p>',
    'id="sidebarFileSize"',
    '<a class="nav-new-analysis secondary-button sidebar-new-analysis" href="#dropZone">换一个文件</a>',
    '<button id="sidebarFieldsToggle"',
    '<ul id="sidebarFieldList"',
    "当前状态",
    "<dt>筛选</dt>",
    '<div id="sidebarHeaderRow" class="hidden"><dt>表头</dt>',
    '<div id="sidebarPlaceholderRow" class="hidden"><dt>已按空值处理</dt>',
    "表格只在这台电脑的浏览器里处理，不会上传。"
  ]);
  assert.match(html, /<div id="sidebarBackdrop" class="sidebar-backdrop" hidden><\/div>/);
  assert.equal((html.match(/class="nav-new-analysis /g) || []).length, 2);
});

test("v3.0.0 shell: the analysis range sits by four module titles; only the overview's carries the ids", () => {
  const html = indexHtmlSource();
  const heading = (id) => {
    const section = html.slice(html.indexOf(`<section id="${id}"`));
    return section.slice(0, section.indexOf("</h2>") + 600);
  };
  assert.match(heading("overviewSection"), /<p id="globalFilterStatus" class="filter-scope" aria-live="polite">当前分析范围：<span id="globalFilterSummary">尚未加载数据<\/span>\s*<button id="globalClearFilters" class="text-button hidden" type="button">清除筛选<\/button><\/p>/);
  ["statisticsSection", "chartsSection", "customSection"].forEach((id) => {
    assert.match(heading(id), /<p class="filter-scope">当前分析范围：<span class="filter-scope-text">尚未加载数据<\/span>\s*<button class="text-button filter-scope-clear hidden" type="button">清除筛选<\/button><\/p>/, id);
  });
  // With the colon: the filter module's own title is 设置当前分析范围.
  ["schemaSection", "qualitySection", "filterSection", "exportSection"].forEach((id) => assert.doesNotMatch(heading(id), /当前分析范围：/, id));
  assert.equal((html.match(/id="globalClearFilters"/g) || []).length, 1);
  assert.match(repoFile("script.js"), /document\.querySelectorAll\("\.filter-scope-clear"\)\.forEach\(\(button\) => button\.addEventListener\("click", clearAllFilters\)\)/);
});

// The three range sentences without ids: the fake page's querySelectorAll returns [], so this
// gives it three of each.
function withFilterScopes(scope) {
  const make = () => Array.from({ length: 3 }, () => {
    const classes = new Set(["hidden"]);
    return { textContent: "", classList: { toggle(name, force) { if (force) classes.add(name); else classes.delete(name); return classes.has(name); }, contains: (name) => classes.has(name) } };
  });
  const parts = { ".filter-scope": make(), ".filter-scope-text": make(), ".filter-scope-clear": make() };
  const document = scope.context.document;
  const originalQuery = document.querySelectorAll;
  document.querySelectorAll = (selector) => parts[selector] || originalQuery.call(document, selector);
  return parts;
}

test("v3.0.0 shell: the range sentences, their 清除筛选 buttons and the sidebar's 筛选 row follow #globalFilterSummary", async () => {
  const orders = await importCsvFile(readExampleCsv("sample-orders.csv"), "sample-orders.csv");
  const parts = withFilterScopes(orders);
  const summary = () => orders.elements.get("globalFilterSummary").textContent;
  orders.evaluate(GUARD_EAST_FILTER);
  assert.equal(summary(), "已筛选 1 项 · 10 / 42 行");
  assert.deepEqual(parts[".filter-scope-text"].map((part) => part.textContent), [summary(), summary(), summary()]);
  assert.deepEqual(parts[".filter-scope-clear"].map((part) => part.classList.contains("hidden")), [false, false, false]);
  assert.deepEqual(parts[".filter-scope"].map((part) => part.classList.contains("active")), [true, true, true]);
  assert.equal(orders.elements.get("sidebarFilterSummary").textContent, summary());
  orders.evaluate("clearAllFilters()");
  assert.equal(summary(), "全部 42 行");
  assert.deepEqual(parts[".filter-scope-text"].map((part) => part.textContent), ["全部 42 行", "全部 42 行", "全部 42 行"]);
  assert.deepEqual(parts[".filter-scope-clear"].map((part) => part.classList.contains("hidden")), [true, true, true]);
  assert.equal(orders.elements.get("sidebarFilterSummary").textContent, "全部 42 行");
});

const sidebarItems = (scope) => [...scope.elements.get("sidebarFieldList").innerHTML.matchAll(/<li class="([^"]*)" title="([^"]*)"><span class="sidebar-type">([^<]*)<\/span><span class="sidebar-field-name">([^<]*)<\/span>(?:<span class="sidebar-unit">([^<]*)<\/span>)?<\/li>/g)]
  .map((match) => [match[1], match[2], match[3], match[4], match[5] || ""]);

test("v3.0.0 shell: the sidebar lists the file, its size and every field with its applied type, and the state rows", async () => {
  const orders = await importCsvFile(readExampleCsv("sample-orders.csv"), "sample-orders.csv");
  const text = (id) => orders.elements.get(id).textContent;
  const hidden = (id) => orders.elements.get(id).classList.contains("hidden");
  assert.equal(text("sourceName"), "sample-orders.csv");
  assert.equal(orders.elements.get("sourceName").title, "sample-orders.csv");
  assert.equal(text("topbarFileName"), "sample-orders.csv");
  assert.equal(text("sidebarFileSize"), "42 行 × 9 列");
  assert.equal(text("sidebarFieldCount"), "9");
  assert.deepEqual(sidebarItems(orders), [
    ["sidebar-field", "订单编号", "#", "订单编号", ""],
    ["sidebar-field", "下单日期", "日", "下单日期", ""],
    ["sidebar-field", "地区", "Aa", "地区", ""],
    ["sidebar-field", "品类", "Aa", "品类", ""],
    ["sidebar-field", "产品", "Aa", "产品", ""],
    ["sidebar-field", "客户类型", "Aa", "客户类型", ""],
    ["sidebar-field", "数量", "12", "数量", ""],
    ["sidebar-field", "金额", "12", "金额", "元"],
    ["sidebar-field", "销售代表", "Aa", "销售代表", ""]
  ]);
  assert.equal(hidden("sidebarHeaderRow"), false);
  assert.equal(text("sidebarHeaderValue"), "第 2 行");
  assert.equal(hidden("sidebarPlaceholderRow"), true);
  assert.equal(text("sidebarFilterSummary"), "全部 42 行");

  // Confirmations count entries; an ignored field greys out with a dash for its mark.
  orders.evaluate("confirmAllCommonPlaceholders()");
  assert.equal(hidden("sidebarPlaceholderRow"), false);
  assert.equal(text("sidebarPlaceholderValue"), "2 处");
  orders.evaluate('stageFieldTypeChange("销售代表", "ignore"); applyFieldConfiguration()');
  assert.deepEqual(sidebarItems(orders).at(-1), ["sidebar-field is-ignored", "销售代表", "—", "销售代表", ""]);
  assert.equal(text("sidebarPlaceholderValue"), "2 处");

  orders.evaluate("resetAnalysisState()");
  ["sourceName", "topbarFileName", "sidebarFileSize", "sidebarFieldCount", "sidebarHeaderValue", "sidebarPlaceholderValue"].forEach((id) => assert.equal(text(id), "", id));
  assert.equal(orders.elements.get("sidebarFieldList").innerHTML, "");
  assert.equal(hidden("sidebarHeaderRow"), true);
  assert.equal(hidden("sidebarPlaceholderRow"), true);
});

test("v3.0.0 shell: an Excel sheet joins the file name; a header in row 1 shows no header row; field names are escaped", () => {
  const scope = createChartContext();
  put(scope.context, "__sidebarRows", Array.from({ length: 12 }, (_, index) => ({ "字段<script>": index % 2 ? "x" : "y" })));
  scope.evaluate('commitTabularData(["字段<script>"], __sidebarRows, "a<b>.xlsx", beginImport(), [], "表一")');
  assert.equal(scope.elements.get("sourceName").textContent, "a<b>.xlsx · 表一");
  assert.equal(scope.elements.get("sourceName").title, "a<b>.xlsx · 表一");
  assert.equal(scope.elements.get("topbarFileName").textContent, "a<b>.xlsx · 表一");
  assert.equal(scope.elements.get("sidebarHeaderRow").classList.contains("hidden"), true);
  const list = scope.elements.get("sidebarFieldList").innerHTML;
  assert.match(list, /字段&lt;script&gt;/);
  assert.doesNotMatch(list, /字段<script>/);
});

test("v3.0.0 shell: the sidebar opens by width until the user chooses; Escape closes an overlay and gives the focus back", () => {
  const scope = createChartContext();
  let wide = true;
  scope.context.window.matchMedia = () => ({ matches: wide, addEventListener() {} });
  const body = scope.context.document.body;
  const toggle = scope.elements.get("sidebarToggle") || scope.context.document.getElementById("sidebarToggle");
  const attributes = {};
  toggle.setAttribute = (name, value) => { attributes[name] = String(value); };
  const focused = [];
  toggle.focus = () => focused.push("toggle");
  scope.context.document.getElementById("sidebarClose").focus = () => focused.push("close");
  const backdrop = scope.context.document.getElementById("sidebarBackdrop");
  commitGuardRows(scope, ["城市"], Array.from({ length: 12 }, (_, index) => ({ 城市: index % 2 ? "上海" : "北京" })), "cities.csv");
  assert.equal(attributes["aria-expanded"], "true");
  scope.evaluate("toggleSidebar()");
  assert.equal(body.classList.contains("sidebar-closed"), true);
  assert.equal(attributes["aria-expanded"], "false");
  scope.evaluate("resetAnalysisState()");
  assert.equal(body.classList.contains("sidebar-closed"), false);
  assert.equal(body.classList.contains("sidebar-open"), false);

  // Narrow: closed by default; opening moves the focus in, Escape closes and returns it.
  wide = false;
  commitGuardRows(scope, ["城市"], Array.from({ length: 12 }, (_, index) => ({ 城市: index % 2 ? "上海" : "北京" })), "cities.csv");
  assert.equal(attributes["aria-expanded"], "false");
  scope.evaluate("toggleSidebar()");
  assert.equal(body.classList.contains("sidebar-open"), true);
  assert.equal(backdrop.hidden, false);
  assert.deepEqual(focused, ["close"]);
  let prevented = 0;
  put(scope.context, "__escape", { key: "Escape", preventDefault() { prevented += 1; } });
  scope.evaluate("handleSidebarKeydown(__escape)");
  assert.equal(body.classList.contains("sidebar-open"), false);
  assert.equal(backdrop.hidden, true);
  assert.deepEqual(focused, ["close", "toggle"]);
  assert.equal(prevented, 1);
  // Escape with the sidebar closed does nothing.
  scope.evaluate("handleSidebarKeydown(__escape)");
  assert.equal(prevented, 1);
});

test("v3.0.0 shell: the sidebar controls are bound once each in script.js", () => {
  const script = repoFile("script.js");
  [
    /dom\.sidebarToggle\.addEventListener\("click", toggleSidebar\)/,
    /dom\.sidebarClose\.addEventListener\("click", \(\) => closeSidebar\(\)\)/,
    /dom\.sidebarBackdrop\.addEventListener\("click", \(\) => closeSidebar\(\)\)/,
    /dom\.sidebarFieldsToggle\.addEventListener\("click", toggleSidebarFields\)/,
    /dom\.appSidebar\.addEventListener\("keydown", handleSidebarKeydown\)/
  ].forEach((pattern) => assert.equal((script.match(new RegExp(pattern.source, "g")) || []).length, 1, pattern.source));
});

// v3.0.0 commit 3: the import message with notes can fold to one line. #statusPanel itself
// keeps its text, classes and role: the fold lives on the wrapper #statusNote.
test("v3.0.0 status note: a wrapper around the unchanged #statusPanel, folded by CSS that keeps the text for screen readers", () => {
  const html = indexHtmlSource();
  const start = html.indexOf('<div id="statusNote" class="status-note">');
  assert.ok(start >= 0);
  const note = html.slice(start, html.indexOf('<section id="sheetSelectionPanel"'));
  assertInOrder(note, [
    '<div class="status-note-bar">',
    '<p id="statusNoteSummary" class="status-note-summary" aria-hidden="true"></p>',
    '<button id="statusNoteToggle" class="text-button status-note-toggle" type="button" aria-controls="statusPanel" aria-expanded="true">收起</button>',
    '<section id="statusPanel" class="status-panel hidden" role="status" aria-live="polite" aria-atomic="true"></section>'
  ]);
  const css = repoFile("style.css");
  const folded = css.match(/\n\.status-note\.is-collapsed #statusPanel \{([^}]*)\}/);
  assert.ok(folded, "folded rule");
  assert.match(folded[1], /position: absolute;/);
  assert.match(folded[1], /clip-path: inset\(50%\);/);
  assert.doesNotMatch(folded[1], /display:|visibility:/);
  // Off the overview only a success hides: warnings and errors (the report export's) show
  // on every module.
  assert.match(css, /body\.has-results:not\(\[data-module="overviewSection"\]\) #statusNote\[data-tone="success"\] \{\s*display: none;\s*\}/);
  assert.equal((css.match(/#statusNote(?:\[[^\]]*\])* \{\s*display: none;/g) || []).length, 1);
  assert.doesNotMatch(css, /#statusNote \{\s*display: none;/);
  assert.match(css, /body\.has-results #statusPanel\.success:not\(\.has-note\)\s*\{\s*display:\s*none;\s*\}/);
  const script = repoFile("script.js");
  assert.equal((script.match(/dom\.statusNoteToggle\.addEventListener\("click", toggleStatusNote\)/g) || []).length, 1);
  const workspace = script.slice(script.indexOf("function setupWorkspaceNavigation("));
  assert.match(workspace, /resetStatusNote\(\);/);
});

test("v3.0.0 status note: a success with notes can fold to 〈file〉已载入，有 N 条说明; #statusPanel never changes", async () => {
  const orders = await importCsvFile(readExampleCsv("sample-orders.csv"), "sample-orders.csv");
  const panel = orders.elements.get("statusPanel");
  const note = orders.elements.get("statusNote");
  const toggle = orders.elements.get("statusNoteToggle");
  const attributes = {};
  toggle.setAttribute = (name, value) => { attributes[name] = String(value); };
  const before = { className: panel.className, text: panel.textContent, role: panel.role };
  assert.equal(before.className, "status-panel success has-note");
  assert.match(before.text, /^sample-orders\.csv 已成功载入，分析结果已生成。表头在第 2 行/);
  assert.equal(note.classList.contains("can-collapse"), true);
  assert.equal(note.classList.contains("is-collapsed"), false);
  assert.equal(orders.elements.get("statusNoteSummary").textContent, "sample-orders.csv 已载入，有 2 条说明");

  orders.evaluate("toggleStatusNote()");
  assert.equal(note.classList.contains("is-collapsed"), true);
  assert.equal(attributes["aria-expanded"], "false");
  assert.equal(toggle.textContent, "看说明");
  assert.deepEqual({ className: panel.className, text: panel.textContent, role: panel.role }, before);
  orders.evaluate("toggleStatusNote()");
  assert.equal(note.classList.contains("is-collapsed"), false);
  assert.equal(attributes["aria-expanded"], "true");
  assert.equal(toggle.textContent, "收起");

  // A later message replaces the import's: nothing left to fold.
  orders.evaluate('toggleStatusNote(); setStatus("HTML 报告已生成：x.html", "success")');
  assert.equal(note.classList.contains("can-collapse"), false);
  assert.equal(note.classList.contains("is-collapsed"), false);
  assert.equal(attributes["aria-expanded"], "true");
  assert.equal(toggle.textContent, "收起");

  // A reset clears the summary.
  orders.evaluate("resetAnalysisState()");
  assert.equal(orders.elements.get("statusNoteSummary").textContent, "");
  assert.equal(note.classList.contains("can-collapse"), false);
});

test("v3.0.0 status note: N counts the notes that are there; a warning never folds", async () => {
  const errors = await importStage3Fixture("公式错误.xlsx");
  assert.equal(errors.elements.get("statusNote").classList.contains("can-collapse"), true);
  assert.equal(errors.elements.get("statusNoteSummary").textContent, "公式错误.xlsx · 公式错误 已载入，有 1 条说明");
  const warned = await importCsvFile("名,额,备注\r\nA,1,x\r\nB\r\nC,3,\r\nD,4,x\r\nE,5,y\r\n", "warned.csv");
  assert.equal(warned.elements.get("statusPanel").className, "status-panel warning");
  assert.equal(warned.elements.get("statusNote").classList.contains("can-collapse"), false);
  // A plain success has no note: no summary, nothing to fold.
  const plain = createChartContext();
  commitGuardRows(plain, ["城市"], Array.from({ length: 12 }, (_, index) => ({ 城市: index % 2 ? "上海" : "北京" })), "plain.csv");
  assert.equal(plain.elements.get("statusPanel").className, "status-panel success");
  assert.equal(plain.elements.get("statusNote").classList.contains("can-collapse"), false);
  assert.equal(plain.elements.get("statusNoteSummary").textContent, "");
});

test("v3.0.0 status note: showModule marks the current module on <body>, so the note shows only on the overview", () => {
  const scope = createScriptContext();
  const body = scope.context.document.body;
  scope.evaluate('showModule("chartsSection")');
  assert.equal(body.dataset.module, "chartsSection");
  scope.evaluate('showModule("overviewSection", { quiet: true })');
  assert.equal(body.dataset.module, "overviewSection");
});

test("v3.0.0 status note: on the export module, setStatus marks warnings and errors so they show; #statusPanel is as setStatus wrote it", async () => {
  const orders = await importCsvFile(readExampleCsv("sample-orders.csv"), "sample-orders.csv");
  const note = orders.elements.get("statusNote");
  const panel = orders.elements.get("statusPanel");
  assert.equal(note.dataset.tone, "success");
  orders.evaluate('showModule("exportSection")');
  assert.equal(orders.context.document.body.dataset.module, "exportSection");
  // report.js:433 and report.js:464 write these two.
  orders.evaluate('setStatus("请先完成数据分析，再导出报告。", "warning")');
  assert.equal(note.dataset.tone, "warning");
  assert.equal(panel.className, "status-panel warning");
  assert.equal(panel.textContent, "请先完成数据分析，再导出报告。");
  orders.evaluate('setStatus("报告导出失败：x", "error")');
  assert.equal(note.dataset.tone, "error");
  assert.equal(panel.className, "status-panel error");
  assert.equal(panel.textContent, "报告导出失败：x");
  orders.evaluate('setStatus("HTML 报告已生成：x.html", "success")');
  assert.equal(note.dataset.tone, "success");
  orders.evaluate('setStatus("正在读取 x.csv ...")');
  assert.equal(note.dataset.tone, "loading");
  orders.evaluate('showError("文件中没有可分析的数据。", state.activeImportId)');
  assert.equal(note.dataset.tone, "error");
  assert.equal(panel.className, "status-panel error");
});

// v3.0.0 commit 4: the upload page and the way back to the last analysis.
test("v3.0.0 upload page: the new title, the drop area and the encoding line; the old intro is gone", () => {
  const html = indexHtmlSource();
  const hero = html.slice(html.indexOf('<div class="upload-hero">'), html.indexOf('<div id="statusNote"'));
  assertInOrder(hero, [
    '<h1 id="heroTitle">上传表格，先做一次体检</h1>',
    '<p class="product-description">不用安装，表格不会离开这台电脑。</p>',
    '<button id="returnToAnalysis" class="secondary-button return-to-analysis hidden" type="button"></button>',
    '<section id="dropZone" class="upload-panel" aria-labelledby="uploadTitle">',
    '<h2 id="uploadTitle">把文件拖到这里</h2>',
    '<p class="upload-limits">CSV、XLSX、XLS，单文件最大 25 MB，最多 100,000 行、200 列</p>',
    '<input id="csvFile" type="file"',
    '<label class="file-button" for="csvFile">',
    "选择数据文件",
    '<button id="loadSample" class="secondary-button" type="button">',
    "使用示例数据",
    '<label class="encoding-control" for="fileEncoding">',
    "<span>CSV 文件编码</span>",
    '<select id="fileEncoding">',
    '<p class="upload-encoding-hint">不确定时保持“自动识别”即可</p>',
    '<a href="examples/sample-orders.csv" download>下载示例 CSV</a>'
  ]);
  // The visually hidden file input stays right before its label (focus ring, style.css).
  assert.match(hero, /<input id="csvFile" type="file"[^>]*>\s*<label class="file-button" for="csvFile">/);
  assert.doesNotMatch(html, /class="eyebrow"|浏览器端 CSV \/ Excel 数据分析工具|快速看懂你的数据|product-capabilities|主要功能|上传数据文件<\/h2>|拖放文件到此处/);
});

// The six rows of examples/sample-orders.csv shown on the upload page, as they are in the file.
test("v3.0.0 upload page: the sample table shows rows 10, 12, 20, 21, 29 and 44 of sample-orders.csv with their marks", () => {
  const html = indexHtmlSource();
  const start = html.indexOf('<section class="sample-preview" aria-labelledby="samplePreviewTitle">');
  assert.ok(start >= 0);
  assert.ok(start > html.indexOf('<div id="statusNote"'), "below the upload card and the status");
  const section = html.slice(start, html.indexOf("</section>", start));
  assertInOrder(section, ['<h2 id="samplePreviewTitle">会帮你查这些</h2>', '<span class="sample-preview-note">示例数据里的几行</span>']);
  assert.match(section, /<thead><tr><th>行<\/th><th>订单编号<\/th><th>下单日期<\/th><th>地区<\/th><th>金额<\/th><th>会标出来<\/th><\/tr><\/thead>/);
  const rows = [...section.matchAll(/<tr>(<td[\s\S]*?)<\/tr>/g)].map((match) => [...match[1].matchAll(/<td(?: class="([^"]*)")?>([\s\S]*?)<\/td>/g)]
    .map((cell) => [cell[1] || "", cell[2].replace(/<span class="visible-space" aria-hidden="true">␣<\/span>/g, " ")]));
  assert.deepEqual(rows.map((cells) => cells[0][1]), ["10", "12", "20", "21", "29", "44"]);
  const lines = readExampleCsv("sample-orders.csv").replace(/^\uFEFF/, "").split(/\r?\n/);
  const header = lines[1].split(",");
  rows.forEach((cells) => {
    const values = lines[Number(cells[0][1]) - 1].split(",");
    ["订单编号", "下单日期", "地区", "金额"].forEach((field, index) => assert.equal(cells[index + 1][1], values[header.indexOf(field)], `${cells[0][1]} ${field}`));
  });
  const marks = rows.map((cells) => cells.slice(1, 5).map(([className], index) => (className ? `${["订单编号", "下单日期", "地区", "金额"][index]}:${className}` : "")).filter(Boolean));
  assert.deepEqual(marks, [
    ["地区:mark mark-look-alike"],
    ["下单日期:mark mark-date"],
    ["订单编号:mark mark-duplicate"],
    ["订单编号:mark mark-duplicate"],
    ["金额:mark mark-placeholder"],
    ["订单编号:mark mark-total"]
  ]);
  assert.deepEqual(rows.map((cells) => cells[5]), [
    ["mark-note mark-look-alike", "后面多了一个空格"],
    ["mark-note mark-date", "日期写法和别的行不一样"],
    ["mark-note", ""],
    ["mark-note mark-duplicate", "和上一行完全相同"],
    ["mark-note mark-placeholder", "像是空值的写法，认不出是数字"],
    ["mark-note mark-total", "合计行混在数据里"]
  ]);
});

test("v3.0.0 upload page: an error shows between the title and the card; loading text stays below the card", () => {
  const css = repoFile("style.css");
  assert.match(css, /body:not\(\.has-results\) \.upload-hero \{\s*order: -2;\s*\}/);
  assert.match(css, /body:not\(\.has-results\) #statusNote\[data-tone="error"\] \{\s*order: -1;\s*\}/);
  assert.match(css, /body\.has-results \.upload-hero,\s*body\.has-results \.sample-preview \{\s*display: none;\s*\}/);
  const html = indexHtmlSource();
  // The sheet choice sits right under the card, above the sample rows.
  assertInOrder(html, ['<div class="upload-hero">', '<section id="dropZone"', '<div id="statusNote"', '<section id="sheetSelectionPanel"', '<section class="sample-preview"', '<section id="results"']);
  assert.match(css, /\.sheet-selection-panel \{\s*grid-template-columns: 1fr;\s*\}/);
  // The upload page's top bar has no sidebar toggle: .icon-button (inline-grid) comes later in
  // the file, so the hiding rule must be more specific than a bare .sidebar-toggle.
  assert.match(css, /body:not\(\.has-results\) \.sidebar-toggle,/);
});

// What 换一个文件 does, and the way back. The fake page ignores addEventListener, so these call
// the functions script.js binds.
function withRenderCounts(scope) {
  scope.evaluate(`(() => {
    globalThis.__renderCalls = 0;
    ["renderOverview", "renderPreview", "renderSchema", "renderSidebar", "renderQuality", "renderNumericStats", "renderCategoryStats", "renderCharts", "renderChartInsights", "populateControls", "initializeFilterControls"].forEach((name) => {
      const original = globalThis[name];
      globalThis[name] = function (...args) { __renderCalls += 1; return original.apply(this, args); };
    });
  })()`);
  return () => scope.evaluate("__renderCalls");
}

test("v3.0.0 way back: 换一个文件 keeps the analysis, offers 回到「名」的分析, and coming back restores the module and the message without recalculating", async () => {
  const orders = await importCsvFile(readExampleCsv("sample-orders.csv"), "sample-orders.csv");
  const record = withModuleNavigation(orders);
  const entry = orders.elements.get("returnToAnalysis");
  const focused = [];
  entry.focus = () => focused.push("entry");
  orders.elements.get("csvFile").focus = () => focused.push("file");
  const body = orders.context.document.body;
  const panel = orders.elements.get("statusPanel");
  const imported = { className: panel.className, text: panel.textContent, summary: orders.elements.get("statusNoteSummary").textContent };
  assert.equal(imported.summary, "sample-orders.csv 已载入，有 2 条说明");
  orders.evaluate('showModule("chartsSection")');
  const renders = withRenderCounts(orders);

  orders.evaluate("showUploadPage()");
  assert.equal(body.classList.contains("has-results"), false);
  assert.equal(orders.elements.get("results").classList.contains("hidden"), true);
  assert.equal(panel.className, "status-panel hidden");
  assert.equal(panel.textContent, "");
  assert.equal(entry.classList.contains("hidden"), false);
  assert.equal(entry.textContent, "回到「sample-orders.csv」的分析");
  assert.deepEqual(focused, ["entry"]);
  assert.equal(orders.evaluate("state.rows.length"), 42);

  record.clear();
  orders.evaluate("returnToAnalysis()");
  assert.equal(body.classList.contains("has-results"), true);
  assert.equal(orders.elements.get("results").classList.contains("hidden"), false);
  assert.equal(entry.classList.contains("hidden"), true);
  assert.deepEqual(record.shown(), ["chartsSection"]);
  assert.deepEqual(record.focused, ["chartsSection"]);
  assert.deepEqual({ className: panel.className, text: panel.textContent, summary: orders.elements.get("statusNoteSummary").textContent }, imported);
  assert.equal(orders.elements.get("statusNote").classList.contains("can-collapse"), true);
  assert.equal(renders(), 0);
});

test("v3.0.0 way back: no entry without an analysis, while a workbook waits for its sheet, or after a failed import", async () => {
  const scope = createChartContext();
  const entry = scope.elements.get("returnToAnalysis") || scope.context.document.getElementById("returnToAnalysis");
  const focused = [];
  scope.context.document.getElementById("csvFile").focus = () => focused.push("file");
  entry.focus = () => focused.push("entry");
  scope.evaluate("showUploadPage()");
  assert.equal(entry.classList.contains("hidden"), true);
  assert.deepEqual(focused, ["file"]);

  // An Excel sheet name joins the file name.
  commitGuardRows(scope, ["城市"], Array.from({ length: 12 }, (_, index) => ({ 城市: index % 2 ? "上海" : "北京" })), "a.xlsx");
  scope.evaluate('state.sourceSheetName = "表一"; showUploadPage()');
  assert.equal(entry.textContent, "回到「a.xlsx · 表一」的分析");

  // A new import begins: the analysis and the recorded message are gone.
  scope.evaluate("beginImport()");
  assert.equal(entry.classList.contains("hidden"), true);
  assert.equal(scope.evaluate("state.importStatus"), null);
  scope.evaluate("showUploadPage()");
  assert.equal(entry.classList.contains("hidden"), true);

  // A workbook waiting for its sheet.
  commitGuardRows(scope, ["城市"], Array.from({ length: 12 }, (_, index) => ({ 城市: index % 2 ? "上海" : "北京" })), "b.csv");
  scope.evaluate('state.pendingExcel = { fileName: "c.xlsx" }; showUploadPage()');
  assert.equal(entry.classList.contains("hidden"), true);

  // A failed import.
  commitGuardRows(scope, ["城市"], Array.from({ length: 12 }, (_, index) => ({ 城市: index % 2 ? "上海" : "北京" })), "d.csv");
  scope.evaluate('showError("文件中没有可分析的数据。", state.activeImportId); showUploadPage()');
  assert.equal(entry.classList.contains("hidden"), true);
  assert.equal(scope.evaluate("returnToAnalysis(); document.body.classList.contains('has-results')"), false);
});

test("v3.0.0 way back: 换一个文件 and the entry are bound in script.js", () => {
  const script = repoFile("script.js");
  assert.match(script, /document\.querySelectorAll\("\.nav-new-analysis"\)\.forEach\(\(link\) => \{\s*link\.addEventListener\("click", \(event\) => \{\s*event\.preventDefault\(\);\s*showUploadPage\(\);/);
  assert.equal((script.match(/dom\.returnToAnalysis\.addEventListener\("click", returnToAnalysis\)/g) || []).length, 1);
});

// v3.0.0 commit 5: the field table (F1) and its search and type filters (F2).
const schemaHeaders = (scope) => [...scope.elements.get("schemaTable").innerHTML.matchAll(/<th(?: class="[^"]*")?>([^<]*)<\/th>/g)].map((match) => match[1]);
const schemaRows = (scope) => [...scope.elements.get("schemaTable").innerHTML.matchAll(/<tr>([\s\S]*?)<\/tr>/g)].slice(1).map((match) => match[1]);
const schemaRowOf = (scope, field) => schemaRows(scope).find((row) => row.includes(`data-field-type="${field}"`));

test("v3.0.0 F1: the field table shows 字段名、自动类型、确认类型、非空值、唯一值、认不出、示例值, without 查看详情", async () => {
  const orders = await importCsvFile(readExampleCsv("sample-orders.csv"), "sample-orders.csv");
  assert.deepEqual(schemaHeaders(orders), ["字段名", "自动类型", "确认类型", "非空值", "唯一值", "认不出", "示例值"]);
  const html = orders.elements.get("schemaTable").innerHTML;
  assert.doesNotMatch(html, /查看详情|<details|字段详情/);
  const profiles = JSON.parse(orders.evaluate("JSON.stringify(state.profiles.map((profile) => [profile.field, profile.inferredTypeKey, profile.inferredType, profile.nonMissingCount, profile.uniqueCount, profile.conversionFailureCount]))"));
  assert.equal(schemaRows(orders).length, profiles.length);
  profiles.forEach(([field, typeKey, typeLabel, nonMissing, unique, failures]) => {
    const row = schemaRowOf(orders, field);
    const cells = [...row.matchAll(/<td(?: class="[^"]*")?>([\s\S]*?)<\/td>/g)].map((match) => match[1].trim());
    assert.equal(cells.length, 7, field);
    assert.equal(cells[0], `<span class="field-name" title="${field}">${field}</span>`);
    assert.equal(cells[1], `<span class="type-badge ${typeKey}">${typeLabel}</span>`);
    assert.match(cells[2], /^<select class="schema-type-select" data-field-type="/);
    assert.equal(cells[3], orders.evaluate(`formatInteger(${nonMissing})`));
    assert.equal(cells[4], orders.evaluate(`formatInteger(${unique})`));
    assert.equal((cells[5].match(/class="conversion-failure-count/g) || []).length, 1, field);
    assert.equal(/class="conversion-failure-count warning"/.test(cells[5]), failures > 0, field);
  });
  // 金额 has one N/A: 认不出 1 with the pink dot (.warning).
  assert.match(schemaRowOf(orders, "金额"), /<span class="conversion-failure-count warning"[^>]*>1<\/span>/);
});

test("v3.0.0 F1: a draft type marks its select and counts its 认不出 under the draft", async () => {
  const orders = await importCsvFile(readExampleCsv("sample-orders.csv"), "sample-orders.csv");
  orders.evaluate('stageFieldTypeChange("客户类型", "numeric"); renderSchema()');
  const changed = schemaRowOf(orders, "客户类型");
  assert.match(changed, /<select class="schema-type-select is-changed" data-field-type="客户类型"/);
  const failures = orders.evaluate('countFieldConversionFailures("客户类型", "numeric")');
  assert.ok(failures > 0);
  assert.match(changed, new RegExp(`<span class="conversion-failure-count warning"[^>]*>${failures}</span>`));
  orders.evaluate('stageFieldTypeChange("客户类型", "category"); renderSchema()');
  const back = schemaRowOf(orders, "客户类型");
  assert.match(back, /<select class="schema-type-select" data-field-type="客户类型"/);
  assert.match(back, /<span class="conversion-failure-count"[^>]*>—<\/span>/);
});

// 21 fields: 6 numeric (one named Price), 5 category, 5 date, 5 ID-like, 20 rows each.
function commitWideSchema(scope) {
  const fields = [
    "数值1", "数值2", "数值3", "数值4", "数值5", "Price",
    "类别1", "类别2", "类别3", "类别4", "类别5",
    "日期1", "日期2", "日期3", "日期4", "日期5",
    "代码1", "代码2", "代码3", "代码4", "代码5"
  ];
  const rows = Array.from({ length: 20 }, (_, index) => Object.fromEntries(fields.map((field, column) => [field,
    column < 6 ? String(index * 3 + column) : column < 11 ? ["甲", "乙", "丙"][(index + column) % 3]
      : column < 16 ? `2025-01-${String(index + 1).padStart(2, "0")}` : `K${column}-${index}`])));
  commitGuardRows(scope, fields, rows, "wide.csv");
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(state.profiles.map((profile) => profile.typeKey))")),
    [...Array(6).fill("numeric"), ...Array(5).fill("category"), ...Array(5).fill("date"), ...Array(5).fill("id")]);
}
const kindCounts = (scope) => JSON.parse(scope.evaluate("JSON.stringify(countSchemaKinds())"));
const shownFields = (scope) => JSON.parse(scope.evaluate("JSON.stringify(listSchemaFieldsShown())"));
const kindButtons = (scope) => [...scope.elements.get("schemaKinds").innerHTML.matchAll(/<button type="button" class="schema-kind" data-schema-kind="(\w+)" aria-pressed="(true|false)">([^<]*)<span class="schema-kind-count">([\d,]+)<\/span><\/button>/g)]
  .map((match) => [match[1], match[2], match[3].trim(), match[4]]);

test("v3.0.0 F2: more than 20 fields show the search and the type filters with counts; zero counts hide", () => {
  const sample = createChartContext();
  commitGuardRows(sample, ["城市"], Array.from({ length: 12 }, (_, index) => ({ 城市: index % 2 ? "上海" : "北京" })), "small.csv");
  assert.equal(sample.elements.get("schemaFilter").classList.contains("hidden"), true);

  const scope = createChartContext();
  commitWideSchema(scope);
  assert.equal(scope.elements.get("schemaFilter").classList.contains("hidden"), false);
  assert.equal(scope.elements.get("schemaSearch").placeholder, "搜索 21 个字段");
  assert.deepEqual(kindCounts(scope), { all: 21, numeric: 6, category: 5, date: 5, id: 5, ignore: 0, failing: 0, changed: 0 });
  assert.deepEqual(kindButtons(scope), [
    ["all", "true", "全部", "21"], ["numeric", "false", "数值", "6"], ["category", "false", "分类", "5"], ["date", "false", "日期", "5"], ["id", "false", "ID", "5"]
  ]);
  assert.equal(shownFields(scope).length, 21);
});

test("v3.0.0 F2: the search matches a part of the name, ignores case and edge spaces; nothing found says so", () => {
  const scope = createChartContext();
  commitWideSchema(scope);
  scope.evaluate('dom.schemaSearch.value = "  price "; handleSchemaSearch()');
  assert.deepEqual(shownFields(scope), ["Price"]);
  assert.equal(scope.elements.get("schemaFilterEmpty").classList.contains("hidden"), true);
  scope.evaluate('dom.schemaSearch.value = "数值"; handleSchemaSearch()');
  assert.deepEqual(shownFields(scope), ["数值1", "数值2", "数值3", "数值4", "数值5"]);
  scope.evaluate('dom.schemaSearch.value = "zzz"; handleSchemaSearch()');
  assert.deepEqual(shownFields(scope), []);
  // Its sentence is static (checked with the toolbar markup below).
  assert.equal(scope.elements.get("schemaFilterEmpty").classList.contains("hidden"), false);
  // Search and type combine.
  scope.evaluate('dom.schemaSearch.value = "1"; handleSchemaSearch(); selectSchemaKind("date")');
  assert.deepEqual(shownFields(scope), ["日期1"]);
});

test("v3.0.0 F2: the types follow the selects; a change while filtered only updates the counts, and the chosen type stays until the user moves on", () => {
  const scope = createChartContext();
  commitWideSchema(scope);
  scope.evaluate("selectSchemaKind(\"date\")");
  assert.deepEqual(shownFields(scope), ["日期1", "日期2", "日期3", "日期4", "日期5"]);
  scope.evaluate("globalThis.__applyCalls = 0; const __apply = applySchemaFilter; applySchemaFilter = function (...args) { __applyCalls += 1; return __apply.apply(this, args); }");
  scope.evaluate('stageFieldTypeChange("日期2", "ignore")');
  assert.equal(scope.evaluate("__applyCalls"), 0, "a select change does not filter again");
  assert.deepEqual(kindCounts(scope), { all: 21, numeric: 6, category: 5, date: 4, id: 5, ignore: 1, failing: 0, changed: 1 });
  assert.deepEqual(kindButtons(scope).map(([kind, pressed, , count]) => [kind, pressed, count]), [
    ["all", "false", "21"], ["numeric", "false", "6"], ["category", "false", "5"], ["date", "true", "4"], ["id", "false", "5"], ["ignore", "false", "1"], ["changed", "false", "1"]
  ]);
  // The chosen type at zero stays on screen until the user picks another one.
  scope.evaluate('selectSchemaKind("changed"); stageFieldTypeChange("日期2", "date")');
  assert.deepEqual(kindButtons(scope).map(([kind, pressed, , count]) => [kind, pressed, count]).filter(([kind]) => kind === "changed"), [["changed", "true", "0"]]);
  scope.evaluate('selectSchemaKind("all")');
  assert.deepEqual(kindButtons(scope).map(([kind]) => kind), ["all", "numeric", "category", "date", "id"]);
  // 有认不出的 follows the draft type.
  scope.evaluate('stageFieldTypeChange("类别1", "numeric")');
  assert.equal(kindCounts(scope).failing, 1);
  assert.deepEqual(kindButtons(scope).find(([kind]) => kind === "failing"), ["failing", "false", "有认不出的", "1"]);
});

test("v3.0.0 F2: applying keeps the search and the type (back to 全部 when it is empty); a new import clears them", () => {
  const scope = createChartContext();
  commitWideSchema(scope);
  scope.evaluate('dom.schemaSearch.value = "日期"; handleSchemaSearch(); selectSchemaKind("date")');
  scope.evaluate('stageFieldTypeChange("代码1", "category"); applyFieldConfiguration()');
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(state.schemaFilter)")), { query: "日期", kind: "date" });
  assert.equal(scope.elements.get("schemaSearch").value, "日期");
  assert.equal(shownFields(scope).length, 5);
  // Every date field becomes a category: the type has no field left and falls back to 全部.
  scope.evaluate('["日期1", "日期2", "日期3", "日期4", "日期5"].forEach((field) => stageFieldTypeChange(field, "category")); applyFieldConfiguration()');
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(state.schemaFilter)")), { query: "日期", kind: "all" });
  commitWideSchema(scope);
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(state.schemaFilter)")), { query: "", kind: "all" });
  assert.equal(scope.elements.get("schemaSearch").value, "");
  scope.evaluate("resetAnalysisState()");
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(state.schemaFilter)")), { query: "", kind: "all" });
});

test("v3.0.0 F2: the toolbar is static, bound once, and the table box is limited only above 640 px", () => {
  const html = indexHtmlSource();
  const schema = html.slice(html.indexOf('<section id="schemaSection"'), html.indexOf('<section id="qualitySection"'));
  assertInOrder(schema, [
    '<div id="schemaFilter" class="schema-filter hidden">',
    '<input id="schemaSearch" class="schema-search" type="search" autocomplete="off" aria-label="搜索字段">',
    '<div id="schemaKinds" class="schema-kinds" role="group" aria-label="按类型筛选字段"></div>',
    '<p id="schemaFilterEmpty" class="empty-state hidden">没有符合条件的字段。</p>',
    '<div class="table-wrap schema-table-wrap" tabindex="0" aria-label="字段配置表，可横向滚动">',
    '<table id="schemaTable"'
  ]);
  const script = repoFile("script.js");
  assert.equal((script.match(/dom\.schemaSearch\.addEventListener\("input", handleSchemaSearch\)/g) || []).length, 1);
  assert.equal((script.match(/dom\.schemaKinds\.addEventListener\("click", handleSchemaKindClick\)/g) || []).length, 1);
  const css = repoFile("style.css");
  assert.match(css, /@media \(min-width: 641px\) \{\s*\.schema-table-wrap \{\s*max-height: [^;]+;\s*overflow: auto;\s*\}\s*\.schema-table-wrap th \{\s*position: sticky;\s*top: 0;/);
  // Printed: no search or filters, no height limit, every field.
  const print = css.match(/@media print \{([\s\S]*?)\n\}/)[1];
  assert.match(print, /\.schema-filter,/);
  assert.match(print, /\.schema-table-wrap \{\s*max-height: none !important;/);
  assert.match(print, /\.schema-table-wrap tr\[hidden\] \{\s*display: table-row !important;/);
});

test("v3.0.0 F2: field names in the table and the kinds are escaped", () => {
  const scope = createChartContext();
  const fields = Array.from({ length: 21 }, (_, index) => (index === 0 ? "字段<script>" : `f${index}`));
  commitGuardRows(scope, fields, Array.from({ length: 12 }, (_, index) => Object.fromEntries(fields.map((field) => [field, index % 2 ? "x" : "y"]))), "escape.csv");
  const html = scope.elements.get("schemaTable").innerHTML;
  assert.match(html, /title="字段&lt;script&gt;">字段&lt;script&gt;<\/span>/);
  assert.doesNotMatch(html + scope.elements.get("schemaKinds").innerHTML, /字段<script>/);
});

// v3.0.0 commit 6 (F3): empty states as text. drawEmptyChart(canvas, text) keeps its name and
// arguments (basic:7002 and the commit-0 guard wrap it) but writes the sentence into a static
// element over the chart instead of painting it on the canvas.
const CHART_CANVAS_IDS = ["numericDistributionChart", "categoryTopChart", "dateTrendChart", "v2TopGroupChart", "v2SelectedTrendChart"];

// The fake page's elements have no id of their own; the browser's do.
function withChartCanvasIds(scope) {
  const painted = [];
  CHART_CANVAS_IDS.forEach((id) => {
    const canvas = scope.context.document.getElementById(id);
    canvas.id = id;
    canvas.getContext = () => ({ clearRect() {}, save() {}, restore() {}, fillText(text) { painted.push([id, text]); } });
  });
  return painted;
}
const chartEmpty = (scope, id) => {
  const element = scope.context.document.getElementById(`${id}Empty`);
  return { text: element.textContent, hidden: element.classList.contains("hidden") };
};

test("v3.0.0 F3: every chart frame has a static, hidden empty-state line after its canvas", () => {
  const html = indexHtmlSource();
  CHART_CANVAS_IDS.forEach((id) => {
    assert.match(html, new RegExp(`<div class="chart-frame"><canvas id="${id}" [^>]*></canvas><p id="${id}Empty" class="empty-state chart-empty hidden"></p></div>`), id);
  });
  const css = repoFile("style.css");
  assert.match(css, /\n\.chart-empty \{[^}]*position: absolute;/);
});

test("v3.0.0 F3: drawEmptyChart writes its sentence as text, word for word, and paints nothing; the next real chart clears it", () => {
  const scope = createChartContext();
  const painted = withChartCanvasIds(scope);
  commitGuardRows(scope, ["名称", "备注"], Array.from({ length: 20 }, (_, index) => ({ 名称: ["甲", "乙", "丙"][index % 3], 备注: index % 2 ? "x" : "y" })), "words.csv");
  assert.deepEqual(chartEmpty(scope, "numericDistributionChart"), { text: "没有可用的数值字段", hidden: false });
  assert.deepEqual(chartEmpty(scope, "dateTrendChart"), { text: "没有可用的日期字段", hidden: false });
  assert.deepEqual(chartEmpty(scope, "categoryTopChart"), { text: "", hidden: true });
  assert.deepEqual(painted, []);
  // A file with numbers and dates: the charts are drawn and the sentences go.
  commitGuardRows(scope, ["日期", "金额", "组"], Array.from({ length: 30 }, (_, index) => ({
    日期: `2025-01-${String(index + 1).padStart(2, "0")}`, 金额: String(100 + index * 10), 组: index % 2 ? "A" : "B"
  })), "numbers.csv");
  scope.evaluate('dom.numericChartField.value = "金额"; renderCharts()');
  assert.deepEqual(chartEmpty(scope, "numericDistributionChart"), { text: "", hidden: true });
  assert.deepEqual(chartEmpty(scope, "dateTrendChart"), { text: "", hidden: true });
  // Chart.js missing: the sentence goes into all three frames.
  scope.context.Chart = scope.context.window.Chart = null;
  scope.evaluate("renderCharts()");
  ["numericDistributionChart", "categoryTopChart", "dateTrendChart"].forEach((id) => {
    assert.deepEqual(chartEmpty(scope, id), { text: "图表组件加载失败，请刷新页面重试", hidden: false }, id);
  });
  // A reset clears all five.
  scope.evaluate("resetAnalysisState()");
  CHART_CANVAS_IDS.forEach((id) => assert.deepEqual(chartEmpty(scope, id), { text: "", hidden: true }, id));
  assert.deepEqual(painted, []);
});

test("v3.0.0 F3: the custom analysis' empty trend says so in its frame, and a new analysis clears it", () => {
  const trend = createChartContext();
  withChartCanvasIds(trend);
  commitGuardRows(trend, ["日期", "金额", "组"], Array.from({ length: 30 }, (_, index) => ({
    日期: index < 25 ? `2025-01-${String(index + 1).padStart(2, "0")}` : "", 金额: String(100 + index * 10), 组: index < 25 ? "A" : "B"
  })), "trend.csv");
  filterGuardRows(trend, "组", ["B"]);
  trend.evaluate('dom.v2MetricField.value = "金额"; dom.v2GroupField.value = "组"; dom.v2AggregateMethod.value = "sum"; dom.v2DateField.value = "日期"; renderV2Analysis()');
  assert.deepEqual(chartEmpty(trend, "v2SelectedTrendChart"), { text: "所选时间字段没有可用的趋势数据", hidden: false });
  trend.evaluate("clearAllFilters()");
  trend.evaluate('dom.v2MetricField.value = "金额"; dom.v2GroupField.value = "组"; dom.v2AggregateMethod.value = "sum"; dom.v2DateField.value = "日期"; renderV2Analysis()');
  assert.deepEqual(chartEmpty(trend, "v2SelectedTrendChart"), { text: "", hidden: true });
});

test("v3.0.0 F3: empty tables and lists get the shared empty-state look by a class on their container; the sentences stay as they were", () => {
  const words = createChartContext();
  commitGuardRows(words, ["名称", "备注"], Array.from({ length: 20 }, (_, index) => ({ 名称: ["甲", "乙", "丙"][index % 3], 备注: index % 2 ? "x" : "y" })), "words.csv");
  const has = (id, name) => words.elements.get(id).classList.contains(name);
  assert.equal(has("outlierTable", "is-empty-table"), true);
  assert.equal(has("numericStatsTable", "is-empty-table"), true);
  assert.equal(has("idDuplicateDetails", "is-empty"), true);
  assert.equal(has("placeholderDetails", "is-empty"), true);
  assert.match(words.elements.get("outlierTable").innerHTML, /<td>未识别到数值字段。<\/td>/);
  // With data the marks go.
  // 订单ID is an ID by its name (two values repeat); 金额 has an N/A; 组 is for the filter below.
  commitGuardRows(words, ["日期", "金额", "订单ID", "组"], Array.from({ length: 20 }, (_, index) => ({ 日期: `2025-01-${String(index + 1).padStart(2, "0")}`, 金额: index === 3 ? "N/A" : String(index), 订单ID: `K${index % 18}`, 组: index % 2 ? "A" : "B" })), "data.csv");
  assert.deepEqual(JSON.parse(words.evaluate("JSON.stringify(state.profiles.map((profile) => profile.typeKey))")), ["date", "numeric", "id", "category"]);
  ["outlierTable", "numericStatsTable"].forEach((id) => assert.equal(has(id, "is-empty-table"), false, id));
  assert.equal(has("idDuplicateDetails", "is-empty"), false);
  assert.equal(has("placeholderDetails", "is-empty"), false);
  // A filtered range with no rows: the preview table.
  filterGuardRows(words, "组", ["没有这个值"]);
  assert.equal(has("previewTable", "is-empty-table"), true);
  words.evaluate("clearAllFilters()");
  assert.equal(has("previewTable", "is-empty-table"), false);

  const css = repoFile("style.css");
  const shared = css.match(/\n(\.empty-state,[^{]*)\{/);
  assert.ok(shared, "one shared rule");
  [".empty-note", ".filter-value-empty", ".filter-empty-note", ".id-duplicate-details.is-empty .id-duplicate-note", ".placeholder-details.is-empty .placeholder-note"]
    .forEach((selector) => assert.ok(shared[1].includes(selector), selector));
  assert.match(css, /\.is-empty-table td \{/);
});

// v3.0.0 commit 7a: module structure (section 二·6, 二·7, 二·9).
const metricItems = (scope) => [...scope.elements.get("overviewCards").innerHTML.matchAll(/<span class="metric-item"><span class="metric-label">([^<]*)<\/span> <strong class="metric-value">([^<]*)<\/strong><\/span>/g)].map((match) => [match[1], match[2]]);

test("v3.0.0 overview: the five numbers are one line of text with their old labels; a filter keeps the 总记录 → 当前分析 form", async () => {
  const orders = await importCsvFile(readExampleCsv("sample-orders.csv"), "sample-orders.csv");
  const counts = JSON.parse(orders.evaluate('JSON.stringify([state.rows.length, state.fields.length, getProfilesByType("numeric").length, getProfilesByType("category").length])'));
  assert.deepEqual(counts, [42, 9, 2, 5]);
  assert.match(orders.elements.get("overviewCards").innerHTML, /^\s*<p class="metric-line">/);
  assert.doesNotMatch(orders.elements.get("overviewCards").innerHTML, /metric-card/);
  assert.deepEqual(metricItems(orders), [["总记录数", "42"], ["字段数", "9"], ["数值字段数", "2"], ["分类字段数", "5"], ["空值概览", "6"]]);
  orders.evaluate(GUARD_EAST_FILTER);
  assert.deepEqual(metricItems(orders)[0], ["总记录 → 当前分析", "42 → 10"]);
});

test("v3.0.0 overview: the preview sits in an open <details> whose summary is the old 数据预览 heading", () => {
  const html = indexHtmlSource();
  const overview = html.slice(html.indexOf('id="overviewSection"'), html.indexOf('id="schemaSection"'));
  assertInOrder(overview, [
    '<details class="preview-details" open>',
    '<summary><h3 class="subsection-title">数据预览</h3></summary>',
    '<table id="previewTable" aria-label="数据预览"></table>',
    '<p id="previewPageInfo"',
    "</details>"
  ]);
});

test("v3.0.0 field configuration: the message and both buttons sit under the table, marked has-draft while drafts wait", async () => {
  const html = indexHtmlSource();
  const schema = html.slice(html.indexOf('<section id="schemaSection"'), html.indexOf('<section id="qualitySection"'));
  assertInOrder(schema, ['<table id="schemaTable"', '<div id="fieldConfigToolbar" class="field-config-toolbar">', '<p id="fieldConfigMessage" class="field-config-message" aria-live="polite"></p>', '<button id="restoreAutoFieldTypes"', '<button id="applyFieldConfig"']);
  const orders = await importCsvFile(readExampleCsv("sample-orders.csv"), "sample-orders.csv");
  const toolbar = orders.elements.get("fieldConfigToolbar");
  assert.equal(toolbar.classList.contains("has-draft"), false);
  orders.evaluate('stageFieldTypeChange("客户类型", "id")');
  assert.equal(toolbar.classList.contains("has-draft"), true);
  orders.evaluate("applyFieldConfiguration()");
  assert.equal(toolbar.classList.contains("has-draft"), false);
  orders.evaluate("resetAnalysisState()");
  assert.equal(toolbar.classList.contains("has-draft"), false);
});

test("v3.0.0 quality: only notes of more than one sentence keep their paragraph and get a 展开 button after it", () => {
  const html = indexHtmlSource();
  const quality = html.slice(html.indexOf('<section id="qualitySection"'), html.indexOf('<section id="filterSection"'));
  const notes = [...quality.matchAll(/<p class="section-note">([^<]*)<\/p>(\s*<button type="button" class="text-button note-toggle" aria-expanded="false">展开<\/button>)?/g)];
  const sentences = notes.map((match) => (match[1].match(/。/g) || []).length);
  // v3.0.0 look2: the module's own one-sentence note became the aside next to the title.
  assert.deepEqual(sentences, [1, 2, 1, 2, 3, 3]);
  // A single sentence shows whole and has no button.
  assert.deepEqual(notes.map((match) => Boolean(match[2])), sentences.map((count) => count > 1));
  assert.equal((quality.match(/class="text-button note-toggle"/g) || []).length, 4);
  // script.js marks the notes before a button; only those are cut to one line.
  assert.match(repoFile("script.js"), /document\.querySelectorAll\("\.note-toggle"\)\.forEach\(\(button\) => button\.previousElementSibling\.classList\.add\("is-clamped"\)\)/);
  assert.match(repoFile("style.css"), /#qualitySection \.section-note\.is-clamped:not\(\.is-expanded\) \{[^}]*white-space: nowrap;/);
  assert.match(quality, /<div id="placeholderSection" class="module-subsection">\s*<h3 class="subsection-title">像是空值的写法<\/h3>/);
  assert.equal((repoFile("script.js").match(/dom\.qualitySection\.addEventListener\("click", toggleSectionNote\)/g) || []).length, 1);
});

test("v3.0.0 quality: 展开 / 收起 opens the note before it and says so", () => {
  const scope = createScriptContext();
  const classes = new Set();
  const note = { classList: { toggle(name, force) { if (force) classes.add(name); else classes.delete(name); }, contains: (name) => classes.has(name) } };
  const attributes = { "aria-expanded": "false" };
  const button = { textContent: "展开", previousElementSibling: note, getAttribute: (name) => attributes[name], setAttribute: (name, value) => { attributes[name] = String(value); } };
  put(scope.context, "__noteEvent", { target: { closest: (selector) => (selector === ".note-toggle" ? button : null) } });
  scope.evaluate("toggleSectionNote(__noteEvent)");
  assert.deepEqual([classes.has("is-expanded"), attributes["aria-expanded"], button.textContent], [true, "true", "收起"]);
  scope.evaluate("toggleSectionNote(__noteEvent)");
  assert.deepEqual([classes.has("is-expanded"), attributes["aria-expanded"], button.textContent], [false, "false", "展开"]);
  // Clicks elsewhere do nothing.
  put(scope.context, "__otherEvent", { target: { closest: () => null } });
  scope.evaluate("toggleSectionNote(__otherEvent)");
});

test("v3.0.0 statistics: each frequency row has a bar as long as its share, and each card title a tooltip", async () => {
  const orders = await importCsvFile(readExampleCsv("sample-orders.csv"), "sample-orders.csv");
  const html = orders.elements.get("categoryStats").innerHTML;
  const rows = [...html.matchAll(/<span class="category-name" title="([^"]*)">[^<]*<\/span>\s*<span class="frequency-bar" aria-hidden="true"><span style="width: ([\d.]+)%"><\/span><\/span>\s*<strong>([^<]*)<\/strong>\s*<span>([^<]*)<\/span>/g)];
  assert.ok(rows.length > 0);
  rows.forEach(([, , width, , percent]) => assert.equal(`${Number(width).toFixed(1)}%`, percent, width));
  assert.match(html, /<h3 title="地区">地区<\/h3>/);
});

test("v3.0.0 export: each option shows its own row count after its title; the separate count row is gone", () => {
  const html = indexHtmlSource();
  const options = html.slice(html.indexOf('<div class="data-export-options">'), html.indexOf('<div id="standardizedExportPanel"'));
  // v3.2.0 commit 10: the words around the number are written by the script (“共 42 行” / “42 rows”).
  assert.match(options, /<h3>当前分析的行<\/h3>\s*<p class="export-count"><span id="dataExportFilteredBefore"><\/span><strong id="dataExportFilteredCount">0<\/strong><span id="dataExportFilteredAfter"><\/span><\/p>/);
  assert.doesNotMatch(options, /去重/);
  assert.match(options, /<h3>写法统一的表<\/h3>\s*<p class="export-count"><span id="dataExportOriginalBefore"><\/span><strong id="dataExportOriginalCount">0<\/strong><span id="dataExportOriginalAfter"><\/span><\/p>/);
  assert.doesNotMatch(html, /数据导出记录统计/);
});

test("v3.0.0 dirty data: a truncated table cell shows its full text on hover; chart titles carry their text as a tooltip", async () => {
  const scope = createScriptContext();
  const cell = (scrollWidth, text) => ({ tagName: "TD", textContent: text, scrollWidth, clientWidth: 100, title: "" });
  const long = cell(300, "很长的一格内容");
  const short = cell(80, "短");
  put(scope.context, "__cells", [long, short]);
  scope.evaluate("__cells.forEach((td) => showTruncatedCellTitle({ target: { closest: (selector) => (selector === \".table-wrap td, .table-wrap th\" ? td : null) } }))");
  assert.equal(long.title, "很长的一格内容");
  assert.equal(short.title, "");
  assert.equal((repoFile("script.js").match(/dom\.results\.addEventListener\("mouseover", showTruncatedCellTitle\)/g) || []).length, 1);

  const orders = await importCsvFile(readExampleCsv("sample-orders.csv"), "sample-orders.csv");
  withCapturableCharts(orders);
  orders.evaluate(`dom.numericChartField.value = "数量"; renderCharts(); ${GUARD_CUSTOM_ANALYSIS}`);
  ["numericChartTitle", "categoryChartTitle", "dateChartTitle", "v2TopGroupTitle", "v2TrendTitle", "v2PrimaryTableTitle"].forEach((id) => {
    const element = orders.elements.get(id);
    assert.ok(element.textContent.length > 0, id);
    assert.equal(element.title, element.textContent, id);
  });
});

test("v3.0.0 dirty data and phones: table cells are capped and cut; the upload sample rows wrap their reason below on phones", () => {
  const css = repoFile("style.css");
  // The tables of raw cell values: the data preview and the problem rows.
  assert.match(css, /#previewTable td,\s*#problemRowsTable td \{[^}]*max-width: 30em;[^}]*overflow: hidden;[^}]*text-overflow: ellipsis;[^}]*white-space: nowrap;/);
  assert.match(css, /\.table-wrap :is\(td, th\):first-child \{[^}]*position: sticky;[^}]*left: 0;/);
  const phone = css.slice(css.indexOf("@media (max-width: 640px)"));
  assert.match(phone, /\.sample-preview-table tr \{[^}]*display: grid;/);
  assert.match(phone, /\.sample-preview-table td\.mark-note,\s*\.sample-preview-table th:last-child \{[^}]*grid-column: 1 \/ -1;/);
  // The quality page's two tables stack (only there: the charts stay side by side).
  assert.match(css, /#qualitySection \.two-column \{\s*grid-template-columns: 1fr;/);
  assert.match(css, /\n\.chart-grid \{\s*display: grid;\s*grid-template-columns: repeat\(2, minmax\(0, 1fr\)\);/);
});

// v3.0.0 commit 7b: the look (section 二·7, 二·11). Styles only: these read style.css.
const cssRule = (css, selector) => {
  const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  // A rule of its own: not the last selector of a list ("a,\n b {").
  const match = css.match(new RegExp(`(?:^|\\n)(?<!,\\r?\\n)${escaped} \\{([^}]*)\\}`));
  assert.ok(match, selector);
  return match[1];
};

test("v3.0.0 look: the colour tokens, the system font stack, seven font sizes and three weights", () => {
  const css = repoFile("style.css");
  const root = cssRule(css, ":root");
  [["--bg", "#f2f2f7"], ["--panel", "#ffffff"], ["--text", "#1c1c1e"], ["--text-secondary", "#636366"], ["--muted", "#636366"], ["--primary", "#0068e0"], ["--sidebar-bg", "#f9f9fb"], ["--card-border", "rgba(28, 28, 30, 0.07)"], ["--radius-card", "16px"]]
    .forEach(([name, value]) => assert.ok(root.includes(`${name}: ${value};`), name));
  const body = cssRule(css, "body");
  assert.match(body, /background: var\(--bg\);/);
  const family = body.match(/font-family: ([^;]+);/)[1];
  assert.doesNotMatch(family, /Inter/);
  assert.ok(family.indexOf('"Segoe UI"') >= 0 && family.indexOf('"Segoe UI"') < family.indexOf('"Microsoft YaHei"'), family);
  const sizes = new Set([...css.matchAll(/font-size: ([^;]+);/g)].map((match) => match[1]));
  sizes.forEach((size) => assert.ok(["12px", "13px", "15px", "17px", "22px", "28px", "clamp(28px, 4vw, 40px)"].includes(size), size));
  const weights = new Set([...css.matchAll(/font-weight: ([^;]+);/g)].map((match) => match[1]));
  weights.forEach((weight) => assert.ok(["400", "600", "700"].includes(weight), weight));
  // The old purple is gone, the focus ring is the new blue.
  assert.doesNotMatch(css, /#4f46e5|#4338ca|rgba\(79, 70, 229/i);
  assert.match(css, /\[tabindex="0"\]:focus-visible \{\s*outline: 3px solid rgba\(0, 104, 224, 0\.45\);/);
});

test("v3.0.0 look: modules sit on the grey page; their content blocks are white cards with a hairline and no shadow", () => {
  const css = repoFile("style.css");
  // The module itself: no card (no fill, border or padding), so its title is on the grey.
  const block = cssRule(css, ".section-block");
  assert.match(block, /padding: 0;/);
  assert.match(block, /border: 0;/);
  assert.match(block, /background: transparent;/);
  assert.doesNotMatch(block, /box-shadow/);
  // The content blocks inside it are the cards.
  const cards = css.match(/\n(\.table-wrap,\s*\.frequency-card,[^{]*)\{([^}]*)\}/);
  assert.ok(cards, "one card rule");
  const cardSelectors = cards[1].split(",").map((part) => part.trim());
  assert.deepEqual(cardSelectors, [".table-wrap", ".frequency-card", ".data-export-options", ".report-export-layout", ".filter-condition", ".filter-toolbar", ".chart-card"]);
  assert.match(cards[2], /border: 1px solid var\(--card-border\);/);
  assert.match(cards[2], /border-radius: var\(--radius-card\);/);
  assert.match(cards[2], /background: #fff;/);
  // Those blocks set no other fill, border or radius of their own.
  cardSelectors.forEach((selector) => {
    const own = css.match(new RegExp(`(?:^|\\n)(?<!,\\r?\\n)${selector.replace(".", "\\.")} \\{([^}]*)\\}`));
    if (own) assert.doesNotMatch(own[1], /background:|border:|border-radius:/, selector);
  });
  // The containers that hold such cards (custom analysis and its details) are not cards
  // themselves: no card inside a card. (v3.0.0 look2: the problem rows became a card.)
  [".analysis-panel", ".analysis-details"].forEach((selector) => {
    const own = cssRule(css, selector);
    assert.match(own, /padding: 0;/, selector);
    assert.match(own, /border: 0;/, selector);
    assert.match(own, /background: transparent;/, selector);
  });
  // Every box-shadow is none, an inset line, a 1 px outline ring, or the sidebar over the content.
  const shadows = [...css.matchAll(/\n\s*([^{}\n][^{}]*?)\s*\{([^{}]*box-shadow: ([^;]+);[^{}]*)\}/g)].map((match) => [match[1].trim(), match[3]]);
  assert.ok(shadows.length > 0);
  shadows.forEach(([selector, value]) => {
    const ok = value === "none" || /^inset /.test(value) || /^0 0 0 1px /.test(value) || selector === "body.has-results.sidebar-open .app-sidebar";
    assert.ok(ok, `${selector}: ${value}`);
  });
});

test("v3.0.0 look: buttons 34 high (small 28), 44 on phones; the navigation is a segmented control; no hover on plain rows", () => {
  const css = repoFile("style.css");
  const base = css.match(/\nbutton,\r?\n\.file-button,\r?\n\.secondary-button \{([^}]*)\}/)[1];
  assert.match(base, /min-height: 34px;/);
  assert.match(base, /border-radius: 10px;/);
  // Secondary buttons are white with a grey line and body-coloured text; text buttons stay blue.
  const secondary = cssRule(css, ".secondary-button");
  assert.match(secondary, /border-color: rgba\(28, 28, 30, 0\.14\);/);
  assert.match(secondary, /background: #fff;/);
  assert.match(secondary, /color: var\(--text\);/);
  assert.match(cssRule(css, ".text-button"), /border: 0;[\s\S]*background: transparent;[\s\S]*color: var\(--primary\);/);
  assert.match(cssRule(css, ".problem-view-button"), /min-height: 28px;[\s\S]*border-radius: 8px;/);
  const nav = cssRule(css, ".module-nav");
  assert.match(nav, /background: rgba\(28, 28, 30, 0\.06\);/);
  assert.match(nav, /border-radius: 11px;/);
  assert.match(css, /\.module-nav a\[aria-current="page"\] \{[^}]*background: #fff;[^}]*font-weight: 600;/);
  const phone = css.slice(css.indexOf("@media (max-width: 640px)"), css.indexOf("@media (max-width: 390px)"));
  assert.match(phone, /\s+button,\s+\.file-button,\s+\.secondary-button,\s+select,\s+input:not\(\[type="checkbox"\], \[type="radio"\], \[type="file"\]\) \{\s*min-height: 44px;/);
  assert.match(phone, /\.problem-view-button,[^{]*\.placeholder-button,[^{]*\{\s*min-height: 44px;/);
  assert.doesNotMatch(css, /tbody tr:hover/);
  // The date inputs shrink with their column instead of sticking out of the filter card.
  assert.match(cssRule(css, ".filter-range-grid label"), /grid-template-columns: minmax\(0, 1fr\);/);
  const th = cssRule(css, "th");
  assert.match(th, /background: #fafafb;/);
  assert.match(th, /font-size: 12px;/);
});

test("v3.0.0 look: quality tags are neutral (white, grey line, grey at 0), no amber, no dots; five subsection titles carry their category colour", () => {
  const css = repoFile("style.css");
  const pill = cssRule(css, ".quality-pill");
  assert.match(pill, /min-height: 24px;/);
  assert.match(pill, /border: 1px solid rgba\(28, 28, 30, 0\.14\);/);
  assert.match(pill, /border-radius: 6px;/);
  assert.match(pill, /background: #fff;/);
  assert.match(pill, /color: var\(--text\);/);
  assert.match(pill, /font-weight: 600;/);
  assert.match(cssRule(css, ".quality-pill.success"), /color: var\(--text-secondary\);[\s\S]*font-weight: 400;/);
  assert.doesNotMatch(css, /\.quality-pill\.(?:warning|danger)/);
  assert.doesNotMatch(css, /\.quality-pill[^{]*::before/);
  // 认不出 above 0 (quality and statistics tables): the pink dot of the field table, no amber.
  const warningCell = cssRule(css, ".cell-warning");
  assert.doesNotMatch(warningCell, /--warning/);
  assert.match(css, /\.cell-warning::before \{[^}]*background: #ff2d55;/);
  // A problem cell in the problem rows: neutral grey and semibold.
  const problemCell = cssRule(css, "td.cell-problem");
  assert.doesNotMatch(problemCell, /--warning/);
  assert.match(problemCell, /background: rgba\(28, 28, 30, 0\.06\);/);
  assert.match(problemCell, /font-weight: 600;/);
  // The import message keeps its amber warning.
  assert.match(cssRule(css, ".status-panel.warning"), /background: var\(--warning-soft\);/);
  // v3.0.0 look2: the privacy badge and its styles are gone.
  assert.doesNotMatch(css, /\.privacy-badge/);
  [["suspectedTotalSection", "#5856d6"], ["formulaErrorSection", "#c93400"], ["dateStyleSection", "#a86400"], ["lookAlikeSection", "#22973f"], ["placeholderSection", "#148a9e"]].forEach(([id, colour]) => {
    assert.match(css, new RegExp(`#${id} \\{ --dot: ${colour}; \\}`), id);
  });
  assert.match(css, /:is\(#suspectedTotalSection, #formulaErrorSection, #dateStyleSection, #lookAlikeSection, #placeholderSection\) > \.subsection-title::before \{[^}]*background: var\(--dot\);/);
  // The sections without a category colour have no dot.
  ["numberConversionSection", "numericLookingSection", "constantFieldSection", "problemRowsSection"].forEach((id) => assert.doesNotMatch(css, new RegExp(`#${id} \\{ --dot`), id));
});

test("v3.0.0 look: one blue button per group, by id only; the class names stay", () => {
  const css = repoFile("style.css");
  // 导出当前分析的行 keeps its markup (no secondary-button) but looks secondary: the data group's
  // blue button is 导出写法统一的表 (in the check panel).
  assert.match(css, /#exportFilteredCsv:not\(:disabled\) \{[^}]*border-color: rgba\(28, 28, 30, 0\.14\);[^}]*background: #fff;[^}]*color: var\(--text\);/);
  // 按空值处理 on each line looks secondary; 常见写法全部按空值处理 stays blue.
  assert.match(css, /#placeholderDetails \.placeholder-item > \.placeholder-button:not\(\.secondary-button, :disabled\) \{[^}]*background: #fff;[^}]*color: var\(--text\);/);
  const html = indexHtmlSource();
  assert.match(html, /<button id="exportFilteredCsv" type="button" disabled>导出当前分析的行<\/button>/);
});

test("v3.0.0 look: the 空值检查 box shows about ten fields without scrolling; print lifts every height limit", () => {
  const css = repoFile("style.css");
  const compact = cssRule(css, ".table-wrap.compact");
  const height = Number(compact.match(/max-height: (\d+)px;/)[1]);
  // A header and ten rows of 42 px, and the two border pixels.
  assert.ok(height >= 42 * 11 + 2, String(height));
  assert.ok(height <= 600, String(height));
  // Below 640 px no table box limits its height (二·9).
  const phone = css.slice(css.indexOf("@media (max-width: 640px)"), css.indexOf("@media (max-width: 390px)"));
  assert.match(phone, /\.table-wrap\.compact \{\s*max-height: none;/);
  const print = css.match(/@media print \{([\s\S]*?)\n\}/)[1];
  assert.match(print, /\.table-wrap\.compact,\s*\.schema-table-wrap \{\s*max-height: none !important;/);
});

test("v3.0.0 look: charts take the new blue and the page's font", () => {
  const scope = createChartContext();
  const theme = JSON.parse(scope.evaluate("JSON.stringify(CHART_THEME)"));
  assert.deepEqual(Object.keys(theme), ["primary", "primaryFill", "secondary", "secondaryFill", "warning", "text", "grid"]);
  assert.equal(theme.primary, "#0068e0");
  assert.equal(theme.text, "#636366");
  // The second series is a lighter blue that still has 3:1 against the white card.
  const luminance = (hex) => {
    const [r, g, b] = [1, 3, 5].map((index) => parseInt(hex.slice(index, index + 2), 16) / 255)
      .map((value) => (value <= 0.03928 ? value / 12.92 : ((value + 0.055) / 1.055) ** 2.4));
    return 0.2126 * r + 0.7152 * g + 0.0722 * b;
  };
  assert.match(theme.secondary, /^#[0-9a-f]{6}$/);
  const [red, green, blue] = [1, 3, 5].map((index) => parseInt(theme.secondary.slice(index, index + 2), 16));
  assert.ok(blue > red && blue > green, theme.secondary);
  assert.ok(luminance(theme.secondary) > luminance(theme.primary), "lighter than the primary");
  assert.ok(1.05 / (luminance(theme.secondary) + 0.05) >= 3, "3:1 on white");
  assert.equal(theme.secondaryFill, `rgba(${red}, ${green}, ${blue}, 0.14)`);
  // The insight under a chart: a light grey box, no side line.
  const insight = cssRule(repoFile("style.css"), ".chart-insight");
  assert.doesNotMatch(insight, /border-left/);
  assert.match(insight, /background: #f6f6f9;/);
  assert.match(insight, /border-radius: 10px;/);
  assert.match(insight, /color: var\(--text-secondary\);/);
  // The custom analysis's result sentence looks the same.
  const summary = cssRule(repoFile("style.css"), ".custom-result-summary");
  assert.doesNotMatch(summary, /border-left|--primary-soft/);
  ["background: #f6f6f9;", "border-radius: 10px;", "color: var(--text-secondary);"].forEach((line) => assert.ok(summary.includes(line), line));
  const script = repoFile("script.js");
  const family = script.match(/Chart\.defaults\.font\.family = '([^']+)';/)[1];
  assert.equal(family, cssRule(repoFile("style.css"), "body").match(/font-family: ([^;]+);/)[1]);
  assert.match(script, /Chart\.defaults\.font\.size = 12;/);
});

// v3.0.0 look, second pass (impl-v3.0.0-look2): fewer template touches. Items 1–7 of that file.
test("v3.0.0 look2: the top bar has only the name (no mark, no subtitle, no privacy badge); the footer only the GitHub link", () => {
  const html = indexHtmlSource();
  const css = repoFile("style.css");
  const header = html.slice(html.indexOf('<header class="app-header">'), html.indexOf("</header>", html.indexOf('<header class="app-header">')));
  assert.match(header, /<strong>表格体检<\/strong>/);
  assert.doesNotMatch(header, /<small>|brand-mark|privacy-badge|数据仅在浏览器本地处理/);
  assert.doesNotMatch(html, /brand-mark|privacy-badge/);
  assert.doesNotMatch(css, /\.brand-mark|\.privacy-badge|\.brand-copy small/);
  assert.match(html, /<title>表格体检 \| 检查 CSV、Excel 表格的数据问题<\/title>/);
  const footer = html.slice(html.indexOf('<footer class="app-footer">'), html.indexOf("</footer>"));
  assert.equal((footer.match(/<p>/g) || []).length, 1);
  assert.match(footer, /<p><a class="footer-github-link" href="https:\/\/github\.com\/wy-data-30\/sheet-checkup">[\s\S]*<span>GitHub 仓库<\/span><\/a><\/p>/);
  assert.doesNotMatch(footer, /表格数据体检工具|表格体检|检查 CSV、Excel 表格的数据问题|不会上传到服务器|<strong>/);
  // The privacy sentences that stay: under the upload title and at the foot of the sidebar.
  assert.match(html, /<p class="product-description">不用安装，表格不会离开这台电脑。<\/p>/);
  assert.match(html, /<p class="sidebar-privacy">表格只在这台电脑的浏览器里处理，不会上传。<\/p>/);
  const footerRule = cssRule(css, ".app-footer");
  assert.match(footerRule, /justify-content: flex-end;/);
  assert.match(footerRule, /font-size: 12px;/);
  assert.match(footerRule, /color: var\(--text-secondary\);/);
});

test("v3.0.0 look2: module titles are 17 px without a rule; three modules carry a one-line aside next to the title", () => {
  const html = indexHtmlSource();
  const css = repoFile("style.css");
  const heading = cssRule(css, ".section-heading");
  assert.doesNotMatch(heading, /border-bottom|padding-bottom/);
  assert.match(heading, /margin-bottom: 14px;/);
  const h2 = cssRule(css, ".section-heading h2");
  assert.match(h2, /font-size: 17px;/);
  assert.match(h2, /font-weight: 600;/);
  assert.match(h2, /letter-spacing: 0;/);
  const aside = cssRule(css, ".section-aside");
  assert.match(aside, /font-size: 13px;/);
  assert.match(aside, /color: var\(--text-secondary\);/);
  assert.match(cssRule(css, ".section-title-line"), /align-items: baseline;/);
  assert.match(cssRule(css, ".section-title-line"), /flex-wrap: wrap;/);
  const asides = [
    ["schemaSection", "schemaSectionTitle", "确认字段类型", "识别不准时再改；应用后统一重新计算"],
    ["qualitySection", "qualitySectionTitle", "空值、重复和其他问题", "始终按导入的原始数据算，不受筛选影响"],
    ["filterSection", "filterSectionTitle", "设置当前分析范围", "只影响分析结果，不改原始数据和数据质量"]
  ];
  for (const [section, id, title, text] of asides) {
    const start = html.indexOf(`<section id="${section}"`);
    const head = html.slice(html.indexOf('<div class="section-heading">', start));
    assert.match(head, new RegExp(`^<div class="section-heading">\\s*<div class="section-title-line">\\s*<h2 id="${id}" tabindex="-1">${title}</h2>\\s*<p class="section-aside">${text}</p>\\s*</div>\\s*</div>`), section);
    // The old paragraph under the heading is gone.
    assert.doesNotMatch(head, /^<div class="section-heading">[\s\S]*?<\/div>\s*<\/div>\s*<p class="section-note">/, section);
  }
  assert.doesNotMatch(html, /系统已自动推断字段类型|本模块始终基于导入的原始数据计算|分类、数值和日期条件可以同时应用/);
  assert.equal((html.match(/class="section-aside"/g) || []).length, 3);
  // The other modules keep their paragraph (custom analysis) and the 当前分析范围 on the right.
  // v3.2.0 commit 10: heading → the note bar → the paragraph, in this order.
  assert.match(html, /<\/div>\s*<p id="customContainsNote" class="analysis-contains-note hidden"><\/p>\s*<p class="section-note">选择一个指标、一个分组维度和汇总方式即可生成结果；时间字段为可选。<\/p>/);
  assert.equal((html.match(/<p class="filter-scope">当前分析范围：/g) || []).length, 3);
});

test("v3.0.0 look2: the quality page is cards: three of their own, and one holding the seven note subsections", () => {
  const html = indexHtmlSource();
  const css = repoFile("style.css");
  const quality = html.slice(html.indexOf('<section id="qualitySection"'), html.indexOf('<section id="filterSection"'));
  const start = quality.indexOf('<div class="quality-notes">');
  assert.ok(start > quality.indexOf('id="problemRowsSection"'), String(start));
  const notes = quality.slice(start);
  assertInOrder(notes, ['id="numberConversionSection"', 'id="numericLookingSection"', 'id="dateStyleSection"', 'id="constantFieldSection"', 'id="lookAlikeSection"', 'id="placeholderSection"', ">编号重复</h3>"]);
  // Exactly those seven subsections inside it, and none of them before it.
  assert.equal((notes.match(/<div (?:id="[^"]+" )?class="module-subsection[^"]*">/g) || []).length, 7);
  assert.equal(quality.slice(0, start).includes('id="numberConversionSection"'), false);
  const box = cssRule(css, ".quality-notes");
  assert.match(box, /border: 1px solid var\(--card-border\);/);
  assert.match(box, /border-radius: var\(--radius-card\);/);
  assert.match(box, /background: #fff;/);
  assert.doesNotMatch(box, /box-shadow/);
  // The three cards of their own are styled by id, so their class lists stay as they were.
  const card = css.match(/\n#suspectedTotalSection,\r?\n#formulaErrorSection,\r?\n#problemRowsSection \{([^}]*)\}/)[1];
  assert.match(card, /padding: 18px 20px;/);
  assert.match(card, /border: 1px solid var\(--card-border\);/);
  assert.match(card, /border-radius: var\(--radius-card\);/);
  assert.match(card, /background: #fff;/);
  // A line only between two shown subsections: hidden ones leave no double line.
  assert.match(css, /\.quality-notes > \.module-subsection:not\(\.hidden\) ~ \.module-subsection:not\(\.hidden\) \{\s*border-top: 1px solid var\(--card-border\);/);
  assert.match(cssRule(css, ".quality-notes > .module-subsection"), /border-top: 0;/);
});

test("v3.0.0 look2: number columns of the three quality and statistics tables carry the num class; the report has none", async () => {
  const orders = await importCsvFile(readExampleCsv("sample-orders.csv"), "sample-orders.csv");
  const classesOf = (tableHtml) => [...tableHtml.matchAll(/<tr>([\s\S]*?)<\/tr>/g)]
    .map((row) => [...row[1].matchAll(/<t[hd](?: class="([^"]*)")?>/g)].map((cell) => (cell[1] || "").split(" ").includes("num")));
  const expectRows = (id, pattern) => {
    const rows = classesOf(orders.elements.get(id).innerHTML);
    assert.ok(rows.length > 1, id);
    rows.forEach((row) => assert.deepEqual(row, pattern, id));
  };
  expectRows("qualityTable", [false, false, true, true, true, false]);
  expectRows("outlierTable", [false, true, true, true, false]);
  expectRows("numericStatsTable", [false, true, true, true, true, true, true, true]);
  // The warning colour still sits on the same cell.
  assert.match(orders.elements.get("qualityTable").innerHTML, /<td class="num cell-warning">/);
  assert.match(repoFile("style.css"), /\nth\.num,\r?\ntd\.num \{[^}]*text-align: right;/);
  const report = orders.evaluate("buildHtmlReport(buildReportData(new Date(2026, 0, 1)))");
  assert.doesNotMatch(report, /class="num/);
});

test("v3.0.0 look2: trend lines are straight with white points; every chart has at most 5 y ticks", async () => {
  const orders = await importCsvFile(readExampleCsv("sample-orders.csv"), "sample-orders.csv");
  const charts = withFakeCharts(orders);
  orders.evaluate(`
    dom.numericChartField.value = "金额";
    renderCharts();
    dom.v2MetricField.value = "金额"; dom.v2GroupField.value = "地区"; dom.v2AggregateMethod.value = "sum"; dom.v2DateField.value = "下单日期";
    renderV2Analysis();
  `);
  const lines = charts.filter((config) => config.type === "line");
  assert.equal(lines.length, 2);
  for (const config of lines) {
    const dataset = config.data.datasets[0];
    assert.equal(dataset.tension, 0);
    assert.equal(dataset.fill, false);
    assert.equal(dataset.borderWidth, 2.5);
    assert.equal(dataset.pointRadius, 4);
    assert.equal(dataset.pointBackgroundColor, "#ffffff");
    assert.equal(dataset.pointBorderColor, dataset.borderColor);
  }
  assert.equal(charts.length, 5);
  charts.forEach((config) => assert.equal(config.options.scales.y.ticks.maxTicksLimit, 5));
  assert.equal(orders.evaluate('chartOptions("记录数").scales.y.ticks.maxTicksLimit'), 5);
  // CHART_THEME keeps its seven keys.
  assert.deepEqual(JSON.parse(orders.evaluate("JSON.stringify(Object.keys(CHART_THEME))")), ["primary", "primaryFill", "secondary", "secondaryFill", "warning", "text", "grid"]);
});

test("v3.0.0 look2: x-axis labels keep a visible gap (autoSkipPadding 16) on every chart", async () => {
  const orders = await importCsvFile(readExampleCsv("sample-orders.csv"), "sample-orders.csv");
  const charts = withFakeCharts(orders);
  orders.evaluate(`
    dom.numericChartField.value = "金额";
    renderCharts();
    dom.v2MetricField.value = "金额"; dom.v2GroupField.value = "地区"; dom.v2AggregateMethod.value = "sum"; dom.v2DateField.value = "下单日期";
    renderV2Analysis();
  `);
  assert.equal(charts.length, 5);
  charts.forEach((config) => {
    assert.equal(config.options.scales.x.ticks.autoSkip, true);
    assert.equal(config.options.scales.x.ticks.autoSkipPadding, 16);
  });
});

test("v3.0.0 look2: the field table's 非空值、唯一值、认不出 columns carry the num class", async () => {
  const orders = await importCsvFile(readExampleCsv("sample-orders.csv"), "sample-orders.csv");
  const rows = [...orders.elements.get("schemaTable").innerHTML.matchAll(/<tr>([\s\S]*?)<\/tr>/g)]
    .map((row) => [...row[1].matchAll(/<t[hd](?: class="([^"]*)")?>/g)].map((cell) => (cell[1] || "").split(" ").includes("num")));
  assert.ok(rows.length > 1);
  rows.forEach((row) => assert.deepEqual(row, [false, false, false, true, true, true, false]));
});

test("v3.0.0 look2: subheadings inside modules are 15 px, weight 600 (the export headings no longer look like module titles)", () => {
  const css = repoFile("style.css");
  const rules = [...css.matchAll(/([^{}]+)\{([^{}]*)\}/g)]
    .map((match) => [match[1].trim().replace(/\s+/g, " "), match[2]])
    .filter(([selector, body]) => /(?:^|[ ,])(?:[^ ,]* )?h3$|\.subsection-title$/.test(selector) && /font-size:/.test(body));
  const selectors = rules.map(([selector]) => selector);
  assert.deepEqual(selectors, [".subsection-title", ".filter-condition-heading h3", ".frequency-card h3", ".chart-header h3", ".export-module-heading h3", ".data-export-option h3", ".report-export-copy h3"]);
  rules.forEach(([selector, body]) => {
    assert.match(body, /font-size: 15px;/, selector);
    assert.match(body, /font-weight: 600;/, selector);
  });
});

test("v3.0.0 look2: the upload icon has no tinted square; the two upload buttons have no icon", () => {
  const html = indexHtmlSource();
  const buttons = html.slice(html.indexOf('<div class="upload-buttons">'), html.indexOf('<div class="upload-footer">'));
  assert.match(buttons, /<label class="file-button" for="csvFile">\s*选择数据文件\s*<\/label>/);
  assert.match(buttons, /<button id="loadSample" class="secondary-button" type="button">\s*使用示例数据\s*<\/button>/);
  assert.doesNotMatch(buttons, /<svg/);
  assert.match(html, /<span class="upload-icon" aria-hidden="true">\s*<svg viewBox="0 0 48 48">/);
  const icon = cssRule(repoFile("style.css"), ".upload-icon");
  assert.doesNotMatch(icon, /background|border-radius/);
  assert.match(icon, /color: var\(--primary\);/);
});

// v3.0.1: the product is called 表格体检. The old name, the old subtitle and the old report title
// never come back on the page (source without comments, title, aria labels, every element after
// importing the sample), in either report, or in an export file name; the new texts are exact.
// The old words are joined from parts so that this file is not a hit when the repository is
// searched for them. The Latin words are matched without case but with a space, so a lowercase
// repository name with hyphens is never a hit. v3.2.2 renamed the repository to sheet-checkup; the
// v3.2.2 guard at the end of this file looks for the old repository name on its own.
const OLD_PRODUCT_WORDS = [["Smart", "Tabular", "Analyzer"].join(" "), ["Tabular", "Analyzer"].join(" "), ["表格数据", "体检工具"].join(""), ["数据分析", "报告"].join("")];
const RENAME_REPO_URL = "https://github.com/wy-data-30/sheet-checkup";

function findOldProductWords(where, text, hits) {
  const lower = String(text).toLowerCase();
  for (const word of OLD_PRODUCT_WORDS) {
    const at = lower.indexOf(word.toLowerCase());
    if (at >= 0) hits.push(`${where}: 「${word}」 in …${String(text).slice(Math.max(0, at - 20), at + word.length + 20)}…`);
  }
}

function findOldWordsInProductScope(scope, hits) {
  for (const [id, element] of scope.elements) {
    findOldProductWords(`#${id}.textContent`, element.textContent, hits);
    findOldProductWords(`#${id}.innerHTML`, element.innerHTML, hits);
  }
}

test("v3.0.1 rename: the old product name and report title never appear; the new name is exact", async () => {
  const hits = [];
  const html = indexHtmlSource().replace(/<!--[\s\S]*?-->/g, "");
  findOldProductWords("index.html", html, hits);
  assert.equal((html.match(/<title>([^<]*)<\/title>/) || [])[1], "表格体检 | 检查 CSV、Excel 表格的数据问题");
  assert.equal((html.match(/<a class="brand" href="#mainContent" aria-label="([^"]*)">/) || [])[1], "表格体检首页");
  const header = html.slice(html.indexOf('<header class="app-header">'), html.indexOf("</header>"));
  assert.deepEqual(header.match(/<strong>[^<]*<\/strong>/g), ["<strong>表格体检</strong>"]);
  for (const label of html.match(/aria-label="[^"]*"/g) || []) findOldProductWords("index.html aria-label", label, hits);

  // The four-character name must not break across lines (642-790 px): nowrap is on the base rule
  // now, not only on phones.
  const css = repoFile("style.css");
  assert.match(cssRule(css, ".brand-copy strong"), /white-space: nowrap;/);
  const phoneBrand = (css.slice(css.indexOf("@media (max-width: 640px)")).match(/\.brand-copy strong \{([^}]*)\}/) || [])[1];
  assert.match(phoneBrand, /font-size: 13px;/);
  assert.doesNotMatch(phoneBrand, /white-space/);

  const orders = await importCsvFile(readExampleCsv("sample-orders.csv"), "sample-orders.csv");
  findOldWordsInProductScope(orders, hits);
  orders.evaluate(GUARD_CUSTOM_ANALYSIS);
  orders.evaluate(GUARD_EAST_FILTER);
  const { html: reportHtml } = reportOf(orders);
  findOldProductWords("html report", reportHtml, hits);
  assert.equal((reportHtml.match(/<title>([^<]*)<\/title>/) || [])[1], "sample-orders.csv - 表格体检报告");
  assert.deepEqual(reportHtml.match(/<h1>[^<]*<\/h1>/g), ["<h1>表格体检报告</h1>"]);
  assert.equal((reportHtml.match(/<footer>由<a href="([^"]*)">([^<]*)<\/a> /) || []).slice(1).join(" "), `${RENAME_REPO_URL} 表格体检`);
  for (const text of [html, reportHtml]) assert.ok(text.includes(RENAME_REPO_URL));

  const names = [
    orders.evaluate(`buildReportFileName("html", new Date(2026, 10, 15))`),
    orders.evaluate(`buildDataExportFileName("filtered", new Date(2026, 10, 15))`),
    orders.evaluate(`buildStandardizedFileName("data", new Date(2026, 10, 15))`), orders.evaluate(`buildStandardizedFileName("changes", new Date(2026, 10, 15), 3)`),
    ...JSON.parse(orders.evaluate("JSON.stringify(state.problemCategories.map((category) => buildProblemRowsFileName(category, new Date(2026, 10, 15))))"))
  ];
  assert.ok(names.length > 6);
  // v3.1.0: 21 names, the deduplicated export is gone; v3.2.0: 20, the Markdown report is gone.
  assert.equal(names.length, 20);
  for (const name of names) {
    assert.ok(name.startsWith("sample-orders"), name);
    findOldProductWords("file name", name, hits);
    assert.ok(!name.includes("表格体检"), name);
  }
  assert.deepEqual(hits.slice(0, 20), [], `${hits.length} hits`);
});

// ---- v3.1.0 commit 1: what treating a value as empty changed ----
// The messages say the counts before and after, that no row went, and why the statistics
// stay; "已清除现有筛选条件。" only when a filter was applied before.

const ID_WITH_NA = ["N/A", "N/A", "N/A", ...Array.from({ length: 11 }, (_, index) => `D${100 + index}`)];

test("v3.1.0 empty values: a numeric field says the counts, that no row went and that the statistics stay", async () => {
  const orders = await clickSampleButton(sampleOrdersFetchedText());
  orders.evaluate('confirmPlaceholder("金额", "N/A")');
  assert.equal(placeholderMessage(orders), "已把「金额」的 1 个「N/A」按空值处理：空值 0 → 1，认不出 1 → 0。只是把这 1 格当作空，这一行没有删，其他字段照常参与分析，所以行数还是 42。平均值等统计不变，因为「N/A」本来就没有算进去。");
  orders.evaluate('undoPlaceholderConfirmation("金额", "N/A")');
  assert.equal(placeholderMessage(orders), "已撤销「金额」的「N/A」：空值 1 → 0，认不出 0 → 1。这 1 格重新按原来的文字处理，行数还是 42。");
});

test("v3.1.0 empty values: a date field says the trend by date stays, not the average", async () => {
  const dates = await importStage3Fixture("订单表.csv");
  dates.evaluate('confirmPlaceholder("下单日期", "—")');
  assert.equal(placeholderMessage(dates), "已把「下单日期」的 1 个「—」按空值处理：空值 0 → 1，认不出 1 → 0。只是把这 1 格当作空，这一行没有删，其他字段照常参与分析，所以行数还是 30。按日期的趋势不变，因为「—」本来就没有算进去。");
  assert.doesNotMatch(placeholderMessage(dates), /平均值/);
});

test("v3.1.0 empty values: a category field loses the class, and the count is every cell the spelling covers", async () => {
  const orders = await clickSampleButton(sampleOrdersFetchedText());
  orders.evaluate('confirmPlaceholder("销售代表", "—")');
  assert.equal(placeholderMessage(orders), "已把「销售代表」的 1 个「—」按空值处理：空值 1 → 2，分类里不再有「—」这一类。只是把这 1 格当作空，这一行没有删，其他字段照常参与分析，所以行数还是 42。");
  orders.evaluate('undoPlaceholderConfirmation("销售代表", "—")');
  assert.equal(placeholderMessage(orders), "已撤销「销售代表」的「—」：空值 2 → 1，分类里重新有「—」这一类。这 1 格重新按原来的文字处理，行数还是 42。");
  const mixed = await importStage3Fixture("写法不一致.csv");
  mixed.evaluate('confirmPlaceholder("客户", "N/A")');
  assert.equal(placeholderMessage(mixed), "已把「客户」的 2 个「N/A」按空值处理：空值 1 → 3，分类里不再有「N/A」这一类。只是把这 2 格当作空，这些行没有删，其他字段照常参与分析，所以行数还是 24。");
});

test("v3.1.0 empty values: an ID field says the duplicate list, a changed type says the type, an ignored field says why", () => {
  const ids = commitColumns({ 订单ID: ID_WITH_NA, 金额: ID_WITH_NA.map((_, index) => String(10 + index)) });
  ids.evaluate('confirmPlaceholder("订单ID", "N/A")');
  assert.equal(placeholderMessage(ids), "已把「订单ID」的 3 个「N/A」按空值处理：空值 0 → 3，编号重复里不再列出「N/A」。只是把这 3 格当作空，这些行没有删，其他字段照常参与分析，所以行数还是 14。");
  ids.evaluate('undoPlaceholderConfirmation("订单ID", "N/A")');
  assert.equal(placeholderMessage(ids), "已撤销「订单ID」的「N/A」：空值 3 → 0，编号重复里重新列出「N/A」。这 3 格重新按原来的文字处理，行数还是 14。");

  const amounts = commitColumns({ 金额: AMOUNT_WITH_NA });
  amounts.evaluate('confirmPlaceholder("金额", "N/A")');
  assert.equal(placeholderMessage(amounts), "已把「金额」的 1 个「N/A」按空值处理：空值 0 → 1，认不出 0 → 1。只是把这 1 格当作空，这一行没有删，其他字段照常参与分析，所以行数还是 6。「金额」现在识别为数值（原来是分类）。");
  amounts.evaluate('undoPlaceholderConfirmation("金额", "N/A")');
  assert.equal(placeholderMessage(amounts), "已撤销「金额」的「N/A」：空值 1 → 0，认不出 1 → 0。这 1 格重新按原来的文字处理，行数还是 6。「金额」现在识别为分类（原来是数值）。");

  const notes = commitColumns({ 备注: ["N/A", "N/A", "甲", "乙", "丙", "丁"], 金额: ["1", "2", "3", "4", "5", "6"] });
  notes.evaluate('confirmPlaceholder("备注", "N/A"); stageFieldTypeChange("备注", "ignore"); applyFieldConfiguration()');
  assert.equal(notes.evaluate('undoPlaceholderConfirmation("备注", "N/A")'), true);
  assert.equal(placeholderMessage(notes), "已撤销「备注」的「N/A」。「备注」已忽略，撤销只影响以后改回其他类型时的结果。");
});

test("v3.1.0 empty values: treating every common spelling as empty names them, type changes first", async () => {
  const orders = await clickSampleButton(sampleOrdersFetchedText());
  orders.evaluate("confirmAllCommonPlaceholders()");
  assert.equal(placeholderMessage(orders), "已把 2 处常见写法按空值处理（「金额」的「N/A」、「销售代表」的「—」）。只是把这些格当作空，行都没有删，所以行数还是 42。");
  // 说明 comes first in the table but keeps its type, so it is named after the two that change.
  const scope = commitColumns({
    说明: ["N/A", "甲", "乙", "丙", "丁", "戊", "己"],
    年份: YEARS_WITH_DASHES,
    金额: ["100", "200", "N/A", "约500", "300", "400", "500"]
  });
  scope.evaluate("confirmAllCommonPlaceholders()");
  assert.equal(placeholderMessage(scope), "已把 3 处常见写法按空值处理（「年份」的「—」、「金额」的「N/A」、「说明」的「N/A」）。「年份」现在识别为日期（原来是分类）；「金额」现在识别为数值（原来是分类）。只是把这些格当作空，行都没有删，所以行数还是 7。");
});

test("v3.1.0 已清除现有筛选条件 only follows a rebuild that had a filter applied", async () => {
  const orders = await clickSampleButton(sampleOrdersFetchedText());
  const fieldMessage = () => orders.elements.get("fieldConfigMessage").textContent;
  orders.evaluate('stageFieldTypeChange("销售代表", "id"); applyFieldConfiguration()');
  assert.equal(fieldMessage(), "字段配置已应用；1 个非空值认不出，已从对应统计和图表中跳过。");
  orders.evaluate("restoreAutomaticFieldTypes()");
  assert.equal(fieldMessage(), "已恢复自动识别，统计、图表和分析结论已重新计算。");
  applyFilterState(orders, { category: { field: "地区", values: ["华东"] } });
  orders.evaluate('stageFieldTypeChange("销售代表", "id"); applyFieldConfiguration()');
  assert.equal(fieldMessage(), "字段配置已应用；1 个非空值认不出，已从对应统计和图表中跳过。 已清除现有筛选条件。");
  applyFilterState(orders, { category: { field: "地区", values: ["华东"] } });
  orders.evaluate("restoreAutomaticFieldTypes()");
  assert.equal(fieldMessage(), "已恢复自动识别，统计、图表和分析结论已重新计算。 已清除现有筛选条件。");
  applyFilterState(orders, { category: { field: "地区", values: ["华东"] } });
  orders.evaluate('confirmPlaceholder("金额", "N/A")');
  assert.match(placeholderMessage(orders), /所以行数还是 42。平均值等统计不变，因为「N\/A」本来就没有算进去。 已清除现有筛选条件。$/);
  orders.evaluate('undoPlaceholderConfirmation("金额", "N/A")');
  assert.doesNotMatch(placeholderMessage(orders), /已清除现有筛选条件/);
});

// ---- v3.1.0 commit 2: the rows counted in the analysis ----
// Three layers: every row of the file (state.rows) -> the rows counted in the analysis
// (getIncludedRows()) -> the filtered rows (getAnalysisRows()). No button yet: the tests call
// setRowExclusion() as the button will. The sample has its total on file row 44 and a
// repeated row on file row 21.

const SAMPLE_TOTAL_VALUES = ["123440", "123,440"];

async function sampleWithExclusions({ totals = true, duplicates = true } = {}) {
  const orders = await clickSampleButton(sampleOrdersFetchedText());
  if (totals) orders.evaluate('setRowExclusion("suspectedTotal", true)');
  if (duplicates) orders.evaluate('setRowExclusion("duplicateRows", true)');
  return orders;
}

// A rebuild that ends with every type as it was: 销售代表 to ID and back.
function rebuildKeepingTypes(scope) {
  scope.evaluate('stageFieldTypeChange("销售代表", "id"); applyFieldConfiguration(); stageFieldTypeChange("销售代表", "category"); applyFieldConfiguration()');
}

const excludedLines = (scope) => JSON.parse(scope.evaluate(`JSON.stringify({
  suspectedTotal: state.excludedRows.suspectedTotal.map((index) => getSourceRowNumber(state.rows[index])),
  duplicateRows: state.excludedRows.duplicateRows.map((index) => getSourceRowNumber(state.rows[index]))
})`));

test("v3.1.0 counted rows: without a click nothing is left out and the arrays are the v3.0.1 ones", async () => {
  const orders = await clickSampleButton(sampleOrdersFetchedText());
  assert.equal(orders.evaluate("getIncludedRows() === state.rows"), true);
  assert.equal(orders.evaluate("state.numericStats === state.qualityNumericStats"), true);
  assert.equal(orders.evaluate("getAnalysisRows().length"), 42);
  assert.deepEqual(excludedLines(orders), { suspectedTotal: [], duplicateRows: [] });
  assert.equal(orders.evaluate("state.suspectedTotals.length"), 1);
  rebuildKeepingTypes(orders);
  assert.equal(orders.evaluate("getIncludedRows() === state.rows && state.numericStats === state.qualityNumericStats"), true);
});

test("v3.1.0 counted rows: every exit counts 40 rows after both kinds are left out, through a rebuild too", async () => {
  const orders = await sampleWithExclusions();
  assert.deepEqual(excludedLines(orders), { suspectedTotal: [44], duplicateRows: [21] });
  const exits = () => JSON.parse(orders.evaluate(`JSON.stringify((() => {
    dom.v2MetricField.value = "金额"; dom.v2GroupField.value = "地区"; dom.v2AggregateMethod.value = "sum"; dom.v2DateField.value = "";
    renderV2Analysis();
    dom.categoryFilterField.value = "地区"; updateCategoryFilterField();
    return {
      included: getIncludedRows().length, analysis: getAnalysisRows().length,
      quantity: state.numericStats.find((item) => item.field === "数量").count,
      amount: state.numericStats.find((item) => item.field === "金额").count,
      region: state.categoryStats.find((item) => item.field === "地区").total,
      histogram: getNumericValues("数量").length,
      trend: buildDateTrend("下单日期", "").values.reduce((sum, value) => sum + value, 0),
      filterOptions: state.categoryFilterOptions.reduce((sum, option) => sum + option.count, 0),
      exportRows: getProcessedExportRows("filtered").length,
      report: buildReportData(new Date(2026, 10, 15)).dataScale,
      custom: state.customAnalysis.grouped.reduce((sum, item) => sum + item.count, 0),
      quality: getQualityNumericStats().find((item) => item.field === "数量").count
    };
  })())`));
  const expected = {
    included: 40, analysis: 40, quantity: 40, amount: 39, region: 40, histogram: 40, trend: 40, filterOptions: 40, exportRows: 40,
    report: { rows: 40, originalRows: 42, filteredOutRows: 0, excludedRows: 2, fields: 9, activeFields: 9, numericFields: 2, categoryFields: 5, dateFields: 1, idFields: 1, ignoredFields: 0 },
    custom: 39, quality: 42
  };
  assert.deepEqual(exits(), expected);
  rebuildKeepingTypes(orders);
  assert.equal(orders.evaluate("state.numericStats === state.qualityNumericStats"), false);
  assert.deepEqual(exits(), expected);
});

test("v3.1.0 counted rows: no analysis output shows the total row's 123440, before or after a rebuild", async () => {
  const orders = await clickSampleButton(sampleOrdersFetchedText());
  const chartConfigs = withFakeCharts(orders);
  orders.evaluate('dom.numericFilterField.value = "金额"; updateNumericFilterField(); setRowExclusion("suspectedTotal", true)');
  const outputs = () => orders.evaluate(`(() => {
    dom.v2MetricField.value = "金额"; dom.v2GroupField.value = "地区"; dom.v2AggregateMethod.value = "sum"; dom.v2DateField.value = "下单日期";
    renderCharts(); renderChartInsights(); renderV2Analysis(); renderNumericStats(); renderCategoryStats();
    dom.numericFilterField.value = "金额"; updateNumericFilterField(); dom.categoryFilterField.value = "地区"; updateCategoryFilterField();
    const report = buildReportData(new Date(2026, 10, 15));
    return [dom.numericStatsTable.innerHTML, dom.categoryStats.innerHTML, dom.v2PrimaryTable.innerHTML, dom.v2SumTable.innerHTML, dom.v2AvgTable.innerHTML,
      dom.v2ResultSummary.textContent, state.chartInsights.category, state.chartInsights.trend, JSON.stringify(report.numericStats),
      JSON.stringify(report.categoryStats), JSON.stringify(report.dateTrends), JSON.stringify(report.customAnalysis), dom.numericFilterRange.textContent,
      dom.numericFilterMax.value, JSON.stringify(state.categoryFilterOptions)].join("\\n");
  })()`) + JSON.stringify(chartConfigs.map((config) => config.data));
  const before = outputs();
  for (const value of SAMPLE_TOTAL_VALUES) assert.ok(!before.includes(value), value);
  assert.ok(chartConfigs.length >= 3, "charts were drawn");
  assert.ok(before.includes("7,200"), "the largest counted amount is there");
  rebuildKeepingTypes(orders);
  const after = outputs();
  for (const value of SAMPLE_TOTAL_VALUES) assert.ok(!after.includes(value), value);
  // The quality page still reads every row.
  assert.ok(orders.evaluate('getQualityNumericStats().find((item) => item.field === "金额").max') === 123440);
});

test("v3.1.0 counted rows: a rebuild keeps the totals still flagged, brings back the others once, and says when one is flagged again", async () => {
  const orders = await sampleWithExclusions();
  const fieldMessage = () => orders.elements.get("fieldConfigMessage").textContent;
  orders.evaluate('stageFieldTypeChange("数量", "ignore"); stageFieldTypeChange("金额", "ignore"); applyFieldConfiguration()');
  assert.equal(fieldMessage(), "字段配置已应用，统计、图表和分析结论已重新计算。 有 1 行不再是疑似合计行，已恢复计入分析。");
  assert.deepEqual(excludedLines(orders), { suspectedTotal: [], duplicateRows: [21] });
  assert.equal(orders.evaluate("getIncludedRows().length"), 41);
  orders.evaluate('stageFieldTypeChange("数量", "numeric"); stageFieldTypeChange("金额", "numeric"); applyFieldConfiguration()');
  assert.equal(fieldMessage(), "字段配置已应用；1 个非空值认不出，已从对应统计和图表中跳过。 第 44 行又被标为疑似合计行，目前仍计入分析，可以再设为不计入分析。");
  assert.deepEqual(excludedLines(orders), { suspectedTotal: [], duplicateRows: [21] });
  assert.equal(orders.evaluate("state.suspectedTotals.length"), 1);
  assert.equal(orders.evaluate("getIncludedRows().length"), 41);
  // Said once: the next rebuild is quiet.
  rebuildKeepingTypes(orders);
  assert.equal(fieldMessage(), "字段配置已应用；1 个非空值认不出，已从对应统计和图表中跳过。");
  // Excluding it again and losing it again restores it again; then it is not "flagged again"
  // after the user clicked, until it is restored once more.
  orders.evaluate('setRowExclusion("suspectedTotal", true)');
  assert.deepEqual(excludedLines(orders), { suspectedTotal: [44], duplicateRows: [21] });
  orders.evaluate('stageFieldTypeChange("数量", "category"); stageFieldTypeChange("金额", "category"); applyFieldConfiguration()');
  assert.match(fieldMessage(), / 有 1 行不再是疑似合计行，已恢复计入分析。$/);
  // A field configuration that keeps the total flagged keeps it out.
  const other = await sampleWithExclusions();
  other.evaluate('stageFieldTypeChange("数量", "category"); applyFieldConfiguration()');
  assert.deepEqual(excludedLines(other), { suspectedTotal: [44], duplicateRows: [21] });
  assert.doesNotMatch(other.elements.get("fieldConfigMessage").textContent, /恢复计入|又被标为/);
});

test("v3.1.0 counted rows: the click keeps filters and the custom analysis, rebuilds nothing, and is refused while drafts wait", async () => {
  const orders = await clickSampleButton(sampleOrdersFetchedText());
  orders.evaluate(`(() => {
    state.filters = { category: { field: "地区", values: ["华北", "华东", "华南", "西北", "西南"] }, numeric: null, date: null };
    state.filteredRows = filterRowsByConfig(getIncludedRows(), state.filters);
    state.filterSourceRows = getIncludedRows();
    refreshFilteredAnalysis();
    dom.v2MetricField.value = "金额"; dom.v2GroupField.value = "品类"; dom.v2AggregateMethod.value = "sum"; dom.v2DateField.value = "";
    renderV2Analysis();
    dom.v2AnalysisResults.classList.remove("hidden");
    globalThis.__calls = {};
    ["populateControls", "renderQuality", "rebuildAnalysisFromProfiles", "initializeFilterControls", "renderSchema", "refreshFilteredAnalysis"].forEach((name) => {
      const original = globalThis[name];
      __calls[name] = 0;
      globalThis[name] = function (...args) { __calls[name] += 1; return original.apply(this, args); };
    });
  })()`);
  assert.equal(orders.evaluate("getAnalysisRows().length"), 41);
  assert.equal(orders.evaluate('setRowExclusion("suspectedTotal", true)'), true);
  assert.deepEqual(JSON.parse(orders.evaluate("JSON.stringify(__calls)")), {
    populateControls: 0, renderQuality: 0, rebuildAnalysisFromProfiles: 0, initializeFilterControls: 0, renderSchema: 0, refreshFilteredAnalysis: 1
  });
  assert.equal(orders.evaluate("hasActiveFilters()"), true);
  // The total row is the only one with an empty 地区, so the filter keeps every counted row.
  assert.equal(orders.evaluate("getAnalysisRows().length"), 41);
  assert.equal(orders.evaluate("Boolean(state.customAnalysis)"), true);
  assert.equal(orders.evaluate("state.customAnalysis.grouped.reduce((sum, item) => sum + item.count, 0)"), 40);
  orders.evaluate('stageFieldTypeChange("销售代表", "id")');
  assert.equal(orders.evaluate('setRowExclusion("duplicateRows", true)'), false);
  assert.deepEqual(excludedLines(orders), { suspectedTotal: [44], duplicateRows: [] });
});

test("v3.1.0 counted rows: imports and resets clear the left-out rows; filters and the standardized check keep them", async () => {
  const orders = await sampleWithExclusions();
  orders.evaluate("checkStandardizedExport()");
  assert.equal(orders.evaluate("Boolean(state.standardizedExport)"), true);
  orders.evaluate('setRowExclusion("suspectedTotal", false)');
  assert.equal(orders.evaluate("Boolean(state.standardizedExport)"), true);
  orders.evaluate('setRowExclusion("suspectedTotal", true)');
  applyFilterState(orders, { category: { field: "地区", values: ["华东"] } });
  orders.evaluate("clearAllFilters()");
  assert.deepEqual(excludedLines(orders), { suspectedTotal: [44], duplicateRows: [21] });
  assert.equal(orders.evaluate("getAnalysisRows().length"), 40);
  orders.evaluate(`commitTabularData(["名称", "金额"], [{ 名称: "甲", 金额: "1" }, { 名称: "乙", 金额: "2" }], "next.csv", beginImport())`);
  assert.deepEqual(excludedLines(orders), { suspectedTotal: [], duplicateRows: [] });
  assert.equal(orders.evaluate("getIncludedRows() === state.rows"), true);
  orders.evaluate("resetAnalysisState()");
  assert.equal(orders.evaluate("state.exclusionSourceRows === null && state.includedRows === null && state.excludedRows.suspectedTotal.length === 0"), true);
});

test("v3.1.0 counted rows: left-out rows of another data set never count", async () => {
  const orders = await sampleWithExclusions();
  orders.evaluate("state.rows = state.rows.slice()");
  assert.equal(orders.evaluate("getIncludedRows() === state.rows"), true);
  assert.equal(orders.evaluate("getExcludedIndexSet().size"), 0);
  assert.equal(orders.evaluate("getAnalysisRows() === state.rows"), true);
});

test("v3.1.0 counted rows: an import builds each row key as often as v3.0.1 did", async () => {
  const counts = {};
  for (const [name, file] of [["sample-orders.csv", path.join(REPO_ROOT, "examples", "sample-orders.csv")],
    ["花名册.xlsx", path.join(STAGE3_FIXTURES, "花名册.xlsx")], ["sample-students.csv", path.join(__dirname, "fixtures", "examples", "sample-students.csv")]]) {
    const scope = createExcelContext();
    put(scope.context, "Papa", Papa);
    scope.context.window.Papa = Papa;
    let calls = 0;
    const original = scope.evaluate("buildExactRowKey");
    put(scope.context, "buildExactRowKey", (...args) => { calls += 1; return original(...args); });
    const bytes = fs.readFileSync(file);
    put(scope.context, "__countedFile", { name, size: bytes.length, arrayBuffer: async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.length) });
    await scope.evaluate("parseDataFile(__countedFile)");
    counts[name] = calls;
  }
  // v3.0.1: one key per row for the repeated rows, plus the identical-row check of duplicated IDs.
  assert.deepEqual(counts, { "sample-orders.csv": 44, "花名册.xlsx": 33, "sample-students.csv": 120 });
});

test("v3.1.0 counted rows: findRepeatedRowIndexes() is what the repeated-row exclusion and the duplicate count share", async () => {
  const orders = await clickSampleButton(sampleOrdersFetchedText());
  const repeated = JSON.parse(orders.evaluate("JSON.stringify(findRepeatedRowIndexes(state.rows, state.fields))"));
  assert.equal(repeated.length, orders.evaluate("state.duplicateRows"));
  assert.deepEqual(repeated.map((index) => orders.evaluate(`getSourceRowNumber(state.rows[${index}])`)), [21]);
  assert.equal(orders.evaluate("typeof deduplicateRows"), "undefined");
});

test("v3.1.0 counted rows: with rows left out, a placeholder message counts the rows of the analysis", async () => {
  const orders = await sampleWithExclusions();
  orders.evaluate('confirmPlaceholder("金额", "N/A")');
  assert.equal(placeholderMessage(orders), "已把「金额」的 1 个「N/A」按空值处理：空值 0 → 1，认不出 1 → 0。只是把这 1 格当作空，这一行没有删，其他字段照常参与分析，所以参与分析的行还是 40。平均值等统计不变，因为「N/A」本来就没有算进去。");
  orders.evaluate('undoPlaceholderConfirmation("金额", "N/A")');
  assert.equal(placeholderMessage(orders), "已撤销「金额」的「N/A」：空值 1 → 0，认不出 0 → 1。这 1 格重新按原来的文字处理，参与分析的行还是 40。");
  orders.evaluate("confirmAllCommonPlaceholders()");
  assert.equal(placeholderMessage(orders), "已把 2 处常见写法按空值处理（「金额」的「N/A」、「销售代表」的「—」）。只是把这些格当作空，行都没有删，所以参与分析的行还是 40。");
});

test("v3.1.0 counted rows: the numeric filter inputs follow the counted rows unless that field is filtered", async () => {
  const orders = await clickSampleButton(sampleOrdersFetchedText());
  const inputs = () => [orders.elements.get("numericFilterMin").value, orders.elements.get("numericFilterMax").value, orders.elements.get("numericFilterRange").textContent];
  orders.evaluate('dom.numericFilterField.value = "金额"; updateNumericFilterField()');
  assert.deepEqual(inputs(), ["480", "123440", "480 至 123,440"]);
  orders.evaluate('setRowExclusion("suspectedTotal", true)');
  assert.deepEqual(inputs(), ["480", "7200", "480 至 7,200"]);
  orders.evaluate('setRowExclusion("suspectedTotal", false)');
  assert.deepEqual(inputs(), ["480", "123440", "480 至 123,440"]);
  // Applied on 金额: the user's numbers stay, the hint follows.
  orders.evaluate(`(() => {
    dom.numericFilterMin.value = "100"; dom.numericFilterMax.value = "200000";
    state.filters = { category: null, numeric: { field: "金额", min: 100, max: 200000 }, date: null };
    state.filteredRows = filterRowsByConfig(getIncludedRows(), state.filters);
    state.filterSourceRows = getIncludedRows();
    refreshFilteredAnalysis();
  })()`);
  orders.evaluate('setRowExclusion("suspectedTotal", true)');
  assert.deepEqual(inputs(), ["100", "200000", "480 至 7,200"]);
});

// ---- v3.1.0 commit 3: the buttons and what the page says about rows left out ----

const exclusionClick = (scope, kind, on) => scope.evaluate(`handleExclusionClick({ target: { closest: () => ({ dataset: { exclusionKind: "${kind}", exclusionOn: "${on}" } }) } })`);
const textOf = (scope, id) => scope.elements.get(id).textContent;
const htmlText = (scope, id) => scope.elements.get(id).innerHTML.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
const isHidden = (scope, id) => scope.elements.get(id).classList.contains("hidden");

// A normal task table: names sometimes end in 汇总, hours are 2, 4 or 8 (review section 5).
function taskTableCsv() {
  let seed = 20261115;
  const random = () => { seed = (seed * 1103515245 + 12345) % 2147483648; return seed / 2147483648; };
  const pick = (list) => list[Math.floor(random() * list.length)];
  const names = ["整理需求", "评审方案", "编写代码", "测试", "部署", "周报汇总", "数据汇总", "月度汇总", "客户回访", "资料归档"];
  const people = ["张敏", "李强", "王芳", "赵磊", "陈静"];
  const lines = ["任务名,负责人,工时,完成日期"];
  for (let i = 0; i < 300; i += 1) {
    const day = String(1 + Math.floor(random() * 28)).padStart(2, "0");
    lines.push(`${pick(names)},${pick(people)},${pick([2, 4, 8])},2026-0${1 + Math.floor(random() * 9)}-${day}`);
  }
  return `${lines.join("\n")}\n`;
}

test("v3.1.0 quality page: every count stays when rows are left out; problem rows only add （不计入分析）", async () => {
  const orders = await clickSampleButton(sampleOrdersFetchedText());
  const snapshot = () => JSON.stringify({
    pills: sampleOrdersPills(orders),
    tables: ["qualityTable", "outlierTable", "idDuplicateDetails", "placeholderDetails", "formulaErrorDetails", "dateStyleDetails", "lookAlikeDetails", "constantFieldDetails"].map((id) => orders.elements.get(id).innerHTML),
    totalMissing: orders.evaluate("state.totalMissing"),
    lists: JSON.parse(orders.evaluate("JSON.stringify(listProblemCategories().map((category) => [category.kind, category.field || '', listProblemRows(category).items.map((item) => item.line)]))"))
  });
  const before = snapshot();
  exclusionClick(orders, "suspectedTotal", true);
  exclusionClick(orders, "duplicateRows", true);
  assert.equal(snapshot(), before);
  orders.evaluate('openProblemView({ kind: "duplicateRows" })');
  const notes = [...orders.elements.get("problemRowsTable").innerHTML.matchAll(/<tr><td>(\d+)<\/td><td>([^<]*)<\/td>/g)].map((match) => [match[1], match[2]]);
  assert.deepEqual(notes, [["20", "第 1 组（保留）"], ["21", "第 1 组（不计入分析）"]]);
  const csv = orders.evaluate('buildProblemRowsCsvContent(listProblemRows({ kind: "suspectedTotal" }), state.fields)');
  assert.match(csv, /^44,数量、金额等于第 3–43 行之和（41 行）（不计入分析）,合计,/m);
});

test("v3.1.0 quality page: the two buttons, their sentences and their own messages", async () => {
  const orders = await clickSampleButton(sampleOrdersFetchedText());
  const totalsFooter = () => orders.elements.get("suspectedTotalDetails").innerHTML;
  assert.match(totalsFooter(), /data-exclusion-kind="suspectedTotal" data-exclusion-on="true"[^>]*>不计入分析<\/button>/);
  assert.equal(isHidden(orders, "suspectedTotalExcludedNote"), true);
  assert.equal(orders.elements.get("suspectedTotalSection").classList.contains("is-excluded"), false);
  assert.equal(isHidden(orders, "qualityScopeNote"), true);
  assert.equal(htmlText(orders, "duplicateExclusionAction"), "不计入分析");
  assert.equal(textOf(orders, "duplicateExclusionText"), "重复行里多出来的 1 行目前仍计入分析。");
  assert.equal(orders.evaluate("formatDuplicateIncludedNote(1234)"), "重复行里多出来的 1,234 行目前仍计入分析。");

  exclusionClick(orders, "suspectedTotal", true);
  assert.equal(textOf(orders, "suspectedTotalMessage"), "已把 1 行疑似合计行设为不计入分析，参与分析的行 42 → 41。");
  assert.equal(textOf(orders, "duplicateExclusionMessage"), "");
  assert.equal(orders.elements.get("suspectedTotalSection").classList.contains("is-excluded"), true);
  assert.equal(isHidden(orders, "suspectedTotalExcludedNote"), false);
  assert.equal(textOf(orders, "suspectedTotalExcludedNote"), "这 1 行已不计入分析：统计、图表、自定义分析和筛选不再包含它。");
  assert.match(totalsFooter(), /data-exclusion-on="false"[^>]*>撤销<\/button>/);
  assert.equal(textOf(orders, "qualityScopeNote"), "质量页检查的是文件里的全部 42 行，包括不计入分析的 1 行（其中有 6 个空值）；偏大或偏小的值也按这 42 行算。");

  exclusionClick(orders, "duplicateRows", true);
  assert.equal(textOf(orders, "duplicateExclusionMessage"), "已把 1 行重复的行设为不计入分析，参与分析的行 41 → 40。");
  assert.equal(textOf(orders, "duplicateExclusionText"), "多出来的 1 行已不计入分析。");
  assert.equal(htmlText(orders, "duplicateExclusionAction"), "撤销");
  assert.equal(textOf(orders, "qualityScopeNote"), "质量页检查的是文件里的全部 42 行，包括不计入分析的 2 行（其中有 6 个空值）；偏大或偏小的值也按这 42 行算。");

  exclusionClick(orders, "duplicateRows", false);
  assert.equal(textOf(orders, "duplicateExclusionText"), "重复行里多出来的 1 行目前仍计入分析。");
  assert.equal(htmlText(orders, "duplicateExclusionAction"), "不计入分析");
  exclusionClick(orders, "duplicateRows", true);

  exclusionClick(orders, "suspectedTotal", false);
  assert.equal(textOf(orders, "suspectedTotalMessage"), "已撤销：1 行疑似合计行重新计入分析，参与分析的行 40 → 41。");
  assert.equal(textOf(orders, "qualityScopeNote"), "质量页检查的是文件里的全部 42 行，包括不计入分析的 1 行；偏大或偏小的值也按这 42 行算。");
  assert.equal(isHidden(orders, "suspectedTotalExcludedNote"), true);

  // Drafts: refused in the section that was clicked, nothing changes.
  orders.evaluate('stageFieldTypeChange("销售代表", "id")');
  exclusionClick(orders, "suspectedTotal", true);
  assert.equal(textOf(orders, "suspectedTotalMessage"), "字段配置有未应用的修改。请先点击“应用字段配置”，再设置不计入分析。");
  assert.equal(orders.evaluate("state.excludedRows.suspectedTotal.length"), 0);
  // A rebuild clears both messages.
  orders.evaluate("applyFieldConfiguration()");
  assert.equal(textOf(orders, "suspectedTotalMessage"), "");
  assert.equal(textOf(orders, "duplicateExclusionMessage"), "");
});

test("v3.1.0 quality page: the totals section of 分区小计表 names the totals that only equal the row above", async () => {
  const region = await importStage3Fixture("分区小计表.csv");
  const single = "其中 1 行只等于上面一行，可能不是合计行，先点“查看”确认。";
  assert.ok(htmlText(region, "suspectedTotalDetails").includes(single));
  region.evaluate('setRowExclusion("suspectedTotal", true)');
  assert.ok(htmlText(region, "suspectedTotalDetails").includes(single));
  assert.equal(textOf(region, "suspectedTotalExcludedNote"), "这 5 行已不计入分析：统计、图表、自定义分析和筛选不再包含它们。");
  for (const scope of [await clickSampleButton(sampleOrdersFetchedText()), await importStage3Fixture("门店报表.xlsx")]) {
    assert.doesNotMatch(htmlText(scope, "suspectedTotalDetails"), /只等于上面一行/);
  }
});

test("v3.1.0 quality page: 另有 N 行 names the labelled rows whose numbers do not add up, only where something was checked", async () => {
  const tasks = await importCsvFile(taskTableCsv(), "task-table.csv");
  assert.equal(tasks.evaluate("state.totalLabelRows.length"), 95);
  assert.equal(tasks.evaluate("state.suspectedTotals.length"), 6);
  assert.ok(htmlText(tasks, "suspectedTotalDetails").includes("另有 89 行也带这类字样，但数字对不上，没有列出。"));
  assert.ok(htmlText(tasks, "suspectedTotalDetails").includes("其中 6 行只等于上面一行，可能不是合计行，先点“查看”确认。"));
  tasks.evaluate('stageFieldTypeChange("工时", "ignore"); applyFieldConfiguration()');
  assert.equal(isHidden(tasks, "suspectedTotalSection"), true);
  assert.equal(tasks.evaluate("formatUnmatchedTotalLabels(95, 0, false)"), "");
  for (const scope of [await clickSampleButton(sampleOrdersFetchedText()), await importStage3Fixture("分区小计表.csv"), await importStage3Fixture("门店报表.xlsx")]) {
    assert.doesNotMatch(htmlText(scope, "suspectedTotalDetails"), /另有 \d+ 行也带这类字样/);
  }
});

test("v3.1.0 quality page: a message counts only the rows this click left out or brought back", () => {
  const scope = createScriptContext();
  const rows = [["甲", "10"], ["乙", "20"], ["丙", "30"], ["合计", "60"], ["合计", "60"]].map(([名称, 金额]) => ({ 名称, 金额 }));
  commitRows(scope, ["名称", "金额"], rows, "twice.csv");
  assert.equal(scope.evaluate("state.suspectedTotals.length"), 2);
  assert.equal(scope.evaluate("state.duplicateRows"), 1);
  // No column tells the records apart: the long note stays, before and after.
  assert.equal(textOf(scope, "duplicateExclusionText"), scope.evaluate("NO_ID_DUPLICATE_NOTE[0]"));
  exclusionClick(scope, "suspectedTotal", true);
  assert.equal(textOf(scope, "suspectedTotalMessage"), "已把 2 行疑似合计行设为不计入分析，参与分析的行 5 → 3。");
  exclusionClick(scope, "duplicateRows", true);
  assert.equal(textOf(scope, "duplicateExclusionMessage"), "多出来的 1 行已经作为疑似合计行不计入分析，参与分析的行仍是 3。");
  exclusionClick(scope, "suspectedTotal", false);
  assert.equal(textOf(scope, "suspectedTotalMessage"), "已撤销：1 行疑似合计行重新计入分析，参与分析的行 3 → 4。");
});

test("v3.1.0 page: scope sentence, filter panel, overview, preview, sidebar and custom analysis count the rows left out", async () => {
  const orders = await clickSampleButton(sampleOrdersFetchedText());
  orders.evaluate('dom.v2MetricField.value = "金额"; dom.v2GroupField.value = "品类"; dom.v2AggregateMethod.value = "sum"; dom.v2DateField.value = ""; renderV2Analysis(); dom.v2AnalysisResults.classList.remove("hidden")');
  const view = () => ({
    scope: textOf(orders, "globalFilterSummary"),
    counts: ["filterOriginalCount", "filterCurrentCount", "filterExcludedCount", "filterNotCountedCount"].map((id) => textOf(orders, id)),
    notCountedHidden: isHidden(orders, "filterNotCountedItem"),
    filterMessage: textOf(orders, "filterMessage"),
    overview: htmlText(orders, "overviewCards"),
    preview: textOf(orders, "previewPageInfo"),
    // The sidebar says it once, in its 筛选 line; it has no row of its own for the rows left out.
    sidebar: textOf(orders, "sidebarFilterSummary"),
    custom: textOf(orders, "v2ResultSummary")
  });
  assert.deepEqual(view(), {
    scope: "全部 42 行", counts: ["42", "42", "0", "0"], notCountedHidden: true, filterMessage: "尚未应用筛选，当前分析使用全部记录。",
    overview: "总记录数 42 字段数 9 数值字段数 2 分类字段数 5 空值概览 6", preview: "第 1 / 1 页，共 42 行", sidebar: "全部 42 行",
    custom: "已按 品类 对 金额 进行求和，结果基于完整数据 42 行。指标字段「金额」有 1 个值认不出是数字，已从计算中排除。"
  });
  exclusionClick(orders, "suspectedTotal", true);
  exclusionClick(orders, "duplicateRows", true);
  assert.deepEqual(view(), {
    scope: "40 行（另有 2 行不计入分析）", counts: ["42", "40", "0", "2"], notCountedHidden: false, filterMessage: "当前分析使用全部 40 行（不含不计入分析的 2 行）。",
    overview: "总记录 → 当前分析 42 → 40 字段数 9 数值字段数 2 分类字段数 5 空值概览 0", preview: "第 1 / 1 页，共 40 行（不含不计入分析的 2 行）", sidebar: "40 行（另有 2 行不计入分析）",
    custom: "已按 品类 对 金额 进行求和，结果基于 40 行（另有 2 行不计入分析）。指标字段「金额」有 1 个值认不出是数字，已从计算中排除。"
  });
  orders.evaluate('dom.categoryFilterField.value = "地区"; updateCategoryFilterField(); state.categoryFilterDraftValues = new Set(["华东", "华南"]); applyFilters()');
  assert.deepEqual(view(), {
    scope: "已筛选 1 项 · 17 / 40 行（另有 2 行不计入分析）", counts: ["42", "17", "23", "2"], notCountedHidden: false, filterMessage: "已应用 1 个筛选条件，保留 17 行，筛除 23 行。",
    overview: "总记录 → 当前分析 42 → 17 字段数 9 数值字段数 2 分类字段数 5 空值概览 0", preview: "第 1 / 1 页，共 17 行", sidebar: "已筛选 1 项 · 17 / 40 行（另有 2 行不计入分析）",
    custom: "已按 品类 对 金额 进行求和，结果基于当前筛选数据 17 / 40 行（另有 2 行不计入分析）。"
  });
  // A click while filtered writes the filter message again with the new numbers.
  exclusionClick(orders, "duplicateRows", false);
  const kept = orders.evaluate("getAnalysisRows().length");
  assert.equal(textOf(orders, "filterMessage"), `已应用 1 个筛选条件，保留 ${kept} 行，筛除 ${41 - kept} 行。`);
  assert.equal(textOf(orders, "filterExcludedCount"), String(41 - kept));
  orders.evaluate("clearAllFilters()");
  assert.equal(textOf(orders, "filterMessage"), "已清除全部筛选，当前分析恢复使用全部 41 行（不含不计入分析的 1 行）。");
  // 当前状态 keeps its three rows: 筛选, 表头, 已按空值处理.
  const html = indexHtmlSource();
  const status = html.slice(html.indexOf('<dl class="sidebar-status">'), html.indexOf("</dl>", html.indexOf('<dl class="sidebar-status">')));
  assert.deepEqual([...status.matchAll(/<dt>([^<]*)<\/dt>/g)].map((match) => match[1]), ["筛选", "表头", "已按空值处理"]);
});

test("v3.1.0 filters: 筛除 counts from the rows counted in the analysis, in the message and in the panel", async () => {
  const orders = await sampleWithExclusions();
  orders.evaluate('dom.categoryFilterField.value = "地区"; updateCategoryFilterField(); state.categoryFilterDraftValues = new Set(["华东", "华南"]); applyFilters()');
  assert.equal(orders.evaluate("getAnalysisRows().length"), 17);
  assert.equal(textOf(orders, "filterMessage"), "已应用 1 个筛选条件，保留 17 行，筛除 23 行。");
  assert.equal(textOf(orders, "filterExcludedCount"), "23");
});

test("v3.1.0 filters: a filter result made from other rows falls back to the 40 counted rows, never all 42", async () => {
  const orders = await sampleWithExclusions();
  orders.evaluate("state.filterSourceRows = state.rows.slice(0, 3); state.filteredRows = state.rows.slice(0, 3)");
  assert.equal(orders.evaluate("getAnalysisRows().length"), 40);
  assert.equal(orders.evaluate("getAnalysisRows() === getIncludedRows()"), true);
  orders.evaluate("state.filterSourceRows = state.rows");
  assert.equal(orders.evaluate("getAnalysisRows().length"), 40);
});

test("v3.1.0 status bar: the totals sentence offers to leave them out and never says 仍计入统计 again", async () => {
  const orders = await clickSampleButton(sampleOrdersFetchedText());
  assert.equal(textOf(orders, "statusPanel"), "sample-orders.csv 已成功载入，分析结果已生成。表头在第 2 行，已跳过上方 1 行：「2025年上半年销售订单明细」。第 44 行「合计」疑似合计行，可以在“数据质量”里设为不计入分析。");
  const oldEnding = "目前仍计入统计，详情见“数据质量”";
  orders.evaluate('setRowExclusion("suspectedTotal", true)');
  const { html } = reportOf(orders);
  for (const text of [textOf(orders, "statusPanel"), orders.elements.get("qualitySection").innerHTML, htmlText(orders, "suspectedTotalDetails"), html]) {
    assert.ok(!text.includes(oldEnding));
  }
  const region = await importStage3Fixture("分区小计表.csv");
  assert.ok(!textOf(region, "statusPanel").includes(oldEnding));
  assert.ok(textOf(region, "statusPanel").endsWith("发现 5 行疑似合计 / 小计（第 5、8、10、14、15 行），可以在“数据质量”里设为不计入分析。"));
});

test("v3.1.0 empty states: without a filter they no longer say 当前筛选", () => {
  const scope = createScriptContext();
  withFakeCharts(scope);
  const emptyTexts = collectEmptyChartTexts(scope);
  const rows = [...Array.from({ length: 6 }, (_, i) => ({ 名称: `商品${i + 1}`, 金额: String((i + 1) * 10), 备注: "" })), { 名称: "合计", 金额: "210", 备注: "手工合计" }];
  commitRows(scope, ["名称", "金额", "备注"], rows, "only-total-note.csv");
  scope.evaluate('setRowExclusion("suspectedTotal", true)');
  emptyTexts();
  scope.evaluate('dom.categoryChartField.value = "备注"; renderCategoryTopChart()');
  assert.ok(scope.elements.get("categoryStats").innerHTML.includes("当前分析的行里没有非空类别值。"));
  assert.deepEqual(emptyTexts(), [["categoryTopChart", "当前分析的行里没有可统计的分类值"]]);
  assert.doesNotMatch(scope.elements.get("categoryStats").innerHTML, /当前筛选/);
});

const CONTAINS_NOTE_IDS = ["statisticsContainsNote", "chartsContainsNote", "customContainsNote"];
const containsNotes = (scope) => CONTAINS_NOTE_IDS.map((id) => (isHidden(scope, id) ? null : htmlText(scope, id)));
const containsTarget = (scope) => (scope.elements.get("statisticsContainsNote").innerHTML.match(/data-quality-target="(\w+)" aria-label="([^"]*)"/) || []).slice(1);
const filterRegion = (scope, ...regions) => scope.evaluate(`dom.categoryFilterField.value = "地区"; updateCategoryFilterField(); state.categoryFilterDraftValues = new Set(${JSON.stringify(regions)}); applyFilters()`);

test("v3.1.0 modules: statistics, charts and custom analysis say what the filtered rows still contain", async () => {
  const orders = await clickSampleButton(sampleOrdersFetchedText());
  const both = "这些结果包含 1 行疑似合计行、1 行重复的行。 去数据质量处理";
  assert.deepEqual(containsNotes(orders), [both, both, both]);
  assert.deepEqual(containsTarget(orders), ["suspectedTotal", "去数据质量处理疑似合计行和重复的行"]);
  // The import built no list of repeated rows for it: state.duplicateRows says how many.
  assert.equal(orders.evaluate("state.repeatedRowSource === state.rows"), false);

  // By the filtered rows: 华东 has neither the total (no region) nor the repeated row (华南).
  filterRegion(orders, "华东");
  assert.deepEqual(containsNotes(orders), [null, null, null]);
  assert.equal(orders.elements.get("statisticsContainsNote").innerHTML, "");
  filterRegion(orders, "华南");
  const repeated = "这些结果包含 1 行重复的行。 去数据质量处理";
  assert.deepEqual(containsNotes(orders), [repeated, repeated, repeated]);
  assert.deepEqual(containsTarget(orders), ["duplicateRows", "去数据质量处理重复的行"]);
  orders.evaluate("clearAllFilters()");
  assert.deepEqual(containsNotes(orders), [both, both, both]);

  // Rows left out are no longer contained; both kinds out, no note.
  orders.evaluate('setRowExclusion("suspectedTotal", true)');
  assert.deepEqual(containsNotes(orders), [repeated, repeated, repeated]);
  orders.evaluate('setRowExclusion("duplicateRows", true)');
  assert.deepEqual(containsNotes(orders), [null, null, null]);
  orders.evaluate('setRowExclusion("suspectedTotal", false)');
  const total = "这些结果包含 1 行疑似合计行。 去数据质量处理";
  assert.deepEqual(containsNotes(orders), [total, total, total]);
  assert.deepEqual(containsTarget(orders), ["suspectedTotal", "去数据质量处理疑似合计行"]);

  // Without a clean ID column the repeated rows are not named (they may be different records).
  const noId = await clickSampleButton(sampleOrdersFetchedText());
  noId.evaluate('stageFieldTypeChange("订单编号", "category"); applyFieldConfiguration()');
  assert.deepEqual(containsNotes(noId), [total, total, total]);
  noId.evaluate("resetAnalysisState()");
  assert.deepEqual(containsNotes(noId), [null, null, null]);
  assert.equal(noId.elements.get("customContainsNote").innerHTML, "");
  assert.equal(formatAnalysisContainsNoteOf(noId, 0, 0), "");
  assert.equal(formatAnalysisContainsNoteOf(noId, 1234, 5678), "这些结果包含 1,234 行疑似合计行、5,678 行重复的行。");
});

test("v3.1.0 modules: the repeated-row list is built once per import, however often the filters change", async () => {
  const orders = await clickSampleButton(sampleOrdersFetchedText());
  orders.evaluate("globalThis.__repeatedListCalls = 0; const __find = findRepeatedRowIndexes; findRepeatedRowIndexes = (...args) => { globalThis.__repeatedListCalls += 1; return __find(...args); }");
  const calls = () => orders.evaluate("globalThis.__repeatedListCalls");
  assert.equal(calls(), 0);
  filterRegion(orders, "华南");
  assert.equal(calls(), 1);
  filterRegion(orders, "华东");
  orders.evaluate("clearAllFilters()");
  exclusionClick(orders, "duplicateRows", true);
  assert.equal(orders.evaluate("state.excludedRows.duplicateRows.length"), 1);
  assert.equal(calls(), 1);
  // Another file: built again, once.
  const rows = [["A-1", "甲", "10"], ["A-2", "乙", "20"], ["A-2", "乙", "20"], ...Array.from({ length: 12 }, (_, i) => [`A-${i + 3}`, "丙", String(30 + i)])].map(([订单ID, 名称, 金额]) => ({ 订单ID, 名称, 金额 }));
  commitRows(orders, ["订单ID", "名称", "金额"], rows, "another.csv");
  assert.equal(orders.evaluate("state.duplicateRows"), 1);
  // A clean ID column: the notes need the list once a filter applies.
  assert.equal(orders.evaluate("hasCleanIdField()"), true);
  assert.equal(calls(), 1);
  orders.evaluate('dom.categoryFilterField.value = "名称"; updateCategoryFilterField(); state.categoryFilterDraftValues = new Set(["乙"]); applyFilters()');
  assert.equal(calls(), 2);
  assert.deepEqual(containsNotes(orders), Array(3).fill("这些结果包含 1 行重复的行。 去数据质量处理"));
  orders.evaluate('state.categoryFilterDraftValues = new Set(["乙", "丙"]); applyFilters()');
  orders.evaluate('setRowExclusion("duplicateRows", true)');
  assert.equal(calls(), 2);
});

function formatAnalysisContainsNoteOf(scope, totals, duplicates) {
  return scope.evaluate(`formatAnalysisContainsNote(${totals}, ${duplicates})`);
}

test("v3.1.0 modules: 去数据质量处理 opens the quality module at the subsection and focuses it", async () => {
  const orders = await clickSampleButton(sampleOrdersFetchedText());
  const calls = [];
  const record = (id, name) => { orders.elements.get(id)[name] = (options) => calls.push(`${id}.${name}${options ? ` ${JSON.stringify(options)}` : ""}`); };
  record("suspectedTotalSection", "scrollIntoView");
  record("suspectedTotalTitle", "focus");
  record("duplicateExclusion", "scrollIntoView");
  const duplicateButton = { focus: (options) => calls.push(`duplicate button.focus ${JSON.stringify(options)}`) };
  orders.elements.get("duplicateExclusionAction").querySelector = (selector) => (selector === "button[data-exclusion-kind]" ? duplicateButton : null);
  const click = (target) => orders.evaluate(`showModule("statisticsSection"); handleAnalysisNoteClick({ target: { closest: () => ({ dataset: { qualityTarget: "${target}" } }) } })`);

  click("suspectedTotal");
  assert.equal(orders.evaluate("document.body.dataset.module"), "qualitySection");
  assert.deepEqual(calls, ['suspectedTotalSection.scrollIntoView {"block":"start"}', 'suspectedTotalTitle.focus {"preventScroll":true}']);
  calls.length = 0;
  click("duplicateRows");
  assert.equal(orders.evaluate("document.body.dataset.module"), "qualitySection");
  assert.deepEqual(calls, ['duplicateExclusion.scrollIntoView {"block":"start"}', 'duplicate button.focus {"preventScroll":true}']);
  // A click beside the button does nothing.
  calls.length = 0;
  orders.evaluate('showModule("chartsSection"); handleAnalysisNoteClick({ target: { closest: () => null } })');
  assert.equal(orders.evaluate("document.body.dataset.module"), "chartsSection");
  assert.deepEqual(calls, []);

  const html = indexHtmlSource();
  assert.match(html, /<h3 id="suspectedTotalTitle" class="subsection-title" tabindex="-1">疑似合计 \/ 小计行<\/h3>/);
  for (const [section, id] of [["statisticsSection", "statisticsContainsNote"], ["chartsSection", "chartsContainsNote"], ["customSection", "customContainsNote"]]) {
    const start = html.indexOf(`<section id="${section}"`);
    const heading = html.slice(html.indexOf('<div class="section-heading">', start), html.indexOf(`id="${id}"`, start));
    // v3.2.0 commit 10: a bar of its own right under the heading (the range sentence is the
    // heading's last child), before the module's first content.
    assert.match(heading, /<\/h2>\s*<\/div>\s*<p class="filter-scope">当前分析范围：[^\n]*<\/p>\s*<\/div>\s*<p $/, section);
    assert.match(html.slice(html.indexOf(`id="${id}"`)), /^id="[^"]+" class="analysis-contains-note hidden"><\/p>\s*<(div|p) class=/, section);
  }
  assert.match(repoFile("script.js"), /\[dom\.statisticsContainsNote, dom\.chartsContainsNote, dom\.customContainsNote\]\.forEach\(\(note\) => note\.addEventListener\("click", handleAnalysisNoteClick\)\)/);
});

test("v3.1.0 quality page: a total whose columns match different ranges is not one that only equals the row above", () => {
  const totalsOf = (rows) => {
    const scope = createScriptContext();
    commitRows(scope, ["名称", "数量", "金额"], rows.map(([名称, 数量, 金额]) => ({ 名称, 数量, 金额 })), "two-ranges.csv");
    return scope;
  };
  // 合计: 数量 40 is only 丙 above it (the segment after 小计); 金额 15 is 甲 + 乙 + 丙 (5 + 7 + 3).
  const mixed = totalsOf([["甲", "10", "5"], ["乙", "20", "7"], ["小计", "30", "12"], ["丙", "40", "3"], ["合计", "40", "15"]]);
  const matches = JSON.parse(mixed.evaluate("JSON.stringify(state.suspectedTotals.map((item) => item.matches.map((match) => [match.scope, match.fields.join(), match.rowCount])))"));
  assert.deepEqual(matches, [[["segment", "数量,金额", 2]], [["segment", "数量", 1], ["above", "金额", 3]]]);
  assert.doesNotMatch(htmlText(mixed, "suspectedTotalDetails"), /只等于上面一行/);
  // Control: with 金额 3, both columns of 合计 only equal 丙, and the sentence names it.
  const single = totalsOf([["甲", "10", "5"], ["乙", "20", "7"], ["小计", "30", "12"], ["丙", "40", "3"], ["合计", "40", "3"]]);
  assert.ok(htmlText(single, "suspectedTotalDetails").includes("其中 1 行只等于上面一行，可能不是合计行，先点“查看”确认。"));
});

test("v3.1.0 quality page: the repeated-rows message counts only the rows this click adds, not the total rows already out", () => {
  const scope = createScriptContext();
  const rows = [["甲", "10"], ["乙", "20"], ["乙", "20"], ["丙", "30"], ["合计", "80"], ["合计", "80"]].map(([名称, 金额]) => ({ 名称, 金额 }));
  commitRows(scope, ["名称", "金额"], rows, "totals-and-repeats.csv");
  assert.equal(scope.evaluate("state.suspectedTotals.length"), 2);
  assert.equal(scope.evaluate("state.duplicateRows"), 2);
  exclusionClick(scope, "suspectedTotal", true);
  assert.equal(textOf(scope, "suspectedTotalMessage"), "已把 2 行疑似合计行设为不计入分析，参与分析的行 6 → 4。");
  exclusionClick(scope, "duplicateRows", true);
  assert.equal(textOf(scope, "duplicateExclusionMessage"), "已把 1 行重复的行设为不计入分析，参与分析的行 4 → 3。");
});

const REPORT_EXCLUDED_KEYS = ["dataScale.excludedRows", "exclusions", "quality.excludedScope", "quality.suspectedTotalsExcluded", "quality.suspectedTotalExtras"];
const scaleCard = (label, value) => `<div class="scale-card"><div class="label">${label}</div><div class="value">${value}</div></div>`;

test("v3.1.0 report: the rows left out are named in the data scale, the filters, the quality part and the totals section", async () => {
  const orders = await sampleWithExclusions();
  const { report, html } = reportOf(orders);
  assert.deepEqual([report.dataScale.rows, report.dataScale.originalRows, report.dataScale.filteredOutRows, report.dataScale.excludedRows], [40, 42, 0, 2]);
  assert.deepEqual(report.exclusions, { suspectedTotal: { count: 1, lines: [44] }, duplicateRows: { count: 1, lines: [21] }, overlap: null });
  assert.deepEqual(report.quality.excludedScope, { totalRows: 42, excludedRows: 2, excludedMissing: 6 });
  assert.equal(report.quality.suspectedTotalsExcluded, 1);

  const left = "不计入分析：疑似合计行 1 行（第 44 行）、重复的行 1 行（第 21 行）。";
  const scope = "质量页检查的是文件里的全部 42 行，包括不计入分析的 2 行（其中有 6 个空值）；偏大或偏小的值也按这 42 行算。";
  const excludedNote = "这 1 行已不计入分析：统计、图表、自定义分析和筛选不再包含它。";
  // HTML: a card after 筛除行数; the sentence between the filter text and the quality scope text;
  // the page's sentence under the quality heading; the totals note as on the page.
  assert.ok(html.includes(`${scaleCard("筛除行数", 0)}${scaleCard("不计入分析行数", 2)}${scaleCard("字段总数", 9)}`));
  const noFilter = "未设置筛选，结果基于全部 40 行（不含不计入分析的 2 行）。";
  assert.match(html, new RegExp(`<p>${noFilter}<\\/p>\\s*<p>${left}<\\/p>\\s*<p>数据质量部分始终基于原始数据，不受筛选影响。<\\/p>`));
  assert.match(html, new RegExp(`<h2>数据质量报告<\\/h2>\\s*<p>${scope}<\\/p>\\s*<p>空值总数：`));
  assert.match(html, new RegExp(`<h3>疑似合计 \\/ 小计行<\\/h3>\\s*<p>${excludedNote}<\\/p>`));
  assert.ok(!html.includes(orders.evaluate("SUSPECTED_TOTAL_NOTE[0]")));
  // The page shows the same sentences.
  assert.equal(textOf(orders, "qualityScopeNote"), scope);
  assert.equal(textOf(orders, "suspectedTotalExcludedNote"), excludedNote);

  // Filtered as well: 筛除 counts from the 40 counted rows.
  filterRegion(orders, "华东", "华南");
  const filtered = reportOf(orders);
  assert.deepEqual([filtered.report.dataScale.rows, filtered.report.dataScale.filteredOutRows, filtered.report.dataScale.excludedRows], [17, 23, 2]);
  assert.ok(filtered.html.includes(`${scaleCard("筛除行数", 23)}${scaleCard("不计入分析行数", 2)}`));
  assert.match(filtered.html, new RegExp(`<\\/table><\\/div>\\s*<p>${left}<\\/p>\\s*<p>数据质量部分始终基于原始数据`));

  // One kind only: only that kind; the totals section keeps its usual note.
  const repeats = await sampleWithExclusions({ totals: false });
  const one = reportOf(repeats);
  assert.equal(textOf(repeats, "qualityScopeNote"), "质量页检查的是文件里的全部 42 行，包括不计入分析的 1 行；偏大或偏小的值也按这 42 行算。");
  // The same words as the filter panel without a filter.
  // HTML: the same four places.
  assert.ok(one.html.includes("<p>不计入分析：重复的行 1 行（第 21 行）。</p>"));
  assert.match(one.html, new RegExp(`<h3>疑似合计 \\/ 小计行<\\/h3>\\s*<p>${regExpText(repeats.evaluate("SUSPECTED_TOTAL_NOTE[0]"))}<\\/p>`));
  assert.match(one.html, new RegExp(`<h2>数据质量报告<\\/h2>\\s*<p>${regExpText(textOf(repeats, "qualityScopeNote"))}<\\/p>`));
  assert.ok(one.html.includes("<p>未设置筛选，结果基于全部 41 行（不含不计入分析的 1 行）。</p>"));
  assert.equal(textOf(repeats, "filterMessage"), "当前分析使用全部 41 行（不含不计入分析的 1 行）。");
  assert.ok(!one.html.includes("完整数据"));
});

test("v3.1.0 report: nothing left out, or an old snapshot, prints none of it", async () => {
  const orders = await clickSampleButton(sampleOrdersFetchedText());
  const plain = reportOf(orders);
  assert.equal(plain.report.dataScale.excludedRows, 0);
  assert.equal(plain.report.exclusions, null);
  assert.equal(plain.report.quality.excludedScope, null);
  assert.equal(plain.report.quality.suspectedTotalsExcluded, 0);
  for (const text of [plain.html]) {
    assert.ok(!text.includes("不计入分析"), text.slice(0, 80));
    assert.ok(text.includes(NO_FILTER_SENTENCE));
  }
  assert.equal(orders.evaluate("formatReportNoFilterText(42, 0)"), "");
  assert.equal(orders.evaluate("formatReportNoFilterText(12345, 1234)"), "未设置筛选，结果基于全部 12,345 行（不含不计入分析的 1,234 行）。");

  // A snapshot from before v3.1.0 has none of the new keys: 0 / null, and no sentence.
  const excluded = await sampleWithExclusions();
  const report = JSON.parse(excluded.evaluate('JSON.stringify(buildReportData(new Date("2026-10-25T00:00:00.000Z")))'));
  for (const key of REPORT_EXCLUDED_KEYS) {
    const [head, tail] = key.split(".");
    if (tail) delete report[head][tail];
    else delete report[head];
  }
  put(excluded.context, "__oldReport", report);
  const normalized = JSON.parse(excluded.evaluate("JSON.stringify(normalizeReportSnapshot(__oldReport))"));
  assert.deepEqual([normalized.dataScale.excludedRows, normalized.exclusions, normalized.quality.excludedScope, normalized.quality.suspectedTotalsExcluded, normalized.quality.suspectedTotalExtras], [0, null, null, 0, null]);
  const html = excluded.evaluate("buildHtmlReport(__oldReport, [])");
  for (const text of [html]) {
    assert.ok(!text.includes("不计入分析"));
    assert.ok(!text.includes("质量页检查的是"));
    assert.ok(text.includes(excluded.evaluate("SUSPECTED_TOTAL_NOTE[0]")));
    assert.ok(text.includes(NO_FILTER_SENTENCE));
  }
  assert.ok(html.includes(`${scaleCard("筛除行数", 0)}${scaleCard("字段总数", 9)}`));
});

test("v3.1.0 report: line numbers up to 20, then 等; none when a row has no line number", () => {
  const scope = createScriptContext();
  const lines = (count, from = 5) => Array.from({ length: count }, (_, i) => from + i);
  const sentence = (exclusions) => scope.evaluate(`formatReportExclusions(${JSON.stringify(exclusions)})`);
  assert.equal(sentence(null), "");
  assert.equal(sentence({ suspectedTotal: { count: 5, lines: [5, 8, 10, 14, 15] }, duplicateRows: { count: 0, lines: [] } }), "不计入分析：疑似合计行 5 行（第 5、8、10、14、15 行）。");
  assert.equal(sentence({ suspectedTotal: { count: 0, lines: [] }, duplicateRows: { count: 20, lines: lines(20) } }), `不计入分析：重复的行 20 行（第 ${lines(20).join("、")} 行）。`);
  assert.equal(sentence({ suspectedTotal: { count: 0, lines: [] }, duplicateRows: { count: 1234, lines: lines(20) } }), `不计入分析：重复的行 1,234 行（第 ${lines(20).join("、")} 行等）。`);
  assert.equal(sentence({ suspectedTotal: { count: 2, lines: null }, duplicateRows: { count: 0, lines: [] } }), "不计入分析：疑似合计行 2 行。");
  // From the state: at most 20 numbers kept, and none at all when a row has no line number.
  const rows = [["甲", "10"], ["乙", "20"], ["丙", "30"], ["合计", "60"], ["合计", "60"]].map(([名称, 金额]) => ({ 名称, 金额 }));
  commitRows(scope, ["名称", "金额"], rows, "no-lines.csv");
  scope.evaluate('setRowExclusion("suspectedTotal", true)');
  const report = JSON.parse(scope.evaluate("JSON.stringify(buildReportData())"));
  assert.deepEqual(report.exclusions, { suspectedTotal: { count: 2, lines: null }, duplicateRows: { count: 0, lines: [] }, overlap: null });
  assert.ok(scope.evaluate("buildHtmlReport(buildReportData(), [])").includes("<p>不计入分析：疑似合计行 2 行。</p>"));
});

test("v3.1.0 report: a row of both kinds is counted once, and named, in the 不计入分析 line", async () => {
  const both = await importCsvFile("编号,名称,金额\n1,甲,10\n2,乙,20\n2,乙,20\n3,丙,30\n,合计,80\n,合计,80\n", "totals-and-repeats.csv");
  both.evaluate('setRowExclusion("suspectedTotal", true); setRowExclusion("duplicateRows", true)');
  const { report, html } = reportOf(both);
  assert.equal(report.dataScale.excludedRows, 3);
  assert.deepEqual(report.exclusions, { suspectedTotal: { count: 2, lines: [6, 7] }, duplicateRows: { count: 2, lines: [4, 7] }, overlap: { count: 1, lines: [7] } });
  const line = "不计入分析：疑似合计行 2 行（第 6、7 行）、重复的行 2 行（第 4、7 行），共 3 行（第 7 行两类都是）。";
  assert.ok(html.includes(`<p>${line}</p>`));
  assert.ok(html.includes(`${scaleCard("不计入分析行数", 3)}`));

  // Without overlap the sentence is unchanged, and the snapshot says null.
  const sample = reportOf(await sampleWithExclusions());
  assert.equal(sample.report.exclusions.overlap, null);
  assert.ok(sample.html.includes("<p>不计入分析：疑似合计行 1 行（第 44 行）、重复的行 1 行（第 21 行）。</p>"));
  // An old snapshot (no overlap key) gets null and the sentence without 共.
  const old = JSON.parse(both.evaluate('JSON.stringify(buildReportData(new Date("2026-10-25T00:00:00.000Z")))'));
  delete old.exclusions.overlap;
  put(both.context, "__oldOverlap", old);
  assert.equal(both.evaluate("normalizeReportSnapshot(__oldOverlap).exclusions.overlap"), null);
  assert.ok(both.evaluate("buildHtmlReport(__oldOverlap, [])").includes("<p>不计入分析：疑似合计行 2 行（第 6、7 行）、重复的行 2 行（第 4、7 行）。</p>"));

  // More than 20 rows of both kinds end in 等; without line numbers only 共 N 行.
  const sentence = (exclusions) => both.evaluate(`formatReportExclusions(${JSON.stringify(exclusions)})`);
  const lines = Array.from({ length: 20 }, (_, i) => i + 2);
  assert.equal(sentence({ suspectedTotal: { count: 30, lines }, duplicateRows: { count: 40, lines }, overlap: { count: 25, lines } }),
    `不计入分析：疑似合计行 30 行（第 ${lines.join("、")} 行等）、重复的行 40 行（第 ${lines.join("、")} 行等），共 45 行（第 ${lines.join("、")} 行等两类都是）。`);
  assert.equal(sentence({ suspectedTotal: { count: 2, lines: null }, duplicateRows: { count: 2, lines: null }, overlap: { count: 1, lines: null } }),
    "不计入分析：疑似合计行 2 行、重复的行 2 行，共 3 行。");
  assert.equal(sentence({ suspectedTotal: { count: 1234, lines: null }, duplicateRows: { count: 1000, lines: null }, overlap: { count: 1, lines: null } }),
    "不计入分析：疑似合计行 1,234 行、重复的行 1,000 行，共 2,233 行。");
});

test("v3.1.0 report: the totals section has the page's 只等于上面一行 and 另有 N 行 sentences", async () => {
  const region = await importStage3Fixture("分区小计表.csv");
  const single = "其中 1 行只等于上面一行，可能不是合计行，先点“查看”确认。";
  assert.ok(htmlText(region, "suspectedTotalDetails").includes(single));
  const regionReport = reportOf(region);
  assert.ok(regionReport.html.includes(`<p>${single}</p>`));
  // Left out: the excluded note and still the same sentence.
  region.evaluate('setRowExclusion("suspectedTotal", true)');
  const excludedReport = reportOf(region);
  assert.match(excludedReport.html, new RegExp(`<h3>疑似合计 \\/ 小计行<\\/h3>\\s*<p>${regExpText(textOf(region, "suspectedTotalExcludedNote"))}<\\/p>`));
  assert.ok(excludedReport.html.includes("<p>不计入分析：疑似合计行 5 行（第 5、8、10、14、15 行）。</p>"));
  assert.ok(excludedReport.html.includes(`<p>${single}</p>`));

  const tasks = await importCsvFile(taskTableCsv(), "task-table.csv");
  const tasksReport = reportOf(tasks);
  for (const sentence of ["另有 89 行也带这类字样，但数字对不上，没有列出。", "其中 6 行只等于上面一行，可能不是合计行，先点“查看”确认。"]) {
    assert.ok(htmlText(tasks, "suspectedTotalDetails").includes(sentence));
    assert.ok(tasksReport.html.includes(`<p>${sentence}</p>`), sentence);
  }
  // The sample has neither.
  const { html } = reportOf(await clickSampleButton(sampleOrdersFetchedText()));
  for (const text of [html]) assert.doesNotMatch(text, /只等于上面一行|也带这类字样/);
});

test("v3.1.0 export: one data export, the rows of the current analysis, and what it leaves out", async () => {
  const html = indexHtmlSource();
  const options = html.slice(html.indexOf('<div class="data-export-options">'), html.indexOf('<div id="standardizedExportPanel"'));
  assert.match(options, /^<div class="data-export-options">\s*<div class="data-export-option">\s*<div>\s*<h3>当前分析的行<\/h3>\s*<p class="export-count"><span id="dataExportFilteredBefore"><\/span><strong id="dataExportFilteredCount">0<\/strong><span id="dataExportFilteredAfter"><\/span><\/p>\s*<p>导出统计、图表用的那批行（筛选后的行）。<span id="dataExportExcludedNote"><\/span><\/p>\s*<\/div>\s*<button id="exportFilteredCsv" type="button" disabled>导出当前分析的行<\/button>\s*<\/div>\s*<div class="data-export-option">\s*<div>\s*<h3>写法统一的表<\/h3>/);

  const orders = await clickSampleButton(sampleOrdersFetchedText());
  const downloads = captureDownloads(orders);
  const note = () => textOf(orders, "dataExportExcludedNote");
  assert.deepEqual([textOf(orders, "dataExportFilteredCount"), note()], ["42", ""]);
  orders.evaluate('setRowExclusion("suspectedTotal", true); setRowExclusion("duplicateRows", true)');
  assert.deepEqual([textOf(orders, "dataExportFilteredCount"), note()], ["40", "不含不计入分析的 2 行。"]);
  assert.equal(orders.evaluate("formatDataExportExcludedNote(1234)"), "不含不计入分析的 1,234 行。");
  orders.evaluate('exportProcessedCsv("filtered")');
  const [file] = await downloads();
  assert.match(file.fileName, /^sample-orders_当前分析_\d{8}\.csv$/);
  assert.equal(textOf(orders, "dataExportStatus"), `已导出当前分析的 40 行：${file.fileName}`);
  assert.deepEqual(file.bom, [0xef, 0xbb, 0xbf]);
  assert.equal(file.text.split("\r\n").length, 41);
  // Filtered as well: still the rows of the statistics, and the note stays.
  filterRegion(orders, "华东", "华南");
  assert.deepEqual([textOf(orders, "dataExportFilteredCount"), note()], ["17", "不含不计入分析的 2 行。"]);
  // Nothing left out: no note.
  orders.evaluate('clearAllFilters(); setRowExclusion("suspectedTotal", false); setRowExclusion("duplicateRows", false)');
  assert.deepEqual([textOf(orders, "dataExportFilteredCount"), note()], ["42", ""]);
  // The deduplicated export is gone, code and page.
  assert.throws(() => orders.evaluate('getProcessedExportRows("deduplicated")'), /不支持的数据导出类型/);
  for (const source of ["app.js", "report.js", "script.js"]) assert.doesNotMatch(repoFile(source), /deduplicated|Deduplicated/, source);
});

test("v3.1.0 export: the page never says 去重后 or 导出筛选后 CSV again", async () => {
  const words = ["去重后", "导出筛选后 CSV", "筛选后数据", "_filtered"];
  const html = indexHtmlSource().replace(/<!--[\s\S]*?-->/g, "");
  for (const word of words) assert.ok(!html.includes(word), word);
  const orders = await clickSampleButton(sampleOrdersFetchedText());
  orders.evaluate('setRowExclusion("suspectedTotal", true); exportProcessedCsv("filtered")');
  for (const [id, element] of orders.elements) {
    for (const word of words) {
      assert.ok(!String(element.textContent).includes(word) && !String(element.innerHTML).includes(word), `${id}: ${word}`);
    }
  }
});

test("v3.1.0 wording: the new sentences use no old interface word and no space between two Chinese characters", async () => {
  const hits = [];
  const orders = await clickSampleButton(sampleOrdersFetchedText());
  const region = await importStage3Fixture("分区小计表.csv");
  const tasks = await importCsvFile(taskTableCsv(), "task-table.csv");
  for (const [name, scope] of [["sample", orders], ["region", region], ["tasks", tasks]]) {
    scope.evaluate('setRowExclusion("suspectedTotal", true); if (state.duplicateRows) setRowExclusion("duplicateRows", true)');
    for (const id of ["qualityScopeNote", "suspectedTotalExcludedNote", "suspectedTotalMessage", "duplicateExclusionText", "duplicateExclusionMessage", "globalFilterSummary", "filterMessage", "previewPageInfo", "statusPanel", ...CONTAINS_NOTE_IDS]) {
      findOldInterfaceWords(`${name} #${id}`, textOf(scope, id), hits);
      if (HAN_SPACE_HAN.test(textOf(scope, id).replace(/「[^」]*」/g, ""))) hits.push(`${name} #${id}: a space between two Chinese characters`);
    }
    findOldInterfaceWords(`${name} #suspectedTotalDetails`, htmlText(scope, "suspectedTotalDetails"), hits);
  }
  const fixed = JSON.parse(orders.evaluate(`JSON.stringify([NO_ID_DUPLICATE_NOTE[0], formatUnmatchedTotalLabels(95, 6, true), formatSingleRowTotals(1),
    formatSuspectedTotalExcludedNote(1), formatSuspectedTotalExcludedNote(5), formatDuplicateIncludedNote(1), formatDuplicateExcludedNote(1), EXCLUSION_DRAFT_REFUSAL[0],
    formatQualityScopeNote(42, 2, 6), formatQualityScopeNote(42, 2, 0), formatAnalysisScope(1, 28, 40, 2), formatNoFilterMessage(40, 2),
    formatAnalysisContainsNote(1, 1), formatAnalysisContainsNote(1, 0), formatAnalysisContainsNote(0, 1), formatGoToQualityAriaLabel(1, 1), GO_TO_QUALITY_LABEL[0],
    formatReportNoFilterText(40, 2), formatDataExportExcludedNote(2), formatReportExclusions({ suspectedTotal: { count: 1, lines: [44] }, duplicateRows: { count: 30, lines: Array.from({ length: 20 }, (_, i) => i + 2) } })])`));
  for (const text of fixed) {
    findOldInterfaceWords("format.js", text, hits);
    if (HAN_SPACE_HAN.test(text)) hits.push(`format.js: a space between two Chinese characters: ${text}`);
  }
  assert.deepEqual(hits, []);
  // 不计入分析 is part of several sentences and the button: the button is matched with its markup.
  assert.match(orders.elements.get("suspectedTotalDetails").innerHTML, /data-exclusion-kind="suspectedTotal" data-exclusion-on="false"[^>]*>撤销<\/button>/);
  assert.doesNotMatch(orders.elements.get("suspectedTotalDetails").innerHTML, />不计入分析<\/button>/);
});

test("v3.2.0 tests read index.html only through indexHtmlSource(), which takes out data-i18n attributes and nothing else", () => {
  const readCall = new RegExp("(?:" + ["readFileSync", "repoFile"].join("|") + ")\\([^;\\n]*index\\.html");
  for (const file of ["basic-data-processing.test.cjs", "core.test.cjs", "data-processing.test.cjs", "destructive-qa.test.cjs"]) {
    const lines = fs.readFileSync(path.join(__dirname, file), "utf8").split(/\r?\n/);
    const reads = lines.filter((line) => readCall.test(line));
    assert.ok(reads.every((line) => line.includes("stripI18nAttributes(")), `${file}: ${reads.filter((line) => !line.includes("stripI18nAttributes(")).join(" | ")}`);
    if (file !== "destructive-qa.test.cjs") {
      assert.equal(reads.length, 1, file);
      assert.ok(lines.includes("const DATA_I18N_ATTRIBUTE = /\\s+data-i18n[\\w-]*=\"[^\"]*\"/g;"), file);
    }
  }
  assert.equal(stripI18nAttributes('<p id="a" data-i18n="upload.title" class="b">表格</p>'), '<p id="a" class="b">表格</p>');
  assert.equal(stripI18nAttributes('<button data-i18n-attr="aria-label=nav.open title=nav.tip" type="button">打开</button>'), '<button type="button">打开</button>');
  assert.equal(stripI18nAttributes('<details data-i18n-open="x" data-i18n-closed="y">'), "<details>");
  // Text that only looks like the attribute stays: no leading space, or not an attribute name.
  assert.equal(stripI18nAttributes("<p>data-i18n=\"x\"</p>"), "<p>data-i18n=\"x\"</p>");
  assert.equal(stripI18nAttributes('<p data-lang="en" data-i18ndex="1">'), '<p data-lang="en">');
  assert.equal(stripI18nAttributes('<p data-lang="en">'), '<p data-lang="en">');
  // The raw file is read by a name the scan above does not see, on purpose: these are the only
  // comparisons with it. (1) Everything the helper takes out is an attribute inside a start tag,
  // and as many as the file has data-i18n names.
  const raw = fs.readFileSync(path.join(REPO_ROOT, ["index", "html"].join(".")), "utf8").replace(/\r\n/g, "\n");
  const taken = [...raw.matchAll(DATA_I18N_ATTRIBUTE)];
  // (Names written inside a data-i18n-attr value, "data-i18n-open=details.open", have no quote.)
  assert.equal(taken.length, (raw.match(/\sdata-i18n[\w-]*="/g) || []).length);
  assert.ok(taken.length > 0);
  for (const match of taken) {
    const before = raw.slice(0, match.index);
    assert.ok(before.lastIndexOf("<") > before.lastIndexOf(">"), `not inside a tag: ${match[0]}`);
    assert.match(before.slice(before.lastIndexOf("<")), /^<[a-z][\w-]*(?:\s+[^\s=>]+(?:="[^"]*")?)*$/, match[0]);
  }
  assert.ok(!indexHtmlSource().includes("data-i18n"));
  // (2) What is left is index.html as it was at commit 2 (1001bb9, before any data-i18n), byte for
  // byte, once the v3.2.0 additions that are not attributes are taken out too: the inline script in
  // <head> and the two language switches. A later commit that changes the Chinese page on purpose
  // changes this hash and says why.
  const withoutAdditions = stripI18nAttributes(raw)
    .replace(/ {2}<!-- v3\.2\.0: the one script[\s\S]*?<\/script>\n/, "")
    .replace(/\n\s*<button id="(?:languageSwitch|sidebarLanguageSwitch)"[^>]*>English<\/button>/g, "");
  // The ten cache parameters (?v=x.y.z of style.css and the nine scripts) are written as ?v=VERSION
  // before hashing, as the report hash blanks the version: a release changes them and nothing else,
  // and the version itself is pinned place by place by the version test (data-processing). So a
  // release stays 17 places, and this hash still sees every other character of the page.
  const cacheParameters = withoutAdditions.match(/\?v=\d+\.\d+\.\d+"/g) || [];
  assert.equal(cacheParameters.length, 10);
  const versionless = withoutAdditions.replace(/\?v=\d+\.\d+\.\d+"/g, "?v=VERSION\"");
  // sha256 of index.html at 1001bb9 with LF line endings was 75fc338d…46851. Commit 10 changed the
  // Chinese page on purpose: the note above statistics / charts / custom analysis moved out of the
  // heading into a bar of its own (3 places), and the export's two counts became “共 N 行” (2 places);
  // that value (7c3f1ed6…38435) was the old content with only those 5 replacements. The release
  // commit took the version out of the hash: the value below is 5e2ac9a's index.html through these
  // steps, and the 3.2.0 page gives the same value (nothing but the version changed). v3.2.2 only
  // changed the repository URL in the footer to sheet-checkup; the value below is the old content
  // (95ef087e…ca866) with just that one replacement, compared byte for byte.
  assert.equal(crypto.createHash("sha256").update(versionless).digest("hex"), "8c1f88e7169f3b6aeb10a60afd154f3d8d16f68846e2ba6b855e778f20bdc80a");
});

// v3.2.0 (R1): long original names. The file name before the date was cut at 80 characters as a
// whole, so a long original name cut off the labels: the standardized table and the change list
// got the same name, _前N条 went missing, and problem-row files of different categories or columns
// got the same name. Now the original name (and the user's column and writing) is cut, the labels
// never; a name that fitted in 80 is byte for byte the same as before.
const LONG_NAME_EN71 = "Quarterly Regional Sales Orders Export for the North America Team v2 Q3";
const LONG_NAME_ZH = (count) => "华东区上半年销售订单明细表".repeat(8).slice(0, count);
const FILE_NAME_DATE = new Date(2026, 10, 21);
// The file names as cd4d2a0 builds them: one sanitize over the whole, cut at 80.
const oldSanitizeFileName = (value) => String(value || "data-report")
  .replace(/[<>:"/\\|?*\u0000-\u001F]/g, "-")
  .replace(/[\u202A-\u202E\u2066-\u2069]/g, "")
  .replace(/\s+/g, " ")
  .trim();
const oldCutFileName = (value) => oldSanitizeFileName(value).slice(0, 80) || "data-report";
const PROBLEM_KIND_WORDS = { missing: "空值", conversion: "认不出", outlier: "偏大或偏小", duplicateRows: "整行重复", idDuplicate: "编号重复", placeholder: "像是空值的写法",
  parseWarnings: "读取有疑问", suspectedTotal: "疑似合计", formulaError: "公式错误", dateStyle: "日期写法", lookAlike: "写法不一样" };

async function longNameExports(stem) {
  const scope = await importCsvFile(readExampleCsv("sample-orders.csv"), `${stem}.csv`);
  put(scope.context, "__fileNameDate", FILE_NAME_DATE);
  const categories = JSON.parse(scope.evaluate("JSON.stringify(state.problemCategories)"));
  const problemRows = categories.map((category) => {
    put(scope.context, "__fileNameCategory", category);
    const joined = [stem, "问题行", PROBLEM_KIND_WORDS[category.kind], category.field, category.key].filter((part) => part !== undefined && part !== null && part !== "").join("_");
    return { category, name: scope.evaluate("buildProblemRowsFileName(__fileNameCategory, __fileNameDate)"), joined };
  });
  return {
    scope,
    problemRows,
    data: { name: scope.evaluate('buildStandardizedFileName("data", __fileNameDate)'), joined: `${stem}_写法统一` },
    changes: { name: scope.evaluate('buildStandardizedFileName("changes", __fileNameDate)'), joined: `${stem}_写法统一改动清单` },
    changesCut: { name: scope.evaluate('buildStandardizedFileName("changes", __fileNameDate, 1048575)'), joined: `${stem}_写法统一改动清单_前1048575条` },
    analysis: scope.evaluate('buildDataExportFileName("filtered", __fileNameDate)'),
    report: scope.evaluate("buildReportFileName(\"html\", __fileNameDate)")
  };
}
const beforeDate = (name) => name.replace(/_20261121\.csv$/, "");

test("v3.2.0 long file names: the labels stay, so no two exports of one file get the same name", async () => {
  for (const stem of [LONG_NAME_EN71, LONG_NAME_ZH(76)]) {
    const files = await longNameExports(stem);
    assert.ok(files.data.name.endsWith("_写法统一_20261121.csv"), files.data.name);
    assert.ok(files.changes.name.endsWith("_写法统一改动清单_20261121.csv"), files.changes.name);
    assert.ok(files.changesCut.name.endsWith("_写法统一改动清单_前1048575条_20261121.csv"), files.changesCut.name);
    const names = [files.data.name, files.changes.name, files.changesCut.name, ...files.problemRows.map((item) => item.name)];
    assert.equal(new Set(names).size, names.length, names.join("\n"));
    for (const { category, name } of files.problemRows) {
      assert.ok(name.includes(`_问题行_${PROBLEM_KIND_WORDS[category.kind]}`), name);
      // The column (and the writing) is still there, maybe cut, after the category word.
      if (category.field) assert.ok(beforeDate(name).split(`_问题行_${PROBLEM_KIND_WORDS[category.kind]}_`)[1].length > 0, name);
    }
    for (const name of names) {
      assert.ok(beforeDate(name).length <= 80, `${beforeDate(name).length} ${name}`);
      assert.ok(name.startsWith(stem.slice(0, 10)), name);
    }
  }
  // The rule itself: the parts that may be cut share what the labels leave, longest first (the
  // first one on a tie); a label is never cut, even past 80; a cut part loses a space it ends with.
  const { evaluate } = createScriptContext();
  const labeled = (parts) => evaluate(`buildLabeledFileName(${JSON.stringify(parts)})`);
  assert.equal(labeled([{ text: "S".repeat(90), cut: true }, { text: "L".repeat(40), cut: false }, { text: "F".repeat(40), cut: true }]),
    `${"S".repeat(19)}_${"L".repeat(40)}_${"F".repeat(19)}`);
  assert.equal(labeled([{ text: "S".repeat(10), cut: true }, { text: "L".repeat(90), cut: false }]), `S_${"L".repeat(90)}`);
  assert.equal(labeled([{ text: `${"a".repeat(74)} b`, cut: true }, { text: "写法统一", cut: false }]), `${"a".repeat(74)}_写法统一`);
  assert.equal(labeled([{ text: "表".repeat(81), cut: true }]), "表".repeat(80));
  // One character over: before, the column lost its last character; now the original name does.
  const over = await longNameExports("B".repeat(69));
  const missing = over.problemRows.find((item) => item.category.kind === "missing" && item.category.field === "下单日期");
  assert.equal(missing.joined.length, 81);
  assert.equal(missing.name, `${"B".repeat(68)}_问题行_空值_下单日期_20261121.csv`);
});

test("v3.2.0 long file names: every name that fitted in 80 characters is byte for byte as before", async () => {
  let same = 0;
  let cut = 0;
  for (const stem of [LONG_NAME_EN71, LONG_NAME_ZH(40), LONG_NAME_ZH(76), "B".repeat(68), "B".repeat(69), "短名", `  ${"甲".repeat(70)} 乙  `, `a:b?${"x".repeat(60)} \u202Ec`]) {
    const files = await longNameExports(stem);
    for (const { name, joined } of [files.data, files.changes, files.changesCut, ...files.problemRows]) {
      if (oldSanitizeFileName(joined).length > 80) { cut += 1; continue; }
      assert.equal(name, `${oldCutFileName(joined)}_20261121.csv`, joined);
      same += 1;
    }
    // These two never cut a label: unchanged for every length.
    assert.equal(files.analysis, `${oldCutFileName(stem)}_当前分析_20261121.csv`);
    assert.ok(files.report.startsWith(`${oldCutFileName(stem)}-20261121-`), files.report);
  }
  // The boundary: 68 + "_问题行_空值_下单日期" is exactly 80 and stays as it was.
  const exact = (await longNameExports("B".repeat(68))).problemRows.find((item) => item.category.kind === "missing" && item.category.field === "下单日期");
  assert.equal(exact.joined.length, 80);
  assert.equal(exact.name, `${exact.joined}_20261121.csv`);
  assert.ok(same > 50 && cut > 10, `${same} same, ${cut} cut`);
});

// An original name that ends in a space, the whole name exactly 80 (or 79) characters: it fits, so
// it is cd4d2a0's name byte for byte, space before "_" included. With "< 80" instead of "<= 80" the
// 80-character names would go the cutting way and lose that space.
test("v3.2.0 long file names: 80 and 79 characters with a space before a label are as before", () => {
  const scope = createScriptContext();
  put(scope.context, "__fileNameDate", FILE_NAME_DATE);
  const nameOf = (stem, expression) => {
    scope.evaluate(`state.sourceFileName = ${JSON.stringify(`${stem}.csv`)}`);
    return scope.evaluate(expression);
  };
  let checked = 0;
  for (const total of [80, 79]) {
    // 问题行 · 空值 · <column>: "订单 " + "_问题行_空值_" + the column.
    const stem = "订单 ";
    const field = "列".repeat(total - stem.length - "_问题行_空值_".length);
    put(scope.context, "__spaceCategory", { kind: "missing", field });
    const problemJoined = `${stem}_问题行_空值_${field}`;
    assert.equal(problemJoined.length, total);
    assert.equal(nameOf(stem, "buildProblemRowsFileName(__spaceCategory, __fileNameDate)"), `${oldCutFileName(problemJoined)}_20261121.csv`);
    assert.ok(nameOf(stem, "buildProblemRowsFileName(__spaceCategory, __fileNameDate)").startsWith("订单 _问题行_"));
    // 写法统一的表, 改动清单, 改动清单（前 N 条）: an original name of the right length ending in a space.
    for (const [label, expression] of [["写法统一", 'buildStandardizedFileName("data", __fileNameDate)'], ["写法统一改动清单", 'buildStandardizedFileName("changes", __fileNameDate)'], ["写法统一改动清单_前1048575条", 'buildStandardizedFileName("changes", __fileNameDate, 1048575)']]) {
      const spaced = `订单${"a".repeat(total - 4 - label.length)} `;
      const joined = `${spaced}_${label}`;
      assert.equal(joined.length, total, label);
      const name = nameOf(spaced, expression);
      assert.equal(name, `${oldCutFileName(joined)}_20261121.csv`, label);
      assert.ok(name.includes(` _${label}_20261121`), name);
      checked += 1;
    }
    checked += 1;
  }
  assert.equal(checked, 8);
});

// ---- v3.2.0 commit 4: the language skeleton and the static guard ----
const i18nGuard = require("./i18n-guard.cjs");
const i18nWhitelist = () => new Map(i18nGuard.I18N_WHITELIST.map(([file, decl, count]) => [`${file}|${decl}`, count]));

test("v3.2.0 static guard: the tokenizer puts every Han character of the nine scripts in a comment or a literal", () => {
  for (const file of i18nGuard.APP_SCRIPT_NAMES) {
    const source = repoFile(file);
    const { all, inTokens } = i18nGuard.lexerSelfCheck(source);
    assert.ok(all > 0, file);
    assert.equal(inTokens, all, file);
  }
  // The guard's own file list is the page's script list.
  assert.deepEqual(i18nGuard.APP_SCRIPT_NAMES, require("./test-context.cjs").APP_SCRIPT_FILES);
});

test("v3.2.0 static guard: the whitelist names declarations that exist, with exactly that many Chinese literals", () => {
  const whitelist = i18nWhitelist();
  assert.equal(whitelist.size, i18nGuard.I18N_WHITELIST.length, "one entry per file and declaration");
  for (const file of i18nGuard.APP_SCRIPT_NAMES) {
    const tokens = i18nGuard.lex(repoFile(file));
    const { used } = i18nGuard.checkScriptTokens(file, tokens, whitelist, { rule1: false });
    for (const [entryFile, decl, count] of i18nGuard.I18N_WHITELIST.filter(([entryFile]) => entryFile === file)) {
      assert.ok(tokens.some((token) => token.type === "ident" && token.value === decl && token.decl === decl), `${entryFile} ${decl} is declared`);
      assert.equal(used.get(`${entryFile}|${decl}`), count, `${entryFile} ${decl}`);
    }
  }
});

test("v3.2.0 static guard: rules 1 and 2 in every script (no Chinese outside L(), none in an English argument)", () => {
  const whitelist = i18nWhitelist();
  // Commit 9: the list of finished files (commits 5–8 added them one by one) is gone.
  assert.equal("I18N_DONE_FILES" in i18nGuard, false);
  assert.equal(i18nGuard.APP_SCRIPT_NAMES.length, 9);
  for (const file of i18nGuard.APP_SCRIPT_NAMES) {
    const { problems } = i18nGuard.checkScriptTokens(file, i18nGuard.lex(repoFile(file)), whitelist);
    assert.deepEqual(problems.map((problem) => `${problem.line} ${problem.decl} ${problem.kind}: ${problem.text}`), [], file);
  }
});

test("v3.2.0 static guard catches each kind of missed text (the review's 13 cases)", () => {
  assert.equal(i18nGuard.SELF_TEST_CASES.length, 13);
  for (const [name, code, expected] of i18nGuard.SELF_TEST_CASES) {
    const { problems } = i18nGuard.checkScriptTokens("test.js", i18nGuard.lex(code), i18nGuard.SELF_TEST_WHITELIST);
    assert.equal(problems.length, expected, `${name}: ${problems.map((problem) => problem.kind).join("; ")}`);
  }
  // The self-check fails when the tokenizer loses its place (a division read as a regex start
  // would swallow code): Han in an identifier is outside every literal.
  const broken = i18nGuard.lexerSelfCheck("const 金额 = 1;");
  assert.notEqual(broken.inTokens, broken.all);
});

test("v3.2.0 static guard: index.html has a key for every Chinese text and attribute, and every key has English", () => {
  const raw = fs.readFileSync(path.join(REPO_ROOT, ["index", "html"].join(".")), "utf8");
  const english = JSON.parse(createScriptContext().evaluate("JSON.stringify(I18N_EN)"));
  // Commit 9: the whole page (commit 4 checked only the top bar, the sidebar and the upload page).
  const { problems, keys } = i18nGuard.checkIndexHtml(raw, english);
  assert.deepEqual(problems.map((problem) => `${problem.line} ${problem.kind}: ${problem.text}`), []);
  assert.equal(keys.size, Object.keys(english).length);
  // Commit 13: the product name's three places use the keys the name test checks in English.
  assert.match(raw, /<title data-i18n="page\.title">表格体检 \| 检查 CSV、Excel 表格的数据问题<\/title>/);
  assert.match(raw, /<a class="brand" href="#mainContent" aria-label="表格体检首页" data-i18n-attr="aria-label=brand\.label">/);
  assert.match(raw, /<strong data-i18n="brand\.name">表格体检<\/strong>/);
  const summary = raw.match(/<summary [^>]*>/)[0];
  assert.match(summary, /data-i18n-open="收起" data-i18n-closed="展开" data-i18n-attr="data-i18n-open=details\.open data-i18n-closed=details\.closed"/);
  // The same check finds a Chinese text without a key, an attribute not named and a key without English.
  const bad = i18nGuard.checkIndexHtml('<p data-i18n="hero.title">甲</p><p>乙</p><b aria-label="丙">x</b><i data-i18n="no.such">丁</i>', english);
  assert.deepEqual(bad.problems.filter((problem) => problem.line > 0).map((problem) => problem.kind), ["text without data-i18n", "attribute aria-label not in data-i18n-attr"]);
  assert.ok(bad.problems.some((problem) => problem.kind === "key not in I18N_EN" && problem.text === "no.such"));
  assert.ok(keys.size > 180, `${keys.size} keys`);
});

const LANGUAGE_CASES = [
  ["", ["zh-CN"], "zh"], ["", ["zh"], "zh"], ["", ["zh-TW", "en"], "zh"], ["", ["zh-HK"], "zh"], ["", ["ZH-cn"], "zh"],
  ["", ["en-US"], "en"], ["", ["en-US", "zh-CN"], "en"], ["", ["ja-JP"], "en"], ["", ["zhx"], "en"], ["", ["fr"], "en"],
  ["?lang=en", ["zh-CN"], "en"], ["?lang=zh", ["en-US"], "zh"], ["?a=1&lang=en", ["zh-CN"], "en"], ["?lang=en&a=1", ["zh-CN"], "en"],
  ["?lang=EN", ["zh-CN"], "zh"], ["?lang=EN", ["en-US"], "en"], ["?lang=fr", ["zh-CN"], "zh"], ["?lang=en-US", ["zh-CN"], "zh"], ["?xlang=en", ["zh-CN"], "zh"],
  ["", [], "zh"], ["", [""], "zh"], ["?lang=en", [], "en"], [undefined, undefined, "zh"]
];

test("v3.2.0 language: ?lang= first (lower case only), then the browser's first language, Chinese without either", () => {
  const scope = createScriptContext();
  for (const [search, languages, expected] of LANGUAGE_CASES) {
    put(scope.context, "__languageCase", [search, languages]);
    assert.equal(scope.evaluate("pickInterfaceLanguage(__languageCase[0], __languageCase[1])"), expected, JSON.stringify([search, languages]));
  }
  const read = (browser) => { put(scope.context, "__browser", browser); return JSON.parse(scope.evaluate("JSON.stringify(readNavigatorLanguages(__browser))")); };
  assert.deepEqual(read(null), []);
  assert.deepEqual(read({ languages: ["en-US", "zh"], language: "zh" }), ["en-US", "zh"]);
  assert.deepEqual(read({ languages: [], language: "en-GB" }), ["en-GB"]);
  assert.deepEqual(read({ language: "zh-TW" }), ["zh-TW"]);
  assert.deepEqual(read({}), []);
});

test("v3.2.0 language: the inline script in <head> decides as pickInterfaceLanguage() does, defines nothing and only marks <html>", () => {
  const html = indexHtmlSource();
  const head = html.slice(0, html.indexOf("</head>"));
  const inline = [...head.matchAll(/<script>([\s\S]*?)<\/script>/g)];
  assert.equal(inline.length, 1);
  // Before the style sheet, so the hidden state is set before the first paint; after <title>.
  assert.ok(head.indexOf("<script>") < head.indexOf('<link rel="stylesheet"'));
  assert.ok(head.indexOf("</title>") < head.indexOf("<script>"));
  const vm = require("node:vm");
  const code = inline[0][1];
  assert.match(code.trim(), /^\(function \(\) \{[\s\S]*\}\)\(\);$/);
  const scope = createScriptContext();
  const browsers = [];
  for (const [search, languages] of LANGUAGE_CASES) browsers.push([search, languages === undefined ? undefined : { languages, language: languages[0] }]);
  browsers.push(["", { languages: [], language: "en-US" }], ["", { languages: [], language: "zh-CN" }], ["", { language: "" }], ["?lang=zh", {}]);
  for (const [search, browser] of browsers) {
    const marks = [];
    const sandbox = vm.createContext({
      window: { location: { search } },
      navigator: browser || {},
      document: { documentElement: { setAttribute: (name, value) => marks.push([name, value]) } },
      String,
      Array
    });
    const before = new Set(Object.keys(sandbox));
    vm.runInContext(code, sandbox);
    assert.deepEqual(Object.keys(sandbox).filter((key) => !before.has(key)), [], "no global names");
    put(scope.context, "__headCase", [search, browser || null]);
    const expected = scope.evaluate("pickInterfaceLanguage(__headCase[0], readNavigatorLanguages(__headCase[1]))");
    assert.deepEqual(marks, expected === "en" ? [["data-lang-pending", ""]] : [], JSON.stringify([search, browser]));
  }
});

test("v3.2.0 language: style.css hides a marked page and shows it after 3 s without any script; script.js removes the mark", () => {
  const css = repoFile("style.css");
  assert.match(cssRule(css, "html[data-lang-pending] body"), /visibility: hidden;\s*animation: lang-pending-reveal 0s linear 3s forwards;/);
  assert.match(css, /@keyframes lang-pending-reveal \{\s*to \{\s*visibility: visible;\s*\}\s*\}/);
  // The preview's 收起 / 展开 come from the attributes now, so they can switch language.
  assert.match(cssRule(css, ".preview-details > summary::after"), /content: attr\(data-i18n-open\);/);
  assert.match(cssRule(css, ".preview-details:not([open]) > summary::after"), /content: attr\(data-i18n-closed\);/);
  assert.doesNotMatch(css, /content: "收起"|content: "展开"/);
  const script = repoFile("script.js");
  // Removed on the way out of a failed startup check, before the notice, and after the language is set.
  const check = script.slice(script.indexOf("function ensureDomElementsPresent()"), script.indexOf("throw new Error", script.indexOf("function ensureDomElementsPresent()")));
  assertInOrder(check, ["if (!missingKeys.length) return;", 'document.documentElement?.removeAttribute("data-lang-pending");', 'const notice = document.createElement("div");']);
  assert.ok(script.indexOf("\nensureDomElementsPresent();") >= 0 && script.indexOf("\nensureDomElementsPresent();") < script.indexOf("const INITIAL_LANGUAGE"));
  const startup = script.slice(script.indexOf("const INITIAL_LANGUAGE"));
  assertInOrder(startup, ["const INITIAL_LANGUAGE = pickInterfaceLanguage(", "if (INITIAL_LANGUAGE !== \"zh\") setInterfaceLanguage(INITIAL_LANGUAGE);", 'document.documentElement?.removeAttribute("data-lang-pending");', "setupResultNavigation();"]);
  // The startup check's representative of format.js is L (still one per file).
  assert.match(script, /\["format\.js", typeof L\],/);
  assert.doesNotMatch(script, /typeof formatNumber/);
});

test("v3.2.0 language: a test context with an English browser goes through the real startup, a Chinese one changes nothing", () => {
  const english = createScriptContext({ language: "en" });
  assert.equal(english.evaluate("currentLanguage"), "en");
  assert.equal(english.evaluate("INITIAL_LANGUAGE"), "en");
  assert.equal(english.elements.get("languageSwitch").textContent, "中文");
  assert.equal(english.elements.get("sidebarLanguageSwitch").textContent, "中文");
  assert.equal(english.evaluate('L("甲", "A")'), "A");
  for (const [options, language] of [[{}, "zh"], [{ language: "zh" }, "zh"], [{ language: "en", search: "?lang=zh" }, "zh"], [{ language: "zh", search: "?lang=en" }, "en"]]) {
    const scope = createScriptContext(options);
    assert.equal(scope.evaluate("currentLanguage"), language, JSON.stringify(options));
  }
  // Chinese: the switches keep the text index.html gives them; nothing writes them.
  const chinese = createScriptContext({ language: "zh" });
  assert.equal(chinese.elements.get("languageSwitch").textContent, "");
});

test("v3.2.0 language: L, Lp / Lpick and enCount follow the language", () => {
  const scope = createScriptContext();
  assert.equal(scope.evaluate('L("甲", "A")'), "甲");
  assert.deepEqual(JSON.parse(scope.evaluate('JSON.stringify(Lp("甲", "A"))')), ["甲", "A"]);
  assert.equal(scope.evaluate('Object.isFrozen(Lp("甲", "A"))'), true);
  assert.equal(scope.evaluate('Lpick(Lp("甲", "A"))'), "甲");
  assert.equal(scope.evaluate("languageSwitchLabel()"), "English");
  scope.evaluate('setInterfaceLanguage("en")');
  assert.equal(scope.evaluate('L("甲", "A")'), "A");
  assert.equal(scope.evaluate('Lpick(Lp("甲", "A"))'), "A");
  assert.equal(scope.evaluate("languageSwitchLabel()"), "中文");
  assert.deepEqual(["enCount(0, \"row\", \"rows\")", "enCount(1, \"row\", \"rows\")", "enCount(2, \"row\", \"rows\")", "enCount(12345, \"row\", \"rows\")"].map((expression) => scope.evaluate(expression)), ["0 rows", "1 row", "2 rows", "12,345 rows"]);
  scope.evaluate('setInterfaceLanguage("fr")');
  assert.equal(scope.evaluate("currentLanguage"), "zh");
});

// Elements as the browser has them: own text nodes, attributes.
function i18nElement(attributes, nodes) {
  const element = {
    attributes: { ...attributes },
    childNodes: nodes.map((node) => (typeof node === "string" ? { nodeType: 3, nodeValue: node } : node)),
    getAttribute(name) { return Object.prototype.hasOwnProperty.call(this.attributes, name) ? this.attributes[name] : null; },
    setAttribute(name, value) { this.attributes[name] = String(value); }
  };
  return element;
}

test("v3.2.0 language: the static text switches by its own text nodes and attributes, and comes back byte for byte", () => {
  const scope = createScriptContext();
  const count = { nodeType: 1, nodeValue: null, textContent: "（9）" };
  const fields = i18nElement({ "data-i18n": "sidebar.fields" }, ["字段 ", count]);
  const choose = i18nElement({ "data-i18n": "upload.choose" }, ["\n              选择数据文件\n            "]);
  const toggle = i18nElement({ "aria-label": "文件和字段", "data-i18n-attr": "aria-label=sidebar.toggle" }, []);
  const summary = i18nElement({ "data-i18n-open": "收起", "data-i18n-closed": "展开", "data-i18n-attr": "data-i18n-open=details.open data-i18n-closed=details.closed" }, []);
  const unknown = i18nElement({ "data-i18n": "no.such.key" }, ["甲"]);
  put(scope.context, "__i18nElements", [fields, choose, toggle, summary, unknown]);
  scope.evaluate("document.querySelectorAll = (selector) => selector === '[data-i18n], [data-i18n-attr]' ? __i18nElements : []");
  const snapshot = () => JSON.stringify([fields.childNodes.map((node) => node.nodeValue), choose.childNodes[0].nodeValue, toggle.attributes, summary.attributes, unknown.childNodes[0].nodeValue]);
  const chinese = snapshot();
  scope.evaluate('setInterfaceLanguage("en")');
  assert.deepEqual(fields.childNodes.map((node) => node.nodeValue), ["Columns ", null]);
  assert.equal(fields.childNodes[1], count, "the count element stays");
  assert.equal(choose.childNodes[0].nodeValue, "\n              Choose a file\n            ");
  assert.equal(toggle.attributes["aria-label"], "File and columns");
  assert.deepEqual([summary.attributes["data-i18n-open"], summary.attributes["data-i18n-closed"]], ["Collapse", "Expand"]);
  assert.equal(unknown.childNodes[0].nodeValue, "甲", "a key without English keeps the Chinese");
  scope.evaluate('setInterfaceLanguage("zh")');
  assert.equal(snapshot(), chinese);
  scope.evaluate('setInterfaceLanguage("en"); setInterfaceLanguage("en"); setInterfaceLanguage("zh")');
  assert.equal(snapshot(), chinese);
});

test("v3.2.0 language: the page's Chinese is read at startup, so a switch leaves the text and attributes a script wrote", () => {
  const scope = createScriptContext();
  const scopeText = i18nElement({ "data-i18n": "scope.none" }, ["尚未加载数据"]);
  const title = i18nElement({ "data-i18n": "charts.numeric" }, ["数值字段分布（记录频数）"]);
  const canvas = i18nElement({ "aria-label": "数值字段分布图", "data-i18n-attr": "aria-label=charts.numericChart" }, []);
  const untouched = i18nElement({ "data-i18n": "filter.apply", "aria-label": "类别值", "data-i18n-attr": "aria-label=filter.values" }, ["应用筛选"]);
  put(scope.context, "__i18nElements", [scopeText, title, canvas, untouched]);
  scope.evaluate("document.querySelectorAll = (selector) => selector === '[data-i18n], [data-i18n-attr]' ? __i18nElements : []");
  scope.evaluate("rememberStaticLanguage()");
  // After an import the scripts write their own text (textContent puts a new text node in) and
  // the chart's aria-label.
  scopeText.childNodes = [{ nodeType: 3, nodeValue: "40 行" }];
  title.childNodes = [{ nodeType: 3, nodeValue: "Amount（元） 数值分布（记录频数）" }];
  canvas.setAttribute("aria-label", "Amount（元） 数值分布（记录频数）");
  scope.evaluate('setInterfaceLanguage("en")');
  assert.equal(scopeText.childNodes[0].nodeValue, "40 行");
  assert.equal(title.childNodes[0].nodeValue, "Amount（元） 数值分布（记录频数）");
  assert.equal(canvas.attributes["aria-label"], "Amount（元） 数值分布（记录频数）");
  assert.deepEqual([untouched.childNodes[0].nodeValue, untouched.attributes["aria-label"]], ["Apply filters", "Values"]);
  scope.evaluate('setInterfaceLanguage("zh")');
  assert.deepEqual([scopeText.childNodes[0].nodeValue, canvas.attributes["aria-label"], untouched.childNodes[0].nodeValue, untouched.attributes["aria-label"]], ["40 行", "Amount（元） 数值分布（记录频数）", "应用筛选", "类别值"]);
  // script.js reads it before the first switch, in both languages.
  const script = repoFile("script.js");
  assert.ok(script.indexOf("rememberStaticLanguage();") > 0 && script.indexOf("rememberStaticLanguage();") < script.indexOf("if (INITIAL_LANGUAGE !== \"zh\") setInterfaceLanguage(INITIAL_LANGUAGE);"));
});

test("v3.2.0 language: every key of the page has English that keeps its numbers, and the notes are those of format.js", () => {
  const raw = fs.readFileSync(path.join(REPO_ROOT, ["index", "html"].join(".")), "utf8");
  const scope = createScriptContext();
  const english = JSON.parse(scope.evaluate("JSON.stringify(I18N_EN)"));
  const pairs = [];
  for (const match of raw.matchAll(/<([a-z0-9]+)\b([^>]*)>([^<]*)/g)) {
    const attrs = Object.fromEntries([...match[2].matchAll(/([^\s=]+)="([^"]*)"/g)].map((item) => [item[1], item[2]]));
    if (attrs["data-i18n"]) pairs.push([attrs["data-i18n"], match[3].trim()]);
    for (const [name, key] of String(attrs["data-i18n-attr"] || "").split(/\s+/).filter(Boolean).map((pair) => pair.split("="))) pairs.push([key, attrs[name]]);
  }
  assert.ok(pairs.length > 200, `${pairs.length}`);
  for (const [key, chinese] of pairs) {
    if (!chinese) continue;
    assert.ok(english[key] && !ENGLISH_BAD.test(english[key]), `${key}: ${english[key]}`);
    // The one exception: the sample's date written as the English sample writes it (02/18/2025).
    if (key !== "sample.otherDate") assert.deepEqual(numbersOf(english[key]).filter((number) => !ENGLISH_EXTRA_NUMBERS.includes(number)), numbersOf(chinese).filter((number) => !ENGLISH_EXTRA_NUMBERS.includes(number)), `${key}: ${chinese} / ${english[key]}`);
  }
  // One key, one Chinese text: a key used twice names the same words.
  const byKey = new Map();
  for (const [key, chinese] of pairs) {
    if (!chinese) continue;
    if (byKey.has(key)) assert.equal(byKey.get(key), chinese, key);
    byKey.set(key, chinese);
  }
  for (const [key, name] of [["quality.totalsNote", "SUSPECTED_TOTAL_NOTE"], ["quality.formulaNote", "FORMULA_ERROR_NOTE"], ["quality.dateStyleNote", "DATE_STYLE_NOTE"], ["quality.constantNote", "CONSTANT_FIELD_NOTE"], ["quality.lookAlikeNote", "LOOK_ALIKE_NOTE"]]) {
    assert.equal(english[key], scope.evaluate(`${name}[1]`), key);
  }
  // The names the scripts write into the same places.
  assert.equal(english["export.check"], "Check which cells will change");
  assert.equal(english["custom.create"], "Create breakdown");
  assert.equal(english["schema.apply"], "Apply column settings");
  // The eight module titles say what the Chinese title says, not the navigation's short names.
  assert.deepEqual(["overview", "schema", "quality", "filter", "statistics", "charts", "custom", "export"].map((name) => english[`${name}.title`]),
    ["Size and preview", "Check column types", "Blanks, duplicates and other issues", "Choose the rows to analyze", "Number and text stats", "Distributions and trends", "Combine columns as you need", "Save the data and the report"]);
  assert.equal(english["sample.row"], "#");
  assert.equal(english["note.expand"], scope.evaluate('(() => { const button = { getAttribute: () => "true", setAttribute() {}, previousElementSibling: { classList: { toggle() {} } } }; setInterfaceLanguage("en"); toggleSectionNote({ target: { closest: () => button } }); return button.textContent; })()'));
});

test("v3.2.0 language: the switch names the other language, sets <html lang> and ?lang=, keeps the module, and waits during an import", () => {
  const scope = createScriptContext({ search: "?x=1" });
  const attributes = new Map();
  for (const id of ["languageSwitch", "sidebarLanguageSwitch"]) {
    const button = scope.elements.get(id);
    button.setAttribute = (name, value) => attributes.set(`${id}.${name}`, String(value));
  }
  const documentElement = { lang: "zh-CN" };
  scope.context.document.documentElement = documentElement;
  const addresses = [];
  scope.context.window.location.href = "http://127.0.0.1/index.html?x=1#chartsSection";
  scope.context.window.history.replaceState = (state, title, address) => addresses.push(address);
  scope.evaluate("switchInterfaceLanguage()");
  assert.equal(scope.evaluate("currentLanguage"), "en");
  assert.equal(documentElement.lang, "en");
  assert.equal(scope.elements.get("languageSwitch").textContent, "中文");
  assert.equal(attributes.get("languageSwitch.lang"), "zh-CN");
  assert.equal(attributes.get("sidebarLanguageSwitch.lang"), "zh-CN");
  assert.deepEqual(addresses, ["/index.html?x=1&lang=en#chartsSection"]);
  scope.context.window.location.href = "http://127.0.0.1/index.html?x=1&lang=en#chartsSection";
  scope.evaluate("switchInterfaceLanguage()");
  assert.equal(documentElement.lang, "zh-CN");
  assert.equal(scope.elements.get("sidebarLanguageSwitch").textContent, "English");
  assert.equal(attributes.get("sidebarLanguageSwitch.lang"), "en");
  assert.equal(addresses[1], "/index.html?x=1&lang=zh#chartsSection");
  // While a file is read, or a worksheet waits: disabled, and a click does nothing.
  scope.evaluate("setImportBusy(true)");
  assert.deepEqual([scope.elements.get("languageSwitch").disabled, scope.elements.get("sidebarLanguageSwitch").disabled], [true, true]);
  scope.evaluate("switchInterfaceLanguage()");
  assert.equal(scope.evaluate("currentLanguage"), "zh");
  scope.evaluate("setImportBusy(false)");
  assert.equal(scope.elements.get("languageSwitch").disabled, false);
  scope.evaluate("state.pendingExcel = { workbook: {}, fileName: 'a.xlsx', importId: 1 }; syncLanguageSwitches()");
  assert.equal(scope.elements.get("languageSwitch").disabled, true);
  scope.evaluate("switchInterfaceLanguage()");
  assert.equal(scope.evaluate("currentLanguage"), "zh");
  scope.evaluate("state.pendingExcel = null; syncLanguageSwitches()");
  assert.equal(scope.elements.get("sidebarLanguageSwitch").disabled, false);
  // script.js: both buttons bound, the worksheet panel followed.
  const script = repoFile("script.js");
  assert.match(script, /\[dom\.languageSwitch, dom\.sidebarLanguageSwitch\]\.forEach\(\(button\) => button\.addEventListener\("click", switchInterfaceLanguage\)\);/);
  assert.match(script, /new MutationObserver\(syncLanguageSwitches\)\.observe\(dom\.sheetSelectionPanel, \{ attributes: true, attributeFilter: \["class"\] \}\);/);
});

test("v3.2.0 language: two switches, in the top bar of the upload page and at the sidebar's foot, written in the other language", () => {
  const html = indexHtmlSource();
  const header = html.slice(html.indexOf('<header class="app-header">'), html.indexOf("</header>"));
  assert.match(header, /<div class="header-actions">\s*<button id="languageSwitch" class="secondary-button language-switch" type="button" lang="en">English<\/button>\s*<a class="nav-new-analysis secondary-button" href="#dropZone">换一个文件<\/a>/);
  const sidebar = html.slice(html.indexOf('<aside id="appSidebar"'), html.indexOf("</aside>"));
  assert.match(sidebar, /<p class="sidebar-privacy">表格只在这台电脑的浏览器里处理，不会上传。<\/p>\s*<button id="sidebarLanguageSwitch" class="secondary-button language-switch sidebar-language-switch" type="button" lang="en">English<\/button>\s*$/);
  assert.equal((html.match(/class="[^"]*\blanguage-switch\b/g) || []).length, 2);
  const css = repoFile("style.css");
  assert.match(css, /body\.has-results \.header-actions \.language-switch \{\s*display: none;\s*\}/);
  assert.deepEqual(domKeys(createScriptContext()).filter((key) => /language/i.test(key)), ["languageSwitch", "sidebarLanguageSwitch"]);
});

// ---- v3.2.0 commit 5: the sentences of format.js in English ----
// The 18 constants set at load time: [中文, English] pairs (Lp), read with Lpick() or [0] / [1].
const FORMAT_PAIRED_CONSTANTS = ["SUSPECTED_TOTAL_NOTE", "FORMULA_ERROR_NOTE", "EXCLUSION_KIND_LABELS", "EXCLUSION_DRAFT_REFUSAL", "NO_ID_DUPLICATE_NOTE", "CATEGORY_STATS_EMPTY_NOTE",
  "CATEGORY_CHART_EMPTY_NOTE", "GO_TO_QUALITY_LABEL", "PLACEHOLDER_MAYBE_NOTE", "PLACEHOLDER_DASH_NOTE", "DATE_STYLE_NOTE", "CONSTANT_FIELD_NOTE", "LOOK_ALIKE_NOTE", "DATE_STYLE_NAMES",
  "VISIBLE_TEXT_NAMES", "LOOK_ALIKE_REASON_NAMES", "STANDARDIZE_ACTION_LABELS", "STANDARDIZE_ORDER_HINT"];
const ENGLISH_BAD = /[\p{Script=Han}\u3000-\u303F\uFF01-\uFF0F\uFF1A-\uFF20\uFF3B-\uFF40\uFF5B-\uFF65]/u;
// The one exception (commit 7): the trend note writes its last date as “up to YYYY-MM-DD” in
// English and “截至 M 月 D 日” in Chinese; that date counts as its month and day, without zeros
// (its year is in the period's name in both languages).
const numbersOf = (text) => (String(text).replace(/(only has data up to )(\d{4})-(\d{2})-(\d{2})(?=,)/g, (_, words, year, month, day) => `${words}${Number(month)} ${Number(day)}`).match(/\d[\d,.]*\d|\d/g) || []).sort();

test("v3.2.0 format.js: the 18 constants are frozen [中文, English] pairs, and the English has no Chinese", () => {
  const scope = createScriptContext();
  assert.equal(FORMAT_PAIRED_CONSTANTS.length, 18);
  let pairs = 0;
  for (const name of FORMAT_PAIRED_CONSTANTS) {
    const value = scope.context[name] ?? scope.evaluate(name);
    const list = Array.isArray(value) ? [value] : Object.values(value);
    assert.ok(Object.isFrozen(value), name);
    for (const pair of list) {
      assert.ok(Array.isArray(pair) && pair.length === 2 && Object.isFrozen(pair), name);
      assert.match(pair[0], /\p{Script=Han}/u, name);
      // The one English that is not a fixed text: the maybe-valid note names this column's writings (R11).
      const english = typeof pair[1] === "function" ? pair[1](["NA"]) : pair[1];
      assert.equal(typeof english, "string", name);
      assert.ok(english.length > 0 && !ENGLISH_BAD.test(english), `${name}: ${english}`);
      pairs += 1;
    }
  }
  // 13 single sentences, EXCLUSION_KIND_LABELS 2, DATE_STYLE_NAMES 10, VISIBLE_TEXT_NAMES 13, LOOK_ALIKE_REASON_NAMES 4, STANDARDIZE_ACTION_LABELS 13.
  assert.equal(pairs, 13 + 2 + 10 + 13 + 4 + 13);
  // The page notes, as index.html will have them in English (commit 9 checks I18N_EN against these).
  assert.equal(scope.evaluate("SUSPECTED_TOTAL_NOTE[1]"), "These rows still count as ordinary data, so their numbers are counted again in sums, averages and the check for unusually high or low values.");
  assert.equal(scope.evaluate("LOOK_ALIKE_NOTE[1]").split(": ")[0], "Leading or trailing spaces");
});

// Every function of format.js that writes text, called with the same arguments in both languages.
const FORMAT_TEXT_CALLS = [
  'formatConversionFailures(3, ["N/A", "x1", "about 500"], "3")',
  'numericAxisLabel("Weight", "Weight")', 'numericAxisLabel("Rate", "Rate")', 'formatFieldLabelWithUnit("Amount", "kg")',
  'formatNumberConversion({ field: "Weight", unit: "kg", unitCount: 5, unitSamples: ["5kg", "6kg", "7kg"], magnitudeCount: 2, fullwidthCount: 1, conflictUnits: [] })',
  'formatNumberConversion({ field: "Weight", unit: "", unitCount: 0, unitSamples: [], magnitudeCount: 0, fullwidthCount: 0, conflictUnits: [{ unit: "kg", count: 5 }, { unit: "t", count: 1 }] })',
  'formatNumericLookingHint({ field: "Floor", valueCount: 6, numericCount: 4, ratio: 0.6667, samples: ["B1", "B2"] })',
  'formatTotalMatchedFields(["a", "b", "c", "d", "e", "f", "g"])', 'formatTotalMatchedFields(["a", "b"])',
  'formatTotalAddends({ scope: "segment", firstLine: 5, lastLine: 32, rowCount: 28, contiguous: true })',
  'formatTotalAddends({ scope: "above", firstLine: 2, lastLine: 13, rowCount: 9, contiguous: false })',
  'formatTotalAddends({ scope: "segment", firstLine: 9, lastLine: 9, rowCount: 1, contiguous: true })',
  'formatTotalAddends({ scope: "all", firstLine: null, lastLine: null, rowCount: 12, contiguous: true })',
  'formatSuspectedTotal({ line: 33, label: "Total", matches: [{ fields: ["Sales", "Cost"], scope: "segment", firstLine: 5, lastLine: 32, rowCount: 28, contiguous: true }, { fields: ["Qty"], scope: "all", firstLine: null, lastLine: null, rowCount: 9, contiguous: true }] })',
  'formatSuspectedTotal({ line: null, label: "Sum", matches: [{ fields: ["Qty"], scope: "above", firstLine: 3, lastLine: 3, rowCount: 1, contiguous: true }] })',
  "formatSuspectedTotalOverflow(21)",
  'formatSuspectedTotalStatus([{ line: 44, label: "Total" }])', 'formatSuspectedTotalStatus([{ line: null, label: "Total" }])',
  'formatSuspectedTotalStatus([1, 2, 3, 4, 5, 6, 7].map((line) => ({ line, label: "Subtotal" })))', 'formatSuspectedTotalStatus([{ line: 5, label: "A" }, { line: 8, label: "B" }])',
  "formatExclusionLines([44])", "formatExclusionLines([1, 2, 3, 4, 5, 6, 7])",
  "formatExclusionRestored(1)", "formatExclusionRestored(3)", "formatTotalsFlaggedAgain([5])", "formatTotalsFlaggedAgain([5, 8])", "formatTotalsFlaggedAgain([null, null])",
  "formatExcludedTail(2)", "formatAnalysisScope(1, 28, 40, 2)", "formatAnalysisScope(2, 28, 40, 0)", "formatAnalysisScope(0, 40, 40, 2)", "formatAnalysisScope(0, 42, 42, 0)", "formatAnalysisScope(0, 1, 1, 0)",
  "formatAllIncludedRows(40, 2)", "formatAllIncludedRows(42, 0)", "formatDataExportExcludedNote(2)", "formatReportNoFilterText(40, 2)", "formatNoFilterMessage(40, 2)", "formatNoFilterMessage(42, 0)",
  "formatFiltersApplied(2, 17, 23)", "formatFiltersApplied(1, 0, 40)", "formatPreviewPageInfo(1, 3, 120, 2, false)", "formatPreviewPageInfo(2, 3, 120, 2, true)",
  'formatCustomAnalysisScope({ groupField: "Region", metricField: "Amount", aggregateLabel: "sum", filtered: true, currentCount: 17, includedCount: 40, totalCount: 42, excludedCount: 2 })',
  'formatCustomAnalysisScope({ groupField: "Region", metricField: "Amount", aggregateLabel: "sum", filtered: false, currentCount: 40, includedCount: 40, totalCount: 42, excludedCount: 2 })',
  'formatCustomAnalysisScope({ groupField: "Region", metricField: "Amount", aggregateLabel: "sum", filtered: false, currentCount: 42, includedCount: 42, totalCount: 42, excludedCount: 0 })',
  "formatQualityScopeNote(42, 2, 6)", "formatQualityScopeNote(42, 1, 0)", "formatSuspectedTotalExcludedNote(1)", "formatSuspectedTotalExcludedNote(5)",
  "formatSingleRowTotals(1)", "formatSingleRowTotals(6)", "formatUnmatchedTotalLabels(95, 6, true)", "formatUnmatchedTotalLabels(7, 6, true)",
  "formatReportExclusions({ suspectedTotal: { count: 1, lines: [44] }, duplicateRows: { count: 30, lines: Array.from({ length: 20 }, (_, i) => i + 2) } })",
  "formatReportExclusions({ suspectedTotal: { count: 2, lines: [6, 7] }, duplicateRows: { count: 2, lines: [4, 7] }, overlap: { count: 1, lines: [7] } })",
  "formatReportExclusions({ suspectedTotal: { count: 2, lines: null }, duplicateRows: { count: 2, lines: null }, overlap: { count: 1, lines: null } })",
  "formatDuplicateIncludedNote(1)", "formatDuplicateIncludedNote(3)", "formatDuplicateExcludedNote(1)", "formatDuplicateExcludedNote(3)",
  "formatAnalysisContainsNote(1, 1)", "formatAnalysisContainsNote(2, 0)", "formatAnalysisContainsNote(0, 3)", "formatGoToQualityAriaLabel(1, 1)", "formatGoToQualityAriaLabel(0, 1)",
  'formatExclusionMessage({ kind: "suspectedTotal", on: true, rows: 1, changed: 1, before: 42, after: 41 })',
  'formatExclusionMessage({ kind: "duplicateRows", on: true, rows: 2, changed: 0, before: 40, after: 40 })',
  'formatExclusionMessage({ kind: "suspectedTotal", on: true, rows: 1, changed: 0, before: 40, after: 40 })',
  'formatExclusionMessage({ kind: "duplicateRows", on: false, rows: 3, changed: 3, before: 39, after: 42 })',
  'formatExclusionMessage({ kind: "suspectedTotal", on: false, rows: 1, changed: 0, before: 41, after: 41 })',
  "formatFormulaErrorStatus([{ count: 3 }, { count: 4 }])", 'formatStatusNoteSummary("orders.csv", 2)', 'formatStatusNoteSummary("orders.csv", 1)',
  "formatSchemaSearchPlaceholder(21)", 'formatReturnToAnalysis("orders.csv")',
  'formatFormulaError({ field: "Margin", count: 3, byError: { "#DIV/0!": 2, "#REF!": 1 }, lines: [3, 5, 9] })',
  'formatFormulaError({ field: "Margin", count: 12, byError: { "#N/A": 12 }, lines: [2, 3, 4, 5, 6, 7, 8, 9, 10, 11], fieldIgnored: true })',
  'formatFormulaError({ field: "Margin", count: 1, byError: { "#VALUE!": 1 }, lines: [null] })',
  'formatHeaderDetectionNote({ headerRowNumber: 2, skippedRows: [{ preview: "Sales 2025" }] })',
  'formatHeaderDetectionNote({ headerRowNumber: 5, skippedRows: [{ preview: "A" }, { preview: "B" }, { preview: "C" }] })', 'formatHeaderDetectionNote({ headerRowNumber: 3, skippedRows: [] })',
  'formatCsvWarningMessage("TooFewFields", { line: 12, cellCount: 2, headerCount: 4 })', 'formatCsvWarningMessage("TooFewFields", { line: 12 })',
  'formatCsvWarningMessage("TooFewFields", { cellCount: 1, headerCount: 4 })', 'formatCsvWarningMessage("TooFewFields")',
  'formatCsvWarningMessage("InvalidQuotes", { line: 7 })', 'formatCsvWarningMessage("InvalidQuotes")',
  'formatCsvWarningMessage("UndetectableDelimiter", { separator: ";" })', 'formatCsvWarningMessage("UndetectableDelimiter", { separator: "|" })',
  'formatCsvWarningMessage("UndetectableDelimiter", { separator: String.fromCharCode(9) })',
  ...["missing", "conversion", "outlier", "duplicateRows", "idDuplicate", "placeholder", "suspectedTotal", "formulaError", "dateStyle", "lookAlike", "parseWarnings"].flatMap((kind) => [
    `formatProblemViewLabel({ kind: "${kind}", field: "Amount", key: "N/A", count: 3 })`,
    `formatProblemCategoryLabel({ kind: "${kind}", field: "Amount", key: "N/A", count: 3 })`,
    `formatProblemRowsSummary({ kind: "${kind}" }, { count: 2, groupCount: 1, items: [1, 2, 3] }, 1, 2)`
  ]),
  "formatRowNumberList([6, 19])", "formatRowNumberList([1, 2, 3, 4, 5, 6, 7, 8, 9, 10], 15)", "formatIdenticalRows(2, 1)", "formatIdenticalRows(4, 2)",
  "formatIdDuplicateOverflow(12, 10)", 'formatCleanIdFieldsNote(["ID", "Name"])', 'formatCleanIdFieldsNote(["ID"], true)',
  'formatMostlyRepeatedIdHint({ field: "User", distinctCount: 5, nonMissingCount: 120, uniqueRatio: 0.0417 })',
  'formatPlaceholderValue({ key: "N/A", rawValues: ["N/A", "n/a"] })', "formatPlaceholderRows([15], 1)", "formatPlaceholderRows([2, 3], 5)", "formatPlaceholderRows([null], 1)",
  'formatPlaceholderChange("Amount", 0, "numeric", { typeKey: "numeric", missingCount: 1 })', 'formatPlaceholderChange("Amount", 2, "category", { typeKey: "numeric", missingCount: 5 })',
  'formatPendingPlaceholder({ field: "Amount", key: "N/A", rawValues: ["N/A"], count: 1, rowNumbers: [15] }, formatPlaceholderChange("Amount", 0, "numeric", { typeKey: "numeric", missingCount: 1 }))',
  'formatConfirmedPlaceholder({ field: "Note", key: "N/A", rawValues: ["N/A"], count: 2, fieldIgnored: true })',
  'formatPlaceholderNotes(["-", "N/A"])', 'formatPlaceholderNotes(["NA", "/"])',
  'formatPlaceholderCounts("N/A", { missing: 0, failures: 1, typeKey: "numeric", listed: true }, { missing: 1, failures: 0, typeKey: "numeric", listed: false }, true)',
  'formatPlaceholderCounts("-", { missing: 1, failures: 0, typeKey: "category", listed: false }, { missing: 0, failures: 0, typeKey: "category", listed: true }, false)',
  'formatPlaceholderTypeChange("Amount", "category", "numeric")', "formatPlaceholderRowCount(42, false)", "formatPlaceholderRowCount(40, true)",
  'formatPlaceholderConfirmMessage({ field: "Amount", key: "N/A", count: 1, before: { missing: 0, failures: 1, typeKey: "numeric", listed: false }, after: { missing: 1, failures: 0, typeKey: "numeric", listed: false }, rowCount: 42 })',
  'formatPlaceholderConfirmMessage({ field: "Date", key: "-", count: 3, before: { missing: 0, failures: 3, typeKey: "date", listed: false }, after: { missing: 3, failures: 0, typeKey: "date", listed: false }, rowCount: 40, rowsLeftOut: true })',
  'formatPlaceholderConfirmMessage({ field: "Size", key: "none", count: 4, before: { missing: 0, failures: 0, typeKey: "category", listed: false }, after: { missing: 4, failures: 0, typeKey: "numeric", listed: false }, rowCount: 42 })',
  'formatPlaceholderUndoMessage({ field: "Amount", key: "N/A", count: 1, before: { missing: 1, failures: 0, typeKey: "numeric", listed: false }, after: { missing: 0, failures: 1, typeKey: "numeric", listed: false }, rowCount: 42 })',
  'formatPlaceholderUndoMessage({ field: "Size", key: "none", count: 4, before: { missing: 4, failures: 0, typeKey: "numeric", listed: false }, after: { missing: 0, failures: 0, typeKey: "category", listed: false }, rowCount: 40, rowsLeftOut: true })',
  'formatPlaceholderUndoMessage({ field: "Note", key: "N/A", count: 2, ignored: true })',
  'formatAllPlaceholdersMessage([{ field: "Amount", key: "N/A" }, { field: "Rep", key: "--" }], [{ field: "Amount", typeKey: "numeric", previousTypeKey: "category" }], 42, false)',
  'formatAllPlaceholdersMessage([{ field: "Amount", key: "N/A" }], [], 40, true)',
  'formatDateStyleLine({ field: "Hired", styles: [{ key: "excelCell", count: 24, sample: "2016-04-01" }, { key: "ymdSlash", count: 2, sample: "2021/3/5" }, { key: "english", count: 4, sample: "Mar 5, 2021" }], otherCount: 6 })',
  'formatDateStyleLine({ field: "Hired", styles: [{ key: "ymdDash", count: 9, sample: "2021-03-05" }, { key: "dayMonthSame", count: 1, sample: "05/05/2021" }], otherCount: 1 })',
  'formatConstantField({ field: "Note", value: "OK", count: 6, missingCount: 3 })', 'formatConstantField({ field: "Note", value: " OK", count: 1, missingCount: 0 })',
  'formatLookAlikeReasons(["compat", "case", "invisible", "innerSpace"])',
  'formatLookAlikeLine({ field: "Region", edgeSpaceCount: 2, edgeSample: "East ", invisibleOnlyCount: 1, groups: [{ variants: [{ text: "HD", count: 2 }, { text: "hd", count: 1 }], variantCount: 3, reasons: ["case"] }], groupCount: 3, rowCount: 18 })',
  'formatLookAlikeLine({ field: "Region", edgeSpaceCount: 0, edgeSample: "", invisibleOnlyCount: 0, groups: [{ variants: [{ text: "a  b", count: 1 }, { text: "a b", count: 4 }], reasons: ["innerSpace"] }], groupCount: 1, rowCount: 5 })',
  'formatDateStyleRowNote("ymdSlash", "excelCell")', 'formatLookAlikeRowNote("East ", { position: 0, reasons: ["case", "compat"] }, true, true)', 'formatLookAlikeRowNote("x", null, true, false)',
  ...["date", "number", "placeholder", "trim", "keptTime", "keptPercent", "keptScientific", "keptCurrencyMix", "keptPendingPlaceholder", "keptDateUnreadable", "keptDateOrder", "keptNumberUnreadable", "keptUnsafe"].map((action) => `formatStandardizeAction("${action}")`),
  "formatStandardizedExportGate(false, false, false)", "formatStandardizedExportGate(true, true, false)", "formatStandardizedExportGate(true, false, false)",
  'formatStandardizeFieldNotes({ kind: "year", unit: "kg", number: 3, field: "Weight", orderHint: true }).join(" | ")',
  'formatStandardizeKept({ kept: 2, keptSamples: ["2024-05-08 10:30", "TBD"] })',
  'formatStandardizedChangesStatus(100, 250, "orders_changes.csv")', 'formatStandardizedChangesStatus(3, 3, "orders_changes.csv")',
  "formatStandardizedSummary({ changed: 12, kept: 3, quietCount: 4, ignoredCount: 1 })", "formatStandardizedSummary({ changed: 0, kept: 0, quietCount: 0, ignoredCount: 2 })",
  "formatStandardizedSummary({ changed: 1, kept: 0, quietCount: 0, ignoredCount: 0 })"
];
// Numbers the English adds on purpose: 万 and 亿 written out.
const ENGLISH_EXTRA_NUMBERS = ["10,000", "100,000,000"];

test("v3.2.0 format.js: every sentence in English keeps the Chinese sentence's numbers and has no Chinese", () => {
  const scope = createScriptContext();
  put(scope.context, "__unitRows", Array.from({ length: 5 }, (_, index) => ({ Weight: `${index + 5}kg`, Rate: `${(index + 1) * 10}%`, Amount: String(index) })));
  scope.evaluate('commitTabularData(["Weight", "Rate", "Amount"], __unitRows, "units.csv", beginImport())');
  // The type names come from profiling.js and are paired in commit 7; until then, their keys.
  scope.evaluate("fieldTypeLabel = (typeKey) => typeKey");
  const failures = [];
  for (const call of FORMAT_TEXT_CALLS) {
    scope.evaluate('setInterfaceLanguage("zh")');
    const zh = scope.evaluate(call);
    scope.evaluate('setInterfaceLanguage("en")');
    const en = scope.evaluate(call);
    if (typeof zh !== "string" || typeof en !== "string") { failures.push(`${call}: not a string`); continue; }
    if (!zh) { failures.push(`${call}: empty`); continue; }
    if (ENGLISH_BAD.test(en)) failures.push(`${call}: Chinese in the English: ${en}`);
    if (/\p{Script=Han}|[、。，：；（）「」]/u.test(zh) && zh === en) failures.push(`${call}: not translated: ${en}`);
    const zhNumbers = numbersOf(zh);
    const enNumbers = numbersOf(en);
    const left = [...enNumbers];
    for (const number of zhNumbers) {
      const at = left.indexOf(number);
      if (at < 0) failures.push(`${call}: ${number} missing from the English: ${zh} / ${en}`);
      else left.splice(at, 1);
    }
    const extra = left.filter((number) => !ENGLISH_EXTRA_NUMBERS.includes(number));
    if (extra.length) failures.push(`${call}: English has ${extra.join(" ")} more: ${zh} / ${en}`);
  }
  assert.deepEqual(failures, []);
  assert.ok(FORMAT_TEXT_CALLS.length > 150, `${FORMAT_TEXT_CALLS.length}`);
});

test("v3.2.0 format.js: the English sentences read as written", () => {
  const scope = createScriptContext();
  scope.evaluate('setInterfaceLanguage("en")');
  const en = (expression) => scope.evaluate(expression);
  assert.equal(en("formatAnalysisScope(1, 28, 40, 2)"), "1 filter · 28 / 40 rows (2 more excluded from analysis)");
  assert.equal(en("formatAnalysisScope(0, 42, 42, 0)"), "All 42 rows");
  assert.equal(en("formatSuspectedTotalExcludedNote(1)"), "1 row is excluded from analysis: stats, charts, breakdowns and filters no longer include it.");
  assert.equal(en("formatReportExclusions({ suspectedTotal: { count: 1, lines: [44] }, duplicateRows: { count: 1, lines: [21] } })"), "Excluded from analysis: 1 possible total row (row 44), 1 duplicate row (row 21).");
  assert.equal(en("formatReportExclusions({ suspectedTotal: { count: 2, lines: [6, 7] }, duplicateRows: { count: 2, lines: [4, 7] }, overlap: { count: 1, lines: [7] } })"),
    "Excluded from analysis: 2 possible total rows (rows 6, 7), 2 duplicate rows (rows 4, 7), 3 rows in all (row 7 is both).");
  assert.equal(en('formatSuspectedTotalStatus([{ line: 44, label: "Total" }])'), "Row 44 “Total” looks like a total row; you can exclude it from analysis in Data quality.");
  assert.equal(en('formatFormulaError({ field: "Margin", count: 3, byError: { "#DIV/0!": 2, "#REF!": 1 }, lines: [3, 5, 9] })'), "“Margin”: 3 (#DIV/0! 2, #REF! 1), rows 3, 5, 9.");
  assert.equal(en('formatHeaderDetectionNote({ headerRowNumber: 2, skippedRows: [{ preview: "Sales 2025" }] })'), "The header is on row 2; 1 row above it was skipped: “Sales 2025”.");
  assert.equal(en('formatProblemCategoryLabel({ kind: "missing", field: "Region", count: 1 })'), "Blanks · Region (1 row)");
  assert.equal(en('formatCsvWarningMessage("TooFewFields", { line: 12, cellCount: 2, headerCount: 4 })'), "Row 12 has only 2 columns (the header has 4); the missing cells are treated as blank.");
  assert.equal(en('formatStandardizeAction("date")'), "Date rewritten as YYYY-MM-DD");
  assert.equal(en('formatVisibleText(" a" + String.fromCharCode(9) + "b ")'), "␣a[tab]b␣");
  assert.equal(en("formatPreviewPageInfo(1, 1, 1, 0, false)"), "Page 1 of 1, 1 row");
  // Back in Chinese the same calls give the Chinese.
  scope.evaluate('setInterfaceLanguage("zh")');
  assert.equal(scope.evaluate("formatAnalysisScope(0, 42, 42, 0)"), "全部 42 行");
  assert.equal(scope.evaluate('formatVisibleText(" a" + String.fromCharCode(9) + "b ")'), "␣a[制表符]b␣");
});

test("v3.2.0 format.js: the English maybe-valid note names only the writings this column has (R11)", () => {
  const scope = createScriptContext();
  scope.evaluate('setInterfaceLanguage("en")');
  assert.equal(scope.evaluate('formatPlaceholderNotes(["N/A", "-"])'), "“-” may also be a real value (for example a yes/no answer); make sure that in this column it means there is no data. In accounting formats, “-” usually means 0.");
  assert.equal(scope.evaluate('formatPlaceholderNotes(["NA", "/"])'), "“NA”, “/” may also be real values (for example a yes/no answer, or NA for North America); make sure that in this column they mean there is no data.");
  assert.equal(scope.evaluate('formatPlaceholderNotes(["NA"])'), "“NA” may also be a real value (for example NA for North America); make sure that in this column it means there is no data.");
  assert.equal(scope.evaluate('formatPlaceholderNotes(["N/A", "NULL"])'), "");
  // Chinese keeps its one fixed sentence.
  scope.evaluate('setInterfaceLanguage("zh")');
  assert.equal(scope.evaluate('formatPlaceholderNotes(["NA", "/"])'), PLACEHOLDER_MAYBE_NOTE_TEXT);
});

test("v3.2.0 format.js: the per-row texts never call L() (both arguments would be built for every row)", () => {
  const scope = createScriptContext();
  for (const language of ["zh", "en"]) {
    scope.evaluate(`setInterfaceLanguage("${language}")`);
    scope.evaluate("globalThis.__lCalls = 0; globalThis.__realL ??= L; L = (...args) => { globalThis.__lCalls += 1; return globalThis.__realL(...args); }");
    scope.evaluate('formatDateStyleRowNote("ymdSlash", "excelCell"); formatLookAlikeRowNote("East ", { position: 0, reasons: ["case"] }, true, true); formatStandardizeAction("trim"); formatVisibleText(" x "); formatCsvWarningMessage("TooFewFields", { line: 3, cellCount: 1, headerCount: 2 }); formatCsvWarningMessage("InvalidQuotes", { line: 3 })');
    assert.equal(scope.evaluate("__lCalls"), 0, language);
    // The counter itself works.
    scope.evaluate('formatExcludedTail(2)');
    assert.equal(scope.evaluate("__lCalls"), 1, language);
    scope.evaluate("L = globalThis.__realL");
  }
});

// ---- v3.2.0 commit 6: the sentences of app.js in English ----
// A table with English names and values that reaches most of app.js: an ID column with a repeated
// ID, a repeated row, N/A, a total row, a short row (a reading issue), two date formats, "east" next
// to "East", a value with a trailing space, a blank region and regions whose order depends on the
// sort key ("(A)", "Z", blank; R3).
function englishScenarioCsv() {
  const regions = ["East", "(A)", "Z", "North"];
  const rows = Array.from({ length: 16 }, (_, index) => ({
    id: `A${String(index + 1).padStart(2, "0")}`,
    region: index === 5 ? "east" : index === 9 ? "" : regions[index % 4],
    amount: index === 3 ? "N/A" : String((index + 1) * 100),
    date: index === 7 || index === 11 ? `01/${String(index + 2).padStart(2, "0")}/2025` : `2025-01-${String(index + 2).padStart(2, "0")}`,
    rep: index === 4 ? "Cy " : ["Ann", "Bob", "Cy", "Dee"][index % 4]
  }));
  const line = (row) => [row.id, row.region, row.amount, row.date, row.rep].join(",");
  const amounts = [...rows, rows[2]].map((row) => Number(row.amount)).filter(Number.isFinite);
  return ["ID,Region,Amount,Date,Rep", ...rows.map(line), line(rows[2]), "A17,East", `Total,,${amounts.reduce((sum, value) => sum + value, 0)},,`].join("\r\n") + "\r\n";
}

async function importCsvInLanguage(text, fileName, language) {
  const scope = createScriptContext({ language });
  put(scope.context, "XLSX", XLSX);
  scope.context.window.XLSX = XLSX;
  put(scope.context, "Papa", Papa);
  scope.context.window.Papa = Papa;
  const bytes = Buffer.from(text, "utf8");
  put(scope.context, "__languageFile", { name: fileName, size: bytes.length, arrayBuffer: async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.length) });
  await scope.evaluate("parseDataFile(__languageFile)");
  return scope;
}

// Every element's text, markup, options and attributes (the fake setAttribute keeps nothing, so
// it is recorded on the shared prototype for the length of the scenario).
function elementTexts(scope) {
  const out = {};
  for (const [id, element] of scope.elements) {
    out[id] = [element.textContent, element.innerHTML, element.value, ...(element.options || []).map((option) => `${option.textContent ?? ""}`), ...Object.values(element.__attributes || {})].join(" ¦ ");
  }
  return out;
}

// The steps, the same in both languages; texts after each.
async function runLanguageScenario(language) {
  const scope = await importCsvInLanguage(englishScenarioCsv(), "orders.csv", language);
  const steps = [["import", elementTexts(scope)]];
  const step = (name, expression) => { scope.evaluate(expression); steps.push([name, elementTexts(scope)]); };
  step("blanks", "confirmAllCommonPlaceholders()");
  step("exclude", `["suspectedTotal", "duplicateRows"].forEach((kind) => handleExclusionClick({ target: { closest: () => ({ dataset: { exclusionKind: kind, exclusionOn: "true" } }) } }))`);
  step("filter", `(() => {
    const values = getCategoryFilterOptions("Region").map((option) => option.value);
    dom.categoryFilterField.value = "Region"; dom.numericFilterField.value = ""; dom.dateFilterField.value = "";
    state.categoryFilterDraftValues = new Set(values.slice(0, 3));
    applyFilters();
  })()`);
  step("field type", 'stageFieldTypeChange("Rep", "ignore"); applyFieldConfiguration()');
  step("custom", 'dom.v2MetricField.value = "Amount"; dom.v2GroupField.value = "Region"; dom.v2AggregateMethod.value = "sum"; dom.v2DateField.value = "Date"; renderV2Analysis()');
  step("standardize", "checkStandardizedExport()");
  const categories = JSON.parse(scope.evaluate("JSON.stringify(state.problemCategories)"));
  categories.forEach((category, index) => {
    put(scope.context, "__scenarioCategory", category);
    step(`problem rows ${index}`, "openProblemView(__scenarioCategory)");
  });
  step("status fold", "toggleStatusNote()");
  step("clear filters", "clearAllFilters()");
  return { scope, steps, categories };
}

async function withRecordedAttributes(action) {
  const proto = Object.getPrototypeOf(createScriptContext().elements.get("languageSwitch") || createScriptContext().evaluate("dom.languageSwitch"));
  const original = proto.setAttribute;
  proto.setAttribute = function setAttribute(name, value) { (this.__attributes ??= {})[name] = String(value); };
  try {
    return await action();
  } finally {
    proto.setAttribute = original;
  }
}

// The Chinese of the scripts' L() / Lp() calls, in pieces of two or more Han characters, without
// the pieces the whitelisted data literals also hold (“空值” of BLANK_GROUP_KEY, say).
function doneFileChineseFragments() {
  const pieces = (file, chineseArgumentOnly) => new Set(i18nGuard.lex(repoFile(file))
    .filter((token) => (token.type === "str" || token.type === "tmpl") && (!chineseArgumentOnly || (token.lFrame && token.lFrame.arg === 0)))
    .flatMap((token) => token.value.match(/\p{Script=Han}{2,}/gu) || []));
  const done = new Set(i18nGuard.APP_SCRIPT_NAMES.flatMap((file) => [...pieces(file, true)]));
  for (const file of i18nGuard.APP_SCRIPT_NAMES) {
    // Also a piece inside one of the data literals outside L().
    const texts = i18nGuard.lex(repoFile(file)).filter((token) => (token.type === "str" || token.type === "tmpl") && !token.lFrame).map((token) => token.value);
    for (const piece of [...done]) if (texts.some((text) => text.includes(piece))) done.delete(piece);
  }
  return [...done];
}

test("v3.2.0 app.js: in English the page never shows the Chinese of the files that are done, and every element keeps the numbers of its Chinese text", async () => {
  const [zh, en] = await withRecordedAttributes(async () => [await runLanguageScenario("zh"), await runLanguageScenario("en")]);
  assert.equal(en.scope.evaluate("currentLanguage"), "en");
  assert.ok(zh.categories.length >= 8, `${zh.categories.length} problem categories`);
  assert.deepEqual(en.categories, zh.categories);
  const fragments = doneFileChineseFragments();
  assert.ok(fragments.length > 300, `${fragments.length} fragments`);
  const leaks = [];
  const numberGaps = [];
  en.steps.forEach(([name, texts], stepIndex) => {
    const zhTexts = zh.steps[stepIndex][1];
    assert.equal(zh.steps[stepIndex][0], name);
    for (const [id, text] of Object.entries(texts)) {
      const leaked = fragments.filter((fragment) => text.includes(fragment));
      if (leaked.length) leaks.push(`${name} #${id}: ${leaked.slice(0, 3).join(" / ")} in ${text.slice(0, 120)}`);
      const zhNumbers = numbersOf(zhTexts[id] || "").filter((number) => !ENGLISH_EXTRA_NUMBERS.includes(number));
      const enNumbers = numbersOf(text).filter((number) => !ENGLISH_EXTRA_NUMBERS.includes(number));
      if (zhNumbers.join(" ") !== enNumbers.join(" ")) numberGaps.push(`${name} #${id}: ${zhNumbers.join(" ")} / ${enNumbers.join(" ")}`);
    }
  });
  assert.deepEqual(leaks.slice(0, 15), [], `${leaks.length} leaks`);
  assert.deepEqual(numberGaps.slice(0, 15), [], `${numberGaps.length} elements with other numbers`);
  // The steps really ran: a filter, two kinds left out, a custom analysis, a standardized check.
  assert.equal(en.scope.evaluate("state.excludedRows.suspectedTotal.length + state.excludedRows.duplicateRows.length"), 2);
  assert.ok(en.scope.evaluate("Boolean(state.customAnalysis && state.standardizedExport)"));
});

test("v3.2.0 app.js: the filter values keep the Chinese order in both languages (R3); only the empty one's label changes", () => {
  const orders = [];
  for (const language of ["zh", "en"]) {
    const scope = createScriptContext({ language });
    commitRows(scope, ["Region"], ["(A)", "Z", "", "(A)", "Z", ""].map((Region) => ({ Region })));
    orders.push(JSON.parse(scope.evaluate('JSON.stringify(getCategoryFilterOptions("Region").map((option) => [option.value, option.label, option.count]))')));
  }
  assert.deepEqual(orders[0], [["", "（空值）", 2], ["(A)", "(A)", 2], ["Z", "Z", 2]]);
  assert.deepEqual(orders[1], [["", "(Blank)", 2], ["(A)", "(A)", 2], ["Z", "Z", 2]]);
});

test("v3.2.0 app.js: the import message is kept as figures and written in the language it is shown in", async () => {
  const scope = await importCsvInLanguage(englishScenarioCsv(), "orders.csv", "zh");
  const status = JSON.parse(scope.evaluate("JSON.stringify(state.importStatus)"));
  assert.deepEqual(Object.keys(status).sort(), ["affectedRows", "formulaErrors", "headerDetection", "sourceDisplayName", "suspectedTotals", "warningCount"]);
  assert.equal(status.warningCount, 1);
  assert.equal(status.affectedRows, 1);
  const chinese = textOf(scope, "statusPanel");
  assert.equal(chinese, "CSV 已载入，但读取时有 1 处疑问，涉及 1 行。请核对数据预览和字段数量。第 20 行「Total」疑似合计行，可以在“数据质量”里设为不计入分析。");
  scope.evaluate('setInterfaceLanguage("en"); showImportStatus(state.importStatus)');
  assert.equal(textOf(scope, "statusPanel"), "The CSV is loaded, but there is 1 reading issue in 1 row. Check the data preview and the number of columns. Row 20 “Total” looks like a total row; you can exclude it from analysis in Data quality.");
  assert.equal(textOf(scope, "statusNoteSummary"), "orders.csv loaded, 1 note");
  // Back from 换一个文件: written again, now in Chinese, byte for byte as after the import.
  scope.evaluate('setInterfaceLanguage("zh"); showUploadPage(); returnToAnalysis()');
  assert.equal(textOf(scope, "statusPanel"), chinese);
});

test("v3.2.0 app.js: one-off messages are written again in the other language, and back byte for byte", async () => {
  const scope = await importCsvInLanguage(englishScenarioCsv(), "orders.csv", "zh");
  scope.evaluate('confirmPlaceholder("Amount", "N/A")');
  scope.evaluate(`handleExclusionClick({ target: { closest: () => ({ dataset: { exclusionKind: "duplicateRows", exclusionOn: "true" } }) } })`);
  scope.evaluate('stageFieldTypeChange("Rep", "ignore")');
  scope.evaluate("applyFilters()");
  scope.evaluate('state.problemView = { category: { kind: "missing", field: "Nope" }, page: 1, list: { items: [] } }; refreshProblemView()');
  const ids = ["placeholderMessage", "duplicateExclusionMessage", "fieldConfigMessage", "filterMessage", "problemRowsMessage"];
  const chinese = ids.map((id) => textOf(scope, id));
  chinese.forEach((text, index) => assert.match(text, /\p{Script=Han}/u, ids[index]));
  scope.evaluate('setInterfaceLanguage("en"); rewriteOneOffMessages()');
  const english = ids.map((id) => textOf(scope, id));
  english.forEach((text, index) => assert.ok(text && !ENGLISH_BAD.test(text), `${ids[index]}: ${text}`));
  assert.equal(english[4], "This category has no problem rows any more.");
  assert.equal(english[2], "Column settings have changes that are not applied. Click “Apply column settings” first, then apply the filters.");
  scope.evaluate('setInterfaceLanguage("zh"); rewriteOneOffMessages()');
  assert.deepEqual(ids.map((id) => textOf(scope, id)), chinese);
  // An empty message forgets the element: nothing is written back into it.
  scope.evaluate('setProblemRowsMessage(""); rewriteOneOffMessages()');
  assert.equal(textOf(scope, "problemRowsMessage"), "");
  assert.equal(scope.evaluate("oneOffMessages.has(dom.problemRowsMessage)"), false);
});

test("v3.2.0 app.js: the rebuild message is built at the click in both languages; a later switch shows the counts of the click", async () => {
  const scope = await importCsvInLanguage(englishScenarioCsv(), "orders.csv", "zh");
  scope.evaluate('confirmPlaceholder("Amount", "N/A")');
  const chinese = textOf(scope, "placeholderMessage");
  // The short row has no Amount: one blank before. 19 rows: 16, the repeated one, the short one and the total.
  assert.match(chinese, /^已把「Amount」的 1 个「N\/A」按空值处理：空值 1 → 2，认不出 1 → 0。.*行数还是 19。/);
  // Afterwards a row is left out (the analysis has 18 rows), and only then the language switches:
  // the English still has the counts of the click.
  scope.evaluate('setRowExclusion("duplicateRows", true)');
  assert.equal(scope.evaluate("getIncludedRows().length"), 18);
  scope.evaluate('setInterfaceLanguage("en"); rewriteOneOffMessages()');
  assert.match(textOf(scope, "placeholderMessage"), /^Treated 1 “N\/A” in “Amount” as blank: blanks 1 → 2, not recognized 1 → 0\..*there are still 19 rows\./);
  scope.evaluate('setInterfaceLanguage("zh"); rewriteOneOffMessages()');
  assert.equal(textOf(scope, "placeholderMessage"), chinese);
  // The language is the one shown before the message was built.
  assert.equal(scope.evaluate("currentLanguage"), "zh");
});

test("v3.2.0 app.js: the (excluded) suffix of a problem row is no L() call, one per listed row", async () => {
  const scope = await importCsvInLanguage(englishScenarioCsv(), "orders.csv", "en");
  scope.evaluate('setRowExclusion("suspectedTotal", true)');
  scope.evaluate("globalThis.__lCalls = 0; globalThis.__realL = L; L = (...args) => { globalThis.__lCalls += 1; return globalThis.__realL(...args); }");
  const notes = JSON.parse(scope.evaluate("JSON.stringify(Array.from({ length: 50 }, () => markExcludedNote({ index: state.excludedRows.suspectedTotal[0], note: \"x\" }, getExcludedIndexSet())))"));
  assert.equal(scope.evaluate("__lCalls"), 0);
  scope.evaluate("L = globalThis.__realL");
  assert.equal(notes[0], "x (excluded from analysis)");
});

test("v3.2.0 format.js English wording: values (not categories), treated as blank, counts with their unit, date formats as written", () => {
  const scope = createScriptContext();
  scope.evaluate('setInterfaceLanguage("en")');
  const en = (expression) => scope.evaluate(expression);
  assert.equal(en("Lpick(CATEGORY_STATS_EMPTY_NOTE)"), "The rows in analysis have no non-blank values.");
  assert.equal(en("Lpick(CATEGORY_CHART_EMPTY_NOTE)"), "The rows in analysis have no values to count");
  assert.equal(en('formatPlaceholderCounts("-", { missing: 0, failures: 0, typeKey: "category", listed: false }, { missing: 1, failures: 0, typeKey: "category", listed: false }, true)'), "blanks 0 → 1, the value counts no longer have “-”");
  assert.equal(en('formatPlaceholderCounts("-", { missing: 1, failures: 0, typeKey: "category", listed: false }, { missing: 0, failures: 0, typeKey: "category", listed: false }, false)'), "blanks 1 → 0, the value counts have “-” again");
  assert.match(en("Lpick(LOOK_ALIKE_NOTE)"), /so value counts, filters and duplicate IDs count them separately\..*so value counts and filters show an extra value you cannot see\.$/);
  // No English sentence of format.js says category or categories any more, nor treated as empty.
  const english = [...repoFile("format.js").matchAll(/Lp?\(\s*(?:"[^"]*"|`[^`]*`)\s*,\s*("[^"]*"|`[^`]*`)/g)].map((match) => match[1]).join("\n");
  assert.ok(english.length > 5000);
  // (category.key and category.count are the variable, not a word.)
  assert.doesNotMatch(english, /categor(?!y\.)|treated as empty|white space/i);
  assert.match(en('formatPlaceholderConfirmMessage({ field: "Amount", key: "N/A", count: 2, before: { missing: 0, failures: 2, typeKey: "numeric", listed: false }, after: { missing: 2, failures: 0, typeKey: "numeric", listed: false }, rowCount: 42 })'), /Only these 2 cells are treated as blank;/);
  assert.match(en('formatAllPlaceholdersMessage([{ field: "Amount", key: "N/A" }], [], 42, false)'), /Only these cells are treated as blank and no rows are deleted/);
  // Every category label counts in its unit.
  assert.deepEqual(["missing", "conversion", "idDuplicate", "parseWarnings"].map((kind) => en(`formatProblemCategoryLabel({ kind: "${kind}", field: "Amount", count: 3 })`)),
    ["Blanks · Amount (3 rows)", "Numbers or dates not recognized · Amount (3 values)", "Duplicate IDs · Amount (3 duplicate values)", "Reading issues (3 issues)"]);
  assert.equal(en('formatProblemCategoryLabel({ kind: "parseWarnings", count: 1 })'), "Reading issues (1 issue)");
  assert.deepEqual(JSON.parse(en("JSON.stringify(Object.fromEntries(Object.entries(DATE_STYLE_NAMES).map(([key, pair]) => [key, pair[1]])))")), {
    excelCell: "Excel date cell", ymdDash: "text YYYY-MM-DD", ymdSlash: "text YYYY/MM/DD", ymdDot: "text YYYY.MM.DD", chinese: "text with Chinese year/month/day characters",
    dmy: "text DD/MM/YYYY", mdy: "text MM/DD/YYYY", dayMonthSame: "text with the same day and month (order unclear)", english: "text with English month names", yearOnly: "year only"
  });
  assert.equal(en("Lpick(VISIBLE_TEXT_OTHER)"), "[whitespace]");
  assert.equal(en('I18N_EN["sidebar.status"]'), "Current status");
});

test("v3.2.0 profiling.js: the type names are [中文, English] pairs (Ignored); profile.type stays the Chinese name", () => {
  const scope = createScriptContext({ language: "en" });
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(Object.fromEntries(Object.keys(FIELD_TYPE_LABELS).map((key) => [key, FIELD_TYPE_LABELS[key]])))")), {
    numeric: ["数值", "Number"], category: ["分类", "Text"], date: ["日期", "Date"], id: ["ID", "ID"], ignore: ["忽略", "Ignored"]
  });
  assert.ok(scope.evaluate("Object.isFrozen(FIELD_TYPE_LABELS)"));
  assert.deepEqual(["numeric", "category", "date", "id", "ignore"].map((key) => scope.evaluate(`fieldTypeLabel("${key}")`)), ["Number", "Text", "Date", "ID", "Ignored"]);
  // The profile is data (the v2.8.1 digest holds it): the Chinese name in both languages.
  commitRows(scope, ["Amount", "Region"], [{ Amount: "1", Region: "East" }, { Amount: "2", Region: "West" }]);
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(state.profiles.map((profile) => profile.type))")), ["数值", "分类"]);
  scope.evaluate('setInterfaceLanguage("zh")');
  assert.deepEqual(["numeric", "ignore"].map((key) => scope.evaluate(`fieldTypeLabel("${key}")`)), ["数值", "忽略"]);
});

test("v3.2.0 stats.js: the group of blanks keeps its Chinese key (R3) and is shown as (Blank) in English", () => {
  const groups = {};
  for (const language of ["zh", "en"]) {
    const scope = createScriptContext({ language });
    commitRows(scope, ["Amount", "Region"], [{ Amount: "1", Region: "" }, { Amount: "2", Region: "East" }, { Amount: "4", Region: "" }]);
    groups[language] = JSON.parse(scope.evaluate('JSON.stringify(groupNumericFieldByCategory("Amount", "Region").map((group) => [group.name, displayGroupName(group.name)]))'));
    assert.equal(scope.evaluate("BLANK_GROUP_KEY"), "（空值）");
  }
  assert.deepEqual(groups.zh.map(([key]) => key), groups.en.map(([key]) => key));
  assert.ok(groups.en.some(([key, label]) => key === "（空值）" && label === "(Blank)"), JSON.stringify(groups.en));
  assert.ok(groups.zh.some(([key, label]) => key === "（空值）" && label === "（空值）"), JSON.stringify(groups.zh));
  // A value that is written （空值） in the data is the same group in both languages (as before).
  const scope = createScriptContext({ language: "en" });
  assert.equal(scope.evaluate('displayGroupName("East")'), "East");
});

test("v3.2.0 importer.js: a formula error with no known writing is #错误 in both languages (it is the cell's text)", () => {
  for (const language of ["zh", "en"]) {
    const scope = createScriptContext({ language });
    assert.equal(scope.evaluate("FORMULA_ERROR_FALLBACK"), "#错误");
  }
});

test("v3.2.0 importer.js: reading issues are kept as code and figures and written in the language shown", async () => {
  // A short row (a row with more cells stops the import).
  const text = "a,b,c\n1,2,3\n4,5\n7,8,9\n10\n";
  const scope = await importCsvInLanguage(text, "short.csv", "zh");
  const stored = JSON.parse(scope.evaluate("JSON.stringify(state.parseWarnings)"));
  assert.ok(stored.length >= 1);
  for (const warning of stored) {
    assert.equal("message" in warning, false, JSON.stringify(warning));
    assert.deepEqual(Object.keys(warning).sort(), ["code", "params", "row", "sourceMessage"]);
  }
  assert.deepEqual(stored.map((warning) => [warning.row, warning.code, warning.params]), [
    [3, "TooFewFields", { line: 3, cellCount: 2, headerCount: 3, separator: "" }],
    [5, "TooFewFields", { line: 5, cellCount: 1, headerCount: 3, separator: "" }]
  ]);
  const chinese = JSON.parse(scope.evaluate("JSON.stringify(state.parseWarnings.map(formatParseWarning))"));
  // The text of v3.1.0, from the figures.
  assert.deepEqual(chinese, ["第 3 行只有 2 列（表头有 3 列），缺少的单元格按空值处理。", "第 5 行只有 1 列（表头有 3 列），缺少的单元格按空值处理。"]);
  scope.evaluate('setInterfaceLanguage("en")');
  const english = JSON.parse(scope.evaluate("JSON.stringify(state.parseWarnings.map(formatParseWarning))"));
  assert.deepEqual(english, ["Row 3 has only 2 columns (the header has 3); the missing cells are treated as blank.", "Row 5 has only 1 column (the header has 3); the missing cells are treated as blank."]);
  english.forEach((message, index) => {
    assert.ok(!ENGLISH_BAD.test(message), message);
    assert.deepEqual(numbersOf(message), numbersOf(chinese[index]), `${message} / ${chinese[index]}`);
  });
  scope.evaluate('setInterfaceLanguage("zh")');
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(state.parseWarnings.map(formatParseWarning))")), chinese);
  // An old message (a string) is shown as it is.
  assert.equal(scope.evaluate('formatParseWarning({ row: 3, code: "X", message: "原样" })'), "原样");
});

test("v3.2.0 stats.js: the trend conclusion is figures (computeTrendInsight) and a sentence (formatTrendInsight)", () => {
  const scope = createScriptContext({ language: "zh" });
  const sentence = (data, language) => {
    scope.evaluate(`setInterfaceLanguage("${language}")`);
    put(scope.context, "__trend", data);
    return scope.evaluate("formatTrendInsight(__trend)");
  };
  const partial = { yearly: false, year: 2025, label: "2025-06", month: 6, day: 15 };
  const cases = [
    [{ dateField: "Date", kind: "none" }, "日期字段「Date」里没有认得出的日期，因此未生成趋势结论。", "The date column “Date” has no recognized dates, so there is no trend summary."],
    [{ dateField: "Date", kind: "few" }, "日期字段「Date」可用于趋势分析，但有效时间点不足以判断变化方向。", "The date column “Date” can show a trend, but it has too few points in time to tell the direction."],
    [{ dateField: "Date", kind: "flat", yearly: false, from: "2025-01", to: "2025-06", ratio: 0, partial: null }, "基于日期字段「Date」观察，记录数趋势 从 2025-01 到 2025-06 基本持平。", "By the date column “Date”, the record count stays about the same from 2025-01 to 2025-06."],
    [{ dateField: "Date", kind: "up", yearly: false, from: "2025-01", to: "2025-06", ratio: 0.25, partial },
      "基于日期字段「Date」观察，记录数趋势 从 2025-01 到 2025-06 呈上升变化，变化幅度约 25.0%。注意：最后一期（2025-06）的数据截至 6 月 15 日，可能不完整，比较结果仅供参考。",
      "By the date column “Date”, the record count rises from 2025-01 to 2025-06, by about 25.0%. Note: the last period (2025-06) only has data up to 2025-06-15, so it may be incomplete and the comparison is only a guide."],
    [{ dateField: "Year", kind: "down", yearly: true, from: "2023", to: "2025", ratio: 0.5, partial: { yearly: true, year: 2025, label: "2025", month: 3, day: 1 } },
      "基于日期字段「Year」观察，年度记录数趋势 从 2023 到 2025 呈下降变化，变化幅度约 50.0%。注意：最后一期（2025 年）的数据截至 3 月 1 日，可能不完整，比较结果仅供参考。",
      "By the date column “Year”, the yearly record count falls from 2023 to 2025, by about 50.0%. Note: the last period (2025) only has data up to 2025-03-01, so it may be incomplete and the comparison is only a guide."]
  ];
  for (const [data, zh, en] of cases) {
    assert.equal(sentence(data, "zh"), zh);
    assert.equal(sentence(data, "en"), en);
    assert.deepEqual(numbersOf(en), numbersOf(zh), en);
  }
  // The exception of numbersOf() is that sentence only, and still tells a wrong day or month.
  assert.deepEqual(numbersOf("only has data up to 2025-06-15, so"), ["15", "6"]);
  assert.deepEqual(numbersOf("from 2025-06-15 to 2025-07-01."), ["2025-06-15", "2025-07-01"].flatMap((date) => date.split("-")).sort());
  assert.notDeepEqual(numbersOf(cases[3][2].replace("2025-06-15", "2025-06-16")), numbersOf(cases[3][1]));
  // (Month and day swapped have the same numbers; the whole sentences above catch that.)
  assert.notDeepEqual(numbersOf(cases[4][2].replace("2025-03-01", "2025-03-02")), numbersOf(cases[4][1]));
});

test("v3.2.0 app.js: the trend figures are kept for a switch and cleared with the analysis", async () => {
  const scope = await importCsvInLanguage(englishScenarioCsv(), "orders.csv", "zh");
  const data = JSON.parse(scope.evaluate("JSON.stringify(state.trendInsightData)"));
  assert.ok(data && ["flat", "up", "down"].includes(data.kind), JSON.stringify(data));
  const chinese = scope.evaluate("state.chartInsights.trend");
  assert.equal(chinese, scope.evaluate("buildTrendInsight(dom.dateChartField.value)"));
  // The figures are not computed again: the sentence from the kept figures in English keeps the
  // numbers. (Commit 11: a switch draws the page again from the rows, so it is the trend's
  // computation that must not run, not every read of the rows.)
  scope.evaluate("globalThis.__realTrend = computeTrendInsight; computeTrendInsight = () => { throw new Error(\"computed\"); }");
  scope.evaluate('setInterfaceLanguage("en")');
  const english = scope.evaluate("formatTrendInsight(state.trendInsightData)");
  assert.equal(scope.evaluate("state.chartInsights.trend"), english);
  scope.evaluate("computeTrendInsight = globalThis.__realTrend");
  assert.ok(!ENGLISH_BAD.test(english), english);
  assert.deepEqual(numbersOf(english), numbersOf(chinese), `${english} / ${chinese}`);
  scope.evaluate("resetAnalysisState()");
  assert.equal(scope.evaluate("state.trendInsightData"), null);
});

test("v3.2.0 quality.js: problem-row notes in English, picked once per list (no L() call per row)", async () => {
  const zh = await importCsvInLanguage(englishScenarioCsv(), "orders.csv", "zh");
  const en = await importCsvInLanguage(englishScenarioCsv(), "orders.csv", "en");
  const categories = JSON.parse(zh.evaluate("JSON.stringify(state.problemCategories)"));
  assert.ok(categories.length >= 8);
  const notesOf = (scope) => categories.map((category) => {
    put(scope.context, "__category", category);
    return JSON.parse(scope.evaluate("JSON.stringify(listProblemRows(__category).items.map((item) => item.note))"));
  });
  const zhNotes = notesOf(zh);
  en.evaluate("globalThis.__lCalls = 0; globalThis.__realL = L; L = (...args) => { globalThis.__lCalls += 1; return globalThis.__realL(...args); }");
  const enNotes = notesOf(en);
  const lCalls = en.evaluate("__lCalls");
  en.evaluate("L = globalThis.__realL");
  // Some lists format a bound or a number (formatFieldNumber may call L once a list); never once a row.
  const rows = enNotes.flat().length;
  assert.ok(rows > 20, `${rows} rows`);
  assert.ok(lCalls < categories.length, `${lCalls} L() calls for ${rows} rows`);
  // Many rows of one kind: a list of 12 blanks, 4 values not recognized and 7 rows of one repeat group
  // makes no L() call at all.
  const many = createScriptContext({ language: "en" });
  commitRows(many, ["Amount", "Note"], Array.from({ length: 36 }, (_, index) => ({ Amount: index < 12 ? "" : index < 16 ? `x${index}` : String(index), Note: index >= 30 ? "same" : `n${index}` })).concat(Array.from({ length: 6 }, () => ({ Amount: "35", Note: "same" }))));
  many.evaluate("globalThis.__lCalls = 0; globalThis.__realL = L; L = (...args) => { globalThis.__lCalls += 1; return globalThis.__realL(...args); }");
  const manyNotes = JSON.parse(many.evaluate('JSON.stringify([listMissingRows("Amount"), listConversionFailureRows("Amount"), listDuplicateRowGroups()].map((list) => list ? list.items.map((item) => item.note) : null))'));
  const manyCalls = many.evaluate("__lCalls");
  many.evaluate("L = globalThis.__realL");
  assert.deepEqual(manyNotes.map((notes) => notes && notes.length), [12, 4, 7]);
  assert.deepEqual([manyNotes[0][0], manyNotes[1][0], manyNotes[2][0]], ["Blank", "Not recognized as a number", "Group 1 (kept)"]);
  assert.equal(manyCalls, 0);
  enNotes.forEach((notes, index) => {
    assert.equal(notes.length, zhNotes[index].length);
    notes.forEach((note, position) => {
      assert.ok(!/\p{Script=Han}/u.test(note.replace(/「[^」]*」|“[^”]*”/g, "")), note);
      assert.deepEqual(numbersOf(note), numbersOf(zhNotes[index][position]), `${note} / ${zhNotes[index][position]}`);
    });
  });
  const all = new Set(enNotes.flat());
  for (const note of ["Blank", "Group 1 (kept)", "Group 1"]) assert.ok(all.has(note), `${note} in ${[...all].slice(0, 20).join(" | ")}`);
  assert.ok([...all].some((note) => /^Above the upper bound /.test(note)), [...all].join(" | "));
});

test("v3.2.0 script.js: the self-check notice and the sample messages have English", () => {
  const source = repoFile("script.js");
  assert.match(source, /"The page was updated\. Refresh it to continue\."/);
  assert.match(source, /"Refresh"/);
  // The notice reads the mark before removing it (format.js may be the missing file).
  assert.ok(source.indexOf('hasAttribute?.("data-lang-pending")') < source.indexOf('removeAttribute("data-lang-pending")'));
});

test("v3.2.0 quality.js English wording: a writing like a blank that is not treated as blank yet, and one that is", async () => {
  const notesOf = (scope) => JSON.parse(scope.evaluate('JSON.stringify([...new Set(listPlaceholderRows("Amount", "N/A").items.map((item) => item.note))])'));
  const en = await importCsvInLanguage(englishScenarioCsv(), "orders.csv", "en");
  const zh = await importCsvInLanguage(englishScenarioCsv(), "orders.csv", "zh");
  assert.deepEqual(notesOf(en), ["Not treated as blank yet"]);
  assert.deepEqual(notesOf(zh), ["待确认"]);
  en.evaluate('confirmPlaceholder("Amount", "N/A")');
  zh.evaluate('confirmPlaceholder("Amount", "N/A")');
  assert.deepEqual(notesOf(en), ["Treated as blank"]);
  assert.deepEqual(notesOf(zh), ["已按空值处理"]);
});

// ---- v3.2.0 commit 8: the report, file names and CSV columns in English ----
// The scenario's steps that reach every part of the report: values like blanks pending, the
// repeated row left out, a filter, a breakdown with a date column, a standardized check.
async function reportScenario(language, csv = englishScenarioCsv(), fileName = "orders.csv") {
  const scope = await importCsvInLanguage(csv, fileName, language);
  const fields = JSON.parse(scope.evaluate("JSON.stringify(state.fields)"));
  if (fields.includes("Region")) {
    scope.evaluate('setRowExclusion("duplicateRows", true)');
    scope.evaluate(`(() => {
      dom.categoryFilterField.value = "Region"; dom.numericFilterField.value = ""; dom.dateFilterField.value = "";
      state.categoryFilterDraftValues = new Set(getCategoryFilterOptions("Region").map((option) => option.value).slice(0, 3));
      applyFilters();
    })()`);
    scope.evaluate('dom.v2MetricField.value = "Amount"; dom.v2GroupField.value = "Region"; dom.v2AggregateMethod.value = "sum"; dom.v2DateField.value = "Date"; renderV2Analysis()');
  }
  scope.evaluate("checkStandardizedExport()");
  // The analysis time is the clock of the import: the same for both languages, or the two reports
  // differ by a second now and then (commit 11 found the number comparison flaky).
  scope.evaluate('state.analysisCompletedAt = new Date("2026-10-25T08:00:00.000Z")');
  return scope;
}

// The report's text a reader sees: no style block, no tags, entities as characters.
const reportVisibleText = (html) => html.replace(/<style>[\s\S]*?<\/style>/, " ").replace(/<[^>]+>/g, " ").replace(/&nbsp;/g, " ").replace(/&amp;/g, "&");

test("v3.2.0 report.js: an English report has <html lang=\"en\">, no Chinese and every number of the Chinese report", async () => {
  const zh = reportOf(await reportScenario("zh"));
  const en = reportOf(await reportScenario("en"));
  assert.match(zh.html, /^<!DOCTYPE html>\n<html lang="zh-CN">/);
  assert.match(en.html, /^<!DOCTYPE html>\n<html lang="en">/);
  const text = reportVisibleText(en.html);
  assert.ok(!ENGLISH_BAD.test(text), text.slice(Math.max(0, text.search(ENGLISH_BAD) - 40), text.search(ENGLISH_BAD) + 40));
  assert.deepEqual(numbersOf(text), numbersOf(reportVisibleText(zh.html)));
  // Every part was there: the filter, the row left out, the breakdown, the pending blank-like values.
  for (const heading of ["Filters", "Column types", "Data quality", "Values that may mean blank", "Breakdown by group", "Date trends", "Automatic conclusions"]) {
    assert.ok(text.includes(heading), heading);
  }
  // The kept values after a colon, each in curly quotes; the blank one is (Blank) without quotes.
  assert.ok(text.includes("Region: keeps 3 of 6 values: (Blank), “(A)”, “east”"), text.slice(text.indexOf("Region: keeps"), text.indexOf("Region: keeps") + 80));
  assert.deepEqual(en.report.filters[0].keptValues, ["(Blank)", "(A)", "east"]);
  assert.match(zh.report.filters[0].description, /^Region：保留 3 \/ 共 6 个类别（（空值）、\(A\)、east）$/);
  assert.match(text, /Metric column:\s+Amount\s+Group by:\s+Region\s+Aggregation:\s+sum\s+Date column:\s+Date/);
  // The two trend sentences, each with its direction.
  assert.match(text, /The date column “Date” has \d+ points in time; the record count goes from \d+ in \d{4}-\d{2}(?:-\d{2})? to \d+ in \d{4}-\d{2}(?:-\d{2})?, (?:rising|falling|about the same) overall\./);
  assert.match(text, /By “Date”, the sum of “Amount” goes from [\d,.]+ in \d{4}-\d{2}(?:-\d{2})? to [\d,.]+ in \d{4}-\d{2}(?:-\d{2})?, (?:rising|falling|about the same) overall\./);
  assert.match(en.html, /<title>orders\.csv - Sheet Checkup report<\/title>/);
  // The version is the report's own (pinned by the version test), so a release does not touch this.
  const reportVersion = en.report.version.replace(/\./g, "\\.");
  assert.match(reportVersion, /^v\d+\\\.\d+\\\.\d+$/);
  assert.match(en.html, new RegExp(`<footer>Created by <a href="https://github\\.com/wy-data-30/sheet-checkup">Sheet Checkup</a> ${reportVersion} in your browser; no data was uploaded\\.</footer>`));
  // The type names of the snapshot are those of the export's language; in Chinese, profile.type.
  assert.deepEqual(en.report.fieldTypes.map((field) => field.appliedType), ["ID", "Text", "Number", "Date", "Text"]);
  assert.deepEqual(zh.report.fieldTypes.map((field) => [field.inferredType, field.appliedType]), JSON.parse((await reportScenario("zh")).evaluate("JSON.stringify(state.profiles.map((profile) => [profile.inferredType, profile.type]))")));
});

test("v3.2.0 report.js: the sample orders (Chinese data) give an English report with every number of the Chinese one", async () => {
  const csv = readExampleCsv("sample-orders.csv");
  const zh = reportOf(await reportScenario("zh", csv, "sample-orders.csv"));
  const en = reportOf(await reportScenario("en", csv, "sample-orders.csv"));
  assert.deepEqual(numbersOf(reportVisibleText(en.html)), numbersOf(reportVisibleText(zh.html)));
  // Only the data and the sample's own Chinese name pieces are Chinese in the English report.
  const dataPieces = new Set((csv.match(/\p{Script=Han}+/gu) || []));
  const leftover = (reportVisibleText(en.html).match(/\p{Script=Han}+/gu) || []).filter((piece) => ![...dataPieces].some((data) => data.includes(piece)));
  assert.deepEqual([...new Set(leftover)], []);
});

test("v3.2.0 report.js: English file names use the label words of the plan; the Chinese ones are as before", async () => {
  const day = 'new Date("2026-10-25T12:00:00")';
  const names = {};
  for (const language of ["zh", "en"]) {
    const scope = await reportScenario(language);
    const categories = JSON.parse(scope.evaluate("JSON.stringify(state.problemCategories)"));
    names[language] = {
      data: scope.evaluate(`buildDataExportFileName("filtered", ${day})`),
      standardized: scope.evaluate(`buildStandardizedFileName("data", ${day})`),
      changes: scope.evaluate(`buildStandardizedFileName("changes", ${day})`),
      first: scope.evaluate(`buildStandardizedFileName("changes", ${day}, 2)`),
      report: scope.evaluate(`buildReportFileName("html", ${day})`),
      problems: categories.map((category) => { put(scope.context, "__nameCategory", category); return [category.kind, scope.evaluate(`buildProblemRowsFileName(__nameCategory, ${day})`)]; })
    };
  }
  assert.deepEqual([names.en.data, names.en.standardized, names.en.changes, names.en.first, names.en.report],
    ["orders_in-analysis_20261025.csv", "orders_standardized_20261025.csv", "orders_standardized-changes_20261025.csv", "orders_standardized-changes_first-2_20261025.csv", names.zh.report]);
  assert.deepEqual([names.zh.data, names.zh.standardized, names.zh.changes, names.zh.first],
    ["orders_当前分析_20261025.csv", "orders_写法统一_20261025.csv", "orders_写法统一改动清单_20261025.csv", "orders_写法统一改动清单_前2条_20261025.csv"]);
  const words = { missing: "blank", conversion: "unrecognized", outlier: "high-or-low", duplicateRows: "duplicate-rows", idDuplicate: "duplicate-ids", placeholder: "blank-like", parseWarnings: "reading-issues", suspectedTotal: "possible-totals", formulaError: "formula-errors", dateStyle: "mixed-date-formats", lookAlike: "variants" };
  names.en.problems.forEach(([kind, name], index) => {
    assert.ok(name.startsWith(`orders_problem-rows_${words[kind]}`) && name.endsWith("_20261025.csv"), name);
    // The same name with the Chinese labels in their place.
    const zhName = names.zh.problems[index][1];
    assert.equal(name.replace(`_problem-rows_${words[kind]}`, ""), zhName.replace(/_问题行_[^_]+/, ""), `${name} / ${zhName}`);
  });
  // The eleven words, each a pair with the Chinese of v3.1.0.
  const scope = createScriptContext({ language: "en" });
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(Object.fromEntries(Object.entries(PROBLEM_KIND_FILE_NAMES).map(([kind, pair]) => [kind, pair[1]])))")), words);
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(Object.values(PROBLEM_KIND_FILE_NAMES).map((pair) => pair[0]))")),
    ["空值", "认不出", "偏大或偏小", "整行重复", "编号重复", "像是空值的写法", "读取有疑问", "疑似合计", "公式错误", "日期写法", "写法不一样"]);
});

test("v3.2.0 report.js: long English names keep their label words (R1)", async () => {
  const stem = "a".repeat(90);
  const scope = await reportScenario("en", englishScenarioCsv(), `${stem}.csv`);
  const category = JSON.parse(scope.evaluate("JSON.stringify(state.problemCategories.find((item) => item.kind === \"missing\"))"));
  put(scope.context, "__longCategory", category);
  const day = 'new Date("2026-10-25T12:00:00")';
  const problem = scope.evaluate(`buildProblemRowsFileName(__longCategory, ${day})`);
  const first = scope.evaluate(`buildStandardizedFileName("changes", ${day}, 1048575)`);
  assert.match(problem, new RegExp(`^a+_problem-rows_blank_${category.field}_20261025\\.csv$`));
  assert.match(first, /^a+_standardized-changes_first-1048575_20261025\.csv$/);
  for (const name of [problem, first]) assert.equal(name.replace(/_20261025\.csv$/, "").length, 80, name);
});

test("v3.2.0 report.js: the CSV columns the exports add are in the language of the export", () => {
  const scope = createScriptContext({ language: "zh" });
  const headers = (fields) => { put(scope.context, "__headerFields", fields); return JSON.parse(scope.evaluate("JSON.stringify(buildProblemRowsCsvHeaders(__headerFields))")); };
  assert.deepEqual(headers(["文件行号", "说明", "x"]), ["文件行号（2）", "说明（2）", "文件行号", "说明", "x"]);
  assert.deepEqual(headers(["File row", "x"]), ["文件行号", "说明", "File row", "x"]);
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(Lpick(STANDARDIZED_CHANGE_HEADERS))")), ["文件行号", "字段", "原值", "导出值", "处理"]);
  scope.evaluate('setInterfaceLanguage("en")');
  assert.deepEqual(headers(["File row", "Note", "File row (2)", "x"]), ["File row (3)", "Note (2)", "File row", "Note", "File row (2)", "x"]);
  assert.deepEqual(headers(["文件行号", "x"]), ["File row", "Note", "文件行号", "x"]);
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(Lpick(STANDARDIZED_CHANGE_HEADERS))")), ["File row", "Column", "Original", "Exported", "Action"]);
});

test("v3.2.0 report.js: the exported CSV files of an English export start with the English columns", async () => {
  const scope = await reportScenario("en");
  const downloads = captureDownloads(scope);
  scope.evaluate('openProblemView(state.problemCategories.find((item) => item.kind === "missing")); exportProblemRows(); exportStandardizedData("changes")');
  const files = await downloads();
  assert.equal(files.length, 2);
  assert.match(files[0].text, /^File row,Note,ID,Region,Amount,Date,Rep\r\n/);
  assert.match(files[1].text, /^File row,Column,Original,Exported,Action\r\n/);
  assert.match(files[0].fileName, /^orders_problem-rows_blank_/);
});

test("v3.2.0 report.js: an old snapshot's sentences come out as written; a missing one is written in the current language", () => {
  const scope = createScriptContext({ language: "en" });
  put(scope.context, "__oldSnapshot", {
    fileName: "old.csv",
    dateTrends: [{ summary: "日期字段「日期」包含 3 个时间点。" }],
    customAnalysis: { completed: false, message: "旧版本写的句子。" },
    insights: ["旧的结论。"]
  });
  const html = scope.evaluate("buildHtmlReport(__oldSnapshot, [])");
  for (const text of ["日期字段「日期」包含 3 个时间点。", "旧版本写的句子。", "旧的结论。"]) assert.ok(html.includes(text), text);
  assert.equal(scope.evaluate("normalizeReportSnapshot({}).customAnalysis.message"), "There is no breakdown by group to export yet.");
  assert.match(scope.evaluate("buildHtmlReport({}, [])"), /<p class="empty">There is no breakdown by group to export yet\.<\/p>/);
  scope.evaluate('setInterfaceLanguage("zh")');
  assert.equal(scope.evaluate("normalizeReportSnapshot({}).customAnalysis.message"), "当前没有可导出的自定义分组分析结果。");
});

test("v3.2.0 report.js: the export status lines are written again in the other language, and back byte for byte", async () => {
  const scope = await reportScenario("zh");
  captureDownloads(scope);
  scope.evaluate('exportProcessedCsv("filtered"); exportStandardizedData("data"); openProblemView(state.problemCategories[0]); exportProblemRows(); exportAnalysisReport()');
  const ids = ["dataExportStatus", "standardizedExportStatus", "problemRowsExportStatus", "reportExportStatus"];
  const chinese = ids.map((id) => textOf(scope, id));
  chinese.forEach((text, index) => assert.match(text, /^已导出|^报告已生成：/, ids[index]));
  scope.evaluate('setInterfaceLanguage("en"); rewriteOneOffMessages()');
  const english = ids.map((id) => textOf(scope, id));
  assert.match(english[0], /^Exported \d+ rows in analysis: orders_当前分析_\d{8}\.csv$/);
  assert.match(english[1], /^Exported the standardized table, 19 rows: orders_写法统一_\d{8}\.csv$/);
  assert.match(english[3], /^The report is ready: orders-\d{8}-\d{6}\.html$/);
  english.forEach((text, index) => assert.deepEqual(numbersOf(text), numbersOf(chinese[index]), text));
  scope.evaluate('setInterfaceLanguage("zh"); rewriteOneOffMessages()');
  assert.deepEqual(ids.map((id) => textOf(scope, id)), chinese);
});

test("v3.2.0 report.js (R10): the report's file name never had the 已载入 prefix since v3.0.0, so its fallback reads the name alone", async () => {
  assert.doesNotMatch(repoFile("report.js").replace(/\/\/.*$/gm, ""), /已载入/);
  const scope = await importCsvInLanguage(englishScenarioCsv(), "orders.csv", "zh");
  assert.equal(scope.evaluate("dom.sourceName.title"), "orders.csv");
  assert.equal(scope.evaluate("getReportSourceFileName()"), "orders.csv");
  scope.evaluate('state.sourceFileName = ""; dom.sourceName.title = "book.xlsx · Sheet1"');
  assert.equal(scope.evaluate("getReportSourceFileName()"), "book.xlsx");
  scope.evaluate('dom.sourceName.title = ""; dom.sourceName.textContent = ""');
  assert.equal(scope.evaluate("getReportSourceFileName()"), "data-report");
});

test("v3.2.0 report.js English wording: a long list of kept values, and the standardized check's name", () => {
  const descriptions = {};
  for (const language of ["zh", "en"]) {
    const scope = createScriptContext({ language });
    const names = Array.from({ length: 12 }, (_, index) => `V${String(index + 1).padStart(2, "0")}`);
    commitRows(scope, ["Kind"], [...names, ""].map((Kind) => ({ Kind })));
    descriptions[language] = scope.evaluate('buildReportFilters({ category: { field: "Kind", values: getCategoryFilterOptions("Kind").map((option) => option.value) } })[0].description');
    if (language === "en") {
      scope.evaluate('exportStandardizedData("data")');
      assert.equal(textOf(scope, "standardizedExportStatus"), "Click “Check which cells will change” first.");
    }
  }
  assert.equal(descriptions.zh, "Kind：保留 13 / 共 13 个类别（（空值）、V01、V02、V03、V04、V05、V06、V07、V08、V09……等 13 个）");
  assert.equal(descriptions.en, "Kind: keeps 13 of 13 values: (Blank), “V01”, “V02”, “V03”, “V04”, “V05”, “V06”, “V07”, “V08”, “V09”, … 13 in all");
});

test("v3.2.0 app.js English wording: the chart's value conclusion and an empty filter search say value", () => {
  const texts = {};
  for (const language of ["zh", "en"]) {
    const scope = createScriptContext({ language });
    commitRows(scope, ["Region"], [..."EEEEEEEEEEEE", ..."W".repeat(20), ..."N".repeat(8)].map((letter) => ({ Region: { E: "East", W: "West", N: "North" }[letter] })));
    const insight = scope.evaluate('buildCategoryInsight("Region")');
    scope.evaluate('dom.categoryFilterField.value = "Region"; updateCategoryFilterField(); dom.categoryFilterSearch.value = "zzz"; renderCategoryFilterValues()');
    texts[language] = [insight, scope.evaluate("dom.categoryFilterValues.innerHTML")];
  }
  assert.deepEqual(texts.zh, ["分类字段「Region」中占比最高的类别是「West」，出现 20 次，占比 50.0%。", '<p class="filter-value-empty">没有匹配的类别。</p>']);
  assert.deepEqual(texts.en, ["In “Region”, the most common value is “West”: 20 times, 50.0%.", '<p class="filter-value-empty">No matching values.</p>']);
  // The problem category's sentence keeps category (it is a category of problem rows).
  assert.match(repoFile("app.js"), /"This category has no problem rows any more\."/);
});

test("v3.2.0 navigation: a second row only when the links overflow, measured without that row; a language switch measures again", () => {
  const scope = createScriptContext();
  // A navigation that overflows in one row (900 > 800) and fits in its own (900 ≤ 1200).
  scope.evaluate(`globalThis.__nav = { get scrollWidth() { return 900; }, get clientWidth() { return document.body.classList.contains("nav-wrap") ? 1200 : globalThis.__navWidth; } };
    document.getElementById = ((original) => (id) => (id === "moduleNav" ? globalThis.__nav : original(id)))(document.getElementById);`);
  const wraps = (width, results) => scope.evaluate(`(() => {
    globalThis.__navWidth = ${width};
    document.body.classList.${results ? "add" : "remove"}("has-results");
    updateNavFit();
    return document.body.classList.contains("nav-wrap");
  })()`);
  assert.equal(wraps(800, true), true);
  // Measured without the class: a navigation that fits in one row loses its second row.
  assert.equal(wraps(950, true), false);
  // One pixel of rounding is allowed (scrollWidth is a whole number).
  assert.equal(wraps(899, true), false);
  assert.equal(wraps(898, true), true);
  // The upload page (no results) never wraps.
  assert.equal(wraps(800, false), false);
  // A language switch measures again: the English words overflow where the Chinese did not.
  wraps(950, true);
  scope.evaluate("globalThis.__navWidth = 800; setInterfaceLanguage(\"en\")");
  assert.equal(scope.evaluate('document.body.classList.contains("nav-wrap")'), true);
  // script.js measures on every size change of the top bar and of the window.
  const script = repoFile("script.js");
  assert.match(script, /window\.addEventListener\("resize", updateNavFit\);/);
  assert.match(script, /const navFitObserver = new ResizeObserver\(updateNavFit\);/);
  assert.match(script, /navFitObserver\.observe\(document\.querySelector\("\.app-header"\)\);/);
  // The navigation itself: a page opened narrow gets its links with the import, the top bar keeps its size.
  assert.match(script, /navFitObserver\.observe\(document\.getElementById\("moduleNav"\)\);/);
  // style.css: the second row, and the fallback scroll of .module-nav stays.
  const css = repoFile("style.css");
  assert.match(css, /body\.has-results\.nav-wrap \.header-inner \{\s*flex-wrap: wrap;/);
  // The second row's track is only as wide as its links (fix after commit 9), never wider than the row.
  assert.match(css, /body\.has-results\.nav-wrap \.module-nav \{\s*flex: 0 1 auto;\s*max-width: 100%;/);
  assert.match(css, /\.module-nav \{[^}]*overflow-x: auto;/);
});

test("v3.2.0 commit 10: the note above the results is a white bar under the heading, sentence left, button right", () => {
  const css = repoFile("style.css");
  const rule = css.match(/\n\.analysis-contains-note \{([^}]*)\}/)[1];
  for (const line of ["display: flex;", "justify-content: space-between;", "border: 1px solid var(--card-border);", "border-radius: var(--radius-card);", "background: #fff;", "font-size: 13px;", "font-weight: 400;"]) {
    assert.ok(rule.includes(line), line);
  }
  const button = css.match(/\n\.analysis-contains-note \.text-button \{([^}]*)\}/)[1];
  assert.ok(button.includes("font-size: 13px;") && button.includes("font-weight: 600;"), button);
  // The heading no longer wraps for it (the rules of v3.1.0 are gone).
  assert.doesNotMatch(css, /\.section-heading:has\(> \.analysis-contains-note/);
  // The text and the button are as before: the bar changes place and frame only.
  const scope = createScriptContext();
  commitRows(scope, ["名称", "金额"], [{ 名称: "甲", 金额: "1" }, { 名称: "乙", 金额: "2" }, { 名称: "合计", 金额: "3" }]);
  const html = scope.evaluate("dom.statisticsContainsNote.innerHTML");
  assert.match(html, /^[^<]+ <button type="button" class="text-button" data-quality-target="suspectedTotal" aria-label="[^"]+">去数据质量处理<\/button>$/);
});

test("v3.2.0 commit 10: the export options say 共 42 行 / 42 rows, the number alone in its <strong>", async () => {
  const words = (language, count) => {
    const scope = createScriptContext({ language });
    return JSON.parse(scope.evaluate(`JSON.stringify(formatExportRowCount(${count}))`));
  };
  assert.deepEqual(words("zh", 42), { before: "共 ", after: " 行" });
  assert.deepEqual(words("zh", 1), { before: "共 ", after: " 行" });
  assert.deepEqual(words("en", 42), { before: "", after: " rows" });
  assert.deepEqual(words("en", 1), { before: "", after: " row" });
  for (const language of ["zh", "en"]) {
    const scope = await importCsvInLanguage(englishScenarioCsv(), "orders.csv", language);
    scope.evaluate('setRowExclusion("duplicateRows", true)');
    const line = (name) => ["Before", "Count", "After"].map((part) => textOf(scope, `dataExport${name}${part}`)).join("");
    assert.deepEqual([line("Filtered"), line("Original")], language === "zh" ? ["共 18 行", "共 19 行"] : ["18 rows", "19 rows"]);
    assert.deepEqual([textOf(scope, "dataExportFilteredCount"), textOf(scope, "dataExportOriginalCount")], ["18", "19"]);
  }
  // The old labels are gone from the page and from the English table.
  const page = indexHtmlSource();
  assert.doesNotMatch(page, /当前分析行数|原始行数 </);
  const english = JSON.parse(createScriptContext().evaluate("JSON.stringify(I18N_EN)"));
  assert.equal("export.analysisCount" in english || "export.originalCount" in english, false);
});

// ---- v3.2.0 commit 11: switch the language in place ----
// The scenario's steps (filter, rows left out, field draft, custom analysis, standardized check,
// a problem category open), ending on a state worth keeping; what a switch must keep.
async function switchScenario(language) {
  const { scope } = await runLanguageScenario(language);
  scope.evaluate('setRowExclusion("duplicateRows", true)');
  // (updateSchemaDraftRow() patches the row through DOM calls the fake page has not; the browser
  // shows what renderSchema() draws, so the scenario draws it.)
  scope.evaluate('stageFieldTypeChange("ID", "category"); renderSchema()');
  scope.evaluate('openProblemView(state.problemCategories.find((category) => category.kind === "missing") || state.problemCategories[0])');
  scope.evaluate('dom.v2MetricField.value = "Amount"; dom.v2GroupField.value = "Region"; dom.v2AggregateMethod.value = "sum"; dom.v2DateField.value = "Date"; renderV2Analysis()');
  return scope;
}

// The two language switches have no text in the fake page until a switch writes it.
const withoutSwitches = (texts) => Object.fromEntries(Object.entries(texts).filter(([id]) => !["languageSwitch", "sidebarLanguageSwitch"].includes(id)));

const KEPT_STATE = `JSON.stringify({
  filters: state.filters, excluded: state.excludedRows, module: document.body.dataset.module ?? null,
  problem: state.problemView && { category: state.problemView.category, page: state.problemView.page },
  previewPage: state.previewPage, drafts: state.fieldTypeDrafts, standardized: Boolean(state.standardizedExport),
  custom: state.customAnalysis && { metricField: state.customAnalysis.metricField, groupField: state.customAnalysis.groupField, grouped: state.customAnalysis.grouped.map((item) => [item.name, item.sum]) },
  selects: [dom.numericChartField, dom.categoryChartField, dom.dateChartField, dom.dateChartMetric, dom.dateChartAggregate, dom.v2MetricField, dom.v2GroupField, dom.v2AggregateMethod, dom.v2DateField, dom.categoryFilterField, dom.numericFilterField, dom.dateFilterField].map((select) => select.value),
  categoryDrafts: [...state.categoryFilterDraftValues]
})`;

test("v3.2.0 commit 11: Chinese → English → Chinese gives every element of the page back byte for byte and keeps the state", async () => {
  await withRecordedAttributes(async () => {
    const scope = await switchScenario("zh");
    const before = withoutSwitches(elementTexts(scope));
    const state = scope.evaluate(KEPT_STATE);
    // A browser's <select> loses its choice when its options are written again; the fake one keeps
    // it, so this does what the browser does (the switch must put every choice back).
    scope.evaluate("globalThis.__realSetSelectOptions = setSelectOptions; setSelectOptions = (select, ...args) => { __realSetSelectOptions(select, ...args); select.value = \"\"; }");
    scope.evaluate('setInterfaceLanguage("en")');
    const english = elementTexts(scope);
    assert.equal(scope.evaluate(KEPT_STATE), state, "the state after the switch to English");
    scope.evaluate('setInterfaceLanguage("zh")');
    const after = withoutSwitches(elementTexts(scope));
    const changed = Object.keys(before).filter((id) => before[id] !== after[id]).map((id) => `${id}: ${String(before[id]).slice(0, 80)} → ${String(after[id]).slice(0, 80)}`);
    assert.deepEqual(changed, []);
    assert.equal(scope.evaluate(KEPT_STATE), state);
    // The English page really was English: no element kept Chinese of the scripts.
    const fragments = doneFileChineseFragments();
    const leaks = Object.entries(english).filter(([, text]) => fragments.some((fragment) => text.includes(fragment))).map(([id, text]) => `${id}: ${text.slice(0, 100)}`);
    assert.deepEqual(leaks, []);
  });
});

test("v3.2.0 commit 11: Chinese switched to English shows what a page opened in English shows (no render function left out)", async () => {
  await withRecordedAttributes(async () => {
    const switched = await switchScenario("zh");
    switched.evaluate('setInterfaceLanguage("en")');
    const native = await switchScenario("en");
    const a = elementTexts(switched);
    const b = elementTexts(native);
    // The status bar's fold and the language announcer come from the switch itself.
    const differ = Object.keys(b).filter((id) => a[id] !== b[id]).map((id) => `${id}: ${String(a[id]).slice(0, 90)} ≠ ${String(b[id]).slice(0, 90)}`);
    assert.deepEqual(differ, []);
  });
});

test("v3.2.0 commit 11: the switch draws again without computing again: no grouping, no trend figures, charts destroyed and drawn", async () => {
  const scope = await switchScenario("zh");
  scope.evaluate(`globalThis.__calls = { grouped: 0, trend: 0, charts: 0, stats: 0 };
    for (const [name, key] of [["buildV2GroupedStats", "grouped"], ["computeTrendInsight", "trend"], ["renderCharts", "charts"], ["buildCategoryStats", "stats"]]) {
      const real = globalThis[name];
      globalThis[name] = (...args) => { __calls[key] += 1; return real(...args); };
    }`);
  scope.evaluate('setInterfaceLanguage("en")');
  const calls = JSON.parse(scope.evaluate("JSON.stringify(__calls)"));
  assert.deepEqual([calls.grouped, calls.trend, calls.stats], [0, 0, 0]);
  // The charts are drawn again (renderCharts() destroys each first, 契约 11; the fake page has no
  // Chart.js, so the count is of the calls).
  assert.equal(calls.charts, 1);
  // The open problem rows are listed again in English, on the same page.
  assert.match(scope.evaluate("state.problemView.list.items[0].note"), /^[A-Z][a-z]/);
});

test("v3.2.0 commit 11: the problem-row viewer keeps its page; the notes are listed again in the new language", () => {
  const scope = createScriptContext();
  commitRows(scope, ["Name", "Amount"], Array.from({ length: 130 }, (_, index) => ({ Name: `n${index}`, Amount: index % 2 ? "" : String(index) })));
  scope.evaluate('openProblemView({ kind: "missing", field: "Amount" }); state.problemView.page = 2; renderProblemView()');
  const notes = () => JSON.parse(scope.evaluate("JSON.stringify([state.problemView.page, state.problemView.list.items[0].note])"));
  assert.deepEqual(notes(), [2, "空值"]);
  scope.evaluate('setInterfaceLanguage("en")');
  assert.deepEqual(notes(), [2, "Blank"]);
  scope.evaluate('setInterfaceLanguage("zh")');
  assert.deepEqual(notes(), [2, "空值"]);
});

test("v3.2.0 commit 11: the status bar, an import error and the worksheet message are written again after a switch", async () => {
  // An error of the importer (a CSV with only a header row): built in both languages.
  const scope = await importCsvInLanguage("a,b\n", "only-header.csv", "zh");
  const chinese = textOf(scope, "statusPanel");
  assert.match(chinese, /\p{Script=Han}/u);
  scope.evaluate('setInterfaceLanguage("en")');
  const english = textOf(scope, "statusPanel");
  assert.ok(english && !ENGLISH_BAD.test(english), english);
  scope.evaluate('setInterfaceLanguage("zh")');
  assert.equal(textOf(scope, "statusPanel"), chinese);
  // A message the report export writes into the status bar.
  const done = await importCsvInLanguage(englishScenarioCsv(), "orders.csv", "zh");
  captureDownloads(done);
  done.evaluate("exportAnalysisReport()");
  assert.match(textOf(done, "statusPanel"), /^报告已生成：orders-\d{8}-\d{6}\.html$/);
  done.evaluate('setInterfaceLanguage("en")');
  assert.match(textOf(done, "statusPanel"), /^The report is ready: orders-\d{8}-\d{6}\.html$/);
  // The worksheet message (a one-off message since commit 11).
  done.evaluate('showSheetSelectionError(() => L("请选择一个工作表后再继续。", "Choose a worksheet to continue."))');
  assert.equal(textOf(done, "sheetSelectionMessage"), "Choose a worksheet to continue.");
  done.evaluate('setInterfaceLanguage("zh")');
  assert.equal(textOf(done, "sheetSelectionMessage"), "请选择一个工作表后再继续。");
  // An error whose message is not ours (PapaParse, the browser) stays as it is.
  assert.equal(done.evaluate('errorMessageBuilder(new Error("plain"))'), "plain");
  assert.equal(done.evaluate('typeof errorMessageBuilder(languageError(Lp("甲", "A")))'), "function");
});

test("v3.2.0 commit 11: a switch says so in a role=status element made at the first switch, written a moment later and taken away", async () => {
  const scope = createScriptContext();
  // A page never switched has no such element (the Chinese page is as before).
  assert.equal(scope.evaluate("languageAnnouncer"), null);
  scope.evaluate("switchInterfaceLanguage()");
  const element = () => JSON.parse(scope.evaluate("JSON.stringify(languageAnnouncer && { text: languageAnnouncer.textContent, className: languageAnnouncer.className })"));
  assert.deepEqual(element(), { text: "", className: "visually-hidden" });
  await new Promise((resolve) => setTimeout(resolve, 80));
  assert.equal(element().text, "Switched to English.");
  scope.evaluate("switchInterfaceLanguage()");
  await new Promise((resolve) => setTimeout(resolve, 80));
  assert.equal(element().text, "界面已切换为中文。");
  assert.equal(scope.evaluate("LANGUAGE_ANNOUNCE_CLEAR_MS"), 5000);
  const app = repoFile("app.js");
  assert.match(app, /languageAnnouncer\.setAttribute\("role", "status"\);/);
  assert.match(app, /document\.getElementById\("mainContent"\)\?\.appendChild\(languageAnnouncer\);/);
  assert.doesNotMatch(repoFile("script.js"), /languageAnnouncer/);
});

// ---- fix after commit 11: the English y-axis titles start with a capital ----
test("v3.2.0 fix: English chart axis titles are capitalized; the words in sentences and the Chinese stay as they were", async () => {
  const axisTitles = async (language) => {
    const scope = await importCsvInLanguage(englishScenarioCsv(), "orders.csv", language);
    const charts = withFakeCharts(scope);
    const titles = [];
    for (const [metric, aggregate] of [["", "count"], ["Amount", "avg"], ["Amount", "sum"]]) {
      charts.length = 0;
      scope.evaluate(`populateControls(); dom.dateChartField.value = "Date"; dom.dateChartMetric.value = ${JSON.stringify(metric)}; dom.dateChartAggregate.value = ${JSON.stringify(aggregate)}; renderDateTrendChart()`);
      titles.push(charts.at(-1).options.scales.y.title.text);
    }
    for (const method of ["sum", "avg", "median"]) {
      charts.length = 0;
      scope.evaluate(`dom.v2MetricField.value = "Amount"; dom.v2GroupField.value = "Region"; dom.v2AggregateMethod.value = ${JSON.stringify(method)}; dom.v2DateField.value = "Date"; renderV2Analysis()`);
      assert.equal(charts.length, 2, method);
      titles.push(charts.map((config) => config.options.scales.y.title.text));
    }
    return { scope, titles };
  };
  const english = await axisTitles("en");
  assert.deepEqual(english.titles, ["Record count", "Average", "Sum", ["Sum", "Sum"], ["Average", "Average"], ["Median", "Median"]]);
  // The sentences keep the lower-case words (the chart titles, the breakdown message).
  assert.deepEqual(JSON.parse(english.scope.evaluate('JSON.stringify([getDateChartAggregateLabel("count"), getDateChartAggregateLabel("sum"), getAggregateMethodLabel("avg"), getAggregateMethodLabel("median")])')),
    ["record count", "sum", "average", "median"]);
  assert.match(english.scope.evaluate("dom.v2TrendTitle.textContent"), /^Amount by date \(median\)$/);
  assert.match(english.scope.evaluate("dom.v2TopGroupTitle.textContent"), /\(median\)$/);
  // A unit or a percent column keeps its words after the capital.
  assert.equal(english.scope.evaluate('numericAxisLabel("Amount", getAggregateMethodAxisLabel("sum"))'), "Sum");
  // Chinese: the same words as before, in the titles and in the sentences.
  const chinese = await axisTitles("zh");
  assert.deepEqual(chinese.titles, ["记录数", "平均值", "求和", ["求和", "求和"], ["平均值", "平均值"], ["中位数", "中位数"]]);
  assert.deepEqual(JSON.parse(chinese.scope.evaluate('JSON.stringify([getDateChartAggregateLabel("count"), getDateChartAggregateLabel("sum"), getAggregateMethodLabel("avg"), getAggregateMethodLabel("median"), getDateChartAxisLabel("avg"), getAggregateMethodAxisLabel("median")])')),
    ["记录数", "求和", "平均值", "中位数", "平均值", "中位数"]);
});

// ---- v3.2.0 commit 12: the English sample and two English fixtures ----
// examples/sample-orders-en.csv: the Chinese sample row by row in English (the same quantities,
// amounts, dates and places of the problems). The English interface's sample button loads it and
// the download link points at it; the Chinese interface is as before.
const SAMPLE_ORDERS_EN = "sample-orders-en.csv";
// Compared after dropping the CR of every CRLF, like SAMPLE_ORDERS_LF_SHA256.
const SAMPLE_ORDERS_EN_LF_SHA256 = "0b79436e00095497f163a5e69d42ea66e07fefc356a203f2035319a786031479";
const sampleOrdersEnLfText = () => fs.readFileSync(path.join(REPO_ROOT, "examples", SAMPLE_ORDERS_EN), "utf8").replace(/\r\n/g, "\n");
const sampleOrdersEnLines = () => sampleOrdersEnLfText().slice(1, -1).split("\n");
const PILLS_ZERO_EN = ["Blanks: 0", "Duplicate rows: 0", "Numbers or dates not recognized: 0 values", "Unusually high or low values: 0 columns", "Duplicate IDs: 0 columns", "Values that may mean blank: 0 columns"];
const SAMPLE_ORDERS_EN_TYPES = ["id", "date", "category", "category", "category", "category", "numeric", "numeric", "category"];

test("v3.2.0 commit 12: the English sample holds the planned rows, title and total line; the total adds the duplicate too", () => {
  const raw = fs.readFileSync(path.join(REPO_ROOT, "examples", SAMPLE_ORDERS_EN));
  assert.deepEqual([...raw.subarray(0, 3)], [0xef, 0xbb, 0xbf]);
  const text = sampleOrdersEnLfText();
  assert.equal(crypto.createHash("sha256").update(text, "utf8").digest("hex"), SAMPLE_ORDERS_EN_LF_SHA256);
  assert.equal(Buffer.byteLength(text, "utf8"), 3306);
  assert.ok(text.endsWith("\n"));
  const lines = sampleOrdersEnLines();
  assert.equal(lines.length, 44);
  // 20 characters: the status line's preview of a skipped row (SKIPPED_ROW_PREVIEW_LENGTH) shows it whole.
  assert.equal(lines[0], "H1 2025 sales orders,,,,,,,,");
  assert.equal(lines[0].replace(/,+$/, "").length, createScriptContext().evaluate("SKIPPED_ROW_PREVIEW_LENGTH"));
  assert.equal(lines[1], "Order ID,Order date,Region,Category,Product,Customer type,Quantity,Amount,Sales rep");
  assert.equal(lines[43], "Total,,,,,,800,123440,");
  assert.equal(lines[19], lines[20]);
  assert.match(lines[19], /^SO-1018,/);
  // The total is every readable quantity and amount above it, the repeated SO-1018 included and the
  // N/A left out (so both columns match; without the duplicate the amount would be 119540).
  const cells = (line) => Papa.parse(line).data[0];
  const data = lines.slice(2, 43).map(cells);
  assert.equal(data.length, 41);
  assert.equal(data.reduce((sum, row) => sum + Number(row[6]), 0), 800);
  const amounts = data.map((row) => Number(row[7].replace(/[$,]/g, ""))).filter(Number.isFinite);
  assert.equal(amounts.length, 40);
  assert.equal(amounts.reduce((sum, value) => sum + value, 0), 123440);
  // The Chinese sample has the same numbers, row by row.
  const chinese = sampleOrdersLines().slice(2, 43).map(cells);
  assert.deepEqual(data.map((row) => row[6]), chinese.map((row) => row[6]));
  assert.deepEqual(data.map((row) => Number(row[7].replace(/[$,]/g, ""))), chinese.map((row) => Number(row[7].replace(/元$/, ""))));
});

test("v3.2.0 commit 12: the sample button and the download link follow the language; the Chinese ones are as before", async () => {
  const script = repoFile("script.js");
  assert.match(script, /fetch\(`examples\/\$\{sampleName\}`\)/);
  const english = await clickSampleButton(new TextDecoder("utf-8").decode(fs.readFileSync(path.join(REPO_ROOT, "examples", SAMPLE_ORDERS_EN))), "en");
  assert.deepEqual(english.fetched, ["examples/sample-orders-en.csv"]);
  assert.deepEqual(english.parsedNames, [SAMPLE_ORDERS_EN, SAMPLE_ORDERS_EN]);
  const uploaded = await importCsvInLanguage(fs.readFileSync(path.join(REPO_ROOT, "examples", SAMPLE_ORDERS_EN), "utf8"), SAMPLE_ORDERS_EN, "en");
  assert.equal(sampleOrdersSnapshot(english), sampleOrdersSnapshot(uploaded));
  const chinese = await clickSampleButton(sampleOrdersFetchedText(), "zh");
  assert.deepEqual(chinese.fetched, ["examples/sample-orders.csv"]);
  assert.deepEqual(chinese.parsedNames, [SAMPLE_ORDERS, SAMPLE_ORDERS]);
  // The loading line names the file being loaded.
  assert.match(script, /setStatus\(L\(`正在加载 \$\{sampleName\} \.\.\.`, `Loading \$\{sampleName\} \.\.\.`\)\);/);
  // The link: the Chinese page is unchanged; the English address is in the English table and is a real
  // file. (The static guard's test requires every English key to be used in index.html, so the link
  // carries upload.downloadHref; the switch below is what the key does.)
  assert.match(indexHtmlSource(),/<a href="examples\/sample-orders\.csv" download>下载示例 CSV<\/a>/);
  const table = JSON.parse(createScriptContext().evaluate("JSON.stringify(I18N_EN)"));
  assert.equal(table["upload.downloadHref"], "examples/sample-orders-en.csv");
  assert.ok(fs.existsSync(path.join(REPO_ROOT, table["upload.downloadHref"])));
  // Switched in place: English, then Chinese again (applyStaticLanguage() over the link alone).
  const scope = createScriptContext();
  const link = i18nElement({ href: "examples/sample-orders.csv", download: "", "data-i18n": "upload.download", "data-i18n-attr": "href=upload.downloadHref" }, ["下载示例 CSV"]);
  put(scope.context, "__i18nElements", [link]);
  scope.evaluate("document.querySelectorAll = (selector) => selector === '[data-i18n], [data-i18n-attr]' ? __i18nElements : []");
  scope.evaluate('rememberStaticLanguage(); setInterfaceLanguage("en")');
  assert.deepEqual([link.getAttribute("href"), link.childNodes[0].nodeValue], ["examples/sample-orders-en.csv", "Download sample CSV"]);
  scope.evaluate('setInterfaceLanguage("zh")');
  assert.deepEqual([link.getAttribute("href"), link.childNodes[0].nodeValue], ["examples/sample-orders.csv", "下载示例 CSV"]);
});

test("v3.2.0 commit 12: the upload page's English rows are rows 10, 12, 20, 21, 29 and 44 of the English sample", () => {
  const html = indexHtmlSource();
  const start = html.indexOf('<section class="sample-preview" aria-labelledby="samplePreviewTitle">');
  assert.ok(start >= 0);
  const section = html.slice(start, html.indexOf("</section>", start));
  const english = JSON.parse(createScriptContext().evaluate("JSON.stringify(I18N_EN)"));
  // The Chinese table's cells (the page as the Chinese interface shows it), with the cells that
  // have words replaced by their English: the six translated cells and their keys, by place.
  const rows = [...section.matchAll(/<tr>(<td[\s\S]*?)<\/tr>/g)].map((match) => [...match[1].matchAll(/<td(?: class="[^"]*")?>([\s\S]*?)<\/td>/g)]
    .map((cell) => cell[1].replace(/<span class="visible-space" aria-hidden="true">␣<\/span>/g, "")));
  const translated = { "0,3": "sample.south", "1,2": "sample.otherDate", "1,3": "sample.southwest", "2,3": "sample.south", "3,3": "sample.south", "4,3": "sample.northwest", "5,1": "sample.total" };
  Object.entries(translated).forEach(([place, key]) => {
    const [row, column] = place.split(",").map(Number);
    rows[row][column] = english[key];
  });
  const lines = sampleOrdersEnLines().map((line) => Papa.parse(line).data[0]);
  const header = lines[1];
  assert.deepEqual(rows.map((cells) => cells[0]), ["10", "12", "20", "21", "29", "44"]);
  rows.forEach((cells) => {
    const values = lines[Number(cells[0]) - 1];
    // The marked trailing space of "South " is drawn as ␣ on the page, not in the words.
    ["Order ID", "Order date", "Region", "Amount"].forEach((field, index) => assert.equal(cells[index + 1], values[header.indexOf(field)].trim(), `${cells[0]} ${field}`));
  });
  assert.equal(lines[9][header.indexOf("Region")], "South ");
  // Every other cell is the same in both samples (order IDs, ISO dates, amounts): no key needed.
  assert.equal(Object.keys(english).filter((key) => /^sample\.(south|otherDate|southwest|northwest|total)$/.test(key)).length, 5);
});

// The English sample without its problems, as in the v2.11.0 test: each fix takes one problem out.
// SO-1026 is 6 headphones at 260 each (1560); every other East order is Emma Chen's.
const SAMPLE_ORDERS_EN_FIXES = {
  title: (lines) => lines.slice(1),
  total: (lines) => lines.filter((line) => !line.startsWith("Total,")),
  duplicate: (lines) => lines.filter((line, index) => !(line.startsWith("SO-1018,") && line === lines[index - 1])),
  na: (lines) => lines.map((line) => (line.startsWith("SO-1026,") ? line.replace(",N/A,", ",1560,") : line)),
  dash: (lines) => lines.map((line) => (line.startsWith("SO-1014,") ? line.replace(/,\u2014$/, ",Emma Chen") : line)),
  money: (lines) => lines.map((line) => line.replace(/,"?\$([\d,]+)\.00"?,/, (match, number) => `,${number.replace(/,/g, "")},`)),
  date: (lines) => lines.map((line) => line
    .replace(/^(SO-\d+),(\d{2})\/(\d{2})\/(\d{4}),/, "$1,$4-$2-$3,")
    .replace(/^(SO-\d+),"([A-Z][a-z]{2}) (\d{1,2}), (\d{4})",/, (match, id, month, day, year) => `${id},${year}-${String(["Jan", "Feb", "Mar", "Apr", "May", "Jun"].indexOf(month) + 1).padStart(2, "0")}-${day.padStart(2, "0")},`)),
  space: (lines) => lines.map((line) => line.replace(/ +,/g, ",").replace(/^(SO-1034,[^,]*),east,/, "$1,East,").replace(/, (Emma Chen)$/, ",$1"))
};

// The total line of the "total only" case adds the rows that are left, like the sample's own total
// (the duplicate and N/A fixes change what it adds up to).
function sampleOrdersEnVariant(keep) {
  let lines = sampleOrdersEnLines();
  for (const [name, fix] of Object.entries(SAMPLE_ORDERS_EN_FIXES)) if (name !== keep) lines = fix(lines);
  if (keep === "total") {
    const data = lines.slice(1, -1).map((line) => line.split(","));
    const sum = (column) => data.reduce((total, row) => total + Number(row[column]), 0);
    lines[lines.length - 1] = `Total,,,,,,${sum(6)},${sum(7)},`;
  }
  return "\uFEFF" + lines.join("\n") + "\n";
}

test("v3.2.0 commit 12: without its problems the English sample flags nothing, and each problem alone flags only itself", async () => {
  const pills = (changes) => PILLS_ZERO_EN.map((pill) => changes[pill.split(":")[0]] ?? pill);
  const cases = {
    clean: { pills: PILLS_ZERO_EN, sections: [], note: false, cells: 0 },
    title: { pills: PILLS_ZERO_EN, sections: [], note: true, cells: 0 },
    total: { pills: pills({ Blanks: "Blanks: 6", "Unusually high or low values": "Unusually high or low values: 2 columns" }), sections: ["suspectedTotalSection"], note: true, cells: 0 },
    duplicate: { pills: pills({ "Duplicate rows": "Duplicate rows: 1", "Duplicate IDs": "Duplicate IDs: 1 column" }), sections: [], note: false, cells: 0 },
    na: { pills: pills({ "Numbers or dates not recognized": "Numbers or dates not recognized: 1 value", "Values that may mean blank": "Values that may mean blank: 1 column" }), sections: [], note: false, cells: 1 },
    dash: { pills: pills({ "Values that may mean blank": "Values that may mean blank: 1 column" }), sections: [], note: false, cells: 0 },
    // $5,400.00, $900.00 and $1,050.00 are plain numbers to the current rules ($ and the thousands
    // separator are read without a unit, unlike 5400元), so the quality page has nothing to say;
    // only the standardized table writes them as 5400.00, 900.00 and 1050.00. Six of the seven
    // problems put in the sample are flagged.
    money: { pills: PILLS_ZERO_EN, sections: [], note: false, cells: 3 },
    date: { pills: PILLS_ZERO_EN, sections: ["dateStyleSection"], note: false, cells: 5 },
    space: { pills: PILLS_ZERO_EN, sections: ["lookAlikeSection"], note: false, cells: 3 }
  };
  for (const [name, expected] of Object.entries(cases)) {
    const scope = await importCsvInLanguage(sampleOrdersEnVariant(name === "clean" ? null : name), SAMPLE_ORDERS_EN, "en");
    assert.deepEqual(sampleOrdersPills(scope), expected.pills, name);
    assert.deepEqual(sampleOrdersShownSections(scope), expected.sections, name);
    assert.equal(scope.elements.get("statusPanel").className.includes("has-note"), expected.note, name);
    assert.equal(scope.evaluate("state.placeholders.length"), ["na", "dash"].includes(name) ? 1 : 0, name);
    assert.equal(scope.evaluate("state.parseWarnings.length"), 0, name);
    assert.equal(scope.evaluate("state.problemCategories.length") === 0, name === "clean" || name === "title" || name === "money", name);
    scope.evaluate("checkStandardizedExport()");
    assert.equal(scope.evaluate("state.standardizedExport.cellCount"), expected.cells, name);
    assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(state.profiles.map((profile) => profile.typeKey))")), SAMPLE_ORDERS_EN_TYPES, name);
    // The total alone matches on both columns.
    if (name === "total") assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(state.suspectedTotals.map((item) => item.matches[0].fields))")), [["Quantity", "Amount"]]);
  }
  const full = sampleOrdersEnLines();
  for (const [name, fix] of Object.entries(SAMPLE_ORDERS_EN_FIXES)) assert.notDeepEqual(fix(full), full, name);
  assert.equal(sampleOrdersEnVariant(null).split("\n").length, 1 + 40 + 1);
  // Each fix takes its whole problem out: the money fix all three amounts, the date fix all five dates.
  assert.equal(SAMPLE_ORDERS_EN_FIXES.money(full).filter((line) => line.includes("$")).length, 0);
  assert.equal(SAMPLE_ORDERS_EN_FIXES.date(full).filter((line) => /^SO-\d+,(?!\d{4}-\d{2}-\d{2},)/.test(line)).length, 0);
});

test("v3.2.0 commit 12: the English sample in English: the status line, the six pills, every hint with its rows and the two conclusions", async () => {
  const scope = await importCsvInLanguage(fs.readFileSync(path.join(REPO_ROOT, "examples", SAMPLE_ORDERS_EN), "utf8"), SAMPLE_ORDERS_EN, "en");
  assert.equal(scope.elements.get("statusPanel").className, "status-panel success has-note");
  assert.equal(scope.elements.get("statusPanel").textContent,
    "sample-orders-en.csv is loaded and the analysis is ready. The header is on row 2; 1 row above it was skipped: “H1 2025 sales orders”. Row 44 “Total” looks like a total row; you can exclude it from analysis in Data quality.");
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(state.profiles.map((profile) => [profile.field, profile.typeKey]))")), [
    ["Order ID", "id"], ["Order date", "date"], ["Region", "category"], ["Category", "category"], ["Product", "category"],
    ["Customer type", "category"], ["Quantity", "numeric"], ["Amount", "numeric"], ["Sales rep", "category"]
  ]);
  assert.deepEqual(sampleOrdersPills(scope), ["Blanks: 6", "Duplicate rows: 1", "Numbers or dates not recognized: 1 value", "Unusually high or low values: 2 columns", "Duplicate IDs: 1 column", "Values that may mean blank: 2 columns"]);
  assert.deepEqual(sampleOrdersShownSections(scope), ["suspectedTotalSection", "dateStyleSection", "lookAlikeSection"]);
  assert.equal(scope.elements.get("categoryChartInsight").textContent, "In “Region”, the most common value is “East”: 9 times, 22.0%.");
  assert.equal(scope.elements.get("dateChartInsight").textContent, "By the date column “Order date”, the record count rises from 2025-01 to 2025-06, by about 66.7%.");
  assert.equal(scope.evaluate("dom.categoryChartField.value"), "Region");
  assert.equal(sampleOrdersText(scope, "suspectedTotalDetails"), "Row 44 “Total”: Quantity, Amount equal the sum of rows 3\u201343 (41 rows). View Exclude from analysis");
  assert.equal(sampleOrdersText(scope, "dateStyleDetails"),
    "“Order date”: the main format is text YYYY-MM-DD: 36 (e.g. 2025-01-03); 5 other rows use other formats: text MM/DD/YYYY: 3 (e.g. 02/18/2025), text with English month names: 2 (e.g. Apr 16, 2025). View");
  assert.equal(sampleOrdersText(scope, "lookAlikeDetails"),
    "“Region”, 11 rows: leading or trailing spaces in 2 cells (e.g. “South\u2423”); 1 value is written differently: “East” 9 / “east” 1 (upper and lower case). View “Sales rep”, 1 row: leading or trailing spaces in 1 cell (e.g. “\u2423Emma Chen”). View");
  assert.deepEqual(placeholderEntries(scope.evaluate), [["Amount", "N/A", "common", 1, [29]], ["Sales rep", "\u2014", "common", 1, [16]]]);
  assert.deepEqual(idDuplicates(scope.evaluate).map((item) => [item.field, item.values.map((value) => [value.value, value.count, value.rowNumbers])]), [["Order ID", [["SO-1018", 2, [20, 21]]]]]);
  const { names, mismatches } = problemRowCheck(scope);
  assert.deepEqual(mismatches, []);
  assert.deepEqual(names, [
    "missing / Order date = 1", "missing / Region = 1", "missing / Category = 1", "missing / Product = 1", "missing / Customer type = 1", "missing / Sales rep = 1",
    "conversion / Amount = 1", "dateStyle / Order date = 5", "outlier / Quantity = 1", "outlier / Amount = 1", "duplicateRows = 2", "suspectedTotal = 1",
    "idDuplicate / Order ID = 2", "lookAlike / Region = 11", "lookAlike / Sales rep = 1", "placeholder / Amount / N/A = 1", "placeholder / Sales rep / \u2014 = 1"
  ]);
  // The same rows as the Chinese sample, except the case variant "east" (row 37), which the Chinese
  // sample has not: its 地区 group is two trailing spaces only.
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(listProblemCategories().map((category) => listProblemRows(category).items.map((item) => item.line)))")), [
    [44], [44], [44], [44], [44], [44], [29], [12, 19, 25, 36, 39], [44], [44], [20, 21], [44], [20, 21], [3, 7, 11, 16, 22, 27, 31, 32, 37, 42, 10], [42], [29], [16]
  ]);
});

// en-issues.csv and en-formula-errors.xlsx (tests/fixtures/stage3/, made by gen-fixtures.cjs): what
// the English sample cannot show, against the hand-written known-issues.json.
test("v3.2.0 commit 12: en-issues.csv: types, reading issues, the kg column, the numeric-looking column, the single value and the variants", async () => {
  const scope = await importStage3Fixture("en-issues.csv");
  const expected = knownIssues["en-issues.csv"];
  const summary = importedSummary(scope.evaluate);
  assert.deepEqual(summary.fields, expected.expectedTypes.map((item) => item.field));
  assert.deepEqual(summary.types, expected.expectedTypes.map((item) => item.expected));
  assert.equal(summary.rows, expected.dataRows);
  assert.equal(summary.duplicateRows, expected.expectedTotals.duplicateRows);
  assert.equal(summary.totalMissing, expected.expectedTotals.totalMissing);
  assert.equal(scope.evaluate("state.headerDetection"), null);
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(state.parseWarnings.map((warning) => ({ line: warning.row, code: warning.code })))")), expected.parseWarnings);
  assert.deepEqual(constantsOf(scope), expected.constantFields);
  assert.deepEqual(lookAlikesOf(scope).map(({ field, edgeSpaceCount, invisibleOnlyCount, groupCount, rowCount, groups }) => ({
    field, edgeSpaceCount, invisibleOnlyCount, groupCount, rowCount,
    groups: groups.map(({ variants, reasons, rowCount: count }) => ({ variants: variants.map(({ text, count: n }) => [text, n]), reasons, rowCount: count }))
  })), expected.lookAlikes);
  assert.equal(scope.evaluate('getFieldProfile("Weight").numberStrategy.unit'), "kg");
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(state.numericLooking.map((item) => item.field))")), ["Floor"]);
  assert.deepEqual(sampleOrdersShownSections(scope), ["numberConversionSection", "numericLookingSection", "constantFieldSection", "lookAlikeSection"]);
  assert.deepEqual(scope.evaluate("state.dateStyles.length"), 0);
  // The fixture's whole point: every one of these is a problem-row category with the rows above.
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(listProblemCategories().map((category) => [category.kind, category.field || '', listProblemRows(category).items.map((item) => item.line)]))")), [
    ["missing", "Status", [13]], ["missing", "Note", [2, 4, 5, 6, 7, 8, 9, 11, 12, 13, 14, 15, 17, 18, 19, 20, 21]],
    ["lookAlike", "Team", expected.issues.find((issue) => issue.id === "N4").fileLines], ["lookAlike", "Note", [10]], ["parseWarnings", "", [16, 13]]
  ]);
});

test("v3.2.0 commit 12: en-formula-errors.xlsx: the Margin errors and the English subtotal match known-issues.json", async () => {
  const scope = await importStage3Fixture("en-formula-errors.xlsx");
  const expected = knownIssues["en-formula-errors.xlsx"];
  assert.deepEqual(formulaErrors(scope), expected.formulaErrors);
  assert.deepEqual(suspectedTotals(scope).map((item) => ({ line: item.line, label: item.label, field: item.field, matches: item.matches })), expected.suspectedTotals);
  const summary = importedSummary(scope.evaluate);
  assert.deepEqual(summary.fields, expected.expectedTypes.map((item) => item.field));
  assert.deepEqual(summary.types, expected.expectedTypes.map((item) => item.expected));
  assert.equal(summary.rows, expected.dataRows);
  assert.equal(summary.totalMissing, expected.expectedTotals.totalMissing);
  assert.equal(scope.evaluate("state.rows.map(getSourceRowNumber).join()"), "2,3,4,5,6,7,8,9,10,11,12");
  assert.deepEqual(sampleOrdersShownSections(scope), ["suspectedTotalSection", "formulaErrorSection"]);
  // Nothing else is flagged: the only high or low value is the subtotal's Qty (no Price).
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(listProblemCategories().map((category) => [category.kind, category.field || '', listProblemRows(category).items.map((item) => item.line)]))")), [
    ["missing", "Price", [12]], ["missing", "Margin", [3, 5, 7, 9, 12]], ["formulaError", "Margin", [3, 5, 7, 9]], ["outlier", "Qty", [12]], ["suspectedTotal", "", [12]]
  ]);
});

// The two fixtures are made by the generator, byte for byte, twice the same.
test("v3.2.0 commit 12: gen-fixtures.cjs writes the two English fixtures byte for byte, and the others as they are", () => {
  const out = fs.mkdtempSync(path.join(require("os").tmpdir(), "stage3-"));
  try {
    const run = require("child_process").spawnSync(process.execPath, [path.join(STAGE3_FIXTURES, "gen-fixtures.cjs"), out], { encoding: "utf8" });
    assert.equal(run.status, 0, run.stderr);
    for (const name of STAGE3_IMPORT_FILES) {
      assert.ok(fs.readFileSync(path.join(out, name)).equals(fs.readFileSync(path.join(STAGE3_FIXTURES, name))), name);
    }
    assert.ok(STAGE3_IMPORT_FILES.includes("en-issues.csv") && STAGE3_IMPORT_FILES.includes("en-formula-errors.xlsx"));
  } finally {
    fs.rmSync(out, { recursive: true, force: true });
  }
  // The invisible character is written as an escape in the generator (CLAUDE.md), and is in the file.
  const generator = fs.readFileSync(path.join(STAGE3_FIXTURES, "gen-fixtures.cjs"), "utf8");
  assert.doesNotMatch(generator, /\u200B/);
  assert.ok(fs.readFileSync(path.join(STAGE3_FIXTURES, "en-issues.csv"), "utf8").includes("Sales\u200B,"));
});

// ---- v3.2.0 commit 13: one name per language, no Chinese in the English interface, same numbers ----

// (1) The product name: the Chinese interface says 表格体检 only, the English one Sheet Checkup only;
// the old names in neither. (The Chinese half is the v3.0.1 test above; this adds the English half
// and that the English name never reaches a Chinese page, report or file name.)
test("v3.2.0 commit 13: each language shows only its own product name, and the old names in neither", async () => {
  const english = JSON.parse(createScriptContext().evaluate("JSON.stringify(I18N_EN)"));
  assert.equal(english["page.title"], "Sheet Checkup | Check CSV and Excel files for data problems");
  assert.equal(english["brand.label"], "Sheet Checkup home");
  assert.equal(english["brand.name"], "Sheet Checkup");
  const hits = [];
  for (const [key, value] of Object.entries(english)) {
    findOldProductWords(`I18N_EN ${key}`, value, hits);
    if (value.includes("表格体检")) hits.push(`I18N_EN ${key}: ${value}`);
  }
  const names = {};
  for (const language of ["zh", "en"]) {
    const own = language === "zh" ? "表格体检" : "Sheet Checkup";
    const other = language === "zh" ? "Sheet Checkup" : "表格体检";
    const file = language === "zh" ? "sample-orders.csv" : "sample-orders-en.csv";
    const scope = await importCsvInLanguage(readExampleCsv(file), file, language);
    const outputs = [];
    const take = (where, text) => {
      outputs.push([where, String(text)]);
      findOldProductWords(`${language} ${where}`, text, hits);
      if (String(text).includes(other)) hits.push(`${language} ${where}: ${String(text).slice(0, 120)}`);
    };
    for (const [id, text] of Object.entries(elementTexts(scope))) take(`#${id}`, text);
    const { html } = reportOf(scope);
    take("report", html);
    const fileNames = [
      scope.evaluate('buildReportFileName("html", new Date(2026, 10, 15))'),
      scope.evaluate('buildDataExportFileName("filtered", new Date(2026, 10, 15))'),
      scope.evaluate('buildStandardizedFileName("data", new Date(2026, 10, 15))'), scope.evaluate('buildStandardizedFileName("changes", new Date(2026, 10, 15), 3)'),
      ...JSON.parse(scope.evaluate("JSON.stringify(state.problemCategories.map((category) => buildProblemRowsFileName(category, new Date(2026, 10, 15))))"))
    ];
    // The report, the data, the two standardized files and one per problem category (16 / 17).
    assert.equal(fileNames.length, language === "zh" ? 20 : 21, language);
    for (const name of fileNames) {
      take("file name", name);
      assert.ok(!name.includes(own), name);
    }
    names[language] = fileNames;
    // The report names the product in its own language: title, heading and footer.
    const title = (html.match(/<title>([^<]*)<\/title>/) || [])[1];
    const heading = html.match(/<h1>[^<]*<\/h1>/g);
    const footer = (html.match(/<footer>([\s\S]*?)<\/footer>/) || [])[1];
    if (language === "zh") {
      assert.equal(title, "sample-orders.csv - 表格体检报告");
      assert.deepEqual(heading, ["<h1>表格体检报告</h1>"]);
      assert.match(footer, new RegExp(`^由<a href="${RENAME_REPO_URL}">表格体检</a> `));
    } else {
      assert.equal(title, "sample-orders-en.csv - Sheet Checkup report");
      assert.deepEqual(heading, ["<h1>Sheet Checkup report</h1>"]);
      assert.match(footer, new RegExp(`^Created by <a href="${RENAME_REPO_URL}">Sheet Checkup</a> `));
    }
    assert.ok(outputs.some(([, text]) => text.includes(own)), `${language}: the own name shows somewhere`);
  }
  assert.deepEqual(hits.slice(0, 20), [], `${hits.length} hits`);
});

// (2) No Chinese in the English interface. Every output of the main operations, in English, on the
// English sample and the two English fixtures, plus three small files built here for what the
// fixtures do not reach. The only Chinese allowed, each for its reason:
//   1. the two switches to Chinese say “中文” (lang="zh-CN"): their text, nothing else on them;
//   2. FORMULA_ERROR_FALLBACK (#错误): the text of an Excel error cell with no readable code, a value
//      of the file, shown as written in both languages (R3);
//   3. BLANK_GROUP_KEY (（空值）), the internal key of the blank group, would be allowed only where it
//      is never shown; it reaches no output at all (pinned below), so no exception is made for it.
// Any other Han character or full-width punctuation fails.
const CHINESE_OR_FULLWIDTH = /[\p{Script=Han}\u3000-\u303F\uFF00-\uFFEF]/u;
const ALL_PROBLEM_KINDS = ["missing", "formulaError", "conversion", "dateStyle", "outlier", "duplicateRows", "suspectedTotal", "idDuplicate", "lookAlike", "placeholder", "parseWarnings"];

function chineseLeftInEnglish(place, text) {
  const scope = chineseLeftInEnglish.constants;
  let rest = String(text);
  if (/^#(languageSwitch|sidebarLanguageSwitch)$/.test(place)) rest = rest.split(" ¦ ").filter((part) => part !== "中文").join(" ¦ ");
  rest = rest.split(scope.fallback).join("");
  const found = rest.match(new RegExp(CHINESE_OR_FULLWIDTH.source, "gu"));
  return found ? `${place}: ${[...new Set(found)].join("")} in …${rest.slice(Math.max(0, rest.search(CHINESE_OR_FULLWIDTH) - 40), rest.search(CHINESE_OR_FULLWIDTH) + 40)}…` : "";
}

// Runs the main operations on an English page and hands every output to take(place, text). Returns
// the problem kinds it opened, each with the operation that reached it.
function walkEnglishOutputs(scope, name, take) {
  const downloads = [];
  put(scope.context, "__downloads", downloads);
  scope.evaluate("downloadUtf8TextFile = (content, fileName) => { __downloads.push([fileName, content]); }");
  const charts = withFakeCharts(scope);
  const reached = {};
  const snapshot = (operation) => {
    for (const [id, text] of Object.entries(elementTexts(scope))) take(`#${id}`, text, `${name} ${operation}`);
    for (const config of charts.splice(0)) take("chart", JSON.stringify([config.data.datasets.map((set) => set.label), Object.values(config.options.scales || {}).map((axis) => axis.title?.text)]), `${name} ${operation}`);
  };
  const flushDownloads = (operation) => {
    for (const [fileName, content] of downloads.splice(0)) {
      take("file name", fileName, `${name} ${operation}`);
      take(`file ${fileName}`, content, `${name} ${operation}`);
    }
  };
  const problemRows = (operation) => {
    const categories = JSON.parse(scope.evaluate("JSON.stringify(state.problemCategories)"));
    categories.forEach((category) => {
      reached[category.kind] ??= `${name}: ${operation}`;
      put(scope.context, "__walkCategory", category);
      take("category label", scope.evaluate("formatProblemCategoryLabel(__walkCategory)"), `${name} ${operation}`);
      take("category aria", scope.evaluate("formatProblemViewLabel(__walkCategory)"), `${name} ${operation}`);
      take("category notes", scope.evaluate("JSON.stringify(listProblemRows(__walkCategory).items.map((item) => item.note))"), `${name} ${operation}`);
      scope.evaluate("openProblemView(__walkCategory); exportProblemRows()");
      flushDownloads(`${operation} ${category.kind}`);
      snapshot(`${operation} ${category.kind}`);
    });
  };
  const fields = (type) => JSON.parse(scope.evaluate(`JSON.stringify(getProfilesByType(${JSON.stringify(type)}).map((profile) => profile.field))`));
  // A breakdown by the first text column (its blank group too, before any row is left out).
  const breakdown = (operation, method) => {
    const [metric] = fields("numeric");
    const [group] = fields("category");
    if (!metric || !group) return;
    put(scope.context, "__walkMetric", metric);
    put(scope.context, "__walkGroup", group);
    put(scope.context, "__walkDate", fields("date")[0] || "");
    scope.evaluate(`dom.v2MetricField.value = __walkMetric; dom.v2GroupField.value = __walkGroup; dom.v2AggregateMethod.value = ${JSON.stringify(method)}; dom.v2DateField.value = __walkDate; renderV2Analysis()`);
    snapshot(operation);
  };

  snapshot("import");
  breakdown("breakdown at import", "sum");
  scope.evaluate("exportAnalysisReport()");
  flushDownloads("report");
  snapshot("report");
  problemRows("problem rows");
  // A filter: the first text column keeps the first half of its values.
  const category = fields("category")[0];
  if (category) {
    put(scope.context, "__walkField", category);
    scope.evaluate(`(() => {
      const values = getCategoryFilterOptions(__walkField).map((option) => option.value);
      dom.categoryFilterField.value = __walkField; updateCategoryFilterField();
      state.categoryFilterDraftValues = new Set(values.slice(0, Math.max(1, Math.ceil(values.length / 2))));
      applyFilters();
    })()`);
    snapshot("filter");
  }
  scope.evaluate('exportProcessedCsv("filtered")');
  flushDownloads("data export");
  // Both kinds left out, where the file has them; the notes of the problem rows say so.
  scope.evaluate('setRowExclusion("suspectedTotal", true); setRowExclusion("duplicateRows", true)');
  snapshot("exclude");
  problemRows("problem rows after excluding");
  // Treat as blank: all, then one undone.
  scope.evaluate("confirmAllCommonPlaceholders()");
  snapshot("treat as blank");
  scope.evaluate("(() => { const entry = state.placeholders.find((item) => item.confirmed); if (entry) undoPlaceholderConfirmation(entry.field, entry.key); })()");
  snapshot("undo blank");
  problemRows("problem rows after blanks");
  // Column settings: a draft (the gates say so), applied (a placeholder of an ignored column), restored.
  const draftField = JSON.parse(scope.evaluate("JSON.stringify(state.placeholders[0]?.field || state.fields[state.fields.length - 1])"));
  put(scope.context, "__walkDraft", draftField);
  scope.evaluate('stageFieldTypeChange(__walkDraft, "ignore"); exportAnalysisReport()');
  snapshot("column draft");
  scope.evaluate("applyFieldConfiguration()");
  snapshot("column applied");
  problemRows("problem rows with an ignored column");
  scope.evaluate("restoreAutomaticFieldTypes()");
  snapshot("columns restored");
  // The standardized table and its change list.
  scope.evaluate('checkStandardizedExport(); exportStandardizedData("data"); exportStandardizedData("changes")');
  flushDownloads("standardized");
  snapshot("standardized");
  // A breakdown and the trend summed by a number, where the file has the columns.
  breakdown("breakdown", "median");
  const [metric] = fields("numeric");
  const [dateField] = fields("date");
  if (metric && dateField) {
    put(scope.context, "__walkMetric", metric);
    put(scope.context, "__walkDate", dateField);
    scope.evaluate('dom.dateChartField.value = __walkDate; dom.dateChartMetric.value = __walkMetric; dom.dateChartAggregate.value = "sum"; renderDateTrendChart()');
    snapshot("trend");
  }
  scope.evaluate("exportAnalysisReport()");
  flushDownloads("report at the end");
  scope.evaluate("clearAllFilters()");
  snapshot("filters cleared");
  return reached;
}

// The three small files: what the English fixtures do not reach.
//   single-column.csv: a reading issue without a row (UndetectableDelimiter, separator “;”);
//   errors-and-dates.xlsx: an error cell with no code (FORMULA_ERROR_FALLBACK) and Excel date cells
//   mixed with text dates (the Excel date cell format of mixed date formats).
function englishExtraFiles() {
  const serial = (y, m, d) => (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000;
  const cells = {
    A1: { t: "s", v: "Item" }, B1: { t: "s", v: "Date" }, C1: { t: "s", v: "Value" }
  };
  const rows = [["A", serial(2025, 1, 3), 10], ["B", serial(2025, 1, 6), 12], ["C", "2025/1/9", { e: 255 }], ["D", serial(2025, 1, 14), 9], ["E", serial(2025, 1, 20), 11], ["F", "2025/1/25", 13]];
  rows.forEach(([item, date, value], index) => {
    const r = index + 2;
    cells[`A${r}`] = { t: "s", v: item };
    cells[`B${r}`] = typeof date === "number" ? { t: "n", v: date, z: "yyyy-mm-dd" } : { t: "s", v: date };
    cells[`C${r}`] = typeof value === "object" ? { t: "e", v: value.e } : { t: "n", v: value };
  });
  return {
    "single-column.csv": "Name;Team\nAnn;Sales\nBob;Support\nCy;Sales\n",
    "errors-and-dates.xlsx": buildCellSheetBuffer(cells, "A1:C7")
  };
}

test("v3.2.0 commit 13: in English, no output of any operation on any English file has Chinese, except the three allowed and named", async () => {
  const constants = createScriptContext();
  chineseLeftInEnglish.constants = { fallback: constants.evaluate("FORMULA_ERROR_FALLBACK"), blankKey: constants.evaluate("BLANK_GROUP_KEY") };
  assert.deepEqual(chineseLeftInEnglish.constants, { fallback: "#错误", blankKey: "（空值）" });
  const problems = [];
  const reached = {};
  const allowedSeen = { switches: 0, fallback: 0, blankKey: 0 };
  let outputs = 0;
  const take = (where, text, operation) => {
    outputs += 1;
    const value = String(text);
    if (/^#(languageSwitch|sidebarLanguageSwitch)$/.test(where) && value.split(" ¦ ").includes("中文")) allowedSeen.switches += 1;
    if (value.includes(chineseLeftInEnglish.constants.fallback)) allowedSeen.fallback += 1;
    if (value.includes(chineseLeftInEnglish.constants.blankKey)) allowedSeen.blankKey += 1;
    const left = chineseLeftInEnglish(where, value);
    if (left) problems.push(`${operation} ${left}`);
  };
  await withRecordedAttributes(async () => {
    const files = [
      [SAMPLE_ORDERS_EN, () => importCsvInLanguage(readExampleCsv(SAMPLE_ORDERS_EN), SAMPLE_ORDERS_EN, "en")],
      ["en-issues.csv", () => importStage3Fixture("en-issues.csv", "en")],
      ["en-formula-errors.xlsx", () => importStage3Fixture("en-formula-errors.xlsx", "en")],
      // Blank values among ordinary rows (not only on a total row), and a value starting with a bracket.
      ["orders.csv", () => importCsvInLanguage(englishScenarioCsv(), "orders.csv", "en")],
      ...Object.entries(englishExtraFiles()).map(([name, content]) => [name, () => (typeof content === "string" ? importCsvInLanguage(content, name, "en") : importWorkbookFile(content, name, "en"))])
    ];
    for (const [name, open] of files) {
      const scope = await open();
      assert.equal(scope.evaluate("currentLanguage"), "en", name);
      assert.ok(scope.evaluate("state.rows.length") > 0, name);
      for (const [kind, where] of Object.entries(walkEnglishOutputs(scope, name, take))) reached[kind] ??= where;
    }
  });
  assert.deepEqual(problems.slice(0, 15), [], `${problems.length} outputs with Chinese`);
  assert.ok(outputs > 5000, `${outputs} outputs`);
  // Every kind of problem rows was opened (its label, notes, viewer and exported file looked at).
  assert.deepEqual(Object.keys(reached).sort(), [...ALL_PROBLEM_KINDS].sort());
  assert.deepEqual(reached, {
    missing: "sample-orders-en.csv: problem rows",
    conversion: "sample-orders-en.csv: problem rows",
    dateStyle: "sample-orders-en.csv: problem rows",
    outlier: "sample-orders-en.csv: problem rows",
    duplicateRows: "sample-orders-en.csv: problem rows",
    suspectedTotal: "sample-orders-en.csv: problem rows",
    idDuplicate: "sample-orders-en.csv: problem rows",
    lookAlike: "sample-orders-en.csv: problem rows",
    placeholder: "sample-orders-en.csv: problem rows",
    parseWarnings: "en-issues.csv: problem rows",
    formulaError: "en-formula-errors.xlsx: problem rows"
  });
  // The allowed ones were really met (the rules are not empty words): the switches on every page,
  // the fallback in the built workbook; the blank group's key nowhere, not even in an attribute.
  assert.ok(allowedSeen.switches > 0 && allowedSeen.fallback > 0, JSON.stringify(allowedSeen));
  assert.equal(allowedSeen.blankKey, 0);
});

// (3) Principle 1: changing the language changes words only. The same file in both languages gives
// the same state, problem rows and line numbers, the same report snapshot once its strings are
// taken out, the same numbers in every element and the same filter order.
const SAME_IN_BOTH_STATE = `JSON.stringify({
  rows: state.rows.length, fields: state.fields, totalMissing: state.totalMissing, duplicateRows: state.duplicateRows,
  profiles: state.profiles, qualityNumericStats: state.qualityNumericStats, numericStats: state.numericStats, categoryStats: state.categoryStats,
  idDuplicates: state.idDuplicates, placeholders: state.placeholders, numericLooking: state.numericLooking, suspectedTotals: state.suspectedTotals,
  formulaErrors: state.formulaErrors, parseWarnings: state.parseWarnings, dateStyles: state.dateStyles, constantFields: state.constantFields,
  lookAlikes: state.lookAlikes, headerDetection: state.headerDetection, trend: state.trendInsightData, categoryChartField: dom.categoryChartField.value,
  problems: listProblemCategories().map((category) => [category, listProblemRows(category).items.map((item) => [item.line, item.rowIndex])]),
  filterOrder: getProfilesByType("category").map((profile) => [profile.field, getCategoryFilterOptions(profile.field).map((option) => [option.value, option.count])]),
  standardized: state.standardizedExport && { counts: state.standardizedExport.counts, cells: state.standardizedExport.cells },
  custom: state.customAnalysis && state.customAnalysis.grouped
})`;
// A breakdown by the first number and text columns (with the date column when there is one).
const SAME_IN_BOTH_BREAKDOWN = `(() => {
  const [metric] = getProfilesByType("numeric"); const [group] = getProfilesByType("category"); const [date] = getProfilesByType("date");
  if (!metric || !group) return;
  dom.v2MetricField.value = metric.field; dom.v2GroupField.value = group.field; dom.v2AggregateMethod.value = "sum"; dom.v2DateField.value = date ? date.field : "";
  renderV2Analysis();
})()`;
const withoutStrings = (value) => (typeof value === "string" ? "" : Array.isArray(value) ? value.map(withoutStrings)
  : value && typeof value === "object" ? Object.fromEntries(Object.entries(value).map(([key, item]) => [key, withoutStrings(item)])) : value);

test("v3.2.0 commit 13: the same file in Chinese and in English gives the same numbers, rows, report figures and filter order", async () => {
  const files = [
    ...STAGE3_IMPORT_FILES.map((name) => [name, (language) => importStage3Fixture(name, language)]),
    ...["sample-orders.csv", SAMPLE_ORDERS_EN].map((name) => [name, (language) => importCsvInLanguage(readExampleCsv(name), name, language)]),
    // Values that start with a bracket next to blanks: the filter's sort key must be the Chinese one in
    // both languages (R3), or “(Blank)” and “（空值）” sort to different places.
    ["orders.csv", (language) => importCsvInLanguage(englishScenarioCsv(), "orders.csv", language)]
  ];
  assert.equal(files.length, 13);
  const gaps = [];
  await withRecordedAttributes(async () => {
  for (const [name, open] of files) {
    const [zh, en] = [await open("zh"), await open("en")];
    const steps = [["import", ""], ["breakdown", SAME_IN_BOTH_BREAKDOWN], ["blanks and standardized", "confirmAllCommonPlaceholders(); checkStandardizedExport()"]];
    for (const [step, expression] of steps) {
      for (const scope of [zh, en]) if (expression) scope.evaluate(expression);
      if (zh.evaluate(SAME_IN_BOTH_STATE) !== en.evaluate(SAME_IN_BOTH_STATE)) {
        const a = JSON.parse(zh.evaluate(SAME_IN_BOTH_STATE));
        const b = JSON.parse(en.evaluate(SAME_IN_BOTH_STATE));
        gaps.push(`${name} ${step} state: ${Object.keys(a).filter((key) => JSON.stringify(a[key]) !== JSON.stringify(b[key])).join(", ")}`);
      }
      // The report writes the time of the analysis, each import's own clock: the same moment for both.
      for (const scope of [zh, en]) scope.evaluate('state.analysisCompletedAt = new Date("2026-10-25T08:00:00.000Z")');
      const reports = [zh, en].map((scope) => reportOf(scope));
      if (JSON.stringify(withoutStrings(reports[0].report)) !== JSON.stringify(withoutStrings(reports[1].report))) gaps.push(`${name} ${step} report snapshot`);
      const reportNumbers = reports.map(({ html }) => numbersOf(html.replace(/<[^>]+>/g, " ")).filter((number) => !ENGLISH_EXTRA_NUMBERS.includes(number)).join(" "));
      if (reportNumbers[0] !== reportNumbers[1]) gaps.push(`${name} ${step} report numbers`);
      const zhTexts = elementTexts(zh);
      const enTexts = elementTexts(en);
      for (const id of Object.keys({ ...zhTexts, ...enTexts })) {
        const a = numbersOf(zhTexts[id] || "").filter((number) => !ENGLISH_EXTRA_NUMBERS.includes(number)).join(" ");
        const b = numbersOf(enTexts[id] || "").filter((number) => !ENGLISH_EXTRA_NUMBERS.includes(number)).join(" ");
        if (a !== b) gaps.push(`${name} ${step} #${id}: ${a.slice(0, 80)} / ${b.slice(0, 80)}`);
      }
    }
  }
  });
  assert.deepEqual(gaps.slice(0, 15), [], `${gaps.length} differences`);
  // The bracket case is really there: the blank value sorts first in both (（ is after ( in zh-CN order).
  const scope = await importCsvInLanguage(englishScenarioCsv(), "orders.csv", "en");
  assert.deepEqual(JSON.parse(scope.evaluate('JSON.stringify(getCategoryFilterOptions("Region").map((option) => [option.value, option.label]))')).slice(0, 2), [["", "(Blank)"], ["(A)", "(A)"]]);
});

// ---- v3.2.1 commit 1: read one worksheet at a time ----

// Puts a counting XLSX into the context: every XLSX.read call's options, in order.
function countXlsxReads(scope) {
  const reads = [];
  const counting = Object.assign(Object.create(XLSX), {
    read(data, options) {
      reads.push({ ...options });
      return XLSX.read(data, options);
    }
  });
  put(scope.context, "XLSX", counting);
  scope.context.window.XLSX = counting;
  return reads;
}

// A workbook file as parseDataFile() gets it.
function putWorkbookFile(scope, buffer, fileName) {
  const bytes = Buffer.from(buffer);
  put(scope.context, "__workbookFile", {
    name: fileName,
    size: bytes.length,
    arrayBuffer: async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.length)
  });
}

const V321_GOOD_ROWS = [["编号", "金额"], ["A-1", 10], ["A-2", 20], ["A-3", 30]];
const v321TwoSheets = () => buildXlsxBuffer([{ name: "说明", rows: [["本工作簿为虚构数据"], ["第二行"]] }, { name: "数据", rows: V321_GOOD_ROWS }]);

test("v3.2.1 commit 1: a multi-sheet file reads its names first and then only the chosen sheet; a single sheet reads twice", async () => {
  const multi = createExcelContext();
  const reads = countXlsxReads(multi);
  putWorkbookFile(multi, v321TwoSheets(), "two.xlsx");
  await multi.evaluate("parseDataFile(__workbookFile)");
  // The list of worksheets shows after one read of the names alone.
  assert.equal(multi.evaluate("Boolean(state.pendingExcel)"), true);
  assert.deepEqual(reads, [{ type: "array", bookSheets: true }]);
  multi.elements.get("excelSheetSelect").value = "数据";
  await multi.evaluate("analyzeSelectedExcelSheet()");
  assert.deepEqual(reads, [{ type: "array", bookSheets: true }, { type: "array", cellDates: true, cellNF: true, sheets: "数据" }]);
  assert.equal(multi.evaluate("state.rows.length"), 3);
  // A single sheet: the names, then that sheet.
  const single = createExcelContext();
  const singleReads = countXlsxReads(single);
  putWorkbookFile(single, buildXlsxBuffer([{ name: "S1", rows: V321_GOOD_ROWS }]), "one.xlsx");
  await single.evaluate("parseDataFile(__workbookFile)");
  assert.deepEqual(singleReads, [{ type: "array", bookSheets: true }, { type: "array", cellDates: true, cellNF: true, sheets: "S1" }]);
  assert.equal(single.evaluate("state.rows.length"), 3);
  // The waiting state keeps the file's bytes and names, never a parsed workbook.
  assert.deepEqual(JSON.parse(multi.evaluate("JSON.stringify(Object.keys(state.pendingExcel || {}))")), []);
  const waiting = createExcelContext();
  countXlsxReads(waiting);
  const waitingBook = Buffer.from(v321TwoSheets());
  putWorkbookFile(waiting, waitingBook, "two.xlsx");
  await waiting.evaluate("parseDataFile(__workbookFile)");
  assert.deepEqual(JSON.parse(waiting.evaluate("JSON.stringify(Object.keys(state.pendingExcel).sort())")), ["bytes", "fileName", "importId", "sheetNames"]);
  // (The buffer comes from the test's realm, so instanceof ArrayBuffer would be false in the VM.)
  assert.equal(waiting.evaluate("Object.prototype.toString.call(state.pendingExcel.bytes)"), "[object ArrayBuffer]");
  assert.equal(waiting.evaluate("state.pendingExcel.bytes.byteLength"), waitingBook.length);
  assert.deepEqual(JSON.parse(waiting.evaluate("JSON.stringify(state.pendingExcel.sheetNames)")), ["说明", "数据"]);
  // A CSV never reads a workbook.
  const csv = createExcelContext();
  const csvReads = countXlsxReads(csv);
  put(csv.context, "Papa", Papa);
  csv.context.window.Papa = Papa;
  csv.evaluate('parseCsvText("编号,金额\\nA-1,10\\nA-2,20", "a.csv")');
  assert.equal(csv.evaluate("state.rows.length"), 2);
  assert.deepEqual(csvReads, []);
});

test("v3.2.1 commit 1: reading one sheet gives exactly what reading the whole workbook gives, errors too, in both languages", () => {
  const scope = createExcelContext();
  const wide = [Array.from({ length: 201 }, (_, index) => `列${index + 1}`), Array.from({ length: 201 }, (_, index) => index)];
  const books = [
    ...["en-formula-errors.xlsx", "公式错误.xlsx", "花名册.xlsx", "门店报表.xlsx"].map((name) => [name, fs.readFileSync(path.join(STAGE3_FIXTURES, name))]),
    ["excel-import-test.xlsx", fs.readFileSync(path.join(__dirname, "fixtures", "excel-import-test.xlsx"))],
    ["checks.xlsx", Buffer.from(buildXlsxBuffer([
      { name: "好表", rows: V321_GOOD_ROWS },
      { name: "空表", rows: [] },
      { name: "只有表头", rows: [["编号", "金额"]] },
      { name: "没有表头", rows: [[1, 2], [3, 4]] },
      { name: "表头重复", rows: [["编号", "金额", "金额"], ["a", 1, 2]] },
      { name: "超列", rows: wide },
      { name: "有标题行", rows: [["2025 年订单"], [], ["编号", "金额"], ["T-1", 10], ["T-2", 20]] }
    ]))]
  ];
  let compared = 0;
  for (const [name, bytes] of books) {
    put(scope.context, "__bytes", bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.length));
    const whole = XLSX.read(scope.context.__bytes, { type: "array", cellDates: true, cellNF: true });
    put(scope.context, "__whole", whole);
    assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(readExcelSheetNames(__bytes))")), whole.SheetNames, name);
    for (const sheetName of whole.SheetNames) {
      put(scope.context, "__sheetName", sheetName);
      for (const language of ["zh", "en"]) {
        const [one, all] = JSON.parse(scope.evaluate(`(() => {
          currentLanguage = ${JSON.stringify(language)};
          const run = (read) => { try { return JSON.stringify(read()); } catch (error) { return "error: " + error.message; } };
          const one = run(() => readExcelSheet(__bytes, __sheetName));
          const all = run(() => {
            const worksheet = __whole.Sheets[__sheetName];
            if (!worksheet || !worksheet["!ref"]) throw languageError(Lp(\`工作表「\${__sheetName}」为空，请选择包含表头和数据的工作表。\`, \`The worksheet “\${__sheetName}” is empty; choose one with a header row and data.\`));
            return readExcelSheetTable(worksheet, __sheetName);
          });
          currentLanguage = "zh";
          return JSON.stringify([one, all]);
        })()`));
        assert.equal(one, all, `${name} / ${sheetName} / ${language}`);
        compared += 1;
      }
    }
  }
  // 4 single sheets, 4 sheets of excel-import-test.xlsx and 7 of checks.xlsx, twice each.
  assert.equal(compared, 30);
});

test("v3.2.1 commit 1: the chosen sheet says it is being read; a sheet that fails leaves the list waiting, the error rewritable", async () => {
  const scope = createExcelContext();
  putWorkbookFile(scope, buildXlsxBuffer([{ name: "好表", rows: V321_GOOD_ROWS }, { name: "表头重复", rows: [["编号", "金额", "金额"], ["a", 1, 2]] }]), "book.xlsx");
  await scope.evaluate("parseDataFile(__workbookFile)");
  let focused = 0;
  scope.elements.get("confirmSheetSelection").focus = () => { focused += 1; };
  scope.elements.get("excelSheetSelect").value = "表头重复";
  const pending = scope.evaluate("state.pendingExcel");
  const done = scope.evaluate("analyzeSelectedExcelSheet()");
  // Before the read: the message, the page busy, the switches locked.
  assert.equal(textOf(scope, "sheetSelectionMessage"), "正在读取并分析工作表「表头重复」...");
  assert.equal(scope.evaluate("document.body.classList.contains('is-loading')"), true);
  assert.equal(scope.elements.get("languageSwitch").disabled, true);
  await done;
  assert.equal(textOf(scope, "sheetSelectionMessage"), "工作表「表头重复」包含重复表头「金额」，请修改后重新上传。");
  assert.equal(scope.evaluate("document.body.classList.contains('is-loading')"), false);
  assert.equal(scope.evaluate("state.pendingExcel") === pending, true);
  assert.equal(focused, 1);
  // The error is kept as a function: written again in English (the switch itself is locked while
  // a sheet waits, so the rewrite is called directly).
  scope.evaluate('setInterfaceLanguage("en"); rewriteOneOffMessages()');
  assert.equal(textOf(scope, "sheetSelectionMessage"), "The worksheet “表头重复” has the header “金额” more than once; change it and upload the file again.");
  scope.evaluate('setInterfaceLanguage("zh")');
  // A good sheet afterwards: analyzed, and the message is gone, also from the rewrite list.
  scope.elements.get("excelSheetSelect").value = "好表";
  await scope.evaluate("analyzeSelectedExcelSheet()");
  assert.equal(scope.evaluate("state.rows.length"), 3);
  assert.equal(textOf(scope, "sheetSelectionMessage"), "");
  assert.equal(scope.evaluate("oneOffMessages.has(dom.sheetSelectionMessage)"), false);
});

test("v3.2.1 commit 1: an unexpected error while the chosen sheet is committed ends the import with the error", async () => {
  const scope = createExcelContext();
  putWorkbookFile(scope, v321TwoSheets(), "two.xlsx");
  await scope.evaluate("parseDataFile(__workbookFile)");
  scope.evaluate('countDuplicateRows = () => { throw new Error("commit broke"); }');
  scope.elements.get("excelSheetSelect").value = "数据";
  let rejected = null;
  await scope.evaluate("analyzeSelectedExcelSheet()").catch((error) => { rejected = error; });
  assert.equal(rejected, null);
  assert.equal(scope.evaluate("document.body.classList.contains('is-loading')"), false);
  assert.equal(scope.elements.get("csvFile").disabled, false);
  assert.equal(scope.elements.get("languageSwitch").disabled, false);
  assert.equal(textOf(scope, "statusPanel"), "commit broke");
  assert.match(scope.elements.get("statusPanel").className, /\berror\b/);
  assert.equal(scope.evaluate("state.rows.length"), 0);
});

// ---- v3.2.1 commit 1b: the check after the frame in the sheet list ----

// A multi-sheet file waiting in the list, with yieldToBrowser() held until release() is called.
async function v321HeldSheetList(sheets) {
  const scope = createExcelContext();
  const reads = countXlsxReads(scope);
  put(scope.context, "Papa", Papa);
  scope.context.window.Papa = Papa;
  putWorkbookFile(scope, buildXlsxBuffer(sheets), "book.xlsx");
  await scope.evaluate("parseDataFile(__workbookFile)");
  assert.equal(scope.evaluate("Boolean(state.pendingExcel)"), true);
  scope.evaluate("globalThis.__realYield = yieldToBrowser; yieldToBrowser = () => new Promise((resolve) => { globalThis.__releaseYield = resolve; })");
  return {
    scope,
    reads,
    release: () => scope.evaluate("globalThis.__releaseYield()"),
    restore: () => scope.evaluate("yieldToBrowser = globalThis.__realYield")
  };
}

const V321_NEW_CSV = "城市,人数\n北京,1\n上海,2\n广州,3";

test("v3.2.1 commit 1b: an import begun while a failing sheet waits for its frame is left alone", async () => {
  const held = await v321HeldSheetList([{ name: "好表", rows: V321_GOOD_ROWS }, { name: "只有表头", rows: [["编号", "金额"]] }]);
  const { scope, reads } = held;
  try {
    assert.equal(reads.length, 1);
    scope.elements.get("excelSheetSelect").value = "只有表头";
    const choosing = scope.evaluate("analyzeSelectedExcelSheet()");
    assert.equal(textOf(scope, "sheetSelectionMessage"), "正在读取并分析工作表「只有表头」...");
    // Another file begins while the frame is drawn: still being read (loading), its own message.
    put(scope.context, "__newCsv", V321_NEW_CSV);
    const importId = scope.evaluate('globalThis.__newImport = beginImport(); setStatus("正在读取 new.csv ..."); __newImport');
    held.release();
    await choosing;
    // The sheet was never read, and nothing of it reached the page or the new import.
    assert.equal(reads.length, 1);
    assert.equal(textOf(scope, "sheetSelectionMessage"), "");
    assert.equal(scope.evaluate("document.body.classList.contains('is-loading')"), true);
    assert.equal(textOf(scope, "statusPanel"), "正在读取 new.csv ...");
    assert.doesNotMatch(scope.elements.get("statusPanel").className, /\berror\b/);
    assert.equal(scope.evaluate("state.activeImportId"), importId);
    // The new import finishes as its own.
    scope.evaluate("parseCsvText(__newCsv, \"new.csv\", __newImport)");
    assert.equal(scope.evaluate("state.sourceFileName"), "new.csv");
    assert.equal(scope.evaluate("state.rows.length"), 3);
    assert.equal(textOf(scope, "statusPanel"), "new.csv 已成功载入，分析结果已生成。");
    assert.equal(textOf(scope, "sheetSelectionMessage"), "");
  } finally {
    held.restore();
  }
});

test("v3.2.1 commit 1b: an import done while a readable sheet waits for its frame is not overwritten", async () => {
  const held = await v321HeldSheetList([{ name: "说明", rows: [["本工作簿为虚构数据"], ["第二行"]] }, { name: "数据", rows: V321_GOOD_ROWS }]);
  const { scope, reads } = held;
  try {
    scope.elements.get("excelSheetSelect").value = "数据";
    const choosing = scope.evaluate("analyzeSelectedExcelSheet()");
    put(scope.context, "__newCsv", V321_NEW_CSV);
    scope.evaluate('parseCsvText(__newCsv, "new.csv")');
    assert.equal(scope.evaluate("state.sourceFileName"), "new.csv");
    const status = textOf(scope, "statusPanel");
    held.release();
    await choosing;
    assert.equal(reads.length, 1);
    assert.equal(scope.evaluate("state.sourceFileName"), "new.csv");
    assert.equal(scope.evaluate("state.sourceSheetName"), "");
    assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(state.fields)")), ["城市", "人数"]);
    assert.equal(scope.evaluate("state.rows.length"), 3);
    assert.equal(textOf(scope, "statusPanel"), status);
    assert.equal(scope.evaluate("document.body.classList.contains('is-loading')"), false);
    assert.equal(textOf(scope, "sheetSelectionMessage"), "");
  } finally {
    held.restore();
  }
});

// ---- v3.2.1 commit 2: keep a multi-sheet file and switch worksheets ----

// The English scenario's orders as a worksheet (text cells, as the CSV reads them), a second sheet
// with other orders, and sheets that fail.
function v321ScenarioRows(shift) {
  const rows = Papa.parse(englishScenarioCsv().trim(), { header: false }).data;
  if (!shift) return rows;
  return rows.map((row, index) => (index === 0 || row[0] === "Total" ? row : row.map((cell, column) => (column === 2 && /^\d+$/.test(cell) ? String(Number(cell) + shift) : cell))));
}
const v321ScenarioBook = () => buildXlsxBuffer([
  { name: "Orders", rows: v321ScenarioRows(0) },
  { name: "Second", rows: v321ScenarioRows(7) },
  { name: "Empty", rows: [] },
  { name: "HeaderOnly", rows: [["ID", "Amount"]] },
  { name: "Doubled", rows: [["ID", "Amount", "Amount"], ["a", 1, 2]] }
]);
const V321_FIXED_TIME = "new Date(2026, 10, 23, 12, 0, 0)";

// A context with the scenario workbook imported and one sheet chosen through the list.
async function v321ImportSheet(sheetName, { buffer = v321ScenarioBook(), language } = {}) {
  const scope = createExcelContext(language === undefined ? {} : { language });
  const reads = countXlsxReads(scope);
  put(scope.context, "Papa", Papa);
  scope.context.window.Papa = Papa;
  putWorkbookFile(scope, buffer, "book.xlsx");
  await scope.evaluate("parseDataFile(__workbookFile)");
  scope.elements.get("excelSheetSelect").value = sheetName;
  await scope.evaluate("analyzeSelectedExcelSheet()");
  assert.equal(scope.evaluate("state.sourceSheetName"), sheetName);
  return { scope, reads };
}

// Everything a user could set on an analysis: blanks, both exclusions, a filter, a column type, a
// breakdown, the standardized check and an open category of problem rows.
function v321ApplyScenario(scope) {
  scope.evaluate("confirmAllCommonPlaceholders()");
  scope.evaluate(`["suspectedTotal", "duplicateRows"].forEach((kind) => handleExclusionClick({ target: { closest: () => ({ dataset: { exclusionKind: kind, exclusionOn: "true" } }) } }))`);
  // (Applying column settings clears the filters, so the type goes first.)
  scope.evaluate('stageFieldTypeChange("Rep", "ignore"); applyFieldConfiguration()');
  scope.evaluate(`(() => {
    const values = getCategoryFilterOptions("Region").map((option) => option.value);
    dom.categoryFilterField.value = "Region"; dom.numericFilterField.value = ""; dom.dateFilterField.value = "";
    state.categoryFilterDraftValues = new Set(values.slice(0, 2));
    applyFilters();
  })()`);
  scope.evaluate('dom.v2MetricField.value = "Amount"; dom.v2GroupField.value = "Region"; dom.v2AggregateMethod.value = "sum"; dom.v2DateField.value = "Date"; renderV2Analysis()');
  scope.evaluate("checkStandardizedExport()");
  scope.evaluate("openProblemView(state.problemCategories[0])");
  // Every one of them took: the comparison below can tell a switch that kept any of them.
  const set = JSON.parse(scope.evaluate(`JSON.stringify({
    placeholders: Object.keys(state.placeholderConfirmations).length,
    excluded: state.excludedRows.suspectedTotal.length + state.excludedRows.duplicateRows.length,
    filter: state.filters.category ? state.filters.category.values.length : 0,
    types: Object.keys(state.manualFieldTypes).length,
    custom: Boolean(state.customAnalysis), standardized: Boolean(state.standardizedExport), problems: Boolean(state.problemView)
  })`));
  assert.deepEqual(set, { placeholders: 1, excluded: 2, filter: 2, types: 1, custom: true, standardized: true, problems: true });
}

// What the page shows of an analysis: the report (time fixed), the status bar, the sidebar, the
// file names and the figures.
function v321Snapshot(scope) {
  scope.evaluate(`state.analysisCompletedAt = ${V321_FIXED_TIME}`);
  return JSON.parse(scope.evaluate(`JSON.stringify({
    report: buildHtmlReport(buildReportData(${V321_FIXED_TIME}), []),
    status: dom.statusPanel.textContent, summary: dom.statusNoteSummary.textContent,
    sourceName: dom.sourceName.textContent, title: dom.sourceName.title, topbar: dom.topbarFileName.textContent,
    size: dom.sidebarFileSize.textContent,
    names: [buildReportFileName("html", ${V321_FIXED_TIME}), buildDataExportFileName("filtered", ${V321_FIXED_TIME}), buildStandardizedFileName("data", ${V321_FIXED_TIME}), buildStandardizedFileName("changes", ${V321_FIXED_TIME}), buildProblemRowsFileName(state.problemCategories[0] || { kind: "missing", field: "x" }, ${V321_FIXED_TIME})],
    fields: state.fields, rows: state.rows.length, types: state.profiles.map((profile) => profile.typeKey),
    module: document.body.dataset.module ?? null,
    // The one-off messages of the old analysis (and the switch's own) must be gone as after an import.
    messages: ["sidebarSheetMessage", "uploadSheetMessage", "placeholderMessage", "fieldConfigMessage", "filterMessage", "problemRowsMessage", "duplicateExclusionMessage", "suspectedTotalMessage", "standardizedExportStatus", "dataExportStatus", "reportExportStatus"].map((id) => [id, document.getElementById(id).textContent])
  })`));
}

// The snapshot without the worksheet lists' own messages (a failed switch writes its error there).
const v321WithoutSheetMessages = (snapshot) => ({ ...snapshot, messages: snapshot.messages.filter(([id]) => !/SheetMessage$/.test(id)) });

test("v3.2.1 commit 2: switching worksheets gives exactly what importing the file and choosing that sheet gives", async () => {
  for (const language of ["zh", "en"]) {
    const { scope } = await v321ImportSheet("Orders", { language });
    v321ApplyScenario(scope);
    await scope.evaluate('switchExcelSheet("Second", "sidebar")');
    const fresh = await v321ImportSheet("Second", { language });
    assert.deepEqual(v321Snapshot(scope), v321Snapshot(fresh.scope), language);
    assert.equal(scope.evaluate("state.sourceSheetName"), "Second");
  }
});

test("v3.2.1 commit 2: a switch clears the filter, the rows left out, the blanks, the column types and the rest", async () => {
  const { scope } = await v321ImportSheet("Orders");
  v321ApplyScenario(scope);
  await scope.evaluate('switchExcelSheet("Second", "sidebar")');
  const after = JSON.parse(scope.evaluate(`JSON.stringify({
    filters: state.filters, excluded: state.excludedRows, placeholders: state.placeholderConfirmations,
    types: state.manualFieldTypes, drafts: state.fieldTypeDrafts, inferred: Object.fromEntries(state.profiles.map((profile) => [profile.field, profile.inferredTypeKey])),
    problems: state.problemView, standardized: state.standardizedExport, custom: state.customAnalysis,
    module: document.body.dataset.module, previewPage: state.previewPage
  })`));
  assert.deepEqual(after.filters, JSON.parse(scope.evaluate("JSON.stringify(createEmptyFilterConfig())")));
  assert.deepEqual(after.excluded, { suspectedTotal: [], duplicateRows: [] });
  assert.deepEqual(after.placeholders, {});
  assert.deepEqual(after.types, {});
  assert.deepEqual(after.drafts, after.inferred);
  assert.equal(after.problems, null);
  assert.equal(after.standardized, null);
  assert.equal(after.custom, null);
  assert.equal(after.module, "overviewSection");
  assert.equal(after.previewPage, 1);
});

test("v3.2.1 commit 2: a sheet that fails leaves the current analysis as it was, the error under the list in both languages", async () => {
  const { scope, reads } = await v321ImportSheet("Orders");
  v321ApplyScenario(scope);
  const before = v321Snapshot(scope);
  const kept = scope.evaluate("state.keptExcel");
  const importId = scope.evaluate("state.activeImportId");
  const filters = scope.evaluate("JSON.stringify(state.filters)");
  const errors = {
    Empty: ["工作表「Empty」为空，请选择包含表头和数据的工作表。", "The worksheet “Empty” is empty; choose one with a header row and data."],
    HeaderOnly: ["工作表「HeaderOnly」只有表头，没有可分析的数据行。", "The worksheet “HeaderOnly” has only a header row and no data to analyze."],
    Doubled: ["工作表「Doubled」包含重复表头「Amount」，请修改后重新上传。", "The worksheet “Doubled” has the header “Amount” more than once; change it and upload the file again."]
  };
  for (const [sheet, [zh, en]] of Object.entries(errors)) {
    const readsBefore = reads.length;
    await scope.evaluate(`switchExcelSheet(${JSON.stringify(sheet)}, "sidebar")`);
    assert.equal(reads.length, readsBefore + 1, sheet);
    assert.equal(textOf(scope, "sidebarSheetMessage"), zh);
    assert.match(scope.elements.get("sidebarSheetMessage").className, /\berror\b/);
    scope.evaluate('setInterfaceLanguage("en")');
    assert.equal(textOf(scope, "sidebarSheetMessage"), en);
    scope.evaluate('setInterfaceLanguage("zh")');
    assert.equal(textOf(scope, "sidebarSheetMessage"), zh);
    assert.equal(scope.evaluate("document.body.classList.contains('is-loading')"), false);
    assert.deepEqual([scope.elements.get("languageSwitch").disabled, scope.elements.get("sidebarLanguageSwitch").disabled], [false, false]);
    assert.equal(scope.evaluate("state.keptExcel") === kept, true);
    assert.equal(scope.evaluate("state.activeImportId"), importId);
    assert.equal(scope.evaluate("JSON.stringify(state.filters)"), filters);
    assert.deepEqual(v321WithoutSheetMessages(v321Snapshot(scope)), v321WithoutSheetMessages(before), sheet);
  }
  // From the upload page the error goes under that list.
  await scope.evaluate('switchExcelSheet("Empty", "upload")');
  assert.equal(textOf(scope, "uploadSheetMessage"), errors.Empty[0]);
  assert.deepEqual(v321WithoutSheetMessages(v321Snapshot(scope)), v321WithoutSheetMessages(before));
});

test("v3.2.1 commit 2: only a file with more than one sheet keeps its bytes, until the next import of any kind", async () => {
  const { scope } = await v321ImportSheet("Orders");
  const kept = JSON.parse(scope.evaluate("JSON.stringify({ keys: Object.keys(state.keptExcel).sort(), fileName: state.keptExcel.fileName, sheetNames: state.keptExcel.sheetNames, tag: Object.prototype.toString.call(state.keptExcel.bytes) })"));
  assert.deepEqual(kept, { keys: ["bytes", "fileName", "sheetNames"], fileName: "book.xlsx", sheetNames: ["Orders", "Second", "Empty", "HeaderOnly", "Doubled"], tag: "[object ArrayBuffer]" });
  scope.evaluate("globalThis.__keptBytes = state.keptExcel.bytes");
  // A switch, a failed switch: the same bytes.
  await scope.evaluate('switchExcelSheet("Second", "sidebar")');
  assert.equal(scope.evaluate("state.keptExcel.bytes === __keptBytes"), true);
  await scope.evaluate('switchExcelSheet("Empty", "sidebar")');
  assert.equal(scope.evaluate("state.keptExcel.bytes === __keptBytes"), true);
  // A CSV (the sample button does this too), a single-sheet file, an import error: gone.
  const again = async (action) => {
    const next = await v321ImportSheet("Orders");
    await action(next.scope);
    return next.scope.evaluate("state.keptExcel");
  };
  assert.equal(await again((other) => other.evaluate('parseCsvText("城市,人数\\n北京,1\\n上海,2", "new.csv", beginImport())')), null);
  assert.equal(await again(async (other) => { putWorkbookFile(other, buildXlsxBuffer([{ name: "S1", rows: V321_GOOD_ROWS }]), "one.xlsx"); await other.evaluate("parseDataFile(__workbookFile)"); }), null);
  assert.equal(await again(async (other) => { putWorkbookFile(other, Buffer.from("not a zip"), "bad.xlsx"); await other.evaluate("parseDataFile(__workbookFile)"); }), null);
  // A single-sheet file never keeps them.
  const single = createExcelContext();
  putWorkbookFile(single, buildXlsxBuffer([{ name: "S1", rows: V321_GOOD_ROWS }]), "one.xlsx");
  await single.evaluate("parseDataFile(__workbookFile)");
  assert.equal(single.evaluate("state.rows.length"), 3);
  assert.equal(single.evaluate("state.keptExcel"), null);
});

test("v3.2.1 commit 2: while a sheet is read the page is busy and the language locked; a switch reads only that sheet", async () => {
  const { scope, reads } = await v321ImportSheet("Orders");
  const readsBefore = reads.length;
  const switching = scope.evaluate('switchExcelSheet("Second", "sidebar")');
  assert.equal(scope.evaluate("document.body.classList.contains('is-loading')"), true);
  assert.deepEqual([scope.elements.get("languageSwitch").disabled, scope.elements.get("sidebarLanguageSwitch").disabled], [true, true]);
  assert.equal(textOf(scope, "sidebarSheetMessage"), "正在读取并分析工作表「Second」...");
  scope.evaluate("switchInterfaceLanguage()");
  assert.equal(scope.evaluate("currentLanguage"), "zh");
  // A click on another sheet meanwhile does nothing: the message stays the first one's.
  const second = scope.evaluate('switchExcelSheet("HeaderOnly", "sidebar")');
  assert.equal(textOf(scope, "sidebarSheetMessage"), "正在读取并分析工作表「Second」...");
  await second;
  await switching;
  assert.deepEqual(reads.slice(readsBefore), [{ type: "array", cellDates: true, cellNF: true, sheets: "Second" }]);
  assert.equal(scope.evaluate("state.sourceSheetName"), "Second");
  assert.equal(scope.evaluate("document.body.classList.contains('is-loading')"), false);
  assert.deepEqual([scope.elements.get("languageSwitch").disabled, scope.elements.get("sidebarLanguageSwitch").disabled], [false, false]);
  // The current sheet and a name not in the file: nothing.
  await scope.evaluate('switchExcelSheet("Second", "sidebar")');
  await scope.evaluate('switchExcelSheet("Nope", "sidebar")');
  assert.equal(reads.length, readsBefore + 1);
});

test("v3.2.1 commit 2: an import begun while a switch waits for its frame wins; the switch reads nothing", async () => {
  const { scope, reads } = await v321ImportSheet("Orders");
  const readsBefore = reads.length;
  scope.evaluate("globalThis.__realYield = yieldToBrowser; yieldToBrowser = () => new Promise((resolve) => { globalThis.__releaseYield = resolve; })");
  try {
    const switching = scope.evaluate('switchExcelSheet("Second", "sidebar")');
    scope.evaluate('parseCsvText("城市,人数\\n北京,1\\n上海,2", "new.csv", beginImport())');
    const status = textOf(scope, "statusPanel");
    scope.evaluate("__releaseYield()");
    await switching;
    assert.equal(reads.length, readsBefore);
    assert.equal(scope.evaluate("state.sourceFileName"), "new.csv");
    assert.equal(scope.evaluate("state.rows.length"), 2);
    assert.equal(scope.evaluate("state.keptExcel"), null);
    assert.equal(textOf(scope, "statusPanel"), status);
    assert.equal(textOf(scope, "sidebarSheetMessage"), "");
    assert.equal(scope.evaluate("document.body.classList.contains('is-loading')"), false);
  } finally {
    scope.evaluate("yieldToBrowser = __realYield");
  }
});

test("v3.2.1 commit 2: a reset clears the worksheet lists and their messages for good", async () => {
  const { scope } = await v321ImportSheet("Orders");
  await scope.evaluate('switchExcelSheet("Empty", "sidebar")');
  await scope.evaluate('switchExcelSheet("Empty", "upload")');
  assert.notEqual(textOf(scope, "sidebarSheetMessage"), "");
  assert.notEqual(textOf(scope, "uploadSheetMessage"), "");
  scope.evaluate("beginImport()");
  for (const place of ["sidebar", "upload"]) {
    assert.equal(textOf(scope, `${place}SheetMessage`), "", place);
    assert.equal(scope.evaluate(`oneOffMessages.has(dom.${place}SheetMessage)`), false, place);
    assert.equal(scope.elements.get(`${place}SheetList`).innerHTML, "", place);
    assert.equal(scope.elements.get(`${place}SheetPanel`).classList.contains("hidden"), true, place);
    assert.equal(scope.elements.get(`${place}SheetToggle`).classList.contains("hidden"), true, place);
  }
  scope.evaluate('setInterfaceLanguage("en")');
  assert.equal(textOf(scope, "sidebarSheetMessage"), "");
  assert.equal(scope.evaluate("state.keptExcel"), null);
});

test("v3.2.1 commit 2: an unexpected error while the other sheet is committed ends the import with the error", async () => {
  const { scope } = await v321ImportSheet("Orders");
  scope.evaluate('countDuplicateRows = () => { throw new Error("commit broke"); }');
  let rejected = null;
  await scope.evaluate('switchExcelSheet("Second", "sidebar")').catch((error) => { rejected = error; });
  assert.equal(rejected, null);
  assert.equal(scope.evaluate("document.body.classList.contains('is-loading')"), false);
  assert.equal(scope.elements.get("csvFile").disabled, false);
  assert.equal(scope.elements.get("languageSwitch").disabled, false);
  assert.equal(textOf(scope, "statusPanel"), "commit broke");
  assert.match(scope.elements.get("statusPanel").className, /\berror\b/);
  assert.equal(scope.evaluate("state.rows.length"), 0);
  assert.equal(scope.evaluate("state.keptExcel"), null);
});

// ---- v3.2.1 commit 3: the worksheet list in the sidebar ----

// The fake page has no getAttribute / querySelector and records no focus; these give every fake
// element (they share one prototype) what the worksheet list needs, for one action:
// setAttribute / getAttribute keep the values (as withRecordedAttributes() does), focus() is
// recorded, and querySelector('[data-sheet-index="N"]') answers an item whose focus() is recorded
// with the list it came from. Restored afterwards.
async function withSheetListDom(action) {
  const proto = Object.getPrototypeOf(createScriptContext().evaluate("dom.languageSwitch"));
  const saved = { setAttribute: proto.setAttribute, getAttribute: proto.getAttribute, querySelector: proto.querySelector, focus: proto.focus };
  const focused = [];
  proto.setAttribute = function setAttribute(name, value) { (this.__attributes ??= {})[name] = String(value); };
  proto.getAttribute = function getAttribute(name) { return this.__attributes && name in this.__attributes ? this.__attributes[name] : null; };
  proto.focus = function focus() { focused.push({ element: this }); };
  proto.querySelector = function querySelector(selector) {
    const list = this;
    const index = /^\[data-sheet-index="(\d+)"\]$/.exec(selector)?.[1];
    if (index === undefined || !String(list.innerHTML).includes(`data-sheet-index="${index}"`)) return null;
    return { focus() { focused.push({ list, item: Number(index) }); } };
  };
  try {
    return await action(focused);
  } finally {
    Object.assign(proto, saved);
  }
}

const EN_MULTI_SHEET = path.join(__dirname, "fixtures", "en-multi-sheet.xlsx");

test("v3.2.1 commit 3: a failed switch puts the focus back on the item that was clicked, in either list", async () => {
  await withSheetListDom(async (focused) => {
    for (const place of ["sidebar", "upload"]) {
      const { scope } = await v321ImportSheet("Orders");
      // Empty is the third sheet (index 2); the list draws it as a button.
      assert.match(scope.elements.get(`${place}SheetList`).innerHTML, /data-sheet-index="2"/, place);
      focused.length = 0;
      await scope.evaluate(`switchExcelSheet("Empty", "${place}")`);
      assert.equal(textOf(scope, `${place}SheetMessage`), "工作表「Empty」为空，请选择包含表头和数据的工作表。");
      assert.equal(focused.length, 1, place);
      assert.equal(focused[0].list, scope.elements.get(`${place}SheetList`), place);
      assert.equal(focused[0].item, 2, place);
    }
  });
});

test("v3.2.1 commit 3: with more than one sheet the sidebar names the file alone and the worksheet below it; one sheet as before", async () => {
  const { scope } = await v321ImportSheet("Orders");
  assert.equal(textOf(scope, "sourceName"), "book.xlsx");
  assert.equal(scope.elements.get("sourceName").title, "book.xlsx");
  assert.equal(textOf(scope, "sidebarSheetLine"), "工作表：Orders（共 5 个）");
  assert.equal(scope.elements.get("sidebarSheetLine").classList.contains("hidden"), false);
  assert.equal(scope.elements.get("sidebarSheetToggle").classList.contains("hidden"), false);
  // The top bar, the import message and the report keep “file · sheet”.
  assert.equal(textOf(scope, "topbarFileName"), "book.xlsx · Orders");
  assert.match(textOf(scope, "statusPanel"), /^book\.xlsx · Orders 已成功载入/);
  assert.match(scope.evaluate(`buildHtmlReport(buildReportData(${V321_FIXED_TIME}), [])`), /book\.xlsx/);
  scope.evaluate('setInterfaceLanguage("en")');
  assert.equal(textOf(scope, "sidebarSheetLine"), "Worksheet: Orders (5 in this file)");
  scope.evaluate('setInterfaceLanguage("zh")');
  await scope.evaluate('switchExcelSheet("Second", "sidebar")');
  assert.equal(textOf(scope, "sourceName"), "book.xlsx");
  assert.equal(textOf(scope, "sidebarSheetLine"), "工作表：Second（共 5 个）");
  // A new import of any kind takes the line away at once (before anything is drawn again).
  scope.evaluate("beginImport()");
  assert.equal(textOf(scope, "sidebarSheetLine"), "");
  assert.equal(scope.elements.get("sidebarSheetLine").classList.contains("hidden"), true);
  // A single sheet: “file · sheet” and no worksheet line or button.
  const single = createExcelContext();
  putWorkbookFile(single, buildXlsxBuffer([{ name: "S1", rows: V321_GOOD_ROWS }]), "one.xlsx");
  await single.evaluate("parseDataFile(__workbookFile)");
  assert.equal(textOf(single, "sourceName"), "one.xlsx · S1");
  assert.equal(textOf(single, "sidebarSheetLine"), "");
  assert.equal(single.elements.get("sidebarSheetLine").classList.contains("hidden"), true);
  assert.equal(single.elements.get("sidebarSheetToggle").classList.contains("hidden"), true);
  assert.equal(single.elements.get("sidebarSheetList").innerHTML, "");
});

test("v3.2.1 commit 3: the list marks the sheet open now, makes the others buttons and escapes the names", async () => {
  const buffer = buildXlsxBuffer([{ name: "a<b>&c", rows: V321_GOOD_ROWS }, { name: "订单", rows: V321_GOOD_ROWS }, { name: "说明", rows: [["备注"]] }]);
  const { scope } = await v321ImportSheet("a<b>&c", { buffer });
  const list = scope.elements.get("sidebarSheetList").innerHTML;
  assert.equal(list, [
    '<li class="sheet-switch-current" aria-current="true"><span class="sheet-switch-name">a&lt;b&gt;&amp;c</span><span class="sheet-switch-now">正在分析</span></li>',
    '<li><button class="sheet-switch-item" type="button" data-sheet-index="1"><span class="sheet-switch-name">订单</span></button></li>',
    '<li><button class="sheet-switch-item" type="button" data-sheet-index="2"><span class="sheet-switch-name">说明</span></button></li>'
  ].join(""));
  assert.equal(scope.elements.get("uploadSheetList").innerHTML, list);
  scope.evaluate('setInterfaceLanguage("en")');
  assert.match(scope.elements.get("sidebarSheetList").innerHTML, /<span class="sheet-switch-now">Current<\/span>/);
  scope.evaluate('setInterfaceLanguage("zh")');
  // A click on an item switches to it (the index, not the name, is in the markup).
  await scope.evaluate('handleSheetSwitchClick({ target: { closest: () => ({ dataset: { sheetIndex: "1" } }) } }, "sidebar")');
  assert.equal(scope.evaluate("state.sourceSheetName"), "订单");
  assert.match(scope.elements.get("sidebarSheetList").innerHTML, /^<li><button[^>]*data-sheet-index="0"><span class="sheet-switch-name">a&lt;b&gt;&amp;c<\/span><\/button><\/li><li class="sheet-switch-current" aria-current="true"><span class="sheet-switch-name">订单<\/span>/);
  // Clicks on the current sheet's line or around the items do nothing.
  await scope.evaluate('handleSheetSwitchClick({ target: { closest: () => null } }, "sidebar")');
  assert.equal(scope.evaluate("state.sourceSheetName"), "订单");
});

test("v3.2.1 commit 3: the button opens and closes the list; Escape closes it, gives the focus back and stops there", async () => {
  await withSheetListDom(async (focused) => {
    const { scope } = await v321ImportSheet("Orders");
    const toggle = scope.elements.get("sidebarSheetToggle");
    const panel = scope.elements.get("sidebarSheetPanel");
    assert.equal(toggle.getAttribute("aria-expanded"), "false");
    assert.equal(panel.classList.contains("hidden"), true);
    scope.evaluate('toggleSheetSwitch("sidebar")');
    assert.deepEqual([toggle.getAttribute("aria-expanded"), panel.classList.contains("hidden")], ["true", false]);
    const event = { key: "Escape", stopped: 0, prevented: 0, stopPropagation() { this.stopped += 1; }, preventDefault() { this.prevented += 1; } };
    put(scope.context, "__escape", event);
    focused.length = 0;
    scope.evaluate('handleSheetSwitchKeydown(__escape, "sidebar")');
    assert.deepEqual([toggle.getAttribute("aria-expanded"), panel.classList.contains("hidden")], ["false", true]);
    assert.deepEqual([event.stopped, event.prevented], [1, 1]);
    assert.equal(focused.length, 1);
    assert.equal(focused[0].element, toggle);
    // Closed: Escape is left to the sidebar (the overlay closes on the next one).
    scope.evaluate('handleSheetSwitchKeydown(__escape, "sidebar")');
    assert.deepEqual([event.stopped, event.prevented], [1, 1]);
    // Other keys do nothing.
    scope.evaluate('toggleSheetSwitch("sidebar")');
    put(scope.context, "__enter", { key: "Enter", stopPropagation() { throw new Error("stopped"); }, preventDefault() { throw new Error("prevented"); } });
    scope.evaluate('handleSheetSwitchKeydown(__enter, "sidebar")');
    assert.equal(toggle.getAttribute("aria-expanded"), "true");
    // The list stays open across a language switch and a new render of the sidebar.
    scope.evaluate('setInterfaceLanguage("en"); renderSidebar()');
    assert.equal(toggle.getAttribute("aria-expanded"), "true");
    assert.equal(panel.classList.contains("hidden"), false);
  });
});

test("v3.2.1 commit 3: the English multi-sheet fixture is what its generator writes", () => {
  const bytes = fs.readFileSync(EN_MULTI_SHEET);
  assert.equal(crypto.createHash("sha256").update(bytes).digest("hex"), "1e7e1e61b264d08a95bf933914b5bf5c26e6554d7ce3a6554083675bd77eb32f");
  const folder = fs.mkdtempSync(path.join(require("node:os").tmpdir(), "en-multi-sheet-"));
  try {
    require("node:child_process").execFileSync(process.execPath, [path.join(__dirname, "fixtures", "gen-en-multi-sheet.cjs"), folder]);
    assert.deepEqual(fs.readFileSync(path.join(folder, "en-multi-sheet.xlsx")), bytes);
  } finally {
    fs.rmSync(folder, { recursive: true, force: true });
  }
  const names = XLSX.read(bytes, { type: "buffer", bookSheets: true }).SheetNames;
  assert.deepEqual(names, ["Orders", "Customers", "Notes"]);
});

test("v3.2.1 commit 3: an English file in English shows no Chinese in the worksheet list, and Chinese → English → Chinese gives it back", async () => {
  await withSheetListDom(async () => {
    const buffer = fs.readFileSync(EN_MULTI_SHEET);
    const { scope } = await v321ImportSheet("Orders", { buffer });
    scope.evaluate('setSheetSwitchOpen("sidebar", true)');
    await scope.evaluate('switchExcelSheet("Notes", "sidebar")');
    assert.equal(textOf(scope, "sidebarSheetMessage"), "工作表「Notes」只有表头，没有可分析的数据行。");
    const before = elementTexts(scope);
    scope.evaluate('setInterfaceLanguage("en")');
    for (const id of ["sidebarSheetLine", "sidebarSheetList", "sidebarSheetMessage", "sourceName"]) {
      const element = scope.elements.get(id);
      assert.doesNotMatch(`${element.textContent} ${element.innerHTML}`, ENGLISH_BAD, id);
    }
    assert.equal(textOf(scope, "sidebarSheetLine"), "Worksheet: Orders (3 in this file)");
    assert.equal(textOf(scope, "sidebarSheetMessage"), "The worksheet “Notes” has only a header row and no data to analyze.");
    assert.match(scope.elements.get("sidebarSheetList").innerHTML, /Current/);
    assert.equal(scope.elements.get("sidebarSheetToggle").getAttribute("aria-expanded"), "true");
    scope.evaluate('setInterfaceLanguage("zh")');
    const after = elementTexts(scope);
    const changed = Object.keys(before).filter((id) => before[id] !== after[id] && !["languageSwitch", "sidebarLanguageSwitch"].includes(id));
    assert.deepEqual(changed, []);
  });
});

test("v3.2.1 commit 3: the sidebar's markup: the note above the list, every text with an English key", () => {
  // (The data-i18n keys themselves are checked by the static guard: every Chinese text of the page
  // has a key and every key an English text.)
  const html = indexHtmlSource();
  const sidebar = html.slice(html.indexOf('<aside id="appSidebar"'), html.indexOf("</aside>"));
  assertInOrder(sidebar, ['<p id="sourceName"', '<p id="sidebarSheetLine"', '<p id="sidebarFileSize"', '<div class="sidebar-file-actions">', '<button id="sidebarSheetToggle"', 'class="nav-new-analysis secondary-button sidebar-new-analysis"', '<div id="sidebarSheetPanel"', '<p class="sheet-switch-note"', '<ul id="sidebarSheetList"', '<p id="sidebarSheetMessage"']);
  assert.match(sidebar, /<button id="sidebarSheetToggle" class="secondary-button sheet-switch-toggle hidden" type="button" aria-expanded="false" aria-controls="sidebarSheetPanel"><span>换一个工作表<\/span><svg [^>]*aria-hidden="true">/);
  assert.ok(sidebar.includes('<ul id="sidebarSheetList" class="sheet-switch-list" aria-label="本文件的工作表"></ul>'));
  assert.ok(sidebar.includes('<p class="sheet-switch-note">换表会重新分析那张表，现在的筛选、不计入分析、按空值处理和改过的字段类型都不会保留。</p>'));
  assert.ok(sidebar.includes('<p id="sidebarSheetMessage" class="sheet-switch-message" role="status" aria-live="polite"></p>'));
  const english = JSON.parse(createScriptContext().evaluate("JSON.stringify(I18N_EN)"));
  assert.equal(english["sheetSwitch.toggle"], "Change worksheet");
  assert.equal(english["sheetSwitch.list"], "Worksheets in this file");
  assert.equal(english["sheetSwitch.note"], "Switching analyzes that worksheet from scratch: the filter, excluded rows, cells treated as blank and column types you changed here are not kept.");
  // R6: on phones the list does not scroll inside the scrolling sidebar.
  const css = repoFile("style.css");
  const phone = css.slice(css.indexOf("@media (max-width: 640px)"));
  assert.match(phone.slice(0, phone.indexOf("\n}\n")), /\.sheet-switch-list \{\s*max-height: none;/);
});

// ---- v3.2.1 commit 4: the way back to the analysis follows the language (R8) ----

test("v3.2.1 commit 4: 回到「…」的分析 on the upload page is written again in the other language, and back byte for byte", async () => {
  for (const [file, name] of [["csv", "orders.csv"], ["xlsx", "book.xlsx · Orders"]]) {
    const scope = file === "csv"
      ? await importCsvInLanguage(englishScenarioCsv(), "orders.csv", "zh")
      : (await v321ImportSheet("Orders")).scope;
    scope.evaluate("showUploadPage()");
    const chinese = textOf(scope, "returnToAnalysis");
    assert.equal(chinese, `回到「${name}」的分析`);
    scope.evaluate('setInterfaceLanguage("en")');
    assert.equal(textOf(scope, "returnToAnalysis"), `Back to the analysis of “${name}”`);
    scope.evaluate('setInterfaceLanguage("zh")');
    assert.equal(textOf(scope, "returnToAnalysis"), chinese);
    // Back to the analysis: the entry is gone, and a later switch does not write it back.
    scope.evaluate("returnToAnalysis()");
    assert.equal(textOf(scope, "returnToAnalysis"), "");
    scope.evaluate('setInterfaceLanguage("en")');
    assert.equal(textOf(scope, "returnToAnalysis"), "");
    scope.evaluate('setInterfaceLanguage("zh")');
    // A new import clears it too, for good.
    scope.evaluate("showUploadPage(); beginImport()");
    assert.equal(textOf(scope, "returnToAnalysis"), "");
    scope.evaluate('setInterfaceLanguage("en")');
    assert.equal(textOf(scope, "returnToAnalysis"), "");
  }
});

test("v3.2.1 commit 4: an upload page opened in English names the way back in English, and a switch writes it in Chinese", async () => {
  const scope = await importCsvInLanguage(englishScenarioCsv(), "orders.csv", "en");
  scope.evaluate("showUploadPage()");
  assert.equal(textOf(scope, "returnToAnalysis"), "Back to the analysis of “orders.csv”");
  scope.evaluate('setInterfaceLanguage("zh")');
  assert.equal(textOf(scope, "returnToAnalysis"), "回到「orders.csv」的分析");
});

// ---- v3.2.1 commit 5: long names wrap in the way back to the analysis (R9) ----

// The fake page has no layout: the browser check (no sideways scroll at 390 / 640 / 768 with long
// names, short names as before) is in the commit's notes; this pins the rule that makes it so.
test("v3.2.1 commit 5: 回到「…」的分析 may wrap, at any character, and never grows wider than the page", () => {
  const css = repoFile("style.css");
  // (Two classes: the shared button rule, later in the file, would win over one.)
  const start = css.indexOf("\n.upload-hero .return-to-analysis {");
  assert.ok(start >= 0);
  const rule = css.slice(start, css.indexOf("}", start) + 1);
  assert.match(rule, /\bmax-width: 100%;/);
  assert.match(rule, /\bwhite-space: normal;/);
  assert.match(rule, /\boverflow-wrap: anywhere;/);
  // Only this button: the shared button rule still keeps every other button on one line, and the
  // font is the buttons' own (no size or weight here).
  assert.match(css.slice(css.indexOf("\n.secondary-button {"), css.indexOf("}", css.indexOf("\n.secondary-button {"))), /white-space: nowrap;/);
  assert.doesNotMatch(rule, /font-(?:size|weight)/);
});

// ---- v3.2.1 commit 6: the worksheet list on the upload page ----

const v321ReportHash = (scope) => crypto.createHash("sha256").update(scope.evaluate(`buildHtmlReport(buildReportData(${V321_FIXED_TIME}), [])`)).digest("hex");
const v321UploadSwitch = (scope) => ({
  toggleHidden: scope.elements.get("uploadSheetToggle").classList.contains("hidden"),
  expanded: scope.elements.get("uploadSheetToggle").getAttribute("aria-expanded"),
  panelHidden: scope.elements.get("uploadSheetPanel").classList.contains("hidden")
});

test("v3.2.1 commit 6 (T14): the upload page's worksheet button comes with 回到「…」的分析 and goes with it", async () => {
  await withSheetListDom(async () => {
    const { scope } = await v321ImportSheet("Orders");
    const closed = { toggleHidden: false, expanded: "false", panelHidden: true };
    // In the analysis it is not there (the sidebar has its own).
    assert.equal(scope.elements.get("uploadSheetToggle").classList.contains("hidden"), true);
    assert.equal(scope.elements.get("sidebarSheetToggle").classList.contains("hidden"), false);
    scope.evaluate("showUploadPage()");
    assert.equal(scope.elements.get("returnToAnalysis").classList.contains("hidden"), false);
    assert.deepEqual(v321UploadSwitch(scope), closed);
    assert.match(scope.elements.get("uploadSheetList").innerHTML, /data-sheet-index="1"><span class="sheet-switch-name">Second</);
    scope.evaluate('toggleSheetSwitch("upload")');
    assert.deepEqual(v321UploadSwitch(scope), { toggleHidden: false, expanded: "true", panelHidden: false });
    // A language switch keeps it there and open.
    scope.evaluate('setInterfaceLanguage("en")');
    assert.deepEqual(v321UploadSwitch(scope), { toggleHidden: false, expanded: "true", panelHidden: false });
    scope.evaluate('setInterfaceLanguage("zh")');
    // Back to the analysis: gone and closed; the next upload page shows it closed.
    scope.evaluate("returnToAnalysis()");
    assert.deepEqual(v321UploadSwitch(scope), { toggleHidden: true, expanded: "false", panelHidden: true });
    assert.equal(scope.elements.get("sidebarSheetToggle").classList.contains("hidden"), false);
    scope.evaluate("showUploadPage()");
    assert.deepEqual(v321UploadSwitch(scope), closed);
    // No way back, no button: a sheet being chosen, an import going on.
    scope.evaluate("returnToAnalysis(); state.pendingExcel = { fileName: \"c.xlsx\" }; showUploadPage()");
    assert.deepEqual([scope.elements.get("returnToAnalysis").classList.contains("hidden"), v321UploadSwitch(scope).toggleHidden], [true, true]);
    scope.evaluate("state.pendingExcel = null; showUploadPage(); returnToAnalysis(); document.body.classList.add(\"is-loading\"); showUploadPage()");
    assert.deepEqual([scope.elements.get("returnToAnalysis").classList.contains("hidden"), v321UploadSwitch(scope).toggleHidden], [true, true]);

    // A CSV and a single-sheet file: the way back, no worksheet button.
    const csv = await importCsvInLanguage(englishScenarioCsv(), "orders.csv", "zh");
    const single = createExcelContext();
    putWorkbookFile(single, buildXlsxBuffer([{ name: "S1", rows: V321_GOOD_ROWS }]), "one.xlsx");
    await single.evaluate("parseDataFile(__workbookFile)");
    for (const other of [csv, single]) {
      other.evaluate("showUploadPage()");
      assert.equal(other.elements.get("returnToAnalysis").classList.contains("hidden"), false);
      assert.equal(other.elements.get("uploadSheetToggle").classList.contains("hidden"), true);
      assert.equal(other.elements.get("uploadSheetList").innerHTML, "");
    }
  });
});

test("v3.2.1 commit 6: a failed switch on the upload page stays there; the way back goes to the same sheet, its report unchanged", async () => {
  await withSheetListDom(async () => {
    const { scope } = await v321ImportSheet("Orders");
    v321ApplyScenario(scope);
    scope.evaluate('showModule("qualitySection")');
    const before = v321Snapshot(scope);
    const hash = v321ReportHash(scope);
    scope.evaluate("showUploadPage()");
    const way = textOf(scope, "returnToAnalysis");
    assert.equal(way, "回到「book.xlsx · Orders」的分析");
    scope.evaluate('toggleSheetSwitch("upload")');
    // Empty is the third sheet (index 2).
    await scope.evaluate('handleSheetSwitchClick({ target: { closest: () => ({ dataset: { sheetIndex: "2" } }) } }, "upload")');
    assert.equal(textOf(scope, "uploadSheetMessage"), "工作表「Empty」为空，请选择包含表头和数据的工作表。");
    assert.equal(textOf(scope, "sidebarSheetMessage"), "");
    // Still the upload page, the way back as it was, nothing busy.
    assert.equal(scope.evaluate("document.body.classList.contains('has-results')"), false);
    assert.equal(scope.elements.get("results").classList.contains("hidden"), true);
    assert.equal(scope.elements.get("returnToAnalysis").classList.contains("hidden"), false);
    assert.equal(textOf(scope, "returnToAnalysis"), way);
    assert.deepEqual(v321UploadSwitch(scope), { toggleHidden: false, expanded: "true", panelHidden: false });
    assert.equal(scope.evaluate("document.body.classList.contains('is-loading')"), false);
    assert.equal(scope.elements.get("csvFile").disabled, false);
    // The way back: the same sheet, the same module, the same report; the list's error is gone.
    scope.evaluate("returnToAnalysis()");
    assert.equal(scope.evaluate("document.body.classList.contains('has-results')"), true);
    assert.equal(scope.evaluate("state.sourceSheetName"), "Orders");
    assert.equal(v321ReportHash(scope), hash);
    assert.deepEqual(v321WithoutSheetMessages(v321Snapshot(scope)), v321WithoutSheetMessages(before));
    assert.equal(textOf(scope, "uploadSheetMessage"), "");
    assert.equal(scope.evaluate("oneOffMessages.has(dom.uploadSheetMessage)"), false);
    assert.deepEqual(v321UploadSwitch(scope), { toggleHidden: true, expanded: "false", panelHidden: true });
  });
});

test("v3.2.1 commit 6: a switch from the upload page opens the other sheet's analysis, as importing it does", async () => {
  await withSheetListDom(async () => {
    const { scope } = await v321ImportSheet("Orders");
    v321ApplyScenario(scope);
    scope.evaluate('showUploadPage(); toggleSheetSwitch("upload")');
    await scope.evaluate('handleSheetSwitchClick({ target: { closest: () => ({ dataset: { sheetIndex: "1" } }) } }, "upload")');
    assert.equal(scope.evaluate("document.body.classList.contains('has-results')"), true);
    assert.equal(scope.elements.get("returnToAnalysis").classList.contains("hidden"), true);
    assert.deepEqual(v321UploadSwitch(scope), { toggleHidden: true, expanded: "false", panelHidden: true });
    const fresh = await v321ImportSheet("Second");
    assert.deepEqual(v321Snapshot(scope), v321Snapshot(fresh.scope));
  });
});

test("v3.2.1 commit 6: the upload page's markup is the sidebar's, next to 回到「…」的分析; a wrapped name is left-aligned", () => {
  const html = indexHtmlSource();
  const hero = html.slice(html.indexOf('<div class="upload-hero">'), html.indexOf('<div class="workspace-intro">'));
  const sidebar = html.slice(html.indexOf('<aside id="appSidebar"'), html.indexOf("</aside>"));
  assertInOrder(hero, ['<div class="return-row">', '<button id="returnToAnalysis"', '<button id="uploadSheetToggle"', "</div>", '<div id="uploadSheetPanel"']);
  // (Indentation between the tags differs: the upload page's is two levels deeper.)
  const part = (source, start, end) => source.slice(source.indexOf(start), source.indexOf(end, source.indexOf(start)) + end.length).replace(/>\s+</g, "><");
  for (const [start, end] of [['<button id="{}SheetToggle"', "</button>"], ['<div id="{}SheetPanel"', "</div>"]]) {
    const upload = part(hero, start.replace("{}", "upload"), end);
    assert.ok(upload.length > end.length, start);
    assert.equal(upload.replace(/upload(?=Sheet)/g, "sidebar"), part(sidebar, start.replace("{}", "sidebar"), end), start);
  }
  // Short names stay centred on one line; a wrapped long one reads from the left.
  const css = repoFile("style.css");
  const start = css.indexOf("\n.upload-hero .return-to-analysis {");
  assert.match(css.slice(start, css.indexOf("}", start)), /\btext-align: left;/);
});

// ---- v3.2.1 commit 7: export names carry the worksheet (R2, R2b) ----

const V321_NAME_TIME = "new Date(2026, 10, 23, 14, 30, 5)";
// The five exports (the changes list also with its first-N cut, the problem rows with and without
// a column and a writing): [report, data, standardized, changes, changes cut, problem rows, duplicate rows].
const V321_EXPORT_NAMES = `[
  buildReportFileName("html", ${V321_NAME_TIME}),
  buildDataExportFileName("filtered", ${V321_NAME_TIME}),
  buildStandardizedFileName("data", ${V321_NAME_TIME}),
  buildStandardizedFileName("changes", ${V321_NAME_TIME}),
  buildStandardizedFileName("changes", ${V321_NAME_TIME}, 1048575),
  buildProblemRowsFileName({ kind: "placeholder", field: "金额", key: "N/A" }, ${V321_NAME_TIME}),
  buildProblemRowsFileName({ kind: "duplicateRows" }, ${V321_NAME_TIME})
]`;

// The export names of one source in one language. sheetNames: the worksheets of a kept file (more
// than one), or null for a CSV and a single-sheet file (nothing is kept).
function v321ExportNames(scope, fileName, sheetName, sheetNames, language) {
  put(scope.context, "__nameSource", { fileName, sheetName, sheetNames });
  return JSON.parse(scope.evaluate(`(() => {
    state.sourceFileName = __nameSource.fileName;
    state.sourceSheetName = __nameSource.sheetName;
    state.keptExcel = __nameSource.sheetNames ? { bytes: new ArrayBuffer(0), fileName: __nameSource.fileName, sheetNames: __nameSource.sheetNames } : null;
    currentLanguage = ${JSON.stringify(language)};
    const names = ${V321_EXPORT_NAMES};
    currentLanguage = "zh";
    return JSON.stringify(names);
  })()`));
}

test("v3.2.1 commit 7 (T8): with more than one worksheet the five exports carry it after the file name, in both languages", () => {
  const scope = createScriptContext();
  assert.deepEqual(v321ExportNames(scope, "2025上半年订单.xlsx", "订单", ["订单", "退货"], "zh"), [
    "2025上半年订单_订单-20261123-143005.html",
    "2025上半年订单_订单_当前分析_20261123.csv",
    "2025上半年订单_订单_写法统一_20261123.csv",
    "2025上半年订单_订单_写法统一改动清单_20261123.csv",
    "2025上半年订单_订单_写法统一改动清单_前1048575条_20261123.csv",
    "2025上半年订单_订单_问题行_像是空值的写法_金额_N-A_20261123.csv",
    "2025上半年订单_订单_问题行_整行重复_20261123.csv"
  ]);
  assert.deepEqual(v321ExportNames(scope, "2025上半年订单.xlsx", "订单", ["订单", "退货"], "en"), [
    "2025上半年订单_订单-20261123-143005.html",
    "2025上半年订单_订单_in-analysis_20261123.csv",
    "2025上半年订单_订单_standardized_20261123.csv",
    "2025上半年订单_订单_standardized-changes_20261123.csv",
    "2025上半年订单_订单_standardized-changes_first-1048575_20261123.csv",
    "2025上半年订单_订单_problem-rows_blank-like_金额_N-A_20261123.csv",
    "2025上半年订单_订单_problem-rows_duplicate-rows_20261123.csv"
  ]);
  // .xls too; unsafe characters in a worksheet name are made safe like those in a file name.
  assert.equal(v321ExportNames(scope, "2025上半年订单.xls", "一月", ["一月", "二月"], "zh")[1], "2025上半年订单_一月_当前分析_20261123.csv");
  assert.equal(v321ExportNames(scope, "book.xlsx", "a<b>c", ["a<b>c", "x"], "zh")[1], "book_a-b-c_当前分析_20261123.csv");
  // The same file with one worksheet: no worksheet in the names (as in 7b6e091).
  assert.deepEqual(v321ExportNames(scope, "2025上半年订单.xlsx", "订单", null, "zh").slice(0, 2),
    ["2025上半年订单-20261123-143005.html", "2025上半年订单_当前分析_20261123.csv"]);
});

test("v3.2.1 commit 7 (T8): a real multi-sheet import names the sheet it analyzes, and the one it switched to", async () => {
  const { scope } = await v321ImportSheet("Orders");
  assert.equal(scope.evaluate(`buildDataExportFileName("filtered", ${V321_NAME_TIME})`), "book_Orders_当前分析_20261123.csv");
  await scope.evaluate('switchExcelSheet("Second", "sidebar")');
  assert.equal(scope.evaluate(`buildDataExportFileName("filtered", ${V321_NAME_TIME})`), "book_Second_当前分析_20261123.csv");
  assert.equal(scope.evaluate(`buildReportFileName("html", ${V321_NAME_TIME})`), "book_Second-20261123-143005.html");
});

test("v3.2.1 commit 7 (T8): a 70-character file name and a 30-character worksheet are both cut, longest first, the labels whole, 80 at most", () => {
  const scope = createScriptContext();
  const file = `${"名".repeat(70)}.xlsx`;
  const sheet = "表".repeat(30);
  const zh = v321ExportNames(scope, file, sheet, [sheet, "另一张"], "zh");
  const en = v321ExportNames(scope, file, sheet, [sheet, "另一张"], "en");
  // Report: the file name loses 21 (49 + "_" + 30 = 80); the data export also makes room for its label.
  assert.equal(zh[0], `${"名".repeat(49)}_${sheet}-20261123-143005.html`);
  assert.equal(zh[1], `${"名".repeat(44)}_${sheet}_当前分析_20261123.csv`);
  assert.equal(en[1], `${"名".repeat(37)}_${sheet}_in-analysis_20261123.csv`);
  // Long labels (70 + 30 + 20 + 13 + 3 "_" = 136, 56 to cut): the file name comes down to the
  // worksheet's 30, then the two lose 8 each.
  assert.equal(en[4], `${"名".repeat(22)}_${"表".repeat(22)}_standardized-changes_first-1048575_20261123.csv`);
  for (const name of [...zh, ...en]) {
    const before = name.replace(/(?:-\d{8}-\d{6}\.html|_\d{8}\.csv)$/, "");
    assert.ok(before.length <= 80, `${before.length} ${name}`);
    assert.match(before, /^名+_表+_?/, name);
  }
});

test("v3.2.1 commit 7 (R2b): the worksheet part is trimmed; a worksheet of spaces adds nothing", () => {
  const scope = createScriptContext();
  assert.deepEqual(v321ExportNames(scope, "book.xlsx", " 订单 ", [" 订单 ", "x"], "zh").slice(0, 3),
    ["book_订单-20261123-143005.html", "book_订单_当前分析_20261123.csv", "book_订单_写法统一_20261123.csv"]);
  // Spaces only: exactly the names of the same file with one worksheet, the file name part as today
  // (here with its own trailing space, which the report and the data export trim as they always did).
  for (const language of ["zh", "en"]) {
    for (const file of ["book.xlsx", " 前后空格 .xlsx", `${"名".repeat(81)}.xlsx`]) {
      assert.deepEqual(v321ExportNames(scope, file, "   ", ["   ", "x"], language), v321ExportNames(scope, file, "   ", null, language), `${file} ${language}`);
    }
  }
});

// R2: a CSV and a single-sheet Excel file keep 7b6e091's names byte for byte, over awkward names:
// spaces around, a name of spaces, unsafe characters, 70–90 characters (the report and the data
// export cut at 80, the others shorten the name), a space or an emoji at the 80th place, emoji,
// combining marks, a bidi control, a tab. (Built from code points: no invisible character here.)
const V321_AWKWARD_STEMS = [
  "订单", "orders", "a<b>c", `x/y${String.fromCharCode(92)}z`, " 前后空格 ", "trailing ", "", "   ", "...",
  "名".repeat(70), "名".repeat(75), "名".repeat(76), "名".repeat(80), "名".repeat(81), "名".repeat(90),
  `${"a".repeat(79)} b`, `${"a".repeat(79)}${String.fromCodePoint(0x1F600)}b`, String.fromCodePoint(0x1F600).repeat(45),
  String.fromCodePoint(0xE9).repeat(50), `${String.fromCodePoint(0x202E)}rtl`, `tab${String.fromCharCode(9)}here`
];
// JSON.stringify([[file, language, the seven names], …]) over the 21 names × .csv/.xlsx/.xls × zh/en,
// computed by running these same calls on `git archive 7b6e091` (v3.2.0).
const V321_7B6E091_NAMES_SHA256 = "8fbe080ede4364964e6afcc56d3045c2debdb9df16aae4076b0b97b0c1c4671a";

test("v3.2.1 commit 7 (R2): a CSV and a single-sheet file keep 7b6e091's export names byte for byte, over 21 awkward names", () => {
  const scope = createScriptContext();
  const all = [];
  for (const stem of V321_AWKWARD_STEMS) {
    for (const ext of [".csv", ".xlsx", ".xls"]) {
      for (const language of ["zh", "en"]) {
        const names = v321ExportNames(scope, stem + ext, ext === ".csv" ? "" : "表一", null, language);
        all.push([stem + ext, language, names]);
        // The report and the data export: 7b6e091's one sanitize-and-cut of the original name.
        const base = (stem + ext).split(/[\\/]/).pop().replace(/\.(?:csv|xlsx|xls)$/i, "");
        assert.equal(names[0], `${oldCutFileName(base || "data-report")}-20261123-143005.html`, JSON.stringify(stem + ext));
        assert.equal(names[1], `${oldCutFileName(base || "data")}_${language === "zh" ? "当前分析" : "in-analysis"}_20261123.csv`, JSON.stringify(stem + ext));
      }
    }
  }
  assert.equal(all.length, 126);
  // All 882 names, as 7b6e091 builds them (computed on a copy of 7b6e091, not on this code).
  assert.equal(crypto.createHash("sha256").update(JSON.stringify(all)).digest("hex"), V321_7B6E091_NAMES_SHA256);
});

// v3.2.2: the repository is called sheet-checkup. The page, both reports and package.json point
// to the new name only; the old repository name is joined from parts so that this file is not a
// hit when the repository is searched for it.
test("v3.2.2 repository: the page and both reports link to sheet-checkup only", async () => {
  const oldName = ["smart", "tabular", "analyzer"].join("-");
  const newUrl = "https://github.com/wy-data-30/sheet-checkup";
  const html = indexHtmlSource();
  assert.equal(html.toLowerCase().indexOf(oldName), -1);
  assert.equal((html.match(/<a class="footer-github-link" href="([^"]*)">/) || [])[1], newUrl);
  for (const lang of ["zh", "en"]) {
    const { html: report } = reportOf(await reportScenario(lang));
    assert.equal(report.toLowerCase().indexOf(oldName), -1, lang);
    assert.deepEqual((report.match(/<footer>[^<]*<a href="([^"]*)">/) || []).slice(1), [newUrl], lang);
  }
  assert.equal(JSON.parse(repoFile("package.json")).name, "sheet-checkup");
});
