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

function createExcelContext() {
  const scope = createScriptContext();
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
    importExcelSheet(readExcelWorkbook(__xlsxBuffer), "formats.xlsx", __xlsxSheet, importId);
  })()`);
}

function tableCells(elements, id) {
  return [...elements.get(id).innerHTML.matchAll(/<td>([^<]*)<\/td>/g)].map((match) => match[1]);
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
  assert.match(scope.elements.get("numericStatsTable").innerHTML, /<th>有效数量<\/th>\s*<th>认不出<\/th>/);
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
  assert.match(html.split("数值字段描述性统计")[1], /<td>金额<\/td><td>12<\/td><td>3<\/td>/);
  assert.match(html, /<td>3（&lt;b&gt;x&lt;\/b&gt;\|y、待定、约500）<\/td>/);
  assert.doesNotMatch(html, /<b>x<\/b>/);

  const markdown = scope.evaluate("buildMarkdownReport(__report)");
  assert.ok(markdown.includes("| 字段 | 有效数量 | 认不出 | 平均值 | 中位数 | 最小值 | 最大值 | 标准差 |"));
  // escapeMarkdownCell backslash-escapes ASCII punctuation, so "/" and "|" both carry a backslash.
  assert.ok(markdown.includes("3（&lt;b&gt;x&lt;\\/b&gt;\\|y、待定、约500）"));

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
  assert.doesNotThrow(() => scope.evaluate("buildMarkdownReport(__legacy)"));
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
    html: scope.evaluate("buildHtmlReport(__filterReport, [])"),
    markdown: scope.evaluate("buildMarkdownReport(__filterReport)")
  };
}

test("reports say when no filter is applied", () => {
  const scope = createChartContext();
  commitExampleCsv(scope, "sample-data.csv");
  const { report, html, markdown } = reportOf(scope);
  assert.deepEqual(report.filters, []);
  for (const text of [html, markdown]) {
    assert.ok(text.includes("筛选条件"));
    assert.ok(text.includes(NO_FILTER_SENTENCE));
    assert.ok(text.includes(QUALITY_SCOPE_SENTENCE));
  }
  assert.ok(html.indexOf("<h2>数据规模</h2>") < html.indexOf("<h2>筛选条件</h2>"));
  assert.ok(html.indexOf("<h2>筛选条件</h2>") < html.indexOf("<h2>字段类型结果</h2>"));
  assert.ok(markdown.indexOf("## 数据规模") < markdown.indexOf("## 筛选条件"));
  assert.ok(markdown.indexOf("## 筛选条件") < markdown.indexOf("## 字段类型结果"));
});

test("reports list applied category, numeric and date filters in order", () => {
  const scope = createChartContext();
  commitExampleCsv(scope, "sample-data.csv");
  applyFilterState(scope, {
    category: { field: "地区", values: ["华东", "华北", "华南"] },
    numeric: { field: "销售额", min: 1000, max: 5000 },
    date: { field: "订单日期", start: "2025-01-01", end: "2025-03-31" }
  });
  const { report, html, markdown } = reportOf(scope);
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
  assert.ok(markdown.includes("| 字段 | 条件 |"));
  assert.ok(markdown.includes("| 地区 | 地区：保留 3 \\/ 共 6 个类别（华北、华东、华南） |"));
  // escapeMarkdownCell backslash-escapes ASCII punctuation, including "," "/" and "-".
  assert.ok(markdown.includes("| 销售额 | 销售额：1\\,000 至 5\\,000 |"));
  assert.ok(markdown.includes("| 订单日期 | 订单日期：2025\\-01\\-01 至 2025\\-03\\-31 |"));
  assert.ok(markdown.includes(QUALITY_SCOPE_SENTENCE));
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

test("report filter descriptions are escaped in HTML and Markdown", () => {
  const scope = createScriptContext();
  commitRows(scope, ["备注"], ["<b>x</b>|y", "第一行\n第二行", "其他"].map((note) => ({ 备注: note })), "escape.csv");
  applyFilterState(scope, { category: { field: "备注", values: ["<b>x</b>|y", "第一行\n第二行"] } });
  const { html, markdown } = reportOf(scope);
  // The same values also appear as field samples, so look only at the filter section.
  const htmlSection = html.split("<h2>筛选条件</h2>")[1]?.split("<h2>字段类型结果</h2>")[0] ?? "";
  const markdownSection = markdown.split("## 筛选条件")[1]?.split("## 字段类型结果")[0] ?? "";
  assert.ok(htmlSection.includes("<td>备注</td><td>备注：保留 2 / 共 3 个类别（&lt;b&gt;x&lt;/b&gt;|y、第一行\n第二行）</td>"));
  assert.ok(!html.includes("<b>x</b>"));
  assert.ok(markdownSection.includes("| 备注 | 备注：保留 2 \\/ 共 3 个类别（&lt;b&gt;x&lt;\\/b&gt;\\|y、第一行<br>第二行） |"));
  assert.ok(!markdown.includes("<b>x</b>"));
});

test("legacy report snapshots without filters still render", () => {
  const scope = createScriptContext();
  assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(normalizeReportSnapshot({}).filters)")), []);
  assert.ok(scope.evaluate("buildHtmlReport({}, [])").includes(NO_FILTER_SENTENCE));
  assert.ok(scope.evaluate("buildMarkdownReport({})").includes(NO_FILTER_SENTENCE));
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
  const outlierCells = [...elements.get("outlierTable").innerHTML.matchAll(/<td>([^<]*)<\/td>/g)].map((match) => match[1]);
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

  const statsCells = [...elements.get("numericStatsTable").innerHTML.matchAll(/<td>([^<]*)<\/td>/g)].map((match) => match[1]);
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
  assert.equal(evaluate('getProcessedExportRows("deduplicated").length'), 2);
  assert.equal(JSON.stringify(rows), originalSnapshot);

  const csv = evaluate('buildCsvExportContent(getProcessedExportRows("deduplicated"), state.fields)');
  assert.equal(
    csv,
    '姓名,备注\r\n张三,"上海,重点客户"\r\n李四,"包含""引号"""'
  );
  assert.deepEqual(Array.from(context.__unparseInput.fields), ["姓名", "备注"]);
  assert.deepEqual(
    Array.from(context.__unparseInput.data, (row) => Array.from(row)),
    [["张三", "上海,重点客户"], ["李四", '包含"引号"']]
  );
  assert.equal(context.__unparseOptions.newline, "\r\n");

  assert.equal(
    evaluate('buildDataExportFileName("filtered", new Date(2026, 6, 18))'),
    "中文销售_filtered_20260718.csv"
  );
  assert.equal(
    evaluate('buildDataExportFileName("deduplicated", new Date(2026, 6, 18))'),
    "中文销售_deduplicated_20260718.csv"
  );

  const blob = evaluate('createUtf8BomBlob("姓名,城市\\r\\n张三,上海", "text/csv;charset=utf-8")');
  const bytes = new Uint8Array(await blob.arrayBuffer());
  assert.deepEqual(Array.from(bytes.slice(0, 3)), [0xef, 0xbb, 0xbf]);

  evaluate("updateExportButtons()");
  assert.equal(elements.get("dataExportOriginalCount").textContent, "3");
  assert.equal(elements.get("dataExportFilteredCount").textContent, "1");
  assert.equal(elements.get("dataExportDeduplicatedCount").textContent, "2");
  assert.equal(elements.get("exportFilteredCsv").disabled, false);
  assert.equal(elements.get("exportDeduplicatedCsv").disabled, false);

  evaluate("state.filteredRows = []");
  assert.equal(
    evaluate('buildCsvExportContent(getProcessedExportRows("filtered"), state.fields)'),
    "姓名,备注"
  );
});

// Stage 3 acceptance fixtures (tests/fixtures/stage3/, see known-issues.json)
// imported through parseDataFile() with the real PapaParse and SheetJS.
const STAGE3_FIXTURES = path.join(__dirname, "fixtures", "stage3");

async function importStage3Fixture(fileName) {
  const scope = createExcelContext();
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

async function importWorkbookFile(buffer, fileName = "t.xlsx") {
  const scope = createExcelContext();
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
    const sheet = readExcelWorkbook(__offsetBuffer).Sheets.S1;
    return readExcelSheetTable(sheet, "S1");
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
  assert.equal(store.elements.get("statusPanel").textContent, `门店报表.xlsx · 销售汇总 已成功载入，分析结果已生成。${STORE_HEADER_NOTE}第 33 行「合计」疑似合计行，目前仍计入统计，详情见“数据质量”。`);
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
  const { report, html, markdown } = reportOf(scope);
  assert.deepEqual(report.headerDetection, JSON.parse(scope.evaluate("JSON.stringify(state.headerDetection)")));
  assert.ok(html.includes(`<div><dt>表头位置</dt><dd>${scope.evaluate("escapeHtml(__headerNote)")}</dd></div>`));
  assert.ok(markdown.includes(`- 表头位置：${scope.evaluate("escapeMarkdownText(__headerNote)")}`));

  const orders = await importStage3Fixture("订单表.csv");
  assert.equal(orders.evaluate("state.headerDetection"), null);
  assert.doesNotMatch(orders.elements.get("statusPanel").textContent, /表头在第/);
  const after = reportOf(orders);
  assert.equal(after.report.headerDetection, null);
  assert.doesNotMatch(after.html, /表头位置/);
  assert.doesNotMatch(after.markdown, /表头位置/);

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
  assert.doesNotMatch(scope.evaluate("buildMarkdownReport({})"), /表头位置/);
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
  scope.evaluate("analyzeSelectedExcelSheet()");
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
  const table = JSON.parse(evaluate("JSON.stringify(readExcelSheetTable(readExcelWorkbook(__offsetBuffer).Sheets.S1, \"S1\"))"));
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
    assert.deepEqual(JSON.parse(scope.evaluate("JSON.stringify(state.parseWarnings)")), warnings, JSON.stringify(text));
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

const codedWarnings = (scope) => JSON.parse(scope.evaluate("JSON.stringify(state.parseWarnings.map(({ row, code, message }) => [row, code, message]))"));

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
  assert.deepEqual(JSON.parse(evaluate("JSON.stringify(state.parseWarnings)")), [
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
  const html = fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8");
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
  const { report, html, markdown } = reportOf(roster);
  assert.deepEqual(report.quality.idDuplicates, idDuplicates(roster.evaluate));
  assert.ok(html.includes("；读取时有疑问：<strong>0</strong> 处；编号重复：<strong>3</strong> 个字段；日期写法混用：<strong>1</strong> 个字段。</p>"));
  const htmlSection = html.slice(html.indexOf("<h3>编号重复</h3>"), html.indexOf("数值字段描述性统计"));
  assertInOrder(html, ["<h3>偏大或偏小的值</h3>", "<h3>编号重复</h3>"]);
  assert.match(htmlSection, /<th>字段<\/th><th>重复值<\/th><th>出现次数<\/th><th>文件行号<\/th><th>说明<\/th>/);
  assert.ok(htmlSection.includes("<td>工号</td><td>E1012</td><td>3</td><td>13、14、25</td><td>其中 2 行内容完全相同</td>"));

  assert.ok(markdown.includes("- 读取时有疑问：0 处\n- 编号重复：3 个字段\n"));
  assertInOrder(markdown, ["### 偏大或偏小的值", "### 编号重复"]);
  assertInOrder(markdown, ["### 编号重复", "## 数值字段描述性统计"]);
  assert.ok(markdown.includes("| 字段 | 重复值 | 出现次数 | 文件行号 | 说明 |"));
  assert.ok(markdown.includes("| 工号 | E1012 | 3 | 13、14、25 | 其中 2 行内容完全相同 |"));
});

test("report ID duplicates show mostly repeated fields, overflow, clean fields and no ID fields", async () => {
  const behavior = await importCsvFile(behaviorCsv(), "sample-user-behavior.csv");
  const behaviorReport = reportOf(behavior);
  put(behavior.context, "__hintItem", idDuplicates(behavior.evaluate)[1]);
  const hint = behavior.evaluate("formatMostlyRepeatedIdHint(__hintItem)");
  assert.ok(behaviorReport.html.includes(`<td>用户ID</td><td>—</td><td>—</td><td>—</td><td>${behavior.evaluate("escapeHtml(formatMostlyRepeatedIdHint(__hintItem))")}</td>`));
  assert.ok(behaviorReport.markdown.includes(`| 用户ID | — | — | — | ${behavior.evaluate(`escapeMarkdownCell(${JSON.stringify(hint)})`)} |`));
  assert.ok(behaviorReport.html.includes("<td>事件ID</td><td>EV-00026</td><td>2</td><td>27、120</td><td>其中 2 行内容完全相同</td>"));
  assert.ok(behaviorReport.html.includes("；编号重复：<strong>1</strong> 个字段。"));

  const ids = [
    ...Array.from({ length: 15 }, () => "S000"),
    ...Array.from({ length: 11 }, (_, index) => [`S${index + 1}`, `S${index + 1}`]).flat(),
    ...Array.from({ length: 40 }, (_, index) => `S${index + 100}`)
  ];
  const long = reportOf(await importCsvFile(`学号\r\n${ids.join("\r\n")}\r\n`));
  assert.ok(long.html.includes("<p>「学号」共 12 个重复值，只列出前 10 个</p>"));
  assert.ok(long.markdown.includes("「学号」共 12 个重复值，只列出前 10 个"));

  const grades = reportOf(await importStage3Fixture("成绩表.csv"));
  const gradeSection = grades.html.slice(grades.html.indexOf("<h3>编号重复</h3>"), grades.html.indexOf("数值字段描述性统计"));
  assert.ok(gradeSection.includes("<p>类型为 ID 的字段没有编号重复（已检查：学号、姓名）</p>"));
  assert.doesNotMatch(gradeSection, /<table>/);
  assert.ok(grades.markdown.includes("### 编号重复\n\n类型为 ID 的字段没有编号重复（已检查：学号、姓名）\n"));
  assert.ok(grades.html.includes("；编号重复：<strong>0</strong> 个字段。"));

  const none = createScriptContext();
  commitRows(none, ["城市"], [{ 城市: "上海" }, { 城市: "北京" }], "cities.csv");
  const noneReport = reportOf(none);
  assert.deepEqual(noneReport.report.quality.idDuplicates, []);
  assert.ok(noneReport.html.includes("<p>未识别到 ID 字段。</p>"));
  assert.ok(noneReport.markdown.includes("### 编号重复\n\n未识别到 ID 字段。\n"));
});

test("report ID duplicate values are escaped for HTML and Markdown", () => {
  const scope = createScriptContext();
  const value = "<b>x</b>|y\nz";
  commitRows(scope, ["编号", "姓名"], [{ 编号: value, 姓名: "甲" }, { 编号: value, 姓名: "乙" }], "escape.csv");
  const { html, markdown } = reportOf(scope);
  assert.ok(html.includes("<td>编号</td><td>&lt;b&gt;x&lt;/b&gt;|y\nz</td><td>2</td><td>—</td><td></td>"));
  assert.doesNotMatch(html, /<b>x<\/b>/);
  assert.ok(markdown.includes(`| 编号 | ${scope.evaluate(`escapeMarkdownCell(${JSON.stringify(value)})`)} | 2 | — |  |`));
  assert.ok(markdown.includes("&lt;b&gt;x&lt;\\/b&gt;\\|y<br>z"));
});

test("legacy report snapshots without ID duplicates leave the section out", () => {
  const scope = createScriptContext();
  assert.equal(scope.evaluate("normalizeReportSnapshot({}).quality.idDuplicates"), null);
  const legacy = { quality: { totalMissing: 1, duplicateRows: 0, parseWarningCount: 0, outliers: [{ field: "金额", outlierCount: 1, lowerBound: "1", upperBound: "9" }] } };
  put(scope.context, "__legacyQuality", legacy);
  const html = scope.evaluate("buildHtmlReport(__legacyQuality, [])");
  const markdown = scope.evaluate("buildMarkdownReport(__legacyQuality)");
  assert.doesNotMatch(html, /<h3>编号重复<\/h3>|；编号重复：/);
  assert.doesNotMatch(markdown, /### 编号重复|- 编号重复：/);
  assert.ok(html.includes("；读取时有疑问：<strong>0</strong> 处。</p>"));
  assert.ok(html.includes("<td>金额</td><td>1</td><td>1</td><td>9</td>"));
  assert.ok(markdown.includes("- 读取时有疑问：0 处\n\n### 各字段空值情况"));
  assert.ok(markdown.includes("| 金额 | 1 | 1 | 9 |"));
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
  assert.equal(examples.length, 7);
  for (const fileName of examples) {
    const scope = await importCsvFile(readExampleCsv(fileName), fileName);
    assert.ok(scope.evaluate("state.rows.length") > 0, fileName);
    const fieldsAndKeys = placeholderEntries(scope.evaluate).map(([field, key]) => [field, key]);
    assert.deepEqual(fieldsAndKeys, fileName === "sample-orders.csv" ? [["金额", "N/A"], ["销售代表", "\u2014"]] : [], fileName);
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
  assert.equal(placeholderMessage(scope), "已把「金额」的「N/A」按空值处理，统计、图表和分析结论已重新计算。 已清除现有筛选条件。");
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
  assert.equal(placeholderMessage(scope), "已把 2 处常见写法按空值处理，统计、图表和分析结论已重新计算。 已清除现有筛选条件。");
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
  assert.equal(placeholderMessage(orders), "已撤销「备注」的「—」，统计、图表和分析结论已重新计算。 已清除现有筛选条件。");
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
    filtered: buildCsvExportContent(getProcessedExportRows("filtered"), state.fields),
    deduplicated: buildCsvExportContent(getProcessedExportRows("deduplicated"), state.fields)
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
  const html = fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8");
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
  assert.ok(before.markdown.includes("- 编号重复：1 个字段\n- 像是空值的写法：4 个字段\n"));
  assertInOrder(before.markdown, ["### 各字段空值情况", "### 像是空值的写法"]);
  assertInOrder(before.markdown, ["### 像是空值的写法", "### 偏大或偏小的值"]);
  const cell = (value) => orders.evaluate(`escapeMarkdownCell(${JSON.stringify(value)})`);
  assert.ok(before.markdown.includes(`| 金额 | ${cell("「N/A」")} | 常见 | 1 | 15 | 按空值处理后，空值 0 → 1，类型仍为数值 |`));

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
  assert.ok(after.markdown.includes(`| 备注 | ${cell("「N/A」")} | 1 | 已按空值处理（字段已忽略） |`));
});

test("report placeholder summaries appear only with candidates, and legacy snapshots leave the section out", async () => {
  for (const scope of [await importStage3Fixture("花名册.xlsx"), await importStage3Fixture("成绩表.csv"), await importCsvFile(behaviorCsv(), "sample-user-behavior.csv")]) {
    const { report, html, markdown } = reportOf(scope);
    assert.deepEqual(report.quality.placeholders, { confirmed: [], pending: [] });
    assert.doesNotMatch(html, /像是空值的写法：/);
    assert.doesNotMatch(markdown, /像是空值的写法：/);
    assert.ok(html.includes("<p>没有发现像是空值的写法。</p>"));
    assert.ok(markdown.includes("### 像是空值的写法\n\n没有发现像是空值的写法。\n"));
  }
  const scope = createScriptContext();
  assert.equal(scope.evaluate("normalizeReportSnapshot({}).quality.placeholders"), null);
  assert.equal(scope.evaluate("normalizeReportSnapshot({ quality: { placeholders: { confirmed: [] } } }).quality.placeholders"), null);
  put(scope.context, "__legacyPlaceholders", { quality: { totalMissing: 1, duplicateRows: 0, parseWarningCount: 0 } });
  assert.doesNotMatch(scope.evaluate("buildHtmlReport(__legacyPlaceholders, [])"), /<h3>像是空值的写法<\/h3>|；像是空值的写法：/);
  assert.doesNotMatch(scope.evaluate("buildMarkdownReport(__legacyPlaceholders)"), /### 像是空值的写法|- 像是空值的写法：/);
});

test("report placeholder sections escape field names and keep Markdown lines from starting with user text", () => {
  const scope = commitColumns({ "<b>备注</b>|x": ["N/A", "好", "-", "好"], "-前缀": ["—", "甲", "乙", "丙"] });
  const { html, markdown } = reportOf(scope);
  const section = html.slice(html.indexOf("<h3>像是空值的写法</h3>"), html.indexOf("<h3>偏大或偏小的值</h3>"));
  assert.ok(section.includes("<td>&lt;b&gt;备注&lt;/b&gt;|x</td><td>「N/A」</td>"));
  assert.ok(section.includes("<td>-前缀</td><td>「—」</td>"));
  assert.doesNotMatch(html, /<b>备注<\/b>/);
  assert.ok(section.includes(`<p>说明：${PLACEHOLDER_MAYBE_NOTE_TEXT}在会计格式里，「-」通常表示 0。</p>`));
  const markdownSection = markdown.slice(markdown.indexOf("### 像是空值的写法"), markdown.indexOf("### 偏大或偏小的值"));
  const cell = (value) => scope.evaluate(`escapeMarkdownCell(${JSON.stringify(value)})`);
  assert.ok(markdownSection.includes(`| ${cell("<b>备注</b>|x")} | ${cell("「N/A」")} |`));
  assert.ok(markdownSection.includes(`| ${cell("-前缀")} | ${cell("「—」")} |`));
  assert.ok(markdownSection.includes(`| ${cell("<b>备注</b>|x")} | ${cell("「-」")} | 可能是有效值 |`));
  for (const line of markdownSection.split("\n")) {
    assert.ok(!line.startsWith("-") && !line.startsWith("<"), line);
  }
  assert.ok(markdownSection.includes(scope.evaluate(`escapeMarkdownText(${JSON.stringify(`说明：${PLACEHOLDER_MAYBE_NOTE_TEXT}在会计格式里，「-」通常表示 0。`)})`)));
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
  const html = fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8");
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
  const { report, html, markdown } = reportOf(orders);
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
  assert.doesNotMatch(markdown, /- 按单位读取：/);
  assert.doesNotMatch(markdown, /像是数字的文本列/);
  assert.ok(markdown.includes(`### 按单位读取的数字\n\n- ${orders.evaluate(`escapeMarkdownText(${JSON.stringify(ORDER_AMOUNT_CONVERSION_TEXT)})`)}\n\n`));
  assert.ok(markdown.includes("| 金额（元） | 29 | 1 |"));
  assert.ok(markdown.includes("- 指标字段：金额（元）"));

  const scope = commitColumns(UNIT_SECTION_COLUMNS);
  const constructed = reportOf(scope);
  assert.deepEqual(constructed.report.quality.numericLooking, [{ field: "楼层", valueCount: 6, numericCount: 4, ratio: 4 / 6, samples: ["B1", "B2"] }]);
  assert.ok(constructed.html.includes("；编号重复：<strong>0</strong> 个字段；像是数字的文本列：<strong>1</strong> 个字段。</p>"));
  assert.doesNotMatch(constructed.html, /按单位读取：<strong>/);
  assert.ok(constructed.html.includes(`<h3>按单位读取的数字</h3>\n    <ul class="summary-list"><li>${WEIGHT_CONFLICT_TEXT}</li><li>${SALES_CONVERSION_TEXT}</li></ul>`));
  assert.ok(constructed.html.includes(`<h3>像是数字的文本列</h3>\n    <ul class="summary-list"><li>${FLOOR_HINT_TEXT}</li></ul>`));
  assert.ok(constructed.markdown.includes("- 编号重复：0 个字段\n- 像是数字的文本列：1 个字段\n"));
  assert.doesNotMatch(constructed.markdown, /- 按单位读取：/);
  assert.ok(constructed.markdown.includes("### 按单位读取的数字\n\n- "));
  assert.ok(constructed.markdown.includes(`### 像是数字的文本列\n\n- ${scope.evaluate(`escapeMarkdownText(${JSON.stringify(FLOOR_HINT_TEXT)})`)}\n\n`));
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
  const markdown = scope.evaluate("buildMarkdownReport(__legacyNumbers)");
  assert.doesNotMatch(html, /按单位读取|像是数字的文本列/);
  assert.doesNotMatch(markdown, /按单位读取|像是数字的文本列/);
  assert.ok(html.includes("<td>金额</td><td>3</td>"));
  assert.ok(markdown.includes("| 金额 | 3 |"));
  // Present but empty lists: no section and no count either.
  put(scope.context, "__emptyNumbers", { ...legacy, quality: { ...legacy.quality, numberConversions: [], numericLooking: [] } });
  assert.doesNotMatch(scope.evaluate("buildHtmlReport(__emptyNumbers, [])"), /按单位读取|像是数字的文本列/);
  assert.doesNotMatch(scope.evaluate("buildMarkdownReport(__emptyNumbers)"), /按单位读取|像是数字的文本列/);
});

test("number sections and unit labels escape field names in HTML and Markdown", () => {
  const field = "<b>重\n量</b>|x";
  const scope = commitColumns({ [field]: ["5kg", "6kg", "7kg", "8kg", "9kg"], "-楼层": ["1", "2", "3", "B1", "B2"] });
  const { html, markdown } = reportOf(scope);
  assert.ok(html.includes("<td>&lt;b&gt;重\n量&lt;/b&gt;|x（kg）</td><td>5</td>"));
  assert.ok(html.includes(`<li>${scope.evaluate(`escapeHtml(${JSON.stringify(`「${field}」：按"kg"读取了 5 个值（5kg、6kg、7kg）`)})`)}</li>`));
  assert.ok(scope.elements.get("numberConversionDetails").innerHTML.includes("「&lt;b&gt;重\n量&lt;/b&gt;|x」"));
  assert.ok(scope.elements.get("numericLookingDetails").innerHTML.includes("「-楼层」"));
  const conversionLine = `- ${scope.evaluate(`escapeMarkdownText(${JSON.stringify(`「${field}」：按"kg"读取了 5 个值（5kg、6kg、7kg）`)})`)}`;
  assert.ok(markdown.includes(conversionLine));
  assert.ok(conversionLine.includes("<br>") && !conversionLine.includes("<b>"));
  for (const line of markdown.split("\n")) {
    assert.ok(!line.startsWith("<") && !line.startsWith("-楼层") && !line.startsWith("量"), line);
  }
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
  assert.equal(scopes.length, 15);
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
  assert.equal(scopes.length, 15);
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
  const html = fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8");
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
  const html = fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8");
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
  for (const mode of ["filtered", "deduplicated"]) {
    assert.equal(orders.evaluate(`buildCsvExportContent(getProcessedExportRows("${mode}"), state.fields)`), expected(`getProcessedExportRows("${mode}")`), mode);
  }
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

  const html = fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8");
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
  assert.deepEqual(fs.readdirSync(path.join(REPO_ROOT, "examples")), ["sample-orders.csv"]);
  assert.deepEqual(fs.readdirSync(path.join(REPO_ROOT, "tests", "fixtures", "examples")).sort(), [
    "sample-data.csv", "sample-sales.csv", "sample-students.csv", "sample-survey.csv", "sample-used-products.csv", "sample-user-behavior.csv"
  ]);
  assert.deepEqual(["examples", path.join("tests", "fixtures", "examples")].flatMap((dir) => fs.readdirSync(path.join(REPO_ROOT, dir))).sort(), [
    "sample-data.csv", "sample-orders.csv", "sample-sales.csv", "sample-students.csv", "sample-survey.csv", "sample-used-products.csv", "sample-user-behavior.csv"
  ]);
});

const repoFile = (name) => fs.readFileSync(path.join(REPO_ROOT, name), "utf8");
const domKeys = (scope) => JSON.parse(scope.evaluate("JSON.stringify(Object.keys(dom))"));
const typesOf = (scope, names) => JSON.parse(scope.evaluate("JSON.stringify(" + JSON.stringify(names) + ".map((name) => typeof globalThis[name]))"));

test("v2.7.0 workflow: the workflow status bar is gone from the page, the styles and the scripts", () => {
  const scope = createScriptContext();
  assert.doesNotMatch(repoFile("index.html"), /workflow-status|workflowStep/);
  assert.doesNotMatch(repoFile("style.css"), /workflow-|\.step-(?:number|copy|state)/);
  assert.deepEqual(typesOf(scope, ["updateWorkflowSteps"]), ["undefined"]);
  assert.deepEqual(domKeys(scope).filter((key) => key.startsWith("workflowStep")), []);
});

test("v2.7.0 quick grouping: the advanced settings block, its chart and its report fallback are gone", () => {
  const scope = createScriptContext();
  assert.doesNotMatch(repoFile("index.html"), /id="(?:metricField|groupField|aggregateMethod|customAnalysisChart|customAnalysisChartTitle)"|高级设置|快速分组/);
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
  const html = repoFile("index.html");
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
  const html = repoFile("index.html");
  assert.doesNotMatch(html, /insightsSection|id="insights"|主要发现|分析结论<\/a>|<span>09<\/span>/);
  const navHrefs = [...html.matchAll(/<a[^>]*href="#(\w+Section)"><span>(\d\d)<\/span>/g)].map((match) => [match[1], match[2]]);
  assert.deepEqual(navHrefs.map(([, number]) => number), ["01", "02", "03", "04", "05", "06", "07", "08"]);
  assert.equal(navHrefs.at(-1)[0], "exportSection");
  const mobile = [...html.slice(html.indexOf('id="mobileSectionNav"')).matchAll(/<option value="(\w+)">/g)].map((match) => match[1]).slice(0, navHrefs.length);
  assert.deepEqual(mobile, navHrefs.map(([id]) => id));
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
  const { report, html, markdown } = reportOf(scope);
  assert.deepEqual(report.insights, [conclusions(scope).category, conclusions(scope).trend]);
  assert.ok(html.includes("<h2>自动分析结论</h2>"));
  assert.ok(html.includes("<li>" + conclusions(scope).category + "</li>"));
  assert.ok(markdown.includes("1. " + conclusions(scope).category + "\n2. "));
  // one sentence only
  scope.evaluate('dom.dateChartField.value = ""; renderTrendInsight()');
  assert.deepEqual(reportOf(scope).report.insights, [conclusions(scope).category]);
  // none
  scope.evaluate('stageFieldTypeChange("部门", "ignore"); stageFieldTypeChange("品名", "ignore"); stageFieldTypeChange("签约日期", "ignore"); stageFieldTypeChange("回访日期", "ignore"); applyFieldConfiguration()');
  const none = reportOf(scope);
  assert.deepEqual(none.report.insights, []);
  assert.ok(none.html.includes("暂无自动分析结论。"));
  assert.ok(none.markdown.includes("暂无自动分析结论。"));
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
  assert.deepEqual(scopes.filter(([, scope]) => suspectedTotals(scope).length).map(([fileName]) => fileName), ["分区小计表.csv", "门店报表.xlsx", "sample-orders.csv"]);
  assert.deepEqual(scopes.filter(([, scope]) => formulaErrors(scope).length).map(([fileName]) => fileName), ["公式错误.xlsx"]);
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
  const html = fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8");
  const positions = ['id="parseWarningSummary"', '<div id="suspectedTotalSection" class="module-subsection row-note-section hidden">', '<div id="formulaErrorSection" class="module-subsection row-note-section hidden">', '<div class="two-column">'].map((text) => html.indexOf(text));
  assert.equal(positions.every((position) => position >= 0), true);
  assert.deepEqual(positions, [...positions].sort((a, b) => a - b));
  assert.match(html, /<p class="section-note">这些行目前仍按普通数据参与统计，其中的数字会被重复计入求和、平均值和偏大或偏小的值的判断。<\/p>/);
  assert.match(html, /<p class="section-note">这些单元格的公式算出了错误，按空值处理（已计入空值）。整行都是错误值的行不会导入，也不在这里列出。<\/p>/);
});

test("v2.8.0 status: the sentences follow the header note, on the warning branch too, one row or many", async () => {
  const region = await importStage3Fixture("分区小计表.csv");
  assert.equal(region.elements.get("statusPanel").textContent, "分区小计表.csv 已成功载入，分析结果已生成。发现 5 行疑似合计 / 小计（第 5、8、10、14、15 行），目前仍计入统计，详情见“数据质量”。");
  const errors = await importStage3Fixture("公式错误.xlsx");
  assert.equal(errors.elements.get("statusPanel").textContent, "公式错误.xlsx · 公式错误 已成功载入，分析结果已生成。发现 7 个公式错误单元格，已按空值处理，详情见“数据质量”。");
  const warned = await importCsvFile("名,额\r\nA,10\r\nB\r\nC,20\r\n合计,30\r\n");
  assert.equal(warned.elements.get("statusPanel").textContent, "CSV 已载入，但读取时有 1 处疑问，涉及 1 行。请核对数据预览和字段数量。第 5 行「合计」疑似合计行，目前仍计入统计，详情见“数据质量”。");
  const many = await importCsvFile(subtotalsCsv(7));
  assert.match(many.elements.get("statusPanel").textContent, /发现 7 行疑似合计 \/ 小计（第 3、5、7、9、11 行等），目前仍计入统计/);
  const five = await importCsvFile(subtotalsCsv(5));
  assert.match(five.elements.get("statusPanel").textContent, /发现 5 行疑似合计 \/ 小计（第 3、5、7、9、11 行），目前仍计入统计/);
  const { context, evaluate, elements } = createScriptContext();
  put(context, "__rows", [{ 名: "A", 额: "10" }, { 名: "合计", 额: "10" }]);
  evaluate('commitTabularData(["名", "额"], __rows, "x.csv", beginImport())');
  assert.match(elements.get("statusPanel").textContent, /有一行「合计」疑似合计行/);
  put(context, "__manyRows", [{ 名: "A", 额: "10" }, { 名: "A小计", 额: "10" }, { 名: "B", 额: "20" }, { 名: "B小计", 额: "20" }]);
  evaluate('commitTabularData(["名", "额"], __manyRows, "x.csv", beginImport())');
  assert.match(elements.get("statusPanel").textContent, /发现 2 行疑似合计 \/ 小计，目前仍计入统计/);
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
  const { report, html, markdown } = reportOf(store);
  assert.equal(report.quality.suspectedTotals.length, 1);
  assert.ok(html.includes("；疑似合计 / 小计行：<strong>1</strong> 行。"));
  assert.ok(html.includes("<h3>疑似合计 / 小计行</h3>"));
  assert.ok(html.includes("<li>第 33 行「合计」：销售额、成本等于第 5–32 行之和（28 行）。</li>"));
  assertInOrder(html, ["<h3>疑似合计 / 小计行</h3>", "<h3>各字段空值情况</h3>"]);
  assert.ok(markdown.includes("- 疑似合计 / 小计行：1 行"));
  assert.ok(markdown.includes("### 疑似合计 / 小计行"));
  assert.ok(markdown.includes("- 第 33 行「合计」：销售额、成本等于第 5–32 行之和（28 行）。"));
  assertInOrder(markdown, ["### 疑似合计 / 小计行", "### 各字段空值情况"]);
  assert.doesNotMatch(html + markdown, /公式错误/);

  const errors = await importStage3Fixture("公式错误.xlsx");
  const errorReport = reportOf(errors);
  assert.ok(errorReport.html.includes("；公式错误：<strong>7</strong> 个单元格。"));
  assert.ok(errorReport.html.includes("<h3>公式错误单元格</h3>"));
  assert.ok(errorReport.html.includes("<li>「利润率」：3 个（#DIV/0! 2 个、#REF! 1 个），第 3、5、9 行。</li>"));
  assert.ok(errorReport.markdown.includes("- 公式错误：7 个单元格"));
  assert.ok(errorReport.markdown.includes("### 公式错误单元格"));
  assert.ok(errorReport.markdown.includes("- 「利润率」：3 个（\\#DIV/0! 2 个、\\#REF! 1 个），第 3、5、9 行。"));
  assert.doesNotMatch(errorReport.html + errorReport.markdown, /疑似合计/);
  errors.evaluate('stageFieldTypeChange("利润率", "ignore"); applyFieldConfiguration()');
  assert.match(reportOf(errors).html, /「利润率」（字段已忽略）：3 个/);

  const clean = reportOf(await importStage3Fixture("订单表.csv"));
  assert.deepEqual([clean.report.quality.suspectedTotals, clean.report.quality.formulaErrors], [[], []]);
  assert.doesNotMatch(clean.html + clean.markdown, /疑似合计|公式错误/);

  delete report.quality.suspectedTotals;
  delete report.quality.formulaErrors;
  put(store.context, "__old", report);
  assert.equal(store.evaluate("JSON.stringify([normalizeReportSnapshot(__old).quality.suspectedTotals, normalizeReportSnapshot(__old).quality.formulaErrors])"), "[null,null]");
  assert.doesNotMatch(store.evaluate("buildHtmlReport(__old, [])") + store.evaluate("buildMarkdownReport(__old)"), /疑似合计|公式错误/);
});

test("v2.8.0 report: the note matches the page, at most 20 totals, the counts and overflow say how many there are", async () => {
  const pageHtml = fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8");
  const exact = await importCsvFile(subtotalsCsv(20));
  const note = exact.evaluate("SUSPECTED_TOTAL_NOTE");
  assert.equal(note, "这些行目前仍按普通数据参与统计，其中的数字会被重复计入求和、平均值和偏大或偏小的值的判断。");
  assert.ok(pageHtml.includes(`<p class="section-note">${note}</p>`));
  const atLimit = reportOf(exact);
  assert.equal(atLimit.report.quality.suspectedTotals.length, 20);
  assert.ok(atLimit.html.includes(`<p>${note}</p>`));
  assert.ok(atLimit.markdown.includes(note));
  assert.doesNotMatch(atLimit.html + atLimit.markdown, /只列出/);

  const over = reportOf(await importCsvFile(subtotalsCsv(21)));
  assert.equal(over.report.quality.suspectedTotals.length, 21);
  assert.ok(over.html.includes("；疑似合计 / 小计行：<strong>21</strong> 行"));
  assert.ok(over.markdown.includes("- 疑似合计 / 小计行：21 行"));
  for (const output of [over.html, over.markdown]) {
    assert.ok(output.includes("第 41 行「A19小计」"));
    assert.ok(!output.includes("第 43 行「A20小计」"));
    assert.ok(output.includes("共 21 行，只列出前 20 行。"));
    assert.doesNotMatch(output, /全部行请点/);
  }
  assert.equal([...over.html.matchAll(/<li>第 \d+ 行「A\d+小计」/g)].length, 20);
  assert.equal([...over.markdown.matchAll(/^- 第 \d+ 行「A\d+小计」/gm)].length, 20);
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
  const { html, markdown } = reportOf(scope);
  assert.ok(html.includes("<li>第 3 行「&lt;b&gt;y|\n小计」：&lt;i&gt;额|x&lt;/i&gt;等于第 2 行（只有 1 行）。</li>"));
  assert.ok(!html.includes("<b>y"));
  const line = markdown.split("\n").find((text) => text.startsWith("- 第 3 行"));
  assert.ok(line.includes("「&lt;b&gt;y\\|<br>小计」"), line);
  assert.ok(!markdown.split("\n").some((text) => text.startsWith("小计")));
  const errors = reportOf(await importWorkbookFile(buildErrorWorkbook([["名", "<i>额"], ["A", 1], ["B", { e: 7 }], ["C", 2]])));
  assert.ok(errors.html.includes("<li>「&lt;i&gt;额」：1 个（#DIV/0! 1 个），第 3 行。</li>"));
  assert.ok(errors.markdown.includes("- 「&lt;i&gt;额」：1 个（\\#DIV/0! 1 个），第 3 行。"));
});

test("v2.8.0 report: the formula error note matches the page and comes before the list, only with errors", async () => {
  const pageHtml = fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8");
  const errors = await importStage3Fixture("公式错误.xlsx");
  const note = errors.evaluate("FORMULA_ERROR_NOTE");
  assert.equal(note, "这些单元格的公式算出了错误，按空值处理（已计入空值）。整行都是错误值的行不会导入，也不在这里列出。");
  assert.ok(pageHtml.includes(`<p class="section-note">${note}</p>`));
  const { html, markdown } = reportOf(errors);
  assert.ok(html.includes(`<h3>公式错误单元格</h3>\n    <p>${note}</p>\n    <ul class="summary-list"><li>「数量」`));
  assert.ok(markdown.includes(`### 公式错误单元格\n\n${note}\n\n- 「数量」`));
  const totalsOnly = reportOf(await importStage3Fixture("门店报表.xlsx"));
  assert.ok(!(totalsOnly.html + totalsOnly.markdown).includes(note));
  const clean = reportOf(await importStage3Fixture("订单表.csv"));
  assert.ok(!(clean.html + clean.markdown).includes(note));
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
  suspectedTotals: state.suspectedTotals, formulaErrors: state.formulaErrors, dedup: deduplicateRows(state.rows, state.fields).length,
  problems: listProblemCategories().filter((category) => ${JSON.stringify(V281_GUARD_KINDS)}.includes(category.kind)).map((category) => [category, listProblemRows(category).items])
})`)).digest("hex").slice(0, 32);

test("v2.9.0 guard: everything v2.8.1 computed stays the same on the stage3 fixtures and examples", async () => {
  const expected = JSON.parse(fs.readFileSync(path.join(__dirname, "fixtures", "v2.8.1-guard.json"), "utf8"));
  const actual = {};
  for (const fileName of STAGE3_IMPORT_FILES) {
    if (fileName === "写法不一致.csv") continue;
    const scope = await importStage3Fixture(fileName);
    actual[fileName] = v281GuardDigest(scope);
    scope.evaluate("confirmAllCommonPlaceholders()");
    actual[`${fileName}（全部确认后）`] = v281GuardDigest(scope);
  }
  for (const fileName of EXAMPLE_CSV_FILES) {
    // v2.11.0: the new sample is newer than the guard, like 写法不一致.csv above.
    if (fileName === "sample-orders.csv") continue;
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
  scope.evaluate('importExcelSheet(readExcelWorkbook(__buffer), "s.xlsx", "S", beginImport())');
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
  assert.deepEqual(scopes.filter(([, scope]) => dateStylesOf(scope).length).map(([name]) => name), ["写法不一致.csv", "花名册.xlsx", "sample-orders.csv"]);
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
  assert.deepEqual(scopes.filter(([, scope]) => constantsOf(scope).length).map(([name]) => name), ["公式错误.xlsx", "写法不一致.csv"]);
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
  assert.deepEqual(scopes.filter(([, scope]) => lookAlikesOf(scope).length).map(([name]) => name), ["写法不一致.csv", "sample-orders.csv"]);
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
  const html = fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8");
  const positions = ['id="numericLookingSection"', 'id="dateStyleSection"', 'id="constantFieldSection"', 'id="lookAlikeSection"', 'id="placeholderDetails"'].map((id) => html.indexOf(id));
  assert.ok(positions.every((position, index) => position >= 0 && (index === 0 || position > positions[index - 1])), positions.join());
  const writing = await importStage3Fixture("写法不一致.csv");
  assert.equal((writing.elements.get("qualitySummary").innerHTML.match(/class="quality-pill/g) || []).length, 6);
  assert.equal(writing.elements.get("statusPanel").textContent, "写法不一致.csv 已成功载入，分析结果已生成。");
});

test("v2.9.0 reports: the notes are the page's sentences", () => {
  const html = fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8");
  const { evaluate } = createScriptContext();
  for (const name of ["DATE_STYLE_NOTE", "CONSTANT_FIELD_NOTE", "LOOK_ALIKE_NOTE"]) assert.ok(html.includes(`<p class="section-note">${evaluate(name)}</p>`), name);
  assert.match(evaluate("LOOK_ALIKE_NOTE"), /只有不可见字符：格子看起来是空的/);
});

test("v2.9.0 reports: three sections with content, two summary counts, constant columns not counted", async () => {
  const roster = reportOf(await importStage3Fixture("花名册.xlsx"));
  assert.ok(roster.html.includes("<h3>日期写法混用</h3>\n    <p>这些写法本工具都能读成同一个日期，但在 Excel 里文本日期不能按日期排序、筛选，导入别的系统时可能读错。</p>"));
  assert.ok(roster.html.includes("<li>「入职日期」：主写法是 Excel 日期单元格 24 个（如 2016-04-01）；另有 6 行是别的写法：文本 年/月/日 2 个（如 2021/3/5）、文本 中文年月日 2 个（如 2021年3月5日）、文本 年.月.日 2 个（如 2021.03.05）。</li>"));
  assert.ok(roster.markdown.includes("- 编号重复：3 个字段\n- 日期写法混用：1 个字段\n"));
  assert.ok(!(roster.html + roster.markdown).includes("整列只有一个值"));

  const errors = reportOf(await importStage3Fixture("公式错误.xlsx"));
  assert.ok(errors.html.includes("<li>「备注」：只有一种取值「正常」（6 行），另有 3 行是空值。</li>"));
  assert.ok(errors.markdown.includes("### 整列只有一个值\n\n这些列的非空值全都一样。不一定是错误：可能是导出时带上的固定列，也可能是漏填或填错。\n\n- 「备注」：只有一种取值「正常」（6 行），另有 3 行是空值。\n"));
  assert.ok(!errors.markdown.includes("- 整列只有一个值"));
  assert.ok(!errors.html.includes("整列只有一个值：<strong>"));

  const writing = reportOf(await importStage3Fixture("写法不一致.csv"));
  assert.ok(writing.html.includes("；日期写法混用：<strong>1</strong> 个字段；同一个值，写法不一样：<strong>3</strong> 个字段。</p>"));
  assert.ok(writing.markdown.includes("- 日期写法混用：1 个字段\n- 同一个值，写法不一样：3 个字段\n"));
  // The section shows whitespace and invisible characters by name; Top 10 and the sample values
  // still print cells as they are.
  const section = writing.html.slice(writing.html.indexOf("<h3>同一个值，写法不一样</h3>"), writing.html.indexOf("<h3>像是空值的写法</h3>"));
  assert.ok(section.includes("「华东[零宽字符]」1 个"));
  assert.ok(section.includes("（如「SKU-002␣」）"));
  assert.ok(!section.includes("\u200B"));
  assert.ok(writing.markdown.includes("「华东\\[零宽字符\\]」1 个"));
  // Order: after the numeric-looking section's place, before the placeholders, in both outputs.
  const order = (text, titles) => titles.map((title) => text.indexOf(title));
  const htmlOrder = order(writing.html, ["<h3>各字段空值情况</h3>", "<h3>日期写法混用</h3>", "<h3>整列只有一个值</h3>", "<h3>同一个值，写法不一样</h3>", "<h3>像是空值的写法</h3>"]);
  assert.ok(htmlOrder.every((position, index) => position >= 0 && (index === 0 || position > htmlOrder[index - 1])), htmlOrder.join());
  const markdownOrder = order(writing.markdown, ["### 各字段空值情况", "### 日期写法混用", "### 整列只有一个值", "### 同一个值，写法不一样", "### 像是空值的写法"]);
  assert.ok(markdownOrder.every((position, index) => position >= 0 && (index === 0 || position > markdownOrder[index - 1])), markdownOrder.join());

  const clean = reportOf(await importStage3Fixture("订单表.csv"));
  assert.deepEqual([clean.report.quality.dateStyles, clean.report.quality.constantFields, clean.report.quality.lookAlikes], [[], [], []]);
  for (const title of ["日期写法混用", "整列只有一个值", "同一个值，写法不一样"]) assert.ok(!(clean.html + clean.markdown).includes(title), title);
});

test("v2.9.0 reports: cells of nothing but invisible characters, escaped field names and values", async () => {
  const field = "<b>备注|*</b>";
  const scope = await importCsvFile(`"${field}",序\r\n\u200B,1\r\n<i>x</i>,2\r\n<I>X</I>,3\r\n`, "r.csv");
  const { report, html, markdown } = reportOf(scope);
  assert.equal(report.quality.lookAlikes[0].invisibleOnlyCount, 1);
  assert.ok(html.includes("只有不可见字符 1 格"));
  assert.ok(markdown.includes("只有不可见字符 1 格"));
  assert.ok(html.includes("「&lt;b&gt;备注|*&lt;/b&gt;」共 3 行"));
  assert.ok(html.includes("「&lt;i&gt;x&lt;/i&gt;」1 个"));
  assert.ok(html.includes("<h3>同一个值，写法不一样</h3>"));
  assert.doesNotMatch(html.slice(html.indexOf("<h3>同一个值，写法不一样</h3>")), /<i>x|<b>备注/);
  assert.ok(markdown.includes("「&lt;b&gt;备注\\|\\*&lt;/b&gt;」共 3 行"));
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
  const oldOutput = store.evaluate("buildMarkdownReport(__old)") + store.evaluate("buildHtmlReport(__old, [])");
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
    const exports = () => scope.evaluate('buildCsvExportContent(getProcessedExportRows("filtered"), state.fields) + "|" + buildCsvExportContent(getProcessedExportRows("deduplicated"), state.fields)');
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
    "公式错误.xlsx": 0, "写法不一致.csv": 13, "分区小计表.csv": 14, "成绩表-标题无逗号.csv": 3, "成绩表.csv": 3, "花名册.xlsx": 8, "订单表.csv": 36, "门店报表.xlsx": 0,
    "sample-data.csv": 0, "sample-orders.csv": 11, "sample-sales.csv": 0, "sample-students.csv": 0, "sample-survey.csv": 0, "sample-used-products.csv": 0, "sample-user-behavior.csv": 119
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
  const html = repoFile("index.html");
  const options = html.slice(html.indexOf('<div class="data-export-options">'), html.indexOf('<div id="standardizedExportPanel"'));
  assert.equal((options.match(/<div class="data-export-option">/g) || []).length, 3);
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
async function clickSampleButton(fetchedText) {
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
    scope = createExcelContext();
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
  assert.match(script, /fetch\("examples\/sample-orders\.csv"\)/);
  assert.doesNotMatch(script, /sample-data\.csv/);
  assert.match(repoFile("index.html"), /<a href="examples\/sample-orders\.csv" download>下载示例 CSV<\/a>/);
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
    "sample-orders.csv 已成功载入，分析结果已生成。表头在第 2 行，已跳过上方 1 行：「2025年上半年销售订单明细」。第 44 行「合计」疑似合计行，目前仍计入统计，详情见“数据质量”。");
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
  assert.equal(sampleOrdersText(scope, "suspectedTotalDetails"), "第 44 行「合计」：数量、金额等于第 3\u201343 行之和（41 行）。 查看");
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
const SHARED_NOTE_NAMES = ["SUSPECTED_TOTAL_NOTE", "FORMULA_ERROR_NOTE", "DATE_STYLE_NOTE", "CONSTANT_FIELD_NOTE", "LOOK_ALIKE_NOTE", "PLACEHOLDER_MAYBE_NOTE", "PLACEHOLDER_DASH_NOTE"];
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
  findOldInterfaceWords(`${name} markdown report`, first.markdown, hits);
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
  findOldInterfaceWords(`${name} confirmed markdown report`, confirmed.markdown, hits);

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
  findOldInterfaceWords(`${name} trend report`, trend.html + trend.markdown, hits);
  return title ? 1 : 0;
}

test("v2.12.0: old interface words no longer appear in pages, reports, file names or exported notes", async () => {
  const hits = [];
  const html = fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8").replace(/<!--[\s\S]*?-->/g, "");
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
    const { html, markdown } = reportOf(scope);
    assert.ok(html.includes("日期字段「日期」里没有认得出的日期。"));
    assert.ok(markdown.includes("日期字段「日期」里没有认得出的日期。"));

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
  assert.match(html(range, "categoryStats"), />当前筛选结果中没有非空类别值。</);
  range.evaluate('dom.categoryChartField.value = "备注"; renderCategoryTopChart(); dom.numericChartField.value = "金额"; renderNumericDistributionChart()');
  const filtered = rangeTexts();
  assert.ok(filtered.some(([id, sentence]) => id === "dateTrendChart" && sentence === "该字段没有认得出的日期"), JSON.stringify(filtered));
  assert.ok(filtered.some(([id, sentence]) => id === "categoryTopChart" && sentence === "当前筛选结果中没有可统计的分类值"), JSON.stringify(filtered));
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
// blanked, is the v2.12.0 text: a new shell changes no word of the report.
const GUARD_REPORT_SHA256 = "a0238e0947b3924235966b011aeb32a353f9d1d23378241865733e6e33fdfbef";
test("v3.0.0 guard: the sample's report text is the v2.12.0 text", async () => {
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
    texts.push(orders.evaluate("buildHtmlReport(__guardReport, [])"), orders.evaluate("buildMarkdownReport(__guardReport)"));
  };
  orders.evaluate(GUARD_CUSTOM_ANALYSIS);
  snapshot();
  orders.evaluate(GUARD_EAST_FILTER);
  snapshot();
  const normalized = texts.join("\n----\n").split(version).join("VERSION");
  assert.ok(normalized.includes("VERSION"));
  assert.equal(crypto.createHash("sha256").update(normalized).digest("hex"), GUARD_REPORT_SHA256);
});
