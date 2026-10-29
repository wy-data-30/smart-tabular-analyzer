const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");
const { createScriptContext } = require("./test-context.cjs");

// The real vendored SheetJS builds and reads workbooks at test time, so no
// binary fixture is committed and the production read options are exercised.
const XLSX = require(path.join(__dirname, "..", "vendor", "xlsx-0.20.3.full.min.js"));
const Papa = require(path.join(__dirname, "..", "vendor", "papaparse-5.4.1.min.js"));
const fs = require("node:fs");

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

function commitExampleCsv(scope, fileName) {
  const text = fs.readFileSync(path.join(__dirname, "..", "examples", fileName), "utf8");
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
  assert.deepEqual(qualityPill(scope.elements, "格式问题"), { tone: "warning", text: "格式问题：3 个值" });
  const amountRow = qualityRow(scope.elements, "金额");
  assert.equal(amountRow[amountRow.length - 1].warning, true);
  assert.equal(amountRow[amountRow.length - 1].text, "3（N/A、待定、约500）");
  assert.deepEqual(qualityRow(scope.elements, "地区").slice(-1)[0], { warning: false, text: "0" });

  const clean = createScriptContext();
  commitRows(clean, ["金额"], ["1", "2", "3"].map((amount) => ({ 金额: amount })), "clean.csv");
  assert.deepEqual(qualityPill(clean.elements, "格式问题"), { tone: "success", text: "格式问题：0 个值" });
});

test("unconvertible dates count as format problems with samples", () => {
  const scope = createScriptContext();
  const dates = ["下周一", "2024-13-45", ...Array.from({ length: 8 }, (_, index) => `2024-01-0${index + 1}`)];
  commitRows(scope, ["日期"], dates.map((value) => ({ 日期: value })), "dates.csv");

  assert.equal(scope.evaluate('getFieldProfile("日期").typeKey'), "date");
  assert.deepEqual(qualityPill(scope.elements, "格式问题"), { tone: "warning", text: "格式问题：2 个值" });
  assert.equal(qualityRow(scope.elements, "日期").slice(-1)[0].text, "2（下周一、2024-13-45）");
});

test("statistics count unconvertible values for the current rows only", () => {
  const scope = createScriptContext();
  commitRows(scope, ["金额", "地区"], BAD_AMOUNT_ROWS);

  assert.equal(scope.evaluate('state.numericStats.find((item) => item.field === "金额").invalidCount'), 3);
  assert.match(scope.elements.get("numericStatsTable").innerHTML, /<th>有效数量<\/th>\s*<th>无法转换<\/th>/);
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
  assert.deepEqual(qualityPill(scope.elements, "格式问题"), { tone: "warning", text: "格式问题：3 个值" });
  assert.equal(qualityRow(scope.elements, "金额").slice(-1)[0].text, "3（N/A、待定、约500）");
});

test("field configuration changes update unconvertible counts and samples", () => {
  const scope = createScriptContext();
  commitRows(scope, ["金额", "地区"], BAD_AMOUNT_ROWS);

  scope.evaluate('stageFieldTypeChange("金额", "category"); applyFieldConfiguration()');
  assert.equal(scope.evaluate('getFieldProfile("金额").conversionFailureCount'), 0);
  assert.deepEqual(JSON.parse(scope.evaluate('JSON.stringify(getFieldProfile("金额").conversionFailureSamples)')), []);
  assert.deepEqual(qualityPill(scope.elements, "格式问题"), { tone: "success", text: "格式问题：0 个值" });

  scope.evaluate('stageFieldTypeChange("金额", "numeric"); applyFieldConfiguration()');
  assert.equal(scope.evaluate('getFieldProfile("金额").conversionFailureCount'), 3);
  assert.deepEqual(JSON.parse(scope.evaluate('JSON.stringify(getFieldProfile("金额").conversionFailureSamples)')), ["N/A", "待定", "约500"]);
  assert.deepEqual(qualityPill(scope.elements, "格式问题"), { tone: "warning", text: "格式问题：3 个值" });

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
  assert.match(bad.elements.get("v2ResultSummary").textContent, /指标字段「金额」有 3 个值无法转换为数字，已从计算中排除。$/);
  assert.equal(bad.evaluate("state.customAnalysis.metricInvalidCount"), 3);

  bad.evaluate(`(() => {
    state.filters = { category: { field: "地区", values: ["华东"] }, numeric: null, date: null };
    state.filteredRows = filterRowsByConfig(state.rows, state.filters);
    state.filterSourceRows = state.rows;
    refreshFilteredAnalysis();
  })()`);
  assert.match(bad.elements.get("v2ResultSummary").textContent, /有 1 个值无法转换为数字/);

  const clean = run(BAD_AMOUNT_ROWS.map((row, index) => ({ ...row, 金额: String((index + 1) * 100) })));
  assert.doesNotMatch(clean.elements.get("v2ResultSummary").textContent, /无法转换/);
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
  assert.match(html.split("数值字段描述性统计")[1], /<th>有效数量<\/th><th>无法转换<\/th>/);
  assert.match(html.split("数值字段描述性统计")[1], /<td>金额<\/td><td>12<\/td><td>3<\/td>/);
  assert.match(html, /<td>3（&lt;b&gt;x&lt;\/b&gt;\|y、待定、约500）<\/td>/);
  assert.doesNotMatch(html, /<b>x<\/b>/);

  const markdown = scope.evaluate("buildMarkdownReport(__report)");
  assert.ok(markdown.includes("| 字段 | 有效数量 | 无法转换 | 平均值 | 中位数 | 最小值 | 最大值 | 标准差 |"));
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
  assert.equal(total.title, "销售额 按月份趋势（合计）");
  assert.deepEqual(total.labels, SAMPLE_MONTHS);
  assert.equal(total.data[0], 14400);
  assert.equal(total.data[5], 89140);
});

test("the date trend chart ignores the quick-grouping metric select", () => {
  const scope = createChartContext();
  commitExampleCsv(scope, "sample-data.csv");

  const before = drawDateTrend(scope);
  scope.evaluate('dom.metricField.value = "销售额"; renderCharts()');
  const after = { ...scope.lastChart("dateTrendChart"), title: scope.elements.get("dateChartTitle").textContent };
  assert.deepEqual(after, before);
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
  assert.equal(scope.elements.get("dateChartTitle").textContent, "销售额 按月份趋势（合计）");

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
  assert.match(total.summary, /销售额合计从 2025-01 的 14,400 变为 2025-06 的 89,140/);
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
  assert.equal(scope.evaluate("state.insights[state.insights.length - 1]"), scope.evaluate('buildTrendInsight("日期")'));
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

test("the trend insight ignores numeric fields and both metric selects", () => {
  const dates = ["2025-01-05", "2025-02-05", "2025-03-05", "2025-04-05", "2025-06-10"];
  const withAmounts = (amounts) => dates.map((date, index) => ({ 日期: date, 金额: String(amounts[index]) }));
  const { scope, text } = trendInsightFor(withAmounts([1, 2, 3, 4, 5]));
  assert.match(text, /记录数趋势/);
  assert.equal(trendInsightFor(withAmounts([500, 1, 900, 2, 0])).text, text);
  assert.equal(trendInsightFor(datedRows(dates)).text, text);

  scope.evaluate('dom.metricField.value = "金额"; dom.dateChartMetric.value = "金额"; dom.dateChartAggregate.value = "sum"; renderCharts()');
  assert.equal(scope.evaluate('buildTrendInsight("日期")'), text);
});

const SCATTERED_CATEGORY_INSIGHT = "分类字段的取值都比较分散（每个字段超过 20 个不同值，或没有重复出现的值），没有适合概括的分类结论。";

function categoryInsight(scope) {
  return JSON.parse(scope.evaluate('JSON.stringify(state.insights.find((text) => text.startsWith("分类字段「")) ?? null)'));
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

test("the category conclusion falls back when every category field is scattered", () => {
  const scattered = createScriptContext();
  const rows = Array.from({ length: 30 }, (_, index) => ({ 甲: `a${index % 25}`, 乙: `b${index % 26}` }));
  commitRows(scattered, ["甲", "乙"], rows, "scattered.csv");
  assert.equal(categoryInsight(scattered), null);
  assert.equal(scattered.evaluate(`state.insights.includes(${JSON.stringify(SCATTERED_CATEGORY_INSIGHT)})`), true);

  const noCategory = createScriptContext();
  commitRows(noCategory, ["日期", "金额"], [{ 日期: "2025-01-01", 金额: "1" }, { 日期: "2025-01-02", 金额: "2" }], "plain.csv");
  assert.equal(categoryInsight(noCategory), null);
  assert.equal(noCategory.evaluate(`state.insights.includes(${JSON.stringify(SCATTERED_CATEGORY_INSIGHT)})`), false);
  assert.equal(noCategory.evaluate('state.insights.includes("未识别到适合生成频数统计的分类字段。")'), true);
});

test("the category conclusion re-picks fields from the filtered rows", () => {
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
  assert.equal(categoryInsight(scope), "分类字段「款号」中占比最高的类别是「g0」，出现 2 次，占比 10.0%。");

  const report = JSON.parse(scope.evaluate('JSON.stringify(buildReportData(new Date("2026-10-25T00:00:00.000Z")))'));
  assert.deepEqual(report.categoryStats.map((item) => [item.field, item.distinctCount]), [["款号", 10], ["区域", 1]]);
  assert.deepEqual(report.insights, JSON.parse(scope.evaluate("JSON.stringify(state.insights)")));
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
  assert.equal(store.elements.get("statusPanel").textContent, `门店报表.xlsx · 销售汇总 已成功载入，分析结果已生成。${STORE_HEADER_NOTE}`);
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
    "CSV 已载入，但发现 1 个解析警告，涉及 1 行。请核对数据预览和字段数量。表头在第 2 行，已跳过上方 1 行：「期中成绩」。");
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
  for (const [label, linebreak, bom] of [["LF", "\n", ""], ["CR", "\r", ""], ["CRLF+BOM", "\r\n", "﻿"]]) {
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
    ["a,b\r\n1\r\n", "status-panel warning", "CSV 已载入，但发现 2 个解析警告，涉及 1 行。请核对数据预览和字段数量。", [
      { row: null, code: "UndetectableDelimiter", message: "Unable to auto-detect delimiting character; defaulted to ','" },
      { row: 2, code: "TooFewFields", message: "Too few fields: expected 2 fields but parsed 1" }
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
  const text = fs.readFileSync(path.join(__dirname, "..", "examples", "sample-user-behavior.csv"), "utf8");
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
const behaviorCsv = () => fs.readFileSync(path.join(__dirname, "..", "examples", "sample-user-behavior.csv"), "utf8");

test("the quality page has a static ID duplicate container in the dom map", () => {
  const html = fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8");
  const quality = html.slice(html.indexOf('id="qualitySection"'), html.indexOf('id="filterSection"'));
  assert.match(quality, /<h3 class="subsection-title">ID 字段重复值<\/h3>\s*<div id="idDuplicateDetails"/);
  const { evaluate } = createScriptContext();
  assert.equal(evaluate('Object.keys(dom).includes("idDuplicateDetails")'), true);
});

test("the ID duplicate pill counts listed fields after the existing pills", async () => {
  const pillTexts = (scope) => [...scope.elements.get("qualitySummary").innerHTML.matchAll(/<span class="quality-pill[^"]*">([^<]*)<\/span>/g)].map((match) => match[1]);
  const orders = await importStage3Fixture("订单表.csv");
  assert.deepEqual(pillTexts(orders).map((text) => text.split("：")[0]), ["缺失值总数", "重复行", "格式问题", "异常值字段", "ID 重复字段", "疑似缺失值"]);
  assert.deepEqual(qualityPill(orders.elements, "ID 重复字段"), { tone: "warning", text: "ID 重复字段：1" });
  const grades = await importStage3Fixture("成绩表.csv");
  assert.deepEqual(qualityPill(grades.elements, "ID 重复字段"), { tone: "success", text: "ID 重复字段：0" });
  // 用户ID and 会话ID are mostly repeated, so only 事件ID counts.
  const behavior = await importCsvFile(behaviorCsv(), "sample-user-behavior.csv");
  assert.deepEqual(qualityPill(behavior.elements, "ID 重复字段"), { tone: "warning", text: "ID 重复字段：1" });
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
  assert.ok(idDetails(grades).includes("ID 字段没有重复值（已检查：学号、姓名）"));

  const none = createScriptContext();
  commitRows(none, ["城市"], [{ 城市: "上海" }, { 城市: "上海" }], "cities.csv");
  assert.ok(idDetails(none).includes("未识别到 ID 字段。"));

  const mixed = createScriptContext();
  commitRows(mixed, ["学号", "订单号"], [{ 学号: "S1", 订单号: "D1" }, { 学号: "S2", 订单号: "D1" }], "mixed.csv");
  assert.ok(idDetails(mixed).includes('<table aria-label="订单号 的重复值">'));
  assert.ok(idDetails(mixed).includes("其余 ID 字段没有重复值（已检查：学号）"));
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
  assert.ok(html.includes("；解析警告：<strong>0</strong>；ID 重复字段：<strong>3</strong>。</p>"));
  const htmlSection = html.slice(html.indexOf("<h3>ID 字段重复值</h3>"), html.indexOf("数值字段描述性统计"));
  assert.ok(html.indexOf("<h3>数值异常值</h3>") < html.indexOf("<h3>ID 字段重复值</h3>"));
  assert.match(htmlSection, /<th>字段<\/th><th>重复值<\/th><th>出现次数<\/th><th>文件行号<\/th><th>说明<\/th>/);
  assert.ok(htmlSection.includes("<td>工号</td><td>E1012</td><td>3</td><td>13、14、25</td><td>其中 2 行内容完全相同</td>"));

  assert.ok(markdown.includes("- 解析警告：0\n- ID 重复字段：3\n"));
  assert.ok(markdown.indexOf("### 数值异常值") < markdown.indexOf("### ID 字段重复值"));
  assert.ok(markdown.indexOf("### ID 字段重复值") < markdown.indexOf("## 数值字段描述性统计"));
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
  assert.ok(behaviorReport.html.includes("；ID 重复字段：<strong>1</strong>。"));

  const ids = [
    ...Array.from({ length: 15 }, () => "S000"),
    ...Array.from({ length: 11 }, (_, index) => [`S${index + 1}`, `S${index + 1}`]).flat(),
    ...Array.from({ length: 40 }, (_, index) => `S${index + 100}`)
  ];
  const long = reportOf(await importCsvFile(`学号\r\n${ids.join("\r\n")}\r\n`));
  assert.ok(long.html.includes("<p>「学号」共 12 个重复值，只列出前 10 个</p>"));
  assert.ok(long.markdown.includes("「学号」共 12 个重复值，只列出前 10 个"));

  const grades = reportOf(await importStage3Fixture("成绩表.csv"));
  const gradeSection = grades.html.slice(grades.html.indexOf("<h3>ID 字段重复值</h3>"), grades.html.indexOf("数值字段描述性统计"));
  assert.ok(gradeSection.includes("<p>ID 字段没有重复值（已检查：学号、姓名）</p>"));
  assert.doesNotMatch(gradeSection, /<table>/);
  assert.ok(grades.markdown.includes("### ID 字段重复值\n\nID 字段没有重复值（已检查：学号、姓名）\n"));
  assert.ok(grades.html.includes("；ID 重复字段：<strong>0</strong>。"));

  const none = createScriptContext();
  commitRows(none, ["城市"], [{ 城市: "上海" }, { 城市: "北京" }], "cities.csv");
  const noneReport = reportOf(none);
  assert.deepEqual(noneReport.report.quality.idDuplicates, []);
  assert.ok(noneReport.html.includes("<p>未识别到 ID 字段。</p>"));
  assert.ok(noneReport.markdown.includes("### ID 字段重复值\n\n未识别到 ID 字段。\n"));
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
  assert.doesNotMatch(html, /ID 字段重复值|ID 重复字段/);
  assert.doesNotMatch(markdown, /ID 字段重复值|ID 重复字段/);
  assert.ok(html.includes("；解析警告：<strong>0</strong>。</p>"));
  assert.ok(html.includes("<td>金额</td><td>1</td><td>1</td><td>9</td>"));
  assert.ok(markdown.includes("- 解析警告：0\n\n### 各字段缺失情况"));
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

test("the grade fixtures and the bundled examples have no placeholder candidates", async () => {
  for (const fileName of ["成绩表.csv", "成绩表-标题无逗号.csv"]) {
    const scope = await importStage3Fixture(fileName);
    assert.equal(scope.evaluate("state.rows.length"), 30, fileName);
    assert.deepEqual(placeholderEntries(scope.evaluate), [], fileName);
  }
  const examples = fs.readdirSync(path.join(__dirname, "..", "examples")).filter((fileName) => fileName.endsWith(".csv"));
  assert.equal(examples.length, 6);
  for (const fileName of examples) {
    const scope = await importCsvFile(fs.readFileSync(path.join(__dirname, "..", "examples", fileName), "utf8"), fileName);
    assert.ok(scope.evaluate("state.rows.length") > 0, fileName);
    assert.deepEqual(placeholderEntries(scope.evaluate), [], fileName);
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
  assert.equal(placeholderMessage(scope), "已把「金额」的「N/A」按缺失处理，统计、图表和分析结论已重新计算。 已清除现有筛选条件。");
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
  assert.equal(placeholderMessage(scope), "已把 2 处常见写法按缺失处理，统计、图表和分析结论已重新计算。 已清除现有筛选条件。");
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
    [["未填写", 3, 120], ["上海", 2, 60], ["北京", 1, 30]]
  );
  assert.deepEqual(
    JSON.parse(evaluate('JSON.stringify(aggregateByCategory("金额", "城市", "sum").map((group) => [group.name, group.value]))')),
    [["未填写", 120], ["上海", 60], ["北京", 30]]
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
  const refusal = "字段配置有未应用的修改。请先点击“应用字段配置”，再处理疑似缺失值。";
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
      dom.metricField.value = "金额"; dom.groupField.value = "备注"; dom.aggregateMethod.value = "sum";
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
  const placeholderSection = quality.slice(quality.indexOf("疑似缺失值"), quality.indexOf("ID 字段重复值"));
  assert.ok(quality.indexOf('<h3 class="subsection-title">疑似缺失值</h3>') < quality.indexOf('<h3 class="subsection-title">ID 字段重复值</h3>'));
  assert.ok(placeholderSection.includes("这些写法常被用来表示“没有数据”。确认后按缺失值处理：不再算作转换失败，也不参与统计、Top 10、分组、筛选选项和 ID 重复检查。原始数据、数据预览和导出的 CSV 不变。"));
  // The message sits outside the container that renderPlaceholderDetails() rewrites.
  assert.match(placeholderSection, /<div id="placeholderDetails" class="placeholder-details"><\/div>/);
  assert.match(placeholderSection, /<p id="placeholderMessage" class="placeholder-message" role="status" tabindex="-1"><\/p>/);
  const { evaluate } = createScriptContext();
  assert.equal(evaluate('["placeholderDetails", "placeholderMessage"].every((id) => Object.keys(dom).includes(id))'), true);
  assert.match(fs.readFileSync(path.join(__dirname, "..", "script.js"), "utf8"), /dom\.placeholderDetails\.addEventListener\("click", handlePlaceholderClick\);/);
});

test("the placeholder pill counts fields with unconfirmed keys after the ID duplicate pill", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  assert.deepEqual(qualityPill(orders.elements, "疑似缺失值"), { tone: "warning", text: "疑似缺失值：4 个字段" });
  orders.evaluate('confirmPlaceholder("金额", "N/A")');
  assert.deepEqual(qualityPill(orders.elements, "疑似缺失值"), { tone: "warning", text: "疑似缺失值：3 个字段" });
  orders.evaluate("confirmAllCommonPlaceholders()");
  assert.deepEqual(qualityPill(orders.elements, "疑似缺失值"), { tone: "success", text: "疑似缺失值：0 个字段" });
  const grades = await importStage3Fixture("成绩表.csv");
  assert.deepEqual(qualityPill(grades.elements, "疑似缺失值"), { tone: "success", text: "疑似缺失值：0 个字段" });
});

test("placeholder details show each key with its rows and effect, one button each and a whole-field line", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  const html = placeholderDetails(orders);
  for (const text of [
    "「下单日期」：「—」1 个（第 17 行）。确认后缺失值 0 → 1，类型仍为日期。",
    "「数量」：「—」1 个（第 12 行）。确认后缺失值 0 → 1，类型仍为数值。",
    "「金额」：「N/A」1 个（第 15 行）。确认后缺失值 0 → 1，类型仍为数值。",
    "「备注」：「N/A」1 个（第 29 行）。确认后缺失值 26 → 27，类型仍为分类。",
    "「备注」：「—」1 个（第 30 行）。确认后缺失值 26 → 27，类型仍为分类。",
    "「备注」全部确认后：缺失值 26 → 28，类型仍为分类。"
  ]) assert.ok(html.includes(text), text);
  assert.deepEqual([...html.matchAll(/data-placeholder-index="(\d+)"[^>]*>按缺失处理</g)].map((match) => match[1]), ["0", "1", "2", "3", "4"]);
  assert.ok(html.includes("全部确认常见写法（5 处）"));
  assert.doesNotMatch(html, /确认后类型会变/);
  assert.doesNotMatch(html, /也可能是有效值/);

  const grades = await importStage3Fixture("成绩表.csv");
  assert.ok(placeholderDetails(grades).includes("没有发现疑似缺失值。"));
});

test("placeholder details name raw texts, maybe-valid keys, type changes and missing row numbers", () => {
  const amount = commitColumns({ 金额: ["100", "200", "N/A", "约500", "300", "400", "500"], 年份: YEARS_WITH_DASHES });
  const amountHtml = placeholderDetails(amount);
  assert.ok(amountHtml.includes("「金额」：「N/A」1 个（行号：—）。确认后缺失值 0 → 1，「金额」会识别为数值（现在是分类）。"));
  assert.ok(amountHtml.includes("「年份」：「—」3 个（行号：—）。确认后缺失值 0 → 3，「年份」会识别为日期（现在是分类）。"));
  assert.ok(amountHtml.includes("全部确认常见写法（2 处）"));
  assert.ok(amountHtml.includes("确认后类型会变：「金额」「年份」"));

  const raw = commitColumns({ 备注: ["Ｎ／Ａ", "－－", "好", "N/A", " N/A "] });
  const rawHtml = placeholderDetails(raw);
  assert.ok(rawHtml.includes("「备注」：「N/A」（原文：Ｎ／Ａ、N/A）3 个（行号：—）。确认后缺失值 0 → 3，类型仍为分类。"));
  assert.ok(rawHtml.includes("「备注」：「--」（原文：－－）1 个（行号：—）。确认后缺失值 0 → 1，类型仍为分类。"));
  assert.ok(rawHtml.includes("「备注」全部确认后：缺失值 0 → 4，类型仍为分类。"));

  const maybe = commitColumns({ 是否会员: ["有", "无", "有", "无", "有", "有"], 金额: ["1,200", "-", " - ", "3,400", "560", "(200)"] });
  const maybeHtml = placeholderDetails(maybe);
  assert.ok(maybeHtml.includes("「是否会员」：「无」2 个（行号：—）。确认后缺失值 0 → 2，类型仍为分类。"));
  assert.equal(maybeHtml.split(PLACEHOLDER_MAYBE_NOTE_TEXT).length - 1, 2);
  assert.ok(maybeHtml.includes(`${PLACEHOLDER_MAYBE_NOTE_TEXT}在会计格式里，「-」通常表示 0。`));
  assert.equal(maybeHtml.split("在会计格式里").length - 1, 1);
  // No common key, so no confirm-all button.
  assert.doesNotMatch(maybeHtml, /全部确认常见写法/);

  const escaped = commitColumns({ "<b>x</b>": ["N/A", "1", "2"] });
  assert.ok(placeholderDetails(escaped).includes("「&lt;b&gt;x&lt;/b&gt;」：「N/A」1 个"));
  assert.doesNotMatch(placeholderDetails(escaped), /<b>x<\/b>/);
});

test("confirmed keys show an undo button, ignored fields are marked, and more than ten rows end with 等", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  orders.evaluate('confirmPlaceholder("金额", "N/A"); confirmPlaceholder("备注", "N/A")');
  orders.evaluate('stageFieldTypeChange("备注", "ignore"); applyFieldConfiguration()');
  const html = placeholderDetails(orders);
  assert.ok(html.includes("「金额」：「N/A」已按缺失处理（1 个）。"));
  assert.ok(html.includes("「备注」（字段已忽略）：「N/A」已按缺失处理（1 个）。"));
  assert.deepEqual([...html.matchAll(/data-placeholder-index="(\d+)"[^>]*>撤销</g)].map((match) => match[1]), ["2", "3"]);
  assert.ok(html.includes("全部确认常见写法（2 处）"));

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
  assert.deepEqual(before.report.quality.placeholders.pending[2], { field: "金额", key: "N/A", tier: "common", rawValues: ["N/A"], count: 1, rowNumbers: [15], effect: "确认后缺失值 0 → 1，类型仍为数值" });
  assert.equal(before.report.quality.placeholders.pending.length, 5);
  assert.ok(before.html.includes("；ID 重复字段：<strong>1</strong>；疑似缺失值：<strong>4</strong> 个字段。</p>"));
  const htmlSection = before.html.slice(before.html.indexOf("<h3>疑似缺失值</h3>"), before.html.indexOf("<h3>数值异常值</h3>"));
  assert.ok(before.html.indexOf("<h3>各字段缺失情况</h3>") < before.html.indexOf("<h3>疑似缺失值</h3>"));
  assert.match(htmlSection, /<th>字段<\/th><th>写法<\/th><th>档位<\/th><th>出现次数<\/th><th>文件行号<\/th><th>确认后的影响<\/th>/);
  assert.ok(htmlSection.includes("<td>金额</td><td>「N/A」</td><td>常见</td><td>1</td><td>15</td><td>确认后缺失值 0 → 1，类型仍为数值</td>"));
  assert.ok(before.markdown.includes("- ID 重复字段：1\n- 疑似缺失值：4 个字段\n"));
  assert.ok(before.markdown.indexOf("### 各字段缺失情况") < before.markdown.indexOf("### 疑似缺失值"));
  assert.ok(before.markdown.indexOf("### 疑似缺失值") < before.markdown.indexOf("### 数值异常值"));
  const cell = (value) => orders.evaluate(`escapeMarkdownCell(${JSON.stringify(value)})`);
  assert.ok(before.markdown.includes(`| 金额 | ${cell("「N/A」")} | 常见 | 1 | 15 | 确认后缺失值 0 → 1，类型仍为数值 |`));

  orders.evaluate('confirmPlaceholder("金额", "N/A"); confirmPlaceholder("备注", "N/A")');
  orders.evaluate('stageFieldTypeChange("备注", "ignore"); applyFieldConfiguration()');
  const after = reportOf(orders);
  assert.deepEqual(after.report.quality.placeholders.confirmed, [
    { field: "金额", key: "N/A", rawValues: ["N/A"], count: 1, fieldIgnored: false },
    { field: "备注", key: "N/A", rawValues: ["N/A"], count: 1, fieldIgnored: true }
  ]);
  assert.ok(after.html.includes("；疑似缺失值：<strong>2</strong> 个字段。</p>"));
  assert.ok(after.html.includes("<td>金额</td><td>「N/A」</td><td>1</td><td>已按缺失处理</td>"));
  assert.ok(after.html.includes("<td>备注</td><td>「N/A」</td><td>1</td><td>已按缺失处理（字段已忽略）</td>"));
  assert.ok(after.markdown.includes(`| 备注 | ${cell("「N/A」")} | 1 | 已按缺失处理（字段已忽略） |`));
});

test("report placeholder summaries appear only with candidates, and legacy snapshots leave the section out", async () => {
  for (const scope of [await importStage3Fixture("花名册.xlsx"), await importStage3Fixture("成绩表.csv"), await importCsvFile(behaviorCsv(), "sample-user-behavior.csv")]) {
    const { report, html, markdown } = reportOf(scope);
    assert.deepEqual(report.quality.placeholders, { confirmed: [], pending: [] });
    assert.doesNotMatch(html, /疑似缺失值：/);
    assert.doesNotMatch(markdown, /疑似缺失值：/);
    assert.ok(html.includes("<p>没有发现疑似缺失值。</p>"));
    assert.ok(markdown.includes("### 疑似缺失值\n\n没有发现疑似缺失值。\n"));
  }
  const scope = createScriptContext();
  assert.equal(scope.evaluate("normalizeReportSnapshot({}).quality.placeholders"), null);
  assert.equal(scope.evaluate("normalizeReportSnapshot({ quality: { placeholders: { confirmed: [] } } }).quality.placeholders"), null);
  put(scope.context, "__legacyPlaceholders", { quality: { totalMissing: 1, duplicateRows: 0, parseWarningCount: 0 } });
  assert.doesNotMatch(scope.evaluate("buildHtmlReport(__legacyPlaceholders, [])"), /疑似缺失值/);
  assert.doesNotMatch(scope.evaluate("buildMarkdownReport(__legacyPlaceholders)"), /疑似缺失值/);
});

test("report placeholder sections escape field names and keep Markdown lines from starting with user text", () => {
  const scope = commitColumns({ "<b>备注</b>|x": ["N/A", "好", "-", "好"], "-前缀": ["—", "甲", "乙", "丙"] });
  const { html, markdown } = reportOf(scope);
  const section = html.slice(html.indexOf("<h3>疑似缺失值</h3>"), html.indexOf("<h3>数值异常值</h3>"));
  assert.ok(section.includes("<td>&lt;b&gt;备注&lt;/b&gt;|x</td><td>「N/A」</td>"));
  assert.ok(section.includes("<td>-前缀</td><td>「—」</td>"));
  assert.doesNotMatch(html, /<b>备注<\/b>/);
  assert.ok(section.includes(`<p>说明：${PLACEHOLDER_MAYBE_NOTE_TEXT}在会计格式里，「-」通常表示 0。</p>`));
  const markdownSection = markdown.slice(markdown.indexOf("### 疑似缺失值"), markdown.indexOf("### 数值异常值"));
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
  assert.deepEqual(qualityPill(orders.elements, "格式问题"), { tone: "warning", text: "格式问题：3 个值" });
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
    assert.equal(qualityPill(scope.elements, "格式问题").text, `格式问题：${formatProblems} 个值`, fileName);
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

test("the sales template sums yuan amounts like the main analysis", async () => {
  const orders = await importStage3Fixture("订单表.csv");
  const result = JSON.parse(orders.evaluate(`JSON.stringify((() => {
    const sales = analyzeSalesTemplate({ salesAmount: "金额", unitPrice: "", quantity: "", category: "产品", region: "大区", date: "下单日期" });
    const regions = Object.fromEntries(sales.secondaryChart.labels.map((label, index) => [label, sales.secondaryChart.values[index]]));
    const main = Object.fromEntries(groupNumericFieldByCategory("金额", "大区").map((group) => [group.name, group.sum]));
    const trend = Object.fromEntries(sales.trendChart.labels.map((label, index) => [label, sales.trendChart.values[index]]));
    const generic = analyzeGenericTemplate({ metric: "金额", group: "大区", date: "下单日期" });
    return { regions, main, trend: [trend["2024-05-08"], trend["2024-05-20"]], genericTrend: generic.trendChart ? [generic.trendChart.values[generic.trendChart.labels.indexOf("2024-05-08")], generic.trendChart.values[generic.trendChart.labels.indexOf("2024-05-20")]] : null };
  })())`));
  assert.equal(result.regions["华南"], 6025);
  assert.deepEqual(result.regions, result.main);
  assert.deepEqual(result.trend, [3560, 4204]);
  assert.deepEqual(result.genericTrend, [3560, 4204]);
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
const FLOOR_HINT_TEXT = "「楼层」：去掉缺失值和疑似缺失值后，6 个值里有 4 个能读成数字（66.7%），其余的如 B1、B2。如果这一列应该是数字，请修正这些值，或在字段配置里把类型改成数值。";
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
    dom.metricField.value = "金额"; dom.groupField.value = "大区"; dom.aggregateMethod.value = "sum";
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
  assert.equal(text("customAnalysisChartTitle"), "金额（元） 按 大区 分组（合计）");
  assert.equal(text("v2PrimaryTableTitle"), "金额（元） 按 大区 分组（求和）");
  assert.equal(text("v2TopGroupTitle"), "金额（元） 按 大区 分组 Top 10（求和）");
  assert.equal(text("v2TrendTitle"), "金额（元） 按日期趋势（求和）");
  assert.ok(html("v2PrimaryTable").includes("<th>金额（元） 求和值</th>"));
  assert.ok(html("v2SummaryCards").includes('title="金额（元）">金额（元）</div>'));
  const axisTitles = charts.map((config) => [config.data.datasets[0].label, config.options.scales.y.title.text]);
  assert.deepEqual(axisTitles, [
    ["金额（元） 分布（记录频数）", "记录频数"],
    ["金额（元） 平均值趋势", "平均值（元）"],
    ["大区 分组的 金额（元） 合计", "合计（元）"],
    ["金额（元） 按 大区 分组（求和）", "求和（元）"],
    ["金额（元） 按日期趋势（求和）", "求和（元）"]
  ]);
  for (const id of ["numericChartField", "dateChartMetric", "metricField", "v2MetricField", "numericFilterField"]) {
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
  const order = ['id="qualityTable"', 'id="numberConversionSection"', 'id="numericLookingSection"', '<h3 class="subsection-title">疑似缺失值</h3>'].map((marker) => quality.indexOf(marker));
  assert.ok(order.every((index, position) => index > 0 && (position === 0 || index > order[position - 1])), JSON.stringify(order));
  assert.match(quality, /<div id="numberConversionSection" class="module-subsection hidden">\s*<h3 class="subsection-title">按单位读取的数字<\/h3>\s*<div id="numberConversionDetails" class="number-note-details"><\/div>/);
  assert.match(quality, /<div id="numericLookingSection" class="module-subsection hidden">\s*<h3 class="subsection-title">疑似数值列<\/h3>\s*<div id="numericLookingDetails" class="number-note-details"><\/div>/);
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
  orders.evaluate('dom.dateChartMetric.value = "金额"; dom.dateChartAggregate.value = "avg"; dom.metricField.value = "金额"; dom.groupField.value = "大区"; dom.aggregateMethod.value = "sum"');
  const { report, html, markdown } = reportOf(orders);
  assert.deepEqual(report.fieldTypes.map((field) => field.numberUnit), ["", "", "", "", "", "", "元", ""]);
  assert.deepEqual(report.quality.numberConversions, [
    { field: "金额", unit: "元", unitCount: 2, unitSamples: ["2200元", "3500元"], magnitudeCount: 0, fullwidthCount: 0, conflictUnits: [] }
  ]);
  assert.deepEqual(report.quality.numericLooking, []);
  assert.ok(report.dateTrends[0].summary.startsWith("日期字段「下单日期」包含 20 个时间点；金额（元）平均值从 2024-05-06 的"), report.dateTrends[0].summary);
  assert.ok(html.includes("；疑似缺失值：<strong>4</strong> 个字段。</p>"));
  assert.doesNotMatch(html, /按单位读取：<strong>/);
  assert.doesNotMatch(html, /疑似数值列/);
  assert.ok(html.includes(`<h3>按单位读取的数字</h3>\n    <ul class="summary-list"><li>${orders.evaluate(`escapeHtml(${JSON.stringify(ORDER_AMOUNT_CONVERSION_TEXT)})`)}</li></ul>`));
  assert.ok(html.includes("<td>金额（元）</td><td>29</td><td>1</td>"));
  assert.ok(html.includes("<td>金额（元）</td><td>1</td><td>-4,275</td><td>9,925</td>"));
  assert.ok(html.includes("<strong>指标字段：</strong>金额（元）"));
  assert.doesNotMatch(markdown, /- 按单位读取：/);
  assert.doesNotMatch(markdown, /疑似数值列/);
  assert.ok(markdown.includes(`### 按单位读取的数字\n\n- ${orders.evaluate(`escapeMarkdownText(${JSON.stringify(ORDER_AMOUNT_CONVERSION_TEXT)})`)}\n\n`));
  assert.ok(markdown.includes("| 金额（元） | 29 | 1 |"));
  assert.ok(markdown.includes("- 指标字段：金额（元）"));

  const scope = commitColumns(UNIT_SECTION_COLUMNS);
  const constructed = reportOf(scope);
  assert.deepEqual(constructed.report.quality.numericLooking, [{ field: "楼层", valueCount: 6, numericCount: 4, ratio: 4 / 6, samples: ["B1", "B2"] }]);
  assert.ok(constructed.html.includes("；ID 重复字段：<strong>0</strong>；疑似数值列：<strong>1</strong> 个字段。</p>"));
  assert.doesNotMatch(constructed.html, /按单位读取：<strong>/);
  assert.ok(constructed.html.includes(`<h3>按单位读取的数字</h3>\n    <ul class="summary-list"><li>${WEIGHT_CONFLICT_TEXT}</li><li>${SALES_CONVERSION_TEXT}</li></ul>`));
  assert.ok(constructed.html.includes(`<h3>疑似数值列</h3>\n    <ul class="summary-list"><li>${FLOOR_HINT_TEXT}</li></ul>`));
  assert.ok(constructed.markdown.includes("- ID 重复字段：0\n- 疑似数值列：1 个字段\n"));
  assert.doesNotMatch(constructed.markdown, /- 按单位读取：/);
  assert.ok(constructed.markdown.includes("### 按单位读取的数字\n\n- "));
  assert.ok(constructed.markdown.includes(`### 疑似数值列\n\n- ${scope.evaluate(`escapeMarkdownText(${JSON.stringify(FLOOR_HINT_TEXT)})`)}\n\n`));
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
  assert.doesNotMatch(html, /按单位读取|疑似数值列/);
  assert.doesNotMatch(markdown, /按单位读取|疑似数值列/);
  assert.ok(html.includes("<td>金额</td><td>3</td>"));
  assert.ok(markdown.includes("| 金额 | 3 |"));
  // Present but empty lists: no section and no count either.
  put(scope.context, "__emptyNumbers", { ...legacy, quality: { ...legacy.quality, numberConversions: [], numericLooking: [] } });
  assert.doesNotMatch(scope.evaluate("buildHtmlReport(__emptyNumbers, [])"), /按单位读取|疑似数值列/);
  assert.doesNotMatch(scope.evaluate("buildMarkdownReport(__emptyNumbers)"), /按单位读取|疑似数值列/);
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
