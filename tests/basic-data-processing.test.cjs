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
  const countingPapa = Object.assign(Object.create(Papa), {
    parse(...args) {
      papaCalls += 1;
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
  return { ...scope, papaCalls: () => papaCalls };
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
  assert.deepEqual(pillTexts(orders).map((text) => text.split("：")[0]), ["缺失值总数", "重复行", "格式问题", "异常值字段", "ID 重复字段"]);
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
