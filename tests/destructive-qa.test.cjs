const test = require("node:test");
const assert = require("node:assert/strict");
const { createScriptContext } = require("./test-context.cjs");

function installDataset(context, evaluate, rows, fields = Object.keys(rows[0] || {})) {
  context.__qaRows = rows;
  context.__qaFields = fields;
  evaluate(`(() => {
    state.rows = __qaRows;
    state.filteredRows = state.rows.slice();
    state.filterSourceRows = state.rows;
    state.fields = __qaFields;
    state.profiles = state.fields.map((field) => buildColumnProfile(field, state.rows));
    state.fieldTypeDrafts = Object.fromEntries(state.profiles.map((profile) => [profile.field, profile.typeKey]));
    state.numericStats = buildNumericStats(state.rows);
    state.qualityNumericStats = buildNumericStats(state.rows);
    state.categoryStats = buildCategoryStats(state.rows);
    state.totalMissing = state.profiles.reduce((sum, profile) => sum + profile.missingCount, 0);
    state.duplicateRows = countDuplicateRows(state.rows, state.fields);
    state.analysisCompletedAt = new Date("2026-07-22T00:00:00.000Z");
  })()`);
}

function installPapaUnparse(context) {
  context.window.Papa = {
    unparse(input, options) {
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
}

test("empty CSV, header-only CSV, missing headers and duplicate headers fail cleanly", () => {
  const { context, elements, evaluate } = createScriptContext();
  const cases = [
    { fields: [], data: [], message: /没有可分析的数据/ },
    { fields: ["name"], data: [], message: /没有可分析的数据/ },
    { fields: ["name", ""], data: [{ name: "Alice", "": "orphan" }], message: /缺少表头/ },
    { fields: ["name", "name"], data: [{ name: "Alice" }], message: /重复字段/ }
  ];

  cases.forEach((entry, index) => {
    context.__qaResult = { errors: [], meta: { fields: entry.fields }, data: entry.data };
    const importId = evaluate("beginImport()");
    context.__qaImportId = importId;
    evaluate(`handleParsedData(__qaResult, "case-${index}.csv", __qaImportId)`);
    assert.equal(evaluate("state.rows.length"), 0);
    assert.match(elements.get("statusPanel").textContent, entry.message);
  });
});

test("very long Chinese field names and emoji survive profiling and HTML escaping", () => {
  const { context, elements, evaluate } = createScriptContext();
  const field = `字段<script>${"超长".repeat(5000)}😀`;
  const rows = [{ [field]: "上海😀" }];
  installDataset(context, evaluate, rows, [field]);
  context.__qaField = field;

  evaluate("renderPreview(); renderSchema()");
  assert.equal(evaluate("state.profiles[0].field === __qaField"), true);
  assert.match(elements.get("previewTable").innerHTML, /上海😀/);
  assert.match(elements.get("schemaTable").innerHTML, /字段&lt;script&gt;/);
  assert.doesNotMatch(elements.get("schemaTable").innerHTML, /字段<script>/);
});

test("comma, quote, newline, Chinese and emoji values survive analysis input", () => {
  const { context, elements, evaluate } = createScriptContext();
  const value = '上海, "重点客户"\n第二行😀';
  context.__qaResult = {
    errors: [],
    meta: { fields: ["描述"] },
    data: [{ 描述: value }]
  };
  const importId = evaluate("beginImport()");
  context.__qaImportId = importId;
  evaluate('handleParsedData(__qaResult, "quoted.csv", __qaImportId)');

  assert.equal(evaluate('state.rows[0]["描述"]'), value);
  assert.match(elements.get("previewTable").innerHTML, /&quot;重点客户&quot;/);
  assert.match(elements.get("previewTable").innerHTML, /第二行😀/);
});

test("UTF-8, GBK and GB18030 decode Chinese and supplementary characters", () => {
  const { evaluate } = createScriptContext();
  const utf8 = Array.from(new TextEncoder().encode("姓名,城市\n张三,上海"));
  const gbk = [
    0xd0, 0xd5, 0xc3, 0xfb, 0x2c, 0xb3, 0xc7, 0xca, 0xd0, 0x0a,
    0xd5, 0xc5, 0xc8, 0xfd, 0x2c, 0xc9, 0xcf, 0xba, 0xa3
  ];
  const gb18030 = [0x95, 0x34, 0xb2, 0x35];

  assert.equal(evaluate(`decodeCsvBuffer(new Uint8Array(${JSON.stringify(utf8)}).buffer, "utf-8").text`), "姓名,城市\n张三,上海");
  assert.equal(evaluate(`decodeCsvBuffer(new Uint8Array(${JSON.stringify(gbk)}).buffer, "gbk").text`), "姓名,城市\n张三,上海");
  assert.equal(evaluate(`decodeCsvBuffer(new Uint8Array(${JSON.stringify(gbk)}).buffer, "auto").text`), "姓名,城市\n张三,上海");
  assert.equal(evaluate(`decodeCsvBuffer(new Uint8Array(${JSON.stringify(gb18030)}).buffer, "gb18030").text`), "𠮷");
});

test("scientific notation and extreme finite numbers remain numeric without overflow", () => {
  const { context, evaluate } = createScriptContext();
  assert.equal(evaluate('toNumber("1e308")'), 1e308);
  assert.equal(evaluate('toNumber("-1E308")'), -1e308);
  assert.equal(evaluate('toNumber("5e-324")'), 5e-324);
  assert.equal(evaluate('toNumber("1e309")'), null);

  const rows = [{ value: "-1e308" }, { value: "1e308" }];
  installDataset(context, evaluate, rows, ["value"]);
  const stats = JSON.parse(evaluate('JSON.stringify(buildNumericStats(state.rows)[0])'));
  assert.equal(stats.mean, 0);
  assert.equal(stats.median, 0);
  assert.equal(stats.std, 1e308);
  assert.equal(Number.isFinite(stats.mean), true);
  assert.equal(Number.isFinite(stats.std), true);
});

test("non-standard and invalid dates are handled conservatively", () => {
  const { evaluate } = createScriptContext();
  assert.equal(evaluate('toDate("2026/7/8").toISOString()'), "2026-07-08T00:00:00.000Z");
  assert.equal(evaluate('toDate("8.7.2026", {order:"dmy"}).toISOString()'), "2026-07-08T00:00:00.000Z");
  assert.equal(evaluate('toDate("07/08/2026")'), null);
  assert.equal(evaluate('toDate("2026-13-01")'), null);
  assert.equal(evaluate('toDate("2026-02-29")'), null);
});

test("empty Excel sheets fail and multi-sheet workbooks enter selection state", () => {
  const { context, elements, evaluate } = createScriptContext();
  const emptyMessage = evaluate(`(() => {
    try { buildExcelTabularData([], [], "空表"); return ""; }
    catch (error) { return error.message; }
  })()`);
  assert.match(emptyMessage, /为空/);

  context.__qaWorkbook = { SheetNames: ["销售", "库存"], Sheets: { 销售: {}, 库存: {} } };
  const importId = evaluate("beginImport()");
  context.__qaImportId = importId;
  evaluate('showExcelSheetSelection(__qaWorkbook, "multi.xlsx", ["销售", "库存"], __qaImportId)');
  assert.equal(evaluate("state.pendingExcel.fileName"), "multi.xlsx");
  assert.equal(elements.get("excelSheetSelect").options.length, 2);
  assert.deepEqual(elements.get("excelSheetSelect").options.map((option) => option.value), ["销售", "库存"]);
});

test("a newer upload invalidates older results and removes stale state", () => {
  const { context, evaluate } = createScriptContext();
  const firstImport = evaluate("beginImport()");
  context.__firstImport = firstImport;
  context.__firstResult = { errors: [], meta: { fields: ["old"] }, data: [{ old: "stale" }] };
  const secondImport = evaluate("beginImport()");
  context.__secondImport = secondImport;
  context.__secondResult = { errors: [], meta: { fields: ["new"] }, data: [{ new: "current" }] };

  evaluate('handleParsedData(__firstResult, "old.csv", __firstImport)');
  assert.equal(evaluate("state.rows.length"), 0);
  evaluate('handleParsedData(__secondResult, "new.csv", __secondImport)');
  assert.equal(evaluate('state.fields.join(",")'), "new");
  assert.equal(evaluate('state.rows[0].new'), "current");
  assert.equal(evaluate('Object.prototype.hasOwnProperty.call(state.rows[0], "old")'), false);
});

test("a newer upload without date data does not keep the previous date chart title", () => {
  const { context, evaluate } = createScriptContext();
  class FakeChart {
    destroy() {}
  }
  context.Chart = FakeChart;
  context.window.Chart = FakeChart;
  const byId = (id) => context.document.getElementById(id);
  const canvasAttributes = {};
  byId("dateTrendChart").setAttribute = (name, value) => { canvasAttributes[name] = value; };
  const defaultTitle = "日期字段趋势（记录数）";

  context.__datedRows = [
    { 日期: "2026-07-01", 金额: "10" },
    { 日期: "2026-07-02", 金额: "20" }
  ];
  evaluate('commitTabularData(["日期", "金额"], __datedRows, "dated.csv", beginImport())');
  byId("dateChartField").value = "日期";
  // The trend chart reads its own metric select; #metricField no longer drives it.
  byId("dateChartMetric").value = "金额";
  evaluate("renderCharts()");
  const previousTitle = byId("dateChartTitle").textContent;
  assert.match(previousTitle, /金额/);

  // The date select keeps a field name the new dataset no longer has valid dates for.
  context.__undatedRows = [{ 城市: "上海", 分数: "90" }, { 城市: "北京", 分数: "80" }];
  evaluate('commitTabularData(["城市", "分数"], __undatedRows, "undated.csv", beginImport())');
  assert.equal(byId("dateChartMetric").value, "");
  evaluate("renderCharts()");
  assert.equal(byId("dateChartTitle").textContent, defaultTitle);
  assert.equal(canvasAttributes["aria-label"], defaultTitle);

  // No date field selected at all.
  byId("dateChartField").value = "";
  canvasAttributes["aria-label"] = previousTitle;
  byId("dateChartTitle").textContent = previousTitle;
  evaluate("renderCharts()");
  assert.equal(byId("dateChartTitle").textContent, defaultTitle);
  assert.equal(canvasAttributes["aria-label"], defaultTitle);
});

test("repeated field type changes rebuild cleanly without leaving filter state", () => {
  const { context, evaluate } = createScriptContext();
  const rows = Array.from({ length: 12 }, (_, index) => ({ value: String(index + 1), group: index % 2 ? "B" : "A" }));
  installDataset(context, evaluate, rows, ["value", "group"]);

  for (let index = 0; index < 20; index += 1) {
    context.__qaType = index % 2 ? "numeric" : "category";
    evaluate('state.fieldTypeDrafts.value = __qaType; applyFieldConfiguration()');
    assert.equal(evaluate('state.profiles.find((profile) => profile.field === "value").typeKey'), context.__qaType);
    assert.equal(evaluate("getAnalysisRows().length"), rows.length);
    assert.equal(evaluate("hasActiveFilters()"), false);
  }
  assert.equal(evaluate('state.numericStats.find((item) => item.field === "value").mean'), 6.5);
});

test("repeated filtering and clearing never mutates or loses source rows", () => {
  const { context, evaluate } = createScriptContext();
  const rows = Array.from({ length: 200 }, (_, index) => ({
    id: index + 1,
    category: index % 2 ? "B" : "A",
    amount: String(index),
    date: `2026-01-${String((index % 28) + 1).padStart(2, "0")}`
  }));
  const snapshot = JSON.stringify(rows);
  installDataset(context, evaluate, rows);
  context.__qaFilters = {
    category: { field: "category", values: ["A"] },
    numeric: { field: "amount", min: 20, max: 120 },
    date: { field: "date", start: "2026-01-05", end: "2026-01-20" }
  };

  for (let index = 0; index < 100; index += 1) {
    evaluate('state.filteredRows = filterRowsByConfig(state.rows, __qaFilters)');
    assert.ok(evaluate("getAnalysisRows().length") > 0);
    evaluate("resetFilterState()");
    assert.equal(evaluate("getAnalysisRows().length"), rows.length);
  }
  assert.equal(JSON.stringify(rows), snapshot);
});

test("building reports does not block later filtering and analysis", () => {
  const { context, evaluate } = createScriptContext();
  const rows = [
    { amount: "10", group: "A", date: "2026-01-01" },
    { amount: "20", group: "B", date: "2026-01-02" },
    { amount: "30", group: "A", date: "2026-01-03" }
  ];
  installDataset(context, evaluate, rows);
  context.__qaReport = JSON.parse(evaluate('JSON.stringify(buildReportData(new Date("2026-07-22T01:00:00.000Z")))'));
  assert.doesNotThrow(() => evaluate("buildHtmlReport(__qaReport, [])"));
  assert.doesNotThrow(() => evaluate("buildMarkdownReport(__qaReport)"));

  context.__qaFilters = { category: { field: "group", values: ["A"] }, numeric: null, date: null };
  evaluate(`(() => {
    state.filters = __qaFilters;
    state.filteredRows = filterRowsByConfig(state.rows, state.filters);
    state.numericStats = buildNumericStats(getAnalysisRows());
  })()`);
  assert.equal(evaluate("getAnalysisRows().length"), 2);
  assert.equal(evaluate('state.numericStats.find((item) => item.field === "amount").mean'), 20);
  assert.equal(evaluate("state.rows.length"), 3);
});

test("CSV export preserves Chinese, emoji, commas, quotes and UTF-8 BOM", async () => {
  const { context, evaluate } = createScriptContext();
  installPapaUnparse(context);
  context.__qaRows = [{ 姓名: "张三😀", 备注: '上海, "重点"\n第二行' }];
  context.__qaFields = ["姓名", "备注"];

  const csv = evaluate("buildCsvExportContent(__qaRows, __qaFields)");
  assert.equal(csv, '姓名,备注\r\n张三😀,"上海, ""重点""\n第二行"');
  context.__qaCsv = csv;
  const blob = evaluate('createUtf8BomBlob(__qaCsv, "text/csv;charset=utf-8")');
  const bytes = new Uint8Array(await blob.arrayBuffer());
  assert.deepEqual(Array.from(bytes.slice(0, 3)), [0xef, 0xbb, 0xbf]);
  assert.equal(new TextDecoder("utf-8").decode(bytes.slice(3)), csv);

  assert.equal(evaluate('sanitizeCsvExportValue("=1+1")'), "'=1+1");
  assert.equal(evaluate('sanitizeCsvExportValue("@SUM(A1:A2)")'), "'@SUM(A1:A2)");
  assert.equal(evaluate('sanitizeCsvExportValue("-42")'), "-42");
  assert.equal(evaluate('sanitizeCsvExportValue("12$34")'), "12$34");
  context.__qaFormulaRows = [{ note: '=HYPERLINK("https://example.invalid", "open")' }];
  context.__qaFormulaFields = ["note"];
  assert.equal(
    evaluate("buildCsvExportContent(__qaFormulaRows, __qaFormulaFields)"),
    'note\r\n"\'=HYPERLINK(""https://example.invalid"", ""open"")"'
  );
});

test("100,000-row analysis and combined filtering complete with correct results", (testContext) => {
  const { context, evaluate } = createScriptContext();
  const rows = Array.from({ length: 100000 }, (_, index) => ({
    id: `ROW-${String(index + 1).padStart(6, "0")}`,
    category: `Group-${index % 20}`,
    amount: String((index % 10000) / 10),
    date: `2026-${String((index % 12) + 1).padStart(2, "0")}-${String((index % 28) + 1).padStart(2, "0")}`,
    note: index % 997 === 0 ? "" : `记录-${index % 100}`
  }));

  const analysisStartedAt = Date.now();
  installDataset(context, evaluate, rows);
  const analysisDuration = Date.now() - analysisStartedAt;
  context.__qaFilters = {
    category: { field: "category", values: ["Group-3", "Group-7"] },
    numeric: { field: "amount", min: 100, max: 500 },
    date: { field: "date", start: "2026-03-01", end: "2026-09-30" }
  };

  const filterStartedAt = Date.now();
  const filteredCount = evaluate("filterRowsByConfig(state.rows, __qaFilters).length");
  const filterDuration = Date.now() - filterStartedAt;

  assert.equal(evaluate("state.rows.length"), 100000);
  assert.equal(evaluate("state.duplicateRows"), 0);
  assert.ok(filteredCount > 0 && filteredCount < rows.length);
  assert.equal(evaluate('state.profiles.find((profile) => profile.field === "amount").typeKey'), "numeric");
  testContext.diagnostic(`100k analysis: ${analysisDuration} ms; combined filter: ${filterDuration} ms; result rows: ${filteredCount}`);
});
