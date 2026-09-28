const test = require("node:test");
const assert = require("node:assert/strict");
const { createScriptContext } = require("./test-context.cjs");
const { execFileSync } = require("node:child_process");
const fs = require("node:fs");
const path = require("node:path");

// Default release gate: exercise the real data functions from script.js without
// coupling core correctness to browser layout or CSS implementation details.

function installDataset(context, evaluate, rows, fields = Object.keys(rows[0] || {})) {
  context.__testRows = rows;
  context.__testFields = fields;
  evaluate(`(() => {
    state.rows = __testRows;
    state.filteredRows = state.rows.slice();
    state.filterSourceRows = state.rows;
    state.fields = __testFields;
    state.profiles = state.fields.map((field) => buildColumnProfile(field, state.rows));
  })()`);
}

function filteredIds(evaluate, filters) {
  return JSON.parse(evaluate(
    `JSON.stringify(filterRowsByConfig(state.rows, ${JSON.stringify(filters)}).map((row) => row.id))`
  ));
}

function createFilterDataset() {
  return [
    { id: 1, category: "A", amount: "10", date: "2026-01-01" },
    { id: 2, category: "B", amount: "20", date: "2026-01-10" },
    { id: 3, category: "C", amount: "30", date: "2026-01-20" },
    { id: 4, category: "A", amount: "40", date: "2026-02-01" },
    { id: 5, category: "B", amount: "50", date: "2026-02-10" },
    { id: 6, category: "A", amount: "", date: "invalid-date" }
  ];
}

for (const timezone of ["Asia/Shanghai", "UTC", "America/Los_Angeles"]) {
  test(`English calendar dates and existing date formats are stable in ${timezone}`, () => {
    const output = execFileSync(process.execPath, ["-e", `
      const { createScriptContext } = require('./tests/test-context.cjs');
      const { evaluate } = createScriptContext();
      console.log(evaluate('JSON.stringify(["Jan 1, 2025", "January 1, 2025", "1 Jan 2025", "2025-Jan-1", "2025年1月1日", "2025-01-01", new Date(2025, 0, 1), new Date(Date.UTC(2025, 0, 1))].map(value => toDate(value)?.toISOString()))'));
    `], { cwd: path.join(__dirname, ".."), env: { ...process.env, TZ: timezone }, encoding: "utf8" });
    assert.deepEqual(JSON.parse(output), Array(8).fill("2025-01-01T00:00:00.000Z"));
  });
}

test("English month names validate calendar days instead of normalizing invalid dates", () => {
  const { context, evaluate } = createScriptContext();
  const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  months.forEach((month, index) => {
    for (const spelling of [month, month.slice(0, 3).toUpperCase()]) {
      context.__date = `${spelling} 15, 2025`;
      assert.equal(evaluate("toDate(__date)?.toISOString()"), `2025-${String(index + 1).padStart(2, "0")}-15T00:00:00.000Z`);
    }
  });
  for (const value of ["Feb 30, 2025", "Feb 29, 2025", "Apr 31, 2025", "Jan 0, 2025", "Jan 32, 2025", "29 Feb 1900", "Januaryish 1, 2025", "Jan 1, 2025 extra"]) {
    context.__date = value;
    assert.equal(evaluate("toDate(__date)"), null, value);
  }
  assert.equal(evaluate('toDate("Feb 29, 2024")?.toISOString()'), "2024-02-29T00:00:00.000Z");
  assert.equal(evaluate('toDate("29 Feb 2000")?.toISOString()'), "2000-02-29T00:00:00.000Z");
});

test("English dates keep the correct calendar day in profiling and inclusive filtering", () => {
  const { context, evaluate } = createScriptContext();
  installDataset(context, evaluate, [
    { id: 1, date: "Jan 1, 2025" }, { id: 2, date: "Jan 2, 2025" },
    { id: 3, date: "Jan 3, 2025" }, { id: 4, date: "Jan 4, 2025" },
    { id: 5, date: "Feb 30, 2025" }
  ]);
  assert.equal(evaluate('state.profiles.find(p => p.field === "date").typeKey'), "date");
  assert.equal(evaluate('state.profiles.find(p => p.field === "date").conversionFailuresByType.date'), 1);
  assert.deepEqual(filteredIds(evaluate, { category: null, numeric: null, date: { field: "date", start: "2025-01-01", end: "2025-01-01" } }), [1]);
});

for (const method of ["median", "sum", "avg"]) {
  test(`${method} custom analysis preserves statistics and exports median in both report formats`, () => {
    const { context, evaluate } = createScriptContext();
    context.__rows = [1, 10, 100].map(amount => ({ group: "A", amount: String(amount) }));
    evaluate('commitTabularData(["group", "amount"], __rows, "report.csv", beginImport())');
    context.__method = method;
    evaluate('dom.v2MetricField.value="amount"; dom.v2GroupField.value="group"; dom.v2AggregateMethod.value=__method; dom.v2DateField.value=""; renderV2Analysis()');
    assert.equal(evaluate("state.customAnalysis.grouped[0].median"), 10);
    context.__report = JSON.parse(evaluate("JSON.stringify(buildReportData())"));
    assert.equal(context.__report.customAnalysis.aggregateMethod, method);
    assert.equal(context.__report.customAnalysis.rows[0].median, "10");
    assert.equal(context.__report.customAnalysis.rows[0].sum, "111");
    assert.equal(context.__report.customAnalysis.rows[0].average, "37");
    const html = evaluate("buildHtmlReport(__report, [])").split("自定义分组分析结果")[1];
    assert.match(html, /<th>中位数<\/th>/);
    assert.match(html, /<td>A<\/td>\s*<td>3<\/td>\s*<td>111<\/td>\s*<td>37<\/td>\s*<td>10<\/td>\s*<td>1<\/td>\s*<td>100<\/td>/);
    const md = evaluate("buildMarkdownReport(__report)").split("## 自定义分组分析结果")[1];
    assert.ok(md.includes("| group | 记录数 | 求和 | 平均值 | 中位数 | 最小值 | 最大值 |"));
    assert.ok(md.includes("| A | 3 | 111 | 37 | 10 | 1 | 100 |"));
  });
}

test("Markdown table cells escape inline syntax and preserve ordinary text", () => {
  const { context, evaluate } = createScriptContext();
  const cases = [
    ["![x](https://example.com/a.png)", String.raw`\!\[x\]\(https\:\/\/example\.com\/a\.png\)`],
    ["[x](javascript:alert(1))", String.raw`\[x\]\(javascript\:alert\(1\)\)`],
    ["**bold** _em_ ~~strike~~ `code`", String.raw`\*\*bold\*\* \_em\_ \~\~strike\~\~ \`code\``],
    ["中文 English 123", "中文 English 123"],
    ["第一行\r\n第二行\n第三行\r第四行|末列", "第一行<br>第二行<br>第三行<br>第四行\\|末列"],
    [String.raw`C:\data\[x]`, String.raw`C\:\\data\\\[x\]`],
    ["<img src=x>&copy;", "&lt;img src\\=x&gt;&amp;copy\\;"],
    ["www.example.com https://example.com", String.raw`www\.example\.com https\:\/\/example\.com`]
  ];
  for (const [input, expected] of cases) {
    context.__cell = input;
    assert.equal(evaluate("escapeMarkdownCell(__cell)"), expected, input);
    assert.equal(evaluate("buildMarkdownTable([__cell], [[__cell]])"), `| ${expected} |\n| --- |\n| ${expected} |`);
  }
});

test("Markdown report exports untrusted field names and values as inert table text", () => {
  const { context, evaluate } = createScriptContext();
  context.__field = "[field](https://example.com)";
  context.__value = "![pixel](https://example.com/pixel.png) **bold** `code`";
  evaluate('commitTabularData([__field, "amount"], [{[__field]:__value, amount:"10"}], "report.csv", beginImport())');
  const report = evaluate("buildMarkdownReport(buildReportData())");
  assert.ok(report.includes(evaluate("escapeMarkdownCell(__field)")));
  assert.ok(report.includes(evaluate("escapeMarkdownCell(__value)")));
  assert.ok(!report.includes(context.__field));
  assert.ok(!report.includes(context.__value));
  assert.ok(report.includes("[Smart Tabular Analyzer](https://github.com/wy-data-30/smart-tabular-analyzer)"));
});

test("public current version markers agree on v2.1.2", () => {
  const root = path.join(__dirname, "..");
  assert.equal(JSON.parse(fs.readFileSync(path.join(root, "package.json"), "utf8")).version, "2.1.2");
  assert.match(fs.readFileSync(path.join(root, "README.md"), "utf8"), /\*\*Current release:\*\* v2\.1\.2/);
  assert.match(fs.readFileSync(path.join(root, "README.zh-CN.md"), "utf8"), /\*\*当前版本：\*\* v2\.1\.2/);
  const { evaluate } = createScriptContext();
  assert.equal(evaluate("buildReportData().version"), "v2.1.2");
  assert.equal(evaluate("normalizeReportSnapshot({}).version"), "v2.1.2");
});

test("numeric parsing handles common tabular formats safely", () => {
  const { evaluate } = createScriptContext();

  assert.equal(evaluate('toNumber("1,234.50")'), 1234.5);
  assert.equal(evaluate('toNumber("$2,000")'), 2000);
  assert.equal(evaluate('toNumber("(1,250)")'), -1250);
  assert.equal(evaluate('toNumber("25%")'), 0.25);
  assert.equal(evaluate('toNumber("-3.5")'), -3.5);
  assert.equal(evaluate('toNumber("1e308")'), 1e308);
  assert.equal(evaluate('toNumber("5e-324")'), 5e-324);
  assert.equal(evaluate('toNumber("1e309")'), null);
  assert.equal(evaluate('toNumber("12$34")'), null);
  assert.equal(evaluate('toNumber("1,23")'), null);
  assert.equal(evaluate('toNumber("1 234.50")'), 1234.5);
  assert.equal(evaluate('toNumber(42)'), 42);
  assert.equal(evaluate('toNumber("")'), null);
  assert.equal(evaluate('toNumber("not-a-number")'), null);
  assert.equal(evaluate('toNumber("2026-07-18")'), null);
});

test("percentage and currency formats are parsed and classified correctly", () => {
  const { evaluate } = createScriptContext();

  assert.equal(evaluate('toNumber("￥1,234.50")'), 1234.5);
  assert.equal(evaluate('toNumber("€ 2,000")'), 2000);
  assert.equal(evaluate('toNumber("(￥1,250)")'), -1250);
  assert.equal(evaluate('toNumber("10%")'), 0.1);
  assert.equal(evaluate('toNumber("(10%)")'), -0.1);
  assert.equal(evaluate('detectNumericFormat(["10%", "25%", "100%"] )'), "percent");
  assert.equal(evaluate('detectNumericFormat(["10%", "0.25"])'), "number");
});

test("date parsing accepts supported formats and rejects invalid or ambiguous dates", () => {
  const { evaluate } = createScriptContext();

  assert.equal(evaluate('toDate("2026-07-18").toISOString()'), "2026-07-18T00:00:00.000Z");
  assert.equal(evaluate('toDate("2026年7月18日").toISOString()'), "2026-07-18T00:00:00.000Z");
  assert.equal(evaluate('toDate("18/07/2026").toISOString()'), "2026-07-18T00:00:00.000Z");
  assert.equal(evaluate('toDate("07/08/2026")'), null);
  assert.equal(evaluate('toDate("07/08/2026", {order:"dmy"}).toISOString()'), "2026-08-07T00:00:00.000Z");
  assert.equal(evaluate('toDate("2026", {allowYearOnly:true}).toISOString()'), "2026-01-01T00:00:00.000Z");
  assert.equal(evaluate('toDate("2026-02-30")'), null);
  assert.equal(evaluate('toDate("not-a-date")'), null);
});

test("field inference distinguishes numeric, categorical, date and ID fields", () => {
  const { context, evaluate } = createScriptContext();
  const rows = Array.from({ length: 12 }, (_, index) => ({
    order_id: `ORD-${String(index + 1).padStart(3, "0")}`,
    amount: String(100 + index * 10),
    order_date: `2026-07-${String(index + 1).padStart(2, "0")}`,
    region: index % 2 === 0 ? "East" : "West"
  }));
  context.__testRows = rows;

  assert.equal(evaluate('buildColumnProfile("order_id", __testRows).typeKey'), "id");
  assert.equal(evaluate('buildColumnProfile("amount", __testRows).typeKey'), "numeric");
  assert.equal(evaluate('buildColumnProfile("order_date", __testRows).typeKey'), "date");
  assert.equal(evaluate('buildColumnProfile("region", __testRows).typeKey'), "category");
});

test("ID inference recognizes Chinese identifiers without treating ordinary counts as IDs", () => {
  const { context, evaluate } = createScriptContext();
  context.__studentRows = Array.from({ length: 12 }, (_, index) => ({
    学号: String(20260001 + index),
    学号数量: String(index + 1),
    订单ID: `ORDER-${index + 1}`,
    普通数值: String(1000 + index)
  }));

  assert.equal(evaluate('buildColumnProfile("学号", __studentRows).typeKey'), "id");
  assert.equal(evaluate('buildColumnProfile("订单ID", __studentRows).typeKey'), "id");
  assert.equal(evaluate('buildColumnProfile("学号数量", __studentRows).typeKey'), "numeric");
  assert.equal(evaluate('buildColumnProfile("普通数值", __studentRows).typeKey'), "numeric");
});

test("mixed-type fields follow the 80 percent inference threshold", () => {
  const { context, evaluate } = createScriptContext();
  context.__mostlyNumeric = ["1", "2", "3", "4", "bad"].map((value) => ({ value }));
  context.__mostlyText = ["1", "2", "3", "bad", "unknown"].map((value) => ({ value }));

  const numericProfile = JSON.parse(evaluate('JSON.stringify(buildColumnProfile("value", __mostlyNumeric))'));
  const categoryProfile = JSON.parse(evaluate('JSON.stringify(buildColumnProfile("value", __mostlyText))'));
  assert.equal(numericProfile.typeKey, "numeric");
  assert.equal(numericProfile.conversionFailureCount, 1);
  assert.equal(categoryProfile.typeKey, "category");
});

test("missing value statistics count empty, whitespace, null and undefined", () => {
  const { context, evaluate } = createScriptContext();
  context.__testRows = [
    { value: "10" },
    { value: "" },
    { value: "   " },
    { value: null },
    { value: undefined },
    { value: "20" }
  ];

  const profile = JSON.parse(evaluate('JSON.stringify(buildColumnProfile("value", __testRows))'));
  assert.equal(profile.nonMissingCount, 2);
  assert.equal(profile.missingCount, 4);
  assert.equal(profile.missingRate, 4 / 6);
});

test("empty datasets return stable empty results", () => {
  const { context, evaluate } = createScriptContext();
  installDataset(context, evaluate, [], ["value"]);

  const profile = JSON.parse(evaluate('JSON.stringify(state.profiles[0])'));
  assert.equal(profile.total, 0);
  assert.equal(profile.nonMissingCount, 0);
  assert.equal(profile.missingCount, 0);
  assert.equal(evaluate("buildNumericStats([]).length"), 0);
  assert.equal(evaluate("buildCategoryStats([]).length"), 1);
  assert.equal(evaluate("filterRowsByConfig([], createEmptyFilterConfig()).length"), 0);
  assert.equal(evaluate('countDuplicateRows([], ["value"])'), 0);
});

test("a single-row single-column Chinese dataset remains analyzable", () => {
  const { context, evaluate } = createScriptContext();
  installDataset(context, evaluate, [{ 成绩: "88" }], ["成绩"]);

  const stats = JSON.parse(evaluate('JSON.stringify(buildNumericStats(state.rows)[0])'));
  assert.equal(evaluate('state.profiles[0].field'), "成绩");
  assert.equal(evaluate('state.profiles[0].typeKey'), "numeric");
  assert.equal(stats.count, 1);
  assert.equal(stats.mean, 88);
  assert.equal(stats.median, 88);
  assert.equal(stats.std, 0);
  assert.equal(stats.outlierCount, 0);
});

test("an all-empty column is reported as missing without numeric statistics", () => {
  const { context, evaluate } = createScriptContext();
  const rows = [{ 备注: "" }, { 备注: "   " }, { 备注: null }];
  installDataset(context, evaluate, rows, ["备注"]);

  const profile = JSON.parse(evaluate('JSON.stringify(state.profiles[0])'));
  assert.equal(profile.typeKey, "category");
  assert.equal(profile.missingCount, 3);
  assert.equal(profile.missingRate, 1);
  assert.equal(evaluate("buildNumericStats(state.rows).length"), 0);
});

test("duplicate row detection compares every original field value exactly", () => {
  const { context, evaluate } = createScriptContext();
  const rows = [
    { name: "Alice", score: "90" },
    { name: "Alice", score: "90" },
    { name: " Alice ", score: "90" },
    { name: "Alice", score: "91" },
    { name: "Bob", score: "90" }
  ];
  installDataset(context, evaluate, rows, ["name", "score"]);

  assert.equal(evaluate('countDuplicateRows(state.rows, state.fields)'), 1);
  assert.equal(evaluate('deduplicateRows(state.rows, state.fields).length'), 4);
  assert.equal(evaluate('deduplicateRows(state.rows, state.fields)[0] === state.rows[0]'), true);
  assert.equal(evaluate('state.rows.length'), 5);
});

test("mean, median and population standard deviation are correct", () => {
  const { context, evaluate } = createScriptContext();
  context.__values = [1, 2, 3, 4, 5];

  assert.equal(evaluate("mean(__values)"), 3);
  assert.equal(evaluate("median(__values)"), 3);
  assert.ok(Math.abs(evaluate("standardDeviation(__values)") - Math.sqrt(2)) < 1e-12);
  assert.equal(evaluate("median([1, 2, 3, 4])"), 2.5);
});

test("IQR outlier detection identifies values outside the fences", () => {
  const { context, evaluate } = createScriptContext();
  const rows = [1, 2, 3, 4, 5, 6, 7, 100].map((value) => ({ value: String(value) }));
  installDataset(context, evaluate, rows, ["value"]);

  const stats = JSON.parse(evaluate('JSON.stringify(buildNumericStats(state.rows)[0])'));
  assert.equal(stats.outlierCount, 1);
  assert.equal(stats.max, 100);
  assert.ok(stats.upperBound < 100);
});

test("an all-equal numeric column has zero spread and no outliers", () => {
  const { context, evaluate } = createScriptContext();
  const rows = Array.from({ length: 8 }, () => ({ value: "7" }));
  installDataset(context, evaluate, rows, ["value"]);

  const stats = JSON.parse(evaluate('JSON.stringify(buildNumericStats(state.rows)[0])'));
  assert.equal(stats.min, 7);
  assert.equal(stats.max, 7);
  assert.equal(stats.std, 0);
  assert.equal(stats.lowerBound, 7);
  assert.equal(stats.upperBound, 7);
  assert.equal(stats.outlierCount, 0);
});

test("IQR detection handles an extreme outlier without numeric overflow", () => {
  const { context, evaluate } = createScriptContext();
  const rows = [10, 11, 12, 13, 14, 15, 16, 1_000_000_000].map((value) => ({ value: String(value) }));
  installDataset(context, evaluate, rows, ["value"]);

  const stats = JSON.parse(evaluate('JSON.stringify(buildNumericStats(state.rows)[0])'));
  assert.equal(stats.outlierCount, 1);
  assert.equal(stats.max, 1_000_000_000);
  assert.equal(Number.isFinite(stats.mean), true);
  assert.equal(Number.isFinite(stats.std), true);
});

test("extreme finite values do not crash histograms or overflow grouped sums", () => {
  const { context, evaluate } = createScriptContext();
  context.__extremeValues = [-1e308, 1e308];
  context.__overflowValues = [1e308, 1e308];

  assert.equal(
    evaluate("buildHistogram(__extremeValues, 10).reduce((total, bin) => total + bin.count, 0)"),
    2
  );
  assert.equal(evaluate("sumNumbers(__extremeValues)"), 0);
  assert.equal(evaluate("sumNumbers(__overflowValues)"), null);
  assert.equal(
    evaluate('summarizeGroupedValues(new Map([["A", __overflowValues]]))[0].sum'),
    null
  );
});

test("categorical filtering supports one or multiple selected values", () => {
  const { context, evaluate } = createScriptContext();
  installDataset(context, evaluate, createFilterDataset());

  assert.deepEqual(filteredIds(evaluate, {
    category: { field: "category", values: ["A", "C"] },
    numeric: null,
    date: null
  }), [1, 3, 4, 6]);
});

test("numeric range filtering is inclusive and excludes missing numbers", () => {
  const { context, evaluate } = createScriptContext();
  installDataset(context, evaluate, createFilterDataset());

  assert.deepEqual(filteredIds(evaluate, {
    category: null,
    numeric: { field: "amount", min: 20, max: 40 },
    date: null
  }), [2, 3, 4]);
});

test("date range filtering is inclusive and excludes invalid dates", () => {
  const { context, evaluate } = createScriptContext();
  installDataset(context, evaluate, createFilterDataset());

  assert.deepEqual(filteredIds(evaluate, {
    category: null,
    numeric: null,
    date: { field: "date", start: "2026-01-10", end: "2026-02-01" }
  }), [2, 3, 4]);
});

test("categorical, numeric and date filters combine with AND semantics", () => {
  const { context, evaluate } = createScriptContext();
  installDataset(context, evaluate, createFilterDataset());

  assert.deepEqual(filteredIds(evaluate, {
    category: { field: "category", values: ["A", "B"] },
    numeric: { field: "amount", min: 15, max: 50 },
    date: { field: "date", start: "2026-01-05", end: "2026-02-05" }
  }), [2, 4]);
});

test("clearing filters restores all rows without mutating the source", () => {
  const { context, evaluate } = createScriptContext();
  const rows = createFilterDataset();
  const originalSnapshot = JSON.stringify(rows);
  installDataset(context, evaluate, rows);
  context.__filters = {
    category: { field: "category", values: ["A"] },
    numeric: { field: "amount", min: 10, max: 40 },
    date: null
  };

  evaluate(`(() => {
    state.filters = __filters;
    state.filteredRows = filterRowsByConfig(state.rows, state.filters);
  })()`);
  assert.equal(evaluate("getAnalysisRows().length"), 2);

  evaluate("resetFilterState()");
  assert.equal(evaluate("getAnalysisRows().length"), rows.length);
  assert.equal(evaluate("state.filteredRows === state.rows"), false);
  assert.equal(JSON.stringify(rows), originalSnapshot);
});

test("oversized files and Excel ranges are rejected before analysis", () => {
  const { context, elements, evaluate } = createScriptContext();
  context.__largeFile = { name: "large.csv", size: 25 * 1024 * 1024 + 1 };

  evaluate("parseDataFile(__largeFile)");
  assert.match(elements.get("statusPanel").textContent, /25 MB/);
  assert.equal(evaluate("state.rows.length"), 0);

  context.XLSX = {
    utils: {
      decode_range() {
        return { s: { r: 0, c: 0 }, e: { r: 100001, c: 0 } };
      }
    }
  };
  const rowLimitMessage = evaluate(`(() => {
    try {
      validateExcelWorksheetLimits({"!ref": "A1:A100002"}, "Large Sheet");
      return "";
    } catch (error) {
      return error.message;
    }
  })()`);
  assert.match(rowLimitMessage, /100,000/);
});

test("scenario templates use every selected mapping field", () => {
  const { context, evaluate } = createScriptContext();
  const rows = Array.from({ length: 12 }, (_, index) => ({
    score: String(70 + index),
    class: index % 2 ? "Class B" : "Class A",
    subject: index % 3 ? "Math" : "English",
    price: String(100 + index * 10),
    item: index % 2 ? "Phone" : "Laptop",
    platform: index % 2 ? "Market A" : "Market B",
    city: index % 3 ? "Shanghai" : "Beijing",
    condition: index % 2 ? "Good" : "Like New",
    userId: `U-${index + 1}`,
    event: index % 2 ? "view" : "click",
    channel: index % 2 ? "search" : "direct",
    device: index % 2 ? "mobile" : "desktop",
    date: `2026-07-${String(index + 1).padStart(2, "0")}`
  }));
  installDataset(context, evaluate, rows);

  context.__scoreMappings = { score: "score", class: "class", subject: "subject", examDate: "date" };
  context.__usedMappings = {
    price: "price", item: "item", platform: "platform", city: "city", condition: "condition", date: "date"
  };
  context.__behaviorMappings = {
    userId: "userId", event: "event", channel: "channel", device: "device", date: "date"
  };

  const scoreResult = JSON.parse(evaluate("JSON.stringify(analyzeScoreTemplate(__scoreMappings))"));
  const usedResult = JSON.parse(evaluate("JSON.stringify(analyzeUsedGoodsTemplate(__usedMappings))"));
  const behaviorResult = JSON.parse(evaluate("JSON.stringify(analyzeBehaviorTemplate(__behaviorMappings))"));

  assert.deepEqual([...new Set(scoreResult.table.rows.map((row) => row[0]))].sort(), ["class", "subject"]);
  assert.deepEqual(
    [...new Set(usedResult.table.rows.map((row) => row[0]))].sort(),
    ["city", "condition", "item", "platform"]
  );
  assert.deepEqual(
    [...new Set(behaviorResult.table.rows.map((row) => row[0]))].sort(),
    ["channel", "device", "event"]
  );
});
