const test = require("node:test");
const assert = require("node:assert/strict");
const { createScriptContext, APP_SCRIPT_FILES } = require("./test-context.cjs");
const { execFileSync } = require("node:child_process");
const fs = require("node:fs");
const path = require("node:path");
const englishDateCorpus = require("./fixtures/english-date-corpus.cjs");

// Default release gate: exercise the real data functions from the app scripts without
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
  test(`fixed English compatibility corpus preserves calendar semantics in ${timezone}`, () => {
    const output = execFileSync(process.execPath, ["-e", `
      const { createScriptContext } = require('./tests/test-context.cjs');
      const { context, evaluate } = createScriptContext();
      context.__corpus = require('./tests/fixtures/english-date-corpus.cjs');
      console.log(evaluate('JSON.stringify(__corpus.map(({input}) => toDate(input)?.toISOString().slice(0, 10) ?? null))'));
    `], { cwd: path.join(__dirname, ".."), env: { ...process.env, TZ: timezone }, encoding: "utf8" });
    const results = JSON.parse(output);
    assert.equal(results.length, englishDateCorpus.length);
    results.forEach((actual, index) => {
      assert.equal(actual, englishDateCorpus[index].expected, englishDateCorpus[index].input);
    });
  });

  test(`trend insight month-end and year-end checks use calendar days in ${timezone}`, () => {
    const output = execFileSync(process.execPath, ["-e", `
      const { createScriptContext } = require('./tests/test-context.cjs');
      const cases = [
        ["日期", ["2023-10-01", "2023-11-01", "2023-12-01", "2024-01-01", "2024-02-29"]],
        ["日期", ["2023-10-01", "2023-11-01", "2023-12-01", "2024-01-01", "2024-02-28"]],
        ["日期", ["2025-01-05", "2025-03-05", "2025-05-05", "2025-06-30"]],
        ["日期", ["2025-01-05", "2025-03-05", "2025-05-05", "2025-06-01"]],
        ["年份", ["2021", "2022", "2023", "2024", "2025-12-31"]],
        ["年份", ["2021", "2022", "2023", "2024", "2025-01-01"]]
      ];
      console.log(JSON.stringify(cases.map(([field, values]) => {
        const { context, evaluate } = createScriptContext();
        context.__rows = values.map((value) => ({ [field]: value }));
        context.__field = field;
        evaluate('commitTabularData([__field], __rows, "tz.csv", beginImport())');
        const text = evaluate('buildTrendInsight(__field)');
        const start = text.indexOf("注意：最后一期");
        return start < 0 ? "" : text.slice(start, text.indexOf("日，", start) + 1);
      })));
    `], { cwd: path.join(__dirname, ".."), env: { ...process.env, TZ: timezone }, encoding: "utf8" });
    assert.deepEqual(JSON.parse(output), [
      "",
      "注意：最后一期（2024-02）的数据截至 2 月 28 日",
      "",
      "注意：最后一期（2025-06）的数据截至 6 月 1 日",
      "",
      "注意：最后一期（2025 年）的数据截至 1 月 1 日"
    ]);
  });

  test(`English calendar dates and existing date formats are stable in ${timezone}`, () => {
    const output = execFileSync(process.execPath, ["-e", `
      const { createScriptContext } = require('./tests/test-context.cjs');
      const { evaluate } = createScriptContext();
      console.log(evaluate('JSON.stringify(["Jan 1, 2025", "January 1, 2025", "1 Jan 2025", "2025-Jan-1", "2025年1月1日", "2025-01-01", new Date(2025, 0, 1), new Date(Date.UTC(2025, 0, 1))].map(value => toDate(value)?.toISOString()))'));
    `], { cwd: path.join(__dirname, ".."), env: { ...process.env, TZ: timezone }, encoding: "utf8" });
    assert.deepEqual(JSON.parse(output), Array(8).fill("2025-01-01T00:00:00.000Z"));
  });

  test(`English date/time compatibility preserves written calendar dates in ${timezone}`, () => {
    const cases = [
      ["Jan 1, 2025", "2025-01-01"],
      ["Jan 1, 2025 12:30:00", "2025-01-01"],
      ["Wed, 01 Jan 2025 12:30:00 GMT", "2025-01-01"],
      ["Wednesday, 1 January 2025 12:30 UTC", "2025-01-01"],
      ["Jan. 1, 2025 12:30 PM", "2025-01-01"],
      ["1-Jan-2025 00:00:00 +1400", "2025-01-01"],
      ["2025-Jan-1 00:00:00 GMT+08:00", "2025-01-01"],
      ["Wed Jan 01 2025 00:00:00 GMT+0800 (China Standard Time)", "2025-01-01"],
      ["Jan 1, 2025 00:00:00.123Z", "2025-01-01"],
      ["Jan 1, 2025 12:00 AM EST", "2025-01-01"],
      ["Jan 1, 2025 GMT", "2025-01-01"],
      ["Dec 31, 2025 23:59:59", "2025-12-31"],
      ["Dec 31, 2025 23:59:59 -1200", "2025-12-31"],
      ["Feb 29, 2024 23:59:59 PST", "2024-02-29"],
      ["Feb 30, 2025", null],
      ["Feb 30, 2025 12:30:00", null],
      ["Sun, 30 Feb 2025 12:30:00 GMT", null],
      ["Jan 32, 2025", null],
      ["Jan 32, 2025 00:00:00 GMT+0800", null],
      ["Apr 31, 2025 12:30 PM", null]
    ];
    const output = execFileSync(process.execPath, ["-e", `
      const { createScriptContext } = require('./tests/test-context.cjs');
      const { context, evaluate } = createScriptContext();
      context.__cases = ${JSON.stringify(cases)};
      console.log(evaluate('JSON.stringify(__cases.map(([value]) => toDate(value)?.toISOString().slice(0, 10) ?? null))'));
    `], { cwd: path.join(__dirname, ".."), env: { ...process.env, TZ: timezone }, encoding: "utf8" });
    assert.deepEqual(JSON.parse(output), cases.map(([, expected]) => expected));
  });

  test(`numeric date-time suffixes keep the written calendar date in ${timezone}`, () => {
    const cases = [
      ["2024-01-15 10:30:00", "2024-01-15"],
      ["2024-01-15 10:30:00.123", "2024-01-15"],
      ["2024-01-15 10:30:00.123456", "2024-01-15"],
      ["2024-01-15T10:30:00.1234567Z", "2024-01-15"],
      ["2024-01-15T10:30:00.123456789Z", "2024-01-15"],
      ["2024-01-15 10:30:00+08", "2024-01-15"],
      ["2024-01-15 10:30:00.123456+08", "2024-01-15"],
      ["2024-01-15 10:30:00+0800", "2024-01-15"],
      ["2024-01-15 10:30:00+08:00", "2024-01-15"],
      ["2024-01-15T00:00:00+14:00", "2024-01-15"],
      ["2024-01-15 23:59:59-12", "2024-01-15"],
      ["15/01/2024 10:30:00.123456+08", "2024-01-15"],
      ["2024年1月15日 10:30:00+08", "2024-01-15"],
      ["2024-01-15T10:30:00+99:99", null],
      ["2024-01-15T10:30:00+24:00", null],
      ["2024-01-15T10:30:00+08:60", null],
      ["2024-01-15T10:30:00+0860", null],
      ["2024-01-15T10:30:00+080", null],
      ["2024-01-15T10:30:00+08:", null],
      ["2024-01-15T10:30:00.1234567890Z", null],
      ["2024-01-15T10:30:00.Z", null],
      ["2024-01-15 24:00:00", null]
    ];
    const output = execFileSync(process.execPath, ["-e", `
      const { createScriptContext } = require('./tests/test-context.cjs');
      const { context, evaluate } = createScriptContext();
      context.__cases = ${JSON.stringify(cases)};
      console.log(evaluate('JSON.stringify(__cases.map(([value]) => toDate(value)?.toISOString().slice(0, 10) ?? null))'));
    `], { cwd: path.join(__dirname, ".."), env: { ...process.env, TZ: timezone }, encoding: "utf8" });
    assert.deepEqual(JSON.parse(output), cases.map(([, expected]) => expected));
  });
}

test("English date/time parsing rejects malformed time and timezone suffixes", () => {
  const { context, evaluate } = createScriptContext();
  for (const suffix of ["24:00:00", "12:60:00", "12:30:60", "00:30 AM", "13:30 PM", "12:30 GMT+2400", "12:30 +0860", "12:30 UNKNOWN", "12:30 GMT extra", "12:30:00 trailing", "extra"]) {
    context.__date = `Jan 1, 2025 ${suffix}`;
    assert.equal(evaluate("toDate(__date)"), null, context.__date);
  }
});

test("English date token extraction rejects misplaced or repeated components", () => {
  const { context, evaluate } = createScriptContext();
  for (const input of [
    "Jan 12:30:00 1 2025", "1 Jan 12:30:00 2025", "Jan 1 2025 12:30:00 13:00:00",
    "Jan 1 2025 GMT 12:30:00", "Jan 1 12:30:00 25", "Jan 1 2025 2026",
    "Jan 1 12:60:00 2025", "Feb 30 12:30:00 2025", "Apr 31 12:30:00 2025",
    "Jan 1 2025, 24:00:00", "Wed Jan 32 12:30:00 2025"
  ]) {
    context.__date = input;
    assert.equal(evaluate("toDate(__date)"), null, input);
  }
});

test("a timezone offset colon is not mistaken for a clock", () => {
  const { context, evaluate } = createScriptContext();
  for (const zone of ["GMT", "UTC", "Z", "+0800", "+08:00", "GMT+08:00", "UTC-12:00", "EST", "PST"]) {
    context.__date = `Jan 1, 2025 ${zone}`;
    assert.equal(evaluate("toDate(__date)?.toISOString()"), "2025-01-01T00:00:00.000Z", zone);
  }
});

test("English date/time columns remain date fields and filter by their written date", () => {
  const { context, evaluate } = createScriptContext();
  installDataset(context, evaluate, [
    { id: 1, date: "Jan 1, 2025 12:30:00" },
    { id: 2, date: "Wed, 01 Jan 2025 12:30:00 GMT" },
    { id: 3, date: "Jan 1, 2025 00:00:00 +1400" },
    { id: 4, date: "Dec 31, 2025 23:59:59 -1200" },
    { id: 5, date: "Feb 30, 2025 12:30:00 GMT" },
    { id: 6, date: "Jan 1, 2025, 12:30:00" },
    { id: 7, date: "Wed Jan 1 12:30:00 2025" }
  ]);
  assert.equal(evaluate('state.profiles.find(p => p.field === "date").typeKey'), "date");
  assert.equal(evaluate('state.profiles.find(p => p.field === "date").conversionFailuresByType.date'), 1);
  assert.deepEqual(filteredIds(evaluate, { category: null, numeric: null, date: { field: "date", start: "2025-01-01", end: "2025-01-01" } }), [1, 2, 3, 6, 7]);
});

test("database and pandas timestamp columns remain date fields and filter by their written date", () => {
  const { context, evaluate } = createScriptContext();
  installDataset(context, evaluate, Array.from({ length: 14 }, (_, index) => {
    const day = 10 + index;
    return {
      id: index + 1,
      pg: `2024-01-${day} 10:30:00.${123456 + index}+08`,
      pandas: `2024-01-${day} 10:30:00.${123456 + index}`,
      mixed: index === 13 ? `2024-01-${day}T10:30:00+99:99` : `2024-01-${day}T10:30:00+08:00`
    };
  }));
  const profile = (field) => `state.profiles.find(p => p.field === "${field}")`;
  const singleDay = (field) => ({ category: null, numeric: null, date: { field, start: "2024-01-15", end: "2024-01-15" } });

  for (const field of ["pg", "pandas"]) {
    assert.equal(evaluate(`${profile(field)}.typeKey`), "date", field);
    assert.equal(evaluate(`${profile(field)}.conversionFailuresByType.date`), 0, field);
    assert.deepEqual(filteredIds(evaluate, singleDay(field)), [6], field);
  }
  assert.equal(evaluate(`${profile("mixed")}.typeKey`), "date");
  assert.equal(evaluate(`${profile("mixed")}.conversionFailuresByType.date`), 1);
  assert.deepEqual(
    filteredIds(evaluate, { category: null, numeric: null, date: { field: "mixed", start: "2024-01-01", end: "2024-01-31" } }),
    Array.from({ length: 13 }, (_, index) => index + 1)
  );
});

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

test("Markdown report text converts every line ending, including a lone CR", () => {
  const { context, evaluate } = createScriptContext();
  const cases = [
    ["a\r---", "a\n---"],
    ["第一行\r\n第二行\r第三行", "第一行\n第二行\n第三行"]
  ];
  for (const [input, lfOnly] of cases) {
    context.__text = input;
    context.__lfOnly = lfOnly;
    const escaped = evaluate("escapeMarkdownText(__text)");
    assert.ok(!/[\r\n]/.test(escaped), JSON.stringify(input));
    assert.equal(escaped, evaluate("escapeMarkdownText(__lfOnly)"), JSON.stringify(input));
  }

  // Our own <br> markup is added after escaping, so it stays a real line break
  // while user-supplied angle brackets remain escaped.
  const exact = [
    ["a\r---", "a<br>---"],
    ["第一行\r\n第二行\n第三行\r第四行", "第一行<br>第二行<br>第三行<br>第四行"],
    ["<b>x</b>\n<br>", "&lt;b&gt;x&lt;/b&gt;<br>&lt;br&gt;"]
  ];
  for (const [input, expected] of exact) {
    context.__text = input;
    assert.equal(evaluate("escapeMarkdownText(__text)"), expected, JSON.stringify(input));
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

test("public current version markers agree on v2.6.1", () => {
  const root = path.join(__dirname, "..");
  assert.equal(JSON.parse(fs.readFileSync(path.join(root, "package.json"), "utf8")).version, "2.6.1");
  assert.match(fs.readFileSync(path.join(root, "README.md"), "utf8"), /\*\*Current release:\*\* v2\.6\.1/);
  assert.match(fs.readFileSync(path.join(root, "README.zh-CN.md"), "utf8"), /\*\*当前版本：\*\* v2\.6\.1/);
  const { evaluate } = createScriptContext();
  assert.equal(evaluate("buildReportData().version"), "v2.6.1");
  assert.equal(evaluate("normalizeReportSnapshot({}).version"), "v2.6.1");

  // index.html cache-busting parameters must follow package.json.
  const packageVersion = JSON.parse(fs.readFileSync(path.join(root, "package.json"), "utf8")).version;
  const html = fs.readFileSync(path.join(root, "index.html"), "utf8");
  assert.equal(html.match(/<link rel="stylesheet" href="style\.css\?v=([^"]+)">/)?.[1], packageVersion);
  const appScriptVersions = [...html.matchAll(/<script src="(?!vendor\/)([^"?]+)\?v=([^"]+)"><\/script>/g)]
    .map((match) => ({ file: match[1], version: match[2] }));
  assert.deepEqual(appScriptVersions.map(({ file }) => file), APP_SCRIPT_FILES);
  appScriptVersions.forEach(({ file, version }) => assert.equal(version, packageVersion, file));
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

// Values a number strategy may convert, with the result under the strategy of a
// column holding only that value.
const NUMBER_STRATEGY_CONVERTS = [
  ["2200元", 2200], ["2,200元", 2200], ["2200 元", 2200], ["-2200元", -2200],
  ["1.2万", 12000], ["1.2万元", 12000], ["3亿元", 300000000], ["0.07万", 700], ["1.15亿", 115000000], ["-0.5万", -5000],
  ["１２３", 123], ["１，２００", 1200], ["１２．５％", 0.125], ["－５", -5],
  ["5KG", 5], ["5Kg", 5], ["30天", 30], ["2小时", 2], ["3人", 3], ["25岁", 25], ["3千克", 3]
];
const NUMBER_STRATEGY_REJECTS = [
  "4G", "8G", "500g", "2楼", "3号线", "第3名", "3人间", "90分", "5m", "3块", "2024年", "5月", "1万2", "1.2 万", "约500",
  "500+", "3-5", "1.2.3", "①", "²", "-", "--", "－", "－－", "—", "N/A", "2200元/件", "￥2200元", "(2200元)"
];

test("number strategies convert whitelisted units, 万/亿 and full-width digits", () => {
  const { context, evaluate } = createScriptContext();
  context.__cases = NUMBER_STRATEGY_CONVERTS;
  const results = JSON.parse(evaluate(`JSON.stringify(__cases.map(([value]) => [value, toNumber(value), toNumber(value, buildNumberStrategy([value]))]))`));
  results.forEach(([value, plain, converted], index) => {
    assert.equal(plain, null, `${value} without a strategy`);
    assert.strictEqual(converted, NUMBER_STRATEGY_CONVERTS[index][1], value);
  });
});

test("万 and 亿 shift the decimal point and give exact integers", () => {
  const { context, evaluate } = createScriptContext();
  context.__cases = [["0.07万", 700], ["1.15亿", 115000000], ["1.2万", 12000], ["-0.5万", -5000], ["1.2万元", 12000], ["3亿元", 300000000], ["1,234.5678万", 12345678]];
  const results = JSON.parse(evaluate(`JSON.stringify(__cases.map(([value]) => {
    const number = toNumber(value, buildNumberStrategy([value]));
    return [number, Number.isInteger(number)];
  }))`));
  results.forEach(([number, isInteger], index) => {
    assert.strictEqual(number, context.__cases[index][1], context.__cases[index][0]);
    assert.equal(isInteger, true, context.__cases[index][0]);
  });
});

test("number strategies leave units off the list, other suffixes and placeholders unconverted", () => {
  const { context, evaluate } = createScriptContext();
  context.__cases = NUMBER_STRATEGY_REJECTS;
  const converted = JSON.parse(evaluate(`JSON.stringify(__cases.filter((value) => {
    // Every whitelisted unit (and none), each with 万/亿 and full width on.
    const widest = ["", ...NUMBER_UNITS].map((unit) => ({ unit: unit.toLowerCase(), unitConflict: false, units: {}, allowMagnitude: true, fullwidth: true }));
    return toNumber(value) !== null
      || toNumber(value, buildNumberStrategy([value])) !== null
      || widest.some((strategy) => toNumber(value, strategy) !== null);
  }))`));
  assert.deepEqual(converted, []);
});

test("without a strategy, or with one that has no switch, toNumber() is the plain parser", () => {
  const { context, evaluate } = createScriptContext();
  context.__values = ["1,234.50", "$2,000", "(1,250)", "25%", "-3.5", "1e308", "1e309", "12$34", "1,23", "1 234.50", 42, "", "  ", null, undefined,
    "not-a-number", "2026-07-18", "￥1,234.50", "(10%)", ...NUMBER_STRATEGY_CONVERTS.map(([value]) => value), ...NUMBER_STRATEGY_REJECTS];
  const mismatches = JSON.parse(evaluate(`JSON.stringify(__values.filter((value) => {
    const plain = toPlainNumber(value);
    return toNumber(value) !== plain
      || toNumber(value, undefined) !== plain
      || toNumber(value, buildNumberStrategy([])) !== plain
      || toNumber(value, { unit: "", unitConflict: true, units: { kg: 1, 千克: 1 }, allowMagnitude: false, fullwidth: false }) !== plain;
  }).map(String))`));
  assert.deepEqual(mismatches, []);
});

test("a number strategy takes one unit, drops conflicting spellings and ignores the column share", () => {
  const { evaluate } = createScriptContext();
  assert.deepEqual(JSON.parse(evaluate('JSON.stringify(buildNumberStrategy(["5kg", "500千克"]))')), {
    unit: "", unitConflict: true, units: { kg: 1, 千克: 1 }, allowMagnitude: false, fullwidth: false
  });
  assert.deepEqual(JSON.parse(evaluate(`JSON.stringify((() => {
    const strategy = buildNumberStrategy(["￥2200", "2200元", "1800"]);
    return { unit: strategy.unit, values: ["￥2200", "2200元", "1800"].map((value) => toNumber(value, strategy)) };
  })())`)), { unit: "元", values: [2200, 2200, 1800] });
  // One value is enough: the 80 percent threshold decides the type, not the strategy.
  assert.equal(evaluate('buildNumberStrategy(["张三", "李四", "王五", "3人"]).unit'), "人");
  assert.deepEqual(JSON.parse(evaluate('JSON.stringify(buildNumberStrategy(["5KG", "2kg", "1.2万", "３"]))')), {
    unit: "kg", unitConflict: false, units: { kg: 2 }, allowMagnitude: true, fullwidth: true
  });
});

test("placeholder spellings never convert or give number strategy evidence (contract 17 premise)", () => {
  const { evaluate } = createScriptContext();
  const result = JSON.parse(evaluate(`JSON.stringify((() => {
    const toFullWidth = (text) => text.replace(/[!-~]/g, (char) => String.fromCharCode(char.charCodeAt(0) + 0xfee0));
    const keys = [...PLACEHOLDER_VALUES.common, ...PLACEHOLDER_VALUES.maybe];
    const forms = keys.flatMap((key) => [key, toFullWidth(key), " " + toFullWidth(key) + " "]);
    const strategies = [
      {},
      { unit: "", unitConflict: false, units: {}, allowMagnitude: true, fullwidth: true },
      ...NUMBER_UNITS.map((unit) => ({ unit: unit.toLowerCase(), unitConflict: false, units: {}, allowMagnitude: true, fullwidth: true }))
    ];
    return {
      formCount: forms.length,
      converted: forms.filter((form) => strategies.some((strategy) => toNumber(form, strategy) !== null)),
      matched: forms.filter((form) => matchNumberWithUnit(form.trim()) || matchNumberWithUnit(toHalfwidthNumberText(form.trim()))),
      strategy: buildNumberStrategy(forms)
    };
  })())`));
  assert.equal(result.formCount, 57);
  assert.deepEqual(result.converted, []);
  assert.deepEqual(result.matched, []);
  assert.deepEqual(result.strategy, { unit: "", unitConflict: false, units: {}, allowMagnitude: false, fullwidth: false });
});

test("getFieldNumberParser() is the plain parser unless the field's strategy has a switch", () => {
  const { evaluate } = createScriptContext();
  evaluate(`state.profiles = [
    { field: "金额", numberStrategy: buildNumberStrategy(["2200元", "800"]) },
    { field: "数量", numberStrategy: buildNumberStrategy(["3", "5"]) },
    { field: "备注" }
  ]`);
  assert.equal(evaluate('getFieldNumberParser("数量") === toPlainNumber'), true);
  assert.equal(evaluate('getFieldNumberParser("备注") === toPlainNumber'), true);
  assert.equal(evaluate('getFieldNumberParser("不存在") === toPlainNumber'), true);
  assert.equal(evaluate('getFieldNumberParser("金额") === toPlainNumber'), false);
  assert.deepEqual(JSON.parse(evaluate('JSON.stringify(["2200元", "800", "5kg", "N/A"].map(getFieldNumberParser("金额")))')), [2200, 800, null, null]);
  assert.equal(evaluate('parseFieldNumber("金额", "3500元")'), 3500);
  assert.equal(evaluate('parseFieldNumber("数量", "3500元")'), null);
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

test("number display keeps small values visible and never shows negative zero", () => {
  const { evaluate } = createScriptContext();

  assert.equal(evaluate("formatNumber(-0)"), "0");
  assert.equal(evaluate("formatNumber(-0.001)"), "-0.001");
  assert.equal(evaluate("formatNumber(0.000123)"), "0.000123");
  assert.equal(evaluate("formatNumber(0.123456)"), "0.123");
  assert.equal(evaluate("formatNumber(-0.00075)"), "-0.00075");
  assert.equal(evaluate("formatNumber(-(0.1 + 0.2 - 0.3))"), "0");
  assert.equal(evaluate("formatNumber(1e-11)"), "0");
  assert.equal(evaluate("formatNumber(-1e-11)"), "0");

  // Values at or above 1 keep the existing two-decimal display.
  assert.equal(evaluate("formatNumber(1234.5)"), "1,234.5");
  assert.equal(evaluate("formatNumber(500 / 3)"), "166.67");
  assert.equal(evaluate("formatNumber(10)"), "10");
  assert.equal(evaluate("formatNumber(-2.5)"), "-2.5");
  assert.equal(evaluate("formatNumber(NaN)"), "-");
  assert.equal(evaluate("formatNumber(Infinity)"), "-");
  assert.equal(evaluate("formatNumber(-Infinity)"), "-");

  assert.equal(evaluate("formatPercent(-0.00001)"), "0.0%");
  assert.equal(evaluate("formatPercent(-0)"), "0.0%");
  assert.equal(evaluate("formatPercent(0.25)"), "25.0%");
  assert.equal(evaluate("formatPercent(NaN)"), "0%");

  evaluate('state.profiles = [{ field: "比例", numericFormat: "percent" }]');
  assert.equal(evaluate('formatFieldNumber("比例", -0.00001)'), "0%");
  assert.equal(evaluate('formatFieldNumber("比例", -1e-11)'), "0%");
  assert.equal(evaluate('formatFieldNumber("比例", 0.3375)'), "33.75%");
  assert.equal(evaluate('formatFieldNumber("比例", -0.25)'), "-25%");
  assert.equal(evaluate('formatFieldNumber("比例", NaN)'), "-");
});

test("Excel number formats count as percent only when % is a real scale token", () => {
  const { evaluate } = createScriptContext();

  for (const format of ["0%", "0.00%", "#,##0.0%", "[Red]0%"]) {
    assert.equal(evaluate(`isPercentNumberFormat(${JSON.stringify(format)})`), true, format);
  }
  for (const format of ['0"%"', "0\\%", '"rate %" 0', "General", "#,##0.00", "¥#,##0.00", "yyyy-mm-dd", "", null, undefined]) {
    assert.equal(evaluate(`isPercentNumberFormat(${JSON.stringify(format) ?? "undefined"})`), false, String(format));
  }
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

test("ID names cover row-number headers and Chinese names ending in ID", () => {
  const { context, evaluate } = createScriptContext();
  const isIdName = (name) => evaluate(`isIdFieldName(${JSON.stringify(name)})`);

  for (const name of ["序号", "行号", "行序号", "Index", "IDX", "No.", "Row No.", "Row No", "#", " 序号 ", "学生ID", "客户 ID", "视频id"]) {
    assert.equal(isIdName(name), true, name);
  }
  for (const name of ["序号说明", "index_value", "indexes", "noise", "No. of employees", "指数", "身份证明", "ID卡", "IDEA", "Android", "学号数量"]) {
    assert.equal(isIdName(name), false, name);
  }
  // Already true before v2.1.4 (deferred finding 6 of v2.1.3); kept unchanged on purpose.
  assert.equal(isIdName("Number of employees"), true);
  assert.equal(isIdName("Key metric"), true);

  context.__countRows = Array.from({ length: 12 }, (_, index) => ({ 学号数量: String(index + 1) }));
  assert.equal(evaluate('buildColumnProfile("学号数量", __countRows).typeKey'), "numeric");
});

test("profiles keep up to three distinct conversion failure samples per type", () => {
  const { context, evaluate } = createScriptContext();
  const amounts = ["N/A", " N/A ", "待定", "约500", "x", ...Array.from({ length: 20 }, (_, index) => String(index * 10))];
  context.__amountRows = amounts.map((value) => ({ 金额: value }));
  const amount = JSON.parse(evaluate('JSON.stringify(buildColumnProfile("金额", __amountRows))'));
  assert.equal(amount.typeKey, "numeric");
  assert.equal(amount.conversionFailureCount, 5);
  assert.deepEqual(amount.conversionFailureSamples, ["N/A", "待定", "约500"]);
  assert.deepEqual(amount.conversionFailureSamplesByType.numeric, ["N/A", "待定", "约500"]);

  const dates = ["下周一", "2024-13-45", ...Array.from({ length: 8 }, (_, index) => `2024-01-0${index + 1}`)];
  context.__dateRows = dates.map((value) => ({ 日期: value }));
  const date = JSON.parse(evaluate('JSON.stringify(buildColumnProfile("日期", __dateRows))'));
  assert.equal(date.typeKey, "date");
  assert.equal(date.conversionFailureCount, 2);
  assert.deepEqual(date.conversionFailureSamples, ["下周一", "2024-13-45"]);

  context.__textRows = ["甲", "乙", "丙"].map((value) => ({ 备注: value }));
  const text = JSON.parse(evaluate('JSON.stringify(buildColumnProfile("备注", __textRows))'));
  assert.equal(text.conversionFailureCount, 0);
  assert.deepEqual(text.conversionFailureSamples, []);
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

test("small integer ranges are counted per value instead of binned", () => {
  const { context, evaluate } = createScriptContext();
  const counts = (values, field = "") => {
    context.__values = values;
    return JSON.parse(evaluate(`JSON.stringify(buildDiscreteValueCounts(__values, ${JSON.stringify(field)}))`));
  };
  const labels = (result) => result.map((bin) => bin.label);
  const tallies = (result) => result.map((bin) => bin.count);

  const scores = counts([2, 3, 4, 9, 9, 3]);
  assert.deepEqual(labels(scores), ["2", "3", "4", "5", "6", "7", "8", "9"]);
  assert.deepEqual(tallies(scores), [1, 2, 1, 0, 0, 0, 0, 2]);
  assert.deepEqual(labels(counts([1, 2, 3, 4, 5, 3])), ["1", "2", "3", "4", "5"]);
  assert.deepEqual(counts([4, 4, 4]), [{ label: "4", count: 3 }]);
  assert.deepEqual(labels(counts([-2, 0, 3])), ["-2", "-1", "0", "1", "2", "3"]);
  assert.equal(counts([0, 11]).length, 12);
  assert.equal(counts([0, 12]), null);
  assert.equal(counts([1, 2.5, 3]), null);
  assert.equal(counts([]), null);
  assert.deepEqual(labels(counts([1000000, 1000003])), ["1,000,000", "1,000,001", "1,000,002", "1,000,003"]);
  assert.equal(counts([-Number.MAX_VALUE, Number.MAX_VALUE]), null);

  evaluate('state.profiles = [{ field: "比例", numericFormat: "percent" }]');
  assert.equal(counts([0, 1], "比例"), null);
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

  // The coarse range check allows HEADER_SCAN_ROWS extra rows for title rows above the header.
  const rejectedDataRows = evaluate("IMPORT_LIMITS.maxRows + HEADER_SCAN_ROWS + 1");
  context.XLSX = {
    utils: {
      decode_range() {
        return { s: { r: 0, c: 0 }, e: { r: rejectedDataRows, c: 0 } };
      }
    }
  };
  const rowLimitMessage = evaluate(`(() => {
    try {
      validateExcelWorksheetLimits({"!ref": "A1:A${rejectedDataRows + 1}"}, "Large Sheet");
      return "";
    } catch (error) {
      return error.message;
    }
  })()`);
  assert.match(rowLimitMessage, /100,000/);
});

test("the coarse Excel range check leaves room for title rows above the header", () => {
  const { context, evaluate } = createScriptContext();
  const allowedDataRows = evaluate("IMPORT_LIMITS.maxRows + HEADER_SCAN_ROWS");
  context.XLSX = {
    utils: {
      decode_range() {
        return { s: { r: 0, c: 0 }, e: { r: allowedDataRows, c: 0 } };
      }
    }
  };
  assert.doesNotThrow(() => evaluate(`validateExcelWorksheetLimits({"!ref": "A1:A${allowedDataRows + 1}"}, "Large Sheet")`));
});

test("Excel sheets over the row limit below a first-row header are rejected when built", () => {
  const { context, evaluate } = createScriptContext();
  let lastDataRow = 0;
  context.XLSX = {
    utils: {
      decode_range() {
        return { s: { r: 0, c: 0 }, e: { r: lastDataRow, c: 0 } };
      }
    }
  };
  const errorMessage = (expression) => evaluate(`(() => {
    try { ${expression}; return ""; } catch (error) { return error.message; }
  })()`);

  // One header row plus maxRows + 1 data rows: inside the coarse margin, over the exact limit.
  lastDataRow = evaluate("IMPORT_LIMITS.maxRows + 1");
  assert.equal(errorMessage('validateExcelWorksheetLimits({ "!ref": "A1" }, "大表")'), "");
  const builtMessage = errorMessage(`(() => {
    const rows = [["值"], ...Array.from({ length: ${lastDataRow} }, (_, index) => [index])];
    buildExcelTabularData(rows, rows, "大表");
  })()`);

  lastDataRow = evaluate("IMPORT_LIMITS.maxRows + HEADER_SCAN_ROWS + 1");
  const coarseMessage = errorMessage('validateExcelWorksheetLimits({ "!ref": "A1" }, "大表")');
  assert.match(coarseMessage, /100,000/);
  assert.equal(builtMessage, coarseMessage);
});

// detectHeaderRow() only picks which row is the header; rows are written the way
// PapaParse header:false (CSV strings) or sheet_to_json raw values (Excel) give them.
function detectedHeader(matrix, kind) {
  const { context, evaluate } = createScriptContext();
  context.__headerMatrix = matrix;
  context.__headerKind = kind;
  const result = JSON.parse(evaluate("JSON.stringify(detectHeaderRow(__headerMatrix, __headerKind))"));
  return { ...result, headerRowNumber: result.headerIndex + 1 };
}

const GRADE_HEADER = ["学号", "姓名", "语文", "数学"];
const GRADE_ROWS = [["S01", "张三", "90", "85"], ["S02", "李四", "78", "92"], ["S03", "王五", "88", "70"]];
const ATTENDANCE_HEADER = ["姓名", "部门", "出勤天数", "迟到次数"];
const ATTENDANCE_ROWS = [["张三", "销售部", 22, 1], ["李四", "技术部", 21, 0], ["王五", "财务部", 20, 2]];

test("header detection skips a CSV title row padded with commas", () => {
  const result = detectedHeader([["初二(3)班期中成绩", "", "", ""], GRADE_HEADER, ...GRADE_ROWS], "csv");
  assert.equal(result.status, "skipped");
  assert.equal(result.headerRowNumber, 2);
  assert.deepEqual(result.skippedIndexes, [0]);
});

test("header detection skips a CSV title row without commas", () => {
  const result = detectedHeader([["初二(3)班期中成绩"], GRADE_HEADER, ...GRADE_ROWS], "csv");
  assert.equal(result.status, "skipped");
  assert.equal(result.headerRowNumber, 2);
});

test("header detection skips a title, a unit note and a blank row", () => {
  const result = detectedHeader([["期中成绩", "", "", ""], ["单位：分", "", "", ""], ["", "", "", ""], GRADE_HEADER, ...GRADE_ROWS], "csv");
  assert.equal(result.status, "skipped");
  assert.equal(result.headerRowNumber, 4);
  assert.deepEqual(result.skippedIndexes, [0, 1, 2]);
});

test("header detection skips an Excel merged title, report date and blank row", () => {
  const result = detectedHeader([
    ["门店报表", "", "", ""],
    ["制表日期：2024-08-01", "", "", ""],
    ["", "", "", ""],
    ["门店", "城市", "销售额", "毛利率"],
    ["A店", "上海", 862400, 0.3],
    ["B店", "杭州", 735200, 0.28],
    ["C店", "南京", 912800, 0.28]
  ], "excel");
  assert.equal(result.status, "skipped");
  assert.equal(result.headerRowNumber, 4);
});

test("header detection skips a CSV row of delimiters only", () => {
  const result = detectedHeader([["", "", "", ""], GRADE_HEADER, ...GRADE_ROWS], "csv");
  assert.equal(result.status, "skipped");
  assert.equal(result.headerRowNumber, 2);
  assert.deepEqual(result.skippedIndexes, [0]);
});

test("header detection skips an Excel title row with a unit note", () => {
  const result = detectedHeader([["考勤表", "单位：天", "", ""], ATTENDANCE_HEADER, ...ATTENDANCE_ROWS], "excel");
  assert.equal(result.status, "skipped");
  assert.equal(result.headerRowNumber, 2);
});

test("header detection skips an Excel title and a school/date row", () => {
  const result = detectedHeader([
    ["期末成绩", "", "", ""],
    ["学校：一中", "日期：2024-06", "", ""],
    ATTENDANCE_HEADER,
    ...ATTENDANCE_ROWS
  ], "excel");
  assert.equal(result.status, "skipped");
  assert.equal(result.headerRowNumber, 3);
  assert.deepEqual(result.skippedIndexes, [0, 1]);
});

test("header detection keeps a valid first-row header and single-column files", () => {
  const normal = detectedHeader([GRADE_HEADER, ...GRADE_ROWS], "csv");
  assert.equal(normal.status, "unchanged");
  assert.equal(normal.headerRowNumber, 1);
  assert.deepEqual(normal.skippedIndexes, []);
  const singleColumn = detectedHeader([["城市"], ["上海"], ["北京"], ["广州"]], "csv");
  assert.equal(singleColumn.status, "unchanged");
  assert.equal(singleColumn.headerRowNumber, 1);
});

test("header detection does not skip a header with one blank cell", () => {
  const csv = detectedHeader([["姓名", "", "出勤天数", "迟到次数"], ["张三", "销售部", "22", "1"], ["李四", "技术部", "21", "0"]], "csv");
  assert.equal(csv.status, "fallback");
  assert.equal(csv.headerRowNumber, 1);
  assert.equal(csv.attempted, false);
});

test("header detection falls back when the data below is a single column", () => {
  const excel = detectedHeader([["城市", ""], ["上海", ""], ["北京", ""], ["广州", "待核实"], ["深圳", ""], ["杭州", ""]], "excel");
  assert.equal(excel.status, "fallback");
  assert.equal(excel.headerRowNumber, 1);
  assert.equal(excel.attempted, true);
  const csv = detectedHeader([["城市"], ["上海"], ["北京"], ["广州", "待核实"], ["深圳"], ["杭州"]], "csv");
  assert.equal(csv.status, "fallback");
  assert.equal(csv.headerRowNumber, 1);
  assert.equal(csv.attempted, true);
});

test("header detection rejects a numeric candidate cell above a numeric column", () => {
  const result = detectedHeader([["姓名", ""], ["学生1", "61"], ["学生2", "75"], ["学生3", "82"], ["学生4", "90"], ["学生5", "68"]], "csv");
  assert.equal(result.status, "fallback");
  assert.equal(result.headerRowNumber, 1);
  assert.equal(result.attempted, true);
});

test("header detection rejects a candidate cell that repeats in its column", () => {
  const departments = [["姓名", ""], ["学生1", "销售部"], ["学生2", "技术部"], ["学生3", "销售部"], ["学生4", "财务部"], ["学生5", "技术部"]];
  for (const kind of ["csv", "excel"]) {
    const result = detectedHeader(departments, kind);
    assert.equal(result.status, "fallback", kind);
    assert.equal(result.headerRowNumber, 1, kind);
  }
});

test("header detection does not skip a two-cell first row without a colon", () => {
  const blankScores = detectedHeader([["姓名", "性别", "", ""], ["张三", "男", "90", "85"], ["李四", "女", "78", "92"], ["王五", "男", "88", "70"]], "csv");
  assert.equal(blankScores.status, "fallback");
  assert.equal(blankScores.attempted, false);
  const distinctText = detectedHeader([["姓名", "工号", "", ""], ["张三", "E01", "销售部", "上海"], ["李四", "E02", "技术部", "北京"], ["王五", "E03", "财务部", "广州"]], "csv");
  assert.equal(distinctText.status, "fallback");
  assert.equal(distinctText.attempted, false);
  const dates = detectedHeader([
    ["姓名", "部门", "", ""],
    ["张三", "销售部", new Date(Date.UTC(2024, 0, 5)), 90],
    ["李四", "技术部", new Date(Date.UTC(2024, 1, 6)), 85],
    ["王五", "财务部", new Date(Date.UTC(2024, 2, 7)), 78]
  ], "excel");
  assert.equal(dates.status, "fallback");
  assert.equal(dates.attempted, false);
});

test("header detection falls back when the header repeats on a later page", () => {
  const result = detectedHeader([
    ["期中成绩", "", "", ""],
    GRADE_HEADER,
    ...GRADE_ROWS,
    GRADE_HEADER,
    ["S04", "赵六", "81", "77"],
    ["S05", "钱七", "69", "88"]
  ], "csv");
  assert.equal(result.status, "fallback");
  assert.equal(result.headerRowNumber, 1);
  assert.equal(result.attempted, true);
});

test("header detection skips a name row above distinct text pairs (已知残留，v2.3 手动选择表头行兜底)", () => {
  const result = detectedHeader([["姓名", ""], ["学生1", "备注1"], ["学生2", "备注2"], ["学生3", "备注3"], ["学生4", "备注4"]], "csv");
  assert.equal(result.status, "skipped");
  assert.equal(result.headerRowNumber, 2);
});

test("skipped-row previews join non-empty cells and cut long text at 20 characters", () => {
  const { evaluate } = createScriptContext();
  assert.equal(evaluate('formatSkippedRowPreview(["学校：一中", "日期：2024-06"])'), "学校：一中 日期：2024-06");
  assert.equal(evaluate(`formatSkippedRowPreview(["${"表".repeat(20)}"])`), "表".repeat(20));
  assert.equal(evaluate(`formatSkippedRowPreview(["${"表".repeat(21)}"])`), `${"表".repeat(20)}…`);
  assert.equal(evaluate('formatSkippedRowPreview(["😀".repeat(21)])'), `${"😀".repeat(20)}…`);
});

test("the header note lists at most two skipped rows and adds 等 for more", () => {
  const { evaluate } = createScriptContext();
  const note = (detection) => evaluate(`formatHeaderDetectionNote(${JSON.stringify(detection)})`);
  assert.equal(note(null), "");
  assert.equal(note({ headerRowNumber: 2, skippedRows: [] }), "表头在第 2 行，已跳过上方 1 行。");
  assert.equal(note({ headerRowNumber: 3, skippedRows: [{ rowNumber: 1, preview: "成绩单" }] }), "表头在第 3 行，已跳过上方 2 行：「成绩单」。");
  assert.equal(
    note({ headerRowNumber: 5, skippedRows: [{ rowNumber: 1, preview: "甲" }, { rowNumber: 2, preview: "乙" }, { rowNumber: 3, preview: "丙" }] }),
    "表头在第 5 行，已跳过上方 4 行：「甲」「乙」等。"
  );
});

test("CSV parse warnings read in Chinese by code, with and without a line", () => {
  const { evaluate } = createScriptContext();
  const message = (code, details) => evaluate(`formatCsvWarningMessage(${JSON.stringify(code)}, ${JSON.stringify(details)})`);
  assert.equal(message("TooFewFields", { line: 3, cellCount: 1, headerCount: 3 }), "第 3 行只有 1 列（表头有 3 列），缺少的单元格按空值处理。");
  assert.equal(message("TooFewFields", { line: 3, cellCount: null, headerCount: 3 }), "第 3 行的列数少于表头，缺少的单元格按空值处理。");
  assert.equal(message("TooFewFields", { line: null, cellCount: 1, headerCount: 3 }), "有一条记录的列数少于表头（只有 1 列，表头有 3 列），缺少的单元格按空值处理。");
  assert.equal(message("TooFewFields", { line: null, cellCount: null, headerCount: 3 }), "有一条记录的列数少于表头，缺少的单元格按空值处理。");
  assert.equal(message("InvalidQuotes", { line: 7 }), "第 7 行有不规范的引号（引号后面紧跟了其他字符），这一行可能被读错。");
  assert.equal(message("InvalidQuotes", { line: null }), "有一处不规范的引号（引号后面紧跟了其他字符），所在的记录可能被读错。");
  assert.equal(message("UndetectableDelimiter", { separator: ";" }), "整份文件只读出 1 列，但表头里有分号（；）。如果原文件是用分号分隔的，请另存为逗号分隔的 CSV 后重新上传。");
  assert.equal(message("UndetectableDelimiter", { separator: "\t" }), "整份文件只读出 1 列，但表头里有制表符。如果原文件是用制表符分隔的，请另存为逗号分隔的 CSV 后重新上传。");
  assert.equal(message("UndetectableDelimiter", { separator: "|" }), "整份文件只读出 1 列，但表头里有竖线（|）。如果原文件是用竖线分隔的，请另存为逗号分隔的 CSV 后重新上传。");
  assert.equal(message("SomethingElse", {}), "");
});

test("line numbers at offsets count CR, LF and CRLF like the header row numbers did", () => {
  const { context, evaluate } = createScriptContext();
  const texts = ["", "a", "a\nb", "a\r\nb", "a\rb", "\r\n\r\n", "\r\r\n\n", "x\n\ry\r\n", "\"q\r\nq\",1\n2\r3"];
  let seed = 20261028;
  const next = () => (seed = (seed * 1103515245 + 12345) % 2147483648);
  for (let index = 0; index < 200; index += 1) {
    let text = "";
    for (let length = next() % 40; length > 0; length -= 1) text += ["a", ",", "\r", "\n", "\r\n", "\""][next() % 6];
    texts.push(text);
  }
  // The rule findCsvHeaderBelowTitleRows() used per offset before the counter was shared.
  const lineNumberAt = (text, offset) => 1 + (text.slice(0, offset).match(/\r\n|\r|\n/g) || []).length;
  texts.forEach((text) => {
    const offsets = Array.from({ length: text.length + 1 }, (_, offset) => offset);
    context.__lineText = text;
    context.__lineOffsets = offsets;
    assert.deepEqual(
      JSON.parse(evaluate("JSON.stringify(lineNumbersAtOffsets(__lineText, __lineOffsets))")),
      offsets.map((offset) => lineNumberAt(text, offset)),
      JSON.stringify(text)
    );
  });
  context.__lineOffsets = [5, 2, 7];
  assert.deepEqual(JSON.parse(evaluate('JSON.stringify(lineNumbersAtOffsets("a\\r\\nb\\nc\\rd", __lineOffsets))')), [3, 2, 4]);
});

// buildIdDuplicateSummary() on plain rows; a row's "line" is its index + 2, as if
// the header sat on line 1.
function idDuplicateSummary(rows, typeKeys, getRowNumber = "(row) => __idRows.indexOf(row) + 2") {
  const { context, evaluate } = createScriptContext();
  context.__idRows = rows;
  context.__idFields = Object.keys(typeKeys);
  context.__idProfiles = Object.entries(typeKeys).map(([field, typeKey]) => ({ field, typeKey }));
  return JSON.parse(evaluate(`JSON.stringify(buildIdDuplicateSummary(__idRows, __idFields, __idProfiles, ${getRowNumber}))`));
}

test("ID duplicates trim values, keep case and skip missing values", () => {
  const rows = [" E1012", "E1012", "E1012 ", "e1012", "", "  ", undefined].map((工号) => ({ 工号, 姓名: "张三" }));
  assert.deepEqual(idDuplicateSummary(rows, { 工号: "id", 姓名: "category" }), [{
    field: "工号",
    nonMissingCount: 4,
    distinctCount: 2,
    uniqueRatio: 0.5,
    mostlyRepeated: false,
    duplicateValueCount: 1,
    values: [{ value: "E1012", count: 3, rowNumbers: [2, 3, 4], identicalRowCount: 0, identicalGroupCount: 0 }]
  }]);
});

test("ID duplicates follow the field order and only check ID fields", () => {
  const rows = [{ a: "1", b: "x", c: "k" }, { a: "1", b: "x", c: "k" }];
  assert.deepEqual(idDuplicateSummary(rows, { b: "id", c: "category", a: "id" }).map((item) => item.field), ["b", "a"]);
  assert.deepEqual(idDuplicateSummary(rows, { a: "numeric", b: "ignore" }), []);
});

test("ID duplicates sort by count, then by first appearance", () => {
  const rows = ["A", "B", "B", "B", "C", "A", "C", "D"].map((id) => ({ id }));
  const [summary] = idDuplicateSummary(rows, { id: "id" });
  assert.deepEqual(summary.values.map((item) => [item.value, item.count, item.rowNumbers]), [
    ["B", 3, [3, 4, 5]],
    ["A", 2, [2, 7]],
    ["C", 2, [6, 8]]
  ]);
});

test("mostly repeated ID fields need 12 values and a unique share under one half", () => {
  const field = (values) => idDuplicateSummary(values.map((id) => ({ id })), { id: "id" })[0];
  const fiveOfTwelve = field(Array.from({ length: 12 }, (_, index) => `U${index % 5}`));
  assert.equal(fiveOfTwelve.mostlyRepeated, true);
  assert.equal(fiveOfTwelve.duplicateValueCount, 5);
  assert.deepEqual(fiveOfTwelve.values, []);
  const sixOfTwelve = field(Array.from({ length: 12 }, (_, index) => `U${index % 6}`));
  assert.equal(sixOfTwelve.mostlyRepeated, false);
  assert.equal(sixOfTwelve.values.length, 6);
  const oneOfEleven = field(Array.from({ length: 11 }, () => "U0"));
  assert.equal(oneOfEleven.mostlyRepeated, false);
  assert.deepEqual(oneOfEleven.values.map((item) => item.count), [11]);
});

test("ID duplicates count identical whole rows per value, in groups", () => {
  const rows = [["1012", "销售"], ["1012", "销售"], ["1012", "财务"], ["1012", "销售"], ["1012", "财务"], ["1013", "技术"], ["1013", "市场"], ["1014", "技术"], ["1014", "技术"]]
    .map(([工号, 部门]) => ({ 工号, 部门 }));
  const [summary] = idDuplicateSummary(rows, { 工号: "id", 部门: "category" });
  assert.deepEqual(summary.values.map((item) => [item.value, item.identicalRowCount, item.identicalGroupCount]), [
    ["1012", 5, 2],
    ["1013", 0, 0],
    ["1014", 2, 1]
  ]);
});

test("ID duplicates list at most 10 values with at most 10 row numbers each", () => {
  // 40 unique IDs keep the unique share at 52 / 77, so the field is not mostly repeated.
  const ids = [
    ...Array.from({ length: 15 }, () => "S000"),
    ...Array.from({ length: 11 }, (_, index) => [`S${index + 1}`, `S${index + 1}`]).flat(),
    ...Array.from({ length: 40 }, (_, index) => `S${index + 100}`)
  ];
  const [summary] = idDuplicateSummary(ids.map((id) => ({ id })), { id: "id" });
  assert.equal(summary.duplicateValueCount, 12);
  assert.equal(summary.values.length, 10);
  assert.equal(summary.values[0].count, 15);
  assert.deepEqual(summary.values[0].rowNumbers, [2, 3, 4, 5, 6, 7, 8, 9, 10, 11]);
});

test("ID duplicates keep null for rows without a row number", () => {
  const [summary] = idDuplicateSummary([{ id: "A" }, { id: "A" }], { id: "id" }, "() => null");
  assert.deepEqual(summary.values[0].rowNumbers, [null, null]);
});

test("ID duplicate texts: row numbers, identical rows, overflow and the repeated-field hint", () => {
  const { evaluate } = createScriptContext();
  assert.equal(evaluate("formatRowNumberList([6, 19], 2)"), "6、19");
  assert.equal(evaluate("formatRowNumberList([2, 3, 4, 5, 6, 7, 8, 9, 10, 11], 15)"), "2、3、4、5、6、7、8、9、10、11 等 15 个");
  assert.equal(evaluate("formatRowNumberList([null, null], 2)"), "—");
  assert.equal(evaluate("formatIdenticalRows(0, 0)"), "");
  assert.equal(evaluate("formatIdenticalRows(2, 1)"), "其中 2 行内容完全相同");
  assert.equal(evaluate("formatIdenticalRows(5, 2)"), "其中 5 行内容完全相同（分 2 组）");
  assert.equal(evaluate("formatIdDuplicateOverflow(10, 10)"), "");
  assert.equal(evaluate("formatIdDuplicateOverflow(12, 10)"), "共 12 个重复值，只列出前 10 个");
  assert.equal(
    evaluate('formatMostlyRepeatedIdHint({ field: "会话ID", distinctCount: 30, nonMissingCount: 120, uniqueRatio: 0.25 })'),
    "「会话ID」的值大多重复出现（30 个不同值，共 120 行，唯一率 25.0%），像是同一对象的多条记录，不逐个列出。如果这一列应该每行唯一，请检查数据；不需要检查的话，可以在字段配置里改类型。"
  );
});

test("the placeholder list has a common tier and a maybe-valid tier", () => {
  const { evaluate } = createScriptContext();
  assert.deepEqual(JSON.parse(evaluate("JSON.stringify(PLACEHOLDER_VALUES)")), {
    common: ["N/A", "n/a", "#N/A", "NULL", "null", "Null", "None", "none", "NaN", "nan", "NAN", "—", "——", "--"],
    maybe: ["无", "暂无", "-", "/", "NA"]
  });
  assert.equal(evaluate('placeholderTier("N/A")'), "common");
  assert.equal(evaluate('placeholderTier("——")'), "common");
  assert.equal(evaluate('placeholderTier("无")'), "maybe");
  assert.equal(evaluate('placeholderTier("NA")'), "maybe");
  assert.equal(evaluate('placeholderTier("缺考")'), "");
});

test("placeholder keys trim, apply NFKC and trim again, then match the whole cell", () => {
  const { context, evaluate } = createScriptContext();
  const key = (value) => {
    context.__placeholderValue = value;
    return evaluate("placeholderKey(__placeholderValue)");
  };
  for (const value of JSON.parse(evaluate("JSON.stringify([...PLACEHOLDER_VALUES.common, ...PLACEHOLDER_VALUES.maybe])"))) {
    assert.equal(key(value), value);
    assert.equal(key(` ${value}\t`), value);
  }
  assert.equal(key("Ｎ／Ａ"), "N/A");
  const ideographicSpace = String.fromCharCode(0x3000);
  assert.equal(key(`${ideographicSpace}N/A${ideographicSpace}`), "N/A");
  assert.equal(key("－－"), "--");
  assert.equal(key("ＮａＮ"), "NaN");
  assert.equal(key("／"), "/");
  // Not on the list, or only part of the cell.
  for (const value of ["Nan", "NAn", "na", "缺考", "未知", "面议", "0", "N/A 待补", "NA/", "---", "–", "", "  ", null, undefined, 0, false]) {
    assert.equal(key(value), "", String(value));
  }
});

test("placeholder summaries list each field and key with tier, raw texts and first rows", () => {
  const { context, evaluate } = createScriptContext();
  const amounts = ["100", "N/A", "Ｎ／Ａ", " N/A ", "—", "缺考", ...Array.from({ length: 10 }, () => "N/A"), "200"];
  context.__placeholderRows = amounts.map((amount, index) => ({
    金额: amount,
    备注: index === 0 ? "N/A" : index === 1 ? "无" : "好",
    姓名: index === 2 ? "Nan" : `学生${index}`
  }));
  const summary = JSON.parse(evaluate(`JSON.stringify((() => {
    const rows = __placeholderRows;
    const fields = ["金额", "备注", "姓名"];
    const profiles = fields.map((field) => buildColumnProfile(field, rows));
    profiles[1].typeKey = "ignore";
    const rowNumber = new Map(rows.map((row, index) => [row, index === 4 ? null : index + 2]));
    return buildPlaceholderSummary(rows, fields, profiles, { 金额: ["—"], 备注: ["无"] }, (row) => rowNumber.get(row) ?? null);
  })())`));
  assert.deepEqual(summary, [
    { field: "金额", key: "N/A", tier: "common", count: 13, rawValues: ["N/A", "Ｎ／Ａ"], rowNumbers: [3, 4, 5, 8, 9, 10, 11, 12, 13, 14], confirmed: false, fieldIgnored: false },
    { field: "金额", key: "—", tier: "common", count: 1, rawValues: ["—"], rowNumbers: [null], confirmed: true, fieldIgnored: false },
    // An ignored field shows only the keys confirmed before it was ignored.
    { field: "备注", key: "无", tier: "maybe", count: 1, rawValues: ["无"], rowNumbers: [3], confirmed: true, fieldIgnored: true }
  ]);
});

test("placeholders never change the date strategy of a field without a year name", () => {
  const { context, evaluate } = createScriptContext();
  // Full-width and compatibility forms that NFKC turns into list keys:
  // Ｎ／Ａ － ／ －－ ＮａＮ ⽆ ﹘ ﹣
  const compatibilityForms = [[0xff2e, 0xff0f, 0xff21], [0xff0d], [0xff0f], [0xff0d, 0xff0d], [0xff2e, 0xff41, 0xff2e], [0x2f46], [0xfe58], [0xfe63]]
    .map((codePoints) => String.fromCodePoint(...codePoints));
  context.__compatibilityForms = compatibilityForms;
  assert.deepEqual(JSON.parse(evaluate("JSON.stringify(__compatibilityForms.map(placeholderKey))")), ["N/A", "-", "/", "--", "NaN", "无", "—", "-"]);
  context.__dateSets = [
    ["13/02/2024", "14/03/2024", "2024-01-05"], // dmy evidence
    ["02/13/2024", "03/14/2024", "2024年1月5日"], // mdy evidence
    ["13/02/2024", "02/14/2024", "05/06/2024", "2024"] // both, plus a bare year
  ];
  const result = JSON.parse(evaluate(`JSON.stringify((() => {
    const placeholders = [...PLACEHOLDER_VALUES.common, ...PLACEHOLDER_VALUES.maybe, ...__compatibilityForms];
    return __dateSets.map((dates) => {
      const base = buildDateStrategy("日期", dates);
      return {
        order: base.order,
        changed: placeholders.filter((value) => JSON.stringify(buildDateStrategy("日期", [...dates, value])) !== JSON.stringify(base)),
        allowYearOnly: placeholders.some((value) => buildDateStrategy("日期", [...dates, value]).allowYearOnly),
        allAtOnce: JSON.stringify(buildDateStrategy("日期", [...dates, ...placeholders])) === JSON.stringify(base)
      };
    });
  })())`));
  assert.deepEqual(result, [
    { order: "dmy", changed: [], allowYearOnly: false, allAtOnce: true },
    { order: "mdy", changed: [], allowYearOnly: false, allAtOnce: true },
    { order: "conflict", changed: [], allowYearOnly: false, allAtOnce: true }
  ]);
});

// v2.6.1 fast paths of isMissing() and toPlainNumber().

// The v2.6.0 body of isMissing(), the reference for the fast path.
const legacyIsMissing = (value) => value === null || value === undefined || String(value).trim() === "";

test("isMissing() agrees with trim() for every UTF-16 code unit alone, at either end and inside", () => {
  const { evaluate } = createScriptContext();
  const isMissing = evaluate("isMissing");
  let compared = 0;
  for (let code = 0; code <= 0xFFFF; code += 1) {
    const ch = String.fromCharCode(code);
    for (const value of [ch, ch + ch, `a${ch}`, `${ch}a`, `a${ch}a`, `1${ch}`, `${ch}1`, ` ${ch}12 `, `${ch}1${ch}`]) {
      assert.equal(isMissing(value), legacyIsMissing(value), `U+${code.toString(16).toUpperCase().padStart(4, "0")} in ${JSON.stringify(value)}`);
      compared += 1;
    }
  }
  assert.ok(compared > 500000);
});

test("isMissing() treats exactly the 25 trim() characters as blank and keeps U+0085, U+180E, U+200B", () => {
  const { evaluate } = createScriptContext();
  const isMissing = evaluate("isMissing");
  const blanks = [0x9, 0xA, 0xB, 0xC, 0xD, 0x20, 0xA0, 0x1680, 0x2000, 0x2001, 0x2002, 0x2003, 0x2004, 0x2005, 0x2006, 0x2007, 0x2008, 0x2009, 0x200A, 0x2028, 0x2029, 0x202F, 0x205F, 0x3000, 0xFEFF];
  blanks.forEach((code) => assert.equal(isMissing(String.fromCharCode(code)), true, `U+${code.toString(16)}`));
  [0x85, 0x180E, 0x200B, 0x200C, 0x200D, 0x2060].forEach((code) => assert.equal(isMissing(String.fromCharCode(code)), false, `U+${code.toString(16)}`));
  assert.equal(isMissing(""), true);
  assert.equal(isMissing(null), true);
  assert.equal(isMissing(undefined), true);
  assert.equal(isMissing(0), false);
  assert.equal(isMissing([]), true);
  // Text above U+3000 takes the fast path; each result still matches the reference.
  ["华东", "\u3001", "\ud800", "\u3000华东", "华东\u3000", "\ufeff", "\ufeff华东"].forEach((value) => {
    assert.equal(isMissing(value), legacyIsMissing(value), JSON.stringify(value));
  });
  assert.equal(isMissing("华东"), false);
  assert.equal(isMissing("\ud800"), false);
  assert.equal(isMissing("\u3000"), true);
  assert.equal(isMissing("\ufeff"), true);
});

test("toPlainNumber() reads plain digit strings and never lets the fast path return Infinity or accept more than digits", () => {
  const { evaluate } = createScriptContext();
  const toPlainNumber = evaluate("toPlainNumber");
  assert.equal(toPlainNumber("0"), 0);
  assert.equal(toPlainNumber("007"), 7);
  assert.equal(toPlainNumber(" 42 "), 42);
  assert.equal(toPlainNumber("\u00a042\u3000"), 42);
  assert.equal(toPlainNumber("9".repeat(15)), 999999999999999);
  assert.equal(toPlainNumber("9".repeat(16)), 9999999999999999);
  assert.equal(toPlainNumber("1" + "0".repeat(308)), 1e308);
  assert.equal(toPlainNumber("9".repeat(309)), null);
  assert.equal(toPlainNumber("9".repeat(400)), null);
  assert.equal(toPlainNumber("1" + "0".repeat(309)), null);
  // Not ASCII digits, or not only digits: the full path decides, as before.
  ["１２３", "٣٤٥", "۱۲۳", "123abc", "abc123", "12-3", "1 2", "1\n2", "2024-01-01", "2024/1/1", "1_000", "0x10"].forEach((text) => assert.equal(toPlainNumber(text), null, text));
  assert.equal(toPlainNumber("1,234"), 1234);
  assert.equal(toPlainNumber("12 345"), 12345);
  assert.equal(toPlainNumber("123%"), 1.23);
  assert.equal(toPlainNumber("(123)"), -123);
  assert.equal(toPlainNumber(123), 123);
  assert.equal(Object.is(toPlainNumber(-0), 0), true);
});

test("ordinary text and digit strings are read without calling trim() twice", () => {
  const scope = createScriptContext();
  scope.evaluate(`globalThis.__trim = String.prototype.trim; String.prototype.trim = function () { globalThis.__trimCalls = (globalThis.__trimCalls || 0) + 1; return globalThis.__trim.call(this); };`);
  const trimCalls = () => scope.context.__trimCalls || 0;
  scope.evaluate('isMissing("East")');
  scope.evaluate('isMissing("12345")');
  assert.equal(trimCalls(), 0);
  // Chinese text takes the fast path too.
  scope.evaluate('isMissing("华东")');
  assert.equal(trimCalls(), 0);
  scope.evaluate('toPlainNumber("12345")');
  assert.equal(trimCalls(), 1);
  // Edge whitespace still takes the full path and is judged by trim().
  const before = trimCalls();
  assert.equal(scope.evaluate('isMissing(" a")'), false);
  assert.ok(trimCalls() > before);
});

// v2.6.1: the statistics table takes its median from the values it has already sorted.

test("median() and the statistics table agree for odd, even, unsorted and overflowing values", () => {
  const { context, evaluate } = createScriptContext();
  const median = evaluate("median");
  assert.equal(median([]), 0);
  assert.equal(median([3, 1, 2]), 2);
  assert.equal(median([5, 1, 4, 2]), 3);
  assert.equal(median([1e308, 1e308]), 1e308);
  assert.equal(median([1.7e308, 1.7e308]), 1.7e308);
  assert.equal(median([-1.7e308, 1.7e308]), 0);

  const unsorted = ["30", "10", "40", "20", "60", "50"];
  const huge = ["1.7e308", "1.7e308", "1.7e308", "1.7e308", "1.7e308", "1.7e308"];
  installDataset(context, evaluate, unsorted.map((value, index) => ({ 乱序: value, 巨大: huge[index] })));
  const stats = JSON.parse(evaluate("JSON.stringify(buildNumericStats(state.rows))"));
  assert.deepEqual(stats.map((item) => [item.field, item.median]), [["乱序", 35], ["巨大", 1.7e308]]);
});
