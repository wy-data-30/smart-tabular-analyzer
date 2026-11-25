const test = require("node:test");
const assert = require("node:assert/strict");
const { createScriptContext, APP_SCRIPT_FILES } = require("./test-context.cjs");
const { execFileSync } = require("node:child_process");
const fs = require("node:fs");
const path = require("node:path");
const englishDateCorpus = require("./fixtures/english-date-corpus.cjs");

// v3.2.0: index.html as the tests compare it, without the data-i18n* attributes that carry the
// English keys, so the Chinese text and markup read as before. Tests read index.html only this way.
const DATA_I18N_ATTRIBUTE = /\s+data-i18n[\w-]*="[^"]*"/g;
const stripI18nAttributes = (html) => html.replace(DATA_I18N_ATTRIBUTE, "");

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
  test(`${method} custom analysis preserves statistics and exports median in the report`, () => {
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
  });
}

test("public current version markers agree on v3.2.2", () => {
  const root = path.join(__dirname, "..");
  assert.equal(JSON.parse(fs.readFileSync(path.join(root, "package.json"), "utf8")).version, "3.2.2");
  assert.match(fs.readFileSync(path.join(root, "README.md"), "utf8"), /\*\*Current release:\*\* v3\.2\.2/);
  assert.match(fs.readFileSync(path.join(root, "README.zh-CN.md"), "utf8"), /\*\*当前版本：\*\* v3\.2\.2/);
  const { evaluate } = createScriptContext();
  assert.equal(evaluate("buildReportData().version"), "v3.2.2");
  assert.equal(evaluate("normalizeReportSnapshot({}).version"), "v3.2.2");

  // index.html cache-busting parameters must follow package.json.
  const packageVersion = JSON.parse(fs.readFileSync(path.join(root, "package.json"), "utf8")).version;
  const html = stripI18nAttributes(fs.readFileSync(path.join(root, "index.html"), "utf8"));
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
      ...NUMBER_UNITS.map((unit) => ({ unit: unit.toLowerCase(), unitConflict: false, units: {}, allowMagnitude: true, fullwidth: true })),
      // v3.3.0: English units and currency codes, keyed by the name matchNumberWithUnit() returns.
      ...[...new Set([...Object.values(ENGLISH_NUMBER_UNITS), ...Object.values(ENGLISH_CURRENCY_PREFIXES)])].map((unit) => ({ unit, unitConflict: false, units: {}, allowMagnitude: true, fullwidth: true }))
    ];
    // v3.3.0: 28 writings (None moved to the maybe tier with Unknown, TBD, ?, n.a. and their cases).
    return {
      formCount: forms.length,
      converted: forms.filter((form) => strategies.some((strategy) => toNumber(form, strategy) !== null)),
      matched: forms.filter((form) => matchNumberWithUnit(form.trim()) || matchNumberWithUnit(toHalfwidthNumberText(form.trim()))),
      strategy: buildNumberStrategy(forms)
    };
  })())`));
  assert.equal(result.formCount, 84);
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

  for (const name of ["序号", "行号", "行序号", "Index", "IDX", "No.", "Row No.", "Row No", "#", " 序号 ", "学生ID", "客户 ID", "视频id",
    "Order Number", "Order No", "Order No.", "Invoice number", "ticket_no", "Phone Number", "Number", "No"]) {
    assert.equal(isIdName(name), true, name);
  }
  for (const name of ["序号说明", "index_value", "indexes", "noise", "No. of employees", "指数", "身份证明", "ID卡", "IDEA", "Android", "学号数量",
    "Number of employees", "number_of_items", "no_previous", "No Show", "Number Sold"]) {
    assert.equal(isIdName(name), false, name);
  }
  // v3.3.0 (the other half of deferred finding 6 of v2.1.3): number and no count only as the last
  // word, so "Number of employees" is a count now. id, uuid, guid, code and key are unchanged.
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
  // v3.1.0: the repeated rows past the first of each group, by index; the first one stays.
  assert.deepEqual(JSON.parse(evaluate('JSON.stringify(findRepeatedRowIndexes(state.rows, state.fields))')), [1]);
  assert.equal(evaluate('findRepeatedRowIndexes(state.rows, state.fields).includes(0)'), false);
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
    common: ["N/A", "n/a", "#N/A", "NULL", "null", "Null", "NaN", "nan", "NAN", "—", "——", "--"],
    maybe: ["无", "暂无", "-", "/", "NA", "None", "none", "NONE", "Unknown", "unknown", "UNKNOWN", "TBD", "tbd", "?", "n.a.", "N.A."]
  });
  assert.equal(evaluate('placeholderTier("N/A")'), "common");
  assert.equal(evaluate('placeholderTier("——")'), "common");
  assert.equal(evaluate('placeholderTier("无")'), "maybe");
  assert.equal(evaluate('placeholderTier("NA")'), "maybe");
  assert.equal(evaluate('placeholderTier("缺考")'), "");
  // v3.3.0: None is an answer in some columns ("damage: None"), so it moved to the maybe tier.
  assert.equal(evaluate('placeholderTier("None")'), "maybe");
  assert.equal(evaluate('placeholderTier("na")'), "");
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

// v2.10.0: one-cell rules of the standardized export (parsing.js). Invisible characters are
// written as escapes only.
function standardizeCases(expression, cases) {
  const { context, evaluate } = createScriptContext();
  context.__cases = cases;
  return JSON.parse(evaluate(`JSON.stringify(__cases.map((args) => ${expression}))`));
}

const STANDARDIZE_NONE = { order: "none", allowYearOnly: false, granularity: "day" };

test("v2.10.0 standardizeDateCell: the boundary table", () => {
  const dmy = { ...STANDARDIZE_NONE, order: "dmy" };
  const mdy = { ...STANDARDIZE_NONE, order: "mdy" };
  const cases = [
    ["2024-05-06", STANDARDIZE_NONE, [], "2024-05-06", ""],
    ["2024-5-6", STANDARDIZE_NONE, [], "2024-05-06", "date"],
    ["2024/05/06", STANDARDIZE_NONE, [], "2024-05-06", "date"],
    ["2024.5.6", STANDARDIZE_NONE, [], "2024-05-06", "date"],
    ["2024年5月6日", STANDARDIZE_NONE, [], "2024-05-06", "date"],
    ["2024年5月6", STANDARDIZE_NONE, [], "2024-05-06", "date"],
    ["2024 年 5 月 6 日", STANDARDIZE_NONE, [], "2024-05-06", "date"],
    ["2024-05-06 10:30", STANDARDIZE_NONE, [], "2024-05-06 10:30", "keptTime"],
    ["2024-05-06T10:30:00Z", STANDARDIZE_NONE, [], "2024-05-06T10:30:00Z", "keptTime"],
    ["2024年5月6日 10:30", STANDARDIZE_NONE, [], "2024年5月6日 10:30", "keptTime"],
    ["06/05/2024", dmy, [], "2024-05-06", "date"],
    ["06/05/2024", mdy, [], "2024-06-05", "date"],
    ["06/05/2024", STANDARDIZE_NONE, [], "06/05/2024", "keptDateOrder"],
    ["06/05/2024", { ...STANDARDIZE_NONE, order: "conflict" }, [], "06/05/2024", "keptDateOrder"],
    ["06/06/2024", STANDARDIZE_NONE, [], "2024-06-06", "date"],
    ["13/05/2024", STANDARDIZE_NONE, [], "2024-05-13", "date"],
    ["Mon, 6 May 2024", STANDARDIZE_NONE, [], "2024-05-06", "date"],
    ["Tue, 6 May 2024", STANDARDIZE_NONE, [], "2024-05-06", "date"],
    ["May 6, 2024 10:30 AM", STANDARDIZE_NONE, [], "May 6, 2024 10:30 AM", "keptTime"],
    ["May 6, 2024 GMT", STANDARDIZE_NONE, [], "May 6, 2024 GMT", "keptTime"],
    // D3: already YYYY-MM-DD, only the edge whitespace goes.
    [" 2024-05-06 ", STANDARDIZE_NONE, [], "2024-05-06", "trim"],
    ["\u30002024/5/6\u00A0", STANDARDIZE_NONE, [], "2024-05-06", "date"],
    // D4: every kept cell stays exactly as written, edge whitespace included.
    [" 2024-05-06 10:30 ", STANDARDIZE_NONE, [], " 2024-05-06 10:30 ", "keptTime"],
    [" 06/05/2024 ", STANDARDIZE_NONE, [], " 06/05/2024 ", "keptDateOrder"],
    [" 2024-02-30 ", STANDARDIZE_NONE, [], " 2024-02-30 ", "keptDateUnreadable"],
    ["31/02/2024", STANDARDIZE_NONE, [], "31/02/2024", "keptDateUnreadable"],
    ["待定", STANDARDIZE_NONE, [], "待定", "keptDateUnreadable"],
    ["N/A", STANDARDIZE_NONE, [], "N/A", "keptPendingPlaceholder"],
    [" N/A ", STANDARDIZE_NONE, [], " N/A ", "keptPendingPlaceholder"],
    [" N/A ", STANDARDIZE_NONE, ["N/A"], "", "placeholder"],
    // D2: a blank cell becomes empty.
    ["   ", STANDARDIZE_NONE, [], "", "trim"],
    ["", STANDARDIZE_NONE, [], "", ""]
  ];
  const results = standardizeCases("standardizeDateCell(args[0], args[1], args[2])", cases);
  results.forEach((result, index) => {
    const [value, , , exported, action] = cases[index];
    assert.deepEqual(result, { value: exported, action }, JSON.stringify(value));
  });
});

test("v2.10.0 standardizeNumberCell: the boundary table, as text and never through a float", () => {
  const plain = { unit: "", unitConflict: false, units: {}, allowMagnitude: false, fullwidth: false };
  const yuan = { ...plain, unit: "元" };
  const wan = { ...plain, allowMagnitude: true };
  const full = { ...plain, fullwidth: true };
  const cases = [
    ["2200", plain, "2200", ""], ["007", plain, "007", ""], ["1.50", plain, "1.50", ""], ["-0", plain, "-0", ""],
    ["12345678901234567890", plain, "12345678901234567890", ""],
    // D3: a plain number with edge whitespace is only trimmed.
    [" 2200 ", plain, "2200", "trim"], ["\u3000007\u00A0", plain, "007", "trim"],
    ["2200元", yuan, "2200", "number"], ["2,200元", yuan, "2200", "number"], ["-2200元", yuan, "-2200", "number"],
    ["+2200", plain, "2200", "number"], [".5", plain, "0.5", "number"], ["-.5", plain, "-0.5", "number"], ["5.", plain, "5", "number"],
    ["1,234.50", plain, "1234.50", "number"], ["1 234", plain, "1234", "number"], ["1\u00A0234", plain, "1234", "number"],
    ["１２３４", full, "1234", "number"], ["１，２３４．５", full, "1234.5", "number"], ["－１２", full, "-12", "number"],
    ["0.07万", wan, "700", "number"], ["1.2万", wan, "12000", "number"], ["3亿", wan, "300000000", "number"],
    ["12万元", { ...yuan, allowMagnitude: true }, "120000", "number"], ["-0.5万", wan, "-5000", "number"], ["1.50万", wan, "15000", "number"],
    ["1.234560万", wan, "12345.60", "number"], ["123456789012万", wan, "1234567890120000", "number"],
    ["￥2200", plain, "2200", "number"], ["¥2,200", plain, "2200", "number"], ["$12", plain, "12", "number"], ["-$12", plain, "-12", "number"], ["12$", plain, "12", "number"],
    ["(200)", plain, "-200", "number"], ["(1,200.5)", plain, "-1200.5", "number"], ["(-200)", plain, "-200", "number"],
    // D7: a rewrite that reads as zero drops its "-".
    ["(0)", plain, "0", "number"], ["(0.00)", plain, "0.00", "number"], ["-0,000", plain, "0000", "number"], ["+0", plain, "0", "number"], ["-0万", wan, "0", "number"],
    ["50%", plain, "50%", "keptPercent"], ["(5%)", plain, "(5%)", "keptPercent"], ["１２％", full, "１２％", "keptPercent"],
    ["1e3", plain, "1e3", "keptScientific"], ["1E-3", plain, "1E-3", "keptScientific"],
    ["缺考", plain, "缺考", "keptNumberUnreadable"], ["2200元", plain, "2200元", "keptNumberUnreadable"], ["1.2万", plain, "1.2万", "keptNumberUnreadable"],
    ["N/A", plain, "N/A", "keptPendingPlaceholder"], ["—", plain, "—", "keptPendingPlaceholder"],
    // D4: kept cells keep their edge whitespace.
    [" 缺考 ", plain, " 缺考 ", "keptNumberUnreadable"], [" 50% ", plain, " 50% ", "keptPercent"], [" 1e3 ", plain, " 1e3 ", "keptScientific"], [" N/A ", plain, " N/A ", "keptPendingPlaceholder"],
    // D2: a blank cell becomes empty.
    ["\u3000", plain, "", "trim"], ["", plain, "", ""]
  ];
  const results = standardizeCases("(() => { const r = standardizeNumberCell(args[0], args[1], [], ''); const a = toPlainNumber(r.value); const b = toNumber(args[0], args[1]); return { ...r, same: Object.is(a, b) || (a === 0 && b === 0) }; })()", cases);
  results.forEach((result, index) => {
    const [value, , exported, action] = cases[index];
    assert.equal(result.value, exported, JSON.stringify(value));
    assert.equal(result.action, action, JSON.stringify(value));
    if (action === "number" || action === "trim") assert.equal(result.same, true, JSON.stringify(value));
  });
  const mixed = standardizeCases("standardizeNumberCell(args[0], {}, [], 'mixed')", [["$12"], ["￥12"], [" $12 "], ["12"], ["1,200"]]);
  assert.deepEqual(mixed, [
    { value: "$12", action: "keptCurrencyMix" }, { value: "￥12", action: "keptCurrencyMix" }, { value: " $12 ", action: "keptCurrencyMix" },
    { value: "12", action: "" }, { value: "1200", action: "number" }
  ]);
  assert.deepEqual(standardizeCases("standardizeNumberCell(args[0], {}, ['N/A'], '')", [["N/A"], [" N/A "]]), [{ value: "", action: "placeholder" }, { value: "", action: "placeholder" }]);
});

test("v2.10.0 standardizeNumberCell: the safety net keeps a rewrite that would read as another number", () => {
  const { evaluate } = createScriptContext();
  assert.equal(evaluate('JSON.stringify(standardizeNumberCell("2,200", {}, [], ""))'), JSON.stringify({ value: "2200", action: "number" }));
  // The parser now reads 2,200 as 2201: the plain text 2200 would not read back the same.
  evaluate('var __realToNumber = toNumber; toNumber = (value, strategy) => (value === "2,200" ? 2201 : __realToNumber(value, strategy))');
  assert.equal(evaluate('JSON.stringify(standardizeNumberCell("2,200", {}, [], ""))'), JSON.stringify({ value: "2,200", action: "keptUnsafe" }));
  assert.equal(evaluate('JSON.stringify(standardizeNumberCell("1,500", {}, [], ""))'), JSON.stringify({ value: "1500", action: "number" }));
  assert.equal(evaluate('formatStandardizeAction("keptUnsafe")'), "无法安全改写，保持原值");
});

test("v2.10.0 standardizeTextCell: edge whitespace goes, invisible characters and look-alikes stay", () => {
  const cases = [
    [" 华东 ", [], "华东", "trim"], ["\u3000华东", [], "华东", "trim"], ["华东\u00A0", [], "华东", "trim"], ["\uFEFF华东", [], "华东", "trim"],
    ["   ", [], "", "trim"], ["\u3000", [], "", "trim"], ["\u200B华东", [], "\u200B华东", ""], ["\u200B", [], "\u200B", ""],
    ["ＨＤ", [], "ＨＤ", ""], ["N/A", [], "N/A", ""], [" N/A ", [], "N/A", "trim"], [" N/A ", ["N/A"], "", "placeholder"], ["", [], "", ""]
  ];
  const results = standardizeCases("standardizeTextCell(args[0], args[1])", cases);
  results.forEach((result, index) => {
    const [value, , exported, action] = cases[index];
    assert.deepEqual(result, { value: exported, action }, JSON.stringify(value));
  });
});

test("v2.10.0 the action texts", () => {
  const { context, evaluate } = createScriptContext();
  context.__actions = ["date", "number", "placeholder", "trim", "keptTime", "keptPercent", "keptScientific", "keptCurrencyMix", "keptPendingPlaceholder", "keptDateUnreadable", "keptDateOrder", "keptNumberUnreadable", "keptUnsafe"];
  assert.deepEqual(JSON.parse(evaluate("JSON.stringify(__actions.map(formatStandardizeAction))")), [
    "日期改写为 YYYY-MM-DD", "数字改写为纯数字", "已按空值处理的写法，清空", "去首尾空格", "带时间，保持原值", "百分比，不改写", "科学计数法，不改写",
    "货币符号不一致，保持原值", "像是空值的写法，未确认", "读不成日期，保持原值", "日月顺序无法确定，保持原值", "读不成数字，保持原值", "无法安全改写，保持原值"
  ]);
  assert.deepEqual(JSON.parse(evaluate("JSON.stringify(STANDARDIZE_KEPT_ACTIONS)")), context.__actions.filter((action) => action.startsWith("kept")));
});

test("v2.10.0 guard: a rewritten date reads as the same day, an unread one stays, a time is never rewritten", () => {
  const values = [];
  for (const body of ["2024-5-6", "2024/05/06", "2024.12.31", "2024年5月6日", "2024 年 5 月 6", "06/05/2024", "13/05/2024", "06/06/2024", "2024-02-30", "May 6, 2024", "6 May 2024", "Mon, 6 May 2024"]) {
    for (const suffix of ["", " 10:30", "T10:30:00Z", " 10:30 PM", " GMT", " 24:00"]) values.push(body + suffix, ` ${body + suffix} `);
  }
  const { context, evaluate } = createScriptContext();
  context.__values = values;
  const bad = JSON.parse(evaluate(`JSON.stringify(["none", "dmy", "mdy", "conflict"].flatMap((order) => __values.flatMap((value) => {
    const strategy = { order, allowYearOnly: false, granularity: "day" };
    const before = toDate(value, strategy);
    const result = standardizeDateCell(value, strategy, []);
    if (!before) return result.value === value ? [] : [["unread changed", value, order]];
    const clock = /\\d:\\d\\d|\\dT\\d|GMT$/.test(value.trim());
    if (clock !== (result.action === "keptTime")) return [["time", value, order, result.action]];
    if (result.action === "keptTime") return result.value === value ? [] : [["time changed", value]];
    const after = toDate(result.value, strategy);
    return after && after.getTime() === before.getTime() && /^\\d{4}-\\d{2}-\\d{2}$/.test(result.value) ? [] : [["day", value, order, result.value]];
  })))`));
  assert.deepEqual(bad, []);
});

test("v2.10.0 guard: every rewritten number reads back as the field parser's number, zero without a sign", () => {
  const values = [];
  for (const sign of ["", "-", "+"]) for (const currency of ["", "$", "￥"]) for (const tail of ["", "元", "万", "%"]) {
    for (const number of ["0", "0.00", "0,000", "007", "0.07", ".5", "5.", "1,234.50", "1 234", "12345678901234567890", "1e3"]) {
      values.push(`${sign}${currency}${number}${tail}`, `(${currency}${number}${tail})`, ` ${sign}${number}${tail} `);
    }
  }
  const { context, evaluate } = createScriptContext();
  context.__values = values;
  const bad = JSON.parse(evaluate(`JSON.stringify([{}, { unit: "元" }, { allowMagnitude: true }, { fullwidth: true }].flatMap((strategy) => __values.flatMap((value) => {
    const result = standardizeNumberCell(value, strategy, [], "");
    if (result.action === "keptUnsafe") return [["safety net", value]];
    if (result.action === "number") {
      const after = toPlainNumber(result.value);
      const before = toNumber(value, strategy);
      const same = Object.is(after, before) || (after === 0 && before === 0);
      return same && /^-?\\d+(?:\\.\\d+)?$/.test(result.value) && !(after === 0 && result.value.startsWith("-")) ? [] : [[value, result.value]];
    }
    if (result.action === "trim" || result.action === "") return [];
    return result.value === value ? [] : [["kept changed", value, result.value]];
  })))`));
  assert.deepEqual(bad, []);
});

// v3.3.0 ①: the three date matchers use patterns built once at load, not a new RegExp per call
// (toDate() runs them for every date-like cell). The counter is checked on a RegExp built here.
test("v3.3.0 date matchers build no RegExp per call and match as before", () => {
  const { evaluate } = createScriptContext();
  const result = JSON.parse(evaluate(`JSON.stringify((() => {
    const Original = RegExp;
    let built = 0;
    RegExp = new Proxy(Original, { construct(target, args) { built += 1; return Reflect.construct(target, args); } });
    try {
      new RegExp("a");
      const sanity = built;
      built = 0;
      const texts = ["2024-05-06", "2024/5/6 10:30", "06/05/2024", "13.05.2024", "2024年5月6日", "2024 年 5 月 6 10:30", "2024-13-45", "abc"];
      const matches = texts.map((text) => [matchYearFirstDate(text), matchDayOrMonthFirstDate(text), matchChineseDate(text)].map((match) => match && [...match]));
      for (let i = 0; i < 10000; i += 1) {
        matchYearFirstDate("2024-05-06");
        matchDayOrMonthFirstDate("06/05/2024");
        matchChineseDate("2024年5月6日");
      }
      return { sanity, built, matches };
    } finally {
      RegExp = Original;
    }
  })())`));
  assert.equal(result.sanity, 1);
  assert.equal(result.built, 0);
  assert.deepEqual(result.matches, [
    [["2024-05-06", "2024", "-", "05", "06", ""], null, null],
    [["2024/5/6 10:30", "2024", "/", "5", "6", " 10:30"], null, null],
    [null, ["06/05/2024", "06", "/", "05", "2024", ""], null],
    [null, ["13.05.2024", "13", ".", "05", "2024", ""], null],
    [null, null, ["2024年5月6日", "2024", "5", "6", ""]],
    [null, null, ["2024 年 5 月 6 10:30", "2024", "5", "6", " 10:30"]],
    [["2024-13-45", "2024", "-", "13", "45", ""], null, null],
    [null, null, null]
  ]);
});

// ---- v3.3.0 ③: None and new spellings that may mean blank ----
test("v3.3.0 placeholders: the new maybe-tier writings, full-width forms included; look-alikes that are not on the list", () => {
  const { context, evaluate } = createScriptContext();
  const key = (value) => {
    context.__placeholderValue = value;
    return evaluate("placeholderKey(__placeholderValue)");
  };
  const tier = (value) => evaluate(`placeholderTier(placeholderKey(${JSON.stringify(value)}))`);
  for (const value of ["None", "none", "NONE", "Unknown", "unknown", "UNKNOWN", "TBD", "tbd", "?", "n.a.", "N.A."]) {
    assert.equal(key(value), value, value);
    assert.equal(key(` ${value} `), value, value);
    assert.equal(tier(value), "maybe", value);
  }
  // Full-width forms through NFKC: ？ (U+FF1F), ＵＮＫＮＯＷＮ, ｎ．ａ．
  const fullWidth = (text) => text.replace(/[!-~]/g, (char) => String.fromCharCode(char.charCodeAt(0) + 0xfee0));
  assert.equal(key(String.fromCharCode(0xff1f)), "?");
  assert.equal(key(fullWidth("UNKNOWN")), "UNKNOWN");
  assert.equal(key(fullWidth("n.a.")), "n.a.");
  // Whole cells only: near misses and other "no data" words are not on the list.
  for (const value of ["na", "NONE.", "Nones", "(blank)", "Missing", "nil", "??", "TBD?", "n/a.", "N.A", "Unknown value"]) {
    assert.equal(key(value), "", value);
  }
  // "常见写法全部按空值处理" covers the common tier only, which no longer holds any None.
  assert.equal(evaluate("PLACEHOLDER_VALUES.common.some((key) => key.toLowerCase() === 'none')"), false);
});

test("v3.3.0 placeholders: the cell length limit follows the longest writing (14), and pure ASCII cells skip NFKC", () => {
  const { context, evaluate } = createScriptContext();
  assert.equal(evaluate("PLACEHOLDER_MAX_CELL_LENGTH"), 14);
  // 14 units can still match (UNKNOWN in full width with spaces would be longer; the limit is on the trimmed text).
  context.__cells = ["UNKNOWN", `${" ".repeat(20)}UNKNOWN`, "x".repeat(14), "x".repeat(15)];
  assert.deepEqual(JSON.parse(evaluate("JSON.stringify(__cells.map(placeholderKey))")), ["UNKNOWN", "UNKNOWN", "", ""]);
  // NFKC leaves every ASCII code unit and every ASCII writing on the list as it is, so skipping it changes nothing.
  const identity = JSON.parse(evaluate(`JSON.stringify((() => {
    const units = Array.from({ length: 128 }, (_, code) => String.fromCharCode(code));
    const writings = [...PLACEHOLDER_VALUES.common, ...PLACEHOLDER_VALUES.maybe].filter((text) => /^[\\x00-\\x7F]*$/.test(text));
    return { units: units.filter((text) => text.normalize("NFKC") !== text).length, writings: writings.filter((text) => text.normalize("NFKC") !== text).length, writingCount: writings.length };
  })())`));
  assert.deepEqual(identity, { units: 0, writings: 0, writingCount: 24 });
  const calls = JSON.parse(evaluate(`JSON.stringify((() => {
    const original = String.prototype.normalize;
    let count = 0;
    String.prototype.normalize = function (...args) { count += 1; return original.apply(this, args); };
    try {
      const ascii = ["Item-0001abc", "UNKNOWN", "n.a.", "?", "ORDER-LONG-1"].map(placeholderKey);
      const asciiCalls = count;
      count = 0;
      const wide = [String.fromCharCode(0xff1f), "暂无"].map(placeholderKey);
      return { ascii, asciiCalls, wide, wideCalls: count };
    } finally {
      String.prototype.normalize = original;
    }
  })())`));
  assert.deepEqual(calls, { ascii: ["", "UNKNOWN", "n.a.", "?", ""], asciiCalls: 0, wide: ["?", "暂无"], wideCalls: 2 });
});

// The maybe-tier note of 3594fba, written out here: a column whose maybe writings are all among
// the five of v2.4.0 keeps it byte for byte in both languages.
function oldPlaceholderNotes(keys, language) {
  const maybe = keys.filter((key) => ["无", "暂无", "-", "/", "NA"].includes(key));
  if (!maybe.length) return "";
  const dash = keys.includes("-");
  if (language === "zh") return `「无」「暂无」「-」「/」「NA」也可能是有效值（例如“有/无”选项、北美地区），请确认这一列里它确实表示没有数据。${dash ? "在会计格式里，「-」通常表示 0。" : ""}`;
  const examples = [maybe.some((key) => key !== "NA") ? "a yes/no answer" : "", maybe.includes("NA") ? "NA for North America" : ""].filter(Boolean).join(", or ");
  const sentence = maybe.length === 1
    ? `“${maybe[0]}” may also be a real value (for example ${examples}); make sure that in this column it means there is no data.`
    : `${maybe.map((key) => `“${key}”`).join(", ")} may also be real values (for example ${examples}); make sure that in this column they mean there is no data.`;
  return `${sentence}${dash ? " In accounting formats, “-” usually means 0." : ""}`;
}

test("v3.3.0 placeholders: a column with only the five old maybe writings keeps the old note in both languages (31 subsets)", () => {
  const old = ["无", "暂无", "-", "/", "NA"];
  for (const language of ["zh", "en"]) {
    const { evaluate } = createScriptContext({ language });
    let checked = 0;
    for (let mask = 1; mask < 32; mask += 1) {
      const subset = old.filter((_, index) => mask & (1 << index));
      for (const keys of [subset, ["N/A", ...subset], [...subset, "NULL"]]) {
        assert.equal(evaluate(`formatPlaceholderNotes(${JSON.stringify(keys)})`), oldPlaceholderNotes(keys, language), `${language} ${keys.join(" ")}`);
        checked += 1;
      }
    }
    assert.equal(checked, 93);
    assert.equal(evaluate('formatPlaceholderNotes(["N/A", "NULL"])'), "");
  }
});

test("v3.3.0 placeholders: a column with a new maybe writing names only its own writings, each kind with its example", () => {
  const zh = createScriptContext({ language: "zh" });
  const en = createScriptContext({ language: "en" });
  const note = (scope, keys) => scope.evaluate(`formatPlaceholderNotes(${JSON.stringify(keys)})`);
  assert.equal(note(zh, ["None"]), "「None」也可能是有效值（例如表示“没有”的回答，比如损坏程度：None），请确认这一列里它确实表示没有数据。");
  assert.equal(note(zh, ["None", "none"]), "「None」「none」也可能是有效值（例如表示“没有”的回答，比如损坏程度：None），请确认这一列里它们确实表示没有数据。");
  assert.equal(note(zh, ["Unknown"]), "「Unknown」也可能是有效值（例如单独的一个选项），请确认这一列里它确实表示没有数据。");
  assert.equal(note(zh, ["TBD"]), "「TBD」也可能是有效值（例如表示“待定”的状态），请确认这一列里它确实表示没有数据。");
  assert.equal(note(zh, ["?"]), "「?」也可能是有效值（例如问卷里的“不确定”），请确认这一列里它确实表示没有数据。");
  assert.equal(note(zh, ["n.a."]), "「n.a.」也可能是有效值（例如“不适用”），请确认这一列里它确实表示没有数据。");
  assert.equal(note(zh, ["N/A", "UNKNOWN", "-"]), "「UNKNOWN」「-」也可能是有效值（例如“有/无”选项；单独的一个选项），请确认这一列里它们确实表示没有数据。在会计格式里，「-」通常表示 0。");
  assert.equal(note(zh, ["None", "Unknown", "TBD", "?", "n.a.", "NA", "无"]), "「None」「Unknown」「TBD」「?」「n.a.」「NA」「无」也可能是有效值（例如“有/无”选项；北美地区；表示“没有”的回答，比如损坏程度：None；单独的一个选项；表示“待定”的状态；问卷里的“不确定”；“不适用”），请确认这一列里它们确实表示没有数据。");

  assert.equal(note(en, ["None"]), "“None” may also be a real value (for example an answer meaning “none”, such as damage: None); make sure that in this column it means there is no data.");
  assert.equal(note(en, ["None", "none"]), "“None”, “none” may also be real values (for example an answer meaning “none”, such as damage: None); make sure that in this column they mean there is no data.");
  assert.equal(note(en, ["Unknown"]), "“Unknown” may also be a real value (for example an option of its own); make sure that in this column it means there is no data.");
  assert.equal(note(en, ["TBD"]), "“TBD” may also be a real value (for example a status such as “to be decided”); make sure that in this column it means there is no data.");
  assert.equal(note(en, ["?"]), "“?” may also be a real value (for example “not sure” in a survey); make sure that in this column it means there is no data.");
  assert.equal(note(en, ["N.A."]), "“N.A.” may also be a real value (for example “not applicable”); make sure that in this column it means there is no data.");
  assert.equal(note(en, ["N/A", "UNKNOWN", "-"]), "“UNKNOWN”, “-” may also be real values (for example a yes/no answer; or an option of its own); make sure that in this column they mean there is no data. In accounting formats, “-” usually means 0.");
  assert.equal(note(en, ["None", "Unknown", "TBD", "?", "n.a.", "NA", "无"]), "“None”, “Unknown”, “TBD”, “?”, “n.a.”, “NA”, “无” may also be real values (for example a yes/no answer; or NA for North America; or an answer meaning “none”, such as damage: None; or an option of its own; or a status such as “to be decided”; or “not sure” in a survey; or “not applicable”); make sure that in this column they mean there is no data.");
});

// The report passes the writings of every column at once, so a writing can come twice; each is
// named once. The Chinese old sentence is fixed text and stays byte for byte; the English old
// sentence loses the repeat and is singular when one writing is left (an intended fix).
test("v3.3.0 placeholders: a writing that comes twice is named once, in both languages and both sentences", () => {
  const zh = createScriptContext({ language: "zh" });
  const en = createScriptContext({ language: "en" });
  const note = (scope, keys) => scope.evaluate(`formatPlaceholderNotes(${JSON.stringify(keys)})`);
  assert.equal(note(zh, ["-", "Unknown", "-"]), "「-」「Unknown」也可能是有效值（例如“有/无”选项；单独的一个选项），请确认这一列里它们确实表示没有数据。在会计格式里，「-」通常表示 0。");
  assert.equal(note(zh, ["NA", "-", "NA", "None", "none"]), "「NA」「-」「None」「none」也可能是有效值（例如“有/无”选项；北美地区；表示“没有”的回答，比如损坏程度：None），请确认这一列里它们确实表示没有数据。在会计格式里，「-」通常表示 0。");
  assert.equal(note(zh, ["Unknown", "Unknown"]), "「Unknown」也可能是有效值（例如单独的一个选项），请确认这一列里它确实表示没有数据。");
  assert.equal(note(zh, ["-", "-", "NA", "NA"]), "「无」「暂无」「-」「/」「NA」也可能是有效值（例如“有/无”选项、北美地区），请确认这一列里它确实表示没有数据。在会计格式里，「-」通常表示 0。");
  assert.equal(note(en, ["-", "-"]), "“-” may also be a real value (for example a yes/no answer); make sure that in this column it means there is no data. In accounting formats, “-” usually means 0.");
  assert.equal(note(en, ["NA", "-", "NA"]), "“NA”, “-” may also be real values (for example a yes/no answer, or NA for North America); make sure that in this column they mean there is no data. In accounting formats, “-” usually means 0.");
  assert.equal(note(en, ["-", "Unknown", "-"]), "“-”, “Unknown” may also be real values (for example a yes/no answer; or an option of its own); make sure that in this column they mean there is no data. In accounting formats, “-” usually means 0.");
});

// ---- v3.3.0 ④: number and no count as IDs only as the last word ----
test("v3.3.0 ID names: a column of counts 0-9 is numeric unless number / no is its last word", () => {
  const { context, evaluate } = createScriptContext();
  const counts = [12, 7, 30, 15, 22, 9, 41, 3, 18, 26, 11, 5, 33, 14, 8, 20].map((n) => String(n % 10));
  const typeOf = (name) => {
    context.__rows = counts.map((value) => ({ [name]: value }));
    return evaluate(`buildColumnProfile(${JSON.stringify(name)}, __rows).typeKey`);
  };
  for (const name of ["Number of Records", "No. of items", "no_previous", "num_of_items", "number_of_items", "No Show", "Number Sold"]) {
    assert.equal(typeOf(name), "numeric", name);
  }
  for (const name of ["number", "Order Number", "Order No", "Order No.", "Customer No", "No", "No.", "Number", "#", "Invoice number", "ticket_no", "Phone Number", "Lot no."]) {
    assert.equal(typeOf(name), "id", name);
  }
  // Known (ARCHITECTURE): a count whose last word is number still reads as an ID.
  assert.equal(typeOf("Total Number"), "id");
});

test("v3.3.0 ID names: a survey header with the word yes is a question, so its last word no is not an ID", () => {
  const { context, evaluate } = createScriptContext();
  const answers = Array.from({ length: 16 }, (_, index) => (index % 3 ? "Yes" : "No"));
  const view = (name) => {
    context.__rows = answers.map((value) => ({ [name]: value }));
    return [evaluate(`isIdFieldName(${JSON.stringify(name)})`), evaluate(`buildColumnProfile(${JSON.stringify(name)}, __rows).typeKey`)];
  };
  for (const name of ["Returning customer: Yes or No", "Smoker Yes No", "Smoker_yes_no", "Smoker yes/no", "YES/NO", "Yes-No", "(Yes) or No", "Returning customer? Yes, No"]) {
    assert.deepEqual(view(name), [false, "category"], name);
  }
  // Not a separate word: "Yesterday", "Eyes". And ordinary IDs stay IDs.
  for (const name of ["Yesterday No", "Eyes No", "Room No", "Order No.", "Order Number", "No"]) {
    assert.deepEqual(view(name), [true, "id"], name);
  }
  // Only the number / no rule gives way: id, uuid, guid, code and key are unchanged.
  for (const name of ["Yes code", "yes_id", "Yes Key"]) assert.equal(view(name)[0], true, name);
});

// ---- v3.3.0 ⑤: English units and currency codes ----
// [value, number read with the column's own strategy, the strategy's unit]
const readWithOwnStrategy = (evaluate, value) => JSON.parse(evaluate(`JSON.stringify((() => {
  const value = ${JSON.stringify(value)};
  const strategy = buildNumberStrategy([value, value]);
  return [toNumber(value, strategy), strategy.unit];
})())`));

test("v3.3.0 units: every English spelling reads as its unit's name, in any case; a dot only after an English unit", () => {
  const { evaluate } = createScriptContext();
  for (const [value, number, unit] of [
    ["200 USD", 200, "USD"], ["200 usd", 200, "USD"], ["5,400 USD", 5400, "USD"], ["-200 USD", -200, "USD"], ["300 RMB", 300, "CNY"], ["300 CNY", 300, "CNY"],
    ["9 EUR", 9, "EUR"], ["9 GBP", 9, "GBP"], ["9 JPY", 9, "JPY"], ["9 CAD", 9, "CAD"], ["9 AUD", 9, "AUD"], ["9 HKD", 9, "HKD"], ["9 CHF", 9, "CHF"], ["9 INR", 9, "INR"], ["9 SGD", 9, "SGD"],
    ["12 pcs", 12, "pcs"], ["1 pc", 1, "pcs"], ["3 pieces", 3, "pcs"], ["1 piece", 1, "pcs"], ["12 units", 12, "units"], ["1 unit", 1, "units"],
    ["5 lbs", 5, "lb"], ["5 lb", 5, "lb"], ["5 lbs.", 5, "lb"], ["5 oz", 5, "oz"],
    ["8 hours", 8, "hours"], ["1 hour", 1, "hours"], ["8 hrs", 8, "hours"], ["8 hr", 8, "hours"], ["30 minutes", 30, "minutes"], ["1 minute", 1, "minutes"], ["30 mins", 30, "minutes"], ["30 min", 30, "minutes"],
    ["5 days", 5, "days"], ["1 day", 1, "days"], ["3 weeks", 3, "weeks"], ["1 week", 1, "weeks"], ["2 months", 2, "months"], ["1 month", 1, "months"],
    ["2 years", 2, "years"], ["1 year", 1, "years"], ["2 yrs", 2, "years"], ["1 yr", 1, "years"], ["10 miles", 10, "miles"], ["1 mile", 1, "miles"], ["6 ft", 6, "ft"], ["6 feet", 6, "ft"], ["6 FT", 6, "ft"]
  ]) assert.deepEqual(readWithOwnStrategy(evaluate, value), [number, unit], value);
  // The dot is English only; Chinese units and kg, cm, mm, km take none.
  for (const value of ["5 kg.", "200元.", "3 cm.", "5 lbs.."]) assert.deepEqual(readWithOwnStrategy(evaluate, value), [null, ""], value);
});

test("v3.3.0 units: currency codes in front; what is not read", () => {
  const { evaluate } = createScriptContext();
  for (const [value, number] of [["USD 200", 200], ["USD200", 200], ["US$200", 200], ["us$ 1,050.50", 1050.5], ["Usd 200", 200], ["USD -200", -200], ["RMB 300", 300]]) {
    assert.deepEqual(readWithOwnStrategy(evaluate, value), [number, value.toLowerCase().startsWith("rmb") ? "CNY" : "USD"], value);
  }
  // A code on both sides, single letters, in and mi, K / M, degrees, CR, words around the number.
  for (const value of ["USD 200 USD", "200 USD EUR", "5 g", "5 m", "5 h", "5 s", "5 d", "5 k", "6 in", "3 mi", "1.2K", "3.4M", "$1.2M", "72°F", "1,234.50 CR", "Day 3", "2 for $5", "(200 USD)", "USD -", "127 Hours long"]) {
    assert.deepEqual(readWithOwnStrategy(evaluate, value), [null, ""], value);
  }
  // Chinese units keep their own lower-case key in the strategy (the summary guard depends on it).
  assert.deepEqual(JSON.parse(evaluate('JSON.stringify(buildNumberStrategy(["5kg", "6 KG", "200元"]).units)')), { kg: 2, 元: 1 });
  assert.deepEqual(JSON.parse(evaluate('JSON.stringify(buildNumberStrategy(["5 usd", "6 USD", "US$7"]).units)')), { USD: 3 });
});

test("v3.3.0 units: two units in a column read none; a currency sign beside its code is no conflict", () => {
  const { evaluate } = createScriptContext();
  const column = (values) => JSON.parse(evaluate(`JSON.stringify((() => { const values = ${JSON.stringify(values)}; const s = buildNumberStrategy(values); return [s.unit, s.unitConflict, s.units, values.map((v) => toNumber(v, s))]; })())`));
  assert.deepEqual(column(["200 USD", "300 EUR"]), ["", true, { USD: 1, EUR: 1 }, [null, null]]);
  assert.deepEqual(column(["10 kg", "10 lb"]), ["", true, { kg: 1, lb: 1 }, [null, null]]);
  assert.deepEqual(column(["5,400 USD", "$900"]), ["USD", false, { USD: 1 }, [5400, 900]]);
  assert.deepEqual(column(["1 hour", "8 hours", "8 hrs"]), ["hours", false, { hours: 3 }, [1, 8, 8]]);
});

test("v3.3.0 units: the standardized table writes English units as plain numbers; only the number decides scientific notation", () => {
  const { evaluate } = createScriptContext();
  const cell = (value, column = [value, value]) => JSON.parse(evaluate(`JSON.stringify((() => { const s = buildNumberStrategy(${JSON.stringify(column)}); const r = standardizeNumberCell(${JSON.stringify(value)}, s, [], ""); return [r.value, r.action]; })())`));
  for (const [value, expected] of [["200 EUR", "200"], ["2 years", "2"], ["10 miles", "10"], ["5 pieces", "5"], ["6 feet", "6"], ["3 weeks", "3"], ["1 minute", "1"], ["USD 1,050.50", "1050.50"], ["US$200", "200"], ["5 lbs.", "5"], ["1.2万 USD", "12000"]]) {
    assert.deepEqual(cell(value), [expected, "number"], value);
  }
  for (const value of ["1e3", "1.2E5", "-3e-2"]) assert.deepEqual(cell(value), [value, "keptScientific"], value);
  // Chinese units and plain numbers as before.
  assert.deepEqual(cell("2,200元"), ["2200", "number"]);
  assert.deepEqual(cell(" 300 "), ["300", "trim"]);
});

test("v3.3.0 units: a text column says nothing about units it read, only about a conflict; number columns as before", () => {
  const { context, evaluate } = createScriptContext();
  const lines = (field, values) => {
    context.__rows = values.map((value) => ({ [field]: value }));
    return JSON.parse(evaluate(`JSON.stringify((() => { const profile = buildColumnProfile(${JSON.stringify(field)}, __rows); return [profile.typeKey, buildNumberConversionSummary([profile]).map(formatNumberConversion)]; })())`));
  };
  // Repeated, so the column is text, not an ID column.
  const titles = ["Jaws", "Heat", "Up", "Cars", "Big", "127 Hours", "Jaws", "Heat", "Up", "Cars", "Big", "127 Hours"];
  assert.deepEqual(lines("title", titles), ["category", []]);
  const mixed = ["2 kg", "3 斤", "4 kg", "5 斤", "6 kg", "7 斤", "2 kg", "3 斤", "4 kg", "5 斤", "6 kg", "7 斤"];
  assert.deepEqual(lines("重量", mixed), ["category", ["「重量」：同时出现 kg（6 个）和 斤（6 个），没有按单位读取，这些值不能读成数字。"]]);
  const usd = Array.from({ length: 12 }, (_, i) => `${100 + i * 7} USD`);
  assert.deepEqual(lines("Amount", usd), ["numeric", ["「Amount」：按\"USD\"读取了 12 个值（100 USD、107 USD、114 USD）"]]);
  const kg = Array.from({ length: 12 }, (_, i) => `${1 + i}kg`);
  assert.deepEqual(lines("重", kg), ["numeric", ["「重」：按\"kg\"读取了 12 个值（1kg、2kg、3kg）"]]);
});

// ---- v3.3.0 ⑥: accounting dashes and brackets after a currency sign ----
test("v3.3.0 accounting: a dash after a currency sign is zero; brackets after it are negative; a lone dash is not a number", () => {
  const { context, evaluate } = createScriptContext();
  const read = (values) => {
    context.__values = values;
    return JSON.parse(evaluate("JSON.stringify(__values.map((value) => toPlainNumber(value)))"));
  };
  const yen = String.fromCharCode(0xffe5);
  assert.deepEqual(read(["$-", " $-   ", "$ -", "\u20ac-", `${yen}-`, "\u00a3-", "\u00a5-"]), [0, 0, 0, 0, 0, 0, 0]);
  assert.deepEqual(read(["$(1,200.00)", "$( 5 )", "$(0)", "\u20ac(12.5)"]), [-1200, -5, 0, -12.5]);
  // As before.
  assert.deepEqual(read(["($1,200.00)", "-$1,200.00", "$-1,200.00", "$900", "-$5", "(5)"]), [-1200, -1200, -1200, 900, -5, -5]);
  // Still not numbers: a lone dash (a maybe-tier placeholder), two dashes, a sign before the
  // currency with a dash, brackets with a dash, a dash then a percent, nested brackets.
  assert.deepEqual(read(["-", "$--", "-$-", "($-)", "$-%", "$((5))", "$(5", "$ ( )"]), [null, null, null, null, null, null, null, null]);
  assert.equal(evaluate('placeholderTier("-")'), "maybe");
});

test("v3.3.0 accounting: the standardized table writes $- as 0 and $(1,200.00) as -1200.00, reading back the same number", () => {
  const { context, evaluate } = createScriptContext();
  const values = [" $-   ", "$ -", "\u20ac-", "$(1,200.00)", "$( 5 )", "$(0)", "$(0.50)", "\u20ac(12.5)", "($1,200.00)", "$900"];
  context.__values = values;
  const cells = JSON.parse(evaluate(`JSON.stringify(__values.map((value) => {
    const result = standardizeNumberCell(value, buildNumberStrategy(__values), [], "");
    return [result.value, result.action, toPlainNumber(result.value) === toPlainNumber(value) || (toPlainNumber(result.value) === 0 && toPlainNumber(value) === 0)];
  }))`));
  assert.deepEqual(cells, [
    ["0", "number", true], ["0", "number", true], ["0", "number", true],
    ["-1200.00", "number", true], ["-5", "number", true], ["0", "number", true], ["-0.50", "number", true], ["-12.5", "number", true],
    ["-1200.00", "number", true], ["900", "number", true]
  ]);
  // A lone dash stays as written (a placeholder that is not confirmed).
  assert.deepEqual(JSON.parse(evaluate('JSON.stringify(standardizeNumberCell("-", buildNumberStrategy(["-", "5"]), [], ""))')), { value: "-", action: "keptPendingPlaceholder" });
});

// ---- v3.3.0 ⑦: 12-hour clock and ordinal days ----
// [toDate with mdy, dateStyleOf with mdy, dateTextHasTime, standardizeDateCell with mdy]
const readDate = (evaluate, value) => JSON.parse(evaluate(`JSON.stringify((() => {
  const value = ${JSON.stringify(value)};
  const strategy = { order: "mdy" };
  const date = toDate(value, strategy);
  const cell = standardizeDateCell(value, strategy);
  return [date ? date.toISOString().slice(0, 10) : null, dateStyleOf(value, strategy), date ? dateTextHasTime(value.trim()) : null, cell.value + " / " + cell.action];
})())`));

test("v3.3.0 dates: a 12-hour clock after a space, hours 1-12; the date is kept and the cell stays as written", () => {
  const { evaluate } = createScriptContext();
  for (const [value, date, style] of [
    ["3/5/2025 2:30 PM", "2025-03-05", "mdy"], ["3/5/2025 2:30:15 PM", "2025-03-05", "mdy"], ["3/5/2025 2:30 p.m.", "2025-03-05", "mdy"],
    ["3/5/2025 2:30pm", "2025-03-05", "mdy"], ["3/5/2025 12:00 AM", "2025-03-05", "mdy"], ["3/5/2025 12:59 am", "2025-03-05", "mdy"], ["3/5/2025 9:05 A.M.", "2025-03-05", "mdy"],
    ["2025-03-05 2:30 PM", "2025-03-05", "ymdDash"], ["2025/3/5 11:59:59 PM", "2025-03-05", "ymdSlash"], ["2025年3月5日 2:30 PM", "2025-03-05", "chinese"]
  ]) assert.deepEqual(readDate(evaluate, value), [date, style, true, `${value} / keptTime`], value);
  // Hours outside 1-12 with AM / PM, a broken clock, and (R5) a 12-hour clock after T are not dates.
  for (const value of ["3/5/2025 0:30 AM", "3/5/2025 13:30 PM", "3/5/2025 2:60 PM", "3/5/2025 2:30 XM", "3/5/2025 2:30 PMM", "3/5/2025 2:30 P", "3/5/2025 2:30 A", "3/5/2025 2:30 P.", "2025-03-05T2:30 PM", "2025-03-05T02:30 PM"]) {
    assert.deepEqual(readDate(evaluate, value).slice(0, 2), [null, null], value);
  }
  // The 24-hour clock is as before, with T too.
  for (const value of ["3/5/2025 14:30", "2025-03-05T14:30", "2025-03-05 00:30"]) assert.equal(readDate(evaluate, value)[0], "2025-03-05", value);
});

test("v3.3.0 dates: an English day may be an ordinal with the right suffix", () => {
  const { evaluate } = createScriptContext();
  for (const [value, date] of [
    ["March 5th, 2025", "2025-03-05"], ["5th March 2025", "2025-03-05"], ["March 1st, 2025", "2025-03-01"], ["March 2nd, 2025", "2025-03-02"], ["March 3rd 2025", "2025-03-03"],
    ["March 11th, 2025", "2025-03-11"], ["March 12th, 2025", "2025-03-12"], ["March 13th, 2025", "2025-03-13"], ["March 21st, 2025", "2025-03-21"], ["March 22nd, 2025", "2025-03-22"],
    ["March 23rd, 2025", "2025-03-23"], ["March 31st, 2025", "2025-03-31"], ["Wednesday, March 5th, 2025", "2025-03-05"], ["Mar 5TH, 2025", "2025-03-05"], ["March 5th, 2025 2:30 PM", "2025-03-05"]
  ]) assert.equal(readDate(evaluate, value)[0], date, value);
  for (const value of ["March 1th, 2025", "March 11st, 2025", "March 12nd, 2025", "March 13rd, 2025", "March 2th, 2025", "March 22th, 2025", "the 5th of March 2025", "March 32nd, 2025", "5th 3 2025"]) {
    assert.equal(readDate(evaluate, value)[0], null, value);
  }
  // Without a time it is rewritten like any English date; the weekday is still not checked.
  assert.deepEqual(readDate(evaluate, "March 5th, 2025"), ["2025-03-05", "english", false, "2025-03-05 / date"]);
});

test("v3.3.0 dates: neither a 12-hour clock nor an ordinal is a new date style", () => {
  const { evaluate } = createScriptContext();
  const style = (value) => evaluate(`dateStyleOf(${JSON.stringify(value)}, { order: "mdy" })`);
  assert.equal(style("3/5/2025 2:30 PM"), style("3/6/2025"));
  assert.equal(style("2025-03-05 2:30 PM"), style("2025-03-06"));
  assert.equal(style("March 5th, 2025"), style("Mar 6, 2025"));
  assert.equal(style("5th March 2025"), style("6 March 2025"));
});

// ---- v3.3.0 ⑧: two-digit years ----
const dateIn = (evaluate, value, order) => evaluate(`(() => { const d = toDate(${JSON.stringify(value)}, { order: ${JSON.stringify(order)} }); return d ? d.toISOString().slice(0, 10) : null; })()`);

test("v3.3.0 two-digit years: slashes and Excel's d-mmm-yy only; 00-29 are 2000-2029, 30-99 are 1930-1999", () => {
  const { evaluate } = createScriptContext();
  for (const [value, none, mdy, dmy] of [
    ["3/5/25", null, "2025-03-05", "2025-05-03"], ["03/05/25", null, "2025-03-05", "2025-05-03"],
    ["13/5/25", "2025-05-13", "2025-05-13", "2025-05-13"], ["3/13/25", "2025-03-13", "2025-03-13", "2025-03-13"], ["5/5/25", "2025-05-05", "2025-05-05", "2025-05-05"],
    ["12/31/99", "1999-12-31", "1999-12-31", "1999-12-31"], ["1/1/00", "2000-01-01", "2000-01-01", "2000-01-01"], ["1/1/29", "2029-01-01", "2029-01-01", "2029-01-01"], ["1/1/30", "1930-01-01", "1930-01-01", "1930-01-01"],
    ["5-Mar-25", "2025-03-05", "2025-03-05", "2025-03-05"], ["05-Mar-25", "2025-03-05", "2025-03-05", "2025-03-05"], ["5-March-25", "2025-03-05", "2025-03-05", "2025-03-05"], ["5-Mar-25 2:30 PM", "2025-03-05", "2025-03-05", "2025-03-05"],
    ["3/5/25 14:30", null, "2025-03-05", "2025-05-03"], ["2/30/25", null, null, null]
  ]) assert.deepEqual(["none", "mdy", "dmy"].map((order) => dateIn(evaluate, value, order)), [none, mdy, dmy], value);
  // The column's day / month order is decided as before, two-digit years counted.
  assert.deepEqual(JSON.parse(evaluate('JSON.stringify(["3/13/25", "3/5/25"].map((v, i, all) => [buildDateStrategy("x", all).order, toDate(v, buildDateStrategy("x", all)).toISOString().slice(0, 10)]))')), [["mdy", "2025-03-13"], ["mdy", "2025-03-05"]]);
  assert.equal(evaluate('buildDateStrategy("x", ["13/3/25", "5/3/25"]).order'), "dmy");
  // Known: a year-first two-digit date reads day first (25/03/05 is 2005-03-25).
  assert.equal(dateIn(evaluate, "25/03/05", "none"), "2005-03-25");
});

test("v3.3.0 two-digit years: not after dots or dashes (version numbers stay text), not in Mar 5, 25", () => {
  const { context, evaluate } = createScriptContext();
  for (const value of ["3.5.25", "3-5-25", "1.13.22", "13.5.25", "13-5-25", "Mar 5, 25", "Jan 1, 25", "5 Mar 25", "March 5, 25", "Wed, 5-Mar-25", "5-Mar-2", "5/3/2", "5/3/025"]) {
    assert.deepEqual(["none", "mdy", "dmy"].map((order) => dateIn(evaluate, value, order)), [null, null, null], value);
  }
  context.__rows = ["1.13.22", "1.13.23", "2.0.1", "2.1.10", "1.13.22", "3.0.0"].map((v) => ({ 版本: v }));
  assert.equal(evaluate('buildColumnProfile("版本", __rows).typeKey'), "category");
});

test("v3.3.0 two-digit years: a style of their own (R1); the standardized table writes four digits", () => {
  const { evaluate } = createScriptContext();
  const style = (value, order = "mdy") => evaluate(`dateStyleOf(${JSON.stringify(value)}, { order: ${JSON.stringify(order)} })`);
  for (const value of ["3/5/25", "03/05/25", "3/13/25", "13/5/25", "5/5/25", "5-Mar-25", "05-Mar-25", "3/5/25 2:30 PM"]) assert.equal(style(value), "twoDigitYear", value);
  assert.equal(style("3/5/2025"), "mdy");
  assert.equal(style("Mar 5, 2025"), "english");
  // The others are as before; Excel date cells are not text.
  assert.equal(evaluate('dateStyleOf(new Date(Date.UTC(2025, 2, 5)), {})'), "excelCell");
  assert.deepEqual(JSON.parse(evaluate('JSON.stringify(["3/5/25", "5-Mar-25", "3/5/25 2:30 PM", " 3/5/25 "].map((v) => standardizeDateCell(v, { order: "mdy" })))')), [
    { value: "2025-03-05", action: "date" }, { value: "2025-03-05", action: "date" }, { value: "3/5/25 2:30 PM", action: "keptTime" }, { value: "2025-03-05", action: "date" }
  ]);
  // Order unknown: kept, and said why.
  assert.deepEqual(JSON.parse(evaluate('JSON.stringify(standardizeDateCell("3/5/25", { order: "none" }))')), { value: "3/5/25", action: "keptDateOrder" });
  assert.deepEqual(JSON.parse(evaluate('JSON.stringify([Lp("文本 两位年份", "text with a two-digit year"), DATE_STYLE_NAMES.twoDigitYear])')), [["文本 两位年份", "text with a two-digit year"], ["文本 两位年份", "text with a two-digit year"]]);
});

test("v3.3.0 accounting: CSV export does not quote $- or $(1,200.00), and keeps a negative as it was", () => {
  const { evaluate } = createScriptContext();
  assert.deepEqual(JSON.parse(evaluate('JSON.stringify(["$-", "$(1,200.00)", "-$1,200.00", "-5", "-", "=1+1"].map(sanitizeCsvExportValue))')), ["$-", "$(1,200.00)", "-$1,200.00", "-5", "'-", "'=1+1"]);
});
