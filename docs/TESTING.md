# 表格体检测试指南

## 1. 运行测试

要求 Node.js 18 或更高版本。当前测试只使用 Node 内置的 `node:test`、`assert` 和 `vm`，不需要安装 npm 依赖。

运行核心纯数据处理测试：

```bash
npm test
```

运行发布前完整回归：

```bash
npm run test:regression
```

完整回归会执行四个测试入口：

```bash
node --test \
  tests/data-processing.test.cjs \
  tests/destructive-qa.test.cjs \
  tests/basic-data-processing.test.cjs \
  tests/core.test.cjs
```

GitHub Actions 在每次 push 和 Pull Request 时使用 Node.js 22 运行 `npm run lint --if-present` 和 `npm run test:regression`。当前项目没有 `lint` script，因此 lint 步骤会跳过。

## 2. 测试目录

| 路径 | 用途 |
| --- | --- |
| `tests/test-context.cjs` | 创建 VM 上下文和轻量 `FakeElement`，按 `index.html` 中的顺序逐个加载全部应用脚本（每个文件单独执行，共享同一个全局作用域）。返回的 `source` 是所有应用脚本按顺序拼接的文本，并导出文件列表 `APP_SCRIPT_FILES`。 |
| `tests/data-processing.test.cjs` | 纯数据逻辑：数值/日期解析、字段与 ID 识别、缺失、重复（重复行和多出来的行）、统计、IQR、筛选、`detectHeaderRow()` 的正例 / 反例 / 已知残留、Excel 行数上限、表头位置说明的摘要截断、`lineNumbersAtOffsets()` 与原来逐个偏移用正则计数的一致性（含随机文本），以及 `buildIdDuplicateSummary()` 的归一、字段顺序、排序、疑似非唯一阈值的边界、整行相同的分组、截断和 ID 重复值的说明文字。疑似缺失值：清单和两档的内容、`placeholderKey()` 的去空格 + NFKC + 整格匹配（全角写法能匹配，“Nan”“缺考”和部分匹配不算）、`buildPlaceholderSummary()` 的原文、行号截断和被忽略字段只列已确认的写法；护栏测试“placeholders never change the date strategy of a field without a year name”固定预估捷径的前提：清单里全部写法，以及 Ｎ／Ａ、－、／、－－、ＮａＮ、⽆、﹘、﹣ 这些经 NFKC 变成清单写法的字符，加进含 dmy、mdy 证据的日期值后，非年份名字段的 `buildDateStrategy()` 结果不变，`allowYearOnly` 为 false。改动清单或日期策略时，这条测试失败说明预估不能再走捷径。数字策略（v2.5.0）：应该转和不应该转的写法（不带策略、带本列策略、带“每个白名单单位都打开万/亿和全角”的最宽策略），单位冲突和货币符号与单位同列。四条护栏：“without a strategy, or with one that has no switch, toNumber() is the plain parser”固定单参数 `toNumber()` 和没有开关的策略都与 `toPlainNumber()` 逐值相同；“placeholder spellings never convert or give number strategy evidence (contract 17 premise)”固定契约 17 的前提：清单的全部写法、全角形式和带空格的全角形式，在上述最宽策略下都不能转换，`matchNumberWithUnit()` 也不匹配，只含这些写法的列算出的策略没有任何开关；“万 and 亿 shift the decimal point and give exact integers”用 `strictEqual` 和 `Number.isInteger` 固定 `0.07万` = 700 等结果，改回浮点乘法会失败；“getFieldNumberParser() is the plain parser unless the field's strategy has a switch”固定没有开关时取到的就是 `toPlainNumber` 本身。 |
| `tests/basic-data-processing.test.cjs` | CSV/Excel 行转换、中文数据、统计联动、原始质量视角和 CSV 导出的集成测试；包括跳过表头上方标题行的端到端导入、文件行号换算、状态栏与报告的“表头位置”、`state.headerDetection` 的清空与保持，以及回退提示。数据行的文件行号：Excel 的 `!ref` 偏移和中间空行、stage3 夹具、同一份 CSV 内容的 LF / 单独 CR / CRLF+BOM 三种写法、表头单元格里的引号内换行、第二遍扫描条数对不上时全部为 `null`、没经过导入的行查不到行号。护栏：只有表头、空文件、少一列、多一列这几种 CSV 的报错和解析警告与改前逐字相同；超过行数上限的 CSV 只调用一次 `Papa.parse`。ID 重复值：stage3 夹具和 `sample-user-behavior.csv` 的结果、Excel 数字与文本合并、截断写法、字段配置和筛选下的变化、新导入清空；质量页的 pill 与详情（含没有数值字段的表和转义）；HTML 报告的“编号重复”一节、摘要、转义和旧快照兼容。疑似缺失值（文件末尾）：订单表的 5 条检出和行号、成绩表和全部示例（`examples/` 与 `tests/fixtures/examples/`）无检出（v2.11.0 起按钮示例 `sample-orders.csv` 恰好检出金额「N/A」和销售代表「—」两条）、忽略字段、新导入清空；确认 / 撤销 / 全部确认后的 profile、`totalMissing`、Top 10、筛选选项与匹配、分组、ID 重复值和年份列的日期策略；手动类型记录（含“确认后再对别的字段应用配置，撤销后手动类型仍在”）、有草稿时拒绝、恢复自动识别不撤销确认；`comparePlaceholderPreviews()` 逐条和整字段比较预估与完整重跑的 `buildColumnProfile()`；护栏：确认前后 `state.rows`、数据预览、CSV 导出和 `duplicateRows` 不变；兜底测试“after confirming every order placeholder, no analysis output shows N/A or a placeholder dash”在全部确认后检查分析区所有元素和图表标签（用替身 Chart 记录），只剥掉应用自己写的“当前类型无需转换”那个“—”，漏改后会让 N/A 或占位符「—」出现在分析区或图表里的，它都会失败；只影响计数的地方由各自的测试覆盖；质量页的静态小节、pill、详情文字、委托的点击处理，以及报告的“像是空值的写法”一节、摘要、转义和旧快照兼容。数字策略（文件末尾，v2.5.0）：订单表金额按“元”读取、格式问题 3 个值、金额统计值，其余 stage3 夹具不变；15 个“xx元”从 ID 变成数值，4G / 5G 和 8G 类写法不变；万/亿列的直方图按值计数；带单位的列上 `comparePlaceholderPreviews()` 和确认 / 撤销前后 `numberStrategy` 不变；数值筛选的范围和判定同口径；手动改类型；表头检测把 5kg 当文本；CSV 导出的 `-2200元` 仍加 `'`；主分析按大区、按日的元金额合计。疑似数值列：排除疑似缺失值后的口径、样例不含疑似写法、跟随已应用类型和确认、筛选不重算、重置清空。显示单位：字段标签的“已含单位”规则，统计表、图表标题和 y 轴带“（元）”而下拉框仍是原字段名，质量页两个小节的文字和隐藏，报告快照的三个新字段、两节内容、摘要只有“像是数字的文本列”的计数、旧快照兼容和转义。v2.6.0：空记录占满行数预览（6 万行每行后跟空行导入全部 60,000 行、首末行号正确；表头下 5 个空行 + 100,001 行报超限；正好 10 万行 + 结尾空行为护栏；带标题行的重试路径同样成立；第一次主解析仍带 `preview: 100001`，调用顺序固定为主解析、计数、整份重解析、行号）。精确行号：表头下空行、引号内换行、CRLF / 单独 CR、只有空白被丢掉的行、InvalidQuotes（`row` 计入表头和空记录）、带标题行的重试，警告和“第 N 行”报错（多一列、引号没有闭合）都按记事本的行号断言；重试里才发现的多一列显示重试自己的第 253 行和表头说明；25 行列数不足计 25 条、状态栏千分位、条数对不上时行号为 `null`。调用次数护栏：订单表 2 次、两个成绩表各 4 次（`成绩表-标题无逗号.csv` 第一次解析就报多一列，行号惰性求值才保持不变），第一行是表头又多一列时为主解析、表头检测、行号 3 次。中文警告：每个 code 有 / 无行号的消息（纯函数在 `data-processing`），分号 / 制表符 / 竖线两列文件各一条 `UndetectableDelimiter` 中文提示，单列“字段”文件 0 条（护栏），逗号多列文件不再列出分隔符警告。共用判断（护栏）：stage3 全部夹具和全部示例 CSV（`examples/` 加 `tests/fixtures/examples/`，加订单表全部确认后）上每个字段的数值、日期两种类型，`getProfileConversionCheck()` 数出的失败数都等于 `conversionFailuresByType[类型]`；profile 的 `confirmedKeys` 跟随确认 / 撤销且是副本；`groupIdValueIndexes()` / `sortIdDuplicates()` 的分组和排序，用“次数最多的值不最先出现”的数据（A/B/A/B/C/B/C）区分排序规则，`buildIdDuplicateSummary()` 同序。问题行：`problemRowCheck()` 在 15 个文件上逐类核对“列出的数量 = 质量页计数”（整行重复为列出行数 − 组数，ID 重复核对值的个数、前 10 个值的顺序、行数和行号）和顺序，订单表“全部确认 → 改类型 → 撤销”每一步都核对；ID 重复按表格顺序且超过 10 个值时完整列出；整行重复的“保留”；各类说明文字；解析警告含被丢掉的空白行和没有行号的条目。数据预览分页：每页 50 行、翻页和按钮禁用、筛选从末页回第 1 页（用筛选后仍有 2 页的条件，夹取碰不巧）、字段配置回第 1 页、页码夹取和没有分页 state、静态结构和绑定。查看区：每个“查看”按钮只带 `data-problem-index` 并打开自己的类别，“查看”列每行都有元素，分页、四种计数单位的摘要、下拉框文字、焦点、字段配置 / 确认 / 撤销后的找回和关闭（含没有数值字段、`renderQuality()` 提前 return 的数据）、导入关闭、筛选不影响、转义、关闭后焦点回到“查看”按钮。问题行导出：经真实下载路径捕获 Blob，核对 BOM、列顺序、“说明”列、公式中和（含额外列和表头）、`-2200元` 仍加 `'`、列名冲突、整类而不是当前页、没有对应行的解析警告、文件名（N/A、长字段名不截掉日期）、门禁和状态行；数据导出经新的共用函数逐字节不变（护栏；v3.1.0 起只有“当前分析的行”一种，另加有不计入分析的行时 40 行的一例）。v2.8.0 的测试见第 7 节。v2.8.1 的测试见第 8 节。 |
| `tests/destructive-qa.test.cjs` | 空数据、异常表头、编码、极端数值、连续上传、反复筛选、报告后继续分析和 100,000 行压力测试。 |
| `tests/i18n-guard.cjs` | v3.2.0 静态护栏的分词器和规则（不是测试入口，由 `basic-data-processing.test.cjs` 读取）：把 9 个应用脚本切成注释、字符串、模板和代码，找出含汉字或全角标点的字面量，判断它是不是 `L` / `Lp` 的中文参数或白名单（按“文件 + 声明名 + 字面量个数”登记）里的数据；`checkIndexHtml()` 检查 `index.html` 每段含汉字的文字和属性都有键、每个键都有英文；带审查时的 13 个自测用例。LF 换行。 |
| `tests/core.test.cjs` | 启动自检、应用脚本顺序和顶层名字唯一性、DOM ID、相对资源路径、vendor 依赖与 SRI、许可证、`.nojekyll`、字段修正、报告生成和其他静态/集成断言。 |
| `tests/fixtures/` | `english-date-corpus.cjs` 被 `tests/data-processing.test.cjs` 引用，用于英文日期解析断言；`stage3/` 见下文；其余 CSV/XLSX 文件目前只用于手工回归，自动测试没有读取。 |
| `tests/fixtures/examples/` | `sample-data.csv`、`sample-sales.csv`、`sample-students.csv`、`sample-survey.csv`、`sample-used-products.csv`、`sample-user-behavior.csv`：只给测试用的 6 个示例（后 5 个 v2.7.0 从 `examples/` 移来，`sample-data.csv` v2.11.0 移来，字节不变）；`examples/` 里只有“使用示例数据”按钮加载的 `sample-orders.csv` 和 v3.2.0 英文界面加载的 `sample-orders-en.csv`。`basic-data-processing.test.cjs` 经 `listExampleCsvFiles()` / `readExampleCsv()` 读两个目录，覆盖 8 个文件。内容按 utf8 文本读取，LF 和 CRLF 都通过（v2.11.0 的内容测试去掉 CR 后再比较），所以不需要 `.gitattributes` 规则。 |
| `tests/fixtures/stage3/` | 表头检测的验收夹具，内容全部虚构、固定。`订单表.csv`（GBK、CRLF，表头在第 1 行）和 `花名册.xlsx`（表头在第 1 行）是“不能误跳表头”的反例；`成绩表.csv`（UTF-8 BOM，标题行补了逗号）、`成绩表-标题无逗号.csv` 和 `门店报表.xlsx`（合并单元格标题 + 制表日期 + 空行，第 4 行是表头，含 0.00% 毛利率和合计行）是正例。`分区小计表.csv`（5 行小计 / 总计，带空格、冒号和括号注释的标签）和 `公式错误.xlsx`（4 个字段 7 个错误格，含 #SPILL! 和整行错误的行）是 v2.8.0 加的。`known-issues.json` 是按构造方式手写的已知问题清单和字段类型预期，手工加条目，保持一行一条的格式，不要用 `JSON.stringify` 重写；v2.8.0 的测试读取其中合计行和公式错误的预期，其余条目仍供人工核对；v2.5.0 起订单表金额按“元”读取 2200元、3500元，金额那条预期改为数值率 29/30，`expectedTotals.formatProblems` 从 5 改为 3。`写法不一致.csv`（UTF-8 BOM、CRLF，表头在第 1 行，24 行：日期写法混用、常量列和写法细节不同的值，不可见字符在生成脚本里写 `\u` 转义）是 v2.9.0 加的，它的预期（`dateStyles`、`constantFields`、`lookAlikes`）由测试读取。v3.2.0 加了两个英文夹具：`en-issues.csv`（UTF-8 BOM、CRLF，20 行：按单位读取 kg、像是数字的文本列、整列只有一个值、写法不一样、只有不可见字符的格、读取时有疑问 2 处）和 `en-formula-errors.xlsx`（Margin 列 4 个公式错误，最后一行 Subtotal），预期都在 `known-issues.json`。v3.3.0 加了 `en-subtotal.csv`（UTF-8 BOM、CRLF，20 行：Excel 分类汇总导出的 4 行“East Total”这类小计和一行 Grand Total，另有一列 Report 里的“Weekly Total”是报表名、不是合计），预期也在 `known-issues.json`。`tests/basic-data-processing.test.cjs` 经 `importStage3Fixture()` 读取这 11 个夹具（可选第二个参数是界面语言）；它用真实 PapaParse 包一层计数，并记下每次调用的 `header` 选项，表头在第 1 行的 CSV 正好调用两次（主解析一次，`header: false` 的行号扫描一次）。 |
| `tests/fixtures/stage3/gen-fixtures.cjs` | 重新生成上述 11 个夹具：在仓库根目录运行 `node tests/fixtures/stage3/gen-fixtures.cjs [输出目录]`，不带参数时写回本目录。只依赖 `vendor/` 里的 SheetJS 和 Node 自带的 `TextDecoder`（反推 GBK 编码表），输出与已提交文件逐字节一致，与运行时区无关（SheetJS 写出的 xlsx 在同一台机器上两次生成也逐字节相同）。不单独在 CI 中运行；v3.2.0 起有一条测试在临时目录里跑它，逐字节比较全部夹具。 |
| `tests/create-excel-fixture.mjs` | 辅助夹具生成脚本，不属于 npm 测试入口，并引用未写入 `package.json` 的外部工具。 |

测试环境不是完整浏览器：Chart.js、PapaParse、SheetJS、FileReader 或下载 API 会按测试需要注入替身。真实浏览器布局、触控交互、canvas 像素和 Excel/WPS 打开效果仍需手工回归。

## 3. 如何增加测试

1. 把测试放到最接近风险的文件：纯函数放 `data-processing`，导入/导出联动放 `basic`，边界与压力放 `destructive-qa`，DOM/资源/报告契约放 `core`。
2. 使用 `createScriptContext()` 创建新上下文，避免前一个测试残留 `state`。
3. 通过 `context` 放入测试数据，再用 `evaluate()` 调用应用脚本中的真实函数。函数都是全局的，调用时不需要关心它在哪个文件。
4. 优先断言输入不变、返回值和关键状态，不要断言无业务价值的 CSS 细节。
5. 修改导入、筛选或字段类型时，同时断言 `state.rows` 未被改变。
6. 修改筛选分析时，同时验证 `numericStats/categoryStats` 更新且原始质量统计保持不变。
7. 修改导出时，验证字段顺序、UTF-8 BOM、中文、引号/换行和公式前缀处理。
8. 先运行目标测试，再运行 `npm run test:regression`。

示例：

```js
const test = require("node:test");
const assert = require("node:assert/strict");
const { createScriptContext } = require("./test-context.cjs");

test("new behavior", () => {
  const { context, evaluate } = createScriptContext();
  context.__rows = [{ value: "10" }];

  assert.equal(evaluate('toNumber(__rows[0].value)'), 10);
});
```

如果新增 fixture，必须在测试中显式读取并断言；仅把文件放入 `tests/fixtures/` 不会让它进入 CI。

`.gitattributes` 中的 `tests/fixtures/stage3/** -text` 不能删除：仓库设置了 `core.autocrlf=true` 时，Git 会把 CRLF 和 GBK 编码的 CSV 当作文本做换行转换，入库内容会变成 LF，CI（Linux）检出的字节就和本地不同，测试的前提也随之改变。新增需要保持原始字节的夹具时，放进这个目录，或在 `.gitattributes` 里为它单独加 `-text` 规则，并且要和夹具在同一个提交里、先改规则再 `git add`。提交后用 `git ls-files --eol <路径>` 确认显示 `attr/-text`，并用 sha256 比对提交前后的字节。

## 4. 常见失败原因

### `node` 或 `npm` 不可用

确认 Node.js 已加入 PATH。没有 npm 时可直接执行 `package.json` 中对应的 `node --test ...` 命令。

### 新增 DOM API 后测试报错

`FakeElement` 只实现当前测试需要的方法。若生产代码开始调用新的 DOM 方法或属性，应在 `tests/test-context.cjs` 中增加最小替身，不要引入完整浏览器框架来掩盖问题。

### `Papa`、`XLSX`、`Chart` 或 `FileReader` 未定义

VM 只加载 `index.html` 中的应用脚本，不会加载 `vendor/` 中的库文件。测试调用相关流程前，需要同时在 `context` 和必要的 `context.window` 属性上注入替身。

### 替身 Papa 收到第二次调用

经 `parseCsvText()` 导入成功的 CSV，在提交前还会调用一次 `Papa.parse(text, { header: false, skipEmptyLines: false, step, ... })` 求文件行号；这次调用没有 `complete`，只靠同步的 `step`。替身 Papa 要在 `options.header === false` 时，对每条记录同步调用一次 `options.step({ data, errors: [], meta: { cursor } })`，`cursor` 是这条记录结束（含换行）之后的位置，然后直接返回；不要在这次调用里调用 `options.complete`（它不存在，会抛错），也不要用它覆盖记录第一次调用选项的变量。替身给出的记录数和主解析数据行数不一致时，导入照常完成，只是全部行号和警告行号为 `null`。直接调用 `handleParsedData()`、不经过 `parseCsvText()` 的测试没有原文，不会触发第二次调用，警告的行号为 `null`，“数据列多于表头”报错不写“第 N 行”。

报错路径（多一列、引号没有闭合）和空记录占满行数预览的计数扫描也是 `header: false` 的调用，它们会在 `step` 里调用第二个参数的 `abort()`；替身走到这些路径时，要把 `{ abort() {} }` 之类的对象作为第二个参数传给 `options.step`。

### VM 对象深比较失败

VM 中创建的对象属于不同 realm。复杂对象可先 `JSON.stringify()` 再解析到测试进程，或只断言必要字段。

### 日期测试在不同时区表现不同

使用明确格式和 UTC 预期，例如 `YYYY-MM-DDT00:00:00.000Z`。不要用含义不明确的 `MM/DD/YYYY` 作为无策略输入。

### `core.test.cjs` 报 DOM ID 或资源缺失

该测试会从所有应用脚本（拼接后的 `source`）收集 `getElementById()`，并检查 `index.html`、相对资源路径、vendor 脚本标签及其 SRI、许可证文件和 `.nojekyll`。它还会检查应用脚本的固定加载顺序（vendor 脚本在前，`script.js` 最后）和全部应用脚本的顶层名字不重复；每个应用脚本都带与 `package.json` 一致的 `?v=`，由 `data-processing.test.cjs` 的版本一致性测试检查。修改 ID、移动根目录资源或升级 vendor 依赖时必须同步 HTML、JS 和测试契约；新增、删除或改名应用脚本时，还要同步 `core.test.cjs` 中的固定文件列表和 `script.js` 自检里的代表函数列表。

### 压力测试变慢

`destructive-qa.test.cjs` 会处理 100,000 行。先确认是否引入了重复全表扫描、嵌套 `filter()` 或无必要的数据复制，再考虑调整测试；不要通过降低数据量隐藏性能回归。

### Excel 夹具生成脚本失败

该脚本不是标准测试依赖，也不在 CI 中运行。发布判断应以已提交夹具的手工回归和正式测试命令为准；若要把生成流程纳入项目，需先明确依赖和可重复执行方式。

## 5. v2.6.1 新增测试固定了什么

| 测试（文件） | 固定的内容 |
| --- | --- |
| `isMissing() agrees with trim() for every UTF-16 code unit…`（data-processing） | 65,536 个码元分别单独、重复、夹在文字前后中间时，`isMissing()` 与旧函数体（写在测试里）逐值相同 |
| `isMissing() treats exactly the 25 trim() characters as blank…`（data-processing） | 25 个空白字符判空，U+0085、U+180E、U+200B 等不判空；中文、大于 U+3000 的标点、单独代理项、全角空格加中文的结果与旧逻辑相同 |
| `toPlainNumber() reads plain digit strings…`（data-processing） | 15 位以内纯数字走快速路径；309 位以上不能返回 `Infinity`；全角 / 阿拉伯-印度数字、数字后跟字母、日期、会计括号、百分号仍走原路径 |
| `ordinary text and digit strings are read without calling trim() twice`（data-processing，护栏） | 给 `String.prototype.trim` 计数：英文、中文、数字串在 `isMissing()` 里 0 次，`toPlainNumber("12345")` 1 次，首字符是空白时仍调用 trim |
| `median() and the statistics table agree for odd, even, unsorted and overflowing values`（data-processing） | `median()` 的奇偶、乱序、相加溢出分支；统计表的中位数在乱序数据上正确 |
| 统计复用 3 条（basic） | 筛选后统计只含保留的行（最大值放在被筛掉的类别里）、清除 / 清空筛选后回到完整统计、质量统计始终不变；设了筛选但行数不变时统计正确；筛选生效时改字段类型，两份统计都按新 profile 重算 |
| `importing and clearing a filter build the numeric statistics as few times as needed`、`the statistics table takes the median from its sorted values…`（basic，护栏） | 给全局 `buildNumericStats` / `median` 计数：导入 1 次、带筛选 1 次、清除 0 次、重建 1 次；统计表里 `median` 0 次 |
| 分类图默认字段 5 条（basic） | 规则的两个边界（不同值 ≤ 20、首位次数 ≥ 2）各有刚好满足 / 刚好不满足的数据；没有集中字段时选第一个分类字段，没有分类字段时为空；重建保留所选、所选字段不再是分类字段时回到规则、新导入不继承；图表下的分类句针对的也是这个字段。每条先断言数据里各字段的实际类型 |

两条调用次数护栏和 trim 次数护栏钉的是**性能性质**，不是结果：以后合理的实现变化（例如换了函数名、换了复用方式）可以随之更新这些测试，但要在 PR 里说明为什么。

假 DOM 的 `<select>` 在改写 `innerHTML` 后 `value` 不会像浏览器那样重置为第一个选项，所以默认字段的测试只能确认代码显式给 `categoryChartField.value` 赋了哪个字段；`setSelectOptions()` 在浏览器里的重置行为要靠手动 / 无头浏览器验证（v2.6.1 已用无头 Chromium 对照过 v2.6.0）。


## 6. v2.7.0 新增测试固定了什么

v2.7.0 只做删减和把“主要发现”降级为图表下的两句结论。新增 12 条测试，名字都以 `v2.7.0` 开头，都在 `basic-data-processing.test.cjs` 末尾，可以用 `--test-name-pattern="v2.7.0"` 单独运行。每条在对应改动之前都已确认会失败。

| 测试 | 固定的内容 |
| --- | --- |
| `v2.7.0 examples: …` | `examples/` 恰好只有 `sample-orders.csv`（v2.11.0 起；之前是 `sample-data.csv`）；`tests/fixtures/examples/` 恰好是另外 6 个；两个目录的并集是 7 个名字（比只数个数的 `examples.length === 7` 更能发现文件静默丢失） |
| `v2.7.0 workflow: …` | `index.html` 没有 `workflow-status` / `workflowStep`，`style.css` 没有 `workflow-` / `.step-*`，`updateWorkflowSteps` 未定义，`dom` 没有 `workflowStep*` |
| `v2.7.0 quick grouping: the advanced settings block…` | 页面没有快速分组的 id 和“高级设置 / 快速分组”字样，`renderCustomAnalysisChart` 未定义，`dom` 和 `REPORT_CHART_TITLES` / `REPORT_CHART_HEADING_IDS` 里没有对应项，`script.js` 没有 `dom.metricField` 等 |
| `v2.7.0 quick grouping: the report has a custom-analysis section only after…` | 没点“生成自定义分析”时 `buildReportData().customAnalysis.completed` 为 `false`，点过之后为 `true`。它在 v2.7.0 之前的版本上会失败，是因为当时的报告用快速分组的下拉框回推这一节；假 DOM 不会像浏览器那样默认选第一项，所以测试里有一行“如果还有这些下拉框就给它们赋值”，在删除之后只是空操作 |
| `v2.7.0 templates: …` | `templates.js` 不存在，页面 / 样式里没有模板，4 个模板全局名未定义，`dom` 没有 template 键，`REPORT_CHART_TITLES` 恰好 5 项，`script.js` 启动自检恰好列出 8 个应用脚本的代表函数 |
| `v2.7.0 conclusions: the findings section and its navigation entries are gone…` | 没有 `#insightsSection` / `#insights` / “主要发现”；v3.0.0 起改为：只有一个导航，恰好 8 个链接，依次指向 8 个模块 id、末项是 `exportSection`，没有 01–08 编号，没有手机下拉框 `#mobileSectionNav`（`core.test.cjs` 的对应一行同时改了，见第 13 节）；两个新元素在对应的 `.chart-card` 里且默认带 `hidden`；`script.js` 给两个下拉框各绑一个监听；`renderInsights` / `renderInsightGroups` 未定义 |
| `v2.7.0 conclusions: each sentence follows its own chart select…` | 默认选择、改分类下拉框只变分类句（含“取值比较分散”的单字段说明）、改日期下拉框只变趋势句，趋势图的指标和汇总方式不影响趋势句 |
| `v2.7.0 conclusions: import, filter, field configuration and placeholder confirmation all refresh them` | 四个触发点都刷新（确认 N/A 后最多的类别从「N/A」变成「A」）；筛选保持所选字段并按筛选后的行重算；筛选到 0 行两句都隐藏；忽略所选分类字段后改选下一个；所有字段都忽略时两句都不显示 |
| `v2.7.0 conclusions: a field with no values and a new import reset the sentences` | 字段全空时写“当前没有可统计的取值，不做概括。”；`resetAnalysisState()` 清空两句并隐藏 |
| `v2.7.0 conclusions: they do not depend on Chart.js being loaded` | `window.Chart` 为 `null` 与有替身 Chart 两种环境结果相同（`renderChartInsights()` 不能放进 `renderCharts()` 的提前返回之后） |
| `v2.7.0 conclusions: the report shows the sentences on the page…` | 报告 `insights` 恰为页面上的 0–2 句；一句都没有时报告写“暂无自动分析结论。”（v3.2.0 前 HTML、Markdown 各写一次）；旧快照的 6 句原样输出，没有 `insights` 的快照得到 `[]` |
| `v2.7.0 statistics helpers: …` | `groupMetricRowsByField`、`frequencyRows`、`countNonMissingValues`、`aggregateByCategory`、`summarizeNumbers` 都未定义 |

另外几条现有测试因这些改动调整过断言或 setup：分类结论的几条测试读 `state.chartInsights.category`；“re-picks fields from the filtered rows”改成“跟随所选字段”（筛选后图表保留所选字段，再选另一个字段才重新读取）；`core.test.cjs` 里原来经场景模板字段选项断言的“忽略字段 / 全部转换失败的字段不进候选项”改成断言图表下拉框的选项；极端数值（125,000 个值）的 min / max 断言改指 `numericExtents()`；报告的数值单位测试用 `renderV2Analysis()` 生成自定义分析。

假 DOM 的 `<select>` 不会自己选第一项，所以 `populateControls()` 显式给 `#dateChartField` 赋值（浏览器里行为相同）；两个下拉框的 change 监听在假 DOM 里不会触发（`addEventListener` 是空函数），测试直接调用 `renderCategoryInsight()` / `renderTrendInsight()`，监听是否绑好只用源码检查，真实触发要靠浏览器验证。

## 7. v2.8.0 新增测试固定了什么

v2.8.0 加了疑似合计 / 小计行和 Excel 公式错误单元格，并让带说明的状态栏可见。新增 43 条测试，都在 `basic-data-processing.test.cjs` 末尾，41 条名字以 `v2.8.0` 开头，可以用 `--test-name-pattern="v2.8.0"` 单独运行，另 2 条是状态栏的（`the import status keeps has-note…`、`only a successful status without a note…`）。除标为护栏的 1 条外，每条在对应改动之前都已确认会失败。读 `known-issues.json` 的有 3 条：`v2.8.0 errors: the fixture's cells…`、`v2.8.0 fixtures: the store report's row 33…`、`v2.8.0 page: formula errors have one button per field…`。

| 测试 | 固定的内容 |
| --- | --- |
| `v2.8.0 errors: the fixture's cells by field, types and missing counts…` | `公式错误.xlsx` 的错误按字段汇总、类型和缺失数都等于 `known-issues.json`；字段配置不改汇总 |
| `v2.8.0 errors: a sheet with error cells imports exactly like the same sheet with blank cells` | 错误格和空格的行、profile、统计逐字相同，只多了汇总 |
| `v2.8.0 errors: !ref away from A1, cells outside the table, rows and columns that drop out` | 行偏移和列偏移（`!ref` 为 `C5:E8`，表外放错误格）；整行错误的行和整列只有错误的列不列出；`lines` 上限 10 |
| `v2.8.0 errors: the text comes from cell.w of error cells only, then the code table, never #WTF?` | 只有 `t: "e"` 才读 `w`；`w` 不合格式时查代码表；`#WTF?` 走“#错误” |
| `v2.8.0 errors: xls files read the same, CSV text stays text` | xls 同样读出；CSV 的“#DIV/0!”是格式问题，“#N/A”是疑似缺失值 |
| `v2.8.0 state: a later commit and resetAnalysisState() clear the formula errors` | 同一个 import id 再提交（`beginImport()` 的 reset 掩盖不了）和 reset 都清空 |
| `v2.8.0 labels: whole cells only…` | 标签的规范化和 D10 的边界：`（合计）`、`合计（元）（含税）`、`合计[元]`、`合计【元】`、共计、合共不算，`合计：（元）` 算 |
| `v2.8.0 labels: a prefix that is only a status word…` | 前缀整个等于状态词才排除（已汇总不算，南非小计、无锡合计算）；状态列两行“已汇总”、最后“合计 9”的表命中最后一行 |
| `v2.8.0 labels: traditional keywords count…` | 7 个繁体写法，带前缀、带全角空格，“已匯總”不算 |
| `v2.8.0 labels: the prefilter lets through every keyword and every code point…` | 全码位枚举：规范化后可能进关键词的码位都能通过预筛 |
| `v2.8.0 performance: a text file without labels…`（性能护栏） | 没有标签的表不做 NFKC、不读数字（计数） |
| `v2.8.0 performance: a rebuild reuses the labels…`（性能护栏） | 重建复用导入时的标签，不重扫（计数） |
| `v2.8.0 fixtures: the store report's row 33 and the region table…` | 门店报表第 33 行、分区小计表 5 行的范围和对上的列都等于 `known-issues.json` |
| `v2.8.0 fixtures: no other fixture or example…` | 其他夹具和全部示例 CSV 没有合计行（v2.11.0 起 `sample-orders.csv` 有一行，是设计好的）、没有公式错误（误报护栏） |
| `v2.8.0 totals: the three ranges, each tried in order` | `segment`、`above`、`all` 的边界和顺序 |
| `v2.8.0 totals: columns that match different ranges are grouped by range` | 同一行的不同列对上不同范围时分组 |
| `v2.8.0 totals: counterexamples stay silent, the tolerance is 0.5% of the sum` | 反例；误差的两个边界 |
| `v2.8.0 totals: one column is enough, ratio columns and cells that are not numbers stay out` | 一列对上就算；比率列和非数字格不算 |
| `v2.8.0 totals: rows whose label did not add up go back among the details…` | 第二轮：带标签的明细回到明细，第一轮通过的行保留为分隔 |
| `v2.8.0 totals: huge values neither match by overflow nor hide a later segment` | D3：滚动和、有限值判断 |
| `v2.8.0 totals: numbers read with the field's strategy, labels in any field…` | 按字段的数字策略读；标签在任何字段，取字段顺序里第一个 |
| `v2.8.0 refresh: field configuration and confirmations rebuild the totals, filters do not` | 字段配置、确认 / 撤销重算；筛选不碰（同一个数组）；状态栏不变 |
| `v2.8.0 state: a later commit and resetAnalysisState() clear the labels and the totals` | 再提交和 reset 都清空 |
| `v2.8.0 totals: the check changes nothing…`（护栏，改前就通过） | 第 33 行仍算数据：行数、缺失数、统计个数、异常值都不变（`known-issues.json` S5） |
| `v2.8.0 page: both subsections are static…` | 两个小节静态、默认隐藏、位置在解析警告和两栏表格之间，说明文字 |
| `v2.8.0 status: the sentences follow the header note…` | 状态栏的句子：表头说明之后、警告分支、1 行、多行、正好 5 个和超过 5 个行号、没有行号、没有问题时不写 |
| `v2.8.0 page: the store report has one line and one button…` | 门店报表一条一个“查看”，打开 1 行，只有标签格高亮 |
| `v2.8.0 page: five subtotals, one button…` | 分区小计表 5 条一个按钮，第 15 行没有占比；导出的 CSV 和文件名；pill 仍是 6 项 |
| `v2.8.0 page: suspected totals list at most 20 lines…` | 20 条没有超出句，21 条列前 20 条和那句，5,000 条写千分位；查看区是全部行 |
| `v2.8.0 page: formula errors have one button per field…` | 4 个字段 4 个按钮，下标对应各自类别；缺失行的说明；空格高亮；文件名；“字段已忽略” |
| `v2.8.0 page: the sentences for row ranges, many columns and missing row numbers` | 范围的几种写法（连续、中间有其他行、只有 1 行、没有行号）、超过 5 列、千分位 |
| `v2.8.0 problem rows: nine kinds in order…` | 9 类的顺序；确认过的疑似值仍写“已按空值处理”，不被公式错误的说明盖掉 |
| `v2.8.0 page: labels and field names are escaped` | 两个小节都转义 |
| `v2.8.0 page: formula errors show without numeric fields…` | 提前 return 之前渲染：没有数值字段的表显示公式错误；最后一个数值字段被改走时合计小节隐藏，查看区关闭并写“这一类已经没有问题行。” |
| `v2.8.0 page: the sections follow field configuration and confirmations…` | 小节跟随字段配置和确认 / 撤销，筛选和状态栏不变 |
| `v2.8.0 state: a later commit and resetAnalysisState() empty both sections` | 再提交和 reset 都清空小节 |
| `the import status keeps has-note…` | 门店报表、两个成绩表、分区小计表、公式错误、带说明的警告导入有 `has-note`；旧示例 `sample-data.csv`、不带说明的警告导入、之后的普通导入和报告消息没有（按钮现在加载的 `sample-orders.csv` 有，见第 11 节） |
| `only a successful status without a note…` | `style.css` 有新选择器 `:not(.has-note)`，没有旧选择器。假 DOM 看不出元素是否显示，真实可见性要在浏览器里看 |
| `v2.8.0 report: sections and summary counts only when there is something…` | D5：没有命中时节和计数都不输出；旧快照为 `null`；节在“各字段空值情况”之前；字段已忽略 |
| `v2.8.0 report: the note matches the page, at most 20 totals…` | `SUSPECTED_TOTAL_NOTE` 和 `index.html` 同文；20 / 21 条的边界；摘要计数写全部行数（21），不是列出的 20；没有“全部行请点”那半句 |
| `v2.8.0 report: the formula error note matches the page…` | `FORMULA_ERROR_NOTE` 和 `index.html` 同文；HTML 报告里紧跟标题、在列表之前（v3.2.0 前 Markdown 也这样）；只有合计或没有问题的表不出现 |
| `v2.8.0 report: the sections ignore filters, and the snapshot shares nothing with state` | 筛选确实减少分析行时报告不变；快照的每一层都不和 `state` 共用对象 |
| `v2.8.0 report: labels and fields are escaped for each output` | HTML 报告转义标签和字段（v3.2.0 删去 Markdown 后只剩 HTML 这一半） |

另外批准改过 4 处现有断言：`scopes.length` 从 11 改成 13（两处，新夹具）；状态栏那条结尾加第 33 行的句子；`problemRowCheck()` 加 `suspectedTotal`、`formulaError` 两个 `expected` 分支。

假 DOM 的 `getElementById` 对任何 id 都返回替身，也不计算样式，所以这些只能靠浏览器验证：状态栏是否看得见（v3.0.0 起还有：收起后再展开，收起时原文仍在无障碍树里、读屏能读到）、高亮的空格是否看得见、两个小节的排版、390 宽度下“说明”列的最小宽度（`#problemRowsTable td:nth-child(2)`）。用 Excel 打开 `公式错误.xlsx` 看错误是否和页面一致，也要人工做。

## 8. v2.8.1 新增测试固定了什么

v2.8.1 修了 Excel 日期格式下的文本和布尔单元格被读错（见 `docs/ARCHITECTURE.md` 同名一段）。新增 5 条测试，都在 `basic-data-processing.test.cjs`，用真实 SheetJS 按单元格对象建表（`buildCellSheetBuffer()`，共享字符串 `bookSST: true`，和 Excel 的存法一样），经 `parseDataFile()` 导入。现有的 `buildXlsxBuffer()` 走 `aoa_to_sheet`，建不出这种格。

| 测试 | 固定的内容 |
| --- | --- |
| `v2.8.1 a whole column under a date format…`（xlsx、biff8 各一条） | 整列（含表头）都是日期格式时能导入，字段名正确；真日期、“待定”、文本“2021/3/5”各读成自己的值；字段仍是日期，“待定”算 1 个格式问题，缺失为 0 |
| `v2.8.1 empty, blank, boolean and number-like cells…` | 空串、全空格、公式返回的空串都算缺失（共 3 个），布尔读成 `"true"`，文本读成 `"45418"`；不出现 `1899-12-31`、`1900-01-01` |
| `v2.8.1 guard: dates stay dates…`（护栏，改前就通过） | 真日期和带时间的日期仍读成日期；没有日期格式的文本、数字、布尔，`0.00` 格式的数字和文本，结果都和 v2.8.0 相同 |
| `v2.8.1 cells under a date format are found on a sheet that does not start at A1` | `!ref` 为 `B3:C6` 时行、列偏移都正确，文件行号是 4、5、6 |

另外把 `tests/basic-data-processing.test.cjs` 换行测试里字面的 U+FEFF 改成了转义，行为不变。

假 DOM 不计算样式，质量页是否看得见要靠浏览器验证。这些文件都由 SheetJS 生成，用 Excel 亲手做的文件要人工验证（见 `docs/RELEASE_CHECKLIST.md`）。

## 9. v2.9.0 新增测试固定了什么

v2.9.0 加了日期写法混用、常量列和写法细节不同的值三类提示。新增 42 条测试（383 → 425），都在 `basic-data-processing.test.cjs`。另外按批准改了 6 处现有断言：`scopes.length` 13 → 14（两处，新夹具）；花名册类别清单加 `dateStyle / 入职日期 = 6`；`PROBLEM_ENTRY_CONTAINERS` 加 `dateStyleDetails`、`lookAlikeDetails`；`problemRowCheck()` 加 `dateStyle`、`lookAlike` 的 `expected` 和 `lookAlike` 组内顺序；花名册报告摘要加“日期写法混用：1 个字段”。

| 测试 | 固定的内容 |
| --- | --- |
| `v2.9.0 fixture: 写法不一致.csv has the bytes, types and totals…` | 夹具的 BOM、CRLF 和关键格；类型、缺失、格式问题、ID 重复值等于 `known-issues.json` |
| `v2.9.0 guard: everything v2.8.1 computed stays the same…`（护栏，改前就通过） | 13 个文件 × 导入 / 全部确认后，profile、统计、质量数字和现有问题行的摘要等于 `tests/fixtures/v2.8.1-guard.json`；`写法不一致.csv`（v2.9.0）和 `sample-orders.csv`（v2.11.0）比基准新，循环里跳过 |
| `v2.9.0 Excel: date cells are recorded…` / `only real date cells count…` / `state: a commit replaces the date cells…` | 日期单元格按行列记下（标题行、`!ref` 从 B3 开始）；文本、空串、布尔、`0.00` 数字不算；只有时间的格算（已知限制）；提交替换、长度不对不写、重置清空 |
| `v2.9.0 Excel: rows with the same date fields share one frozen record…` / `guard: text columns named like Object members never count as date cells`（护栏，改前就通过） | 日期字段相同的行共用同一个冻结、没有原型的记录（性能：每行新建记录时 10 万行 xlsx 导入 +5.0%，共用后 +3.1%），两个日期字段的行记录正确；名为 constructor、toString、hasOwnProperty 的文本列不算日期单元格（换成 `{}` 时会算） |
| `v2.9.0 dateStyleOf: …` / `guard: dateStyleOf() is null exactly when toDate() is null` | 每种写法的分类（含日或月在前不分分隔符）；超过 10 万个组合穷举“写法 ⇔ `toDate()`” |
| `v2.9.0 date styles: …`（4 条）、`a field read by year has no date styles` | 夹具和花名册、其他文件 0 命中；主写法规则（最多的不最先出现）；补零、时间、读不出的值；跟随字段配置和确认，筛选不重算；按年份读的字段不查 |
| `v2.9.0 a rebuild reads the date cells only…`、`look-alike work runs once per distinct text…`（性能护栏，计数） | 每个不同文本只分类一次；重建复用扫描缓存，新改类型的字段才扫一次 |
| `v2.9.0 state: a later commit and resetAnalysisState() replace …`（两条） | 同字段名再提交不用旧缓存；重置清空 |
| `v2.9.0 texts: the date style line and how whitespace is shown` | “主写法是 Excel”的空格；`formatVisibleText()` 的每种写法 |
| `v2.9.0 constant columns: …`（5 条） | 两个命中文件；5 个的门槛、去首尾空格、1 / 1.0、全角、大小写、全空列、三种类型；疑似缺失值写法不报（确认前后）；确认后新出现 / 消失；只读 profile；重置 |
| `v2.9.0 hasEdgeWhitespace() agrees with trim()…` | 全部 65536 个码元在开头、结尾、中间 |
| `v2.9.0 lookAlikeKey: …`（2 条）、`lookAlikeReasons: …` | 合并 / 不合并的边界（含 U+00AD、U+200E、U+200F，变体选择符不去）；纯 ASCII 不做 NFKC（计数）；原因判断 |
| `v2.9.0 look-alike values: …`（3 条）、`cells of nothing but invisible characters…`、`look-alike lists stop at the limits…` | 夹具等于 `known-issues.json`、其他文件 0 命中；排除疑似缺失值、数值和日期字段；只有不可见字符的格单独计数、进并集一次；10 组 / 10 种上限和排序 |
| `v2.9.0 problem rows: …`（3 条）、`the viewer follows the field configuration…` | 日期写法的行和说明；写法不同按组、说明的可见形式、只有不可见字符的行；类别顺序、标签、按钮、高亮、文件名、导出保留原值；改类型后查看区关闭；常量列没有类别 |
| `v2.9.0 the three sections render without a numeric field…`、`the sections are static, in place…` | 提前 return 之前渲染、转义（字段名也是）、为空隐藏、重置清空页面；位置；pill 仍是 6 项、状态栏不变 |
| `v2.9.0 reports: …`（4 条） | 说明句和 `index.html` 相同；三节内容、顺序、摘要两个计数（常量列不计）；只有不可见字符的计数和转义；旧快照为 `null`、快照每一层都不和 state 共用、筛选不影响 |

**重新生成摘要基准**：`tests/fixtures/v2.8.1-guard.json` 是 v2.8.1（`196d8a9`）上 13 个文件各两种状态的摘要（profile、统计、质量数字、现有问题行，sha256 截取 32 位；不含比它新的 `写法不一致.csv` 和 v2.11.0 的 `sample-orders.csv`），在 Linux + Node 22 和 Windows + Node 24 上生成的结果逐字节相同。以后哪一版**有意**改变了这些结果，就在那一版用改动前的代码重新生成，或者删掉这条护栏，并在 PR 里说明。生成方法：在改动前的提交上，对每个夹具和示例 CSV 经 `parseDataFile()` 导入，按测试里 `v281GuardDigest()` 的同一个表达式算摘要，再 `confirmAllCommonPlaceholders()` 后算一次。

**换了等价的表达式时**（v3.1.0）：摘要里 `dedup` 一项从 `deduplicateRows(state.rows, state.fields).length` 换成 `state.rows.length - state.duplicateRows`（`deduplicateRows()` 删了）。两个表达式在全部状态上逐个相等，所以哈希不变、`v2.8.1-guard.json` 不用重新生成；改前的代码上先实跑一次护栏确认。

**只改了说法时**（v2.12.0）：问题行的说明是被哈希的原文的一部分，界面说法一变，哈希就变，但结果没有变。这时不在改前的代码上重算，而是做等价比较：在改前、改后的代码上各导出一次哈希前的 JSON 原文（同一个表达式，不取哈希，导入顺序和护栏相同），先确认 `JSON.stringify(JSON.parse(原文))` 和原文相同，再只对改前原文的 `problems[*][1][*].note` 做这次的文字替换，要求每个 key 都和改后原文逐字节相等；然后只把变了的那几个 key 的哈希换成改后的值，手工改、一行一条，并把比较脚本和逐 key 结果写进 PR。如果需要 note 以外的替换才相等，说明改的不只是说法，要停下来查。v2.12.0 只做了三处替换（无法转换为数值 → 认不出是数字、无法转换为日期 → 认不出是日期、已按缺失处理 → 已按空值处理），26 个 key 全部相等，变了的是成绩表、成绩表-标题无逗号、花名册、订单表各两种状态共 8 个。

假 DOM 不计算样式，三个小节是否看得见、`␣` 和 `[零宽字符]` 是否显示得出来要靠浏览器验证（见 `docs/RELEASE_CHECKLIST.md`）。

## 10. v2.10.0 新增测试固定了什么

v2.10.0 加了“标准化数据”导出（v2.12.0 起叫“写法统一的表”）（检查表、标准化 CSV、改动清单）。新增 28 条测试（425 → 453）：7 条纯函数在 `data-processing.test.cjs`，21 条集成在 `basic-data-processing.test.cjs`。没有改任何现有断言，没有新增夹具，`v2.8.1-guard.json` 不变。测试数据都是测试里内联的 CSV（`STANDARDIZE_CSV`：编号 / 日期 / 金额（元）/ 区域 / 年份 / 备注各一种类型，`DMY_EVIDENCE_REWRITTEN`：日月顺序只靠一个会被改写的格），看不见的字符一律写 `\u` 转义。

| 测试 | 固定的内容 |
| --- | --- |
| `v2.10.0 standardizeDateCell: the boundary table` | 各种日期写法的导出值和处理：三种分隔符、中文、日或月在前（dmy / mdy / none / conflict）、英文（含星期写错）、带时间和只写时区的保持原值、已是 `YYYY-MM-DD` 只带空格的记“去首尾空格”、保持原值的格保留首尾空格、读不成日期与日月顺序无法确定、疑似缺失值未确认 / 已确认、只有空白的格 |
| `v2.10.0 standardizeNumberCell: the boundary table, as text and never through a float` | 纯数字（含 `007`、`1.50`、`-0`、20 位）不改写只去空格；单位、千分位、空格和 U+00A0 分组、全角、万 / 亿移位、货币、会计负数、`+`、`.5`、`5.` 的改写结果；读出 0 的改写去掉 `-`（`(0)`、`(0.00)`、`-0,000` → `0000`、`+0`、`-0万`）；百分比、科学计数法、读不成、疑似缺失值和货币混用保持原值且保留首尾空格 |
| `v2.10.0 standardizeNumberCell: the safety net keeps a rewrite that would read as another number` | 兜底走得到：测试里把 `toNumber` 换成对 `2,200` 读出 2201，结果是“无法安全改写，保持原值”；别的值照常改写 |
| `v2.10.0 standardizeTextCell: …` | 首尾半角空格、U+3000、U+00A0、U+FEFF 去掉；只有空白的格清空；U+200B、全角字母不改；未确认的疑似缺失值只去空格 |
| `v2.10.0 the action texts` | 13 种处理的文字和 `STANDARDIZE_KEPT_ACTIONS` 的顺序 |
| `v2.10.0 guard: a rewritten date reads as the same day…` / `guard: every rewritten number reads back…`（语料护栏） | 日期：改写后同一天、读不出原样、带时间判断与独立判断一致（4 种顺序）；数字：`toPlainNumber(新值)` 与字段解析结果 `Object.is` 相同（两边都是 0 时按 `===`），改写结果是纯数字且读出 0 时不带 `-`，兜底不触发，保持原值的格逐字不变 |
| `v2.10.0 guard: the one-cell rules never need the safety net…`（护栏，改前就通过的性质） | 14 个文件导入后和全部确认后，数值 / 日期字段逐格：兜底不触发、改写读回同一个数 |
| `v2.10.0 the constructed table: …`、`confirmed placeholders become empty…` | 构造表的类型、每列计数、合计；年份字段按文本去首尾空格、`2021年` 不改写日期；已确认的疑似缺失值在数值、分类、年份字段清空，被忽略的字段整列原样 |
| `v2.10.0 currency: …` | 只有一种货币符号时去掉；`¥` 和 `￥` 算一种；两种以上时带符号的格保持原值；百分比格的符号也计入（`$50%` 让 `￥12` 保持原值） |
| `v2.10.0 guard: the check never runs on import, filter, field configuration or confirmations` | 给 `buildStandardizedExport` 计数：再次导入、筛选、应用配置、确认、恢复自动识别都是 0 次，点检查 1 次 |
| `v2.10.0 a filter keeps the check; import, commit, …` | 筛选、清除筛选不作废；应用配置、恢复、确认、撤销、导入、重置都作废。“直接 `commitTabularData()`”这一步是必要的：页面导入都先经 `beginImport()` → `resetAnalysisState()`，只测导入抓不到“提交时不作废” |
| `v2.10.0 D8: …` | 顺序证据格被改写、带时间的格保持原值时提示，并且再导入确实变成分类；证据格带时间（证据还在）不提示；普通 dmy、mdy 列全部改写不提示且再导入仍是日期；没有确定顺序的字段不看。故意改坏时，去掉“顺序变了”这个条件测不出来，属于等价改坏：两种策略下 `toDate()` 只有日月顺序不同才会一边读得出、一边读不出（`allowYearOnly` 只会让重算后的策略读得更多），所以“有格读不出了”本身就意味着顺序变了；条件保留是为了读起来清楚 |
| `v2.10.0 all fixtures and examples: rows and exports unchanged, counts = list, round trip…`（护栏 + 数字） | 15 个文件（v2.11.0 起）：检查前后 `state.rows` 和两种现有导出逐字相同；每列计数 = 清单里这一列的条数；都不出顺序提示；标准化 CSV 再导入每个字段的类型相同；各文件清单条数（订单表 36、写法不一致 13、分区小计表 14、两个成绩表各 3、花名册 8、sample-orders 11、sample-user-behavior 119，其余 0） |
| `v2.10.0 the check needs data, a finished analysis and no unapplied field draft` | 无数据、有草稿、分析未完成时不计算；有草稿时旧结果保留，撤销草稿后仍在 |
| `v2.10.0 the check table: …`（2 条） | 摘要句、表头、每行的数字和保持原值样例（`␣` 显示空白）、年份 / 单位 / 顺序三种说明、合计、被忽略字段只在摘要句计数、转义；没有改动时只有摘要句、“导出改动清单”禁用 |
| `v2.10.0 the panel follows the check: …` | 筛选后面板、表格、状态行都不变；应用配置、确认、直接提交、重置后面板收起，表格、摘要、状态行清空，两个导出按钮禁用 |
| `v2.10.0 gate: …` | 无数据时三个按钮禁用并写原因；有草稿时三个按钮禁用并写原因、面板仍在；撤销草稿后恢复 |
| `v2.10.0 the page: …`、`the export buttons are bound once each` | `index.html` 的第三个选项、静态隐藏面板和按钮；打印样式隐藏 `.standardized-export-actions`；`script.js` 三个按钮各绑定一次 |
| `v2.10.0 the two downloads: …` | 构造表的改动清单全文（按行再按字段、`'-50元` 中和）和标准化 CSV 全文（`'=1+1` 中和）、文件名、BOM、类型、两条状态行；导出时不再调用 `buildStandardizedExport` |
| `v2.10.0 D13: …` | 上限取 3：超出时只导出前 3 条、文件名带“前3条”、状态行“只导出了前 3 条，共 21 条”；正好等于和不超出时完整导出、文件名不变；检查表计数不变；上限常量 1,048,575；千分位 |
| `v2.10.0 file names: …` | 文件名清理非法字符、日期总在末尾、带“前N条” |
| `v2.10.0 the exports need a check: …` | 没检查过、检查已作废时写“请先点“检查会改哪些格”。”；0 条时改动清单不导出，标准化 CSV 照常 |
| `v2.10.0 all fixtures and examples: the downloads match the check…` | 14 个文件的下载内容与检查结果一致、清单条数 = 计数、全是 CRLF、再导入类型相同 |
| `v2.10.0 the downloads ignore the filter and never guess a file row` | 筛选后两份下载仍是全部行；没有文件行号的行，清单第一列为空 |

**性能**：新代码不在刷新路径里。10 万行 × 20 列（固定种子生成的 CSV，含日期、数值、分类、ID 列），与 v2.9.0 交替各 5 次取中位数：导入 +0.4%、筛选 +1.1%、重建 −0.2%。检查约 1.6 s（干净数据）/ 3.9 s（每格都要处理的 200 万格）；最坏情况导出改动清单到上限约 2.55 s、峰值 RSS 约 1.22 GB。

假 DOM 不算样式：检查表是否可见、1280 / 390 宽度有没有横向滚动、打印时按钮是否隐藏，要靠浏览器验证（见 `docs/RELEASE_CHECKLIST.md`）。v3.0.0 起还要在浏览器里看：停在别的模块时导出 HTML 报告，看不见的模块里的图表仍按原尺寸导出（假 DOM 没有尺寸）。

## 11. v2.11.0 新增测试固定了什么

v2.11.0 换了“使用示例数据”加载的文件：新的 `examples/sample-orders.csv` 是虚构的 2025 年上半年订单明细（40 笔订单 + 1 行整行重复 + 1 行合计 = 42 行数据，9 列，UTF-8 BOM + LF），故意放了七类问题：第 1 行标题行（带尾随逗号）、第 44 行合计、SO-1018 第 20、21 行整行重复、金额「N/A」（第 29 行）和销售代表「—」（第 16 行）、金额「5400元」等 3 个、5 个非主写法的日期、地区「华南 」「华东 」两个结尾空格。旧的 `sample-data.csv` 字节不变地移到 `tests/fixtures/examples/`，原来引用它的测试照常通过。应用代码的判断规则一条没改；新增 6 条测试（453 → 459），都在 `basic-data-processing.test.cjs` 末尾；`v2.8.1-guard.json` 不变。下表按顺序称为测试 ①–⑥。

| 测试 | 固定的内容 |
| --- | --- |
| `v2.11.0 sample: the file holds the planned rows, title and total line` | BOM 三字节；**去掉 CRLF 里的 CR 后**的 sha256 等于 `SAMPLE_ORDERS_LF_SHA256`；结尾换行、44 行、第 1 行标题原文（含尾随逗号）、表头、第 44 行合计原文、第 20、21 行相同 |
| `v2.11.0 sample: the button fetches the new file and names the import after it, like an upload` | `script.js` 有新的 `fetch`、不再出现 `sample-data.csv`（也管住“正在加载 …”文字）；`index.html` 的下载链接；真正点一次按钮：请求的地址、传给 `parseCsvText()` 的名字（处理函数一次、标题行重试一次），以及点击和上传两个入口的状态栏、pill、各小节显隐和全部详情 HTML、数据预览、问题行和行号、类型逐项相同 |
| `v2.11.0 sample: the status line, the six pills and every hint with its rows` | 状态栏全文和 `has-note`；9 列类型；6 个 pill 全文；显示的 4 个提示小节和顺序；**两句图表结论**（分类句“「华东」出现 10 次，占比 24.4%”、趋势句）；疑似合计、按单位读取、日期写法、写法不同四段文字；疑似缺失值两条（档位、行号）；ID 重复；16 类问题行的名字和每类的文件行号 |
| `v2.11.0 sample: confirming every common placeholder moves the two values to missing and keeps the total` | 全部确认后 6 个 pill（空值 8、认不出的数字或日期 0、像是空值的写法 0 个字段）、合计仍检出、金额和销售代表的缺失数 |
| `v2.11.0 sample: the standardized check counts 11 cells, 12 after confirming` | 每列计数（日期改写 5、去首尾空格 2、数字改写 3、保持原值 1），无顺序提示；全部确认后 12 格 |
| `v2.11.0 sample: without its problems nothing is flagged, and each problem alone flags only itself` | 在测试里用代码去掉问题（不另存干净文件）：干净版 6 个 pill 全 0、没有提示小节、没有 `has-note`、标准化 0 格；只放回一类（标题、合计、重复、N/A、—、元、日期、空格）时只出这一类的提示；每个“去掉”函数单独作用于原文件都确实改了东西（防止正则失配后变体悄悄不含问题） |

**改动的 10 条现有断言**（新文件放进 `examples/` 后被按目录枚举的测试自动用上）：`core.test.cjs` 的 `fetch` 路径；`scopes.length` 14 → 15（两处）；`v2.7.0 examples` 的目录清单；疑似合计、日期写法、写法不同三条“no other file”加 `sample-orders.csv`；`v2.9.0 guard` 的示例循环跳过它（和 `写法不一致.csv` 一样比基准新）；`v2.10.0 all fixtures…` 加 `"sample-orders.csv": 11`；疑似缺失值候选那条从“每个示例都没有”改成“**新示例恰好是金额 N/A、销售代表 —，其余示例都没有**”——不写成“跳过新示例”，因为跳过以后新示例多一个或少一个疑似缺失值都测不出来。另有 4 个测试名在末尾注明 v2.11.0 的变化。

**按钮点击怎么测**：假 DOM 的 `addEventListener` 是空函数，`script.js` 的处理函数拿不到。`clickSampleButton()` 不改 `test-context.cjs`，而是利用“所有上下文的假 DOM 元素共用同一个 `FakeElement` 原型”：先建一个探测上下文拿到原型，临时把 `addEventListener` 换成记录函数，再建真正的上下文（`script.js` 的绑定被记下），`finally` 里恢复；再放一个 `fetch` 替身，返回 UTF-8 解码、去掉 BOM 的文本（和 `fetch().text()` 一样）。以后 `test-context.cjs` 改成每个上下文一个元素类时，这个助手要跟着改。包装 `parseCsvText()` 记录名字时**必须转发全部参数**：标题行重试会再调一次 `parseCsvText()` 并带第 4 个参数（含 `lineOffset`），只转发 3 个会让行号全部少 1、状态栏没有表头说明。

**为什么去掉 CR 再算 sha256**：`examples/` 没有 `-text` 规则。仓库和 GitHub Pages 上是 LF（3206 字节）；Windows + `core.autocrlf=true` 重新检出后工作区是 CRLF（3250 字节，sha256 不同），分析结果完全相同（CRLF 算一次换行）。所以测试读成文本、把 `\r\n` 换成 `\n` 后再比较，Windows 工作区和 CI 都能过（实测把文件换成 CRLF 后完整回归照样通过）。**以后有意修改示例文件时，要同步改 `SAMPLE_ORDERS_LF_SHA256`**（以及测试 ③–⑥ 里受影响的数字），并在 PR 里说明。

**故意改坏的结果**（v2.11.0 提交后，本机时区完整回归，改一处跑一次）：

| 改坏 | 失败的测试 |
| --- | --- |
| 按钮 `fetch` 回 `examples/sample-data.csv` | `core.test.cjs`、测试 ② |
| `parseCsvText()` 的显示名不改 | 测试 ② |
| “正在加载 …”文字不改 | 测试 ② |
| 下载链接回旧文件 | `core.test.cjs`（本地资源不存在）、测试 ② |
| 示例删掉 N/A | 疑似缺失值候选那条、`v2.10.0 all fixtures…`、测试 ①③④⑤⑥ |
| 标题行去掉尾随逗号 | 只有测试 ①（去掉逗号后走“多一列”的检测路径，结果完全相同） |
| 合计 800 → 801 | 测试 ①⑥（仍在 0.5% 的容差内，照样判为疑似合计，状态栏和 pill 不变） |
| 某个产品名后多一个空格 | `v2.10.0 all fixtures…`、测试 ①③⑤⑥ |
| 数据换回地区并列第一的第一版（华东、华南各 9） | 只有测试 ①（内容）和测试 ③（分类句） |
| 对照：示例换成 CRLF | 全部通过 |

## 12. v2.12.0 新增测试固定了什么

v2.12.0 统一了界面说法（13 组，见 `docs/ARCHITECTURE.md`“界面说法和代码里的名字（v2.12.0）”），判断规则一条没改。新增 2 条测试（459 → 461），都在 `basic-data-processing.test.cjs` 末尾；跟着说法改了约 200 行现有断言（`basic-data-processing` 196 行、`data-processing` 2 行），`v2.8.1-guard.json` 按第 9 节“只改了说法时”的做法换了 8 个哈希。

| 测试 | 固定的内容 |
| --- | --- |
| `v2.12.0: old interface words no longer appear in pages, reports, file names or exported notes` | 禁用词（21 个）：缺失、未填写、疑似缺失值、疑似值、按缺失处理、格式问题、无法转换、转换失败、异常值、数值异常、写法细节不同、写法不同、解析警告、条警告、ID 字段重复、ID 重复、ID重复、常量列、疑似数值列、标准化、可安全转换。stage3 全部夹具和全部示例 CSV 各走一遍：导入后页面所有元素的 innerHTML / textContent / 选项文字、HTML 和 Markdown 报告（v3.2.0 起只有 HTML）、每个问题行类别的类别名 / aria / 导出文件名 / 全部说明 / 查看区、“检查会改哪些格”之后的面板和两个文件名、全部按空值处理后再检查一次（含改动清单 CSV 的内容）和报告、日期趋势图切到数值字段 + 求和后的标题（必须以“（求和）”结尾）和报告；去掉注释的 `index.html`，以及 `#dateChartAggregate` 的选项（有“求和”、没有“合计”）；页面和报告共用的 7 个说明常量里没有“汉字 空格 汉字” |
| `v2.12.0: charts, insights and the report say 认得出 when a filtered range has no dates or numbers` | 30 行的表（前 25 行组 A 有日期和金额，后 5 行组 B 都为空）筛选到组 B：日期趋势空图“该字段没有认得出的日期”、`state.chartInsights.trend`、HTML / Markdown 报告的趋势节（v3.2.0 起只有 HTML）；手动选“金额”后数值分布空图“该字段没有认得出的数字”。筛选后 `fieldHasNumericData()` 会把“金额”从 `#numericChartField` 去掉（浏览器里也一样，页面显示“没有可用的数值字段”），所以这一句要手动设下拉框才能走到 |

**为什么护栏里有一个内联 CSV**：stage3 和示例里没有一个文件带解析警告，也没有自定义分组的空组，“读取时有疑问（N 处）”“共 N 处”和空组名“（空值）”在夹具上碰不到，“条警告”“未填写”永远查不到。护栏最后导入一个 5 行的 CSV（一行少列、一行备注为空）并按备注做一次自定义分组。同理，“已按空值处理的写法，清空”只在按空值处理之后再检查时出现，而且只在改动清单 CSV 里，所以护栏在全部按空值处理后再点一次检查，并扫改动清单的内容。

**改动的现有断言**：pill 文字和按前缀找 pill 的查找词、状态栏、报告节标题 / 摘要 / 列头、像是空值的写法的详情 / 按钮 / 消息 / 快照 `effect`、问题行的类别名 / aria / 说明 / 文件名 / 查看区摘要、写法统一的表的标题 / 按钮 / 列头 / 文件名 / 状态行、格式一致性三节、趋势图“求和”、自定义分组空组名、`formatStandardizeAction()` 的两个标签。只换说法会**永远通过**的断言另外改了写法：

- 否定匹配换成新词后还要不被别的句子误伤：“编号重复”也出现在 `LOOK_ALIKE_NOTE` 和像是空值的写法的说明里，“像是空值的写法”也出现在像是数字的文本列的提示里，所以旧快照不出这一节的检查（`legacy report snapshots without ID duplicates…`、`report placeholder summaries appear only with candidates, and legacy snapshots…`）匹配标题或带冒号的摘要（`/<h3>编号重复<\/h3>|；编号重复：/` 等），不用裸词。
- 两条审查指令里没列出、grep 找到的：`markdown.indexOf("### ID 字段重复值") < markdown.indexOf("## 数值字段描述性统计")`（标题不在时 -1 < N 仍然成立）和 `doesNotMatch(maybeHtml, /全部确认常见写法/)`；实施时又找到一条：`doesNotMatch(html, /确认后类型会变/)`（v2.12.0 改成“按空值处理后类型会变”）。
- 切片：像是空值的写法小节按 `<h3 class="subsection-title">…</h3>` 切（按裸词“编号重复”切会切到前面的说明句，得到空串）；同一个值，写法不一样一节的否定匹配前先断言标题存在（标题不在时 `slice(-1)` 只剩最后一个字符）。
- 顺序：9 条“A 在 B 之前”改用 `assertInOrder(text, markers)`：每个 marker 先断言存在，再断言位置严格递增。只写 `indexOf(a) < indexOf(b)` 时 a 不在也会通过。报告大标题顺序那 4 条（`<h2>数据规模</h2>` 等）没有涉及改名，没有改。

**故意改坏的结果**（在分支快照上，用软断言记录每一处失败的断言行）：审查时的 15 个变异（让应该隐藏的节 / 计数 / 按钮出现、把标题改回旧说法、报告一节不转义）都让对应断言失败；另加 4 个：`#dateChartAggregate` 选项改回“合计”、`getDateChartAggregateLabel()` 改回“合计”、`LOOK_ALIKE_NOTE` 里加回空格，都让新护栏失败；没有类型变化时也显示“按空值处理后类型会变”，让 `doesNotMatch(html, /按空值处理后类型会变/)` 失败。“认得出的”四句各改回旧文字一次，第二条新测试都失败。

## 13. v3.0.0 新增测试固定了什么

v3.0.0 只换外壳和样式（一次显示一个模块、顶栏和侧栏、可收起的状态栏、新的上传页和“回到分析”、字段配置表 F1 F2、空状态 F3、新外观），检测规则、计数、报告文字和导出文件一样不变。新增 65 条测试（461 → 526），都在 `basic-data-processing.test.cjs`，名字以 `v3.0.0` 开头；`tests/test-context.cjs` 的假 `window` 加了 `scrollTo()` 和 `history.replaceState()`（`showModule()` 要用）。

| 测试组（名字前缀） | 条数 | 固定的内容 |
| --- | --- | --- |
| `v3.0.0 guard:` | 3 | 护栏，在 v2.12.0 上就通过：导入示例、生成自定义分析、应用筛选后 HTML 报告的图表张数和标题；每一句空状态文字在各自的触发条件下仍出现；示例数据的报告文字与 v2.12.0 相同（`buildHtmlReport(report, [])` 的 SHA-256，不含图片，所以不受图表配色影响；v3.0.1 起这一条改名为 `v3.0.1 guard:`，见第 14 节） |
| `v3.0.0 modules:` | 5 | 8 个模块在一个舞台里；看不见的模块用 `visibility`，从不 `display: none` 或 `.hidden`，打印时全部显示；`showModule()` 只显示一个、只标一个链接、写地址、回到顶部、聚焦标题；导入停在第一个模块（v3.4.0 起叫“体检结果”，之前叫“数据概览”）；应用字段配置、按空值处理、筛选后停在当前模块 |
| `v3.0.0 shell:` | 8 | 顶栏的组成，没有 `section-label`；侧栏的静态部分和手机面板里的“换一个文件”；四个模块标题旁的数据范围句（只有第一个模块那句带 id；v3.4.0 起这句只在有筛选或有不计入分析的行时显示）和它们的“清除筛选”、侧栏的“筛选”一行都跟着 `#globalFilterSummary`；侧栏的字段、类型、单位和状态行；Excel 工作表名、表头在第 1 行时不显示、字段名转义；按宽度的默认开合和 Esc；`script.js` 里每个控件只绑一次 |
| `v3.0.0 status note:` | 5 | `#statusPanel` 原样、外面包一层；收起用 `clip-path`，读屏仍能读到；“〈文件名〉已载入，有 N 条说明”的 N；`warning` 不收起（v3.4.0 起导入后默认收起、`warning` 也收起，这两条测试按第 20 节改了）；成功消息只在第一个模块显示，导出时的警告和报错在每个模块都显示 |
| `v3.0.0 upload page:` / `v3.0.0 way back:` | 3 / 3 | 上传页的新文字和示例小表（第 10、12、20、21、29、44 行的值和说明，第 20 行最后一列为空）；报错排在标题和卡片之间、载入文字在卡片下面；“换一个文件”保留分析、给出“回到「名」的分析”并回到原来的模块；没有分析、等选工作表、导入失败时不给入口 |
| `v3.0.0 F1:` / `v3.0.0 F2:` | 2 / 6 | 字段配置表的 7 列和“认不出”的粉点、草稿即时更新；多于 20 个字段时的搜索和类型切换、个数、隐藏行、应用后保留、新导入清空、转义 |
| `v3.0.0 F3:` | 4 | 每个图表框有静态的空状态段落；`drawEmptyChart()` 逐字写文字、不在画布上画，下一张图清掉它；自定义分析的空趋势；空表格和空列表靠容器上的类得到统一外观，句子不变 |
| `v3.0.0 overview:` 等（7a） | 9 | 5 个数字写成一行；数据预览在 `<details open>` 里；字段配置的说明和按钮在表格下方、有草稿时加 `has-draft`；只有多于一句的小节说明有“展开”；频数比例条；导出页各选项标题后的行数；截断单元格的悬停全文和图表标题的提示；单元格限宽、手机上示例小表的两行排法、质量页两张表上下排 |
| `v3.0.0 look:` | 7 | 只读 `style.css`、`app.js` 的 `CHART_THEME` 和 `script.js` 的 Chart.js 默认值：颜色变量、系统字体栈（`Segoe UI` 在微软雅黑之前）、7 档字号和 3 档字重、焦点环；模块不是卡片，内容块是白卡，`box-shadow` 只给侧栏浮层；按钮高度、次按钮和文字按钮、手机上的 44 px、分段导航、表头；质量页标签中性、没有琥珀色和圆点，5 个小节的类别色圆点，“认不出”的粉点，问题格浅灰加粗；一组一个主按钮（只用 id 选择器）；空值检查表约 10 行不滚动、640 以下不限高、打印不限高；图表主色 `#0068e0`、第二色比主色浅且对白底不低于 3:1、图表结论和自定义分析结果句的浅灰框 |
| `v3.0.0:`（查看） | 1 | “查看”把问题行查看区滚到顶端（`scrollIntoView({ block: "start" })`，标题 `focus({ preventScroll: true })`）；换类别、翻页、重新列出都不滚动 |
| `v3.0.0 look2:` | 9 | 顶栏只有名字、页脚只有 GitHub 链接；模块标题 17 px 无下划线、三个模块标题旁一行说明；质量页三张卡 + 一张包 7 个小节的卡；空值检查、偏大或偏小、数值统计三张表和字段配置表的数字列带 `num` 类，报告没有；趋势线直线、白色数据点，每张图 y 轴最多 5 个刻度；x 轴 `autoSkipPadding: 16`；模块里的小标题 15 px / 600；上传图标没有方块、两个上传按钮没有图标 |

**改动的现有断言（2 处，都和导航有关）**：`basic` 的 `v2.7.0 conclusions: the findings section and its navigation entries are gone…` 原来固定导航编号 01–08 和手机下拉框，改成只有一个导航、恰好 8 个链接、依次指向 8 个模块 id、没有编号、没有 `#mobileSectionNav`（旧写法里的 `<span>09</span>` 一支在去掉编号后会永远匹配不到）；`core.test.cjs` 原来断言有 `#mobileSectionNav`，改成断言没有它，并且恰好 8 个指向模块的链接。读样式表的 4 条旧断言（`workflow-`、`.template-`、`:not(.has-note)`、打印块隐藏 `.standardized-export-actions`）都没改，仍然通过。

**第二次样式提交（look2）改动的现有断言**：`core.test.cjs` 原来断言顶栏是“名字 + `<small>` 副标题”、页脚是“名字 · 副标题”，改成切出顶栏和页脚：顶栏只有名字、没有 `<small>`，页脚只有 GitHub 链接、不重复副标题。`basic` 里四类：取 `<td>` / `<th>` 单元格的正则放开 class（数字列加了 `num`）；删掉顶栏隐私徽章和 `.privacy-badge` 样式的断言（改成断言它们不存在）；质量页小节说明的句数从 `[1, 1, 2, 1, 2, 3, 3]` 改成 `[1, 2, 1, 2, 3, 3]`（模块自己的一句说明成了标题旁的短句）；“卡片里没有卡片”去掉 `.problem-rows`（问题行变成了一张卡）。

**外观测试是怎么验证的**：`v3.0.0 look:` 的每一条都先在改样式前的代码上跑过，确认失败，改完再通过。后来补的两句断言（筛选页日期框不超出卡片、自定义分析结果句的样式）也各自确认过：去掉对应的规则就失败。

假 DOM 不算样式和尺寸：模块切换后的可见性、Tab 走不进看不见的模块、隐藏模块里图表的导出尺寸、吸顶表头、侧栏在 1100 / 641–1099 / ≤640 三档的行为、横向滚动、打印，都要在浏览器里看（`docs/RELEASE_CHECKLIST.md` 第 7 节）。

## 14. v3.0.1 新增测试固定了什么

v3.0.1 只换产品的显示名：页面和报告叫“表格体检”，报告标题“表格体检报告”，英文名 Sheet Checkup 只在两份 README。页面 title 是“表格体检 | 检查 CSV、Excel 表格的数据问题”，品牌链接的 aria-label 是“表格体检首页”，两处报告页脚是“由表格体检 …”“> 由[表格体检](…)”（汉字之间不留空格）。另有一行样式：`.brand-copy strong` 的 `white-space: nowrap` 从 640 px 以下的规则挪到基础规则，名字在 642–790 px 不再被拆成两行。判断、计数、顺序、导出文件一样不变。新增 1 条测试（526 → 527）。

| 测试 | 固定的内容 |
| --- | --- |
| `v3.0.1 rename: the old product name and report title never appear; the new name is exact` | `index.html` 去掉注释后的源码（含全部 aria-label）、导入示例后假 DOM 的全部元素、生成自定义分析并筛选到华东后的 HTML / Markdown 报告、22 个导出文件名（v3.1.0 删去去重后的那个，是 21 个）里，旧产品名、旧副标题、旧报告标题都不出现；旧词用 `join` 拼出来，所以测试文件本身不是 grep 旧名的命中；拉丁词不分大小写但要求中间是空格，仓库地址 `smart-tabular-analyzer` 不命中（并断言三份文本都含仓库地址）。逐字固定 title、aria-label、顶栏唯一的 `<strong>`、报告的 title、h1、Markdown 第一行和两处页脚的链接文字和地址（v3.2.0 删去 Markdown 后只剩 HTML 的 title、h1 和页脚，文件名 20 个）；导出文件名以原文件名开头、不含产品名。`.brand-copy strong` 的基础规则有 `nowrap`，640 px 以下的规则只有字号 |

**改动的现有断言（10 行）和报告哈希**：`core.test.cjs` 204、207、211、759、776，`basic` 7365、8428、8432、8436（v3.0.0 时的行号），`data-processing` 342。759、776 和 `data-processing` 342 连同“由”和链接地址一起匹配：“表格体检”是“表格体检报告”的子串，只写新名时页脚改回旧名也能通过。core 211、basic 8436 是页脚的 `doesNotMatch`，旧文字消失后会永远通过，所以补上新名和新 title 的后半句。报告文字的哈希护栏改名为 `v3.0.1 guard: the sample's report text is the v2.12.0 text with only the name changed`；新哈希按等价比较得到，不是拿新代码重算：改前的原文只做三种替换（旧报告标题 → 表格体检报告 6 处，两种页脚“由 + 旧名链接” → “由 + 表格体检链接”各 2 处），和改后原文逐字节相同（31819 → 31743 字符）。

**改坏验证**：title、顶栏名字、两处报告标题和两处页脚改回旧文字，各自有原来的测试失败；品牌链接的 aria-label 改回旧名、去掉基础规则的 `nowrap`，只有新护栏失败（改名前 aria-label 没有任何测试覆盖）；旧名只写在 `index.html` 的注释里，按设计通过。

假 DOM 没有 `document.title`（取 `index.html` 的 `<title>`，脚本从不改它），`setAttribute` 是空函数，脚本动态加的 aria-label 只能从 `innerHTML` 里看到；名字是否一行要在浏览器里看。

## 15. v3.1.0 新增测试固定了什么

v3.1.0 让用户把疑似合计 / 小计行和整行重复里多出来的行设为不计入分析（见 `docs/ARCHITECTURE.md`“不计入分析的行（v3.1.0）”）。新增 40 条测试（527 → 567），都在 `basic-data-processing.test.cjs`，名字以 `v3.1.0` 开头。不点按钮时的结果由现有护栏固定：`v2.8.1-guard.json` 的摘要护栏（哈希没变，见下面“改动的现有断言”）和 `v3.0.1 guard` 报告哈希都照常通过。

| 测试 | 固定的内容 |
| --- | --- |
| `v3.1.0 empty values: …`（5 条）、`已清除现有筛选条件 only follows a rebuild that had a filter applied` | 按空值处理 / 撤销 / 全部按空值处理的消息说具体数字：空值、认不出前后，“这一行没有删”，数值字段“平均值等统计不变”、日期字段“按日期的趋势不变”、分类字段“不再有「—」这一类”、ID 字段说编号重复、类型变了说新类型、被忽略字段说为什么；“1 个”是这个写法覆盖的格数；“已清除现有筛选条件。”只在之前有筛选时出现 |
| `v3.1.0 counted rows: …`（11 条） | 不点按钮时 `getIncludedRows() === state.rows`、两份统计仍是同一个数组；两类都不计入后每个出口（统计、图表、预览、筛选选项、导出、报告）都是 40 行，经过一次重建也是；分析区找不到合计行的 123440 / 123,440；重建取交集（仍被标出的留下、不再被标出的恢复一次并说、又被标出时说一次）；点按钮不重建、保留筛选和自定义分析、有字段草稿时拒绝；导入和重置清空、筛选和“检查会改哪些格”不清；另一份数据的下标不算；导入时 `buildExactRowKey()` 的调用次数和 v3.0.1 相同（性能护栏）；重复那一类和质量页的重复数共用 `findRepeatedRowIndexes()`；有不计入的行时按空值处理的消息数计入的行；数值筛选的输入框跟着计入的行（这个字段已应用筛选时不动） |
| `v3.1.0 quality page: …`（7 条） | 有不计入的行时质量页每个数字都不变，问题行只在说明末尾加“（不计入分析）”（查看区和“导出这一类”）；两个按钮、说明句、各自的消息和焦点；分区小计表的 R1 那句（示例、门店报表没有）；“另有 N 行”只在核对过数值列时出现（自建任务表 95 个标签行、6 行疑似合计）；消息的“N 行”只数这次新加或恢复的行（两张自建表）；两列对上不同范围的合计行不算“只等于上面一行”（every，不是 some） |
| `v3.1.0 page: …`、`v3.1.0 filters: …`（2 条） | 范围句、筛选面板四格、筛选消息、概览、预览页码、侧栏“筛选”一行、自定义分析范围句在排除前、排除后、再筛选三种状态下的原文；“筛除”按计入的行算（消息和面板）；筛选结果来自别的数组时退回计入的 40 行，不退回 42 行 |
| `v3.1.0 status bar: …`、`v3.1.0 empty states: …` | 导入时那句以“可以在“数据质量”里设为不计入分析。”结尾，旧结尾（连“目前仍计入统计，”一起匹配）不在状态栏、质量页和两份报告里；两句空状态不再说“当前筛选” |
| `v3.1.0 modules: …`（3 条） | 三个模块顶上的提示按筛选后的行数（只看华东时不提）、只剩一类时只提那一类、没有干净编号列时不提重复行、aria-label；重复行列表一次导入只建一次（计数，再导入另一份文件再建一次）；“去数据质量处理”切到质量模块、滚到小节、焦点在小节标题（重复行时在那一行的按钮） |
| `v3.1.0 report: …`（5 条） | 数据规模、筛选条件（没有筛选时“结果基于全部 40 行（不含不计入分析的 2 行）”）、质量部分开头那句、疑似合计说明句，HTML 报告里的位置和原文（v3.2.0 前 Markdown 也有），和页面同文；没有不计入的行、旧快照（补 0 / `null`）时一句都不输出，仍写“结果基于完整数据”；行号最多 20 个、后面写“等”、查不到行号时不写括号；一行两类都是时只数一次（“，共 3 行（第 7 行两类都是）”）；R1 和“另有 N 行”两句和页面同文 |
| `v3.1.0 export: …`（2 条） | 导出区只有“当前分析的行”和“写法统一的表”两项；说明句后半句“不含不计入分析的 2 行。”；文件名 `_当前分析_`、状态行“已导出当前分析的 40 行”；去重导出的代码和页面都没了；页面上不再出现“去重后”“导出筛选后 CSV”“筛选后数据”“_filtered” |
| `v3.1.0 wording: …` | 新句子没有 21 个旧词、两个汉字之间没有空格（页面元素、三张表各自排除后、format.js 的函数） |

**改动的现有断言**（行号是改动时的行号）：

- ① 消息：`basic` 2258、2297、2355（按空值处理 / 全部 / 撤销的旧消息“……统计、图表和分析结论已重新计算。 已清除现有筛选条件。”换成新消息）。
- ② 计入分析的行：`data-processing` 726–727（`deduplicateRows()` → `findRepeatedRowIndexes()`）；`basic` 5331（v2.8.1 摘要护栏的 `dedup` 换成 `state.rows.length - state.duplicateRows`：改前的代码上实跑、四个时区都通过，`v2.8.1-guard.json` 没动，30 个状态里两个表达式逐个相等）。
- ③ 质量页和各处文字：`basic` 1482、4830、6821（状态栏结尾）；经用户同意另改 7 处：4834、4836、4838、4845（状态栏结尾的其他写法）、6832（示例疑似合计一节的文字多了按钮“不计入分析”）、7178、7182（两句空状态）。
- ⑤ 报告：经用户同意改 `v3.1.0 counted rows: every exit counts 40 rows…` 里 `report.dataScale` 的预期（加 `excludedRows: 2`）；修复提交里两条 v3.1.0 报告测试的 `exclusions` 预期加 `overlap: null`。
- ⑥ 导出：`basic` 1199–1232（去重的行数、计数、按钮、文件名删去；CSV 内容改用当前分析的行，另加不筛选时 3 行完整内容）、2525、3999（护栏只留当前分析的行，加有不计入的行时 40 行、经共用函数逐字节等于旧写法）、6389、6539（导出选项 3 → 2）、8202–8203、8370、8376、8664（改名护栏的文件名 22 → 21 个，加数量断言）。

**改坏验证**：每个提交至少两处，结果在实施汇总里。值得记下的几处：几处改回 v3.0.1 的写法（`getAnalysisRows()` 退回 `state.rows`、筛除按全部行算、筛选选项和数值范围按全部行算、有不计入的行时重建仍共用两份统计）各有测试失败；R1 的 every 改成 some、消息按这一类全部行数（两种改法）只有 v3.1.0 补的那几条测试抓得到；导入时就建重复行列表，调用次数护栏失败；报告“不计入分析行数”为 0 也输出，报告哈希护栏失败。

假 DOM 测不到：按钮和“查看”一样高、留 8 px，手机上 44 px；提示在标题区里、范围句下面；“去数据质量处理”滚到顶栏下面；有不计入的行时 HTML 报告的图表张数。这些要在浏览器里看（`docs/RELEASE_CHECKLIST.md` 第 5、7、8 节）。

## 16. v3.2.0 新增测试固定了什么

v3.2.0 加了英文界面（见 `docs/ARCHITECTURE.md`“界面语言（v3.2.0）”），删去了 Markdown 报告。完整回归从 567 条变成 636 条：按测试名数，新增 77 条、删去 8 条（`data-processing` 里 3 条只测 Markdown 的测试；另有 5 条去掉了名字里的 “Markdown”，算一删一增）。新测试名都以 `v3.2.0` 开头，按提交分组；中文界面的输出由 `scratchpad` 里的中文基线脚本逐提交比较（不进仓库），允许的不同只有删 Markdown、超过 80 个字符的文件名、两个切换按钮和 `data-i18n` 属性、`<head>` 里定语言的内联脚本、导航放不下时换到第二排、统计 / 图表 / 自定义分析顶上的提示条和导出区“共 N 行”。

| 测试 | 固定的内容 |
| --- | --- |
| `report filter descriptions are escaped in HTML` 等 4 条、`${method} custom analysis … exports median in the report` | 删 Markdown 时改名，只剩 HTML 那一半；另在原来只有 Markdown 断言的地方补了 15 处 HTML 断言（先在 v3.1.0 的代码上跑通） |
| `v3.2.0 tests read index.html only through indexHtmlSource()…` | 测试读 `index.html` 只经过 `indexHtmlSource()`（去掉 `data-i18n` 开头的属性，别的不动）；去掉属性后的 sha256 和改前相同（v3.2.0 外观提交有意改过一次） |
| `v3.2.0 long file names: …`（3 条）、`v3.2.0 report.js: long English names keep their label words (R1)` | 长原文件名下各个导出不同名，类别词、前N条 / first-N、日期不截；原来不超过 80 个字符的文件名逐字节不变（含 80、79 字符、标签前有空格的边界） |
| `v3.2.0 static guard: …`（5 条） | 分词器自检（9 个脚本的汉字全在注释或字面量里）、白名单项都存在且字面量个数恰好相等、规则 1 和 2 对全部脚本生效、审查的 13 个自测用例、`index.html` 每段中文有键且每个键有英文；产品名的三处（`<title>`、品牌链接、顶栏名字）用 `page.title`、`brand.label`、`brand.name` |
| `v3.2.0 language: …`（10 条） | `?lang=` 优先（只认小写）、浏览器第一语言、都没有时中文；`<head>` 内联脚本和 `pickInterfaceLanguage()` 一致、不定义全局名字；CSS 3 秒兜底；测试上下文可选英文并走真实启动；`L` / `Lp` / `Lpick` / `enCount`；静态文字切换和逐字节切回；启动时读下页面原文；每个键的英文保留数字；切换按钮的文字、`lang`、`?lang=`、模块保留、导入中等待 |
| `v3.2.0 format.js: …`（5 条）、`v3.2.0 format.js English wording: …` | 18 个成对常量冻结、英文无汉字；每句英文保留中文的数字；逐句钉住的英文；第二档空值写法的说明只列这一列找到的写法（R11；v3.3.0 起中文也分两种句子，见第 19 节）；逐行文字不调 `L()` |
| `v3.2.0 app.js: …`（7 条）、`v3.2.0 app.js English wording: …` | 英文界面不出现脚本里的中文、每个元素数字和中文相同；筛选顺序两种语言相同（R3）；导入消息存数据；一次性消息切换后重写、切回逐字节；重建消息点击时两种语言都生成；“（不计入分析）”后缀每张列表取一次；趋势数字留给切换用、随分析一起清掉 |
| `v3.2.0 profiling.js / stats.js / importer.js / quality.js / script.js: …`（8 条，含 `v3.2.0 quality.js English wording: …`） | 类型名成对、`profile.type` 仍是中文；空值组键是中文、显示 “(Blank)”；`#错误` 两种语言相同；读取疑问存代码和参数；趋势结论拆成算和拼句、切换时不重算；问题行说明的英文；自检提示和示例消息的英文 |
| `v3.2.0 report.js: …`（8 条，含 R10；R1 那条见上）、`v3.2.0 report.js English wording: …` | 英文报告 `<html lang="en">`、没有汉字、数字和中文报告相同；英文文件名用计划里的词；导出加的 CSV 列跟语言；旧快照的句子原样；导出状态行切换后重写；R10（报告文件名的“已载入”回退） |
| `v3.2.0 navigation: …` | 导航放不下时才换第二排，量的时候先去掉那一排；切换语言后重新判断 |
| `v3.2.0 commit 10: …`（2 条） | 三个模块顶上的提示条（白卡片、句子在左、按钮在右）；导出区“共 42 行 / 42 rows”，`<strong>` 里只有数字 |
| `v3.2.0 commit 11: …`（6 条） | 中→英→中整页逐字节回到原样并保留状态（测试里模拟浏览器 `<select>` 改写选项后丢掉选择）；切到英文和一开始就是英文完全相同；只重画不重算；问题行页码保留；状态栏、导入报错、工作表消息切换后重写；`role="status"` 确认句 |
| `v3.2.0 fix: English chart axis titles are capitalized…` | 英文纵轴标题首字母大写，句子里的写法仍是小写，中文不变 |
| `v3.2.0 commit 12: …`（8 条） | 英文示例的字节（去掉 CR 后 sha256 `SAMPLE_ORDERS_EN_LF_SHA256`）、合计行连重复行一起算；按钮和下载链接跟语言；上传页英文那几行等于英文示例的行；干净版零提示、每类问题单独放进去只出自己的提示（`$` 金额不提示，写明原因；v3.3.0 起示例改成 USD，七类都提示，见第 19 节）；英文界面下的状态句、pill、各节和行号、两句结论；两个英文夹具对 `known-issues.json`；生成器逐字节 |
| `v3.2.0 commit 13: each language shows only its own product name…` | 中文界面只有“表格体检”、英文界面只有 “Sheet Checkup”，旧名两种都没有（页面元素、报告 title / h1 / 页脚、全部文件名、英文表） |
| `v3.2.0 commit 13: in English, no output of any operation…` | 英文界面导入英文示例、两个英文夹具、⑪ 的英文场景表和测试里构造的两个文件（只有 1 列的 CSV：没有行号的读取疑问；工作簿：错误码读不出的错误格和 Excel 日期单元格），走导入、报告、自定义分析、每一类问题行（4 遍）、筛选、导出、不计入分析、按空值处理和撤销、字段草稿 / 应用 / 恢复、写法统一和两个导出、日期趋势、清除筛选；约 4.8 万份输出（元素文字和属性、图表标题、下载的文件名和全文）没有汉字和全角标点。允许的只有：两个切换按钮的“中文”；`FORMULA_ERROR_FALLBACK` “#错误”（单元格数据）。`BLANK_GROUP_KEY` 实测 0 次，钉住 0 次、不开例外。钉住 11 类问题行都打开过，以及每类由哪个文件触发 |
| `v3.2.0 commit 13: the same file in Chinese and in English…` | stage3 全部、两个示例和英文场景表（括号开头的值和空值同列），中英各导入一次，在导入后、自定义分析后、全部按空值处理加写法统一检查后比较：state（计数、profile、统计、问题行和行号、筛选顺序、写法统一结果、分组结果）逐字相等，报告快照去掉字符串后相等，报告和每个元素的数字序列相等；报告的分析时间两边固定成同一时刻 |

**改动的现有断言**（每个提交的完整清单和改坏结果在实施汇总里）：

- ① 删 Markdown：`reportOf()` 去掉 `markdown` 键；各测试里 Markdown 那一半删掉，拼接 `html + markdown` 的地方只留 HTML；禁用词、名字、状态栏护栏和 core 的旧名扫描去掉 Markdown 那一路；名字护栏的文件名 21 → 20；报告哈希换成只算 HTML 的 `e32ff801…`（在 v3.1.0 的代码上只取 HTML 两段算出，另写改法复算一致）。
- ⑤ `format.js` 的 18 个常量成对：测试读常量的地方改成 `[0]`；其中 `!includes(常量)` 这类不改会永远通过（数组转成字符串后中英两句连在一起），一起改了。
- ⑦ 读取疑问改存代码和参数：警告表和 `codedWarnings()` 经 `formatParseWarning()` 比较，期望不变；core 的自检提示加英文运行。
- ⑨ 静态护栏两条改名（规则 1 对全部脚本生效、`index.html` 全页检查）。
- ⑩ 外观：`v3.0.0 look2` 的位置断言改成“标题 → 提示条 → 说明段”；`index.html` 去掉属性后的 sha256 换成 `7c3f1ed6…`（注释写明有意改动）；导出区两项改成 span + strong + span 的结构。
- ⑪ `the trend figures are kept for a switch…` 从“不读行”改成“趋势数字不重算”（切换要按行重画概览）；⑧ 的报告测试把分析时间固定（原来偶尔差一秒失败）。
- ⑫ 按目录枚举的断言：示例 7 → 8 个、`scopes.length` 15 → 18、疑似合计 / 公式错误 / 日期写法 / 常量列 / 写法不一样各清单加上新文件、写法统一的格数加三项、疑似空值写法加英文示例的 Amount N/A 和 Sales rep —；v2.8.1 摘要护栏照 v2.11.0 的做法跳过三个新文件（`v2.8.1-guard.json` 没动）；`fetch` 的断言改钉 `` fetch(`examples/${sampleName}`) `` 和 `const sampleName = L("sample-orders.csv", "sample-orders-en.csv")`（basic 和 core 各一处）。

**测不到、要在浏览器里看的**：英文浏览器打开不带参数的地址是英文、没有中文闪一下；`?lang=en` 下删掉 `script.js` 或让自检失败，页面 3 秒内出现；导航在 642–1366 两种语言都没有被裁的项（640 以下是手机布局，导航是可以横向滑动的一条，后几项在屏外（v3.0.0 起如此，这一版不改））；切换前后图表一样；打印英文报告（`docs/RELEASE_CHECKLIST.md` 第 7、8 节）。

## 17. v3.2.1 新增测试固定了什么

v3.2.1 让多工作表的 Excel 不用重新上传就能换工作表（见 `docs/ARCHITECTURE.md`“换工作表（v3.2.1）”）。完整回归从 636 条变成 669 条，新增 33 条，没有删去。新测试都在 `basic-data-processing.test.cjs` 末尾，名字以 `v3.2.1 commit N` 开头，按提交分组。英文多表夹具 `tests/fixtures/en-multi-sheet.xlsx`（3 张表：Orders、Customers、Notes，Notes 只有表头）由同目录的 `gen-en-multi-sheet.cjs` 用 `vendor/` 的 SheetJS 生成，两次生成逐字节相同，测试钉住 sha256 并重新生成比较。它不放 `stage3/`，因为那个目录会被按目录枚举的护栏和计数断言用上，而多表文件在那些地方要先选表。

| 测试 | 固定的内容 |
| --- | --- |
| `commit 1: a multi-sheet file reads its names first…` | `XLSX.read` 的调用次数和选项（给它包一层计数）：多表出选表面板时 1 次 `{ type: "array", bookSheets: true }`，选表后第 2 次带 `sheets: 表名`；单表 2 次；CSV 0 次。`pendingExcel` 只有 `bytes / fileName / sheetNames / importId`，`bytes` 是 ArrayBuffer |
| `commit 1: reading one sheet gives exactly what reading the whole workbook gives…` | stage3 的 4 个 xlsx、`excel-import-test.xlsx` 的 4 张、7 张自造的表（含空表、只有表头、没有表头、表头重复、超列、有标题行），中英文：`readExcelSheet()` 和“整本读 + 同样的步骤”结果或报错原文逐字相同。钉的是 SheetJS 的行为，升级 SheetJS 时有用 |
| `commit 1: the chosen sheet says it is being read…`、`commit 1b`（2 条） | 选表后先出读表提示、加载中、语言按钮禁用；读失败时报错留在面板、焦点回“分析所选工作表”、报错切换语言能重写；选好表后 `#sheetSelectionMessage` 清空（R10）。让出那一帧期间开始了别的导入（失败的表、能读的表各一条）时，选表这一路什么都不做 |
| `commit 1 / commit 2: an unexpected error while … is committed…`（2 条，R11） | 让 `countDuplicateRows()` 抛错：Promise 不被拒绝，不再加载中，文件输入和语言按钮可用，状态栏是报错 |
| `commit 2`（其余 7 条） | 换表的结果和重新上传选那张逐项相同（报告、状态栏、侧栏、五种文件名、一次性消息；先设好筛选、两类不计入分析、按空值处理、字段类型、自定义分析、写法统一检查、问题行，否则分不出）；这些都被清空、打开第一个模块（v3.4.0 起叫“体检结果”）；读失败时当前分析不变、报错写在列表下面并能切换语言；只有多表文件留字节，字节在换表和换失败后都是同一个对象，任何一次新导入都清掉；读表时加载中、语言锁住、只读那一张；让出那一帧期间开始的导入优先；重置清掉两个列表和提示 |
| `commit 3`（7 条） | 换表失败后焦点回到点的那一项（侧栏和上传页的列表都测；假 DOM 临时补 `getAttribute` / `querySelector` / `focus`，`withSheetListDom()`，`finally` 里还原）；多表时 `#sourceName` 只写文件名、下一行写工作表，单表照旧；列表标出当前那张、其余是按钮、名字转义、点了按序号换；按钮开合、Escape 收起并 `stopPropagation`、切换语言后仍展开；英文夹具和生成器；英文文件在英文界面里列表没有汉字、中→英→中逐字节回到原样；侧栏标记（说明句在列表上面，三个 `I18N_EN` 键，640 以下列表不限高） |
| `commit 4`（2 条，R8） | 上传页“回到「…」的分析”切换语言后重写、切回逐字节相同；回到分析或新导入后清空，不会被切换写回 |
| `commit 5`（1 条，R9） | 样式规则：`.upload-hero .return-to-analysis` 有 `max-width: 100%`、`white-space: normal`、`overflow-wrap: anywhere`，没有字号字重；按钮共用规则仍是 `nowrap`。必须是两个类：只用一个类时会被后面的 `.secondary-button` 盖掉 |
| `commit 6`（4 条） | 上传页的按钮只跟着“回到”出现：多表时有，单表、CSV、选表中、导入中都没有；回到分析后隐藏并收起（T14）。上传页换表失败时留在上传页、“回到”还在，回去是原来那张表，报告 sha256 和换表前相同，报错清空。从上传页换表成功和重新导入那张相同。上传页的标记去掉缩进后和侧栏的相同；`text-align: left` |
| `commit 7`（5 条，R2、R2b） | 多表时中英文 7 个文件名逐字写死；真实导入和换表后的名字；70 字原名 + 30 字表名的截断逐字写死，日期前不超过 80；表名先 trim，全是空格时和单表相同；CSV 和单表 Excel 用 21 个刁钻原名 × 3 种扩展名 × 中英文，报告和当前分析的行逐个和 7b6e091 的公式比，882 个名字整体的 sha256 是在 `git archive 7b6e091` 上用同样的调用算出的 |

**改动的现有断言**（每个提交的完整清单和改坏结果在实施汇总里）：

- ① 读一张表：`readExcelWorkbook()` 删去，测试里经过它的 4 处改用 `readExcelSheet()`（`importXlsxSheet()` 助手、偏移表的两处、`importExcelSheet()` 的一处），断言不变。选表那条测试多断言一句读表提示。`destructive-qa` 里 `showExcelSheetSelection()` 改传 ArrayBuffer，并加断言 `bytes`、`sheetNames`、`importId`，不再有 `workbook`（收紧）。
- ②、③、⑥：`index.html` 有意改动，整页 sha256 依次换成 `71260d2e…`、`7331133e…`、`95ef087e…`，每次在汇总里贴了 diff。
- 其余提交没有改现有断言。

**测不到、要在浏览器里看的**：列表在 1366 / 1100 / 768 / 640 / 390 宽下中英文的样子，包括长表名折行、30 张表时列表自己滚动、640 以下和“换一个文件”并排、箭头看得见；读表提示真的先画出来；1100 以下的侧栏换表成功后关上、焦点在第一个模块（v3.4.0 起“体检结果”）的标题；Tab 和 Escape 的实际焦点；上传页长名时没有横向滚动；改过磁盘上的文件后仍能换表；读表时的内存（`docs/RELEASE_CHECKLIST.md` 第 4、7 节）。

## 18. v3.2.2 新增测试固定了什么

v3.2.2 只把仓库地址换成 sheet-checkup：页面页脚的 GitHub 链接、中英文报告页脚的链接和 `package.json` 的 name。判断、计数、界面文字、导出文件一样不变，报告除页脚链接地址外逐字节不变。完整回归从 669 条变成 670 条，新增 1 条，没有删去。

| 测试 | 固定的内容 |
| --- | --- |
| `v3.2.2 repository: the page and both reports link to sheet-checkup only` | `index.html` 含注释的源码（经 `indexHtmlSource()`，只去掉 `data-i18n` 属性）和中英文报告（`reportScenario()` 的场景）里都没有旧仓库名（不分大小写）；页面页脚 `footer-github-link` 的 href 和两份报告页脚里唯一的链接都等于 `https://github.com/wy-data-30/sheet-checkup`；`package.json` 的 name 是 `sheet-checkup`。旧仓库名用 `join` 拼出来，所以测试文件本身不是 grep 旧名的命中。查的是含注释的源码：旧名只写在 `index.html` 的注释里也会失败（v3.0.1 的改名护栏查去掉注释的源码，按设计不管注释） |

**两个哈希按等价比较换值，不是拿新代码重算**（照 v3.0.1 的做法）：在 `git archive` 出的改前、改后两份副本里，把哈希断言换成“把被哈希的文本写进文件”，只跑那一条测试，导出两份原文；改前原文只把 `https://github.com/wy-data-30/smart-tabular-analyzer` 换成 `https://github.com/wy-data-30/sheet-checkup`（只做这一种替换），和改后原文逐字节相同。

- 报告文字 `GUARD_REPORT_SHA256`（`v3.0.1 guard: the sample's report text…`）：替换 2 处，23752 → 23734 字符，`e32ff801…21421` → `937a0034…82f5f2`。
- `index.html` 去掉属性和版本号后的 sha256（`v3.2.0 tests read index.html only through indexHtmlSource()…`）：替换 1 处，35023 → 35014 字符，`95ef087e…ca866` → `8c1f88e7…bdc80a`。

两处上方的注释写明了这一点。

**改动的现有断言（7 行，2f49ec8 上的行号）**：`core.test.cjs` 239（页面页脚）、788、789（中文报告页脚和地址）；`basic` 8375（`v3.0.0 look2` 页脚正则）、8550（`RENAME_REPO_URL`，v3.0.1 改名护栏断言三份文本都含它）、10703（英文报告页脚）：正则或字符串里只把仓库名换成 sheet-checkup，其余字符不变。另改 `basic` 8547 的注释：拉丁旧词要求中间是空格，所以小写连字符的仓库名本来就不会命中，旧仓库名由 v3.2.2 的护栏单独查。加上两个哈希，没有放宽或删去断言。

**改坏验证**：页面页脚、中文报告页脚、英文报告页脚分别改回旧地址，各有原来的测试和新护栏失败；`package.json` 的 name 改回旧名、旧仓库名只写进 `index.html` 的一段注释，前者只有新护栏失败，后者是新护栏和 `index.html` 哈希失败。

**测不到、要在线上看的**：GitHub 上改名之后旧的 github.io 地址（含 `?lang=en` 和示例 CSV 的深链接）和 github.com 旧仓库地址能跳到新地址；这些不在仓库的测试里。

## 19. v3.3.0 新增测试固定了什么

v3.3.0 加了英文表格的读法（见 `docs/ARCHITECTURE.md`“英文表格的读法（v3.3.0）”）：英文合计字样和“同一字样过半”、None / Unknown 等空值写法、编号字段名、英文单位和货币代码、货币符号后的会计写法、AM/PM 和序数日、两位年份。完整回归从 670 条变成 709 条，新增 39 条（`data-processing` 22 条、`basic-data-processing` 17 条），没有删去。新测试名字都以 `v3.3.0` 开头；每条都先在改前的代码上跑过、确认失败（护栏性质的几条改前就通过，下表写明）。

| 测试 | 固定的内容 |
| --- | --- |
| `v3.3.0 date matchers build no RegExp per call…` | 用 `Proxy` 给全局 `RegExp` 计 `construct` 次数：三个日期匹配器各调 1 万次构造 0 次（改前 30024 次）；8 个样本的匹配结果逐组钉死 |
| `v3.3.0 totalRowLabel(): an English prefix…` | 16 个正例（East Total、East - Total、East – Total、East:Total、East Total (USD)、EAST TOTAL、恰好 40 个字符的前缀、Totals、TOTALS: 等）的返回值；反例 Total East、Total for East、Subtotal: East、Overall、Summary、Teetotal、41 个字符的前缀、Total Units、East Totality、已汇总 等 |
| `v3.3.0 totals: …`（11 条） | Excel 分类汇总三种写法全部标出（有没有 Grand Total 都测）；七张“字样是值”的普通表一行不标（指标长表、工时表、品牌叫 Total、KPI、订单类型、中文任务名、中文状态）；品牌表 + 真 Grand Total 只标真合计，`suspectedTotalLabelCount` 不数被去掉的组，再加一行单独的 Promo Subtotal 时页面和报告写“另有 1 行”；组按“列 + 字样”分；恰好一半对得上的组留下，不到一半才去；中文表结果不变；最多去 3 轮、`checkTotalRows()` 最多 8 次（链式表）；`findTotalLabelRows()` 的 `key` 是小写，`state.suspectedTotals` 的每一项不加键、重建不重新扫格；组不再对上时已不计入分析的行恢复计入、再被标出时提醒；v3.1.0 的任务表三组字样一行不标；中文句子里英文字段名和“等于”之间有空格 |
| `v3.3.0 placeholders: …`（5 条 data-processing、2 条 basic） | 新的第二档写法（含全角形式）和清单外的近似写法；单元格长度上限 14、纯 ASCII 不做 NFKC；只有原来五个第二档写法时中英文说明句逐字节不变（31 个子集）；有新写法时只列这一列的写法、每种给例子；同一写法只列一次；None、Unknown 逐个列出，“常见写法全部按空值处理”不含它们，原值不变；报告的全表一句里同一写法只列一次 |
| `v3.3.0 ID names: …`（2 条） | 同一列 0–9 的计数换 21 个字段名：number / no 只在最后一个词时判 ID；字段名有单独的 yes 时结尾的 no 不算 |
| `v3.3.0 units: …`（5 条） | 每个英文拼写读成规范名、大小写不分、只有英文单位能带点；前置货币代码和读不出的写法（`1.2K`、`72°F`、`USD -`、`6 in`…）；两种单位整列不读、`$` 和 USD 同列不冲突；写法统一的表把英文单位写成纯数字，科学计数法只看数字部分（R4）；非数值字段只说单位冲突（R2） |
| `v3.3.0 accounting: …`（3 条） | `$-`、`€-`、`￥-` 读成 0，`$(1,200.00)` 读成 −1200，单独的 `-` 不是数；写法统一的表写成 `0`、`-1200.00` 并读回同一个数；CSV 导出不给它们加 `'`（护栏性质，改前就通过） |
| `v3.3.0 dates: …`（3 条） | 12 小时制只跟在空格后面、小时 1–12，读出日期、写法统一的表保持原值；序数日后缀要对（11th、21st、22nd、23rd 对，1th、11st 不对）；两者都不是新写法 |
| `v3.3.0 two-digit years: …`（3 条 data-processing、1 条 basic） | 只收斜杠和 `d-mmm-yy`，00–29 / 30–99 的分界，三种日月顺序下的读法；点号、横杠、`Mar 5, 25` 不收；两位年份是单独一种写法 `twoDigitYear`，写法统一的表写成四位；一列里两位、四位混用时质量页中英文写出两种写法 |
| `v3.3.0 English sample: three amounts in USD…` | 英文示例 LF 3306 字节，三格是 `"5,400 USD"`、`900 USD`、`"1,050 USD"`，文件里不再有 `$` 加数字；英文界面 Amount 是数值、单位 USD，“按单位读取的数字”可见 |
| `v3.3.0 fixture: en-subtotal.csv…` | BOM 和 21 个 CRLF；中英文各导入一次：行数、5 列类型、5 条疑似合计（4 个小计 `segment`，Grand Total 是 `above`、跨 15 行、不连续）等于 `known-issues.json`；Report 列的 Weekly Total 不标 |

**改动的现有断言**（每个提交的原文、改后和改坏结果在实施汇总里）：

- ② `v3.1.0 quality page: 另有 N 行…` 和 `v3.1.0 report: …` 原来用任务表（95 行字样、标 6 行、“另有 89 行”），新规则下这张表一行都不标（另写新测试钉住）；改用 `singleLabelTotalsCsv()`（5 行字样、标 2 行、“另有 3 行”“其中 1 行”）。新表在改前的代码上同样通过，数字都是精确相等，测的仍是句子怎么出现。
- ②b `v2.8.0 page: the sentences for row ranges…`：英文字段名后的“等于”前多一个空格（`a、b、c、d、e 等于`），“等 6 列等于”不变。
- ③ `data-processing` 契约 17 前提的护栏 `formCount` 57 → 84（清单变长）；两档清单的断言照新清单，另加 `placeholderTier("None") === "maybe"`、`placeholderTier("na") === ""`。
- ④ `ID names cover row-number headers…`：原来钉着 `Number of employees` 判 ID（注释写“v2.1.4 前就如此，有意保留”），改成不判；正例多 8 个、反例多 5 个，没有删名字。
- ⑤ 契约 17 前提的护栏的策略列表加英文单位（加一行），`formCount 84`、`converted: []`、`matched: []` 原样通过。
- ⑧ `dateStyleOf() is null exactly when toDate() is null` 的写法键清单加 `twoDigitYear`；成对常量计数 `DATE_STYLE_NAMES` 10 → 11；英文写法名表加一项。
- ⑨ 按目录枚举的断言：stage3 + 示例 18 → 19（两处）、中英比较的文件 13 → 14；v2.8.1 摘要护栏的跳过名单加 `en-subtotal.csv`；有疑似合计的文件清单和写法统一改动数表加上它；`SAMPLE_ORDERS_EN_LF_SHA256` 换成 `900a6751…59f78`（字节数 3306 不变）；英文示例金额的读法和“无问题版”的替换正则认 ` USD`；金额那一类的提示从 `[]` 变成 `["numberConversionSection"]`，英文界面显示的小节加这一节并钉住句子。
- ①、⑥、⑦ 没有改现有断言。两个结果护栏（v2.8.1 摘要、报告文字哈希 `GUARD_REPORT_SHA256`）每个提交都原样通过，没有重新生成；静态护栏只把白名单里的 `matchChineseDate` 改成 `CHINESE_DATE_PATTERN`（个数仍是 1）。

**等价比较（`equivalence.cjs`，不进仓库）**：“不含新写法的表结果不变”靠它检查。脚本在仓库外，和 `harness.cjs`、`corpus.txt` 放在一起，原文在维护记录里。用法：

```text
node equivalence.cjs <改前副本> <改后副本> <输出.json> [--allow "<表>|<键>" …] [--only corpus|fixtures]
```

- 两个副本是 `git archive` 或 `git worktree` 出的两份完整仓库（只读）；`harness.cjs` 用各自的 `tests/test-context.cjs` 和 `vendor/` 建上下文，走真实 PapaParse 和 SheetJS 导入。
- 比较的表：corpus 120 张（`corpus.txt` 列的 96 张公开数据 CSV，加 Power BI 示例工作簿的 24 张工作表；数据仓库和提交号见审查报告附录 C）和改前副本 `tests/fixtures`、`examples` 里的 26 张表（改后副本新加的夹具不参加）。每张表中英文各导入一次。
- 比较的项：摘要（类型、自动识别的类型、认不出的数、单位、标签行、疑似合计、空值写法、编号重复、像是数字的文本列、整行重复、疑似非唯一 ID、常量列、偏大或偏小、写法不一样、日期写法、读取疑问数、`buildNumberConversionSummary()` 的结果）、导入后的状态栏，以及中英文 HTML 报告（分析和导出时间固定成 2026-10-25T00:00:00Z，版本号换成同一个占位）。
- `--allow "<表>|<键>"` 列出这个提交允许的不同（表名是相对数据目录或仓库的路径，工作表写在 `#` 后面）；其余任何不同都记 UNEXPECTED，退出码 1。每个提交和上一个提交比一次、和 v3.2.2（`3594fba`）比一次。
- 它分得清对错：v3.2.2 对审查原型跑 `--only corpus`，35 处不同正好落在审查时的 7 张表上。

**测不到、要在浏览器里看的**：质量页疑似合计、像是空值的写法、按单位读取的数字、日期写法混用几节在 1366 / 390 宽下中英文的样子（各提交的截图在实施汇总里）；Edge 里导入、筛选、清除筛选、改字段类型的耗时和 v3.2.2 比（发版提交量）。

## 20. v3.4.0 新增测试固定了什么

v3.4.0 把第一个模块换成“体检结果”（见 `docs/ARCHITECTURE.md`“体检结果（v3.4.0）”），并改了写法不一样的数法（主写法不算）。完整回归从 709 条变成 749 条，新增 40 条，都在 `basic-data-processing.test.cjs`，名字以 `v3.4.0` 开头，可以用 `--test-name-pattern="v3.4.0"` 单独运行。每条都先在上一个提交上跑过、确认失败；护栏性质的几条（`v3.4.0 words`、读缓存的几条）改前就通过，下表写明。**质量页跟随不计入分析（G）留到 v3.4.1**，审查时为 G 写的那 6 条现有测试（报告第 1 节 13、14、15、17、18、19）这一版原样通过、没有改。

新夹具：`tests/fixtures/checkup/messy.csv`、`messy-en.csv`（298 行的“乱表”，画板生成器的输出原样复制；`.gitattributes` 加 `tests/fixtures/checkup/** -text`，测试钉两个文件的字节 sha256）。不放 stage3，所以不进按目录枚举的护栏。

| 测试（按名字开头） | 条数 | 固定的内容 |
| --- | --- | --- |
| `v3.4.0 look-alike:` | 2 | 主写法不算、并列整组算、只有不可见字符的格算、首尾空格的主写法格也算，列表 = 计数；一组里最多的是全角写法时它就是主写法（只算 1 格） |
| `v3.4.0 checkup fixtures:` / `checkup:` / `checkup reads caches only` | 9 | 示例中文 N = 12（合计 1、重复 2、空值写法 2、日期 5、写法不一样 2），示例英文 14，乱表中文 73 / 多类 7，乱表英文 75 / 多类 7，每类的行和 K 逐个精确；同一格只算一次；整行重复含保留行；排除合计 + 重复后 N = 9；比例条各段相加 = 计入的行；卡片顺序；体检计数调 `forEachDateStyle()`、`buildExactRowKey()`、`totalRowLabel()` 0 次；只在提交、重建、不计入分析时算，筛选不算 |
| `v3.4.0 status fold:` / `first screen:` / `bar` | 5 | 载入说明默认收起、警告也收起，`#statusPanel` 读屏不变；大数字、比例条（无 `tabindex`、段的 `title` / `aria-label`）、图例；导航和标题是“体检结果”，页面不再有“数据概览”“数据规模与预览”；范围那一行只在有筛选或有不计入的行时显示；预览默认收起，打印时展开；“导出体检报告”就是 HTML 报告 |
| `v3.4.0 cards:` | 6 | 8 类卡片的标题、行数、小样、说明句、“其中 K 行”、操作（中英文逐字）；R11 日期卡的主写法个数；R12 “查看”的字样和打开的类别；整行重复、按空值处理、写法统一的操作调现有函数；卡片只算一次，筛选和换语言只拼句子 |
| `v3.4.0 also noticed:` / `quality page:` / `checked (R21):` / `also noticed and checked:` | 10 | 其他发现的顺序和句子（7 张表中英文、单位冲突）；R7 的标签行（乱表 5 行，v3.3.0 那 7 张“字样是值”的表 0 条，质量页小节 0 行对上时也显示）；C3 的半句；每条跳到质量页的小节；R8 D 那句；干净表的查过的项目和两个按钮；不计入分析和撤销时其他发现一列都不扫、改一列类型只扫那一列，结果和不用缓存逐字相同 |
| `v3.4.0 perf:` | 2 | 卡片数据读格 0 次、不调 `listProblemRows()`（不计入分析、撤销时也是）；全部 stage3 夹具、示例、乱表和两张内联表上，卡片的“查看”行数和小样、写法、失败行等于逐格读、排序的老算法，排除前后各一次 |
| `v3.4.0 handled:` / `toast:` / `R9:` | 3 | N 12 → 11 → 9 再撤销回 12，按空值处理 12 → 10 再撤销；提示条和已处理的中英文原句、撤销按钮、aria-label；切换语言重写、切回逐字节相同；焦点不动；切到别的模块或新导入时提示条消失、已处理留着；有字段草稿时首屏的三句拒绝写在提示条里，质量页自己的调用写在原处；写法统一的门禁不通过时焦点在门禁说明上 |
| `v3.4.0 sidebar dots:` | 1 | 每类一个点、按卡片顺序、颜色同卡片；整行重复只在第一个字段（花名册有三个 ID 列）；疑似合计和读取疑问不画点；不计入分析、按空值处理后点跟着变；干净表没有点的标记 |
| `v3.4.0 words:` / `phone:` | 2 | 11 种状态下首屏上体检写的全部句子：英文没有汉字、category、empty、缩写，中文汉字之间不留空格、英文字母或数字后面加空格（护栏，改前就通过，改坏 5 种措辞都能抓到）；手机布局：大数字和字一行、按钮在卡片底部、“需要处理”的说明换行，宽屏“其他发现”的标签不折行 |

**改动的现有断言**（每条的原文、改后和改坏结果在实施汇总里）：

- ① 写法不一样的新数法：`known-issues.json` 里“写法不一致.csv”三项和 `en-issues.csv` 一项 `rowCount`（3 / 18 / 10 → 3 / 6 / 2、8 → 2，手改、一行一条）；写法不一致区域 18 → 6 行、客户 10 → 2 行的问题行和说明列（带“主写法是”），查看区 18 → 6 行；英文示例 Region 11 → 3、摘要加 “(most common)”；`rows > 20` 收紧成 `rows === 20`。另加的断言：每组主写法、主写法的行不在列表里。
- ③ 首屏：状态栏两条（默认收起、警告也收起，测试名注明 v3.4.0）；“orders.csv loaded, 1 note” → “2 notes”；预览 `<details>` 不带 `open`；字号白名单加 52px（只出现一次）；整页哈希；第一个标题的英文 “Size and preview” → “Checkup”。
- ④ 卡片：`v3.0.0 look2` 里单独列出 `.checkup-caption h3`，断言 17 / 600，其余 15 / 600 不动；整页哈希；③ 自己那条干净表的写法。
- ⑤ 其他发现：整页哈希。
- ⑥ 卡片操作：2555 那条兜底测试的跳过清单加 `checkupHandled`，另断言它里面「N/A」「—」各出现一次；整页哈希；③ 那条“导出体检报告”的绑定改成 `exportCheckupReport`（严格程度不变）。
- ⑦ 侧栏：`sidebarItems()` 多取一列点的标签，示例中文的期望逐个写明。
- ⑧ 手机：字号白名单加 40px（用户选的，只在手机块里出现一次，行高 44 px）。
- ⑨ 文档：v2.12.0 旧词护栏的清单加“数据概览”（R18）。
- 整页哈希每次都按惯例换：旧原文只做写明的插入或替换 = 新原文，逐字节比较后再换，不拿新代码重算。

三个结果护栏（v2.8.1 摘要、`GUARD_REPORT_SHA256`、`SAMPLE_ORDERS_EN_LF_SHA256`）每个提交都原样通过、没有换值。

**LF 副本**（④ 的 CI 失败之后加的规矩）：工作区是 CRLF（`core.autocrlf=true`），切 CSS 的测试按 `"\n}"` 找规则时在 CRLF 下可能取到别处而通过，CI（Linux，LF）上才失败。每个提交另用 `git -c core.autocrlf=false archive <提交>` 做一份 LF 副本、换上转成 LF 的改动文件，确认 CR 字节为 0（`-text` 的夹具和二进制文件除外），再跑一次完整回归；普通的 `git archive` 在这台机器上写的是 CRLF，复现不了。push 后再看 GitHub Actions 的结果。

**等价比较**：沿用第 19 节的 `equivalence.cjs`（v3.4.0 用的是 149 张表：120 张公开表 + 29 张夹具和示例）。性能小提交起多比两项：`cards`（`state.checkupCards`）和 `cardWords`（导入那种语言的卡片句子）；和 9158a4b 比时用不含这两项的旧脚本（9158a4b 没有卡片）。① 允许的不同是 8 张表的 `lookAlikeRows` 和中英文报告里写法不一样那一行（共 24 处）；② 以后每个提交和上一个提交 0 处不同，和 9158a4b 比仍是这 24 处、逐字节相同。

**测不到、要在浏览器里看的**：首屏在 1366 / 1100 / 768 / 390 宽下中英文的样子（卡片三列 / 两列 / 一栏、8 类时 3 + 3 + 2、手机上大数字一行和按钮在底部）；比例条的段宽和颜色、点段滚到卡片；提示条是否被读屏念出（无头浏览器测不到）；侧栏圆点在 200 列的表里不撑破；打印时预览展开、比例条保留颜色。各提交的截图和实测数字在实施汇总里。
