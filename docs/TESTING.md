# Smart Tabular Analyzer 测试指南

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
| `tests/data-processing.test.cjs` | 纯数据逻辑：数值/日期解析、字段与 ID 识别、缺失、重复、统计、IQR、筛选、去重、`detectHeaderRow()` 的正例 / 反例 / 已知残留、Excel 行数上限、表头位置说明的摘要截断、`lineNumbersAtOffsets()` 与原来逐个偏移用正则计数的一致性（含随机文本），以及 `buildIdDuplicateSummary()` 的归一、字段顺序、排序、疑似非唯一阈值的边界、整行相同的分组、截断和 ID 重复值的说明文字。疑似缺失值：清单和两档的内容、`placeholderKey()` 的去空格 + NFKC + 整格匹配（全角写法能匹配，“Nan”“缺考”和部分匹配不算）、`buildPlaceholderSummary()` 的原文、行号截断和被忽略字段只列已确认的写法；护栏测试“placeholders never change the date strategy of a field without a year name”固定预估捷径的前提：清单里全部写法，以及 Ｎ／Ａ、－、／、－－、ＮａＮ、⽆、﹘、﹣ 这些经 NFKC 变成清单写法的字符，加进含 dmy、mdy 证据的日期值后，非年份名字段的 `buildDateStrategy()` 结果不变，`allowYearOnly` 为 false。改动清单或日期策略时，这条测试失败说明预估不能再走捷径。数字策略（v2.5.0）：应该转和不应该转的写法（不带策略、带本列策略、带“每个白名单单位都打开万/亿和全角”的最宽策略），单位冲突和货币符号与单位同列。四条护栏：“without a strategy, or with one that has no switch, toNumber() is the plain parser”固定单参数 `toNumber()` 和没有开关的策略都与 `toPlainNumber()` 逐值相同；“placeholder spellings never convert or give number strategy evidence (contract 17 premise)”固定契约 17 的前提：清单的全部写法、全角形式和带空格的全角形式，在上述最宽策略下都不能转换，`matchNumberWithUnit()` 也不匹配，只含这些写法的列算出的策略没有任何开关；“万 and 亿 shift the decimal point and give exact integers”用 `strictEqual` 和 `Number.isInteger` 固定 `0.07万` = 700 等结果，改回浮点乘法会失败；“getFieldNumberParser() is the plain parser unless the field's strategy has a switch”固定没有开关时取到的就是 `toPlainNumber` 本身。 |
| `tests/basic-data-processing.test.cjs` | CSV/Excel 行转换、中文数据、统计联动、原始质量视角和 CSV 导出的集成测试；包括跳过表头上方标题行的端到端导入、文件行号换算、状态栏与报告的“表头位置”、`state.headerDetection` 的清空与保持，以及回退提示。数据行的文件行号：Excel 的 `!ref` 偏移和中间空行、stage3 夹具、同一份 CSV 内容的 LF / 单独 CR / CRLF+BOM 三种写法、表头单元格里的引号内换行、第二遍扫描条数对不上时全部为 `null`、没经过导入的行查不到行号。护栏：只有表头、空文件、少一列、多一列这几种 CSV 的报错和解析警告与改前逐字相同；超过行数上限的 CSV 只调用一次 `Papa.parse`。ID 重复值：stage3 夹具和 `sample-user-behavior.csv` 的结果、Excel 数字与文本合并、截断写法、字段配置和筛选下的变化、新导入清空；质量页的 pill 与详情（含没有数值字段的表和转义）；HTML / Markdown 报告的“编号重复”一节、摘要、转义和旧快照兼容。疑似缺失值（文件末尾）：订单表的 5 条检出和行号、成绩表和全部示例（`examples/` 与 `tests/fixtures/examples/`）无检出（v2.11.0 起按钮示例 `sample-orders.csv` 恰好检出金额「N/A」和销售代表「—」两条）、忽略字段、新导入清空；确认 / 撤销 / 全部确认后的 profile、`totalMissing`、Top 10、筛选选项与匹配、分组、ID 重复值和年份列的日期策略；手动类型记录（含“确认后再对别的字段应用配置，撤销后手动类型仍在”）、有草稿时拒绝、恢复自动识别不撤销确认；`comparePlaceholderPreviews()` 逐条和整字段比较预估与完整重跑的 `buildColumnProfile()`；护栏：确认前后 `state.rows`、数据预览、CSV 导出和 `duplicateRows` 不变；兜底测试“after confirming every order placeholder, no analysis output shows N/A or a placeholder dash”在全部确认后检查分析区所有元素和图表标签（用替身 Chart 记录），只剥掉应用自己写的“当前类型无需转换”那个“—”，漏改后会让 N/A 或占位符「—」出现在分析区或图表里的，它都会失败；只影响计数的地方由各自的测试覆盖；质量页的静态小节、pill、详情文字、委托的点击处理，以及报告的“像是空值的写法”一节、摘要、转义和旧快照兼容。数字策略（文件末尾，v2.5.0）：订单表金额按“元”读取、格式问题 3 个值、金额统计值，其余 stage3 夹具不变；15 个“xx元”从 ID 变成数值，4G / 5G 和 8G 类写法不变；万/亿列的直方图按值计数；带单位的列上 `comparePlaceholderPreviews()` 和确认 / 撤销前后 `numberStrategy` 不变；数值筛选的范围和判定同口径；手动改类型；表头检测把 5kg 当文本；CSV 导出的 `-2200元` 仍加 `'`；主分析按大区、按日的元金额合计。疑似数值列：排除疑似缺失值后的口径、样例不含疑似写法、跟随已应用类型和确认、筛选不重算、重置清空。显示单位：字段标签的“已含单位”规则，统计表、图表标题和 y 轴带“（元）”而下拉框仍是原字段名，质量页两个小节的文字和隐藏，报告快照的三个新字段、两节内容、摘要只有“像是数字的文本列”的计数、旧快照兼容和转义。v2.6.0：空记录占满行数预览（6 万行每行后跟空行导入全部 60,000 行、首末行号正确；表头下 5 个空行 + 100,001 行报超限；正好 10 万行 + 结尾空行为护栏；带标题行的重试路径同样成立；第一次主解析仍带 `preview: 100001`，调用顺序固定为主解析、计数、整份重解析、行号）。精确行号：表头下空行、引号内换行、CRLF / 单独 CR、只有空白被丢掉的行、InvalidQuotes（`row` 计入表头和空记录）、带标题行的重试，警告和“第 N 行”报错（多一列、引号没有闭合）都按记事本的行号断言；重试里才发现的多一列显示重试自己的第 253 行和表头说明；25 行列数不足计 25 条、状态栏千分位、条数对不上时行号为 `null`。调用次数护栏：订单表 2 次、两个成绩表各 4 次（`成绩表-标题无逗号.csv` 第一次解析就报多一列，行号惰性求值才保持不变），第一行是表头又多一列时为主解析、表头检测、行号 3 次。中文警告：每个 code 有 / 无行号的消息（纯函数在 `data-processing`），分号 / 制表符 / 竖线两列文件各一条 `UndetectableDelimiter` 中文提示，单列“字段”文件 0 条（护栏），逗号多列文件不再列出分隔符警告。共用判断（护栏）：stage3 全部夹具和全部示例 CSV（`examples/` 加 `tests/fixtures/examples/`，加订单表全部确认后）上每个字段的数值、日期两种类型，`getProfileConversionCheck()` 数出的失败数都等于 `conversionFailuresByType[类型]`；profile 的 `confirmedKeys` 跟随确认 / 撤销且是副本；`groupIdValueIndexes()` / `sortIdDuplicates()` 的分组和排序，用“次数最多的值不最先出现”的数据（A/B/A/B/C/B/C）区分排序规则，`buildIdDuplicateSummary()` 同序。问题行：`problemRowCheck()` 在 15 个文件上逐类核对“列出的数量 = 质量页计数”（整行重复为列出行数 − 组数，ID 重复核对值的个数、前 10 个值的顺序、行数和行号）和顺序，订单表“全部确认 → 改类型 → 撤销”每一步都核对；ID 重复按表格顺序且超过 10 个值时完整列出；整行重复的“保留”；各类说明文字；解析警告含被丢掉的空白行和没有行号的条目。数据预览分页：每页 50 行、翻页和按钮禁用、筛选从末页回第 1 页（用筛选后仍有 2 页的条件，夹取碰不巧）、字段配置回第 1 页、页码夹取和没有分页 state、静态结构和绑定。查看区：每个“查看”按钮只带 `data-problem-index` 并打开自己的类别，“查看”列每行都有元素，分页、四种计数单位的摘要、下拉框文字、焦点、字段配置 / 确认 / 撤销后的找回和关闭（含没有数值字段、`renderQuality()` 提前 return 的数据）、导入关闭、筛选不影响、转义、关闭后焦点回到“查看”按钮。问题行导出：经真实下载路径捕获 Blob，核对 BOM、列顺序、“说明”列、公式中和（含额外列和表头）、`-2200元` 仍加 `'`、列名冲突、整类而不是当前页、没有对应行的解析警告、文件名（N/A、长字段名不截掉日期）、门禁和状态行；两种数据导出经新的共用函数逐字节不变（护栏）。v2.8.0 的测试见第 7 节。v2.8.1 的测试见第 8 节。 |
| `tests/destructive-qa.test.cjs` | 空数据、异常表头、编码、极端数值、连续上传、反复筛选、报告后继续分析和 100,000 行压力测试。 |
| `tests/core.test.cjs` | 启动自检、应用脚本顺序和顶层名字唯一性、DOM ID、相对资源路径、vendor 依赖与 SRI、许可证、`.nojekyll`、字段修正、报告生成和其他静态/集成断言。 |
| `tests/fixtures/` | `english-date-corpus.cjs` 被 `tests/data-processing.test.cjs` 引用，用于英文日期解析断言；`stage3/` 见下文；其余 CSV/XLSX 文件目前只用于手工回归，自动测试没有读取。 |
| `tests/fixtures/examples/` | `sample-data.csv`、`sample-sales.csv`、`sample-students.csv`、`sample-survey.csv`、`sample-used-products.csv`、`sample-user-behavior.csv`：只给测试用的 6 个示例（后 5 个 v2.7.0 从 `examples/` 移来，`sample-data.csv` v2.11.0 移来，字节不变）；`examples/` 里只剩“使用示例数据”按钮加载的 `sample-orders.csv`。`basic-data-processing.test.cjs` 经 `listExampleCsvFiles()` / `readExampleCsv()` 读两个目录，覆盖 7 个文件。内容按 utf8 文本读取，LF 和 CRLF 都通过（v2.11.0 的内容测试去掉 CR 后再比较），所以不需要 `.gitattributes` 规则。 |
| `tests/fixtures/stage3/` | 表头检测的验收夹具，内容全部虚构、固定。`订单表.csv`（GBK、CRLF，表头在第 1 行）和 `花名册.xlsx`（表头在第 1 行）是“不能误跳表头”的反例；`成绩表.csv`（UTF-8 BOM，标题行补了逗号）、`成绩表-标题无逗号.csv` 和 `门店报表.xlsx`（合并单元格标题 + 制表日期 + 空行，第 4 行是表头，含 0.00% 毛利率和合计行）是正例。`分区小计表.csv`（5 行小计 / 总计，带空格、冒号和括号注释的标签）和 `公式错误.xlsx`（4 个字段 7 个错误格，含 #SPILL! 和整行错误的行）是 v2.8.0 加的。`known-issues.json` 是按构造方式手写的已知问题清单和字段类型预期，手工加条目，保持一行一条的格式，不要用 `JSON.stringify` 重写；v2.8.0 的测试读取其中合计行和公式错误的预期，其余条目仍供人工核对；v2.5.0 起订单表金额按“元”读取 2200元、3500元，金额那条预期改为数值率 29/30，`expectedTotals.formatProblems` 从 5 改为 3。`写法不一致.csv`（UTF-8 BOM、CRLF，表头在第 1 行，24 行：日期写法混用、常量列和写法细节不同的值，不可见字符在生成脚本里写 `\u` 转义）是 v2.9.0 加的，它的预期（`dateStyles`、`constantFields`、`lookAlikes`）由测试读取。`tests/basic-data-processing.test.cjs` 经 `importStage3Fixture()` 读取这 8 个夹具；它用真实 PapaParse 包一层计数，并记下每次调用的 `header` 选项，表头在第 1 行的 CSV 正好调用两次（主解析一次，`header: false` 的行号扫描一次）。 |
| `tests/fixtures/stage3/gen-fixtures.cjs` | 重新生成上述 8 个夹具：在仓库根目录运行 `node tests/fixtures/stage3/gen-fixtures.cjs [输出目录]`，不带参数时写回本目录。只依赖 `vendor/` 里的 SheetJS 和 Node 自带的 `TextDecoder`（反推 GBK 编码表），输出与已提交文件逐字节一致，与运行时区无关。不在 CI 中运行。 |
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
| `v2.7.0 conclusions: the report shows the sentences on the page…` | 报告 `insights` 恰为页面上的 0–2 句；一句都没有时 HTML / Markdown 写“暂无自动分析结论。”；旧快照的 6 句原样输出，没有 `insights` 的快照得到 `[]` |
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
| `v2.8.0 report: the formula error note matches the page…` | `FORMULA_ERROR_NOTE` 和 `index.html` 同文；HTML、Markdown 都紧跟标题、在列表之前；只有合计或没有问题的表不出现 |
| `v2.8.0 report: the sections ignore filters, and the snapshot shares nothing with state` | 筛选确实减少分析行时报告不变；快照的每一层都不和 `state` 共用对象 |
| `v2.8.0 report: labels and fields are escaped for each output` | HTML 和 Markdown 各自转义，单元格内的换行变 `<br>` |

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
| `v2.12.0: old interface words no longer appear in pages, reports, file names or exported notes` | 禁用词（21 个）：缺失、未填写、疑似缺失值、疑似值、按缺失处理、格式问题、无法转换、转换失败、异常值、数值异常、写法细节不同、写法不同、解析警告、条警告、ID 字段重复、ID 重复、ID重复、常量列、疑似数值列、标准化、可安全转换。stage3 全部夹具和全部示例 CSV 各走一遍：导入后页面所有元素的 innerHTML / textContent / 选项文字、HTML 和 Markdown 报告、每个问题行类别的类别名 / aria / 导出文件名 / 全部说明 / 查看区、“检查会改哪些格”之后的面板和两个文件名、全部按空值处理后再检查一次（含改动清单 CSV 的内容）和报告、日期趋势图切到数值字段 + 求和后的标题（必须以“（求和）”结尾）和报告；去掉注释的 `index.html`，以及 `#dateChartAggregate` 的选项（有“求和”、没有“合计”）；页面和报告共用的 7 个说明常量里没有“汉字 空格 汉字” |
| `v2.12.0: charts, insights and the report say 认得出 when a filtered range has no dates or numbers` | 30 行的表（前 25 行组 A 有日期和金额，后 5 行组 B 都为空）筛选到组 B：日期趋势空图“该字段没有认得出的日期”、`state.chartInsights.trend`、HTML / Markdown 报告的趋势节；手动选“金额”后数值分布空图“该字段没有认得出的数字”。筛选后 `fieldHasNumericData()` 会把“金额”从 `#numericChartField` 去掉（浏览器里也一样，页面显示“没有可用的数值字段”），所以这一句要手动设下拉框才能走到 |

**为什么护栏里有一个内联 CSV**：stage3 和示例里没有一个文件带解析警告，也没有自定义分组的空组，“读取时有疑问（N 处）”“共 N 处”和空组名“（空值）”在夹具上碰不到，“条警告”“未填写”永远查不到。护栏最后导入一个 5 行的 CSV（一行少列、一行备注为空）并按备注做一次自定义分组。同理，“已按空值处理的写法，清空”只在按空值处理之后再检查时出现，而且只在改动清单 CSV 里，所以护栏在全部按空值处理后再点一次检查，并扫改动清单的内容。

**改动的现有断言**：pill 文字和按前缀找 pill 的查找词、状态栏、报告节标题 / 摘要 / 列头、像是空值的写法的详情 / 按钮 / 消息 / 快照 `effect`、问题行的类别名 / aria / 说明 / 文件名 / 查看区摘要、写法统一的表的标题 / 按钮 / 列头 / 文件名 / 状态行、格式一致性三节、趋势图“求和”、自定义分组空组名、`formatStandardizeAction()` 的两个标签。只换说法会**永远通过**的断言另外改了写法：

- 否定匹配换成新词后还要不被别的句子误伤：“编号重复”也出现在 `LOOK_ALIKE_NOTE` 和像是空值的写法的说明里，“像是空值的写法”也出现在像是数字的文本列的提示里，所以旧快照不出这一节的检查（`legacy report snapshots without ID duplicates…`、`report placeholder summaries appear only with candidates, and legacy snapshots…`）匹配标题或带冒号的摘要（`/<h3>编号重复<\/h3>|；编号重复：/` 等），不用裸词。
- 两条审查指令里没列出、grep 找到的：`markdown.indexOf("### ID 字段重复值") < markdown.indexOf("## 数值字段描述性统计")`（标题不在时 -1 < N 仍然成立）和 `doesNotMatch(maybeHtml, /全部确认常见写法/)`；实施时又找到一条：`doesNotMatch(html, /确认后类型会变/)`（v2.12.0 改成“按空值处理后类型会变”）。
- 切片：像是空值的写法小节按 `<h3 class="subsection-title">…</h3>` 切（按裸词“编号重复”切会切到前面的说明句，得到空串）；同一个值，写法不一样一节的否定匹配前先断言标题存在（标题不在时 `slice(-1)` 只剩最后一个字符）。
- 顺序：9 条“A 在 B 之前”改用 `assertInOrder(text, markers)`：每个 marker 先断言存在，再断言位置严格递增。只写 `indexOf(a) < indexOf(b)` 时 a 不在也会通过。报告大标题顺序那 4 条（`<h2>数据规模</h2>` 等）没有涉及改名，没有改。

**故意改坏的结果**（在分支快照上，用软断言记录每一处失败的断言行）：审查时的 15 个变异（让应该隐藏的节 / 计数 / 按钮出现、把标题改回旧说法、报告一节不转义）都让对应断言失败；另加 4 个：`#dateChartAggregate` 选项改回“合计”、`getDateChartAggregateLabel()` 改回“合计”、`LOOK_ALIKE_NOTE` 里加回空格，都让新护栏失败；没有类型变化时也显示“按空值处理后类型会变”，让 `doesNotMatch(html, /按空值处理后类型会变/)` 失败。“认得出的”四句各改回旧文字一次，第二条新测试都失败。

## 13. v3.0.0 新增测试固定了什么

v3.0.0 只换外壳和样式（一次显示一个模块、顶栏和侧栏、可收起的状态栏、新的上传页和“回到分析”、字段配置表 F1 F2、空状态 F3、新外观），检测规则、计数、报告文字和导出文件一样不变。新增 55 条测试（461 → 516），都在 `basic-data-processing.test.cjs`，名字以 `v3.0.0` 开头；`tests/test-context.cjs` 的假 `window` 加了 `scrollTo()` 和 `history.replaceState()`（`showModule()` 要用）。

| 测试组（名字前缀） | 条数 | 固定的内容 |
| --- | --- | --- |
| `v3.0.0 guard:` | 3 | 护栏，在 v2.12.0 上就通过：导入示例、生成自定义分析、应用筛选后 HTML 报告的图表张数和标题；每一句空状态文字在各自的触发条件下仍出现；示例数据的报告文字与 v2.12.0 相同（`buildHtmlReport(report, [])` 的 SHA-256，不含图片，所以不受图表配色影响） |
| `v3.0.0 modules:` | 5 | 8 个模块在一个舞台里；看不见的模块用 `visibility`，从不 `display: none` 或 `.hidden`，打印时全部显示；`showModule()` 只显示一个、只标一个链接、写地址、回到顶部、聚焦标题；导入停在数据概览；应用字段配置、按空值处理、筛选后停在当前模块 |
| `v3.0.0 shell:` | 8 | 顶栏的组成，没有 `section-label`；侧栏的静态部分和手机面板里的“换一个文件”；四个模块标题旁的数据范围句（只有数据概览那句带 id）和它们的“清除筛选”、侧栏的“筛选”一行都跟着 `#globalFilterSummary`；侧栏的字段、类型、单位和状态行；Excel 工作表名、表头在第 1 行时不显示、字段名转义；按宽度的默认开合和 Esc；`script.js` 里每个控件只绑一次 |
| `v3.0.0 status note:` | 5 | `#statusPanel` 原样、外面包一层；收起用 `clip-path`，读屏仍能读到；“〈文件名〉已载入，有 N 条说明”的 N；`warning` 不收起；成功消息只在数据概览显示，导出时的警告和报错在每个模块都显示 |
| `v3.0.0 upload page:` / `v3.0.0 way back:` | 3 / 3 | 上传页的新文字和示例小表（第 10、12、20、21、29、44 行的值和说明，第 20 行最后一列为空）；报错排在标题和卡片之间、载入文字在卡片下面；“换一个文件”保留分析、给出“回到「名」的分析”并回到原来的模块；没有分析、等选工作表、导入失败时不给入口 |
| `v3.0.0 F1:` / `v3.0.0 F2:` | 2 / 6 | 字段配置表的 7 列和“认不出”的粉点、草稿即时更新；多于 20 个字段时的搜索和类型切换、个数、隐藏行、应用后保留、新导入清空、转义 |
| `v3.0.0 F3:` | 4 | 每个图表框有静态的空状态段落；`drawEmptyChart()` 逐字写文字、不在画布上画，下一张图清掉它；自定义分析的空趋势；空表格和空列表靠容器上的类得到统一外观，句子不变 |
| `v3.0.0 overview:` 等（7a） | 9 | 5 个数字写成一行；数据预览在 `<details open>` 里；字段配置的说明和按钮在表格下方、有草稿时加 `has-draft`；只有多于一句的小节说明有“展开”；频数比例条；导出页各选项标题后的行数；截断单元格的悬停全文和图表标题的提示；单元格限宽、手机上示例小表的两行排法、质量页两张表上下排 |
| `v3.0.0 look:` | 7 | 只读 `style.css`、`app.js` 的 `CHART_THEME` 和 `script.js` 的 Chart.js 默认值：颜色变量、系统字体栈（`Segoe UI` 在微软雅黑之前）、7 档字号和 3 档字重、焦点环；模块不是卡片，内容块是白卡，`box-shadow` 只给侧栏浮层；按钮高度、次按钮和文字按钮、手机上的 44 px、分段导航、表头；质量页标签中性、没有琥珀色和圆点，5 个小节的类别色圆点，“认不出”的粉点，问题格浅灰加粗；一组一个主按钮（只用 id 选择器）；空值检查表约 10 行不滚动、640 以下不限高、打印不限高；图表主色 `#0068e0`、第二色比主色浅且对白底不低于 3:1、图表结论和自定义分析结果句的浅灰框 |

**改动的现有断言（2 处，都和导航有关）**：`basic` 的 `v2.7.0 conclusions: the findings section and its navigation entries are gone…` 原来固定导航编号 01–08 和手机下拉框，改成只有一个导航、恰好 8 个链接、依次指向 8 个模块 id、没有编号、没有 `#mobileSectionNav`（旧写法里的 `<span>09</span>` 一支在去掉编号后会永远匹配不到）；`core.test.cjs` 原来断言有 `#mobileSectionNav`，改成断言没有它，并且恰好 8 个指向模块的链接。读样式表的 4 条旧断言（`workflow-`、`.template-`、`:not(.has-note)`、打印块隐藏 `.standardized-export-actions`）都没改，仍然通过。

**外观测试是怎么验证的**：`v3.0.0 look:` 的每一条都先在改样式前的代码上跑过，确认失败，改完再通过。后来补的两句断言（筛选页日期框不超出卡片、自定义分析结果句的样式）也各自确认过：去掉对应的规则就失败。

假 DOM 不算样式和尺寸：模块切换后的可见性、Tab 走不进看不见的模块、隐藏模块里图表的导出尺寸、吸顶表头、侧栏在 1100 / 641–1099 / ≤640 三档的行为、横向滚动、打印，都要在浏览器里看（`docs/RELEASE_CHECKLIST.md` 第 7 节）。
