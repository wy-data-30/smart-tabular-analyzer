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
| `tests/data-processing.test.cjs` | 纯数据逻辑：数值/日期解析、字段与 ID 识别、缺失、重复、统计、IQR、筛选、去重、模板、`detectHeaderRow()` 的正例 / 反例 / 已知残留、Excel 行数上限、表头位置说明的摘要截断、`lineNumbersAtOffsets()` 与原来逐个偏移用正则计数的一致性（含随机文本），以及 `buildIdDuplicateSummary()` 的归一、字段顺序、排序、疑似非唯一阈值的边界、整行相同的分组、截断和 ID 重复值的说明文字。疑似缺失值：清单和两档的内容、`placeholderKey()` 的去空格 + NFKC + 整格匹配（全角写法能匹配，“Nan”“缺考”和部分匹配不算）、`buildPlaceholderSummary()` 的原文、行号截断和被忽略字段只列已确认的写法；护栏测试“placeholders never change the date strategy of a field without a year name”固定预估捷径的前提：清单里全部写法，以及 Ｎ／Ａ、－、／、－－、ＮａＮ、⽆、﹘、﹣ 这些经 NFKC 变成清单写法的字符，加进含 dmy、mdy 证据的日期值后，非年份名字段的 `buildDateStrategy()` 结果不变，`allowYearOnly` 为 false。改动清单或日期策略时，这条测试失败说明预估不能再走捷径。数字策略（v2.5.0）：应该转和不应该转的写法（不带策略、带本列策略、带“每个白名单单位都打开万/亿和全角”的最宽策略），单位冲突和货币符号与单位同列。四条护栏：“without a strategy, or with one that has no switch, toNumber() is the plain parser”固定单参数 `toNumber()` 和没有开关的策略都与 `toPlainNumber()` 逐值相同；“placeholder spellings never convert or give number strategy evidence (contract 17 premise)”固定契约 17 的前提：清单的全部写法、全角形式和带空格的全角形式，在上述最宽策略下都不能转换，`matchNumberWithUnit()` 也不匹配，只含这些写法的列算出的策略没有任何开关；“万 and 亿 shift the decimal point and give exact integers”用 `strictEqual` 和 `Number.isInteger` 固定 `0.07万` = 700 等结果，改回浮点乘法会失败；“getFieldNumberParser() is the plain parser unless the field's strategy has a switch”固定没有开关时取到的就是 `toPlainNumber` 本身。 |
| `tests/basic-data-processing.test.cjs` | CSV/Excel 行转换、中文数据、统计联动、原始质量视角和 CSV 导出的集成测试；包括跳过表头上方标题行的端到端导入、文件行号换算、状态栏与报告的“表头位置”、`state.headerDetection` 的清空与保持，以及回退提示。数据行的文件行号：Excel 的 `!ref` 偏移和中间空行、stage3 夹具、同一份 CSV 内容的 LF / 单独 CR / CRLF+BOM 三种写法、表头单元格里的引号内换行、第二遍扫描条数对不上时全部为 `null`、没经过导入的行查不到行号。护栏：只有表头、空文件、少一列、多一列这几种 CSV 的报错和解析警告与改前逐字相同；超过行数上限的 CSV 只调用一次 `Papa.parse`。ID 重复值：stage3 夹具和 `sample-user-behavior.csv` 的结果、Excel 数字与文本合并、截断写法、字段配置和筛选下的变化、新导入清空；质量页的 pill 与详情（含没有数值字段的表和转义）；HTML / Markdown 报告的“ID 字段重复值”一节、摘要、转义和旧快照兼容。疑似缺失值（文件末尾）：订单表的 5 条检出和行号、成绩表和 `examples/` 无检出、忽略字段、新导入清空；确认 / 撤销 / 全部确认后的 profile、`totalMissing`、Top 10、筛选选项与匹配、分组、ID 重复值和年份列的日期策略；手动类型记录（含“确认后再对别的字段应用配置，撤销后手动类型仍在”）、有草稿时拒绝、恢复自动识别不撤销确认；`comparePlaceholderPreviews()` 逐条和整字段比较预估与完整重跑的 `buildColumnProfile()`；护栏：确认前后 `state.rows`、数据预览、CSV 导出和 `duplicateRows` 不变；兜底测试“after confirming every order placeholder, no analysis output shows N/A or a placeholder dash”在全部确认后检查分析区所有元素和图表标签（用替身 Chart 记录），只剥掉应用自己写的“当前类型无需转换”那个“—”，漏改后会让 N/A 或占位符「—」出现在分析区或图表里的，它都会失败；只影响计数的地方由各自的测试覆盖；质量页的静态小节、pill、详情文字、委托的点击处理，以及报告的“疑似缺失值”一节、摘要、转义和旧快照兼容。数字策略（文件末尾，v2.5.0）：订单表金额按“元”读取、格式问题 3 个值、金额统计值，其余 stage3 夹具不变；15 个“xx元”从 ID 变成数值，4G / 5G 和 8G 类写法不变；万/亿列的直方图按值计数；带单位的列上 `comparePlaceholderPreviews()` 和确认 / 撤销前后 `numberStrategy` 不变；数值筛选的范围和判定同口径；手动改类型；表头检测把 5kg 当文本；CSV 导出的 `-2200元` 仍加 `'`；销售模板和主分析的合计相同。疑似数值列：排除疑似缺失值后的口径、样例不含疑似写法、跟随已应用类型和确认、筛选不重算、重置清空。显示单位：字段标签的“已含单位”规则，统计表、图表标题和 y 轴带“（元）”而下拉框仍是原字段名，质量页两个小节的文字和隐藏，报告快照的三个新字段、两节内容、摘要只有“疑似数值列”的计数、旧快照兼容和转义。v2.6.0：空记录占满行数预览（6 万行每行后跟空行导入全部 60,000 行、首末行号正确；表头下 5 个空行 + 100,001 行报超限；正好 10 万行 + 结尾空行为护栏；带标题行的重试路径同样成立；第一次主解析仍带 `preview: 100001`，调用顺序固定为主解析、计数、整份重解析、行号）。精确行号：表头下空行、引号内换行、CRLF / 单独 CR、只有空白被丢掉的行、InvalidQuotes（`row` 计入表头和空记录）、带标题行的重试，警告和“第 N 行”报错（多一列、引号没有闭合）都按记事本的行号断言；重试里才发现的多一列显示重试自己的第 253 行和表头说明；25 行列数不足计 25 条、状态栏千分位、条数对不上时行号为 `null`。调用次数护栏：订单表 2 次、两个成绩表各 4 次（`成绩表-标题无逗号.csv` 第一次解析就报多一列，行号惰性求值才保持不变），第一行是表头又多一列时为主解析、表头检测、行号 3 次。中文警告：每个 code 有 / 无行号的消息（纯函数在 `data-processing`），分号 / 制表符 / 竖线两列文件各一条 `UndetectableDelimiter` 中文提示，单列“字段”文件 0 条（护栏），逗号多列文件不再列出分隔符警告。共用判断（护栏）：stage3 全部夹具和 `examples/` 全部 CSV（加订单表全部确认后）上每个字段的数值、日期两种类型，`getProfileConversionCheck()` 数出的失败数都等于 `conversionFailuresByType[类型]`；profile 的 `confirmedKeys` 跟随确认 / 撤销且是副本；`groupIdValueIndexes()` / `sortIdDuplicates()` 的分组和排序，用“次数最多的值不最先出现”的数据（A/B/A/B/C/B/C）区分排序规则，`buildIdDuplicateSummary()` 同序。问题行：`problemRowCheck()` 在 11 个文件上逐类核对“列出的数量 = 质量页计数”（整行重复为列出行数 − 组数，ID 重复核对值的个数、前 10 个值的顺序、行数和行号）和顺序，订单表“全部确认 → 改类型 → 撤销”每一步都核对；ID 重复按表格顺序且超过 10 个值时完整列出；整行重复的“保留”；各类说明文字；解析警告含被丢掉的空白行和没有行号的条目。数据预览分页：每页 50 行、翻页和按钮禁用、筛选从末页回第 1 页（用筛选后仍有 2 页的条件，夹取碰不巧）、字段配置回第 1 页、页码夹取和没有分页 state、静态结构和绑定。查看区：每个“查看”按钮只带 `data-problem-index` 并打开自己的类别，“查看”列每行都有元素，分页、四种计数单位的摘要、下拉框文字、焦点、字段配置 / 确认 / 撤销后的找回和关闭（含没有数值字段、`renderQuality()` 提前 return 的数据）、导入关闭、筛选不影响、转义、关闭后焦点回到“查看”按钮。问题行导出：经真实下载路径捕获 Blob，核对 BOM、列顺序、“说明”列、公式中和（含额外列和表头）、`-2200元` 仍加 `'`、列名冲突、整类而不是当前页、没有对应行的解析警告、文件名（N/A、长字段名不截掉日期）、门禁和状态行；两种数据导出经新的共用函数逐字节不变（护栏）。 |
| `tests/destructive-qa.test.cjs` | 空数据、异常表头、编码、极端数值、连续上传、反复筛选、报告后继续分析和 100,000 行压力测试。 |
| `tests/core.test.cjs` | 启动自检、应用脚本顺序和顶层名字唯一性、DOM ID、相对资源路径、vendor 依赖与 SRI、许可证、`.nojekyll`、字段修正、报告生成和其他静态/集成断言。 |
| `tests/fixtures/` | `english-date-corpus.cjs` 被 `tests/data-processing.test.cjs` 引用，用于英文日期解析断言；`stage3/` 见下文；其余 CSV/XLSX 文件目前只用于手工回归，自动测试没有读取。 |
| `tests/fixtures/stage3/` | 表头检测的验收夹具，内容全部虚构、固定。`订单表.csv`（GBK、CRLF，表头在第 1 行）和 `花名册.xlsx`（表头在第 1 行）是“不能误跳表头”的反例；`成绩表.csv`（UTF-8 BOM，标题行补了逗号）、`成绩表-标题无逗号.csv` 和 `门店报表.xlsx`（合并单元格标题 + 制表日期 + 空行，第 4 行是表头，含 0.00% 毛利率和合计行）是正例。`known-issues.json` 是按构造方式手写的已知问题清单和字段类型预期，供人工核对，测试不读取；v2.5.0 起订单表金额按“元”读取 2200元、3500元，金额那条预期改为数值率 29/30，`expectedTotals.formatProblems` 从 5 改为 3。`tests/basic-data-processing.test.cjs` 经 `importStage3Fixture()` 读取这 5 个夹具；它用真实 PapaParse 包一层计数，并记下每次调用的 `header` 选项，表头在第 1 行的 CSV 正好调用两次（主解析一次，`header: false` 的行号扫描一次）。 |
| `tests/fixtures/stage3/gen-fixtures.cjs` | 重新生成上述 5 个夹具：在仓库根目录运行 `node tests/fixtures/stage3/gen-fixtures.cjs [输出目录]`，不带参数时写回本目录。只依赖 `vendor/` 里的 SheetJS 和 Node 自带的 `TextDecoder`（反推 GBK 编码表），输出与已提交文件逐字节一致，与运行时区无关。不在 CI 中运行。 |
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
| 分类图默认字段 5 条（basic） | 规则的两个边界（不同值 ≤ 20、首位次数 ≥ 2）各有刚好满足 / 刚好不满足的数据；没有集中字段时选第一个分类字段，没有分类字段时为空；重建保留所选、所选字段不再是分类字段时回到规则、新导入不继承；与“主要发现”选同一个字段。每条先断言数据里各字段的实际类型 |

两条调用次数护栏和 trim 次数护栏钉的是**性能性质**，不是结果：以后合理的实现变化（例如换了函数名、换了复用方式）可以随之更新这些测试，但要在 PR 里说明为什么。

假 DOM 的 `<select>` 在改写 `innerHTML` 后 `value` 不会像浏览器那样重置为第一个选项，所以默认字段的测试只能确认代码显式给 `categoryChartField.value` 赋了哪个字段；`setSelectOptions()` 在浏览器里的重置行为要靠手动 / 无头浏览器验证（v2.6.1 已用无头 Chromium 对照过 v2.6.0）。
