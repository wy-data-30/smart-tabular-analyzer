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
| `tests/data-processing.test.cjs` | 纯数据逻辑：数值/日期解析、字段与 ID 识别、缺失、重复、统计、IQR、筛选、去重、模板、`detectHeaderRow()` 的正例 / 反例 / 已知残留、Excel 行数上限、表头位置说明的摘要截断、`lineNumbersAtOffsets()` 与原来逐个偏移用正则计数的一致性（含随机文本），以及 `buildIdDuplicateSummary()` 的归一、字段顺序、排序、疑似非唯一阈值的边界、整行相同的分组、截断和 ID 重复值的说明文字。疑似缺失值：清单和两档的内容、`placeholderKey()` 的去空格 + NFKC + 整格匹配（全角写法能匹配，“Nan”“缺考”和部分匹配不算）、`buildPlaceholderSummary()` 的原文、行号截断和被忽略字段只列已确认的写法；护栏测试“placeholders never change the date strategy of a field without a year name”固定预估捷径的前提：清单里全部写法，以及 Ｎ／Ａ、－、／、－－、ＮａＮ、⽆、﹘、﹣ 这些经 NFKC 变成清单写法的字符，加进含 dmy、mdy 证据的日期值后，非年份名字段的 `buildDateStrategy()` 结果不变，`allowYearOnly` 为 false。改动清单或日期策略时，这条测试失败说明预估不能再走捷径。 |
| `tests/basic-data-processing.test.cjs` | CSV/Excel 行转换、中文数据、统计联动、原始质量视角和 CSV 导出的集成测试；包括跳过表头上方标题行的端到端导入、文件行号换算、状态栏与报告的“表头位置”、`state.headerDetection` 的清空与保持，以及回退提示。数据行的文件行号：Excel 的 `!ref` 偏移和中间空行、stage3 夹具、同一份 CSV 内容的 LF / 单独 CR / CRLF+BOM 三种写法、表头单元格里的引号内换行、第二遍扫描条数对不上时全部为 `null`、没经过导入的行查不到行号。护栏：只有表头、空文件、少一列、多一列这几种 CSV 的报错和解析警告与改前逐字相同；超过行数上限的 CSV 只调用一次 `Papa.parse`。ID 重复值：stage3 夹具和 `sample-user-behavior.csv` 的结果、Excel 数字与文本合并、截断写法、字段配置和筛选下的变化、新导入清空；质量页的 pill 与详情（含没有数值字段的表和转义）；HTML / Markdown 报告的“ID 字段重复值”一节、摘要、转义和旧快照兼容。疑似缺失值（文件末尾）：订单表的 5 条检出和行号、成绩表和 `examples/` 无检出、忽略字段、新导入清空；确认 / 撤销 / 全部确认后的 profile、`totalMissing`、Top 10、筛选选项与匹配、分组、ID 重复值和年份列的日期策略；手动类型记录（含“确认后再对别的字段应用配置，撤销后手动类型仍在”）、有草稿时拒绝、恢复自动识别不撤销确认；`comparePlaceholderPreviews()` 逐条和整字段比较预估与完整重跑的 `buildColumnProfile()`；护栏：确认前后 `state.rows`、数据预览、CSV 导出和 `duplicateRows` 不变；兜底测试“after confirming every order placeholder, no analysis output shows N/A or a placeholder dash”在全部确认后检查分析区所有元素和图表标签（用替身 Chart 记录），只剥掉应用自己写的“当前类型无需转换”那个“—”，漏改后会让 N/A 或占位符「—」出现在分析区或图表里的，它都会失败；只影响计数的地方由各自的测试覆盖；质量页的静态小节、pill、详情文字、委托的点击处理，以及报告的“疑似缺失值”一节、摘要、转义和旧快照兼容。 |
| `tests/destructive-qa.test.cjs` | 空数据、异常表头、编码、极端数值、连续上传、反复筛选、报告后继续分析和 100,000 行压力测试。 |
| `tests/core.test.cjs` | 启动自检、应用脚本顺序和顶层名字唯一性、DOM ID、相对资源路径、vendor 依赖与 SRI、许可证、`.nojekyll`、字段修正、报告生成和其他静态/集成断言。 |
| `tests/fixtures/` | `english-date-corpus.cjs` 被 `tests/data-processing.test.cjs` 引用，用于英文日期解析断言；`stage3/` 见下文；其余 CSV/XLSX 文件目前只用于手工回归，自动测试没有读取。 |
| `tests/fixtures/stage3/` | 表头检测的验收夹具，内容全部虚构、固定。`订单表.csv`（GBK、CRLF，表头在第 1 行）和 `花名册.xlsx`（表头在第 1 行）是“不能误跳表头”的反例；`成绩表.csv`（UTF-8 BOM，标题行补了逗号）、`成绩表-标题无逗号.csv` 和 `门店报表.xlsx`（合并单元格标题 + 制表日期 + 空行，第 4 行是表头，含 0.00% 毛利率和合计行）是正例。`known-issues.json` 是按构造方式手写的已知问题清单和字段类型预期，供人工核对，测试不读取。`tests/basic-data-processing.test.cjs` 经 `importStage3Fixture()` 读取这 5 个夹具；它用真实 PapaParse 包一层计数，并记下每次调用的 `header` 选项，表头在第 1 行的 CSV 正好调用两次（主解析一次，`header: false` 的行号扫描一次）。 |
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

经 `parseCsvText()` 导入成功的 CSV，在提交前还会调用一次 `Papa.parse(text, { header: false, skipEmptyLines: false, step, ... })` 求文件行号；这次调用没有 `complete`，只靠同步的 `step`。替身 Papa 要在 `options.header === false` 时，对每条记录同步调用一次 `options.step({ data, errors: [], meta: { cursor } })`，`cursor` 是这条记录结束（含换行）之后的位置，然后直接返回；不要在这次调用里调用 `options.complete`（它不存在，会抛错），也不要用它覆盖记录第一次调用选项的变量。替身给出的记录数和主解析数据行数不一致时，导入照常完成，只是全部行号为 `null`。直接调用 `handleParsedData()`、不经过 `parseCsvText()` 的测试没有原文，不会触发第二次调用。

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

