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
| `tests/test-context.cjs` | 创建 VM 上下文和轻量 `FakeElement`，把真实 `script.js` 加载到隔离环境。 |
| `tests/data-processing.test.cjs` | 纯数据逻辑：数值/日期解析、字段与 ID 识别、缺失、重复、统计、IQR、筛选、去重和模板。 |
| `tests/basic-data-processing.test.cjs` | CSV/Excel 行转换、中文数据、统计联动、原始质量视角和 CSV 导出的集成测试。 |
| `tests/destructive-qa.test.cjs` | 空数据、异常表头、编码、极端数值、连续上传、反复筛选、报告后继续分析和 100,000 行压力测试。 |
| `tests/core.test.cjs` | DOM ID、相对资源路径、vendor 依赖与 SRI、许可证、`.nojekyll`、字段修正、报告生成和其他静态/集成断言。 |
| `tests/fixtures/` | `english-date-corpus.cjs` 被 `tests/data-processing.test.cjs` 引用，用于英文日期解析断言；其余 CSV/XLSX 文件目前只用于手工回归，自动测试没有读取。 |
| `tests/create-excel-fixture.mjs` | 辅助夹具生成脚本，不属于 npm 测试入口，并引用未写入 `package.json` 的外部工具。 |

测试环境不是完整浏览器：Chart.js、PapaParse、SheetJS、FileReader 或下载 API 会按测试需要注入替身。真实浏览器布局、触控交互、canvas 像素和 Excel/WPS 打开效果仍需手工回归。

## 3. 如何增加测试

1. 把测试放到最接近风险的文件：纯函数放 `data-processing`，导入/导出联动放 `basic`，边界与压力放 `destructive-qa`，DOM/资源/报告契约放 `core`。
2. 使用 `createScriptContext()` 创建新上下文，避免前一个测试残留 `state`。
3. 通过 `context` 放入测试数据，再用 `evaluate()` 调用 `script.js` 中的真实函数。
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

## 4. 常见失败原因

### `node` 或 `npm` 不可用

确认 Node.js 已加入 PATH。没有 npm 时可直接执行 `package.json` 中对应的 `node --test ...` 命令。

### 新增 DOM API 后测试报错

`FakeElement` 只实现当前测试需要的方法。若生产代码开始调用新的 DOM 方法或属性，应在 `tests/test-context.cjs` 中增加最小替身，不要引入完整浏览器框架来掩盖问题。

### `Papa`、`XLSX`、`Chart` 或 `FileReader` 未定义

VM 只加载 `script.js`，不会加载 `vendor/` 中的库文件。测试调用相关流程前，需要同时在 `context` 和必要的 `context.window` 属性上注入替身。

### VM 对象深比较失败

VM 中创建的对象属于不同 realm。复杂对象可先 `JSON.stringify()` 再解析到测试进程，或只断言必要字段。

### 日期测试在不同时区表现不同

使用明确格式和 UTC 预期，例如 `YYYY-MM-DDT00:00:00.000Z`。不要用含义不明确的 `MM/DD/YYYY` 作为无策略输入。

### `core.test.cjs` 报 DOM ID 或资源缺失

该测试会从 `script.js` 收集 `getElementById()`，并检查 `index.html`、相对资源路径、vendor 脚本标签及其 SRI、许可证文件和 `.nojekyll`。修改 ID、移动根目录资源或升级 vendor 依赖时必须同步 HTML、JS 和测试契约。

### 压力测试变慢

`destructive-qa.test.cjs` 会处理 100,000 行。先确认是否引入了重复全表扫描、嵌套 `filter()` 或无必要的数据复制，再考虑调整测试；不要通过降低数据量隐藏性能回归。

### Excel 夹具生成脚本失败

该脚本不是标准测试依赖，也不在 CI 中运行。发布判断应以已提交夹具的手工回归和正式测试命令为准；若要把生成流程纳入项目，需先明确依赖和可重复执行方式。

