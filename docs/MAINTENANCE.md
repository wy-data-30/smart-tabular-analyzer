# Smart Tabular Analyzer 维护指南

## 1. 新增功能的推荐步骤

1. 先确定功能使用原始数据还是筛选数据：质量检查使用 `state.rows`，普通分析使用 `getAnalysisRows()`。
2. 优先把计算写成接收参数并返回结果的纯函数，再由事件处理函数更新 `state` 和 DOM。
3. 新增控件时同步修改 `index.html`、`app.js` 顶部的 `dom` 映射、`script.js` 中的事件监听和对应渲染函数。
4. 新增状态时明确其创建、重置、连续上传和字段类型修正时的生命周期。
5. 新增图表时使用唯一 canvas ID，把实例存入 `state.charts`，并在重绘前调用 `destroyChart(id)`。
6. 新增用户内容到 `innerHTML` 或导出报告时必须经过对应转义函数。
7. 为纯逻辑和主要边界增加测试，并运行 `npm run test:regression`。
8. 更新 README 或 `docs/` 中受影响的行为说明。

浏览器逻辑按职责分在 10 个普通 `<script>` 文件里（加载顺序和各文件职责见 `docs/ARCHITECTURE.md` 的“应用脚本”表），没有打包器或模块系统，所有文件共享同一个全局作用域：

- 新代码放进职责对应的文件；小改动应遵循现有函数边界，不要仅为代码风格进行大范围搬动。
- 加载时立即执行的代码（事件绑定、全局设置等）只放在最后加载的 `script.js`，其他文件只声明函数和常量。
- 顶层名字在全部应用脚本中必须唯一：重复的 `const` / `let` 会让整个文件无法执行，重复的函数声明会静默覆盖前一个。
- 新增应用脚本时，要同时更新 `index.html`（放在 `script.js` 之前，并带 `?v=<版本>`）、`tests/core.test.cjs` 中的固定文件列表，以及 `script.js` 启动自检 `ensureDomElementsPresent()` 中的代表函数列表。

## 2. 不要随便修改的核心逻辑

### 导入状态

- `beginImport()`、`activeImportId`、`activeFileReader` 和 `resetAnalysisState()` 防止连续上传时旧任务覆盖新数据。
- `commitTabularData()` 是 CSV 和 Excel 的共同提交入口，负责建立所有派生状态。它的第 9 个参数 `rowNumbers` 与 `rows` 一一对应；新增生成行对象的导入路径时要同时产出文件行号，不要往行对象里加字段（行号存在 `state.sourceRowNumbers`，经 `getSourceRowNumber()` 读取）。
- 修改这些函数必须覆盖空文件、多次上传、多工作表和运行时依赖失败。

### 原始数据与筛选数据

- `state.rows` 不得因筛选、统计或去重被原地修改。
- `filteredRows` 只保存当前结果，`filterSourceRows` 用于确认它属于当前数据集。
- 数据质量报告必须继续使用原始数据视角；不要把 `qualityNumericStats` 替换成筛选统计。

### 类型和转换

- `isMissing()`、`toNumber()`、`toDate()`、`buildDateStrategy()`、`isIdFieldName()` 和 `buildColumnProfile()` 共同定义字段识别契约。
- 当前阈值是日期或数值转换成功率达到 80%；ID 识别优先于日期和数值。
- 手动类型修正不改写原值，只记录转换失败并跳过无效值。
- 调整规则前应增加中文字段、混合类型、百分比、货币、年份和歧义日期测试。

### 质量与统计

- 重复行通过完整字段顺序和值类型生成精确键。
- ID 字段重复值用 `normalizeValue()` 比较（去首尾空格、区分大小写），缺失值不参与，基于原始行，导入和字段配置变化时重算、筛选不重算。疑似非唯一的阈值（`ID_MOSTLY_REPEATED_*`）和列出上限（`ID_DUPLICATE_*`）是 `quality.js` 里的常量，修改时同步测试和 `docs/ARCHITECTURE.md`；“其中 N 行内容完全相同”复用 `buildExactRowKey()`，不要另写一套整行比较。
- 异常值使用 `Q1 - 1.5 * IQR` 和 `Q3 + 1.5 * IQR`。
- 标准差当前按总体标准差计算。
- 极端有限数值通过缩放求和和直方图计算避免 `Infinity`；不要退回直接 `max - min` 或无保护求和。

### 图表与导出安全

- 创建 Chart.js 实例前必须销毁同 ID 的旧实例。
- CSV 导出必须保留 UTF-8 BOM、字段顺序和公式前缀中和。
- HTML 报告中的数据必须经过 `escapeHtml()`；图像只接受安全的 PNG、JPEG 或 WebP Data URL（与 `isSafeChartDataUrl()` 一致）。
- 文件名清理负责移除非法字符和双向文本控制字符。

## 3. 依赖位置（vendor/ 自行托管）

浏览器运行时依赖保存在仓库的 `vendor/` 目录，由 `index.html` 顶部按以下顺序以相对路径加载，并且都在所有应用脚本之前：

| 库 | 当前版本 | 文件 | 许可证 | 用途 |
| --- | --- | --- | --- | --- |
| PapaParse | 5.4.1 | `vendor/papaparse-5.4.1.min.js` | `vendor/licenses/papaparse-LICENSE`（MIT） | CSV 解析和 CSV 生成 |
| Chart.js | 4.4.1 | `vendor/chart-4.4.1.umd.min.js` | `vendor/licenses/chartjs-LICENSE.md`（MIT） | 页面图表 |
| SheetJS | 0.20.3 | `vendor/xlsx-0.20.3.full.min.js` | `vendor/licenses/sheetjs-LICENSE`（Apache-2.0） | `.xlsx` / `.xls` 读取 |

三个脚本都保留 SHA-384 `integrity`，页面不再请求第三方 CDN。仓库根目录的 `.nojekyll` 让 GitHub Pages 原样发布 `vendor/`。`tests/core.test.cjs` 会校验：没有 http(s) 外部脚本、三个 vendor 文件名与顺序、每个文件的实际 SHA-384 与 `integrity` 一致、三个许可证文件和 `.nojekyll` 存在。

升级依赖时：

1. 从该库官方发布渠道下载新版本的精确文件，不使用 `latest`。
2. 替换 `vendor/` 中的旧文件，文件名使用新版本号，并同步修改 `index.html` 中的 `src`。
3. 对新文件重新计算 SHA-384 Base64，更新对应 `integrity`：

   ```bash
   node -e "console.log('sha384-'+require('crypto').createHash('sha384').update(require('fs').readFileSync(process.argv[1])).digest('base64'))" vendor/<文件名>
   ```

4. 从同一个包中取得许可证文件，替换 `vendor/licenses/` 中对应的文件并核对许可证类型。
5. 同步修改 `tests/core.test.cjs` 中的文件名，运行完整回归，并在浏览器中实际加载页面。升级 PapaParse 时，还要对照新版源码确认 `findCsvRecordLineNumbers()` 依赖的两条行为没变：`skipEmptyLines: true` 跳过的是只有一个空字符串的记录、并且在取表头之前过滤；字符串输入时 `step` 同步调用，`meta.cursor` 是这条记录结束（含换行）之后的位置。stage3 夹具和行号测试要全部通过。
6. 发布后确认线上 `vendor/` 下三个文件都返回 200。

不要只替换文件而保留旧 SRI，否则浏览器会拒绝执行脚本。

`package.json` 当前只有测试脚本，没有生产依赖或开发依赖。测试使用 Node 内置模块。`tests/create-excel-fixture.mjs` 引用的外部工具不是标准项目依赖，也不由 CI 执行。

## 4. GitHub Pages 部署

项目不需要构建产物，发布目录就是仓库根目录。必须保持：

- 入口为根目录 `index.html`。
- `style.css`、10 个应用脚本、示例 CSV 和 `assets/` 使用相对路径。
- 不添加本地绝对路径、开发服务器代理或后端 API 依赖。
- `vendor/` 下三个库文件随仓库发布、线上返回 200，且 SRI 与内容一致；根目录保留 `.nojekyll`，`.gitattributes` 保持 `vendor/** -text` 和 `tests/fixtures/stage3/** -text`，避免换行转换改变文件字节。

仓库内的 `.github/workflows/ci.yml` 只运行测试，没有 Pages 部署步骤或写权限，因此不会替代 Pages 配置。当前代码适合 GitHub Pages 的“从分支部署”；实际来源分支和目录应在仓库 `Settings > Pages` 中确认。

当前演示地址是：

```text
https://wy-data-30.github.io/smart-tabular-analyzer/
```

仓库名或所有者变化时，需要同步 README 的 Live Demo、报告页脚中的仓库链接和 Pages 地址。

## 5. 常见 bug 排查

### 页面空白或按钮无响应

1. 查看浏览器控制台是否有语法错误或 DOM ID 为 `null`。
2. 在 Network 面板检查 `vendor/` 下三个库文件是否为 200，以及控制台是否出现 SRI（integrity）拒绝；若本地出现 SRI 拒绝，检查文件是否被换行转换改成了 CRLF。
3. 确认 10 个应用脚本都在页面底部按顺序加载、`script.js` 在最底部最后加载，且 HTML ID 与 `app.js` 中的 `dom` 映射一致。页面顶部出现“页面已更新，请刷新后再使用。”时，控制台错误会列出缺失的脚本文件或 DOM 键名。
4. 逐个文件运行 `node --check`，再运行 `npm run test:regression`。`node --check a.js b.js` 只检查第一个文件，因此要用循环：

   ```bash
   for f in parsing.js profiling.js format.js stats.js quality.js app.js importer.js templates.js report.js script.js; do node --check "$f"; done
   ```

### CSV 中文乱码

1. 手动切换 UTF-8、GBK、GB18030。
2. 检查 `decodeCsvBuffer()` 和浏览器 `TextDecoder` 支持。
3. 自动模式只在 UTF-8 严格解码失败后回退 GB18030，不做统计式编码探测。

### CSV 行列错位或字段丢失

检查 PapaParse 的 `errors`、`TooManyFields`、引号和表头。项目会拒绝可能丢失额外单元格的数据，不应把该错误降级为普通警告。

表头上方有标题行时：第一次解析报出表头类错误（数据列多于表头、空表头下有数据、重复表头）后，会用 `findCsvHeaderBelowTitleRows()` 自动检测表头行，并从表头行切开文本重新解析一次；检测或重试失败时显示第一次的报错，如果第一个非空行像标题，报错后面会附加“第 N 行看起来像标题……”的提示。规则见 `docs/ARCHITECTURE.md` 的“表头行检测”。

### Excel 无法导入

1. 确认扩展名与 ZIP/OLE 签名一致。
2. 检查工作表 `!ref`、行列上限和是否为空。
3. 表头默认是第一个非空行；它通不过表头校验时，会在前 10 行内自动跳过标题行、说明行和空行（见 `docs/ARCHITECTURE.md` 的“表头行检测”）。最终选定的表头必须是唯一、非空的文本表头；识别不出时显示原来的报错，第一个非空行像标题时附加删除标题行的提示。
4. 多工作表问题重点检查 `pendingExcel` 和 `activeImportId` 是否仍有效。

### 筛选后统计不一致

1. 业务统计是否调用 `getAnalysisRows()`。
2. 质量报告是否仍读取 `state.rows` / `qualityNumericStats`。
3. 是否意外修改了行对象或 `state.rows` 数组。
4. 字段类型修改后是否通过 `rebuildAnalysisFromProfiles()` 清除旧筛选。

### 图表重复、空白或内存增长

1. 检查新建 Chart 前是否调用 `destroyChart(id)`。
2. 检查实例是否写入 `state.charts[id]`。
3. 无数据时应调用 `drawEmptyChart()`，不要创建空 Chart。
4. 连续上传后确认 `resetAnalysisState()` 已销毁全部图表。

### 报告导出按钮不可用

`canExportReport()` 要求存在数据、字段画像、`analysisCompletedAt`，且没有未应用的字段类型草稿。先应用或恢复字段配置，再确认筛选刷新已经完成。

### GitHub Pages 上示例数据加载失败

确认 `fetch("examples/sample-data.csv")` 和所有静态资源仍为相对路径，并直接访问部署后的文件 URL。通过本地静态服务器测试，不要只双击 `index.html`。
