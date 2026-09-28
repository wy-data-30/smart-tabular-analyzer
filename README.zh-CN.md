# Smart Tabular Analyzer

> 一款在浏览器中完成 CSV 和 Excel 快速探索分析的隐私友好型工具。

**当前版本：** v2.1.2

## 语言

[English](./README.md) | 中文

## 在线演示

- **在线演示：** [打开 Smart Tabular Analyzer](https://wy-data-30.github.io/smart-tabular-analyzer/)
- **GitHub 仓库：** [github.com/wy-data-30/smart-tabular-analyzer](https://github.com/wy-data-30/smart-tabular-analyzer)

打开在线演示并点击 **使用示例数据**，无需准备文件即可完成一次分析。

## 项目简介

Smart Tabular Analyzer 用于帮助用户理解字段结构陌生或不一致的结构化数据文件。选择 CSV 或 Excel 文件后，应用会完成数据预览、字段类型识别、常见数据质量检查、统计计算、图表生成、筛选与分组分析、基于已有数据的基础结论生成以及结果导出。

该项目是静态网站，不依赖后端服务、数据库、账号系统或构建流程，也不要求固定字段名。

## 核心功能

- 导入 CSV、`.xlsx` 和 `.xls` 文件，支持多工作表选择、中文字段名与内容，以及自动识别、UTF-8、GBK 和 GB18030 等 CSV 编码选项。
- 预览前 10 行数据，并汇总数据规模、字段构成、缺失值和重复行。
- 自动识别数值、分类、日期和 ID 字段，并可在需要时人工修正为数值、分类、日期、ID 或忽略。
- 报告缺失值及缺失率、完全重复行、基于 IQR 的数值异常值，以及数值或日期转换失败。
- 生成描述性统计、分类 Top 10 频数、数值分布和日期趋势。
- 应用分类、数值和日期筛选，进行分组分析，或使用适配常见结构化数据的场景模板。
- 基于已有数据生成基础结论，不推断外部业务背景。
- 在浏览器中导出筛选后或去重后的 CSV，保留字段顺序并中和电子表格公式前缀。
- 将分析导出为 HTML 或 Markdown；HTML 报告会嵌入当前已渲染的图表，可离线查看。

详细处理规则和数据流请参阅[架构说明](./docs/ARCHITECTURE.md)。

## 项目截图

### 首页

上传 CSV 或 Excel 文件，或直接使用内置示例数据开始分析。

![Smart Tabular Analyzer 首页与文件上传](./assets/homepage.png)

### 数据质量

在同一视图中检查缺失值、重复行和基于 IQR 的数值异常值。

![Smart Tabular Analyzer 数据质量检查](./assets/data-quality.png)

### 分析结果

通过交互图表查看数值分布、分类 Top 10 和日期趋势。

![Smart Tabular Analyzer 分析图表](./assets/analysis-results.png)

## 隐私说明

CSV 和 Excel 解析、筛选、分析、处理后数据导出及报告生成均在浏览器本地完成。应用不会上传或存储用户文件及生成的报告，也不会将数据集内容发送给第三方 API 或要求用户注册账号。

页面会从 CDN 加载固定版本的 PapaParse、Chart.js 和 SheetJS，并通过子资源完整性（SRI）校验。浏览器加载页面时会访问这些 CDN 域名，但上传文件的内容仍保留在浏览器运行环境中。

## 技术栈

- HTML5
- CSS3
- Vanilla JavaScript
- [PapaParse 5.4.1](https://www.papaparse.com/)：CSV 解析
- [Chart.js 4.4.1](https://www.chartjs.org/)：数据可视化
- [SheetJS 0.20.3](https://sheetjs.com/)：Excel 工作簿解析

## 快速开始

### 使用在线演示

1. 打开[在线演示](https://wy-data-30.github.io/smart-tabular-analyzer/)。
2. 点击 **使用示例数据**，或上传 `.csv`、`.xlsx` 或 `.xls` 文件。
3. 查看自动识别的字段类型，仅在需要时进行修正，然后通过分析导航查看或导出结果。

### 本地运行

克隆仓库并启动静态服务器：

```bash
git clone https://github.com/wy-data-30/smart-tabular-analyzer.git
cd smart-tabular-analyzer
python -m http.server 8000
```

然后打开：

```text
http://localhost:8000
```

直接打开 `index.html` 也可以上传文件，但更推荐使用本地服务器，因为部分浏览器会限制 `file://` 页面中的示例数据 `fetch()` 请求。

## 测试与 CI

自动化测试需要 Node.js 18 或更高版本，使用内置的 `node:test`，不需要安装测试依赖。

运行专注于核心数据处理的测试：

```bash
npm test
```

运行覆盖导入、导出、报告、边界和 HTML 集成的完整回归测试：

```bash
npm run test:regression
```

[GitHub Actions CI](./.github/workflows/ci.yml) 会在每次 push 和 Pull Request 时使用 Node.js 22 运行完整回归测试。测试范围与故障排查请参阅[测试指南](./docs/TESTING.md)。

## 项目文档

- [架构说明](./docs/ARCHITECTURE.md)：架构、状态、导入、字段识别、筛选、分析、图表和导出数据流。
- [测试指南](./docs/TESTING.md)：测试命令、覆盖范围、夹具说明和常见失败原因。
- [维护指南](./docs/MAINTENANCE.md)：核心约束、依赖更新、部署和故障排查。
- [发布检查清单](./docs/RELEASE_CHECKLIST.md)：发布前使用的自动化与人工检查。

### 项目结构

```text
.
├── index.html
├── style.css
├── script.js
├── examples/
│   ├── sample-data.csv          # “使用示例数据”加载的文件
│   ├── sample-sales.csv
│   ├── sample-students.csv
│   ├── sample-survey.csv
│   ├── sample-used-products.csv
│   └── sample-user-behavior.csv
├── assets/
├── docs/
├── tests/
└── package.json
```

## 已知限制

- 单次导入限制为 25 MB、100,000 行数据和 200 列。
- 解析和分析在浏览器主线程中运行；接近导入上限时可能暂时阻塞界面，移动设备上的影响更明显。
- 字段识别基于启发式规则，严格日期解析可能要求用户人工确认含义不明确的字段。
- Excel 工作表的第一个非空行必须包含唯一、非空的文本表头；空表、无表头、受保护或结构异常的工作表可能被拒绝。
- 项目聚焦基础探索式分析，不包含高级数据清洗、预测或机器学习。
- 分析结果不会持久化；刷新页面会清除当前分析。
- CDN 依赖需要网络访问，除非将相关库改为本地托管。

## 许可证

本项目基于 MIT License 开源。详情请查看 [LICENSE](./LICENSE)。
