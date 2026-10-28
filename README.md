# Smart Tabular Analyzer

> A browser-based CSV and Excel analysis tool for fast, privacy-friendly exploration of structured data.

**Current release:** v2.4.0

## Language

English | [中文](./README.zh-CN.md)

## Live Demo

- **Live Demo:** [Open Smart Tabular Analyzer](https://wy-data-30.github.io/smart-tabular-analyzer/)
- **Repository:** [github.com/wy-data-30/smart-tabular-analyzer](https://github.com/wy-data-30/smart-tabular-analyzer)

Open the demo and choose **使用示例数据** ("Use Sample Data") to complete an analysis without preparing a file.

## About

Smart Tabular Analyzer helps users understand structured files with unfamiliar or inconsistent schemas. After a CSV or Excel file is selected, the application previews the data, infers field types, checks common data quality issues, calculates statistics, renders charts, supports filtered and grouped analysis, produces data-grounded observations, and exports results.

The project is a static website with no backend service, database, account system, or build step. It does not require fixed column names.

## Core Features

- Import CSV, `.xlsx`, and `.xls` files, including multi-sheet workbook selection, Chinese field names and values, and CSV encoding options for Auto Detect, UTF-8, GBK, and GB18030. Excel cells formatted as percentages are shown as percentages.
- Skip title, note, and blank rows above the header automatically, and state in the status bar and reports which row the header was found on.
- Preview the first 10 rows and summarize the dataset size, field composition, missing values, and duplicate rows.
- Infer Numeric, Categorical, Date, and ID fields, with manual correction to Numeric, Categorical, Date, ID, or Ignore when needed.
- Report missing values and rates, complete duplicate rows, IQR-based numeric outliers, and values that cannot be converted to numbers or dates, shown as format problems on the quality page with up to three samples per field and counted in the statistics table.
- Check fields of type ID for duplicate values: the quality page and reports list up to 10 duplicate values per field with how often each appears, the file rows it appears on (the Excel row, or the CSV line where the record starts), and how many of those rows are identical. ID columns whose values mostly repeat, such as a user ID in an event log, get a short note instead of a list.
- Find values such as N/A, NULL, or — that are often written for “no data” (suspected missing values): the quality page lists each one per field with its count, file rows, and how the missing count and the field type would change once it is confirmed. Nothing changes until you confirm; after confirming one value in one field (or all common values at once), it is treated as missing, so it no longer counts as a conversion failure and stays out of statistics, Top 10 lists, groups, filter options, and the ID duplicate check. Every confirmation can be undone. The original data, the data preview, and exported CSV files keep the original text, and reports list which values were treated as missing and which are still unconfirmed.
- Generate descriptive statistics, categorical Top 10 frequencies, numeric distributions (small integer ranges such as ratings are counted per value), and date trends that show record counts by default or the average or sum of a chosen numeric field.
- Apply categorical, numeric, and date filters; run grouped analysis; or use scenario templates for common structured datasets.
- Produce data-grounded observations without inferring external business context; trend observations compare record counts and flag a possibly incomplete last month or year, and the category observation skips fields whose values are too scattered.
- Export filtered or deduplicated rows to CSV in the browser while preserving field order and neutralizing spreadsheet-formula prefixes.
- Export the analysis as HTML or Markdown; reports also record the applied filters, and HTML reports embed the currently rendered charts for offline viewing.

Detailed processing rules and data flows are documented in [Architecture](./docs/ARCHITECTURE.md).

## Screenshots

### Home

Upload CSV or Excel files, or start immediately with the built-in sample data.

![Smart Tabular Analyzer home and file upload](./assets/homepage.png)

### Data Quality

Review missing values, duplicate rows, format problems, IQR-based numeric outliers, and duplicate ID values in one view.

![Smart Tabular Analyzer data quality checks](./assets/data-quality.png)

### Analysis Results

Compare numeric distributions, categorical Top 10 results, and date trends through interactive charts.

![Smart Tabular Analyzer analysis charts](./assets/analysis-results.png)

## Privacy

CSV and Excel parsing, filtering, analysis, processed-data export, and report generation run locally in the browser. The application does not upload or store user files or generated reports. It does not send dataset contents to a third-party API or require an account.

The page loads pinned versions of PapaParse, Chart.js, and SheetJS from the repository's own `vendor/` directory and verifies them with Subresource Integrity (SRI). It does not contact third-party CDN hosts, and uploaded file contents remain in the browser runtime. Because the libraries are hosted with the page, files you import can still be analyzed after the page has loaded, even without a network connection; the built-in sample data is fetched from the site and needs the connection.

## Tech Stack

- HTML5
- CSS3
- Vanilla JavaScript
- [PapaParse 5.4.1](https://www.papaparse.com/) for CSV parsing
- [Chart.js 4.4.1](https://www.chartjs.org/) for data visualization
- [SheetJS 0.20.3](https://sheetjs.com/) for Excel workbook parsing

## Quick Start

### Use the Live Demo

1. Open the [Live Demo](https://wy-data-30.github.io/smart-tabular-analyzer/).
2. Choose **使用示例数据** ("Use Sample Data"), or upload a `.csv`, `.xlsx`, or `.xls` file.
3. Review the inferred field types, correct them only when needed, and use the analysis navigation to explore or export the results.

### Run Locally

Clone the repository and start a static server:

```bash
git clone https://github.com/wy-data-30/smart-tabular-analyzer.git
cd smart-tabular-analyzer
python -m http.server 8000
```

Then open:

```text
http://localhost:8000
```

Opening `index.html` directly also works for file uploads, but a local server is recommended because some browsers restrict the sample-data `fetch()` request on `file://` pages.

## Testing and CI

The automated tests require Node.js 18 or later and use the built-in `node:test` runner without test dependencies.

Run the focused core data-processing suite:

```bash
npm test
```

Run the full import, export, report, boundary, and HTML integration regression suite:

```bash
npm run test:regression
```

[GitHub Actions CI](./.github/workflows/ci.yml) runs the full regression suite with Node.js 22 on every push and pull request. See the [Testing Guide](./docs/TESTING.md) for suite coverage and troubleshooting.

## Documentation

- [Architecture](./docs/ARCHITECTURE.md): architecture, state, import, field inference, filtering, analysis, charts, and export data flows.
- [Testing Guide](./docs/TESTING.md): test commands, suite coverage, fixture guidance, and common failures.
- [Maintenance Guide](./docs/MAINTENANCE.md): core invariants, dependency updates, deployment, and troubleshooting.
- [Release Checklist](./docs/RELEASE_CHECKLIST.md): automated and manual checks used before publication.

### Project Structure

```text
.
├── index.html
├── style.css
├── parsing.js                   # Missing values, number and date parsing
├── profiling.js                 # Field type detection and field profiles
├── format.js                    # Display formatting and HTML escaping
├── stats.js                     # Descriptive stats, IQR, grouping, trends
├── quality.js                   # Conversion failures, duplicates, ID duplicates, missing rates
├── app.js                       # State, DOM map, refresh paths, rendering
├── importer.js                  # CSV / Excel import
├── templates.js                 # Scenario templates
├── report.js                    # Reports and CSV export
├── script.js                    # Startup: self-check and event bindings (loads last)
├── vendor/                      # PapaParse, Chart.js, SheetJS and their licenses
├── examples/
│   ├── sample-data.csv          # Loaded by "Use Sample Data"
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

## Known Limitations

- Each import is limited to 25 MB, 100,000 data rows, and 200 columns.
- Parsing and analysis run on the browser main thread and may temporarily block the interface near the import limits, especially on mobile devices.
- Field inference is heuristic, and strict date parsing may require manual confirmation for ambiguous columns.
- Title, note, and blank rows above the header are skipped automatically; the header row is looked for within 10 rows starting at the first non-empty row. A row can be skipped when it holds a single text cell, or two text cells where one contains a colon (such as “制表日期：…”). When rows are skipped, the status bar and the reports state “表头在第 N 行” (the header is on row N).
- If no header row can be confirmed, the error message suggests removing the title rows; choosing the header row manually is not supported yet.
- Multi-row headers are not supported, total rows such as “合计” are treated as data, and several tables on one worksheet are not supported.
- The header itself must still be unique and non-empty for every column that holds data, and Excel headers must be text cells; empty, headerless, protected, or structurally irregular worksheets may be rejected.
- Excel percentage formats are detected from the first 1,000 data rows below the header; other number formats such as currency are not read.
- ID duplicates compare values after trimming spaces, with case kept. Placeholder text such as N/A or — in an ID column is not treated as missing until it is confirmed on the quality page, so repeated unconfirmed placeholders are listed as duplicate values.
- Suspected missing values come from a fixed list and must match the whole cell (after trimming spaces and folding full-width characters); other text such as “缺考” or “未知” is not listed. Values that contain a zero-width space (U+200B) do not match.
- “无”, “暂无”, “-”, “/”, and “NA” can also be real values, so they are not included when confirming all common values. In accounting formats “-” usually means 0; check what the column means before confirming it.
- Excel formula error cells (#N/A, #DIV/0!, and so on) are imported as missing values directly and do not appear in the suspected missing values list.
- Excel cells are read by their stored value: an ID shown as 001012 through a number format is checked and shown as 1012, and the number 1012 and the text 1012 count as the same ID.
- CSV row numbers count text lines, as a plain-text editor shows them. When a cell contains a line break, Excel or WPS shows the records after it on lower row numbers than the tool reports.
- The project focuses on basic exploratory analysis rather than advanced data cleaning, forecasting, or machine learning.
- Analysis results are not persisted; refreshing the page clears the current analysis.
- The third-party libraries are self-hosted in `vendor/`; upgrading them means replacing those files, their integrity values, and their license files.

## License

This project is licensed under the MIT License. See [LICENSE](./LICENSE).
