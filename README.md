# Sheet Checkup (表格体检)

> A privacy-first checkup for spreadsheet data: CSV and Excel files are checked in your browser, nothing is uploaded, and you see what is wrong and which rows it is in.

Formerly Smart Tabular Analyzer. The repository name and URLs are unchanged.

**Current release:** v3.1.0

## Language

English | [中文](./README.zh-CN.md)

## Live Demo

- **Live Demo:** [Open Sheet Checkup](https://wy-data-30.github.io/smart-tabular-analyzer/?lang=en)
- **Repository:** [github.com/wy-data-30/smart-tabular-analyzer](https://github.com/wy-data-30/smart-tabular-analyzer)

Open the demo and choose **Use sample data** to complete an analysis without preparing a file. The interface is in English or Chinese (see Interface Language below); in English the button loads `examples/sample-orders-en.csv`, in Chinese it loads `examples/sample-orders.csv`.

The sample is a fictional order list for the first half of 2025 with a few common problems put in on purpose — a title row above the header, a total row, a duplicated order, "N/A" and "—" placeholders, amounts written in a different way ("$5,400.00" in the English sample, "5400元" in the Chinese one), mixed date formats, and leading or trailing spaces — so the Data quality page has something to show. The two samples hold the same orders and the same numbers. The total row is still counted, which inflates the averages and the distribution chart. In the English sample the "$" amounts are read as plain numbers, so the Data quality page does not point them out; only the standardized export rewrites them.

## Interface Language

The interface is available in English and Chinese. The language follows the browser's first language (Chinese for any `zh` language, English otherwise); `?lang=en` or `?lang=zh` in the address takes precedence. Switch with the "中文" / "English" button at the top right of the upload page, or at the foot of the sidebar after a file is loaded; the switch keeps everything on the page (filters, rows left out of the analysis, the open module, chart and breakdown choices, the problem rows and their page, unapplied column settings) and writes the language into the address. Nothing is stored.

Changing the language changes the words only: the same file gives the same numbers, rows, types, problem rows, order, and charts in both languages. Reports, exported file names, and the columns the tool adds to CSV files follow the language at the moment of export.

## About

Sheet Checkup helps users understand structured files with unfamiliar or inconsistent schemas. After a CSV or Excel file is selected, the application previews the data, infers field types, checks common data quality issues, calculates statistics, renders charts, supports filtered and grouped analysis, produces data-grounded observations, and exports results.

The project is a static website with no backend service, database, account system, or build step. It does not require fixed column names.

## Core Features

- Import CSV, `.xlsx`, and `.xls` files, including multi-sheet workbook selection, Chinese field names and values, and CSV encoding options for Auto Detect, UTF-8, GBK, and GB18030. Excel cells formatted as percentages are shown as percentages.
- Skip title, note, and blank rows above the header automatically, and state in the status bar and reports which row the header was found on.
- Describe CSV parse warnings (such as a row with fewer cells than the header) in the interface language with the file line they are on; the quality page shows how many warnings there are and how many lines they cover, and the “more data columns than the header” and “unclosed quote” errors name their line too. A file separated by semicolons, tabs, or bars that comes out as a single column gets a hint.
- Preview the data page by page (50 rows a page, with the file row in the first column, following the filters) and summarize the dataset size, field composition, missing values, and duplicate rows.
- Infer Numeric, Categorical, Date, and ID fields, with manual correction to Numeric, Categorical, Date, ID, or Ignore when needed.
- Report missing values and rates, complete duplicate rows, IQR-based numeric outliers, and values that cannot be converted to numbers or dates, shown as format problems on the quality page with up to three samples per field and counted in the statistics table.
- Check fields of type ID for duplicate values: the quality page and reports list up to 10 duplicate values per field with how often each appears, the file rows it appears on (the Excel row, or the CSV line where the record starts), and how many of those rows are identical. ID columns whose values mostly repeat, such as a user ID in an event log, get a short note instead of a list.
- Find values such as N/A, NULL, or — that are often written for “no data” (suspected missing values): the quality page lists each one per field with its count, file rows, and how the missing count and the field type would change once it is confirmed. Nothing changes until you confirm; after confirming one value in one field (or all common values at once), it is treated as missing, so it no longer counts as a conversion failure and stays out of statistics, Top 10 lists, groups, filter options, and the ID duplicate check. Every confirmation can be undone. The original data, the data preview, and exported CSV files keep the original text, and reports list which values were treated as missing and which are still unconfirmed.
- Read numbers written with a unit, 万 / 亿, or full-width digits, column by column: a column of amounts such as 2200元 is read in 元 (yuan), 1.2万 is read as 12000 and 3亿 as 300000000, and full-width digits such as １２３ are read as 123. The recognized units are 元; kg, 克, 千克, 公斤, 斤, 吨; 件, 个, 箱, 台, 人, 次; 秒, 分钟, 小时, 天; 米, cm, mm, km; and 岁 (Latin letters in any case). Statistics, chart titles and axes, and reports show the unit after the field name, such as “金额（元）”; the data preview and exported CSV files keep the original text. The quality page lists which fields were read this way (“按单位读取的数字”, numbers read by unit) and points out category fields where most values, but not enough for a numeric field, can be read as numbers (“像是数字的文本列”, text columns that look numeric), with up to three values that could not be read.
- Flag suspected total and subtotal rows: a row with a label such as 合计, 小计, 汇总, or Total whose number in a numeric column equals the sum of the detail rows it covers (within 0.5%) is listed on the quality page and in reports, with the columns that matched and the rows they add up to. The quality page and reports list up to 20 such rows. By default these rows still count in sums, averages, and outliers; they can be left out of the analysis on the quality page (next item).
- Leave suspected totals and the extra rows of complete duplicates out of the analysis: the quality page has a button for each kind, and either can be undone; these rows then stay out of the statistics, charts, custom analysis, and filters, and every place says how many rows are left out. The quality page, the field types, and the standardized export still use every row of the file. The statistics, charts, and custom analysis modules say at the top how many such rows their results still contain, with a button to go and handle them on the quality page.
- List Excel formula error cells (#DIV/0!, #N/A, #REF!, #SPILL!, and so on) per field, with how many of each error and their file rows, on the quality page and in reports. They still count as missing values; the list of a field's missing rows says which ones were errors.
- Flag format inconsistencies on the quality page and in reports, without changing any result: a date column that mixes styles (Excel date cells, 2024-05-06, 2024/5/6, 2024.5.6, 2024年5月6日, and so on) with the rows not in the main style; constant columns whose non-missing values are all the same; and values written in more than one way — leading or trailing spaces, cells holding only invisible characters, and writings that differ only in full or half width, case, zero-width characters, or how many spaces sit between words (such as HD, ＨＤ, and hd). The dates and the values written in more than one way have “查看” (view) to list their rows.
- See the problem rows: next to each count on the quality page, “查看” (view) lists the rows behind it — missing values, formula errors, and format problems per field, outliers per numeric field, complete duplicate rows by group (the first row of each group marked as kept), suspected total rows, every row of every duplicate ID value, suspected missing values per field and value (confirmed ones too), and CSV parse warnings. Each row shows its file row, a short note, and all original columns, with the problem cell highlighted, 50 rows a page, and a list to switch categories; “导出这一类” (export this category) writes every row of the category to CSV, with “文件行号” (file row) and “说明” (note) as the first two columns and the same BOM and formula-prefix neutralizing. The listed rows come from the same checks as the quality counts, use the original data, and do not follow the filters; rows left out of the analysis end their note with “（不计入分析）” (left out), in the export of a category too.
- Generate descriptive statistics, categorical Top 10 frequencies, numeric distributions (small integer ranges such as ratings are counted per value), and date trends that show record counts by default or the average or sum of a chosen numeric field.
- Apply categorical, numeric, and date filters, and run grouped analysis (one metric, one grouping field, optionally a date field).
- Show one data-grounded conclusion under the category chart and one under the date trend chart, each following the field selected in that chart and without inferring external business context; the trend conclusion compares record counts and flags a possibly incomplete last month or year, and the category conclusion says so instead of summarizing a field whose values are too scattered. Reports carry the same sentences.
- Export the rows of the current analysis (the filtered rows, without the rows left out) to CSV in the browser, with a file name ending in “_当前分析_<date>” (current analysis), while preserving field order and neutralizing spreadsheet-formula prefixes.
- Export with unified writing: first click “检查会改哪些格” (check which cells will change) to see, field by field, how many cells would be rewritten and which stay as written; then export every original row with dates written as YYYY-MM-DD, numbers without units, 万/亿, full-width characters, thousands separators, or currency signs, values confirmed as blank emptied, and edge whitespace removed from text. A change list can be exported too, naming every cell rewritten or kept as written (file row, field, original value, exported value, action).
- Export the analysis as an HTML report (print it to PDF from the browser); the report records the applied filters and embeds the currently rendered charts for offline viewing.

Detailed processing rules and data flows are documented in [Architecture](./docs/ARCHITECTURE.md).

## Screenshots

### Home

Upload CSV or Excel files, or start immediately with the built-in sample data.

![Sheet Checkup home and file upload](./assets/homepage.png)

*Screenshot of the previous interface.*

### Data Quality

Review missing values, duplicate rows, format problems, IQR-based numeric outliers, and duplicate ID values in one view.

![Sheet Checkup data quality checks](./assets/data-quality.png)

*Screenshot of the previous interface.*

### Analysis Results

Compare numeric distributions, categorical Top 10 results, and date trends through interactive charts.

![Sheet Checkup analysis charts](./assets/analysis-results.png)

*Screenshot of the previous interface.*

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

1. Open the [Live Demo](https://wy-data-30.github.io/smart-tabular-analyzer/?lang=en).
2. Choose **Use sample data**, or upload a `.csv`, `.xlsx`, or `.xls` file.
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
├── format.js                    # Display formatting, HTML escaping, interface language (L(), I18N_EN)
├── stats.js                     # Descriptive stats, IQR, grouping, trends
├── quality.js                   # Conversion failures, duplicates, ID duplicates, suspected totals, formula errors, missing rates, problem rows
├── app.js                       # State, DOM map, refresh paths, rendering
├── importer.js                  # CSV / Excel import
├── report.js                    # Reports, CSV export and problem-row export
├── script.js                    # Startup: self-check and event bindings (loads last)
├── vendor/                      # PapaParse, Chart.js, SheetJS and their licenses
├── examples/
│   ├── sample-orders-en.csv     # Loaded by "Use sample data" in English
│   └── sample-orders.csv        # Loaded by "使用示例数据" in Chinese
├── assets/
├── docs/
├── tests/                       # Tests and fixtures (tests/fixtures/examples/ holds the other sample CSVs, including the former sample-data.csv)
└── package.json
```

## Known Limitations

- Each import is limited to 25 MB, 100,000 data rows, and 200 columns.
- Parsing and analysis run on the browser main thread and may temporarily block the interface near the import limits, especially on mobile devices.
- Field inference is heuristic, and strict date parsing may require manual confirmation for ambiguous columns.
- Title, note, and blank rows above the header are skipped automatically; the header row is looked for within 10 rows starting at the first non-empty row. A row can be skipped when it holds a single text cell, or two text cells where one contains a colon (such as “制表日期：…”). When rows are skipped, the status bar and the reports state “表头在第 N 行” (the header is on row N).
- If no header row can be confirmed, the error message suggests removing the title rows; choosing the header row manually is not supported yet.
- Multi-row headers are not supported, and several tables on one worksheet are not supported.
- Suspected total rows still count in the statistics by default and can be left out of the analysis on the quality page. The button records the rows flagged at that moment: after a field configuration change or a value treated as empty, rows no longer flagged count again (the message says so). Newly flagged rows are not left out automatically; when a row that was counted again is flagged again, the message says it can be left out again. Importing again counts every row. A label must be the whole cell, after removing spaces, trailing colons, and one trailing note in parentheses: up to 12 characters followed by 合计, 总计, 小计, 总和, 汇总, 累计 (or 合計, 總計, 小計, 總和, 匯總, 彙總, 累計), or Total, Grand Total, Subtotal, Sub-total, Sum in any case. Not recognized: 共计, 合共, （合计）, 合计（元）（含税）, 合计[元], 合计【元】; a prefix that is only a status word (已, 未, 不, 无, 無, 非, 待, 是否, 尚未, 暂未, 暂不, 暫未, 暫不), such as 已汇总; Sub-total written with the hyphen U+2010; and a zero-width space in the middle of the word.
- A detail row that happens to carry a label can still hide a total: labelled rows that do not add up get one more check, only one, so a table whose first detail row has a label may flag its subtotals but miss the grand total.
- Some suspected totals only equal the one row above them (a subtotal of a single detail row, or a row labelled 汇总 in an hours or quantity column with few distinct values that happens to equal the row above), and may not be totals at all. The quality page and reports say “其中 N 行只等于上面一行，可能不是合计行，先点“查看”确认。” (N of them only equal the row above and may not be totals; check them with “查看” first); leaving the totals out leaves these out too, so check them first.
- The quality page always uses every row of the file, including rows left out of the analysis: once rows are left out, the statistics and the outlier bounds differ from those on the quality page.
- Without an ID column that tells the rows apart (a column of type ID), identical rows may still be different records, so the duplicate-row line asks to check them with “查看重复行” (view duplicate rows) first. This check is loose: an ID column with repeated values still counts as telling the rows apart as long as at least half of its values are distinct, and the duplicate-row line then no longer asks to check first.
- The header itself must still be unique and non-empty for every column that holds data, and Excel headers must be text cells; empty, headerless, protected, or structurally irregular worksheets may be rejected.
- Excel percentage formats are detected from the first 1,000 data rows below the header; other number formats such as currency are not read.
- ID duplicates compare values after trimming spaces, with case kept. Placeholder text such as N/A or — in an ID column is not treated as missing until it is confirmed on the quality page, so repeated unconfirmed placeholders are listed as duplicate values.
- Suspected missing values come from a fixed list and must match the whole cell (after trimming spaces and folding full-width characters); other text such as “缺考” or “未知” is not listed. Values that contain a zero-width space (U+200B) do not match.
- “无”, “暂无”, “-”, “/”, and “NA” can also be real values, so they are not included when confirming all common values. In accounting formats “-” usually means 0; check what the column means before confirming it.
- Excel formula error cells (#N/A, #DIV/0!, and so on) are imported as missing values and listed under “公式错误单元格” (formula error cells), not among suspected missing values. Not listed: rows where every cell is an error (they are not imported), columns with no header and only errors, and errors in the header or title rows. In a CSV, “#DIV/0!” is plain text: it counts as a format problem in a numeric column, and #N/A is a suspected missing value.
- Numbers with units must be written as a whole cell of number, optional 万 / 亿, and optional unit. ￥2200元 (a currency sign and a unit), (2200元) (an accounting negative with a unit), 1.2 万 (a space before 万), and 2200元/件 are not read and still count as format problems. g, m, and 分 are not recognized as units, because they are ambiguous (4G / 5G, meters or minutes, points or cents).
- When one column mixes units, such as kg and 千克, or kg and 克, nothing is converted between them and the whole column is not read by unit; the quality page says which units appeared and how often.
- A field named in 万元 whose values are written like 1.2万 is still read in 元, as 12000.
- In accounting formats “-” usually means 0, but it is still listed as a suspected missing value and never read as 0.
- Excel cells are read by their stored value: an ID shown as 001012 through a number format is checked and shown as 1012, and the number 1012 and the text 1012 count as the same ID.
- Format inconsistencies are only hints. Writings that differ in more than width, case, invisible characters, or spacing are not grouped: 华东 and 华东区, simplified and traditional characters (華东 and 华东), and 张 三 and 张三 (a space removed altogether) count as different values. Zero padding and a time part do not make a new date style (2024-5-6 and 2024-05-06 10:30 are both year-month-day). A numeric column with 1 and 1.0 is not a constant column. Case-sensitive IDs such as AbC and aBc are reported as written in more than one way.
- An Excel cell that holds only a time (formatted hh:mm) is a number cell and reads as the date 1899-12-31.
- An Excel date cell with a time (formatted yyyy-mm-dd hh:mm, for example) imports as its date only; the time is in neither the analysis nor any CSV export.
- The standardized export does not rewrite percentages (50%) or scientific notation (1e3), and keeps dates written with a time (2024-05-06 10:30, for example) as they are; zero-width characters, full-width letters, and letter case are not unified.
- After the standardized export, numbers read with a unit are plain numbers and the unit is not added to the header (“金额” stays “金额”, not “金额（元）”).
- The weekday of an English date is not checked: Tue, 6 May 2024 reads as 2024-05-06 (a Monday), and the standardized export writes only the date, 2024-05-06.
- When a date column’s day-month order rests only on cells that the standardized export rewrites, the day-or-month-first cells written with a time lose what decided their order, and importing the exported file into this tool again may not read them as dates; the check table notes this under the column.
- The change list holds at most 1,048,575 cells (the rows of an Excel sheet minus the header). Beyond that only the first 1,048,575 are exported, and the file name and the status line say so; the check table still counts every cell.
- When Excel opens an exported CSV directly, it drops leading zeros such as in 007 and loses the precision of numbers longer than 15 digits; the tool writes the original text.
- CSV row numbers count text lines, as a plain-text editor shows them; data rows, parse warnings, and errors all name the line where the record starts. When a cell contains a line break, Excel or WPS shows the records after it on lower row numbers than the tool reports.
- Problem rows can be viewed and exported on the quality page only; the report does not include the problem-row tables, and there is no view of all problem rows together.
- The English interface changes the words only. Tables are still read by the same rules in both languages: total labels, units, and values that may mean blank come from the lists above (mostly Chinese, plus Total, Subtotal, Sum and the Latin units); reading English tables more fully is planned for the next version.
- Amounts written with "$" and thousands separators (such as $5,400.00) are read as plain numbers, so they are not pointed out on the Data quality page; the standardized export writes them without the sign and separators.
- File names are cut to 80 characters before the date: only the original file name and the user's column or value in it are shortened, never the words the tool adds, so two exports of one file never get the same name.
- The project focuses on basic exploratory analysis rather than advanced data cleaning, forecasting, or machine learning.
- Analysis results are not persisted; refreshing the page clears the current analysis.
- The third-party libraries are self-hosted in `vendor/`; upgrading them means replacing those files, their integrity values, and their license files.

## License

This project is licensed under the MIT License. See [LICENSE](./LICENSE).
