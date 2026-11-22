// v3.2.0 static guard for the English interface: every Chinese text shown to the user goes through
// L(中文, English) or Lp(中文, English), and every Chinese text in index.html has a data-i18n key.
// Used by basic-data-processing.test.cjs; no dependencies.
//
// The tokenizer is enough for the nine application scripts: comments, '…' and "…" strings,
// template literals with nested ${ … } (any depth), regex literals, identifiers, numbers and
// punctuation. Every string-like token (str, tmpl = one piece of a template, regex) carries:
//   line   – the line of its first character
//   decl   – the top-level declaration it sits in (function NAME / const|let|var NAME / class
//            NAME), "" for top-level statements outside any declaration
//   lFrame – null outside a call of L / Lp; otherwise { arg, helper } of the NEAREST enclosing
//            call (arg 0 = Chinese, 1 = English)
// Regex or division: "/" starts a regex unless the previous significant token is an identifier
// (other than a keyword such as return or typeof), a number, a string, a template end, ")" or "]".
// The self-check (every Han character of a file lies in a comment or a literal) catches the
// tokenizer losing its place.
"use strict";

const KEYWORDS_BEFORE_EXPR = new Set(["return", "typeof", "case", "do", "else", "in", "of", "new", "delete", "void", "throw", "instanceof", "yield", "await"]);
const isIdStart = (c) => /[A-Za-z_$\u00C0-\uFFFF]/.test(c);
const isId = (c) => /[A-Za-z0-9_$\u00C0-\uFFFF]/.test(c);
const HELPERS = ["L", "Lp"];

function lex(src, helperNames = HELPERS) {
  const helpers = new Set(helperNames);
  const tokens = [];
  let i = 0;
  let line = 1;
  const stack = []; // { open, callee, arg }
  let prev = null;
  let decl = "";
  let pendingDecl = false;

  const lFrame = () => {
    for (let k = stack.length - 1; k >= 0; k -= 1) {
      if (helpers.has(stack[k].callee) && stack[k].open === "(") return { arg: stack[k].arg, helper: stack[k].callee };
    }
    return null;
  };
  const push = (type, value, startLine, extra = {}) => {
    const token = { type, value, line: startLine, decl, ...extra };
    if (type === "str" || type === "tmpl" || type === "regex") token.lFrame = lFrame();
    tokens.push(token);
    if (type !== "comment") prev = token;
    return token;
  };
  const advance = (n = 1) => {
    for (let k = 0; k < n; k += 1) {
      if (src[i] === "\n") line += 1;
      i += 1;
    }
  };
  const regexAllowed = () => {
    if (!prev) return true;
    if (prev.type === "num" || prev.type === "str" || prev.type === "regex") return false;
    if (prev.type === "tmpl") return !prev.endsTemplate;
    if (prev.type === "ident") return KEYWORDS_BEFORE_EXPR.has(prev.value);
    if (prev.type === "punct") return !(prev.value === ")" || prev.value === "]");
    return true;
  };

  function readTemplate() {
    const startLine = line;
    advance();
    let chunk = "";
    let chunkLine = line;
    for (;;) {
      if (i >= src.length) throw new Error(`unterminated template from line ${startLine}`);
      const c = src[i];
      if (c === "\\") { chunk += src.slice(i, i + 2); advance(2); continue; }
      if (c === "`") {
        push("tmpl", chunk, chunkLine, { endsTemplate: true });
        advance();
        return;
      }
      if (c === "$" && src[i + 1] === "{") {
        push("tmpl", chunk, chunkLine, { endsTemplate: false });
        advance(2);
        stack.push({ open: "${", callee: null, arg: 0 });
        prev = null;
        readUntil(stack.length);
        chunk = "";
        chunkLine = line;
        continue;
      }
      chunk += c;
      advance();
    }
  }

  // Tokens until the bracket frame at height `base` closes (or the end, for base 0).
  function readUntil(base) {
    while (i < src.length) {
      const c = src[i];
      const n = src[i + 1];
      if (c === "\n" || c === " " || c === "\t" || c === "\r" || c === "\uFEFF") { advance(); continue; }
      if (c === "/" && n === "/") {
        const s = i; const l = line;
        while (i < src.length && src[i] !== "\n") advance();
        push("comment", src.slice(s, i), l);
        continue;
      }
      if (c === "/" && n === "*") {
        const s = i; const l = line;
        advance(2);
        while (i < src.length && !(src[i] === "*" && src[i + 1] === "/")) advance();
        advance(2);
        push("comment", src.slice(s, i), l);
        continue;
      }
      if (c === "'" || c === '"') {
        const l = line; let v = "";
        advance();
        while (i < src.length && src[i] !== c) {
          if (src[i] === "\\") { v += src.slice(i, i + 2); advance(2); continue; }
          if (src[i] === "\n") throw new Error(`newline in string at line ${line}`);
          v += src[i]; advance();
        }
        advance();
        push("str", v, l);
        continue;
      }
      if (c === "`") { readTemplate(); continue; }
      if (c === "/" && regexAllowed()) {
        const l = line; let v = ""; let inClass = false;
        advance();
        while (i < src.length) {
          const d = src[i];
          if (d === "\\") { v += src.slice(i, i + 2); advance(2); continue; }
          if (d === "[") inClass = true;
          else if (d === "]") inClass = false;
          else if (d === "/" && !inClass) break;
          else if (d === "\n") throw new Error(`newline in regex at line ${line}`);
          v += d; advance();
        }
        advance();
        let flags = "";
        while (i < src.length && /[a-z]/.test(src[i])) { flags += src[i]; advance(); }
        push("regex", v, l, { flags });
        continue;
      }
      if (/[0-9]/.test(c) || (c === "." && /[0-9]/.test(n))) {
        const s = i;
        while (i < src.length && /[0-9A-Za-z_.]/.test(src[i])) advance();
        push("num", src.slice(s, i), line);
        continue;
      }
      if (isIdStart(c)) {
        const s = i;
        while (i < src.length && isId(src[i])) advance();
        const word = src.slice(s, i);
        if (stack.length === 0) {
          if (pendingDecl && word !== "async" && word !== "function") { decl = word; pendingDecl = false; }
          else if (["function", "const", "let", "var", "class"].includes(word)) pendingDecl = true;
        }
        push("ident", word, line);
        continue;
      }
      if (c === "(" || c === "[" || c === "{") {
        const callee = c === "(" && prev && prev.type === "ident" ? prev.value : null;
        stack.push({ open: c, callee, arg: 0 });
        push("punct", c, line);
        advance();
        continue;
      }
      if (c === ")" || c === "]" || c === "}") {
        const frame = stack.pop();
        if (!frame) throw new Error(`unbalanced ${c} at line ${line}`);
        if (frame.open === "${") {
          advance();
          if (stack.length < base) return;
          continue;
        }
        push("punct", c, line);
        advance();
        if (stack.length < base) return;
        continue;
      }
      if (c === ",") {
        if (stack.length) stack[stack.length - 1].arg += 1;
        push("punct", c, line); advance();
        continue;
      }
      if (c === ";" && stack.length === 0) pendingDecl = false;
      push("punct", c, line);
      advance();
    }
  }

  readUntil(0);
  if (stack.length) throw new Error(`unclosed ${stack.map((f) => f.open).join("")} at end`);
  return tokens;
}

const HAN = /\p{Script=Han}/u;
const HAN_ALL = /\p{Script=Han}/gu;
// Full-width and CJK punctuation that must not appear in English text. Curly quotes, … and — are
// English punctuation too, so they are allowed there.
const CJK_PUNCT_STRICT = /[\u3000-\u303F\uFF01-\uFF0F\uFF1A-\uFF20\uFF3B-\uFF40\uFF5B-\uFF65]/u;

const APP_SCRIPT_NAMES = ["parsing.js", "profiling.js", "format.js", "stats.js", "quality.js", "app.js", "importer.js", "report.js", "script.js"];

// Literals with Chinese that are data or detection rules, not interface text: file + top-level
// declaration + how many such literals it holds (a new Chinese text inside one of them fails).
const I18N_WHITELIST = Object.freeze([
  ["parsing.js", "isYearOnlyDateValue", 1, "只写年份的日期 2025年"],
  ["parsing.js", "PLACEHOLDER_VALUES", 2, "空值写法 无 / 暂无"],
  ["parsing.js", "TOTAL_LABEL_CHINESE_PATTERN", 1, "合计字样"],
  ["parsing.js", "TOTAL_LABEL_STATUS_PREFIXES", 13, "合计字样的状态前缀"],
  ["parsing.js", "TOTAL_LABEL_PREFILTER", 1, "合计预筛"],
  ["parsing.js", "matchChineseDate", 1, "中文日期"],
  ["parsing.js", "FULLWIDTH_NUMBER_CHARS", 5, "全角数字标点"],
  ["parsing.js", "toHalfwidthNumberText", 1, "全角数字标点"],
  ["parsing.js", "NUMBER_UNITS", 18, "中文单位"],
  ["parsing.js", "NUMBER_WITH_UNIT_PATTERN", 1, "万 / 亿"],
  ["parsing.js", "toNumberWithStrategy", 1, "万"],
  ["parsing.js", "toDate", 1, "只写年份"],
  ["parsing.js", "dateStyleOf", 1, "只写年份"],
  ["parsing.js", "standardizeNumberCell", 1, "万"],
  ["profiling.js", "isIdFieldName", 3, "编号字段名"],
  ["profiling.js", "isYearFieldName", 1, "年份字段名"],
  ["profiling.js", "buildDateStrategy", 1, "只写年份"],
  ["importer.js", "isSkippableHeaderRow", 1, "说明：两格行"],
  ["format.js", "LANGUAGE_NAME_ZH", 1, "切换按钮上的“中文”"],
  ["format.js", "formatFieldLabelWithUnit", 2, "字段名结尾已有单位的识别规则：【单位】"],
  ["stats.js", "BLANK_GROUP_KEY", 1, "空值分组的键（数据，两种语言相同；显示用 displayGroupName）"],
  ["importer.js", "FORMULA_ERROR_FALLBACK", 1, "认不出的公式错误写成的值（数据，原样显示）"],
  ["script.js", "ensureDomElementsPresent", 4, "启动自检失败的提示：format.js 可能没加载，不能用 L()；语言看 data-lang-pending"]
]);

// The files whose interface text is all in L() / Lp() (rule 1 applies). Only grows; when all nine
// are here the list goes and rule 1 applies to every file.
const I18N_DONE_FILES = Object.freeze(["format.js", "app.js", "importer.js", "stats.js", "quality.js", "profiling.js", "script.js", "report.js", "parsing.js"]);

const isChineseLiteral = (token) => (token.type === "str" || token.type === "tmpl" || token.type === "regex") && (HAN.test(token.value) || CJK_PUNCT_STRICT.test(token.value));

// Han characters outside comments and literals: the tokenizer lost its place.
function lexerSelfCheck(source, tokens = lex(source)) {
  const all = (source.match(HAN_ALL) || []).length;
  const inTokens = tokens
    .filter((token) => ["comment", "str", "tmpl", "regex"].includes(token.type))
    .reduce((sum, token) => sum + (token.value.match(HAN_ALL) || []).length, 0);
  return { all, inTokens };
}

// Rules 1–3 for one file. `whitelist` maps "file|decl" to the number of literals allowed.
function checkScriptTokens(file, tokens, whitelist, { rule1 = true } = {}) {
  const problems = [];
  const used = new Map();
  for (const token of tokens) {
    if (!isChineseLiteral(token)) continue;
    const han = HAN.test(token.value);
    if (token.lFrame && token.lFrame.arg === 0) continue;
    if (token.lFrame && token.lFrame.arg >= 1) {
      // Rule 2: the English argument has no Han and no full-width punctuation.
      problems.push({ file, line: token.line, decl: token.decl, kind: han ? "Han in the English argument" : "full-width punctuation in the English argument", text: token.value });
      continue;
    }
    const key = `${file}|${token.decl}`;
    if (whitelist.has(key)) {
      used.set(key, (used.get(key) || 0) + 1);
      continue;
    }
    // Rule 1: Chinese outside L() / Lp().
    if (rule1) problems.push({ file, line: token.line, decl: token.decl, kind: han ? "Chinese outside L()" : "full-width punctuation outside L()", text: token.value });
  }
  // Rule 3: no more literals in a whitelisted declaration than it allows.
  for (const [key, count] of used) {
    if (count > whitelist.get(key)) problems.push({ file: key.split("|")[0], line: 0, decl: key.split("|")[1], kind: `whitelisted declaration has ${count} literals, ${whitelist.get(key)} allowed`, text: "" });
  }
  return { problems, used };
}

// index.html: a small tag scanner. Text belongs to the innermost open tag. `within(offset)` says
// which parts are checked (v3.2.0 commit 4: only the blocks that have keys so far).
function checkIndexHtml(html, english, within = () => true) {
  const source = html.replace(/<!--[\s\S]*?-->/g, (comment) => comment.replace(/[^\n]/g, " "));
  const problems = [];
  const keys = new Set();
  const stack = [];
  const lineAt = (offset) => source.slice(0, offset).split("\n").length;
  const tagPattern = /<\/?([a-zA-Z][\w-]*)((?:\s+[^\s=>]+(?:="[^"]*")?)*)\s*\/?>|([^<]+)/g;
  const VOID = new Set(["meta", "link", "input", "br", "img", "hr", "source", "col", "path", "rect", "circle"]);
  let match;
  while ((match = tagPattern.exec(source))) {
    if (match[3] !== undefined) {
      const owner = stack[stack.length - 1];
      if (!HAN.test(match[3]) || !owner || owner.tag === "script" || owner.tag === "style" || !within(match.index)) continue;
      if (!owner.attrs["data-i18n"]) problems.push({ line: lineAt(match.index), kind: "text without data-i18n", text: match[3].trim().slice(0, 40), tag: owner.tag });
      continue;
    }
    const tag = match[1].toLowerCase();
    if (match[0].startsWith("</")) {
      const at = stack.map((item) => item.tag).lastIndexOf(tag);
      if (at >= 0) stack.length = at;
      continue;
    }
    const attrs = {};
    for (const attribute of match[2].matchAll(/([^\s=>]+)(?:="([^"]*)")?/g)) attrs[attribute[1]] = attribute[2] ?? "";
    const pairs = String(attrs["data-i18n-attr"] || "").split(/\s+/).filter(Boolean).map((pair) => pair.split("="));
    const named = Object.fromEntries(pairs);
    if (attrs["data-i18n"]) keys.add(attrs["data-i18n"]);
    pairs.forEach(([, key]) => keys.add(key));
    if (within(match.index)) {
      for (const [name, value] of Object.entries(attrs)) {
        if (name === "data-i18n" || name === "data-i18n-attr" || !HAN.test(value)) continue;
        if (!named[name]) problems.push({ line: lineAt(match.index), kind: `attribute ${name} not in data-i18n-attr`, text: value.slice(0, 40), tag });
      }
    }
    if (!VOID.has(tag) && !match[0].endsWith("/>")) stack.push({ tag, attrs });
  }
  // Every key used is in the English table, and the table has no key the page does not use.
  for (const key of keys) if (!Object.prototype.hasOwnProperty.call(english, key)) problems.push({ line: 0, kind: "key not in I18N_EN", text: key });
  for (const key of Object.keys(english)) if (!keys.has(key)) problems.push({ line: 0, kind: "I18N_EN key not used in index.html", text: key });
  // The English itself: no Han, no full-width punctuation.
  for (const [key, value] of Object.entries(english)) {
    if (HAN.test(value) || CJK_PUNCT_STRICT.test(value)) problems.push({ line: 0, kind: "Chinese in I18N_EN", text: `${key}: ${value}` });
  }
  return { problems, keys };
}

// The review's 13 cases: each snippet holds the stated number of missed texts.
const SELF_TEST_CASES = Object.freeze([
  ["a plain string", 'function a() { return "没有翻译"; }', 1],
  ["a template (two pieces)", "function a(n) { return `共 ${n} 行`; }", 2],
  ["a template nested in ${} of the English argument", "function a(n) { return L(`共 ${n} 行`, `${n ? `第 ${n}` : \"\"} rows`); }", 1],
  ["one branch of a conditional", 'function a(n) { return n ? L("有", "Yes") : "无"; }', 1],
  ["string concatenation", 'function a(n) { return L("甲", "A") + "乙"; }', 1],
  ["an object literal value", 'const A = Object.freeze({ x: Lp("甲", "A"), y: "乙" });', 1],
  ["an array constant", 'const A = ["甲", Lp("乙", "B")[0]];', 1],
  ["throw new Error(…)", 'function a() { throw new Error("出错了"); }', 1],
  ["Han in the English argument", 'function a() { return L("甲", "A（备注）"); }', 1],
  ["full-width punctuation in the English argument", 'function a() { return L("甲", "A：B"); }', 1],
  ["a call nested in the Chinese argument", 'function a(f) { return L(`「${f}」：${foo("内层")}`, `“${f}”: ${bar("inner")}`); }', 0],
  ["Han in a regex of a whitelisted function", 'function isYearOnlyDateValue(raw) { return /^(\\d{4})\\s*年$/.test(raw); }', 0],
  ["a second Chinese literal in a whitelisted function", 'function isYearOnlyDateValue(raw) { return /年$/.test(raw) ? "年份" : ""; }', 1]
]);
const SELF_TEST_WHITELIST = new Map([["test.js|isYearOnlyDateValue", 1]]);

module.exports = {
  lex,
  HAN,
  CJK_PUNCT_STRICT,
  APP_SCRIPT_NAMES,
  I18N_WHITELIST,
  I18N_DONE_FILES,
  isChineseLiteral,
  lexerSelfCheck,
  checkScriptTokens,
  checkIndexHtml,
  SELF_TEST_CASES,
  SELF_TEST_WHITELIST
};
