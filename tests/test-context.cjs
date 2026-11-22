const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

class FakeElement {
  constructor() {
    const classes = new Set();
    this.value = "";
    this.innerHTML = "";
    this.textContent = "";
    this.dataset = {};
    this.className = "";
    this.options = [];
    this.selectedIndex = 0;
    this.disabled = false;
    this.classList = {
      add(...names) { names.forEach((name) => classes.add(name)); },
      remove(...names) { names.forEach((name) => classes.delete(name)); },
      toggle(name, force) {
        if (force === true) classes.add(name);
        else if (force === false) classes.delete(name);
        else if (classes.has(name)) classes.delete(name);
        else classes.add(name);
        return classes.has(name);
      },
      contains(name) { return classes.has(name); }
    };
  }

  addEventListener() {}
  appendChild(child) { this.options.push(child); return child; }
  focus() {}
  removeAttribute() {}
  scrollIntoView() {}
  setAttribute() {}
  querySelectorAll() { return []; }
  getContext() {
    return { clearRect() {}, save() {}, restore() {}, fillText() {} };
  }
}

// Local app scripts in the order index.html loads them (vendor scripts excluded).
const APP_SCRIPT_FILES = [...fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8")
  .matchAll(/<script\b[^>]*\bsrc="([^"]+)"[^>]*><\/script>/g)]
  .map((match) => match[1].split("?")[0])
  .filter((src) => !src.startsWith("vendor/"));

// v3.2.0: { language: "en" } puts a browser whose first language is English (or zh-CN) in place
// before the scripts load, and { search: "?lang=en" } an address, so script.js picks the language
// the way the page does. Without them there is no navigator and no location: Chinese.
function createScriptContext({ language, search } = {}) {
  const elements = new Map();
  const document = {
    getElementById(id) {
      if (!elements.has(id)) elements.set(id, new FakeElement());
      return elements.get(id);
    },
    querySelectorAll() { return []; },
    createElement() { return new FakeElement(); },
    body: new FakeElement()
  };

  const context = vm.createContext({
    console,
    document,
    window: {
      Chart: null,
      addEventListener() {},
      innerHeight: 900,
      print() {},
      requestAnimationFrame(callback) { callback(); return 1; },
      // v3.0.0: showModule() scrolls to the top and writes the module into the address.
      scrollTo() {},
      history: { replaceState() {} },
      scrollY: 0,
      ...(search === undefined ? {} : { location: { search, hash: "", pathname: "/index.html", href: `http://127.0.0.1/index.html${search}` } })
    },
    ...(language === undefined ? {} : { navigator: { languages: [language === "en" ? "en-US" : language === "zh" ? "zh-CN" : language], language: language === "en" ? "en-US" : "zh-CN" } }),
    TextDecoder,
    Uint8Array,
    Intl,
    Blob,
    URL,
    setTimeout,
    clearTimeout
  });

  // One runInContext per file, like separate <script> tags sharing one global scope.
  const scripts = APP_SCRIPT_FILES.map((file) => ({
    file,
    text: fs.readFileSync(path.join(__dirname, "..", file), "utf8")
  }));
  scripts.forEach(({ file, text }) => vm.runInContext(text, context, { filename: file }));
  const source = scripts.map(({ text }) => text).join("\n");

  return {
    context,
    elements,
    source,
    evaluate(expression) {
      return vm.runInContext(expression, context);
    }
  };
}

module.exports = { createScriptContext, APP_SCRIPT_FILES };
