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

function createScriptContext() {
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
      scrollY: 0
    },
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
