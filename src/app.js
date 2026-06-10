"use strict";

const FORMATTERS = {
  "airwallex-transactions": {
    title: "Airwallex — Transactions",
    vendor: "airwallex",
    description: "Splits settlement CSVs by currency and keeps Payment rows for the chosen month.",
    module: "airwallex_transactions",
    months: [{ key: "month", label: "Month (1–12)" }],
  },
  "airwallex-frozen": {
    title: "Airwallex — Frozen Funds (Užšaldytos)",
    vendor: "airwallex",
    description: "Keeps rows where Created month = X and Settled month = Y.",
    module: "airwallex_frozen",
    months: [
      { key: "month_created", label: "Created month (1–12)" },
      { key: "month_settled", label: "Settled month (1–12)" },
    ],
  },
  "paypal-all": {
    title: "PayPal — All Transactions",
    vendor: "paypal",
    description: "Converts CSR Date/Time to Europe/Vilnius and keeps the chosen LT month.",
    module: "paypal_all",
    months: [{ key: "month", label: "LT month (1–12)" }],
  },
  "paypal-customer": {
    title: "PayPal — Customer Payments Only",
    vendor: "paypal",
    description: "Same as All Transactions, but only Express Checkout Payment rows.",
    module: "paypal_customer",
    months: [{ key: "month", label: "LT month (1–12)" }],
  },
};

const MODULE_FILES = [
  "airwallex_transactions",
  "airwallex_frozen",
  "paypal_all",
  "paypal_customer",
];

let pyodide = null;
let pyodideReady = null;
const statusEl = document.getElementById("status");
const view = document.getElementById("view");

function setStatus(text, cls = "") {
  statusEl.textContent = text;
  statusEl.className = "status " + cls;
}

async function initPyodide() {
  setStatus("Loading Python runtime…");
  pyodide = await loadPyodide({
    indexURL: "https://cdn.jsdelivr.net/pyodide/v0.26.4/full/",
  });

  setStatus("Loading pandas…");
  await pyodide.loadPackage(["pandas", "micropip"]);

  setStatus("Loading timezone data…");
  await pyodide.runPythonAsync(`
import micropip
await micropip.install('tzdata')
`);

  setStatus("Loading formatter scripts…");
  // Stub tkinter so the scripts' module-level imports succeed in the browser.
  // The web UI doesn't call main()/Tk; only run_batch/run_one/run_split.
  await pyodide.runPython(`
import sys, types
for name in ['tkinter', 'tkinter.filedialog', 'tkinter.messagebox', 'tkinter.simpledialog']:
    if name not in sys.modules:
        sys.modules[name] = types.ModuleType(name)
sys.modules['tkinter'].filedialog = sys.modules['tkinter.filedialog']
sys.modules['tkinter'].messagebox = sys.modules['tkinter.messagebox']
sys.modules['tkinter'].simpledialog = sys.modules['tkinter.simpledialog']
sys.modules['tkinter'].Tk = lambda: None

# Emscripten libc doesn't support GNU %-m / %-d strftime extensions.
# Substitute them per-element before delegating to libc, so the scripts'
# original "6/1/25 17:08" output format is preserved verbatim.
import pandas as pd
from pandas.core.indexes.accessors import DatetimeProperties
_orig_strftime = DatetimeProperties.strftime
def _strftime_gnu(self, date_format):
    if '%-m' in date_format or '%-d' in date_format:
        def _fmt(ts):
            if pd.isna(ts):
                return None
            f = date_format.replace('%-m', str(ts.month)).replace('%-d', str(ts.day))
            return ts.strftime(f)
        return pd.Series(self._parent).map(_fmt)
    return _orig_strftime(self, date_format)
DatetimeProperties.strftime = _strftime_gnu
`);

  pyodide.FS.mkdirTree("/scripts");
  for (const name of MODULE_FILES) {
    const res = await fetch(`scripts/${name}.py?v=1`);
    if (!res.ok) throw new Error(`Failed to fetch scripts/${name}.py`);
    const text = await res.text();
    pyodide.FS.writeFile(`/scripts/${name}.py`, text);
  }

  await pyodide.runPython(`
import sys
if '/scripts' not in sys.path:
    sys.path.insert(0, '/scripts')
import airwallex_transactions, airwallex_frozen, paypal_all, paypal_customer
`);

  setStatus("Ready", "ready");
}

function startPyodide() {
  if (!pyodideReady) {
    pyodideReady = initPyodide().catch((err) => {
      console.error(err);
      setStatus("Python failed to load", "error");
      throw err;
    });
  }
  return pyodideReady;
}

// ---------------------------------------------------------------------------
// Views
// ---------------------------------------------------------------------------

function el(tag, attrs = {}, ...children) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === "class") e.className = v;
    else if (k === "html") e.innerHTML = v;
    else if (k.startsWith("on") && typeof v === "function") {
      e.addEventListener(k.slice(2).toLowerCase(), v);
    } else if (v !== false && v != null) {
      e.setAttribute(k, v);
    }
  }
  for (const c of children) {
    if (c == null) continue;
    e.append(c.nodeType ? c : document.createTextNode(c));
  }
  return e;
}

function render() {
  const hash = location.hash || "#/";
  view.innerHTML = "";

  if (hash === "#/" || hash === "") {
    view.append(renderHome());
  } else if (hash === "#/paypal") {
    view.append(renderVendor("PayPal", "paypal"));
  } else if (hash === "#/airwallex") {
    view.append(renderVendor("Airwallex", "airwallex"));
  } else if (hash.startsWith("#/run/")) {
    const key = hash.slice("#/run/".length);
    if (FORMATTERS[key]) {
      view.append(renderRun(key));
    } else {
      view.append(el("p", {}, "Unknown formatter."));
    }
  } else {
    view.append(el("p", {}, "Not found."));
  }
}

function renderHome() {
  const root = el("section", { class: "hero" });
  root.append(
    el("h1", {}, "CSV Formatter"),
    el("p", { class: "lede" }, "Pick a source to format."),
  );
  const grid = el("div", { class: "card-grid" });
  grid.append(
    el("a", { class: "card", href: "#/paypal" },
      el("h2", {}, "Format PayPal docs"),
      el("p", {}, "All transactions · Customer payments only"),
    ),
    el("a", { class: "card", href: "#/airwallex" },
      el("h2", {}, "Format Airwallex docs"),
      el("p", {}, "Transactions · Frozen funds (Užšaldytos)"),
    ),
  );
  root.append(grid);
  return root;
}

function renderVendor(label, vendor) {
  const root = el("section");
  root.append(el("h1", {}, label), el("p", { class: "lede" }, "Choose a formatter."));
  const grid = el("div", { class: "card-grid" });
  for (const [key, cfg] of Object.entries(FORMATTERS)) {
    if (cfg.vendor !== vendor) continue;
    const shortTitle = cfg.title.split("—")[1].trim();
    grid.append(
      el("a", { class: "card", href: `#/run/${key}` },
        el("h2", {}, shortTitle),
        el("p", {}, cfg.description),
      ),
    );
  }
  root.append(grid);
  root.append(el("p", {}, el("a", { class: "back", href: "#/" }, "← Back")));
  return root;
}

function renderRun(key) {
  const cfg = FORMATTERS[key];
  const root = el("section");
  root.append(el("h1", {}, cfg.title), el("p", { class: "lede" }, cfg.description));

  const form = el("form", { class: "run-form" });

  const fileLabel = el("label", { class: "field" });
  fileLabel.append(
    el("span", {}, "CSV file(s)"),
    el("input", { type: "file", name: "files", accept: ".csv", multiple: "" }),
    el("small", {}, "You can select multiple files."),
  );
  form.append(fileLabel);

  for (const m of cfg.months) {
    const lbl = el("label", { class: "field" });
    lbl.append(
      el("span", {}, m.label),
      el("input", { type: "number", name: m.key, min: 1, max: 12, required: "" }),
    );
    form.append(lbl);
  }

  const submit = el("button", { type: "submit" }, "Format & download");
  const back = el("a", { class: "back", href: `#/${cfg.vendor}` }, "← Back");
  const actions = el("div", { class: "actions" }, submit, back);
  form.append(actions);

  const message = el("div");
  form.append(message);

  form.addEventListener("submit", async (ev) => {
    ev.preventDefault();
    message.className = "";
    message.textContent = "";

    const fileInput = form.querySelector('input[type="file"]');
    const files = fileInput.files;
    if (!files || files.length === 0) {
      message.className = "message error";
      message.textContent = "Please select at least one CSV.";
      return;
    }

    const months = {};
    for (const m of cfg.months) {
      const v = parseInt(form.querySelector(`input[name="${m.key}"]`).value, 10);
      if (!(v >= 1 && v <= 12)) {
        message.className = "message error";
        message.textContent = `${m.label} must be 1–12.`;
        return;
      }
      months[m.key] = v;
    }

    submit.disabled = true;
    message.className = "message info";
    message.textContent = "Loading Python (first run takes ~15 seconds)…";

    try {
      await startPyodide();
      message.textContent = "Processing…";
      const result = await runFormatter(key, Array.from(files), months);
      if (result.errors.length > 0 && result.outputs.length === 0) {
        message.className = "message error";
        message.textContent = "No output rows matched.\n" + result.errors.join("\n");
      } else {
        triggerDownload(result, key);
        message.className = "message ok";
        const lines = [`Saved ${result.outputs.length} file(s).`];
        for (const o of result.outputs) lines.push("• " + o.name);
        if (result.errors.length) lines.push("\nWarnings:\n" + result.errors.join("\n"));
        message.textContent = lines.join("\n");
      }
    } catch (err) {
      console.error(err);
      message.className = "message error";
      message.textContent = "Error: " + (err.message || String(err));
    } finally {
      submit.disabled = false;
    }
  });

  root.append(form);
  return root;
}

// ---------------------------------------------------------------------------
// Pyodide invocation
// ---------------------------------------------------------------------------

async function runFormatter(key, files, months) {
  const cfg = FORMATTERS[key];

  // Reset working dirs
  pyodide.runPython(`
import shutil, os
from pathlib import Path
for d in ['/work/in', '/work/out']:
    p = Path(d)
    if p.exists():
        shutil.rmtree(p)
    p.mkdir(parents=True)
`);

  // Write uploaded files
  for (const f of files) {
    const buf = new Uint8Array(await f.arrayBuffer());
    pyodide.FS.writeFile(`/work/in/${f.name}`, buf);
  }

  // Build call
  let pyCall;
  if (cfg.months.length === 1) {
    pyCall = `outputs, errors = ${cfg.module}.run_batch(file_paths, ${months[cfg.months[0].key]}, out_dir)`;
  } else {
    const a = months[cfg.months[0].key];
    const b = months[cfg.months[1].key];
    pyCall = `outputs, errors = ${cfg.module}.run_batch(file_paths, ${a}, ${b}, out_dir)`;
  }

  const code = `
from pathlib import Path
import ${cfg.module}
in_dir = Path('/work/in')
out_dir = Path('/work/out')
file_paths = sorted(in_dir.iterdir())
${pyCall}
([str(p) for p in outputs], list(errors))
`;

  const result = await pyodide.runPythonAsync(code);
  const [outPaths, errors] = result.toJs({ create_proxies: false });
  result.destroy();

  const outputs = outPaths.map((p) => {
    const bytes = pyodide.FS.readFile(p);
    const name = p.split("/").pop();
    return { name, bytes };
  });

  return { outputs, errors };
}

function triggerDownload({ outputs, errors }, key) {
  if (outputs.length === 1 && errors.length === 0) {
    const o = outputs[0];
    downloadBlob(new Blob([o.bytes], { type: "text/csv" }), o.name);
    return;
  }
  // Zip via JSZip-free hand-rolled? Easier: just download each. But user asked for clean UX.
  // Use a simple bundle: trigger each download sequentially.
  outputs.forEach((o, i) => {
    setTimeout(() => downloadBlob(new Blob([o.bytes], { type: "text/csv" }), o.name), i * 200);
  });
}

function downloadBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

// ---------------------------------------------------------------------------
// Init
// ---------------------------------------------------------------------------

window.addEventListener("hashchange", render);
render();
// Eagerly start loading Pyodide in the background so it's ready when needed.
startPyodide();
