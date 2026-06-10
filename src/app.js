"use strict";

// Each formatter declares its inputs generically:
//   fileFields: one or more upload fields (named, single or multiple)
//   params:     scalar inputs (number or text)
//   output:     "csv" | "xlsx"
//   pyInvoke:   builds the Python that sets `outputs` (list of paths) + `errors`
const FORMATTERS = {
  "airwallex-transactions": {
    title: "Airwallex — Transactions",
    vendor: "airwallex",
    description: "Splits settlement CSVs by currency and keeps Payment rows for the chosen month.",
    module: "airwallex_transactions",
    output: "csv",
    fileFields: [
      { key: "files", label: "CSV file(s)", accept: ".csv", multiple: true, hint: "You can select multiple files." },
    ],
    params: [{ key: "month", label: "Month (1–12)", type: "number", min: 1, max: 12, required: true }],
    pyInvoke: (f, p) => `
file_paths = sorted(Path(x) for x in ${f.files})
outputs, errors = airwallex_transactions.run_batch(file_paths, ${p.month}, out_dir)
`,
  },
  "airwallex-frozen": {
    title: "Airwallex — Frozen Funds (Užšaldytos)",
    vendor: "airwallex",
    description: "Keeps rows where Created month = X and Settled month = Y.",
    module: "airwallex_frozen",
    output: "csv",
    fileFields: [
      { key: "files", label: "CSV file(s)", accept: ".csv", multiple: true, hint: "You can select multiple files." },
    ],
    params: [
      { key: "month_created", label: "Created month (1–12)", type: "number", min: 1, max: 12, required: true },
      { key: "month_settled", label: "Settled month (1–12)", type: "number", min: 1, max: 12, required: true },
    ],
    pyInvoke: (f, p) => `
file_paths = sorted(Path(x) for x in ${f.files})
outputs, errors = airwallex_frozen.run_batch(file_paths, ${p.month_created}, ${p.month_settled}, out_dir)
`,
  },
  "paypal-all": {
    title: "PayPal — All Transactions",
    vendor: "paypal",
    description: "Converts CSR Date/Time to Europe/Vilnius and keeps the chosen LT month.",
    module: "paypal_all",
    output: "csv",
    fileFields: [
      { key: "files", label: "CSV file(s)", accept: ".csv", multiple: true, hint: "You can select multiple files." },
    ],
    params: [{ key: "month", label: "LT month (1–12)", type: "number", min: 1, max: 12, required: true }],
    pyInvoke: (f, p) => `
file_paths = sorted(Path(x) for x in ${f.files})
outputs, errors = paypal_all.run_batch(file_paths, ${p.month}, out_dir)
`,
  },
  "paypal-customer": {
    title: "PayPal — Customer Payments Only",
    vendor: "paypal",
    description: "Same as All Transactions, but only Express Checkout Payment rows.",
    module: "paypal_customer",
    output: "csv",
    fileFields: [
      { key: "files", label: "CSV file(s)", accept: ".csv", multiple: true, hint: "You can select multiple files." },
    ],
    params: [{ key: "month", label: "LT month (1–12)", type: "number", min: 1, max: 12, required: true }],
    pyInvoke: (f, p) => `
file_paths = sorted(Path(x) for x in ${f.files})
outputs, errors = paypal_customer.run_batch(file_paths, ${p.month}, out_dir)
`,
  },
  "gisko-sales": {
    title: "Gisko — Pardavimų ataskaita",
    vendor: "gisko",
    backHref: "#/",
    description:
      "Sujungia Shopify Orders + Transaction histories eksportus į formatuotą Excel su mokėjimo būdų išskaidymu (mišrūs apmokėjimai pažymimi geltonai, grąžinimai raudonai).",
    module: "gisko_sales",
    output: "xlsx",
    fileFields: [
      {
        key: "orders",
        label: "Shopify Orders eksportas (CSV)",
        accept: ".csv",
        multiple: false,
        hint: "Admin → Orders → Export → „Export orders“",
      },
      {
        key: "transactions",
        label: "Shopify Transaction histories eksportas (CSV)",
        accept: ".csv",
        multiple: false,
        hint: "Admin → Orders → Export → „Export transaction histories“",
      },
    ],
    params: [
      {
        key: "period",
        label: "Laikotarpis (nebūtina, pvz. 2026-05)",
        type: "text",
        required: false,
        placeholder: "2026-05",
      },
    ],
    pyInvoke: (f, p) => `
period = ${p.period}
name = "Gisko_pardavimai_" + (period if period else "report") + ".xlsx"
out_path = str(out_dir / name)
gisko_sales.run(${f.orders}[0], ${f.transactions}[0], out_path, period)
outputs = [out_path]
errors = []
`,
  },
};

const MODULE_FILES = [
  "airwallex_transactions",
  "airwallex_frozen",
  "paypal_all",
  "paypal_customer",
  "gisko_sales",
];

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

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

  setStatus("Loading timezone + Excel support…");
  await pyodide.runPythonAsync(`
import micropip
await micropip.install(['tzdata', 'openpyxl'])
`);

  setStatus("Loading formatter scripts…");
  // Stub tkinter so the scripts' module-level imports succeed in the browser.
  // The web UI doesn't call main()/Tk; only run_batch/run_one/run_split/run.
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
import airwallex_transactions, airwallex_frozen, paypal_all, paypal_customer, gisko_sales
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
    el("a", { class: "card", href: "#/run/gisko-sales" },
      el("h2", {}, "Gisko sales report"),
      el("p", {}, "Shopify Orders + Transactions → formatted Excel"),
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

  for (const ff of cfg.fileFields) {
    const fileLabel = el("label", { class: "field" });
    fileLabel.append(
      el("span", {}, ff.label),
      el("input", {
        type: "file",
        name: ff.key,
        accept: ff.accept || ".csv",
        multiple: ff.multiple ? "" : false,
      }),
      ff.hint ? el("small", {}, ff.hint) : null,
    );
    form.append(fileLabel);
  }

  for (const pr of cfg.params) {
    const lbl = el("label", { class: "field" });
    const input =
      pr.type === "number"
        ? el("input", {
            type: "number",
            name: pr.key,
            min: pr.min,
            max: pr.max,
            required: pr.required ? "" : false,
          })
        : el("input", {
            type: "text",
            name: pr.key,
            placeholder: pr.placeholder || "",
            required: pr.required ? "" : false,
          });
    lbl.append(el("span", {}, pr.label), input);
    form.append(lbl);
  }

  const submit = el("button", { type: "submit" }, "Format & download");
  const back = el("a", { class: "back", href: cfg.backHref || `#/${cfg.vendor}` }, "← Back");
  const actions = el("div", { class: "actions" }, submit, back);
  form.append(actions);

  const message = el("div");
  form.append(message);

  form.addEventListener("submit", async (ev) => {
    ev.preventDefault();
    message.className = "";
    message.textContent = "";

    const fail = (msg) => {
      message.className = "message error";
      message.textContent = msg;
    };

    // Collect files per field
    const fieldFiles = {};
    for (const ff of cfg.fileFields) {
      const input = form.querySelector(`input[type="file"][name="${ff.key}"]`);
      const files = Array.from(input.files || []);
      if (files.length === 0) {
        fail(`Please select a file: ${ff.label}`);
        return;
      }
      fieldFiles[ff.key] = files;
    }

    // Collect + validate params
    const params = {};
    for (const pr of cfg.params) {
      const input = form.querySelector(`[name="${pr.key}"]`);
      if (pr.type === "number") {
        const v = parseInt(input.value, 10);
        if (pr.min != null && pr.max != null && !(v >= pr.min && v <= pr.max)) {
          fail(`${pr.label} must be ${pr.min}–${pr.max}.`);
          return;
        }
        params[pr.key] = v;
      } else {
        const v = (input.value || "").trim();
        if (pr.required && !v) {
          fail(`${pr.label} is required.`);
          return;
        }
        params[pr.key] = v;
      }
    }

    submit.disabled = true;
    message.className = "message info";
    message.textContent = "Loading Python (first run takes ~15 seconds)…";

    try {
      await startPyodide();
      message.textContent = "Processing…";
      const result = await runFormatter(key, fieldFiles, params);
      if (result.errors.length > 0 && result.outputs.length === 0) {
        message.className = "message error";
        message.textContent = "No output produced.\n" + result.errors.join("\n");
      } else {
        triggerDownload(result);
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

async function runFormatter(key, fieldFiles, params) {
  const cfg = FORMATTERS[key];

  // Reset working dirs
  pyodide.runPython(`
import shutil
from pathlib import Path
for d in ['/work/in', '/work/out']:
    p = Path(d)
    if p.exists():
        shutil.rmtree(p)
    p.mkdir(parents=True)
`);

  // Write uploaded files, grouped per field to avoid name collisions
  const fsPaths = {};
  for (const ff of cfg.fileFields) {
    fsPaths[ff.key] = [];
    pyodide.FS.mkdirTree(`/work/in/${ff.key}`);
    for (const f of fieldFiles[ff.key]) {
      const buf = new Uint8Array(await f.arrayBuffer());
      const p = `/work/in/${ff.key}/${f.name}`;
      pyodide.FS.writeFile(p, buf);
      fsPaths[ff.key].push(p);
    }
  }

  // Render python literals: file paths as JSON arrays (valid Python list of str),
  // number params bare, text params JSON-quoted.
  const f = {};
  for (const k in fsPaths) f[k] = JSON.stringify(fsPaths[k]);
  const p = {};
  for (const pr of cfg.params) {
    p[pr.key] = pr.type === "number" ? params[pr.key] : JSON.stringify(params[pr.key] ?? "");
  }

  const body = cfg.pyInvoke(f, p);
  const code = `
from pathlib import Path
import ${cfg.module}
out_dir = Path('/work/out')
${body}
([str(x) for x in outputs], list(errors))
`;

  const result = await pyodide.runPythonAsync(code);
  const [outPaths, errors] = result.toJs({ create_proxies: false });
  result.destroy();

  const mime = cfg.output === "xlsx" ? XLSX_MIME : "text/csv";
  const outputs = outPaths.map((pp) => {
    const bytes = pyodide.FS.readFile(pp);
    const name = pp.split("/").pop();
    return { name, bytes, type: mime };
  });

  return { outputs, errors };
}

function triggerDownload({ outputs, errors }) {
  if (outputs.length === 1 && errors.length === 0) {
    const o = outputs[0];
    downloadBlob(new Blob([o.bytes], { type: o.type }), o.name);
    return;
  }
  outputs.forEach((o, i) => {
    setTimeout(() => downloadBlob(new Blob([o.bytes], { type: o.type }), o.name), i * 200);
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
