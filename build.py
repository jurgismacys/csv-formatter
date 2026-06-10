#!/usr/bin/env python3
"""Bundle the CSV Formatter into a single self-contained index.html.

Inlines style.css and app.js, and embeds the four Python formatter scripts as
JS string constants (so the app no longer fetches scripts/*.py at runtime).
The bundled file in build/index.html is then encrypted with StatiCrypt and the
result becomes docs/index.html — the only file served by GitHub Pages.

Source of truth lives in src/. Run: python3 build.py
"""
from __future__ import annotations
import json
from pathlib import Path

ROOT = Path(__file__).parent
SRC = ROOT / "src"
BUILD = ROOT / "build"

MODULE_FILES = [
    "airwallex_transactions",
    "airwallex_frozen",
    "paypal_all",
    "paypal_customer",
    "gisko_sales",
]

# --- read source assets -----------------------------------------------------
index_html = (SRC / "index.html").read_text()
style_css = (SRC / "style.css").read_text()
app_js = (SRC / "app.js").read_text()
scripts = {name: (SRC / "scripts" / f"{name}.py").read_text() for name in MODULE_FILES}

# --- embed python scripts into the JS, replacing the runtime fetch loop -----
FETCH_BLOCK = '''  pyodide.FS.mkdirTree("/scripts");
  for (const name of MODULE_FILES) {
    const res = await fetch(`scripts/${name}.py?v=1`);
    if (!res.ok) throw new Error(`Failed to fetch scripts/${name}.py`);
    const text = await res.text();
    pyodide.FS.writeFile(`/scripts/${name}.py`, text);
  }'''

EMBED_BLOCK = '''  pyodide.FS.mkdirTree("/scripts");
  for (const name of MODULE_FILES) {
    pyodide.FS.writeFile(`/scripts/${name}.py`, EMBEDDED_SCRIPTS[name]);
  }'''

if FETCH_BLOCK not in app_js:
    raise SystemExit("Could not find the fetch block in app.js — did it change?")
app_js = app_js.replace(FETCH_BLOCK, EMBED_BLOCK)

embedded = "const EMBEDDED_SCRIPTS = " + json.dumps(scripts) + ";\n\n"
app_js = embedded + app_js

# --- inline css + js into the html ------------------------------------------
html = index_html.replace(
    '  <link rel="stylesheet" href="style.css">',
    "  <style>\n" + style_css + "\n  </style>",
)
html = html.replace(
    '  <script src="app.js"></script>',
    "  <script>\n" + app_js + "\n  </script>",
)

BUILD.mkdir(exist_ok=True)
(BUILD / "index.html").write_text(html)
print(f"Wrote {BUILD / 'index.html'} ({len(html):,} bytes)")
