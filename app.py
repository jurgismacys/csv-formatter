"""
Flask wrapper around the four existing formatter scripts.

The scripts themselves are imported as-is via importlib (their file names
contain spaces, so a normal `import` doesn't work). We re-use their
`run_batch` functions verbatim — no edits to formatting logic.

Routes
------
GET  /                                 landing page
GET  /paypal                           pick PayPal formatter
GET  /airwallex                        pick Airwallex formatter
GET  /run/<key>                        upload form for a specific formatter
POST /run/<key>                        process upload, return ZIP or single CSV
"""

from __future__ import annotations

import importlib.util
import io
import shutil
import sys
import tempfile
import uuid
import zipfile
from pathlib import Path

from flask import (
    Flask,
    abort,
    render_template,
    request,
    send_file,
)

# ---------------------------------------------------------------------------
# Load the four formatter scripts as modules (without editing them)
# ---------------------------------------------------------------------------

SCRIPTS_ROOT = Path(__file__).resolve().parent.parent


def _load(path: Path, alias: str):
    spec = importlib.util.spec_from_file_location(alias, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load module at {path}")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[alias] = mod
    spec.loader.exec_module(mod)
    return mod


airwallex_tx = _load(
    SCRIPTS_ROOT / "Airwallex formatter" / "transaction formatter.py",
    "airwallex_tx",
)
airwallex_uzs = _load(
    SCRIPTS_ROOT / "Airwallex formatter" / "uzsaldytos.py",
    "airwallex_uzs",
)
paypal_all = _load(
    SCRIPTS_ROOT / "PayPal formatter" / "All transactions.py",
    "paypal_all",
)
paypal_cust = _load(
    SCRIPTS_ROOT / "PayPal formatter" / "Customer payments only.py",
    "paypal_cust",
)


# ---------------------------------------------------------------------------
# Formatter registry
# ---------------------------------------------------------------------------

FORMATTERS = {
    "airwallex-transactions": {
        "title": "Airwallex — Transactions",
        "vendor": "airwallex",
        "description": "Splits settlement CSVs by currency and keeps Payment rows for the chosen month.",
        "module": airwallex_tx,
        "months": ["month"],
        "month_labels": {"month": "Month (1–12)"},
    },
    "airwallex-frozen": {
        "title": "Airwallex — Frozen Funds (Užšaldytos)",
        "vendor": "airwallex",
        "description": "Keeps rows where Created month = X and Settled month = Y.",
        "module": airwallex_uzs,
        "months": ["month_created", "month_settled"],
        "month_labels": {
            "month_created": "Created month (1–12)",
            "month_settled": "Settled month (1–12)",
        },
    },
    "paypal-all": {
        "title": "PayPal — All Transactions",
        "vendor": "paypal",
        "description": "Converts CSR Date/Time to Europe/Vilnius and keeps the chosen LT month.",
        "module": paypal_all,
        "months": ["month"],
        "month_labels": {"month": "LT month (1–12)"},
    },
    "paypal-customer": {
        "title": "PayPal — Customer Payments Only",
        "vendor": "paypal",
        "description": "Same as All Transactions, but only Express Checkout Payment rows.",
        "module": paypal_cust,
        "months": ["month"],
        "month_labels": {"month": "LT month (1–12)"},
    },
}


# ---------------------------------------------------------------------------
# Flask app
# ---------------------------------------------------------------------------

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024  # 64 MB upload cap


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/paypal")
def paypal():
    items = {k: v for k, v in FORMATTERS.items() if v["vendor"] == "paypal"}
    return render_template("vendor.html", vendor="PayPal", items=items)


@app.route("/airwallex")
def airwallex():
    items = {k: v for k, v in FORMATTERS.items() if v["vendor"] == "airwallex"}
    return render_template("vendor.html", vendor="Airwallex", items=items)


@app.route("/run/<key>", methods=["GET", "POST"])
def run(key: str):
    cfg = FORMATTERS.get(key)
    if cfg is None:
        abort(404)

    if request.method == "GET":
        return render_template("run.html", key=key, cfg=cfg)

    files = request.files.getlist("files")
    files = [f for f in files if f and f.filename]
    if not files:
        return render_template("run.html", key=key, cfg=cfg, error="No files uploaded."), 400

    months: dict[str, int] = {}
    for name in cfg["months"]:
        raw = request.form.get(name, "").strip()
        try:
            val = int(raw)
        except ValueError:
            return (
                render_template("run.html", key=key, cfg=cfg, error=f"{cfg['month_labels'][name]} must be 1–12."),
                400,
            )
        if not (1 <= val <= 12):
            return (
                render_template("run.html", key=key, cfg=cfg, error=f"{cfg['month_labels'][name]} must be 1–12."),
                400,
            )
        months[name] = val

    work_root = Path(tempfile.mkdtemp(prefix=f"fmt_{key}_"))
    in_dir = work_root / "in"
    out_dir = work_root / "out"
    in_dir.mkdir()
    out_dir.mkdir()

    try:
        input_paths: list[Path] = []
        for f in files:
            safe_name = Path(f.filename).name
            dest = in_dir / safe_name
            f.save(dest)
            input_paths.append(dest)

        module = cfg["module"]
        if cfg["months"] == ["month"]:
            outputs, errors = module.run_batch(input_paths, months["month"], out_dir)
        elif cfg["months"] == ["month_created", "month_settled"]:
            outputs, errors = module.run_batch(
                input_paths, months["month_created"], months["month_settled"], out_dir
            )
        else:
            raise RuntimeError(f"Unsupported months config: {cfg['months']}")

        if not outputs:
            msg = "No output rows matched the filters."
            if errors:
                msg += " Errors: " + " | ".join(errors)
            return render_template("run.html", key=key, cfg=cfg, error=msg), 400

        if len(outputs) == 1 and not errors:
            data = outputs[0].read_bytes()
            return send_file(
                io.BytesIO(data),
                as_attachment=True,
                download_name=outputs[0].name,
                mimetype="text/csv",
            )

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for p in outputs:
                zf.write(p, arcname=p.name)
            if errors:
                zf.writestr("_errors.txt", "\n".join(errors))
        buf.seek(0)
        zip_name = f"{key}-{uuid.uuid4().hex[:8]}.zip"
        return send_file(buf, as_attachment=True, download_name=zip_name, mimetype="application/zip")

    finally:
        shutil.rmtree(work_root, ignore_errors=True)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5050, debug=True)
