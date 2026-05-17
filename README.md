# CSV Formatter

A small in-browser tool that runs four CSV-formatting scripts (PayPal + Airwallex) entirely client-side via Pyodide. No backend.

## What it does

- **PayPal — All Transactions**: converts CSR Date/Time to Europe/Vilnius, keeps the chosen LT month.
- **PayPal — Customer Payments Only**: same as above but only Express Checkout Payment rows.
- **Airwallex — Transactions**: splits settlement CSVs by currency and keeps Payment rows for the chosen month.
- **Airwallex — Frozen Funds (Užšaldytos)**: keeps rows where Created month = X and Settled month = Y.

## Run locally

```bash
python3 -m http.server 8000 --directory docs
# then open http://127.0.0.1:8000
```

There's also a Flask version (`app.py`) that uses the same scripts server-side — see `app.py`. To run:

```bash
pip3 install flask pandas
python3 app.py
```

## Structure

- `docs/` — the static GitHub Pages site (Pyodide-powered).
- `docs/scripts/` — the four formatter scripts, unmodified copies of the originals.
- `app.py`, `templates/`, `static/` — the Flask version used during local development.
