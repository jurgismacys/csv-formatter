# CSV Formatter

A small in-browser tool that runs four CSV-formatting scripts (PayPal + Airwallex) entirely client-side via Pyodide. No backend.

**Live (password-protected):** https://jurgismacys.github.io/csv-formatter/

The published site is a single AES-256-encrypted page (via [StatiCrypt](https://github.com/robinmoisson/staticrypt)). Visitors must enter the access password to decrypt and use it in the browser. The password is **not** stored in this repo.

## What it does

- **PayPal — All Transactions**: converts CSR Date/Time to Europe/Vilnius, keeps the chosen LT month.
- **PayPal — Customer Payments Only**: same as above but only Express Checkout Payment rows.
- **Airwallex — Transactions**: splits settlement CSVs by currency and keeps Payment rows for the chosen month.
- **Airwallex — Frozen Funds (Užšaldytos)**: keeps rows where Created month = X and Settled month = Y.
- **Gisko — Pardavimų ataskaita**: takes **two** Shopify exports (Orders + Transaction histories) and outputs a formatted `.xlsx` with per-payment-method amount splits (mixed payments highlighted yellow, refunds/cancellations red, totals row). Two upload fields → one Excel file. Uses `openpyxl` (loaded in-browser via micropip).

## Structure

- `src/` — the source of truth: `index.html`, `style.css`, `app.js`, and `scripts/*.py` (the four formatters). Edit here.
- `build.py` — bundles `src/` into a single self-contained `build/index.html` (inlines CSS/JS, embeds the Python scripts — no runtime `fetch`).
- `docs/index.html` — the **encrypted** bundle. This is the only file GitHub Pages serves.
- `app.py`, `templates/`, `static/` — legacy Flask version used during early local development.

## Rebuild & re-encrypt (after editing `src/`)

```bash
# 1. bundle src/ -> build/index.html
python3 build.py

# 2. encrypt the bundle with the access password
npx staticrypt build/index.html --password "YOUR_PASSWORD" \
  -d build/enc --remember 30 --template-title "CSV Formatter — Locked"

# 3. promote the encrypted file to the served dir and push
cp build/enc/index.html docs/index.html
git add docs/index.html && git commit -m "Rebuild encrypted site" && git push
```

GitHub Pages (Settings → Pages → `main` / `/docs`) rebuilds automatically on push.

### Changing the password

Re-run the rebuild steps above with the new `--password`. The old encrypted file is fully replaced; previously-shared links keep working but require the new password.

## Run the *unencrypted* version locally (dev)

```bash
python3 build.py
python3 -m http.server 8000 --directory build
# open http://127.0.0.1:8000  (no password prompt — for testing only)
```
