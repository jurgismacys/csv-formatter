"""
process_settlement_csv_split.py
===============================
Cleans settlement CSVs **and** automatically splits the output by currency, producing
one file per distinct value in column **M – Transaction currency**.

Now supports:
* **Multi-select** – pick one **or many** CSVs and it will process them all.
* **Choose output folder** – all formatted files are saved where you choose.
* **No overwrites** – if a filename already exists, the script appends (2), (3), …

Per-currency outputs look like:
    EUR Užsaldytos.csv
    NOK Užsaldytos.csv

Workflow applied to *each* currency-specific subset:

1. Drop rows whose **Transaction currency** is empty.
2. Keep only rows whose **Transaction type** is "Payment".
3. Convert **Settled at** (column L) using the per-row **Time zone** (column J,
   values like `UTC+02:00` / `UTC+01:00`) to UTC, then to Europe/Vilnius; this
   automatically handles summer/winter time (DST).
4. Rename column to **Settled at LT Laikas**.
5. Sort rows by that timestamp (oldest → newest).
6. Convert **Created at** (column K) in the same way; rename to **Created at LT Laikas**.
7. Retain only rows whose *Created at* month = **X** (user-chosen) **and** whose
   *Settled at* month = **Y** (user-chosen).
8. Save the cleaned subset as ``<CURRENCY> Užsaldytos.csv`` in the chosen folder.

Run the script, pick **one or more** CSVs, enter month X and month Y, choose an output folder.
A summary lists all generated files and any inputs that were skipped with errors.
"""

from __future__ import annotations

import json
import sys
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog

import pandas as pd
from zoneinfo import ZoneInfo

# ----------------------------------------------------------------------------
# Configuration ----------------------------------------------------------------
# ----------------------------------------------------------------------------
EET = ZoneInfo("Etc/GMT-2")             # legacy fallback: fixed UTC+02:00 if no Time zone column
LOCAL_TZ = ZoneInfo("Europe/Vilnius")   # applies DST automatically
DATE_FMT_OUT = "%#m/%#d/%y %H:%M" if sys.platform.startswith("win") else "%-m/%-d/%y %H:%M"

# Column names (matching sample files) -----------------------------------------
COL_CURRENCY = "Transaction currency"   # column M
COL_TYPE = "Transaction type"           # column F
COL_CREATED_IN = "Created at"           # column K
COL_SETTLED_IN = "Settled at"           # column L
COL_TIMEZONE = "Time zone"              # column J (e.g. 'UTC+02:00', 'UTC+01:00')

COL_CREATED_OUT = "Created at LT Laikas"
COL_SETTLED_OUT = "Settled at LT Laikas"

# Remembers last-used input/output folders between runs
_CONFIG_PATH = Path(__file__).resolve().parent / ".settlement_split_prefs.json"


# ----------------------------------------------------------------------------
# Persistent folder memory
# ----------------------------------------------------------------------------

def _load_prefs() -> dict:
    try:
        return json.loads(_CONFIG_PATH.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def _save_prefs(prefs: dict) -> None:
    _CONFIG_PATH.write_text(json.dumps(prefs, indent=2), encoding="utf-8")


# ----------------------------------------------------------------------------
# Utility: ensure unique filenames in the destination folder
# ----------------------------------------------------------------------------

def _unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem, suffix = path.stem, path.suffix
    i = 2
    while True:
        candidate = path.with_name(f"{stem} ({i}){suffix}")
        if not candidate.exists():
            return candidate
        i += 1


@dataclass(slots=True)
class Cleaner:
    """Applies the full cleaning workflow to a DataFrame subset."""

    df: pd.DataFrame
    month_created: int  # 1–12
    month_settled: int  # 1–12

    # ---------------------------------------------------------------------
    def process(self) -> pd.DataFrame:
        df = self.df.copy()

        # 2. Keep only Payment rows ----------------------------------------------------
        df = df[df[COL_TYPE].eq("Payment")]
        if df.empty:
            return df

        # Use the same Time zone column for both Created and Settled -------------------
        tz_series = df.get(COL_TIMEZONE)

        # 3 & 6. Parse and convert timestamps -----------------------------------------
        dt_created = self._parse_convert(df[COL_CREATED_IN], tz_series)
        dt_settled = self._parse_convert(df[COL_SETTLED_IN], tz_series)

        # 7. Filter by months ----------------------------------------------------------
        mask = dt_created.dt.month.eq(self.month_created) & dt_settled.dt.month.eq(self.month_settled)
        df = df[mask].copy()
        dt_created = dt_created[mask]
        dt_settled = dt_settled[mask]
        if df.empty:
            return df

        # 5. Sort by Settled timestamp -------------------------------------------------
        df["__dt_settled__"] = dt_settled
        df = df.sort_values("__dt_settled__")

        # 4 & 6. Finalise timestamp columns & rename -----------------------------------
        df[COL_SETTLED_IN] = dt_settled.dt.strftime(DATE_FMT_OUT)
        df[COL_CREATED_IN] = dt_created.dt.strftime(DATE_FMT_OUT)
        df = df.rename(columns={
            COL_SETTLED_IN: COL_SETTLED_OUT,
            COL_CREATED_IN: COL_CREATED_OUT,
        })
        return df.drop(columns="__dt_settled__")

    # ---------------------------------------------------------------- helper ---------
    def _parse_convert(self, series: pd.Series, tz_series: pd.Series | None) -> pd.Series:
        """
        Parse naive timestamps in `series` and convert them to Europe/Vilnius.

        If `tz_series` is provided (e.g. column "Time zone" with values like
        'UTC+02:00' / 'UTC+01:00'), we:
          - interpret the timestamp as local time in that offset,
          - convert to UTC,
          - then to Europe/Vilnius (LOCAL_TZ), which handles DST correctly.

        If `tz_series` is missing, we fall back to the old behaviour: treat
        all times as fixed UTC+02:00 (EET) and convert to LOCAL_TZ.
        """
        dt_parsed = pd.to_datetime(series, errors="coerce", utc=False)
        if dt_parsed.isna().any():
            bad = series[dt_parsed.isna()].head().tolist()
            raise ValueError(
                "Unparsable timestamps (samples): " + ", ".join(map(str, bad))
            )

        # If we don't have a time-zone column, behave as before (fixed UTC+02).
        if tz_series is None:
            return dt_parsed.dt.tz_localize(EET).dt.tz_convert(LOCAL_TZ)

        # Example values: 'UTC+02:00', 'UTC+01:00'
        tz_str = tz_series.astype(str)
        # Extract the hour part (+02, +01, etc.) as an integer.
        match = tz_str.str.extract(r"UTC([+-]\d+):")[0]

        if match.isna().any():
            bad = tz_str[match.isna()].head().tolist()
            raise ValueError(
                "Unparsable time-zone values (samples): " + ", ".join(map(str, bad))
            )

        offsets = match.astype(int)

        # Local time -> UTC: local = UTC + offset  =>  UTC = local - offset
        utc_dt = dt_parsed - pd.to_timedelta(offsets, unit="h")
        utc_dt = utc_dt.dt.tz_localize("UTC")

        # UTC -> Europe/Vilnius (handles DST, so summer vs winter is correct)
        return utc_dt.dt.tz_convert(LOCAL_TZ)


# ----------------------------------------------------------------------------
# Main application logic (per file) -------------------------------------------
# ----------------------------------------------------------------------------

def run_one(src_path: Path, month_created: int, month_settled: int, dest_dir: Path) -> list[Path]:
    """Return list of output file paths for a single input file."""
    df_all = pd.read_csv(src_path)

    # 1. Drop rows without currency ---------------------------------------------------
    df_all = df_all[
        df_all[COL_CURRENCY].notna()
        & df_all[COL_CURRENCY].astype(str).str.strip().ne("")
    ]
    if df_all.empty:
        raise ValueError("The file contains no rows with Transaction currency.")

    outputs: list[Path] = []
    for currency_code, df_curr in df_all.groupby(COL_CURRENCY, sort=False):
        cleaned = Cleaner(df_curr, month_created, month_settled).process()
        if cleaned.empty:
            # Nothing matches the month filters; skip creating an empty file.
            continue
        out_name = f"{currency_code.strip()} Užsaldytos{src_path.suffix}"
        out_path = _unique_path(dest_dir / out_name)
        cleaned.to_csv(out_path, index=False, encoding="utf-8-sig")
        outputs.append(out_path)
    return outputs


def run_batch(file_paths: list[Path], month_created: int, month_settled: int, dest_dir: Path) -> tuple[list[Path], list[str]]:
    outputs: list[Path] = []
    errors: list[str] = []
    for p in file_paths:
        try:
            outputs.extend(run_one(p, month_created, month_settled, dest_dir))
        except Exception as exc:  # noqa: BLE001
            errors.append(f"\u2716 {p.name}: {exc}")
    return outputs, errors


# ----------------------------------------------------------------------------
# Tiny Tkinter UI (multi-select + choose output folder) -----------------------
# ----------------------------------------------------------------------------

def main() -> None:
    root = tk.Tk()
    root.withdraw()  # we only need dialogs

    prefs = _load_prefs()

    # Pick one or many CSV files (start from last-used input folder)
    file_paths = filedialog.askopenfilenames(
        title="Select one or more settlement CSV files",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        initialdir=prefs.get("last_input_dir"),
    )
    if not file_paths:
        return  # user cancelled

    # Remember where the user picked files from
    prefs["last_input_dir"] = str(Path(file_paths[0]).parent)

    month_created = simpledialog.askinteger(
        "Pick month X (Created)",
        "Enter month number (1–12) to keep for *Created at*:",
        parent=root,
        minvalue=1,
        maxvalue=12,
    )
    if month_created is None:
        return

    month_settled = simpledialog.askinteger(
        "Pick month Y (Settled)",
        "Enter month number (1–12) to keep for *Settled at*:",
        parent=root,
        minvalue=1,
        maxvalue=12,
    )
    if month_settled is None:
        return

    # Choose destination folder (start from last-used output folder,
    # or fall back to where the input files live)
    fallback_dir = prefs.get("last_output_dir") or prefs["last_input_dir"]
    dest = filedialog.askdirectory(
        title="Select output folder for formatted files",
        mustexist=True,
        initialdir=fallback_dir,
    )
    if not dest:
        # Fall back to the first file's folder
        dest_dir = Path(file_paths[0]).parent
        messagebox.showinfo(
            "No folder selected",
            f"No output folder chosen. Saving next to the first input file:\n{dest_dir}",
            parent=root,
        )
    else:
        dest_dir = Path(dest)

    # Remember the output folder
    prefs["last_output_dir"] = str(dest_dir)
    _save_prefs(prefs)

    outputs, errors = run_batch([Path(p) for p in file_paths], month_created, month_settled, dest_dir)

    # Summarise results
    msg_lines = [f"Saved {len(outputs)} file(s) to:\n{dest_dir}"]
    if outputs:
        msg_lines.append("")
        msg_lines.extend(p.name for p in outputs)
    if errors:
        msg_lines.append("\nSkipped with errors:")
        msg_lines.extend(errors)

    messagebox.showinfo("Done", "\n".join(msg_lines), parent=root)


if __name__ == "__main__":
    main()