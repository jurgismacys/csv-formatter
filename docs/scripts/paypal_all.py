"""
process_csr_lt_month.py
=======================
Takes PayPal CSR CSV files where Date/Time are in the CSR time zone
(e.g. America/Los_Angeles), converts them to Europe/Vilnius time and
keeps only the rows that fall into a user-chosen month (based on LT time).

Features:
- Multi-select: pick one or many CSR CSV files.
- Choose output folder.
- No overwrites: if a filename already exists, appends (2), (3), …

For each input file, outputs ONE CSV named:
    Visos transakcijos pagal LT laiko zona.csv
(or "Visos transakcijos pagal LT laiko zona (2).csv", etc. if needed)

Workflow per file:
1. Combine "Date" + "Time" and parse as datetimes.
2. Interpret those as local times in the file's "Time Zone" (e.g. America/Los_Angeles),
   using Python's zoneinfo (handles DST, including summer/winter changes).
3. Convert to Europe/Vilnius.
4. Filter rows where LT month == user-chosen month (1–12).
5. Sort rows chronologically by LT time.
6. Add a column "DateTime LT Laikas" formatted as 6/1/25 17:08.
7. Save the filtered file in the chosen folder.
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
# Configuration
# ----------------------------------------------------------------------------

DATE_COL = "Date"
TIME_COL = "Time"
TZ_COL = "Time Zone"

LT_TZ = ZoneInfo("Europe/Vilnius")  # target timezone with DST rules

# CSV-friendly datetime format, same style as other scripts
DATE_FMT_OUT = "%#m/%#d/%y %H:%M" if sys.platform.startswith("win") else "%-m/%-d/%y %H:%M"

# Remembers last-used input/output folders between runs
_CONFIG_PATH = Path(__file__).resolve().parent / ".all_transactions_prefs.json"


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
    """Return a path that does not yet exist, appending (2), (3), ... if needed."""
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
    """Convert CSR times to LT time, filter by LT month, and sort chronologically."""

    df: pd.DataFrame
    month: int  # 1–12

    # ---------------------------------------------------------------------
    def process(self) -> pd.DataFrame:
        df = self.df.copy()

        # Parse and convert to LT time ------------------------------------
        dt_lt = self._parse_and_convert_to_lt(df)

        # Filter by LT month ----------------------------------------------
        mask = dt_lt.dt.month.eq(self.month)
        df = df[mask].copy()
        dt_lt = dt_lt[mask]
        if df.empty:
            return df

        # Sort chronologically by LT time ---------------------------------
        df["__dt_lt__"] = dt_lt
        df = df.sort_values("__dt_lt__")

        # Expose a formatted LT datetime column ---------------------------
        df["DateTime LT Laikas"] = df["__dt_lt__"].dt.strftime(DATE_FMT_OUT)

        return df.drop(columns="__dt_lt__")

    # ---------------------------------------------------------------- helper
    def _parse_and_convert_to_lt(self, df: pd.DataFrame) -> pd.Series:
        """
        Combine Date + Time, interpret them in the CSR time zone from TZ_COL,
        and convert to Europe/Vilnius.

        Uses Python's zoneinfo on each timestamp, which avoids the pandas
        tz_localize() ambiguity error at DST transitions.
        """
        # Combine Date + Time as strings
        combined = (
            df[DATE_COL].astype(str).str.strip()
            + " "
            + df[TIME_COL].astype(str).str.strip()
        )

        # Parse naive datetimes (CSR uses m/d/Y H:M:S)
        dt_naive = pd.to_datetime(
            combined,
            format="%m/%d/%Y %H:%M:%S",
            errors="coerce",
        )
        if dt_naive.isna().any():
            bad = combined[dt_naive.isna()].head().tolist()
            raise ValueError(
                "Unparsable Date/Time values (samples): " + ", ".join(map(str, bad))
            )

        # Detect source timezone from the file (e.g. "America/Los_Angeles")
        if TZ_COL not in df.columns:
            raise ValueError(f"Column '{TZ_COL}' not found in CSR file.")

        tz_values = df[TZ_COL].dropna().unique()
        if len(tz_values) == 0:
            raise ValueError("No non-empty values found in 'Time Zone' column.")
        src_tz_name = str(tz_values[0])
        try:
            src_tz = ZoneInfo(src_tz_name)
        except Exception as exc:  # noqa: BLE001
            raise ValueError(f"Unknown time zone {src_tz_name!r}: {exc}") from exc

        # Convert each timestamp using Python's zoneinfo directly ----------
        lt_list = []
        for ts in dt_naive:
            # ts is a pandas Timestamp (naive); convert to Python datetime
            py = ts.to_pydatetime().replace(tzinfo=src_tz)
            lt_list.append(py.astimezone(LT_TZ))

        return pd.Series(lt_list, index=df.index)


# ----------------------------------------------------------------------------
# Core logic
# ----------------------------------------------------------------------------

def run_one(src_path: Path, month: int, dest_dir: Path) -> list[Path]:
    """Process a single CSR CSV. Returns list of output file paths (0 or 1)."""
    df_all = pd.read_csv(src_path)

    cleaned = Cleaner(df_all, month).process()
    if cleaned.empty:
        # Nothing in that LT month; no output file
        return []

    # Always use the same base filename
    base_name = "Visos transakcijos pagal LT laiko zona.csv"
    out_path = _unique_path(dest_dir / base_name)
    cleaned.to_csv(out_path, index=False, encoding="utf-8-sig")
    return [out_path]


def run_batch(file_paths: list[Path], month: int, dest_dir: Path) -> tuple[list[Path], list[str]]:
    """Process multiple files; collect outputs and any error messages."""
    outputs: list[Path] = []
    errors: list[str] = []
    for p in file_paths:
        try:
            outputs.extend(run_one(p, month, dest_dir))
        except Exception as exc:  # noqa: BLE001
            errors.append(f"\u2716 {p.name}: {exc}")
    return outputs, errors


# ----------------------------------------------------------------------------
# Tiny Tkinter UI (multi-select + choose output folder) -----------------------
# ----------------------------------------------------------------------------

def main() -> None:
    root = tk.Tk()
    root.withdraw()  # dialogs only

    prefs = _load_prefs()

    # Pick one or many CSR CSV files (start from last-used input folder)
    file_paths = filedialog.askopenfilenames(
        title="Select one or more PayPal CSR CSV files",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        initialdir=prefs.get("last_input_dir"),
    )
    if not file_paths:
        return  # user cancelled

    # Remember where the user picked files from
    prefs["last_input_dir"] = str(Path(file_paths[0]).parent)

    # Ask for LT month number
    month = simpledialog.askinteger(
        "Pick month (LT time)",
        "Enter month number (1–12) to retain (based on LT time):",
        parent=root,
        minvalue=1,
        maxvalue=12,
    )
    if month is None:
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

    outputs, errors = run_batch([Path(p) for p in file_paths], month, dest_dir)

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