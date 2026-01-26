from __future__ import annotations

from pathlib import Path

import pandas as pd
import tkinter as tk
from tkinter import filedialog


def _pick_latest(files: list[Path], label: str) -> Path:
    if not files:
        raise FileNotFoundError(f"No {label} file found in {Path.cwd()}")
    if len(files) == 1:
        return files[0]
    latest = max(files, key=lambda p: p.stat().st_mtime)
    names = ", ".join(p.name for p in sorted(files))
    print(f"Multiple {label} files found ({names}). Using latest: {latest.name}")
    return latest


def _normalize_yes(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str).str.strip().str.lower().eq("yes")


def _select_file(title: str, initial_dir: Path) -> Path | None:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    file_path = filedialog.askopenfilename(
        title=title,
        initialdir=str(initial_dir),
        filetypes=[("Excel files", "*.xlsx;*.xls")],
    )
    root.destroy()
    return Path(file_path) if file_path else None


def _resolve_input_paths() -> tuple[Path, Path]:
    cwd = Path.cwd()
    sales_path = _select_file("Select Sales Compensation Report", cwd)
    people_path = _select_file("Select People file", cwd)
    if sales_path and people_path:
        return sales_path, people_path

    # Fallback to auto-discovery if user cancels either picker.
    sales_files = list(cwd.glob("Sales Compensation Report*.xls*"))
    people_files = list(cwd.glob("People*.xls*"))
    return _pick_latest(sales_files, "Sales Compensation Report"), _pick_latest(
        people_files, "People"
    )


def main() -> None:
    cwd = Path.cwd()
    sales_path, people_path = _resolve_input_paths()

    sales = pd.read_excel(sales_path, skiprows=3)
    required_sales_cols = {"Employee ID", "Active Status", "On Leave"}
    missing_sales = required_sales_cols - set(sales.columns)
    if missing_sales:
        raise KeyError(
            f"Missing columns in Sales report: {sorted(missing_sales)}. "
            f"Available columns: {list(sales.columns)}"
        )

    active_yes = _normalize_yes(sales["Active Status"])
    on_leave_yes = _normalize_yes(sales["On Leave"])
    active_not_on_leave = sales.loc[active_yes & ~on_leave_yes, "Employee ID"]
    active_ids = pd.to_numeric(active_not_on_leave, errors="coerce").dropna().astype("Int64")
    active_ids = active_ids.dropna().unique()

    people = pd.read_excel(people_path)
    if people.shape[1] < 105:
        raise ValueError(
            f"People file has {people.shape[1]} columns; need at least 105 to reach column DA."
        )

    employee_id_col = people.columns[0]
    status_col = people.columns[10]

    people_ids = pd.to_numeric(people[employee_id_col], errors="coerce").astype("Int64")
    status_loa = people[status_col].fillna("").astype(str).str.strip().str.upper().eq("LOA")
    valid_ids = people_ids.notna()
    in_sales = people_ids.isin(active_ids)
    mask = valid_ids & in_sales & status_loa

    selected_indices = [0, 6, 8, 9, 68, 10, 50, 104]
    selected_cols = [people.columns[i] for i in selected_indices]
    output = people.loc[mask, selected_cols].copy()
    output.columns = [
        "# Employee ID*",
        "First Name",
        "Last Name",
        "Region",
        "Country",
        "Employee Status",
        "Analyst_Name",
        "Plan_Type",
    ]

    output_path = cwd / "LOA return update.xlsx"
    output.to_excel(output_path, index=False)
    print(f"Saved {len(output)} rows to {output_path}")


if __name__ == "__main__":
    main()
