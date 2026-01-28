from __future__ import annotations

from pathlib import Path
import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

try:
    import olefile
except ImportError:  # pragma: no cover - optional dependency
    olefile = None

OLE_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def _normalize_yes(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str).str.strip().str.lower().eq("yes")


def _is_ole_file(path: Path) -> bool:
    try:
        with path.open("rb") as handle:
            return handle.read(8) == OLE_SIGNATURE
    except OSError:
        return False


def _is_encrypted_ole(path: Path) -> bool:
    if olefile is None or not _is_ole_file(path):
        return False
    try:
        ole = olefile.OleFileIO(path)
        streams = ole.listdir()
        ole.close()
        return any("EncryptedPackage" in stream for stream in streams) or any(
            "DRMEncrypted" in stream for stream in streams
        )
    except Exception:
        return False


def _read_excel(path: Path, **kwargs) -> pd.DataFrame:
    if _is_encrypted_ole(path):
        raise ValueError(
            "The file appears to be protected by sensitivity labels or IRM. "
            "Open it in Excel and save an unprotected copy, then retry."
        )
    try:
        return pd.read_excel(path, **kwargs)
    except Exception:
        if _is_ole_file(path):
            return pd.read_excel(path, engine="xlrd", **kwargs)
        raise


def _latest_match(pattern: str) -> Path | None:
    matches = list(Path.cwd().glob(pattern))
    if not matches:
        return None
    return max(matches, key=lambda p: p.stat().st_mtime)


def process_files(
    sales_path: Path,
    people_path: Path,
    output_dir: Path,
    progress_cb,
) -> tuple[Path, int]:
    progress_cb(5, "Reading Sales Compensation Report...")
    sales = _read_excel(sales_path, skiprows=3)
    required_sales_cols = {"Employee ID", "Active Status", "On Leave"}
    missing_sales = required_sales_cols - set(sales.columns)
    if missing_sales:
        raise KeyError(
            f"Missing columns in Sales report: {sorted(missing_sales)}. "
            f"Available columns: {list(sales.columns)}"
        )

    progress_cb(25, "Filtering active employees...")
    active_yes = _normalize_yes(sales["Active Status"])
    on_leave_yes = _normalize_yes(sales["On Leave"])
    active_not_on_leave = sales.loc[active_yes & ~on_leave_yes, "Employee ID"]
    active_ids = pd.to_numeric(active_not_on_leave, errors="coerce").dropna().astype("Int64")
    active_ids = active_ids.dropna().unique()

    progress_cb(45, "Reading People file...")
    people = _read_excel(people_path)
    if people.shape[1] < 105:
        raise ValueError(
            f"People file has {people.shape[1]} columns; need at least 105 to reach column DA."
        )

    progress_cb(65, "Filtering LOA records...")
    employee_id_col = people.columns[0]
    status_col = people.columns[10]

    people_ids = pd.to_numeric(people[employee_id_col], errors="coerce").astype("Int64")
    status_loa = people[status_col].fillna("").astype(str).str.strip().str.upper().eq("LOA")
    valid_ids = people_ids.notna()
    in_sales = people_ids.isin(active_ids)
    mask = valid_ids & in_sales & status_loa

    selected_indices = [0, 6, 8, 9, 68, 10, 12, 50, 104]
    selected_cols = [people.columns[i] for i in selected_indices]
    output = people.loc[mask, selected_cols].copy()
    output.columns = [
        "# Employee ID*",
        "First Name",
        "Last Name",
        "Region",
        "Country",
        "Employee Status",
        "Termination Date",
        "Analyst_Name",
        "Plan_Type",
    ]

    progress_cb(85, "Writing output file...")
    output_path = output_dir / "LOA return update.xlsx"
    output.to_excel(output_path, index=False)
    progress_cb(100, f"Saved {len(output)} rows.")
    return output_path, len(output)


class App:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("LOA Return Update")
        self.root.resizable(False, False)

        self.sales_var = tk.StringVar()
        self.people_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready.")
        self.count_var = tk.StringVar(value="-")
        self.output_var = tk.StringVar(value="-")
        self.progress_var = tk.DoubleVar(value=0)
        self.last_output_path: Path | None = None

        self._queue: queue.Queue[tuple] = queue.Queue()
        self._running = False

        self._build_ui()
        self._prefill_paths()

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=12)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)

        ttk.Label(frame, text="Sales Compensation Report").grid(
            row=0, column=0, sticky="w"
        )
        self.sales_entry = ttk.Entry(
            frame, textvariable=self.sales_var, width=80, state="readonly"
        )
        self.sales_entry.grid(row=1, column=0, sticky="we", pady=(2, 8))
        self.sales_button = ttk.Button(
            frame, text="Browse...", command=self._browse_sales
        )
        self.sales_button.grid(row=1, column=1, padx=(8, 0))

        ttk.Label(frame, text="People file").grid(row=2, column=0, sticky="w")
        self.people_entry = ttk.Entry(
            frame, textvariable=self.people_var, width=80, state="readonly"
        )
        self.people_entry.grid(row=3, column=0, sticky="we", pady=(2, 8))
        self.people_button = ttk.Button(
            frame, text="Browse...", command=self._browse_people
        )
        self.people_button.grid(row=3, column=1, padx=(8, 0))

        ttk.Label(frame, text="Progress").grid(row=4, column=0, sticky="w")
        self.progress = ttk.Progressbar(
            frame, variable=self.progress_var, maximum=100, mode="determinate"
        )
        self.progress.grid(row=5, column=0, columnspan=2, sticky="we", pady=(2, 4))
        ttk.Label(frame, textvariable=self.status_var).grid(
            row=6, column=0, columnspan=2, sticky="w"
        )

        ttk.Label(frame, text="Rows").grid(row=7, column=0, sticky="w", pady=(8, 0))
        ttk.Label(frame, textvariable=self.count_var).grid(
            row=7, column=1, sticky="w", pady=(8, 0)
        )

        ttk.Label(frame, text="Output file").grid(row=8, column=0, sticky="w")
        ttk.Label(frame, textvariable=self.output_var, wraplength=620).grid(
            row=9, column=0, columnspan=2, sticky="w"
        )

        self.open_button = ttk.Button(
            frame, text="Open output file", command=self._open_output, state="disabled"
        )
        self.open_button.grid(row=10, column=0, columnspan=2, pady=(8, 0))

        self.run_button = ttk.Button(frame, text="Run", command=self._run)
        self.run_button.grid(row=11, column=0, columnspan=2, pady=(6, 0))

    def _prefill_paths(self) -> None:
        sales_path = _latest_match("Sales Compensation Report*.xls*")
        people_path = _latest_match("People*.xls*")
        if sales_path:
            self.sales_var.set(str(sales_path))
        if people_path:
            self.people_var.set(str(people_path))

    def _browse_sales(self) -> None:
        path = filedialog.askopenfilename(
            title="Select Sales Compensation Report",
            initialdir=str(Path.cwd()),
            filetypes=[("Excel files", "*.xlsx;*.xls")],
        )
        if path:
            self.sales_var.set(path)

    def _browse_people(self) -> None:
        path = filedialog.askopenfilename(
            title="Select People file",
            initialdir=str(Path.cwd()),
            filetypes=[("Excel files", "*.xlsx;*.xls")],
        )
        if path:
            self.people_var.set(path)

    def _set_running(self, running: bool) -> None:
        state = "disabled" if running else "normal"
        self.run_button.configure(state=state)
        self.sales_entry.configure(state="readonly" if not running else "disabled")
        self.people_entry.configure(state="readonly" if not running else "disabled")
        self.sales_button.configure(state=state)
        self.people_button.configure(state=state)
        if running:
            self.open_button.configure(state="disabled")
        else:
            self.open_button.configure(
                state="normal" if self.last_output_path else "disabled"
            )
        self._running = running

    def _queue_progress(self, value: int, message: str) -> None:
        self._queue.put(("progress", value, message))

    def _worker(self, sales_path: Path, people_path: Path) -> None:
        try:
            output_path, count = process_files(
                sales_path, people_path, Path.cwd(), self._queue_progress
            )
            self._queue.put(("done", output_path, count))
        except Exception as exc:
            self._queue.put(("error", str(exc)))

    def _poll_queue(self) -> None:
        while True:
            try:
                message = self._queue.get_nowait()
            except queue.Empty:
                break
            kind = message[0]
            if kind == "progress":
                _, value, status = message
                self.progress_var.set(value)
                self.status_var.set(status)
            elif kind == "done":
                _, output_path, count = message
                self.progress_var.set(100)
                self.status_var.set("Completed.")
                self.count_var.set(str(count))
                self.output_var.set(str(output_path))
                self.last_output_path = Path(output_path)
                self._set_running(False)
            elif kind == "error":
                _, error_message = message
                self.status_var.set("Error.")
                self.last_output_path = None
                self._set_running(False)
                messagebox.showerror("Error", error_message)

        if self._running:
            self.root.after(100, self._poll_queue)

    def _run(self) -> None:
        sales_value = self.sales_var.get().strip()
        people_value = self.people_var.get().strip()
        if not sales_value or not people_value:
            messagebox.showwarning("Missing files", "Select both input files first.")
            return

        sales_path = Path(sales_value)
        people_path = Path(people_value)
        if not sales_path.exists():
            messagebox.showerror("File not found", f"Sales file not found:\n{sales_path}")
            return
        if not people_path.exists():
            messagebox.showerror("File not found", f"People file not found:\n{people_path}")
            return

        self.progress_var.set(0)
        self.status_var.set("Starting...")
        self.count_var.set("-")
        self.output_var.set("-")
        self.last_output_path = None
        self._set_running(True)

        thread = threading.Thread(
            target=self._worker, args=(sales_path, people_path), daemon=True
        )
        thread.start()
        self.root.after(100, self._poll_queue)

    def run(self) -> None:
        self.root.mainloop()

    def _open_output(self) -> None:
        if self.last_output_path is None:
            messagebox.showwarning("No output", "No output file is available yet.")
            return
        if not self.last_output_path.exists():
            messagebox.showerror(
                "File not found",
                f"Output file not found:\n{self.last_output_path}",
            )
            return
        try:
            os.startfile(self.last_output_path)
        except OSError as exc:
            messagebox.showerror("Open failed", str(exc))


def main() -> None:
    App().run()


if __name__ == "__main__":
    main()
