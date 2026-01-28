from __future__ import annotations

from datetime import datetime
from pathlib import Path
import uuid

from flask import Flask, jsonify, render_template, request, send_from_directory
from werkzeug.utils import secure_filename
import pandas as pd

try:
    import olefile
except ImportError:  # pragma: no cover - optional dependency
    olefile = None

OLE_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
OUTPUT_NAME = "LOA return update.xlsx"
OUTPUT_ROOT = Path.cwd() / "web_outputs"

app = Flask(__name__)


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


def process_files(sales_path: Path, people_path: Path, output_dir: Path) -> tuple[Path, int]:
    sales = _read_excel(sales_path, skiprows=3)
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

    people = _read_excel(people_path)
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

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / OUTPUT_NAME
    output.to_excel(output_path, index=False)
    return output_path, len(output)


def _new_run_dir() -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_id = f"{stamp}_{uuid.uuid4().hex[:6]}"
    return OUTPUT_ROOT / run_id


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/process")
def process():
    sales_file = request.files.get("sales_file")
    people_file = request.files.get("people_file")
    if not sales_file or not people_file:
        return jsonify({"status": "error", "message": "Both files are required."}), 400

    run_dir = _new_run_dir()
    run_dir.mkdir(parents=True, exist_ok=True)

    sales_name = secure_filename(sales_file.filename or "sales.xlsx")
    people_name = secure_filename(people_file.filename or "people.xlsx")
    sales_path = run_dir / sales_name
    people_path = run_dir / people_name
    sales_file.save(sales_path)
    people_file.save(people_path)

    try:
        output_path, count = process_files(sales_path, people_path, run_dir)
    except Exception as exc:
        return jsonify({"status": "error", "message": str(exc)}), 400

    return jsonify(
        {
            "status": "ok",
            "row_count": count,
            "download_url": f"/download/{run_dir.name}",
            "output_path": str(output_path),
        }
    )


@app.get("/download/<run_id>")
def download(run_id: str):
    run_dir = OUTPUT_ROOT / run_id
    return send_from_directory(run_dir, OUTPUT_NAME, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=False)
