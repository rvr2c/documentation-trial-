"""Progress-tracking utilities for the optimisation pipeline.

Maintains the per-campaign progress table (one row per trial) that the
optimiser, CAD stage and cluster stage share. The table is stored as
``progress_live.csv`` (source of truth) with an optional Excel snapshot.

Public API:
    now_str, progress_fieldnames, get_runs_dir, get_progress_paths,
    atomic_write_csv, read_csv_rows, write_xlsx_snapshot, normalise_job_id,
    load_progress_only, save_progress, set_row, count_active
"""
from __future__ import annotations

import csv
import os
import time
from pathlib import Path

from openpyxl import Workbook

# Folder containing this file, and the project root (three levels above it).
UTILS_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = UTILS_DIR.parent.parent.parent


def now_str() -> str:
    """Return the current local time as a timestamp string.

    Args:
        None.

    Returns:
        str formatted as ``YYYY-MM-DD HH:MM:SS``.
    """
    return time.strftime("%Y-%m-%d %H:%M:%S")


def progress_fieldnames(param_names: list[str], metric_names: list[str]) -> list[str]:
    """Build the ordered list of column names for the progress table.

    Column order is: fixed bookkeeping columns, then metric columns, then
    parameter columns.

    Args:
        param_names: Names of the design parameters (one column each).
        metric_names: Names of the objective metrics (one column each).

    Returns:
        list[str] of all column names.
    """
    base_fields = [
        "trial_id",
        "case_name",
        "varied_param",
        "level",
        "status",
        "retry_count",
        "success",
        "state",
        "reason",
        "job_id",
        "jobname",
        "remote_dir",
        "cluster_name",
        "cluster_user",
        "cluster_host",
        "cluster_remote_base",
        "cluster_base_dir",
        "geom_source_path",
        "geom_xt_path",
        "created_at",
        "updated_at",
    ]

    return base_fields + metric_names + param_names


def get_progress_paths(username: str, campaign_id: str) -> tuple[Path, Path]:
    """Get the progress CSV and Excel file paths for a campaign.

    Also creates the campaign's ``runs`` folder if it does not exist.
    Note: the Excel name here is ``progress_live.xlsx``, whereas
    optimisation.py uses ``progress_open.xlsx``.

    Args:
        username: Owner of the campaign.
        campaign_id: Campaign identifier.

    Returns:
        Tuple of two Paths: (progress_live.csv, progress_live.xlsx).
    """
    campaign_dir = get_runs_dir(username, campaign_id)
    return campaign_dir / "progress_live.csv", campaign_dir / "progress_live.xlsx"


def get_runs_dir(username: str, campaign_id: str) -> Path:
    """Get (and create if needed) a campaign's ``runs`` folder.

    Path: ``<PROJECT_ROOT>/storage/<username>/campaigns/<campaign_id>/runs``.

    Args:
        username: Owner of the campaign.
        campaign_id: Campaign identifier.

    Returns:
        Path to the runs folder.
    """
    campaign_dir = PROJECT_ROOT / "storage" / username / "campaigns" / campaign_id / "runs"
    campaign_dir.mkdir(parents=True, exist_ok=True)
    return campaign_dir


def atomic_write_csv(path: Path, fieldnames: list[str], rows: list[dict]) -> None:
    """Write a CSV file safely, avoiding half-written files.

    Writes to ``<path>.tmp`` first, then swaps it into place with
    ``os.replace``. If the swap fails (e.g. the file is open in Excel) it
    retries up to 12 times, 0.5 s apart. If it still fails it writes
    directly to ``path`` and removes the temp file. If that also fails, the
    last swap error is raised.

    Args:
        path: Destination CSV path (parent folders are created).
        fieldnames: Column names, in order.
        rows: List of row dicts keyed by column name.

    Returns:
        None.

    Raises:
        PermissionError | OSError: If the file cannot be replaced or written.
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")

    def _write_csv(target: Path) -> None:
        """Write the header and all rows to ``target``."""
        with target.open("w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(rows)

    _write_csv(tmp)

    last_err = None
    for _ in range(12):
        try:
            os.replace(tmp, path)
            return
        except (PermissionError, OSError) as exc:
            last_err = exc
            time.sleep(0.5)

    try:
        _write_csv(path)
        try:
            if tmp.exists():
                tmp.unlink()
        except OSError:
            pass
        return
    except Exception:
        pass

    if last_err is not None:
        raise last_err


def read_csv_rows(path: Path) -> list[dict]:
    """Read a CSV file into a list of row dicts.

    Args:
        path: CSV file to read.

    Returns:
        list[dict] with one dict per row (all values are strings). Returns
        an empty list if the file does not exist.
    """
    if not path.exists():
        return []
    with path.open(newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def write_xlsx_snapshot(path: Path, sheet_name: str, fieldnames: list[str], rows: list[dict]) -> None:
    """Write the table to an Excel workbook (a read-only style snapshot).

    Missing values are written as empty strings. If saving fails with a
    PermissionError (e.g. the workbook is open in Excel) it retries up to
    12 times, 0.5 s apart, then raises the last error.

    Args:
        path: Destination ``.xlsx`` path.
        sheet_name: Name of the worksheet.
        fieldnames: Column headers, in order.
        rows: List of row dicts keyed by column name.

    Returns:
        None.

    Raises:
        PermissionError: If the file stays locked after all retries.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    ws.append(fieldnames)
    for row in rows:
        ws.append([row.get(col, "") for col in fieldnames])

    last_err = None
    for _ in range(12):
        try:
            wb.save(path)
            return
        except PermissionError as exc:
            last_err = exc
            time.sleep(0.5)

    if last_err is not None:
        raise last_err


def normalise_job_id(job_id) -> str:
    """Clean a cluster job ID into a plain string.

    Handles values that spreadsheets or pandas have mangled: None, empty
    text and "nan" become ``""``, and a float-style ID such as ``"12345.0"``
    becomes ``"12345"``.

    Args:
        job_id: Raw job ID (str, int, float, or None).

    Returns:
        The cleaned job ID as str (empty string if there is none).
    """
    s = "" if job_id is None else str(job_id).strip()
    if not s or s.lower() == "nan":
        return ""
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s


def load_progress_only(progress_csv: Path, param_names: list[str], metric_names: list[str]) -> dict[int, dict]:
    """Load the progress table from CSV into a dict keyed by trial ID.

    Any expected column missing from a row is added as an empty string, and
    ``job_id`` is cleaned with ``normalise_job_id``.

    Args:
        progress_csv: Path to ``progress_live.csv``.
        param_names: Design parameter names (to build the expected columns).
        metric_names: Objective metric names (to build the expected columns).

    Returns:
        dict mapping trial_id (int) -> row dict. Empty dict if the file is
        missing or has no rows.

    Raises:
        ValueError | KeyError: If a row has a missing or non-integer
            ``trial_id``.
    """
    rows = read_csv_rows(progress_csv)
    if not rows:
        return {}

    expected_fields = progress_fieldnames(param_names, metric_names)
    progress: dict[int, dict] = {}

    for row in rows:
        trial_id = int(row["trial_id"])
        for field in expected_fields:
            row.setdefault(field, "")
        row["job_id"] = normalise_job_id(row.get("job_id", ""))
        progress[trial_id] = row

    return progress


def save_progress(
    progress_csv: Path,
    progress_xlsx: Path,
    progress: dict[int, dict],
    param_names: list[str],
    metric_names: list[str],
) -> None:
    """Save the progress table to CSV and refresh the Excel snapshot.

    Rows are written sorted by trial ID. The CSV is written atomically. A
    failure to write the Excel file is only printed as a warning, so it
    never stops the pipeline.

    Args:
        progress_csv: Destination CSV path.
        progress_xlsx: Destination Excel path.
        progress: dict mapping trial_id -> row dict.
        param_names: Design parameter names (columns).
        metric_names: Objective metric names (columns).

    Returns:
        None.
    """
    fieldnames = progress_fieldnames(param_names, metric_names)
    rows = [progress[k] for k in sorted(progress)]

    atomic_write_csv(progress_csv, fieldnames, rows)

    try:
        write_xlsx_snapshot(progress_xlsx, "progress", fieldnames, rows)
    except Exception as exc:
        print(f"[WARN] Could not refresh {progress_xlsx.name}: {exc}")


def set_row(progress: dict[int, dict], trial_id: int, **updates) -> None:
    """Update fields of one trial's row in place.

    Every value is stored as a string (None becomes ""). A ``job_id`` value
    is cleaned with ``normalise_job_id``. ``updated_at`` is always set to
    the current time.

    Args:
        progress: dict mapping trial_id -> row dict (modified in place).
        trial_id: ID of the row to update. It must already exist.
        **updates: Column name = new value pairs, e.g.
            ``status="RUNNING", job_id=123``.

    Returns:
        None.

    Raises:
        KeyError: If ``trial_id`` is not in ``progress``.
    """
    row = progress[trial_id]
    for key, value in updates.items():
        if key == "job_id":
            value = normalise_job_id(value)
        row[key] = "" if value is None else str(value)
    row["updated_at"] = now_str()


def count_active(progress: dict[int, dict]) -> int:
    """Count trials currently on the cluster.

    A trial counts as active if its status is SUBMITTED or RUNNING.

    Args:
        progress: dict mapping trial_id -> row dict.

    Returns:
        int number of active trials.
    """
    return sum(
        1
        for row in progress.values()
        if row.get("status", "") in {"SUBMITTED", "RUNNING"}
    )
