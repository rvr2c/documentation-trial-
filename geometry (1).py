"""SolidWorks geometry generation for the optimisation pipeline.

For each optimisation trial this module:
    1. Copies the campaign's master SolidWorks part into a trial folder.
    2. Writes the trial's parameter values into the part's Equation Manager
       (global variables).
    3. Rebuilds the model and exports it as a Parasolid (.x_t) file.

Requirements: Windows, SolidWorks installed, and the ``pywin32`` package
(provides ``pythoncom`` and ``win32com``).

Public API:
    generate_geom(...) -> dict

All functions whose names start with an underscore are internal helpers.
"""
from __future__ import annotations

import os
import shutil
import time
from pathlib import Path

import pythoncom
import win32com.client as win32

# Folder containing this file, and the project root (three levels above it).
UTILS_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = UTILS_DIR.parent.parent.parent

# SolidWorks API constants.
SW_DOC_PART = 1      # swDocPART: document type "part"
SW_OPEN_SILENT = 1   # swOpenDocOptions_Silent: open without dialogs


def _ensure_exists(path: Path, label: str) -> None:
    """Raise an error if a required path does not exist.

    Args:
        path: Path that must exist.
        label: Human-readable name used in the error message.

    Returns:
        None.

    Raises:
        FileNotFoundError: If ``path`` does not exist.
    """
    if not path.exists():
        raise FileNotFoundError(f"{label} not found: {path}")


def _trial_paths(campaign_dir: Path, iteration: int) -> tuple[Path, Path, Path]:
    """Build (and create) the folder layout for one trial.

    Creates ``<campaign_dir>/runs/trial_NNNN/`` if it does not exist.

    Args:
        campaign_dir: Root folder of the campaign.
        iteration: Trial number, zero-padded to 4 digits in file names.

    Returns:
        Tuple of three Paths:
            trial_dir: The trial folder.
            trial_part_path: ``trial_NNNN.SLDPRT`` inside the trial folder.
            export_xt_path: ``geom_iter_NNNN.x_t`` inside the trial folder.
    """
    trial_dir = campaign_dir / "runs" / f"trial_{iteration:04d}"
    trial_dir.mkdir(parents=True, exist_ok=True)

    trial_part_path = trial_dir / f"trial_{iteration:04d}.SLDPRT"
    export_xt_path = trial_dir / f"geom_iter_{iteration:04d}.x_t"

    return trial_dir, trial_part_path, export_xt_path


def _format_equation(name: str, value: float | int) -> str:
    """Format a SolidWorks global-variable equation string.

    Example: ``_format_equation("width", 12.5)`` -> ``'"width" = 12.500000'``.

    Args:
        name: Global variable name.
        value: Numeric value (converted to float, 6 decimal places).

    Returns:
        The equation text as a str.
    """
    return f'"{name}" = {float(value):.6f}'


def _find_equation_index(eq_mgr, variable_name: str, n_eqs: int) -> int:
    """Find the index of the equation that mentions a global variable.

    Searches each equation for the variable name wrapped in double quotes.
    Note: this is a substring match, so an equation that merely references
    the variable on its right-hand side could match first.

    Args:
        eq_mgr: SolidWorks Equation Manager object.
        variable_name: Name of the global variable to find.
        n_eqs: Number of equations in the manager.

    Returns:
        Index (int) of the first matching equation.

    Raises:
        ValueError: If no equation contains the variable.
    """
    needle = f'"{variable_name}"'
    for i in range(n_eqs):
        eq = str(eq_mgr.Equation(i))
        if needle in eq:
            return i
    raise ValueError(f"Global variable '{variable_name}' not found in Equation Manager.")


def _open_solidworks():
    """Start (or attach to) SolidWorks through COM.

    Initialises COM for the current thread and tries to make the window
    visible (errors setting visibility are ignored).

    Returns:
        The SolidWorks application COM object (``SldWorks.Application``).
    """
    pythoncom.CoInitialize()
    sw_app = win32.Dispatch("SldWorks.Application")
    try:
        sw_app.Visible = True
    except Exception:
        pass
    return sw_app


def _open_part(sw_app, part_path: Path):
    """Open a SolidWorks part file silently.

    Args:
        sw_app: SolidWorks application object from ``_open_solidworks``.
        part_path: Path to the ``.SLDPRT`` file.

    Returns:
        Tuple (model, errors, warnings):
            model: The opened document COM object.
            errors: Integer error code from OpenDoc6.
            warnings: Integer warning code from OpenDoc6.

    Raises:
        RuntimeError: If SolidWorks fails to open the part.
    """
    errors = win32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    warnings = win32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)

    model = sw_app.OpenDoc6(
        str(part_path),
        SW_DOC_PART,
        SW_OPEN_SILENT,
        "",
        errors,
        warnings,
    )

    if model is None:
        raise RuntimeError(f"Failed to open SolidWorks part: {part_path}")

    return model, int(errors.value), int(warnings.value)


def _update_global_variables(model, params: dict[str, float]) -> None:
    """Overwrite global variables in the part's Equation Manager.

    For each entry in ``params`` the matching equation is located and
    replaced with ``"name" = value``. Old and new equations are printed.

    Args:
        model: Open SolidWorks part document.
        params: Mapping of global variable name -> new numeric value.

    Returns:
        None.

    Raises:
        RuntimeError: If the Equation Manager cannot be accessed.
        ValueError: If a variable in ``params`` is not in the part.
    """
    eq_mgr = model.GetEquationMgr
    if eq_mgr is None:
        raise RuntimeError("Could not access SolidWorks Equation Manager.")

    n_eqs = int(eq_mgr.GetCount)

    for name, value in params.items():
        idx = _find_equation_index(eq_mgr, name, n_eqs)
        old_eq = str(eq_mgr.Equation(idx))
        new_eq = _format_equation(name, value)

        eq_mgr.Equation(idx, new_eq)
        print(f"[CAD] {name}: {old_eq} -> {new_eq}")


def _rebuild_model(model) -> None:
    """Rebuild the model so the new variable values take effect.

    Calls ``EditRebuild3`` (handling COM exposing it as a property instead
    of a method), then ``ForceRebuild3(False)`` with errors ignored. A
    1-second pause follows each step to let SolidWorks finish.

    Args:
        model: Open SolidWorks part document.

    Returns:
        None. The rebuild result is printed.
    """
    try:
        ok = model.EditRebuild3()
    except TypeError:
        ok = model.EditRebuild3

    print(f"[CAD] Rebuild result: {ok}")
    time.sleep(1.0)

    try:
        model.ForceRebuild3(False)
    except Exception:
        pass

    time.sleep(1.0)


def _export_parasolid(model, export_path: Path) -> Path:
    """Export the model as a Parasolid (.x_t) file.

    Creates the parent folder, deletes any existing file, then tries
    ``SaveAs3`` and falls back to ``SaveAs2``. Verifies the file exists
    after a 1-second wait.

    Args:
        model: Open SolidWorks part document.
        export_path: Destination path for the ``.x_t`` file.

    Returns:
        The ``export_path`` that was written.

    Raises:
        RuntimeError: If both save calls fail, or the file was not created.
    """
    export_path.parent.mkdir(parents=True, exist_ok=True)

    if export_path.exists():
        export_path.unlink()

    save_ok = None

    try:
        save_ok = model.SaveAs3(str(export_path), 0, 0)
    except Exception:
        try:
            save_ok = model.SaveAs2(str(export_path), 0, True, True)
        except Exception as e:
            raise RuntimeError(f"Parasolid export call failed: {e}")

    time.sleep(1.0)

    if not export_path.exists():
        raise RuntimeError(f"Parasolid export failed; file not created: {export_path}")

    print(f"[CAD] Exported Parasolid: {export_path}")
    print(f"[CAD] Save return value: {save_ok}")
    return export_path


def _close_doc(sw_app, model) -> None:
    """Close an open SolidWorks document, ignoring any errors.

    The document is closed by its title. SolidWorks itself is left running.

    Args:
        sw_app: SolidWorks application object.
        model: The document to close.

    Returns:
        None.
    """
    try:
        title = model.GetTitle()
    except TypeError:
        title = model.GetTitle

    try:
        sw_app.CloseDoc(title)
    except Exception:
        pass


def generate_geom(
    username: str,
    campaign_id: str,
    master_part_name: str,
    iteration: int,
    params: dict[str, float],
    keep_trial_part: bool = True,
) -> dict:
    """Generate one trial's geometry from the master SolidWorks part.

    The master part is never modified: it is copied into
    ``runs/trial_NNNN/`` and the copy is edited and exported.

    Args:
        username: Owner of the campaign (folder under ``storage/``).
        campaign_id: Campaign identifier (folder under the user's campaigns).
        master_part_name: File name of the master ``.SLDPRT`` inside the
            campaign's ``inputs`` folder.
        iteration: Trial number, used to name the trial folder and files.
        params: Mapping of SolidWorks global variable name -> value to set.
        keep_trial_part: If True (default) keep the modified trial
            ``.SLDPRT``; if False delete it after export.

    Returns:
        dict with two absolute-path strings:
            "geom_xt_path": the exported Parasolid file.
            "trial_part_path": the modified trial part file.

    Raises:
        FileNotFoundError: If the master part does not exist.
        RuntimeError: If SolidWorks cannot open the part, the Equation
            Manager is unavailable, or the export fails.
        ValueError: If a name in ``params`` is not a global variable.
    """
    # Path layout: storage/<username>/campaigns/<campaign_id>
    campaign_dir = PROJECT_ROOT / "storage" / username / "campaigns" / campaign_id
    master_part_path = campaign_dir / "inputs" / master_part_name

    _ensure_exists(master_part_path, "Master SolidWorks part")

    trial_dir, trial_part_path, export_xt_path = _trial_paths(campaign_dir, iteration)

    print("=" * 80)
    print(f"[CAD] Trial {iteration:04d}")
    print(f"[CAD] Master part : {master_part_path}")
    print(f"[CAD] Trial dir : {trial_dir}")
    print(f"[CAD] Trial part : {trial_part_path}")
    print(f"[CAD] Export path : {export_xt_path}")
    print("=" * 80)

    shutil.copy2(master_part_path, trial_part_path)
    print(f"[CAD] Copied master part -> {trial_part_path}")

    sw_app = None
    model = None

    try:
        sw_app = _open_solidworks()
        model, open_errors, open_warnings = _open_part(sw_app, trial_part_path)

        print(f"[CAD] Open errors : {open_errors}")
        print(f"[CAD] Open warnings : {open_warnings}")

        _update_global_variables(model, params)
        _rebuild_model(model)

        try:
            model.Save3(1, 0, 0)
        except Exception:
            pass

        export_path = _export_parasolid(model, export_xt_path)

        return {
            "geom_xt_path": str(export_path.resolve()),
            "trial_part_path": str(trial_part_path.resolve()),
        }
    finally:
        if sw_app is not None and model is not None:
            _close_doc(sw_app, model)

        if not keep_trial_part and trial_part_path.exists():
            try:
                trial_part_path.unlink()
                print(f"[CAD] Deleted temporary trial part: {trial_part_path}")
            except Exception as e:
                print(f"[CAD] Warning: could not delete trial part: {e}")
