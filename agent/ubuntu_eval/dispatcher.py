"""
Hybrid comparison dispatcher for the Ubuntu port.

``compare_workbooks(gt_path, res_path, check_boards)`` is a drop-in replacement
for ``agent/utils/compare_sheets.compare_workbooks`` (the Windows/Excel-COM
function used by ``evaluation.py``).  It returns the same ``(report, success)``
shape and reproduces the same control flow:

* worksheet-count mismatch short-circuits to failure;
* each checked sheet's report is the union of per-category sub-reports;
* a sheet/task passes only if every requested property equals 1.

Category routing (the "hybrid" split chosen for this port):

* ``cells``, ``format_conditions``, ``filters``  -> openpyxl  (``oxl_engine``)
* ``view`` (frozen panes)                        -> openpyxl, with the original's
  workbook-level gating quirk preserved;
* ``charts``, ``pivot_tables``                   -> LibreOffice UNO (``uno_engine``),
  because they need recomputed formula/series/pivot data.

The UNO backend is only spun up when at least one checked sheet actually
requests charts or pivot tables, so pure-cell tasks never pay for LibreOffice.
"""

import openpyxl

from . import oxl_engine


def check_success(report):
    """Port of ``compare_sheets.check_success``: every leaf must be truthy."""
    if isinstance(report, dict):
        for value in report.values():
            if isinstance(value, (dict, list)):
                if not check_success(value):
                    return False
            else:
                if not value:
                    return False
    elif isinstance(report, list):
        for item in report:
            if not check_success(item):
                return False
    return True


def _init_sheet_report(config):
    """Build the initial per-sheet report: every requested property starts at 1
    (matched); engines flip properties to 0 on mismatch."""
    report = {}
    for category, properties in config.items():
        if not isinstance(properties, dict):
            continue
        for prop, need_check in properties.items():
            if need_check:
                report.setdefault(category, {})[prop] = 1
    return report


def _safe_load(path):
    """Load with openpyxl; if the file's XML is malformed (some Excel files use
    constructs openpyxl rejects but Excel/LibreOffice tolerate), repair it once
    by round-tripping through headless LibreOffice and retry."""
    try:
        return oxl_engine.load_values_workbook(path)
    except Exception:
        from .soffice import convert_to_xlsx

        repaired = convert_to_xlsx(path)
        return oxl_engine.load_values_workbook(repaired)


def _needs_uno(check_boards):
    """True if any checked sheet requests a chart or pivot-table property."""
    for config in check_boards.values():
        if not isinstance(config, dict):
            continue
        if any((config.get("charts") or {}).values()):
            return True
        if any((config.get("pivot_tables") or {}).values()):
            return True
    return False


def compare_workbooks(gt_path, res_path, check_boards, soffice_instance=None,
                      disable_uno=False):
    """Compare a ground-truth workbook against a result workbook.

    Parameters
    ----------
    gt_path, res_path : str
        Ground-truth and result ``.xlsx`` paths (``file1``/``file2`` in the
        original, i.e. GT is ``ws1`` and result is ``ws2``).
    check_boards : dict
        The ``check_board`` mapping from a ``*_check.yaml`` (sheet-id -> config).
    soffice_instance : SofficeInstance, optional
        A running LibreOffice instance to reuse for chart/pivot extraction.  If
        omitted and UNO is required, the per-process singleton is used.
    disable_uno : bool
        If True, skip chart/pivot-table comparison entirely (those properties
        stay "matched"); openpyxl-only, no LibreOffice.  Useful for fast debugging.

    Returns
    -------
    (report, success) : (dict, bool)
    """
    wb_gt = _safe_load(gt_path)
    wb_res = _safe_load(res_path)

    report = {}

    # Worksheet-count mismatch -> immediate failure (matches the original).
    if len(wb_gt.worksheets) != len(wb_res.worksheets):
        report["worksheet_count"] = False
        return report, False

    need_uno = (not disable_uno) and _needs_uno(check_boards)
    gt_objs = res_objs = None
    if need_uno:
        from . import uno_engine
        from .soffice import get_process_instance

        inst = soffice_instance or get_process_instance()
        doc_gt = inst.load(gt_path)
        doc_res = inst.load(res_path)
        try:
            gt_objs = uno_engine.extract_workbook_objects(doc_gt)
            res_objs = uno_engine.extract_workbook_objects(doc_res)
        finally:
            try:
                doc_gt.close(False)
            except Exception:
                pass
            try:
                doc_res.close(False)
            except Exception:
                pass

    # Compare each *existing* sheet that has a check-board entry.  The original
    # iterates sheets 1..count and looks the check-board up by index, so entries
    # referring to non-existent sheets are silently skipped -- we do the same.
    for sheet_idx in range(1, len(wb_gt.worksheets) + 1):
        sheet_id_str = str(sheet_idx)
        config = check_boards.get(sheet_id_str)
        if not isinstance(config, dict):
            continue
        ws_gt = wb_gt.worksheets[sheet_idx - 1]
        ws_res = wb_res.worksheets[sheet_idx - 1]

        sheet_report = _init_sheet_report(config)

        if any((config.get("cells") or {}).values()):
            oxl_engine.compare_cells(ws_gt, ws_res, sheet_report)

        if any((config.get("format_conditions") or {}).values()):
            oxl_engine.compare_format_conditions(ws_gt, ws_res, sheet_report)

        if any((config.get("filters") or {}).values()):
            oxl_engine.compare_filters(ws_gt, ws_res, sheet_report)

        if need_uno and any((config.get("charts") or {}).values()):
            from . import uno_engine

            uno_engine.compare_charts_on_sheet(
                gt_objs.get(sheet_idx, {}).get("charts", []),
                res_objs.get(sheet_idx, {}).get("charts", []),
                sheet_report,
            )

        if need_uno and any((config.get("pivot_tables") or {}).values()):
            from . import uno_engine

            uno_engine.compare_pivots_on_sheet(
                gt_objs.get(sheet_idx, {}).get("pivots", []),
                res_objs.get(sheet_idx, {}).get("pivots", []),
                sheet_report,
            )

        report[sheet_id_str] = sheet_report

    # Frozen panes: the original only evaluates them when *every* checked sheet
    # carries a 'view' sub-report (otherwise freeze stays matched).  Preserve it.
    if report and all("view" in sr for sr in report.values()):
        for sheet_id_str, sheet_report in report.items():
            if "freeze_pane" in sheet_report.get("view", {}):
                sheet_idx = int(sheet_id_str)
                oxl_engine.compare_view(
                    wb_gt.worksheets[sheet_idx - 1],
                    wb_res.worksheets[sheet_idx - 1],
                    sheet_report,
                )

    return report, check_success(report)
