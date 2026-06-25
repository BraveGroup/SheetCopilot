"""
Ubuntu/Linux outcome-based evaluation backend for SheetCopilot.

This package is a drop-in, Excel-free re-implementation of the comparison logic
in ``agent/utils/compare_sheets.py`` (which requires Windows + Excel via
``win32com``).  It evaluates the *same* ``*_check.yaml`` check-boards against the
*same* result/ground-truth ``.xlsx`` files and produces the *same* metrics
(Exec@1, Pass@1, A_mean/A50/A90), but runs entirely on Ubuntu using:

* **openpyxl** for cells, conditional formatting, filters and frozen panes, and
* **headless LibreOffice via UNO** for charts and pivot tables.

The original Windows scripts are left untouched.
"""

from .dispatcher import compare_workbooks, check_success

__all__ = ["compare_workbooks", "check_success"]
