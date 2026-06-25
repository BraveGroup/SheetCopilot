"""
LibreOffice-UNO comparison engine for charts and pivot tables.

These two object categories require recomputed formula/series/pivot data, which
openpyxl cannot provide, so the dispatcher routes them here.  We extract a plain
Python *signature* for every chart and pivot table in a loaded LibreOffice
document (no live UNO objects escape this module), then compare the signatures
with the exact logic of the Windows engine
(``compare_sheets.compare_charts`` / ``compare_pivot_tables`` and their
worksheet-level matching loops).

Faithfulness note: both the ground-truth and result workbooks are loaded by the
*same* LibreOffice build, so a chart's extracted type/series and a pivot's
fields are represented identically for equivalent Excel objects -- exactly the
property exploited throughout this port.
"""

import numpy as np

TOLERANCE = 1e-8


# --------------------------------------------------------------------------- #
# Extraction
# --------------------------------------------------------------------------- #
def _norm_cat(v):
    if isinstance(v, str):
        return v.strip()
    return v


def _extract_one_chart(model):
    """Build a chart signature from a chart *model* (works for both a table
    chart's ``getEmbeddedObject()`` and a DrawPage OLE chart's ``.Model``)."""
    sig = {
        "type": None,
        "has_title": False,
        "title": "",
        "has_legend": False,
        "legend_align": None,
        "axes": None,
        "categories": (),
        "series": [],
    }

    # Chart type via the chart2 model (clean, e.g. 'ColumnChartType'),
    # disambiguated with old-API diagram orientation/stacking flags.
    diag = model.Diagram
    try:
        d2 = model.getFirstDiagram()
        types = tuple(
            ct.getChartType()
            for cs in d2.getCoordinateSystems()
            for ct in cs.getChartTypes()
        )
    except Exception:
        types = ()
    flags = tuple(
        bool(getattr(diag, attr, False))
        for attr in ("Vertical", "Stacked", "Percent", "Dim3D", "Deep")
    )
    sig["type"] = (types, flags)

    # Title
    try:
        sig["has_title"] = bool(model.HasMainTitle)
        if sig["has_title"]:
            sig["title"] = model.Title.String
    except Exception:
        pass

    # Legend (+ position only when a legend exists)
    try:
        sig["has_legend"] = bool(model.HasLegend)
        if sig["has_legend"]:
            sig["legend_align"] = getattr(model.Legend, "Alignment", None)
    except Exception:
        pass

    # Axes presence + titles (count- and title-equivalent of the original)
    sig["axes"] = tuple(
        bool(getattr(diag, attr, False))
        for attr in (
            "HasXAxis", "HasYAxis", "HasZAxis",
            "HasSecondaryXAxis", "HasSecondaryYAxis",
            "HasXAxisTitle", "HasYAxisTitle",
        )
    )

    # Series + categories from the (recomputed) chart data
    try:
        data = model.Data
        col_desc = list(data.ColumnDescriptions)
        sig["categories"] = tuple(_norm_cat(x) for x in data.RowDescriptions)
        matrix = data.Data  # rows x cols
        ncols = len(col_desc)
        series = []
        for j in range(ncols):
            col_vals = tuple(matrix[i][j] for i in range(len(matrix)))
            try:
                marker = diag.getDataRowProperties(j).SymbolType
            except Exception:
                marker = None
            series.append((col_vals, marker))
        sig["series"] = series
    except Exception:
        pass

    return sig


def _extract_sheet_charts(sheet):
    charts = []
    # Primary: table charts.
    try:
        tcharts = sheet.Charts
        for nm in tcharts.ElementNames:
            charts.append(_extract_one_chart(tcharts.getByName(nm).getEmbeddedObject()))
    except Exception:
        pass
    # Fallback: charts imported as DrawPage OLE shapes.
    if not charts:
        try:
            dp = sheet.DrawPage
            for i in range(dp.Count):
                shape = dp.getByIndex(i)
                model = getattr(shape, "Model", None)
                if model is not None and hasattr(model, "getFirstDiagram"):
                    charts.append(_extract_one_chart(model))
        except Exception:
            pass
    return charts


def _func_str(fn):
    return str(getattr(fn, "value", fn))


def _extract_sheet_pivots(sheet):
    pivots = []
    try:
        dps = sheet.DataPilotTables
    except Exception:
        return pivots
    for nm in dps.ElementNames:
        try:
            dp = dps.getByName(nm)
            src = dp.SourceRange
            source = (src.Sheet, src.StartColumn, src.StartRow, src.EndColumn, src.EndRow)
            row_fields = [dp.RowFields.getByIndex(i).Name for i in range(dp.RowFields.Count)]
            col_fields = [dp.ColumnFields.getByIndex(i).Name for i in range(dp.ColumnFields.Count)]
            data_fields = [
                (dp.DataFields.getByIndex(i).Name, _func_str(dp.DataFields.getByIndex(i).Function))
                for i in range(dp.DataFields.Count)
            ]
            try:
                n_fields = dp.DataPilotFields.Count
            except Exception:
                n_fields = len(row_fields) + len(col_fields) + len(data_fields)
            pivots.append({
                "source": source,
                "n_fields": n_fields,
                "row_fields": row_fields,
                "col_fields": col_fields,
                "data_fields": data_fields,
            })
        except Exception:
            continue
    return pivots


def extract_workbook_objects(doc):
    """Return ``{sheet_index_1based: {'charts': [...], 'pivots': [...]}}``."""
    out = {}
    sheets = doc.Sheets
    for si in range(sheets.Count):
        sheet = sheets.getByIndex(si)
        out[si + 1] = {
            "charts": _extract_sheet_charts(sheet),
            "pivots": _extract_sheet_pivots(sheet),
        }
    return out


# --------------------------------------------------------------------------- #
# Comparison helpers
# --------------------------------------------------------------------------- #
def _vec_eq(a, b, atol=TOLERANCE):
    """Element-wise equality: strings compared stripped, numbers within tol."""
    if len(a) != len(b):
        return False
    for x, y in zip(a, b):
        if isinstance(x, str) or isinstance(y, str):
            if str(x).strip() != str(y).strip():
                return False
        elif x is None or y is None:
            if x != y:
                return False
        else:
            try:
                if not np.isclose(x, y, atol=atol):
                    return False
            except Exception:
                if x != y:
                    return False
    return True


def _series_match(gt_series, res_series):
    """Exhaustive 1:1 matching of series by (values, marker), mirroring the
    original chart series matching."""
    remaining = list(res_series)
    for gv, gm in gt_series:
        for i, (rv, rm) in enumerate(remaining):
            if gm == rm and _vec_eq(gv, rv):
                remaining.pop(i)
                break
        else:
            return False
    return len(remaining) == 0


# --------------------------------------------------------------------------- #
# Chart comparison (port of compare_charts + worksheet charts branch)
# --------------------------------------------------------------------------- #
def _compare_one_chart(g, r, props):
    out = dict(props)
    if "chart_type" in out and g["type"] != r["type"]:
        out["chart_type"] = 0
    if "title" in out:
        if g["has_title"] != r["has_title"]:
            out["title"] = 0
        elif g["has_title"] and g["title"] != r["title"]:
            out["title"] = 0
    if "legend" in out:
        if g["has_legend"] != r["has_legend"] or (
            g["has_legend"] and g["legend_align"] != r["legend_align"]
        ):
            out["legend"] = 0
    if "axes" in out and g["axes"] != r["axes"]:
        out["axes"] = 0
    if "series" in out:
        ok = (
            len(g["series"]) == len(r["series"])
            and _vec_eq(g["categories"], r["categories"])
            and _series_match(g["series"], r["series"])
        )
        if not ok:
            out["series"] = 0
    return out


def compare_charts_on_sheet(gt_charts, res_charts, report):
    """Port of the ``charts`` branch of ``compare_worksheets``."""
    if len(gt_charts) != len(res_charts):
        report["charts"]["count"] = 0
        return
    if len(gt_charts) == 0:
        return
    props = report["charts"]
    pairs = [(g, r) for g in gt_charts for r in res_charts]
    n_true = 0
    for g, r in pairs:
        mr = _compare_one_chart(g, r, props)
        if all(mr.values()):
            n_true += 1
    if n_true != len(gt_charts):
        for k in report["charts"]:
            if k == "count":
                continue
            report["charts"][k] = 0


# --------------------------------------------------------------------------- #
# Pivot comparison (port of compare_pivot_tables + worksheet pivot branch)
# --------------------------------------------------------------------------- #
def _compare_one_pivot(g, r, props):
    out = dict(props)
    if "source" in out and g["source"] != r["source"]:
        out["source"] = 0
    if "filters" in out and g["n_fields"] != r["n_fields"]:
        out["filters"] = 0
    if "rows" in out:
        gr, gc, rr, rc = g["row_fields"], g["col_fields"], r["row_fields"], r["col_fields"]
        if not ((gr == rr and gc == rc) or (gr == rc and gc == rr)):
            out["rows"] = 0
            if "columns" in out:
                out["columns"] = 0
    if "values" in out:
        remaining = list(r["data_fields"])
        ok = True
        for df in g["data_fields"]:
            if df in remaining:
                remaining.remove(df)
            else:
                ok = False
                break
        if not ok:
            out["values"] = 0
    return out


def compare_pivots_on_sheet(gt_pivots, res_pivots, report):
    """Port of the ``pivot_tables`` branch of ``compare_worksheets``."""
    if len(gt_pivots) != len(res_pivots):
        report["pivot_tables"]["count"] = 0
        return
    if len(gt_pivots) == 0:
        return
    props = report["pivot_tables"]
    pairs = [(g, r) for g in gt_pivots for r in res_pivots]
    for g, r in pairs:
        mr = _compare_one_pivot(g, r, props)
        if all(mr.values()):
            return
    for k in report["pivot_tables"]:
        if k == "count":
            continue
        report["pivot_tables"][k] = 0
