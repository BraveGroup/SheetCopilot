import os
import win32com.client as win32
import json
import pythoncom
import time
from collections import defaultdict
from copy import deepcopy
import yaml

# Refer to https://analysistabs.com/excel-vba/colorindex/
COLOR_DICT = {
    'red': [3, 9, 30, 53],
    'green': [4, 10, 14, 31, 43, 50, 51],
    'yellow': [6, 19, 36, 27, 44],
    'orange': [22, 40, 45, 46],
    'purple': [7, 17, 24, 26, 29, 38, 39, 47, 54],
    'black': [1, 48, 56],
    'blue': [5, 8, 20, 23, 25, 28, 32, 33, 34, 37, 41, 42, 49, 55],
    'white': [2]
}

def compare_format_conditions(formatCondition1, formatCondition2, report):
    format_properties = deepcopy(report["format_conditions"])

    # formatCondition can be on of many types, such as 'Top10', and 'FormatCondition'.
    # If the condition types are not matched or they are matched but the formulas are not identical, the result will be judged as incorrect.
    if 'formula1' in format_properties.keys() and (type(formatCondition1) is not type(formatCondition2)) or formatCondition1.Formula1 != formatCondition2.Formula1:
        format_properties['formula1'] = 0

    if 'color' in format_properties.keys():
        for v in COLOR_DICT.values():
            if formatCondition1.Font.ColorIndex in v and formatCondition2.Font.ColorIndex in v:
                break
        else:
            format_properties['color'] = 0

    if 'fill_color' in format_properties.keys():
        for v in COLOR_DICT.values():
            if formatCondition1.Interior.ColorIndex in v and formatCondition2.Interior.ColorIndex in v:
                break
        else:
            format_properties['fill_color'] = 0
    
    if 'font' in format_properties.keys() and formatCondition1.Font.Name != formatCondition2.Font.Name:
        format_properties['font'] = 0

    if 'bold'  in format_properties.keys() and formatCondition1.Font.Bold != formatCondition2.Font.Bold:
        format_properties['bold'] = False

    if 'italic' in format_properties.keys() and formatCondition1.Font.Italic != formatCondition2.Font.Italic:
        format_properties['italic'] = False

    if 'underline' in format_properties.keys() and formatCondition1.Font.Underline != formatCondition2.Font.Underline:
        format_properties['underline'] = False

    return format_properties

def compare_filters_by_properties(filter1, filter2, report):
    filter_properties = deepcopy(report["filters"])

    if filter1.FilterMode != filter2.FilterMode:
        for k in filter_properties.keys():
            filter_properties[k] = 0
    else:
        # Collect filtered fields and criteria
        filter1_fields, filter1_criteria, filter2_fields, filter2_criteria = [], [], [], []
        
        # The number of Filters is equal to the number of columns
        for i, field in enumerate(filter1.Filters):
            try:
                # If the column has not been filtered, field.Criteria1 will throw an exception
                filter1_criteria.append(field.Criteria1)
                filter1_fields.append(i)
            except:
                pass

        for i, field in enumerate(filter2.Filters):
            try:
                # If the column has not been filtered, field.Criteria1 will throw an exception
                filter2_criteria.append(field.Criteria1)
                filter2_fields.append(i)
            except:
                pass
        
        if filter1_fields != filter2_fields:
            filter_properties["fields"] = 0
        
        if filter1_criteria != filter2_criteria:
            filter_properties["criteria"] = 0

    return filter_properties

def compare_filters_by_visible_range(ws1, ws2, report):
    filter_properties = deepcopy(report["filters"])
    mismatch = False

    ws2_filter = ws2.AutoFilter

    if ws2_filter is None:
        mismatch = True
    else:
        visible_range1 = ws1.AutoFilter.Range.SpecialCells(12)  # 12 represents xlCellTypeVisible
        visible_range2 = ws2_filter.Range.SpecialCells(12)

        if len(visible_range1) != len(visible_range2):
            mismatch = True
        else:
            for cell1, cell2 in zip(visible_range1, visible_range2):
                if cell1.Value != cell2.Value:
                    mismatch = True
                    break
    
    if mismatch:
        for k in filter_properties.keys():
            filter_properties[k] = 0
    
    return filter_properties

def compare_pivot_tables(pivot1, pivot2, report):
    # Initialize report
    chart_properties = deepcopy(report['pivot_tables'])

    # Compare name
    # if config['name'] and pivot1.Name != pivot2.Name:
    #     report["name"] = False
    #     print(f"Pivot table name mismatch in sheet {pivot1.Parent.Name}")

    # Compare source
    if 'source' in chart_properties.keys() and pivot1.SourceData != pivot2.SourceData:
        chart_properties["source"] = 0

    # Compare filters
    if 'filters' in chart_properties.keys() and pivot1.PivotFields().Count != pivot2.PivotFields().Count:
        chart_properties["filters"] = False

    # Compare rows
    if "rows" in chart_properties.keys():
        # Compare row fields
        pt1_row_fields = [field.SourceName for field in pivot1.RowFields]
        pt2_row_fields = [field.SourceName for field in pivot2.RowFields]
        pt1_col_fields = [field.SourceName for field in pivot1.ColumnFields]
        pt2_col_fields = [field.SourceName for field in pivot2.ColumnFields]
        
        if not (pivot1.RowFields.Count == pivot2.RowFields.Count and pt1_row_fields == pt2_row_fields and pt1_col_fields == pt2_col_fields) and\
            not (pivot1.RowFields.Count == pivot2.ColumnFields.Count and pt1_row_fields == pt2_col_fields and pt1_col_fields == pt2_row_fields):
            chart_properties["rows"], chart_properties["columns"] = 0, 0

    # Compare values
    if 'values' in chart_properties.keys() and [field.SourceName for field in pivot1.DataFields] != [field.SourceName for field in pivot2.DataFields]:
        chart_properties["values"] = 0

    return chart_properties

def compare_charts(chart1, chart2, report):
    # Initialize report
    chart_properties = deepcopy(report['chart'])

    # Compare name
    # if config['name'] and chart1.Name != chart2.Name:
    #     report["name"] = False
    #     print(f"Chart name mismatch in sheet {chart1.Parent.Name}")

    # Compare chart type
    if 'chart_type' in chart_properties.keys() and chart1.Chart.ChartType != chart2.Chart.ChartType:
        chart_properties["chart_type"] = 0

    # Compare title
    if 'title' in chart_properties.keys():
        if chart1.Chart.HasTitle != chart2.Chart.HasTitle:
            chart_properties["title"] = 0
        elif chart1.Chart.HasTitle:
            if chart1.Chart.ChartTitle.Text != chart2.Chart.ChartTitle.Text:
                chart_properties["title"] = 0

    # Compare legend
    if 'legend' in chart_properties.keys() and chart1.Chart.HasLegend != chart2.Chart.HasLegend:
        chart_properties["legend"] = 0

    # Compare axes
    if 'axes' in chart_properties.keys():
        if chart1.Chart.Axes().Count != chart2.Chart.Axes().Count:
            chart_properties["axes"] = 0
        else:
            for axis1, axis2 in zip(chart1.Chart.Axes(), chart2.Chart.Axes()):
                if axis1.HasTitle != axis2.HasTitle:
                    chart_properties["axes"] = 0

    # Compare series
    if 'series' in chart_properties.keys() or 'trendlines' in chart_properties.keys():
        mismatch = False
        if chart1.Chart.SeriesCollection().Count != chart2.Chart.SeriesCollection().Count:
            mismatch = True
        else:
            chart1_series, chart2_series = list(chart1.Chart.SeriesCollection()), list(chart2.Chart.SeriesCollection())

            mismatch = False
            for series1 in chart1_series:
                i = 0
                while i < len(chart2_series):
                    if series1.Name == chart2_series[i].Name:
                        break
                else: # If no matched series in the result chart
                    mismatch = True
                    break
                    
                chart2_series.pop(i)
            
        if mismatch:
            chart_properties["series"] = 0

    return chart_properties

def compare_cells_itercell(ws1, ws2, report):
    """
    Compare cell-by-cell
    """
    NumberOfRows = min(ws1.Range('A1').End(-4121).Row,ws1.UsedRange.Rows.Count)
    NumberOfColumns = min(ws1.Range('A1').End(-4161).Column, ws1.UsedRange.Columns.Count)

    ws1_values = ws1.Range(ws1.Cells(1, 1), ws1.Cells(NumberOfRows, NumberOfColumns)).Value
    ws2_values = ws2.Range(ws2.Cells(1, 1), ws2.Cells(NumberOfRows, NumberOfColumns)).Value

    assert ws1_values is not None and ws2_values is not None

    # Start checking ...
    mismatch = False

    for row in range(0, NumberOfRows):
        for col in range(0, NumberOfColumns):
            cell1 = ws1_values[row][col]
            cell2 = ws2_values[row][col]
            # Compare cell values
            if 'values' in report['cells'].keys() and cell1 != cell2:
                mismatch = True
                report['cells']['values'] = 0
                break
            
            # Compare cell formatting
            if 'formatting' in report['cells'].keys() and \
                (cell1.Font.Name != cell2.Font.Name \
                    or cell1.Font.Size != cell2.Font.Size \
                    or cell1.Font.Color != cell2.Font.Color \
                    or cell1.Font.Bold != cell2.Font.Bold \
                    or cell1.Font.Italic != cell2.Font.Italic \
                    or cell1.Font.Underline != cell2.Font.Underline \
                    or cell1.Interior.Color != cell2.Interior.Color):
                mismatch = True
                report['cells']['formatting'] = 0
                break
            
            # In our tasks, hyperlinks are only assigned to headers, so we only check the first row for the sake of efficiency
            if row <= 2 and 'hyperlink' in report['cells'].keys() and cell1.Hyperlink != cell2.Hyperlink:
                mismatch = True
                report['cells']['hyperlink'] = 0
                break

        if mismatch: break

def compare_cells_itercolumn(ws1, ws2, report):
    """
    Compare column-by-column
    arg1: ws1. GT sheet
    arg2: ws2. Result sheet
    report: a dict recording the comparison results

    retval: True if ws1 matches ws2; else False
    """

    NumberOfRows = min(ws1.Range('A1').End(-4121).Row,ws1.UsedRange.Rows.Count)
    NumberOfColumns = min(ws1.Range('A1').End(-4161).Column, ws1.UsedRange.Columns.Count)
    # NumberOfRows = ws1.UsedRange.Rows.Count
    # NumberOfColumns = ws1.UsedRange.Columns.Count
    mismatch = False

    if NumberOfRows != min(ws2.Range('A1').End(-4121).Row, ws2.UsedRange.Rows.Count) \
        or NumberOfColumns != min(ws2.Range('A1').End(-4161).Column, ws2.UsedRange.Columns.Count):
    # if NumberOfRows != ws2.UsedRange.Rows.Count\
    #     or NumberOfColumns != ws2.UsedRange.Columns.Count:
        mismatch = True
        report['cells']['values'] = 0
        
    else:
        ws1_cells = ws1.Range(ws1.Cells(1, 1), ws1.Cells(NumberOfRows, NumberOfColumns))
        ws2_cells = ws2.Range(ws2.Cells(1, 1), ws2.Cells(NumberOfRows, NumberOfColumns))

        assert ws1_cells is not None and ws2_cells is not None

        # Start checking ...
        for col in range(NumberOfColumns):
            header_ws1 = ws1_cells.Value[0][col]

            # iterate through all headers to find the matched column in the result sheet
            match_col_id = 0
            while match_col_id < NumberOfColumns:
                if header_ws1 == ws2_cells.Value[0][match_col_id]:
                    break
                match_col_id += 1
            else:
                mismatch = True
                report['cells']['values'] = 0
                break
            
            # Firtly, compare all cells under this header
            # If the row number is too large, we check the column every ten rows
            step = 100 if NumberOfRows > 200 else 1

            for row in range(1, NumberOfRows, step):
                # NOTE：the index of elements in worksheet.Cells starts from 1, not 0
                cell1 = ws1_cells.Cells(row + 1, col + 1) # gt.
                cell2 = ws2_cells.Cells(row + 1, col + 1) # result

                # Compare cell values
                if 'values' in report['cells'].keys() and cell1.Value != cell2.Value:
                    mismatch = True
                    report['cells']['values'] = 0
                    break
            
                # Compare cell formatting
                if 'formatting' in report['cells'].keys() and \
                    (cell1.Font.Name != cell2.Font.Name \
                        or cell1.Font.Size != cell2.Font.Size \
                        or cell1.Font.Color != cell2.Font.Color \
                        or cell1.Font.Bold != cell2.Font.Bold \
                        or cell1.Font.Italic != cell2.Font.Italic \
                        or cell1.Font.Underline != cell2.Font.Underline \
                        or cell1.Interior.Color != cell2.Interior.Color):
                    mismatch = True
                    report['cells']['formatting'] = 0
                    break

                if 'hyperlink' in report['cells'].keys():
                    # If two cells both have no hyperlinks, they are considered matched
                    if list(cell1.Hyperlinks) != list(cell2.Hyperlinks):
                        mismatch = True
                        report['cells']['hyperlink'] = 0
                        break
            
            if mismatch: break

    return not mismatch

def get_matching_pairs(ws1_objects, ws2_objects):
    pairs = []
    for chart in ws1_objects:
        for target_chart in ws2_objects:
            pairs.append([chart, target_chart])

    return pairs

def compare_worksheets(ws1, ws2, config):
    # Initialize report

    # 0 means not matched; 1 means matched; -1 means not checked yet.
    # If one property of a category (e.g. "values" of "cells") is not matched, other properties of this category will not be checked to save time.
    report = defaultdict(dict)
    for category, properties in config.items():
        for property, need_check in properties.items():
            if need_check:
                report[category][property] = 1

    # Compare cells
    if any(config['cells'].values()):
        compare_cells_itercolumn(ws1, ws2, report)

    # Compare charts
    if any(config['charts'].values()):
        if ws1.ChartObjects().Count != ws2.ChartObjects().Count:
            report["charts"]["count"] = 0
        else:
            pairs_to_compare = get_matching_pairs(ws1.ChartObjects(), ws2.ChartObjects())

            # Exhaustive matching
            for chart1, chart2 in pairs_to_compare:
                match_results = compare_charts(chart1, chart2, report)

                if all(match_results.values()):
                    break
            else:
                for k in report["charts"].keys():
                    if k == "count": continue
                    report["charts"][k] = 0
    
    # Compare pivot tables
    if any(config['pivot_tables'].values()):
        if ws1.PivotTables().Count != ws2.PivotTables().Count:
            report["pivot_tables"]["count"] = False
        else:
            pairs_to_compare = get_matching_pairs(ws1.PivotTables(), ws2.PivotTables())

            # Exhaustive matching
            for pivot1, pivot2 in pairs_to_compare:
                match_results = compare_pivot_tables(pivot1, pivot2, report)

                if all(match_results.values()):
                    break
            else:
                for k in report["pivot_tables"].keys():
                    if k == "count": continue
                    report["pivot_tables"][k] = 0

    # Compare filters
    if any(config['filters'].values()):
        report["filters"] = compare_filters_by_visible_range(ws1, ws2, report)

    # Compare format conditions
    if any(config['format_conditions'].values()):
        if ws1.UsedRange.FormatConditions.Count != ws2.UsedRange.FormatConditions.Count:
            report["format_conditions"]["count"] = 0
        elif ws1.UsedRange.FormatConditions.Count == 0 and ws2.UsedRange.FormatConditions.Count == 0:
            for k in report["format_conditions"]:
                report["format_conditions"][k] = 1
        else:
            pairs_to_compare = get_matching_pairs(ws1.UsedRange.FormatConditions, ws2.UsedRange.FormatConditions)


            for condition1, condition2 in pairs_to_compare:
                match_results = compare_format_conditions(condition1, condition2, report)

                if all(match_results.values()):
                    break
            else:
                # If multiple conditional formatting objetcs exist, it is hard to judge the matching between the two sheets.
                # Therefore, we set the check state all to 0
                if len(pairs_to_compare) > 1:
                    for k in report["format_conditions"].keys():
                        if k == "count": continue
                        report["format_conditions"][k] = 0
                # If only one conditional formatting objetc exist
                else:
                    for k in report["format_conditions"].keys():
                        if k == "count": continue
                        report["format_conditions"][k] = match_results[k]

    return report

def compare_frozen_panes(excel, wb1, wb2):
    # Iterate over all sheets
    wb1_frozen_ranges,  wb2_frozen_ranges = [], []
    for wb1_sheet, wb2_sheet in zip(wb1.Sheets, wb2.Sheets):
        # Get the number of rows and columns of the frozen pane
        try:
            # Activate the sheet
            wb1_sheet.Activate(); 

            # Get the active window
            active_window = excel.ActiveWindow
            wb1_sheet_frozen_panes = {
                'row': active_window.SplitRow,
                'column': active_window.SplitColumn
            }

            wb2_sheet.Activate()
            active_window = excel.ActiveWindow
            wb2_sheet_frozen_panes = {
                'row': active_window.SplitRow,
                'column': active_window.SplitColumn
            }
        except:
            wb1_sheet_frozen_panes, wb2_sheet_frozen_panes = None, None

        wb1_frozen_ranges.append(wb1_sheet_frozen_panes)
        wb2_frozen_ranges.append(wb2_sheet_frozen_panes)

    # Compare
    match = [1] * len(wb1_frozen_ranges)
    for i, frozen_pane1, frozen_pane2 in zip(range(len(wb1_frozen_ranges)), wb1_frozen_ranges, wb2_frozen_ranges):
        # The number of rows of the frozen panes must be equal
        # The number of columns of the frozen panes must be equal or the that of the result sheet is 0 (0 means only the rows are frozen)
        if not (wb1_sheet_frozen_panes is not None and frozen_pane1['row'] == frozen_pane2['row'] and (frozen_pane1['column'] == frozen_pane2['column'] or frozen_pane2['column'] == 0)) :
            match[i] = 0
    
    return match

def check_success(report):
    if isinstance(report, dict):
        for key, value in report.items():
            if isinstance(value, dict) or isinstance(value, list):
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

def compare_workbooks(file1, file2, check_boards):
    # Start Excel application
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = True
    excel.DisplayAlerts = False

    # Open workbooks
    time.sleep(0.2)
    wb2 = excel.Workbooks.Open(os.path.abspath(file2))
    time.sleep(0.3)
    wb1 = excel.Workbooks.Open(os.path.abspath(file1))

    # Initialize report
    report = {
    }

    # Compare worksheet count
    if wb1.Worksheets.Count != wb2.Worksheets.Count: # TOdo
        report["worksheet_count"] = False
        wb1.Close(SaveChanges=False)
        wb2.Close(SaveChanges=False)
        return report, False
    
    # Loop through all worksheets in both workbooks
    for sheet_id, ws1, ws2 in zip(range(1, wb1.Worksheets.Count + 1), wb1.Worksheets, wb2.Worksheets):
        check_board = check_boards.get(str(sheet_id), None)

        if check_board is None: continue
        report[sheet_id] = compare_worksheets(ws1, ws2, check_board)

    # Frozen panes need to be checked using wb
    for i, sheet_report in enumerate(report.values()):
        if 'view' in sheet_report.keys():
            break
    else:
        match_lst = compare_frozen_panes(excel, wb1, wb2)
        for i, sheet_report in enumerate(report.values()):
            if 'view' in sheet_report.keys():
                sheet_report['view']["freeze_pane"] = match_lst[i]

    sussess = check_success(report)

    # Close workbooks and quit Excel
    wb1.Close(SaveChanges=False)
    wb2.Close(SaveChanges=False)
    # excel.Application.Quit()

    # pythoncom.CoUninitialize()

    return report, sussess

if __name__ == "__main__":
    # Provide the paths to your workbooks
    rpa_processed_file = r"D:\SheetCopilot_data\action_diff_granularity\Chart1\4_StockChange\4_StockChange_1.xlsx"
    ground_truth_file = r"D:\Github\ActionTransformer\Excel_data\example_sheets_part1\task_sheet_answers\StockChange\4_StockChange\4_StockChange_gt1.xlsx"
    check_board_file = ground_truth_file.replace(".xlsx", "_check.yaml")

    with open(check_board_file, 'r') as f:
        check_boards = yaml.load(f, yaml.Loader)
    report, sussess = compare_workbooks(ground_truth_file, rpa_processed_file, check_boards["check_board"])

    print(report)

    print(sussess)