import os
import win32com.client as win32
import json
import pythoncom
import time

check_board = {
    "cells": {
        "values": False,
        "formatting": False
    },
    "charts": {
        "name": False,
        "chart_type": True,
        "title": False,
        "legend": False,
        "axes": False,
        "series": False,
        "trendlines": False
    },
    "pivot_tables": {
        "name": False,
        "source": False,
        "filters": False,
        "rows": False,
        "columns": False,
        "values": False
    },
    "filters": {
        "name": False,
        "source": False,
        "filters": False
    },
    'format_conditions': {
        'type': True,
        'operator': True,
        'formula1': True,
        'color': True,
        'fill_color': True,
        'font': True,
        'bold': True,
        'italic': True,
        'underline': True,
    }
}

def compare_format_conditions(formatCondition1, formatCondition2, config):
    report = {
        # 'type': True,
        # 'operator': True,
        'formula1': True,
        'color': True,
        'fill_color': True,
        'font': True,
        'bold': True,
        'italic': True,
        'underline': True,
    }
    # if config['type'] and formatCondition1.Type != formatCondition2.Type:
    #     report['type'] = False
    #     print(f"Format condition type mismatch in sheet {formatCondition1.Parent.Parent.Name}")
    # if config['operator'] and formatCondition1.Operator != formatCondition2.Operator:
    #     report['operator'] = False
    #     print(f"Format condition operator mismatch in sheet {formatCondition1.Parent.Parent.Name}")
    if config['formula1'] and formatCondition1.Formula1 != formatCondition2.Formula1:
        report['formula1'] = False
        print(f"Format condition formula1 mismatch in sheet {formatCondition1.Parent.Parent.Name}")
    if config['color'] and formatCondition1.Font.ColorIndex != formatCondition2.Font.ColorIndex:
        report['color'] = False
        print(f"Format condition color mismatch in sheet {formatCondition1.Parent.Parent.Name}")
    if config['fill_color'] and formatCondition1.Interior.ColorIndex != formatCondition2.Interior.ColorIndex:
        report['fill_color'] = False
        print(f"Format condition fill color mismatch in sheet {formatCondition1.Parent.Parent.Name}")
    if config['font'] and formatCondition1.Font.Name != formatCondition2.Font.Name:
        report['font'] = False
        print(f"Format condition font mismatch in sheet {formatCondition1.Parent.Parent.Name}")
    if config['bold'] and formatCondition1.Font.Bold != formatCondition2.Font.Bold:
        report['bold'] = False
        print(f"Format condition bold mismatch in sheet {formatCondition1.Parent.Parent.Name}")
    if config['italic'] and formatCondition1.Font.Italic != formatCondition2.Font.Italic:
        report['italic'] = False
        print(f"Format condition italic mismatch in sheet {formatCondition1.Parent.Parent.Name}")
    if config['underline'] and formatCondition1.Font.Underline != formatCondition2.Font.Underline:
        report['underline'] = False
        print(f"Format condition underline mismatch in sheet {formatCondition1.Parent.Parent.Name}")

    return report

def compare_filters(filter1, filter2, config):
    if filter1.FilterMode != filter2.FilterMode:
        print(f"Filter mode mismatch in sheet {filter1.Parent.Name}")
        report = {
            "FilterMode": False
        }
    else:
        report = {
            "FilterMode": True
        }
    return report

def compare_pivot_tables(pivot1, pivot2, config):
    # Initialize report
    report = {
        "name": True,
        "source": False,
        "filters": False,
        "rows": True,
        "columns": False,
        "values": True
    }

    # Compare name
    if config['name'] and pivot1.Name != pivot2.Name:
        report["name"] = False
        print(f"Pivot table name mismatch in sheet {pivot1.Parent.Name}")

    # Compare source
    if config['source'] and pivot1.SourceData != pivot2.SourceData:
        report["source"] = False
        print(f"Pivot table source mismatch in sheet {pivot1.Parent.Name}")

    # Compare filters
    if config['filters'] and pivot1.PivotFields().Count != pivot2.PivotFields().Count:
        report["filters"] = False
        print(f"Pivot table filter count mismatch in sheet {pivot1.Parent.Name}")

    # Compare rows
    if config['rows'] and pivot1.RowFields.Count != pivot2.RowFields.Count:
        report["rows"] = False
        print(f"Pivot table row count mismatch in sheet {pivot1.Parent.Name}")

    # Compare columns
    if config['columns'] and pivot1.ColumnFields.Count != pivot2.ColumnFields.Count:
        report["columns"] = False
        print(f"Pivot table column count mismatch in sheet {pivot1.Parent.Name}")

    # Compare values
    if config['values'] and pivot1.DataFields.Count != pivot2.DataFields.Count:
        report["values"] = False
        print(f"Pivot table value count mismatch in sheet {pivot1.Parent.Name}")

    return report

def compare_charts(chart1, chart2, config):
    # Initialize report
    report = {
        "name": True,
        "chart_type": True,
        "title": True,
        "legend": True,
        "axes": True,
        "series": True,
        "trendlines": True
    }

    # Compare name
    if config['name'] and chart1.Name != chart2.Name:
        report["name"] = False
        print(f"Chart name mismatch in sheet {chart1.Parent.Name}")

    # Compare chart type
    if config['chart_type'] and chart1.Chart.ChartType != chart2.Chart.ChartType:
        report["chart_type"] = False
        print(f"Chart type mismatch in sheet {chart1.Parent.Name}")

    # Compare title
    if config['title']:
        if chart1.Chart.HasTitle != chart2.Chart.HasTitle:
            report["title"] = False
            print(f"Chart title mismatch in sheet {chart1.Parent.Name}")
        elif chart1.Chart.HasTitle:
            if chart1.Chart.ChartTitle.Text != chart2.Chart.ChartTitle.Text:
                report["title"] = False
                print(f"Chart title text mismatch in sheet {chart1.Parent.Name}")

    # Compare legend
    if config['legend'] and chart1.Chart.HasLegend != chart2.Chart.HasLegend:
        report["legend"] = False
        print(f"Chart legend mismatch in sheet {chart1.Parent.Name}")

    # Compare axes
    if config['axes']:
        if chart1.Chart.Axes().Count != chart2.Chart.Axes().Count:
            report["axes"] = False
            print(f"Chart axes count mismatch in sheet {chart1.Parent.Name}")
        else:
            for axis1, axis2 in zip(chart1.Chart.Axes(), chart2.Chart.Axes()):
                if axis1.HasTitle != axis2.HasTitle:
                    report["axes"] = False
                    print(f"Chart axes title mismatch in sheet {chart1.Parent.Name}")

    # Compare series
    if config['series']:
        if chart1.Chart.SeriesCollection().Count != chart2.Chart.SeriesCollection().Count:
            report["series"] = False
            print(f"Chart series count mismatch in sheet {chart1.Parent.Name}")
        else:
            for series1, series2 in zip(chart1.Chart.SeriesCollection(), chart2.Chart.SeriesCollection()):
                if series1.Name != series2.Name:
                    report["series"] = False
                    print(f"Chart series name mismatch in sheet {chart1.Parent.Name}")

    return report

def compare_worksheets(ws1, ws2, config):
    # Initialize report
    report = {
        "cells": [],
        "charts": [],
        "pivot_tables": [],
        "filters": [],
        "format_conditions": [],
    }

    # Compare cell
    if any(config['cells'].values()):
        mismatch_cell_count = 0
        NumberOfRows = min(ws1.Range('A1').End(-4121).Row,ws1.UsedRange.Rows.Count)
        NumberOfColumns = min(ws1.Range('A1').End(-4161).Column, ws1.UsedRange.Columns.Count)
        start = time.time()
        ws1_values = ws1.Range(ws1.Cells(1, 1), ws1.Cells(NumberOfRows, NumberOfColumns)).Value
        ws2_values = ws2.Range(ws2.Cells(1, 1), ws2.Cells(NumberOfRows, NumberOfColumns)).Value
        if ws1_values is None and ws2_values is None:
            report["cells"].append({
                "values": True,
            })
        elif ws1_values is None or ws2_values is None:
            report["cells"].append({
            "values": mismatch_cell_count == 0,
        })
        else:
            for row in range(0, NumberOfRows):
                for col in range(0, NumberOfColumns):
                    # Compare cell values
                    cell1 = ws1_values[row][col]
                    cell2 = ws2_values[row][col]
                    mismatch = False
                    if config['cells']['values'] and cell1 != cell2:
                        mismatch = True
                        # print(f"Cell value mismatch in sheet {ws1.Name} at {cell1.Address}")
                    
                    # Compare cell formatting
                    # if config['cells']['formatting']:
                    #     if cell1.Font.Name != cell2.Font.Name:
                    #         mismatch = True
                    #         # print(f"Cell font name mismatch in sheet {ws1.Name} at {cell1.Address}")
                    #     if cell1.Font.Size != cell2.Font.Size:
                    #         mismatch = True
                    #         # print(f"Cell font size mismatch in sheet {ws1.Name} at {cell1.Address}")
                    #     if cell1.Font.Color != cell2.Font.Color:
                    #         mismatch = True
                    #         # print(f"Cell font color mismatch in sheet {ws1.Name} at {cell1.Address}")
                    #     if cell1.Font.Bold != cell2.Font.Bold:
                    #         mismatch = True
                    #         # print(f"Cell font bold mismatch in sheet {ws1.Name} at {cell1.Address}")
                    #     if cell1.Font.Italic != cell2.Font.Italic:
                    #         mismatch = True
                    #         # print(f"Cell font italic mismatch in sheet {ws1.Name} at {cell1.Address}")
                    #     if cell1.Font.Underline != cell2.Font.Underline:
                    #         mismatch = True
                    #         # print(f"Cell font underline mismatch in sheet {ws1.Name} at {cell1.Address}")
                    #     if cell1.Interior.Color != cell2.Interior.Color:
                    #         mismatch = True
                    #         # print(f"Cell interior color mismatch in sheet {ws1.Name} at {cell1.Address}")

                    mismatch_cell_count += mismatch

            # value1 = ws1.Range(ws1.Cells(1, 1), ws1.Cells(NumberOfRows, NumberOfColumns)).Value
            # value2 = ws2.Range(ws2.Cells(1, 1), ws2.Cells(NumberOfRows, NumberOfColumns)).Value
            # for row in range(NumberOfRows):
            #     for col in range(NumberOfColumns):
            #         if value1[row][col] != value2[row][col]:
            #             mismatch_cell_count += 1
                        # print(f"Cell value mismatch in sheet {ws1.Name} at {cell1.Address}")

            report["cells"].append({
                # "values": 1 - mismatch_cell_count/(NumberOfRows*NumberOfColumns),
                "values": mismatch_cell_count == 0,
            })

    # Compare charts
    if any(config['charts'].values()):
        if ws1.ChartObjects().Count != ws2.ChartObjects().Count:
            report["charts_count"] = False
            print(f"Chart count mismatch in sheet {ws1.Name}")
        else:
            for chart1, chart2 in zip(ws1.ChartObjects(), ws2.ChartObjects()):
                report["charts"].append(compare_charts(chart1, chart2, config['charts']))
    
    # Compare pivot tables
    if any(config['pivot_tables'].values()):
        if ws1.PivotTables().Count != ws2.PivotTables().Count:
            report["pivot_tables_count"] = False
            print(f"Pivot table count mismatch in sheet {ws1.Name}")
        else:
            for pivot1, pivot2 in zip(ws1.PivotTables(), ws2.PivotTables()):
                report["pivot_tables"].append(compare_pivot_tables(pivot1, pivot2, config['pivot_tables']))

    # Compare filters
    # if any(config['filters'].values()):
    #     report["filters"].append(compare_filters(ws1.AutoFilter, ws2.AutoFilter, config['filters']))

    # Compare format conditions
    if any(config['format_conditions'].values()):
        if ws1.UsedRange.FormatConditions.Count != ws2.UsedRange.FormatConditions.Count:
            report["format_conditions_count"] = False
            print(f"Format condition count mismatch in sheet {ws1.Name}")
        else:
            for condition1, condition2 in zip(ws1.UsedRange.FormatConditions, ws2.UsedRange.FormatConditions):
                report["format_conditions"].append(compare_format_conditions(condition1, condition2, config['format_conditions']))

    return report

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

def compare_workbooks(file1, file2, config = check_board):

    # pythoncom.CoInitialize()

    # Start Excel application
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = True
    excel.DisplayAlerts = False

    # Open workbooks
    time.sleep(0.2)
    wb1 = excel.Workbooks.Open(os.path.abspath(file1))
    time.sleep(0.3)
    wb2 = excel.Workbooks.Open(os.path.abspath(file2))

    # Initialize report
    report = {
        "worksheets": []
    }

    # Compare worksheet count
    if wb1.Worksheets.Count != wb2.Worksheets.Count: # TOdo
        report["worksheet_count"] = False
        print(f"Worksheet count mismatch in workbook {wb1.Name} and {wb2.Name}")
        wb1.Close(SaveChanges=False)
        wb2.Close(SaveChanges=False)
        return report, False
    
    # Loop through all worksheets in both workbooks
    for ws1, ws2 in zip(wb1.Worksheets, wb2.Worksheets):
        report["worksheets"].append(compare_worksheets(ws1, ws2, config))

    sussess = check_success(report["worksheets"])

    # Close workbooks and quit Excel
    wb1.Close(SaveChanges=False)
    wb2.Close(SaveChanges=False)
    # excel.Application.Quit()

    # pythoncom.CoUninitialize()

    return report, sussess

if __name__ == "__main__":
    # Provide the paths to your workbooks
    ground_truth_file = "../example_sheets_part1/Physics/XYScatterPlot_Ans.xlsx"
    rpa_processed_file = "../example_sheets_part1/Physics/XYScatterPlot_1.xlsx"
    dir = "../ActionTransformer/Excel_data/example_sheets_part1/Business_Res/Invoices"
    files = os.listdir(dir)
    count = len(files)
    consisstent = [0 for i in range(count)]
    for i in range(count):
        for j in range(i+1,count):
            files1 = dir + "/" + files[i]
            files2 = dir + "/" + files[j]
            report, sussess = compare_workbooks(files1, files2, check_board)
            if sussess:
                consisstent[i] += 1
                consisstent[j] += 1
            # report = compare_workbooks(ground_truth_file, rpa_processed_file, check_board)
            print(json.dumps(report, indent=2))
    print(consisstent)