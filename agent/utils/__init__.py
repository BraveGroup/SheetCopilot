from .ChatGPT import ChatGPT
from .StateMachine import StateMachine
from .compare_sheets import compare_workbooks

action2API = {
    "Update cell value": "Write",
    "Delete": "Delete",
    "Merge cells": "Merge",
    "Split text to columns": "SplitText",
    "Insert row": "InsertRow",
    "Insert column": "InsertColumn",
    "Autofill": "AutoFill",
    "Copy-paste": "CopyPaste",
    "Copy-paste format": "CopyPasteFormat",
    "Find and replace": "FindReplace",
    "Set hyperlink": "SetHyperlink",
    "Remove duplicates": "RemoveDuplicate",
    "Create sheet": "CreateSheet",
    "Clear": "Clear",
    "Sort": "Sort",
    "Filter": "Filter",
    "Hide rows": "HideRow",
    "Hide columns": "HideColumn",
    "Create named range": "CreateNamedRange",
    "Split panes": "SplitPanes",
    "Freeze panes": "FreezePanes",
    "Format cells": "SetFormat",
    "Set data type": "SetDataType",
    "Set border": "SetBorderAround",
    "Resize cells": "ResizeRowColumn",
    "Conditional formatting": "SetConditionalFormat",
    "Lock and unlock": "LockOrUnlockCells",
    "Protect": "ProtectSheet",
    "Display formulas": "DisplayFormulas",
    "Create chart": "CreateChart",
    "Create Pivot Chart": "CreateChartFromPivotTable",
    "Set chart title": "SetChartTitle",
    "Set chart axis": "SetChartAxis",
    "Set chart has axis": "SetChartHasAxis",
    "Set chart legend": "SetChartLegend",
    "Set chart type": "SetChartType",
    "Set chart color": "SetChartColor",
    "Set chart marker": "SetChartMarker",
    "Set trend line": "SetChartTrendline",
    "Add data labels": "AddDataLabels",
    "Remove data labels": "RemoveDataLabels",
    "Add data series": "AddDataSeries",
    "Remove data series": "RemoveDataSeries",
    "Set data series source": "SetDataSeriesSource",
    "Add error bars": "AddErrorBars",
    "Create Pivot Table": "CreatePivotTable",
    "Set summary type": "SetPivotTableSummaryFunction",
    "Sort Pivot Table": "SortPivotTable"
}

from colorama import Fore, Style

def print_dialog(context):
    for message in context:
        if message['role'] in ['system', 'user', 'human']:
            print(Fore.YELLOW, f"{message['role']}] >>>\n{message['content']}\n<<<"); print(Fore.RESET)
        else:
            print(Fore.CYAN, f"{message['role']}] >>>\n{message['content']}\n<<<"); print(Fore.RESET)