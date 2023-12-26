from win32com.client import constants as c

class constants:
    ChartType = {
        '3DArea': -4098,
        '3DAreaStacked': 78,
        '3DAreaStacked100': 79,
        '3DBarClustered': 60,
        '3DBarStacked': 61,
        '3DBarStacked100': 62,
        '3DColumn': -4100,
        '3DColumnClustered': 54,
        '3DColumnStacked': 55,
        '3DColumnStacked100': 56,
        '3DLine': -4101,
        '3DPie': -4102,
        '3DPieExploded': 70,
        'Area': 1,
        'AreaStacked': 76,
        'AreaStacked100': 77,
        'BarClustered': 57,
        'BarOfPie': 71,
        'BarStacked': 58,
        'BarStacked100': 59,
        'Bubble': 15,
        'Bubble3DEffect': 87,
        'ColumnClustered': 51,
        'ColumnStacked': 52,
        'ColumnStacked100': 53,
        'ConeBarClustered': 102,
        'ConeBarStacked': 103,
        'ConeBarStacked100': 104,
        'ConeCol': 105,
        'ConeColClustered': 99,
        'ConeColStacked': 100,
        'ConeColStacked100': 101,
        'CylinderBarClustered': 95,
        'CylinderBarStacked': 96,
        'CylinderBarStacked100': 97,
        'CylinderCol': 98,
        'CylinderColClustered': 92,
        'CylinderColStacked': 93,
        'CylinderColStacked100': 94,
        'Doughnut': -4120,
        'DoughnutExploded': 80,
        'Line': 4,
        'LineMarkers': 65,
        'LineMarkersStacked': 66,
        'LineMarkersStacked100': 67,
        'LineStacked': 63,
        'LineStacked100': 64,
        'Pie': 5,
        'PieExploded': 69,
        'PieOfPie': 68,
        'PyramidBarClustered': 109,
        'PyramidBarStacked': 110,
        'PyramidBarStacked100': 111,
        'PyramidCol': 112,
        'PyramidColClustered': 106,
        'PyramidColStacked': 107,
        'PyramidColStacked100': 108,
        'Radar': -4151,
        'RadarFilled': 82,
        'RadarMarkers': 81,
        'StockHLC': 88,
        'StockOHLC': 89,
        'StockVHLC': 90,
        'StockVOHLC': 91,
        'Surface': 83,
        'SurfaceTopView': 85,
        'SurfaceTopViewWireframe': 86,
        'SurfaceWireframe': 84,
        'Scatter': -4169,
        'XYScatter': -4169,
        'XYScatterLines': 74,
        'XYScatterLinesNoMarkers': 75,
        'XYScatterSmooth': 72,
        'XYScatterSmoothNoMarkers': 73,
    }

    AxisType = {
        'x': 1,
        'y': 2,
        'z': 3,
        'X': 1,
        'Y': 2,
        'Z': 3,
    }

    AxisGroup = {
        'primary': 1,
        'secondary': 2,
    }

    TrendlineType = {
        'exponential': 5,
        'linear': -4132,
        'logarithmic': -4133,
        'movingAvg': 6,
        'polynomial': 3,
        'power': 4,
    }

    PageOrientation ={
        'landscape': 2,
        'portrait': 1,
    }

    PaperSize = {
        'A3': 8,
        'A4': 9,
        'A4Small': 10,
        'A5': 11,
        'B4': 12,
        'B5': 13,
        'Csheet': 24,
    }

    ColorIndex = {
        'automatic': -4105,
        'none': -4142,
    }

    BorderWeight = {
        'hairline': 1,
        'medium': -4138,
        'thick': 4,
        'thin': 2,
    }
    
    InsertShiftDirection = {
        'down': -4121,
        'right': -4161,
    }

    FormatConditionType = {
        'aboveAverage': 12,
        'blanks': 10,
        'cellValue': 1,
        'colorScale': 3,
        'dataBar': 4,
        'errors': 16,
        'expression': 2,
        'iconSet': 6,
        'noBlanks': 13,
        'noErrors': 17,
        'textString': 9,
        'timePeriod': 11,
        'top10': 5,
        'uniqueValues': 8,
    }

    FormatConditionOperator = {
        'between': 1,
        'equal': 3,
        'greaterThan': 5,
        'greaterThanOrEqual': 7,
        'lessThan': 6,
        'lessThanOrEqual': 8,
        'notBetween': 2,
        'notEqual': 4,
    }

    ColorIndex = {
        'black': 1,
        'white': 2,
        'red': 3,
        'green': 4,
        'blue': 5,
        'yellow': 6,
        'magenta': 7,
        'cyan': 8,
        'dark_red': 9,
        'dark_green': 10
    }

    ValidationType = {
        'date': 4,
        'decimal': 2,
        'list': 3,
        'textLength': 6,
        'time': 5,
        'wholeNumber': 1,
    }
    
    AutoFilterOperator = {
        'and': 1,
        'bottom10Items': 4,
        'bottom10Percent': 6,
        'cellColor': 8,
        'dynamic': 11,
        'fontColor': 9,
        'icon': 10,
        'or': 2,
        'top10Items': 3,
        'top10Percent': 5,
        'values': 7,
    }
    
    MarkerStyle = {
        'auto': -4105,
        'circle': 8,
        'dash': -4115,
        'diamond': 2,
        'dot': -4118,
        'none': -4142,
        'picture': -4147,
        'plus': 9,
        'square': 1,
        'star': 5,
        'triangle': 3,
        'x': -4168,
    }

    ConsolidationFunction = {
        'average': -4106,
        'count': -4112,
        'countNums': -4113,
        'max': -4136,
        'min': -4139,
        'product': -4149,
        'stDev': -4155,
        'stDevP': -4156,
        'sum': -4157,
        'unknown': 1000,
        'var': -4164,
        'varP': -4165,
    }

    SortOrder = {
        'ascending': 1,
        'descending': 2,
    }

    DataType = {
        'date': 'm/d/yyyy',
        'text': '@',
        'number': '0.00',
        'currency': '$#,##0.00',
        'time': 'h:mm:ss',
        'general': 'General',
        'percentage': '0.00%',
    }

    HorizontalAlignment = {
        'left': -4131,
        'center': -4108,
        'right': -4152,
        'justify': -4130
    }

    # Refer to https://learn.microsoft.com/en-us/office/vba/api/excel.xlconsolidationfunction
    SummarizationFunction = {
    'sum': c.xlSum,
    'count': c.xlCount,
    'countNumbers': c.xlCountNums,
    'average': c.xlAverage,
    'avg': c.xlAverage,
    'max': c.xlMax,
    'min': c.xlMin,
    'product': c.xlProduct,
    'var': c.xlVar,
    'varP': c.xlVarP,
    'standardDeviation': c.xlStDev,
    'standardDeviationP': c.xlStDevP
    }
    # 'sum', 'count', 'average', 'max', 'min', 'product', 'countNumbers', 'standardDeviation', 'standardDeviationP', 'var', or 'varP'
