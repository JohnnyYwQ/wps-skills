/**
 * Input: Excel 操作参数
 * Output: Excel 操作结果
 * Pos: macOS 加载项 Excel 处理器。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * WPS Excel Handler - Mac版 JavaScript API实现
 * 对标Windows PowerShell COM接口，提供一致的API
 * @author 老王
 */

// 获取工作表，支持索引或名称
function getSheet(wb, sheet) {
    if (!sheet) return wb.ActiveSheet;
    return wb.Sheets.Item(sheet);
}

// 获取当前工作簿信息
function getActiveWorkbook(params) {
    try {
        var app = Application;
        var wb = app.ActiveWorkbook;
        if (!wb) return { success: false, error: '没有打开的工作簿' };

        var sheets = [];
        for (var i = 1; i <= wb.Sheets.Count; i++) {
            sheets.push(wb.Sheets.Item(i).Name);
        }
        return {
            success: true,
            data: { name: wb.Name, path: wb.FullName, sheetCount: wb.Sheets.Count, sheets: sheets }
        };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 读取单元格值
function getCellValue(params) {
    try {
        var wb = Application.ActiveWorkbook;
        if (!wb) return { success: false, error: '没有打开的工作簿' };

        var sheet = getSheet(wb, params.sheet);
        var cell = sheet.Cells.Item(params.row, params.col);
        return {
            success: true,
            data: { value: cell.Value2, text: cell.Text, formula: cell.Formula }
        };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 写入单元格值
function setCellValue(params) {
    try {
        var wb = Application.ActiveWorkbook;
        if (!wb) return { success: false, error: '没有打开的工作簿' };

        var sheet = getSheet(wb, params.sheet);
        sheet.Cells.Item(params.row, params.col).Value2 = params.value;
        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 读取范围数据
function getRangeData(params) {
    try {
        var wb = Application.ActiveWorkbook;
        if (!wb) return { success: false, error: '没有打开的工作簿' };

        var sheet = getSheet(wb, params.sheet);
        var range = sheet.Range(params.range);
        var data = [];
        for (var r = 1; r <= range.Rows.Count; r++) {
            var row = [];
            for (var c = 1; c <= range.Columns.Count; c++) {
                row.push(range.Cells.Item(r, c).Value2);
            }
            data.push(row);
        }
        return { success: true, data: { data: data } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 写入范围数据
function setRangeData(params) {
    try {
        var wb = Application.ActiveWorkbook;
        if (!wb) return { success: false, error: '没有打开的工作簿' };

        var sheet = getSheet(wb, params.sheet);
        var range = sheet.Range(params.range);
        for (var r = 0; r < params.data.length; r++) {
            for (var c = 0; c < params.data[r].length; c++) {
                range.Cells.Item(r + 1, c + 1).Value2 = params.data[r][c];
            }
        }
        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置公式
function setFormula(params) {
    try {
        var wb = Application.ActiveWorkbook;
        if (!wb) return { success: false, error: '没有打开的工作簿' };

        var sheet = getSheet(wb, params.sheet);
        sheet.Cells.Item(params.row, params.col).Formula = params.formula;
        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 获取Excel上下文（表头、选中区域等）
function getExcelContext(params) {
    try {
        var app = Application;
        var wb = app.ActiveWorkbook;
        if (!wb) return { success: false, error: '没有打开的工作簿' };

        var sheet = app.ActiveSheet;
        var usedRange = sheet.UsedRange;
        var headers = [];

        if (usedRange.Rows.Count > 0) {
            var headerRow = usedRange.Rows.Item(1);
            var colCount = Math.min(headerRow.Columns.Count, 26);
            for (var i = 1; i <= colCount; i++) {
                headers.push({
                    column: String.fromCharCode(64 + i),
                    value: headerRow.Cells.Item(1, i).Value2
                });
            }
        }

        var sheets = [];
        for (var i = 1; i <= wb.Sheets.Count; i++) {
            sheets.push(wb.Sheets.Item(i).Name);
        }

        return {
            success: true,
            data: {
                workbookName: wb.Name,
                currentSheet: sheet.Name,
                allSheets: sheets,
                selectedCell: app.Selection.Address(),
                headers: headers,
                usedRange: usedRange.Address()
            }
        };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 排序
function sortRange(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.range);
        var keyCol = sheet.Range(params.keyColumn);
        var order = params.order === "desc" ? 2 : 1;
        range.Sort(keyCol, order);
        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 筛选
function autoFilter(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.range);
        if (params.criteria) {
            range.AutoFilter(params.field, params.criteria);
        } else {
            range.AutoFilter();
        }
        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 创建图表
function createChart(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.dataRange);
        var chartTypes = { column: 51, bar: 57, line: 4, pie: 5, area: 1, scatter: -4169 };
        var chartType = chartTypes[params.chartType] || 51;

        var left = params.left || (range.Left + range.Width + 20);
        var top = params.top || range.Top;

        var chartObj = sheet.ChartObjects().Add(left, top, 400, 300);
        chartObj.Chart.SetSourceData(range);
        chartObj.Chart.ChartType = chartType;

        if (params.title) {
            chartObj.Chart.HasTitle = true;
            chartObj.Chart.ChartTitle.Text = params.title;
        }
        return { success: true, data: { chartName: chartObj.Name } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 去重
function removeDuplicates(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.range);
        var cols = params.columns || [1];
        range.RemoveDuplicates(cols, 1);
        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 导出图表为图片（Chart.Export 原生 API）
// 调用 chartObj.Chart.Export(FileName, FilterName) 实现 1:1 像素级图表导出
function exportChartAsImage(params) {
    try {
        var wb = Application.ActiveWorkbook;
        if (!wb) return { success: false, error: '没有打开的工作簿' };
        var sheet = getSheet(wb, params.sheet);
        var outputPath = params.outputPath || params.path;
        if (!outputPath) return { success: false, error: '缺少输出路径 outputPath' };
        var chartName = params.chartName;
        if (!chartName) return { success: false, error: '缺少图表名 chartName' };
        var rawFormat = (params.format || 'PNG').toString().toUpperCase();
        // JPEG 在底层 COM 中按 JPG 滤镜处理
        var filterName = rawFormat === 'JPEG' ? 'JPG' : rawFormat;

        var chartObj = sheet.ChartObjects(chartName);
        chartObj.Chart.Export(outputPath, filterName);
        return {
            success: true,
            data: {
                chartName: chartName,
                outputPath: outputPath,
                format: filterName
            }
        };
    } catch (e) {
        return { success: false, error: '导出图表为图片失败: ' + e.message };
    }
}

// 导出区域为图片（Range.CopyPicture + 临时 ChartObject + Chart.Export 经典做法）
// 1. range.CopyPicture(xlScreen=1, xlBitmap=2) 复制区域到剪贴板
// 2. ChartObjects.Add(0, 0, range.Width, range.Height) 创建同尺寸临时图表
// 3. tempChart.Chart.Paste() 把剪贴板位图贴入图表
// 4. tempChart.Chart.Export(outputPath, filterName) 导出
// 5. tempChart.Delete() 清理临时图表
function exportRangeAsImage(params) {
    var tempChart = null;
    try {
        var wb = Application.ActiveWorkbook;
        if (!wb) return { success: false, error: '没有打开的工作簿' };
        var sheet = getSheet(wb, params.sheet);
        var outputPath = params.outputPath || params.path;
        if (!outputPath) return { success: false, error: '缺少输出路径 outputPath' };
        if (!params.range) return { success: false, error: '缺少区域 range' };
        var rawFormat = (params.format || 'PNG').toString().toUpperCase();
        var filterName = rawFormat === 'JPEG' ? 'JPG' : rawFormat;

        var range = sheet.Range(params.range);
        // xlScreen=1 (Appearance), xlBitmap=2 (Format)
        range.CopyPicture(1, 2);

        // 创建同尺寸临时 ChartObject 承载剪贴板位图
        tempChart = sheet.ChartObjects().Add(0, 0, range.Width, range.Height);
        tempChart.Activate();
        tempChart.Chart.Paste();
        tempChart.Chart.Export(outputPath, filterName);
        tempChart.Delete();
        tempChart = null;

        return {
            success: true,
            data: {
                range: params.range,
                outputPath: outputPath,
                format: filterName
            }
        };
    } catch (e) {
        // 异常清理：如果临时图表未删除，尝试回滚
        if (tempChart) {
            try { tempChart.Delete(); } catch (cleanupErr) {}
        }
        return { success: false, error: '导出区域为图片失败: ' + e.message };
    }
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        getActiveWorkbook: getActiveWorkbook,
        getCellValue: getCellValue,
        setCellValue: setCellValue,
        getRangeData: getRangeData,
        setRangeData: setRangeData,
        setFormula: setFormula,
        getExcelContext: getExcelContext,
        sortRange: sortRange,
        autoFilter: autoFilter,
        createChart: createChart,
        removeDuplicates: removeDuplicates,
        exportChartAsImage: exportChartAsImage,
        exportRangeAsImage: exportRangeAsImage
    };
}
