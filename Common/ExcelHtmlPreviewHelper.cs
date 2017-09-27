/// <summary>
/// Excel转html预览
/// </summary>
/// <param name="sheet">Excel单元格</param>
/// <returns></returns>
public string ExcelPreViewHtmlHelper(ISheet sheet)
{
    var html = string.Empty;
    int rowsCount = sheet.PhysicalNumberOfRows;
    var headers = sheet.GetRow(1).Cells;//获取表格数据的同时将监测项名称转换为id
    var headerEffect = new List<int>();//记录有效的表头列号
    if (headers.Count > 0)
    {
        html += @"<table class='table table-bordered table-striped table-hover fancyTable' >
                    <thead>
                        <tr>";
        foreach (var temp in headers)
        {
            if (temp.GetCellValue(false).IsNullOrEmpty()) { continue; }
            html += "<th style='width: 100px;'>" + temp.GetCellValue(false) + "</th>";
            headerEffect.Add(temp.ColumnIndex);
        }
        html += @"</tr></thead>
                    <tbody>";
        //循环行
        for (int x = 0; x < rowsCount; x++)
        {
            if (sheet.GetRow(x) == null) { rowsCount++; continue; }
            //为防止空行，从0开始读,而实际需要读取的值从2开始
            if (x < 2) { continue; }
            var row = sheet.GetRow(x);

            //循环列
            var style = "";
            html += "<tr>";
            for (var c = 0; c <= headerEffect.Max(); c++)
            {
                if (!headerEffect.Contains(c)) { continue; }
                html += "<td{0}>".FormatTo(style);
                if (row.GetCell(c) != null)
                {
                    if (row.GetCell(c).CellType == CellType.Numeric)
                    {
                        if (DateUtil.IsCellDateFormatted(row.GetCell(c)))
                        {
                            html += row.GetCell(c).DateCellValue.ToString(Constant.TimeFormat);
                        }
                        else
                        {
                            html += row.GetCell(c).NumericCellValue;
                        }
                    }
                    else
                    {
                        html += row.GetCell(c).GetCellValue(false);
                    }
                }
                html += "</td>";
            }
            html += "</tr>";
        }
        html += @"</tbody>
    </table>";
    }
    return html;
}