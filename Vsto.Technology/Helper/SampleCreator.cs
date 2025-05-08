using System.Drawing;
using Microsoft.Office.Interop.Excel;

namespace Vsto.Technology.Helper;

public abstract class SampleCreator
{
    public static void Create(Range target)
    {
        var dataArray = new string[11, 5]
        {
            { "", "工艺报价模板", "", "", "" },
            { "", "父节点", "设备", "加工费类别", "修正系数" },
            { "数据示例", "", "分区名", "", "" },
            { "", "分区名", "子分区名", "", "0.95" },
            { "1", "", "前处理区", "", "0.95" },
            { "1.1", "前处理区", "前处理1", "C", "0.95" },
            { "1.2", "前处理区", "前处理2", "C", "0.95" },
            { "1.3", "前处理区", "前处理3", "C", "0.95" },
            { "2", "", "电泳区", "", "0.95" },
            { "2.1", "电泳区", "电泳区1", "C", "0.95" },
            { "2.2", "电泳区", "电泳区2", "C", "0.95" }
        };

        var worksheet = target.Worksheet;
        worksheet.Name = "原始表";

        worksheet.Cells[1, 1].Resize[dataArray.GetLength(0), dataArray.GetLength(1)].Value = dataArray;
        worksheet.Columns["B:E"].HorizontalAlignment = XlHAlign.xlHAlignCenter;

        worksheet.Range["B1", "E1"].Merge();
        worksheet.Range["B1", "E4"].Font.Bold = true;
        SetBorder(worksheet.Range["B1", "E1"], [XlBordersIndex.xlEdgeBottom]);

        worksheet.Range["A3", "A4"].Merge();

        SetBorder(worksheet.Range["B2", "E2"], [XlBordersIndex.xlEdgeBottom]);

        SetBorder(worksheet.Range["B4", "E4"], [XlBordersIndex.xlEdgeBottom]);
        worksheet.Range["B3", "E4"].Interior.Color = ColorTranslator.ToOle(Color.AliceBlue);

        worksheet.UsedRange.Columns.AutoFit();
    }

    private static void SetBorder(Range range, XlBordersIndex[] borderIndexes)
    {
        foreach (var borderIndex in borderIndexes)
        {
            var border = range.Borders[borderIndex];
            border.LineStyle = XlLineStyle.xlContinuous;
            border.Weight = XlBorderWeight.xlMedium;
            border.Color = ColorTranslator.ToOle(Color.Black);
        }
    }
}