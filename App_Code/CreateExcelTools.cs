/*********************************************************************************
 * 机器名称：YUEYEZ-PC
 * 公司名称：励图高科
 * 命名空间：
 * 文件名：  CreateExcel
 * 创建人：  胡勇超
 * 创建时间：2015/7/16 17:37:29
 * 描述：
 ********************************************************************************/
using System.Data;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using System;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Excel.Application;
using Chart = Microsoft.Office.Interop.Excel.Chart;
using DataTable = System.Data.DataTable;
using Range = Microsoft.Office.Interop.Excel.Range;
using XlHAlign = Microsoft.Office.Interop.Excel.XlHAlign;
using XlRowCol = Microsoft.Office.Interop.Excel.XlRowCol;

/// <summary>
/// CreateExcel 的摘要说明
/// </summary>
public class CreateExcelTools : IDisposable
{
    private object omissing = Missing.Value;
    private Application xlApp;
    private Workbook xlBook;
    public Worksheet xlSheet;

    /// <summary>
    /// 初始化调用office生成excel的程序
    /// </summary>
    /// <param name="isVisible"></param>
    public void initExcel(bool isVisible)
    {
        try
        {
            xlApp = new Application();
            xlBook = xlApp.Workbooks.Add(Missing.Value);    //创建Excel工作表  （）中参数 System.Reflection.Missing.Value 或  XlWBATemplate.xlWBATWorksheet
            xlSheet = (Worksheet)xlBook.Worksheets.get_Item(1);
            //显示导出的过程
            xlApp.Visible = isVisible;
            // xlApp.DisplayAlerts = false;
        }
        catch (Exception ec)
        {
            throw ec;
        }
    }

    /// <summary>
    /// 将图片插入到excel中
    /// </summary>
    public void AddPicture(string picPath)
    {
        xlSheet.Shapes.AddPicture(picPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 100, 200, 200, 300);
    }

    /// <summary>
    /// 插入图文框
    /// </summary>
    /// <param name="text"></param>
    public void AddTextEffect(string text)
    {
        xlSheet.Shapes.AddTextEffect(Microsoft.Office.Core.MsoPresetTextEffect.msoTextEffect1, text, "Red", 15, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 150, 200);
    }

    /// <summary>
    /// 指定范围内的数据生成图表
    /// </summary>
    /// <param name="rangstart">Excel的选择表达式A1:F1</param>
    /// <param name="chartTitle"></param>
    /// <param name="xlable"></param>
    /// <param name="ylable"></param>
    public void AddChart(string rangs, string chartTitle, string xlable, string ylable)
    {
        //Chart xlChart = (Chart)xlBook.Charts.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);
        //Range chartRage = xlSheet.get_Range(rangstart, rangend);
        //xlChart.ChartWizard(chartRage, XlChartType.xl3DColumn, Missing.Value, XlRowCol.xlColumns, 1, 1, true, chartTitle, xlable, ylable, Missing.Value);
        Chart xlChart = xlSheet.Shapes.AddChart(XlChartType.xlColumnClustered, 100, 100, 300, 200).Chart;
        //Range chartRage = xlSheet.get_Range(rangstart, rangend);
        Range chartRage = xlSheet.get_Range(rangs, omissing);

        //xlChart.ChartWizard(chartRage, XlChartType.xl3DColumn, Missing.Value, XlRowCol.xlColumns, 1, 1, true, chartTitle, xlable, ylable, Missing.Value);
        xlChart.SetSourceData(chartRage, omissing);
        xlChart.HasTitle = true;
        xlChart.ChartTitle.Text = chartTitle;
    }

    /// <summary>
    /// 添加头部
    /// </summary>
    /// <param name="heads">头内容的数组</param>
    /// <param name="start">头开始的坐标，A1...</param>
    /// <param name="end">头结束的坐标,F1...</param>
    public void AddHead(string[] heads, string start, string end)
    {
        //Range xlRange = xlSheet.get_Range("A1", "F1");
        Range xlRange = xlSheet.get_Range(start, end);
        xlRange.set_Value(Missing.Value, heads);
        xlRange.Font.Bold = true;
        xlRange.Font.Name = "宋体";
        xlRange.Font.Size = 10;
        xlRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
    }

    /// <summary>
    /// 添加数据
    /// </summary>
    /// <param name="datas"></param>
    /// <param name="start"></param>
    public void AddData(string[,] datas, string start)
    {
        Range range = xlSheet.get_Range(start, omissing);
        int xs = range.Row;
        int ys = range.Column;

        int rowNum = datas.GetLength(0);
        int colNum = datas.GetLength(1);

        for (int i = 0; i < rowNum; i++)
        {
            for (int j = 0; j < colNum; j++)
            {
                xlSheet.Cells[xs + i, ys + j] = datas[i, j];
            }
        }
    }

    /// <summary>
    /// 保存Excel文件
    /// </summary>
    /// <param name="basePath"></param>
    /// <param name="fileName"></param>
    public void SaveExcel(string basePath, string fileName)
    {
        if (!Directory.Exists(basePath))
        {
            Directory.CreateDirectory(basePath);
        }
        string filePath = basePath + fileName + ".xls";

        xlBook.SaveAs(filePath, omissing, omissing, omissing, omissing, omissing, XlSaveAsAccessMode.xlNoChange, omissing, omissing, omissing, omissing, omissing);
    }

    #region 不使用office组件，直接导出excel的方法

    /// <summary>
    /// 获得创建excel的语句
    /// </summary>
    /// <param name="dt"></param>
    /// <param name="colHeader"></param>
    /// <returns></returns>
    public static string CreateExcel(DataTable dt, string[] colHeader)
    {
        string excelStr = "";

        string colHeaders = "", ls_item = "";
        int i = 0;
        for (i = 0; i < colHeader.Length - 1; i++)
            colHeaders += colHeader[i] + "\t";
        colHeaders += colHeader[i] + "\n";
        excelStr += colHeaders;


        //逐行处理数据
        DataRow[] myRow = dt.Select("");
        foreach (DataRow row in myRow)
        {
            //在当前行中，逐列获得数据，数据之间以\t分割，结束时加回车符\n
            for (i = 0; i < dt.Columns.Count - 1; i++)
                ls_item += row[i] + "\t";
            ls_item += row[i] + "\n";

            //当前行数据写入HTTP输出流，并且置空ls_item以便下行数据
            excelStr += ls_item;
            ls_item = "";
        }

        return excelStr;
    }


    #endregion

    #region 善后工作

    public void CloseExcel()
    {
        if (xlApp != null)
        {
            xlBook.Close(false, omissing, omissing);
            xlBook = null;
            xlApp.Quit();
            xlApp = null;
        }
    }

    ~CreateExcelTools()
    {
        if (xlApp != null)
        {
            xlBook.Close(false, omissing, omissing);
            xlBook = null;
            xlApp.Quit();
            xlApp = null;
        }
    }

    public void Dispose()
    {
        if (xlApp != null)
        {
            xlBook.Close(false, omissing, omissing);
            xlBook = null;
            xlApp.Quit();
            xlApp = null;
        }
    }

    #endregion
}