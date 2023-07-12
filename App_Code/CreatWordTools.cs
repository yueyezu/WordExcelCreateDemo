/*********************************************************************************
 * 机器名称：YUEYEZ-PC
 * 公司名称：励图高科
 * 文件名：  CreateSeaWord
 * 创建人：  胡勇超
 * 创建时间：2015/2/12 9:02:37
 * 描述：关于该文档的api可以参考网址：https://msdn.microsoft.com/zh-cn/library/office/ff198329.aspx
 ********************************************************************************/
using System;
using System.Web;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Graph;

/// <summary>
/// 创建海区月报
/// </summary>
public class CreatWordTools : IDisposable
{
    private object Nothing = Missing.Value;

    private _Application wordApp;   //word的应用程序调用
    private _Document wordDocument; //word文档对象

    /// <summary>
    /// 获取创建的文档对象
    /// </summary>
    public _Document WordDocument
    {
        get { return wordDocument; }
    }
    private object oEndOfDoc = "\\endofdoc";    //获得当前文档末尾

    object oPageBreak = WdBreakType.wdPageBreak;//分页符
    public CreatWordTools()
    {
        wordApp = new ApplicationClass();
        wordDocument = wordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);//.Add(Template, NewTemplate, DocumentType, Visible)
    }

    /// <summary>
    /// 新建Word文档
    /// </summary>
    /// <param name="wordTitle">Word标题</param>
    /// <param name="headerText">页眉的值</param>
    /// <param name="wordVisible">Word是否可见</param>
    public void initWord(string wordTitle, string headerText, bool wordVisible)
    {
        try
        {
            wordApp.Visible = wordVisible;
            //wordApp.ActiveWindow.View.Type = WdViewType.wdOutlineView;//大纲视图

            /******以下为设置文档页面的情况****************/
            wordDocument.PageSetup.PaperSize = WdPaperSize.wdPaperA4;//设置页面A4纸
            wordDocument.PageSetup.Orientation = WdOrientation.wdOrientPortrait;//页面方向
            wordDocument.PageSetup.TopMargin = 57.0f;//设置边距
            wordDocument.PageSetup.BottomMargin = 57.0f;
            wordDocument.PageSetup.LeftMargin = 57.0f;
            wordDocument.PageSetup.RightMargin = 57.0f;
            wordDocument.PageSetup.HeaderDistance = 28; //页眉

            /********** 设置页眉 ******************/
            wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekCurrentPageHeader;//进入页眉设置
            wordApp.ActiveWindow.ActivePane.Selection.InsertAfter(headerText);//插入页眉,以及其内容
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;//居左对齐
            wordApp.ActiveWindow.ActivePane.Selection.Borders[WdBorderType.wdBorderBottom].Visible = false; //
            wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekMainDocument;//进入文本内容设置区

            /*********** 设置页脚页码 ***********/
            PageNumbers pns = wordApp.Selection.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].PageNumbers;
            pns.NumberStyle = WdPageNumberStyle.wdPageNumberStyleNumberInDash; //指定要应用于页码的样式。带划线数字样式。参考:https://msdn.microsoft.com/ZH-CN/library/office/ff845847.aspx
            pns.HeadingLevelForChapter = 0;//设置文档中章节标题应用的标题级别样式
            pns.IncludeChapterNumber = false;//不包含章节号码  //
            pns.ChapterPageSeparator = WdSeparatorType.wdSeparatorHyphen;   //设置章节号和页码之间的分隔字符。参考：https://msdn.microsoft.com/ZH-CN/library/office/ff834539.aspx
            pns.RestartNumberingAtSection = false;  //如果指定节的页码从1重新开始，则该属性值为 True。Boolean 类型，可读写。
            pns.StartingNumber = 0; //设置注释的起始编号、 行号或页码
            object pagenmbetal = WdPageNumberAlignment.wdAlignPageNumberCenter; //指定段落的对齐方式
            object first = true;   //是否是起始页
            wordApp.Selection.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].PageNumbers.Add(ref pagenmbetal, ref first); //.Add(PageNumberAlignment, FirstPage)

            /******** 正文的设置 *********/
            wordApp.Selection.ParagraphFormat.LineSpacing = 15f;//段落间距
        }
        catch (Exception ex)
        {
            throw ex;
        }
    }

    /// <summary>
    /// 插入新段落以及设置其样式
    /// </summary>
    /// <param name="text">插入段落内容</param>
    public void InsertNewPara(string text)
    {
        object range = wordDocument.Bookmarks.get_Item(ref oEndOfDoc).Range;
        Paragraph newPara;//插入段落内容
        newPara = wordDocument.Content.Paragraphs.Add(ref range);
        newPara.Range.Text = text;
        newPara.Range.Font.Size = 12;
        newPara.Range.Font.Bold = 0;
        newPara.Format.SpaceAfter = 2;
        newPara.CharacterUnitFirstLineIndent = 2;
        newPara.Range.InsertParagraphAfter();
    }

    /// <summary>
    /// 插入标题
    /// </summary>
    public void AddHeader(string text)
    {
        object range = wordDocument.Bookmarks.get_Item(ref oEndOfDoc).Range;
        wordDocument.Paragraphs.Last.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        wordDocument.Paragraphs.Last.Range.Font.Bold = 1;
        wordDocument.Paragraphs.Last.Range.Font.Size = 13;
        wordDocument.Paragraphs.Last.Range.Text = text + "\n";
    }

    /// <summary>
    /// 向Word中插入表格以及表格单元格的数据
    /// </summary>
    /// <param name="range">插入表格的位置</param>
    /// <param name="result">表格中内容的数据源</param>
    /// <returns></returns>
    public Table InsertTable(System.Data.DataTable result, string title)
    {
        Microsoft.Office.Interop.Word.Range range = wordDocument.Bookmarks.get_Item(ref oEndOfDoc).Range;

        Table newTable = wordDocument.Tables.Add(range, result.Rows.Count + 2, result.Columns.Count, ref Nothing, ref Nothing);
        newTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;//wdLineStyleThickThinLargeGap;
        newTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
        //newTable.Range.ParagraphFormat.SpaceAfter = 6;
        newTable.AllowAutoFit = false;
        newTable.Cell(1, 1).Range.Text = title;//表格标题                   
        newTable.Cell(1, 1).Range.Font.Color = WdColor.wdColorDarkBlue;//设置字体颜色
        newTable.Cell(1, 1).Range.Bold = 4;
        newTable.Cell(1, 1).Range.Font.Size = 16;
        newTable.Cell(1, 1).Merge(newTable.Cell(1, result.Columns.Count));//合并单元格      
        newTable.Select();
        newTable.Rows.Height = 2f;
        wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
        wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        //往表格中写入数据

        for (int column = 0; column < result.Columns.Count; column++)
        {
            newTable.Cell(2, column + 1).Range.Text = result.Columns[column].ColumnName;
            newTable.Cell(2, column + 1).Range.Font.Size = 9;
        }
        for (int row = 0; row < result.Rows.Count; row++)
        {
            for (int column = 0; column < result.Columns.Count; column++)
            {
                newTable.Cell(row + 3, column + 1).Range.Text = result.Rows[row][column].ToString();
                newTable.Cell(row + 3, column + 1).Range.Font.Size = 9;
            }
        }

        for (int i = 1; i <= newTable.Columns.Count; i++)
        {
            newTable.Cell(2, i).Range.Shading.ForegroundPatternColor = WdColor.wdColorLightBlue;
        }

        return newTable;
    }

    /// <summary>
    /// 向Word中插入图片
    /// </summary>
    /// <param name="FileName">图片的绝对路径</param>
    /// <param name="LinkToFile"></param>
    /// <param name="SaveWithDocument"></param>
    /// <param name="ImageNumber">图片编号</param>
    public void InsertImage(string FileName, object LinkToFile, object SaveWithDocument, int ImageNumber)
    {
        //直接控制图片的位置不方便，先将其放在一个表格中
        try
        {
            Microsoft.Office.Interop.Word.Range range = wordDocument.Bookmarks.get_Item(ref oEndOfDoc).Range;

            Table newTable1 = wordDocument.Tables.Add(range, 1, 1, ref Nothing, ref Nothing);
            newTable1.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone;
            newTable1.Borders.InsideLineStyle = WdLineStyle.wdLineStyleNone;
            newTable1.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;//整个表格居中显示
            newTable1.Cell(1, 1).Select();
            object Anchor1 = newTable1.Cell(1, 1).Range;
            wordDocument.Application.ActiveDocument.InlineShapes.AddPicture(FileName, ref LinkToFile, ref SaveWithDocument, ref Anchor1);
            wordDocument.Application.ActiveDocument.InlineShapes[ImageNumber].Width = 300f;//图片宽度
            wordDocument.Application.ActiveDocument.InlineShapes[ImageNumber].Height = 200f;//图片高度              
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            newTable1.Cell(1, 1).Row.Height = wordDocument.Application.ActiveDocument.InlineShapes[ImageNumber].Height;//设置表格的长宽与图片的相等

            newTable1.Cell(1, 1).Column.Width = wordDocument.Application.ActiveDocument.InlineShapes[ImageNumber].Width + 10;
        }
        catch
        {

        }
    }

    /// <summary>
    /// 在word中插入图表
    /// </summary>
    /// <param name="data">要插入图表中的数据</param>
    /// <param name="chartType">图表类型：1-柱状图，2、饼状图，3、折线图</param>
    public void AddSimpleChart(string[,] data, int chartType)
    {
        //插入chart    
        object oClassType = "MSGraph.Chart.8";
        Microsoft.Office.Interop.Word.Range range = wordDocument.Bookmarks.get_Item(ref oEndOfDoc).Range;
        InlineShape oShape = range.InlineShapes.AddOLEObject(ref oClassType, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);

        //Demonstrate use of late bound oChart and oChartApp objects to manipulate the chart object with MSGraph.  
        object oChart = oShape.OLEFormat.Object;
        object oChartApp = oChart.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, oChart, null);

        //Change the chart type to Line.  
        object[] Parameters = new Object[1];
        Parameters[0] = 4; //xlLine = 4  
        oChart.GetType().InvokeMember("ChartType", BindingFlags.SetProperty, null, oChart, Parameters);

        Microsoft.Office.Interop.Graph.Chart objChart = (Microsoft.Office.Interop.Graph.Chart)oShape.OLEFormat.Object;
        switch (chartType)
        {
            case 1:
                objChart.ChartType = XlChartType.xlColumnClustered;
                break;
            case 2:
                objChart.ChartType = XlChartType.xlPie;
                break;
            case 3:
                objChart.ChartType = XlChartType.xlLine;
                break;
        }

        //绑定数据  
        DataSheet dataSheet = objChart.Application.DataSheet;

        int rownum = data.GetLength(0);
        int columnnum = data.GetLength(1);
        for (int i = 1; i <= rownum; i++)
            for (int j = 1; j <= columnnum; j++)
            {
                dataSheet.Cells[i, j] = data[i - 1, j - 1];
            }

        objChart.Application.Update();
        oChartApp.GetType().InvokeMember("Update", BindingFlags.InvokeMethod, null, oChartApp, null);
        oChartApp.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, oChartApp, null);
        //设置大小  
        oShape.Width = wordApp.InchesToPoints(6.25f);
        oShape.Height = wordApp.InchesToPoints(3.57f);
    }

    /// <summary>
    /// 保存文档
    /// </summary>
    /// <param name="basePath">存放word临时文件的文件夹路径</param>
    /// <param name="wordName">word的名称</param>
    public void SaveWord(string basePath, string wordName)
    {
        //string dir = HttpRuntime.AppDomainAppPath + "\\wordReport\\";
        if (!Directory.Exists(basePath))
        {
            Directory.CreateDirectory(basePath);
        }
        object destFileName = basePath + wordName + ".doc";
        wordDocument.SaveAs(ref destFileName, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
    }

    #region 善后工作

    /// <summary>
    /// 关闭word对象
    /// </summary>
    public void CloseWord()
    {
        if (wordApp != null)
        {
            wordDocument.Close(ref Nothing, ref Nothing, ref Nothing);
            wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);
            wordApp = null;
        }
    }

    ~CreatWordTools()
    {
        if (wordApp != null)
        {
            wordDocument.Close(ref Nothing, ref Nothing, ref Nothing);
            wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);
            wordApp = null;
        }
    }

    public void Dispose()
    {
        if (wordApp != null)
        {
            wordDocument.Close(ref Nothing, ref Nothing, ref Nothing);
            wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);
            wordApp = null;
        }
    }

    #endregion
}