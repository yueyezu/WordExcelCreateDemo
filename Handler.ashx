<%@ WebHandler Language="C#" Class="Handler" %>
/*********************************************************************************
 * 机器名称：YUEYEZ-PC
 * 公司名称：励图高科
 * 文件名：  Handler
 * 创建人：  胡勇超
 * 创建时间：2015/2/12 11:32:56
 * 描述：
 ********************************************************************************/
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Web;

public class Handler : IHttpHandler
{
    public void ProcessRequest(HttpContext context)
    {
        context.Response.ContentType = "text/plain";

        string action = context.Request["action"];
        switch (action)
        {
            case "createword":
                StartCreateWord(context);
                break;
            case "getword":
                GetWord(context);
                break;
            case "createExcel":
                CreateExcel(context);
                break;
            case "getexcel":
                GetExcel(context);
                break;
        }
    }


    /// <summary>
    /// 导出excel
    /// </summary>
    /// <param name="context"></param>
    private void CreateExcel(HttpContext context)
    {
        string basePath = HttpRuntime.AppDomainAppPath + "excelReport\\";
        string fileName = context.Request["excelName"];
        string[] header = new[] { "类型", "未完成数", "完成比例" };
        string[,] lists1 = {
           {"月任务","33","12"},
           {"周任务","43","67"},
           {"日任务","53","50"}
        };

        using (CreateExcelTools excelTools = new CreateExcelTools())
        {
            excelTools.initExcel(true);
            excelTools.AddHead(header, "A1", "C1");
            excelTools.AddData(lists1, "A2");
            excelTools.AddPicture(basePath + "tu.jpg");
            excelTools.AddChart("A1:A4,C1:C4", "测试的一个比例图表\n只是测试一下是不是可以使用", "类型", "比例");
            excelTools.SaveExcel(basePath, fileName);
        }

        context.Response.Clear();
        context.Response.Write("{\"success\":true}");
        context.Response.End();
    }

    /// <summary>
    /// 获取Excel文件
    /// </summary>
    /// <param name="context"></param>
    private void GetExcel(HttpContext context)
    {
        string excelName = context.Request["excelName"];
        string basePath = HttpRuntime.AppDomainAppPath + "excelReport\\";

        string wordPath = basePath + excelName + ".xls";
        context.Response.Clear();
        if (File.Exists(wordPath))
        {
            context.Response.Buffer = true;
            //Stream stream = new FileStream(wordPath, FileMode.Open);
            //byte[] tempBytes = new byte[stream.Length];
            //stream.Read(tempBytes, 0, tempBytes.Length);
            context.Response.AddHeader("Content-Type", "application/vnd.ms-excel");
            //HttpContext.Current.Response.AddHeader("Content-Disposition",   "inline;");
            context.Response.AddHeader("Content-Disposition", "attachment;filename=" + excelName + ".xls");
            //context.Response.BinaryWrite(tempBytes);
            context.Response.WriteFile(wordPath);
        }
        else
        {
            context.Response.Write("文件不存在");
        }
        context.Response.End();
    }

    /// <summary>
    /// 导出excel，不带图片，直接下载
    /// </summary>
    /// <param name="context"></param>
    private void CreateExcelPos(HttpContext context)
    {
        string basePath = HttpRuntime.AppDomainAppPath + "excelReport\\";
        string fileName = context.Request["excelName"];

        context.Response.AddHeader("Content-Type", "application/vnd.ms-excel");
        //resp.AddHeader("Content-Disposition", "inline;");
        context.Response.ContentEncoding = System.Text.Encoding.GetEncoding("GB2312");
        context.Response.AppendHeader("Content-Disposition", "attachment;filename=" + fileName);

        string[] header = new[] { "test1", "13ttt", "12ttt" };
        List<string[]> lists1 = new List<string[]>()
        {
           new [] {"test1","testt2","ttt3"},
           new [] {"12","13","12"},
           new [] {"12","13","12"},
           new [] {"12","13","12"}
        };
        DataTable edt = ConvertDataTable(lists1);
        string excelStr = CreateExcelTools.CreateExcel(edt, header);
        context.Response.Write(excelStr);

        //FileStream stream = new FileStream(basePath + "tu.jpg", FileMode.Open);
        //long len = stream.Length;
        //byte[] bytes = new byte[len];
        //stream.Read(bytes, 0, (int)len);
        //context.Response.BinaryWrite(bytes);
        context.Response.End();
    }

    /// <summary>
    /// 下载word文件
    /// </summary>
    /// <param name="context"></param>
    private void GetWord(HttpContext context)
    {
        string wordName = context.Request["wordName"];
        string basePath = HttpRuntime.AppDomainAppPath + "wordReport\\";

        string wordPath = basePath + wordName + ".doc";
        context.Response.Clear();
        if (File.Exists(wordPath))
        {
            //Stream stream = new FileStream(wordPath, FileMode.Open);
            //byte[] tempBytes = new byte[stream.Length];
            //stream.Read(tempBytes, 0, tempBytes.Length);
            context.Response.AddHeader("Content-Type", "application/msword");
            //HttpContext.Current.Response.AddHeader("Content-Disposition",   "inline;");
            context.Response.AddHeader("Content-Disposition", "attachment;filename=" + wordName + ".doc");
            //context.Response.BinaryWrite(tempBytes);
            context.Response.WriteFile(wordPath);
        }
        else
        {
            context.Response.Write("文件不存在");
        }
        context.Response.End();
    }


    /// <summary>
    /// 开始创建word
    /// </summary>
    /// <param name="context"></param>
    private void StartCreateWord(HttpContext context)
    {
        string wordName = context.Request["wordName"];

        string basePath = HttpRuntime.AppDomainAppPath + "wordReport\\";
        List<string[]> lists1 = new List<string[]>()
            {
                new []{"类群","极值","日本竹荚鱼","鹰爪虾","翡翠贻贝","大鳞鲻","中华管鞭虾","波罗门赤虾（宽突赤虾）","叫姑鱼"},
                new []{"鱼类","范围","-","-","-","-","-","-","-"},
                new []{"","平均值","0.0000","0.0000","0.0000","0.0000","0.0000","0.0000","0.01"},
                new []{"","平均值","0.0000","0.0000","0.0000","0.0000","0.0000","0.0000","0.01"},
                new []{"","平均值","0.0000","0.0000","0.0000","0.0000","0.0000","0.0000","0.01"},
            };
        DataTable tempTab = ConvertDataTable(lists1);

        string[,] data =
        {
            {"","第一月","第二月","第三月","第四月"},
            {"东部","50","50","50","50"},
            {"西部","60","60","60","60"},
            {"中部","40","40","40","40"},
        };

        using (CreatWordTools wordTools = new CreatWordTools())
        {
            wordTools.initWord("测试文档", "励图高科信息技术有限公司", false);
            wordTools.InsertNewPara("今天做一个app时发现一个问题，应用html5中的video标签加载视频，在Android手机上默认播放大小，但是换成iPhone手机上出问题了，默认弹出全屏播放，查找了好多论坛，都没有谈论这个的。然后几经波折终于找到其解决的方法了，在video标签下的source中加入这个-webkit-play");
            wordTools.InsertTable(tempTab, "test tables");
            wordTools.InsertNewPara("");
            wordTools.InsertImage(basePath + "tu.jpg", false, true, 1);
            wordTools.AddSimpleChart(data, 1);

            wordTools.SaveWord(basePath, wordName);
        }

        context.Response.Clear();
        context.Response.Write("{\"success\":true}");
        context.Response.End();
    }


    /// <summary>
    /// 将list转化成DataTable
    /// </summary>
    /// <param name="datas"></param>
    /// <returns></returns>
    private DataTable ConvertDataTable(List<string[]> datas)
    {
        string[] names = datas[0];

        System.Data.DataTable table = new System.Data.DataTable();
        for (int i = 0; i < names.Length; i++)
        {
            table.Columns.Add(names[i], typeof(string));
        }

        for (int i = 1; i < datas.Count; i++)
        {
            DataRow drtemp = table.NewRow();
            for (int k = 0; k < datas[i].Length; k++)
            {
                drtemp[k] = datas[i][k];
            }
            table.Rows.Add(drtemp);
        }

        return table;
    }

    public bool IsReusable
    {
        get
        {
            return false;
        }
    }
}