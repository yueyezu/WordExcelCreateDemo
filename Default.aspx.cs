/*********************************************************************************
 * 机器名称：YUEYEZ-PC
 * 公司名称：励图高科
 * 文件名：  Default
 * 创建人：  胡勇超
 * 创建时间：2015/2/12 8:28:49
 * 描述：
 ********************************************************************************/
using System;
using System.Collections.Generic;
using System.Data;
using System.Web;

public partial class _Default : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }

    /// <summary>
    /// 测试生成word
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    protected void Button1_Click(object sender, EventArgs e)
    {
        string basePath = HttpRuntime.AppDomainAppPath + "\\wordReport\\";
        List<string[]> lists1 = new List<string[]>()
            {
                new []{"类群","极值","日本竹荚鱼","鹰爪虾","翡翠贻贝","大鳞鲻","中华管鞭虾","波罗门赤虾（宽突赤虾）","叫姑鱼"},
                new []{"鱼类","范围","-","-","-","-","-","-","-"},
                new []{"","平均值","0.0000","0.0000","0.0000","0.0000","0.0000","0.0000","0.01"},
                new []{"","平均值","0.0000","0.0000","0.0000","0.0000","0.0000","0.0000","0.01"},
                new []{"","平均值","0.0000","0.0000","0.0000","0.0000","0.0000","0.0000","0.01"},
            };
        DataTable tempTab = ConvertDataTable(lists1);

        string[,] data = new string[4, 5];
        data[0, 1] = "第一月";
        data[0, 2] = "第二月";
        data[0, 3] = "第三月";
        data[0, 4] = "第四月";
        data[1, 0] = "东部";
        data[1, 1] = "50";
        data[1, 2] = "50";
        data[1, 3] = "40";
        data[1, 4] = "50";
        data[2, 0] = "西部";
        data[2, 1] = "60";
        data[2, 2] = "60";
        data[2, 3] = "70";
        data[2, 4] = "80";
        //data[3,6] = "0";     
        data[3, 0] = "中部";
        data[3, 1] = "50";
        data[3, 2] = "50";
        data[3, 3] = "40";
        data[3, 4] = "50";


        using (CreatWordTools wordTools = new CreatWordTools())
        {
            wordTools.initWord("测试文档", "励图高科信息技术有限公司", true);
            wordTools.InsertNewPara("今天做一个app时发现一个问题，应用html5中的video标签加载视频，在Android手机上默认播放大小，但是换成iPhone手机上出问题了，默认弹出全屏播放，查找了好多论坛，都没有谈论这个的。然后几经波折终于找到其解决的方法了，在video标签下的source中加入这个-webkit-play");
            wordTools.InsertTable(tempTab, "test tables");
            wordTools.InsertNewPara("");
            wordTools.InsertImage(basePath + "tu.jpg", false, true, 1);

            wordTools.AddSimpleChart(data, 1);

            wordTools.SaveWord(basePath, "test");
        }
    }

    /// <summary>
    /// 将list转化成DataTable
    /// </summary>
    /// <param name="datas"></param>
    /// <returns></returns>
    private DataTable ConvertDataTable(List<string[]> datas)
    {
        string[] names = datas[0];

        DataTable table = new DataTable();
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
}