using dblist;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using System.Windows.Forms;

namespace 测试功能专用
{
    public partial class frmExcel : Form
    {
        public frmExcel()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //GetKEYnfo();
            //GetKEYnfo_renminribao();

            //  GetKEYnfo_kaoqincuowu();


            GetKEYnfo_列出客户所在的区或县();
        }

        public List<clsKEYinfo> GetKEYnfo()
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "Resources\\航宇星111 - 副本.xls";



            List<clsKEYinfo> MAPPINGResult = new List<clsKEYinfo>();
            try
            {
                List<clsKEYinfo> WANGYINResult = new List<clsKEYinfo>();
                System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                Microsoft.Office.Interop.Excel.Application excelApp;
                {

                    excelApp = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook analyWK = excelApp.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing,
                        "htc", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    MAPPINGResult = new List<clsKEYinfo>();
                    Microsoft.Office.Interop.Excel.Worksheet WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["Sheet2"];
                    Microsoft.Office.Interop.Excel.Range rng;
                    rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 16]];
                    int rowCount = WS.UsedRange.Rows.Count;
                    object[,] o = new object[1, 1];
                    o = (object[,])rng.Value2;
                    int wscount = analyWK.Worksheets.Count;

                    for (int i = 2; i <= rowCount; i++)
                    {
                        clsKEYinfo temp = new clsKEYinfo();

                        #region 基础信息

                        temp.diqu = "";
                        if (o[i, 1] != null)
                            temp.diqu = o[i, 1].ToString().Trim();

                        temp.fuzeren = "";
                        if (o[i, 2] != null)
                            temp.fuzeren = o[i, 2].ToString().Trim();


                        temp.maichangname = "";
                        if (o[i, 3] != null)
                            temp.maichangname = o[i, 3].ToString().Trim();

                        //卖场代码

                        temp.Maichangdaima = "";
                        if (o[i, 4] != null)
                            temp.Maichangdaima = o[i, 4].ToString().Trim();

                        temp.naishuidanwei = "";
                        if (o[i, 5] != null)
                            temp.naishuidanwei = o[i, 5].ToString().Trim();

                        #endregion
                        MAPPINGResult.Add(temp);
                    }

                    #region 读取模板列对应名称
                    List<clsR2RBankMappinginfo> R2RBankMapping = new List<clsR2RBankMappinginfo>();

                    WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["Sheet1"];

                    rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 16]];
                    rowCount = WS.UsedRange.Rows.Count;
                    o = new object[1, 1];
                    o = (object[,])rng.Value2;
                    wscount = analyWK.Worksheets.Count;

                    for (int i = 2; i <= rowCount; i++)
                    {
                        clsR2RBankMappinginfo temp = new clsR2RBankMappinginfo();

                        #region 基础信息

                        temp.jiaoyiriqi = "";
                        if (o[i, 3] != null)
                            temp.jiaoyiriqi = o[i, 3].ToString().Trim();
                        if (temp.jiaoyiriqi == null || temp.jiaoyiriqi == "")
                            continue;

                        temp.jiefang = "";
                        if (o[i, 4] != null)
                            temp.jiefang = o[i, 4].ToString().Trim();


                        temp.yongtu = "";
                        if (o[i, 8] != null)
                            temp.yongtu = o[i, 8].ToString().Trim();

                        #endregion
                        List<clsKEYinfo> mlist = MAPPINGResult.FindAll(oe => oe.maichangname != null && oe.maichangname == temp.jiaoyiriqi && oe.naishuidanwei == temp.yongtu).ToList();
                        if (mlist.Count > 0)
                        {

                            WS.Cells[i, 26] = "del";

                        }
                        R2RBankMapping.Add(temp);
                    }

                    #endregion
                    excelApp.Visible = true;
                    excelApp.ScreenUpdating = true;



                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: 01032" + ex);
                return null;

                throw;
            }
            return MAPPINGResult;

        }
        public List<clsKEYinfo> GetKEYnfo_renminribao()
        {
            //string path = AppDomain.CurrentDomain.BaseDirectory + "Resources\\人民日报.xlsx";
            //string pathtxt = AppDomain.CurrentDomain.BaseDirectory + "Resources\\reminribao.txt";

            //string path = AppDomain.CurrentDomain.BaseDirectory + "Resources\\Files\\0415\\LNRB.xlsx";
            string path = AppDomain.CurrentDomain.BaseDirectory + "Resources\\Files\\0415\\All_RMRB.xlsx";
         
            string pathtxt = AppDomain.CurrentDomain.BaseDirectory + "Resources\\reminribao_20190415.txt";

            //string path = AppDomain.CurrentDomain.BaseDirectory + "Resources\\人民日报.xlsx";
            //string pathtxt = AppDomain.CurrentDomain.BaseDirectory + "Resources\\reminribao.txt";

            string[] fileText = File.ReadAllLines(pathtxt);




            List<clsKEYinfo> MAPPINGResult = new List<clsKEYinfo>();


            for (int i = 0; i < fileText.Length; i++)
            {

                string[] fileTextG = System.Text.RegularExpressions.Regex.Split(fileText[i].Replace(" ", "").Replace("  ", "").Trim(), "、");

                for (int j = 0; j < fileTextG.Length; j++)
                {
                    if (fileTextG[j] != null && fileTextG[j] != "")
                    {
                        clsKEYinfo item = new clsKEYinfo();
                        item.maichangname = fileTextG[j];
                        MAPPINGResult.Add(item);
                    }
                }
            }




            try
            {

                System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                Microsoft.Office.Interop.Excel.Application excelApp;
                {

                    excelApp = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook analyWK = excelApp.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing,
                        "htc", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                    //Microsoft.Office.Interop.Excel.Worksheet WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets[1];
                    Microsoft.Office.Interop.Excel.Worksheet WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["Sheet2"];
                    Microsoft.Office.Interop.Excel.Range rng;
                    rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 16]];
                    int rowCount = WS.UsedRange.Rows.Count;
                    object[,] o = new object[1, 1];
                    o = (object[,])rng.Value2;
                    int wscount = analyWK.Worksheets.Count;



                    #region 读取模板列对应名称
                    List<clsR2RBankMappinginfo> R2RBankMapping = new List<clsR2RBankMappinginfo>();

                    //WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["bz"];
                    WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["Sheet2"];
                    rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 16]];
                    rowCount = WS.UsedRange.Rows.Count;
                    o = new object[1, 1];
                    o = (object[,])rng.Value2;
                    wscount = analyWK.Worksheets.Count;

                    for (int i = 2; i <= rowCount; i++)
                    {
                        clsR2RBankMappinginfo temp = new clsR2RBankMappinginfo();

                        #region 基础信息

                        //temp.jiaoyiriqi = "";
                        //if (o[i, 2] != null)
                        //    temp.jiaoyiriqi = o[i, 2].ToString().Trim();

                        //temp.jiefang = "";
                        //if (o[i, 5] != null)
                        //    temp.jiefang = o[i, 5].ToString().Trim();

                        temp.jiaoyiriqi = "";
                        if (o[i,4] != null)
                            temp.jiaoyiriqi = o[i,4].ToString().Trim();

                        temp.jiefang = "";
                        if (o[i,8] != null)
                            temp.jiefang = o[i,8].ToString().Trim();

                        #endregion



                        List<clsKEYinfo> mlist = MAPPINGResult.FindAll(oe => temp.jiaoyiriqi.Contains(oe.maichangname) || temp.jiefang.Contains(oe.maichangname)).ToList();
                        if (mlist.Count > 0)
                        {
                            if (mlist[0].maichangname == "蒙古")
                            {
                                if (mlist[0].maichangname == "蒙古" && (temp.jiaoyiriqi.Contains("内蒙古") || temp.jiefang.Contains("内蒙古")))
                                {

                                }
                                else
                                {
                                    //WS.Cells[i, 6] = "留下";
                                    WS.Cells[i, 10] = "留下";
                                }
                            }
                            else
                            {
                                //WS.Cells[i, 6] = "留下";
                                WS.Cells[i, 10] = "留下";
                            }
                        }
                        R2RBankMapping.Add(temp);
                    }

                    #endregion
                    excelApp.Visible = true;
                    excelApp.ScreenUpdating = true;



                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: 01032" + ex);
                return null;

                throw;
            }
            return MAPPINGResult;

        }


        public List<clsKEYinfo> GetKEYnfo_kaoqincuowu()
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "Resources\\异常考勤统计.xlsx";

            List<string> NAMEAlist = new List<string>();

            List<clsKEYinfo> MAPPINGResult = new List<clsKEYinfo>();
            try
            {
                List<clsKEYinfo> WANGYINResult = new List<clsKEYinfo>();
                System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                Microsoft.Office.Interop.Excel.Application excelApp;
                {

                    excelApp = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook analyWK = excelApp.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing,
                        "htc", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    MAPPINGResult = new List<clsKEYinfo>();
                    Microsoft.Office.Interop.Excel.Worksheet WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["2.13"];
                    Microsoft.Office.Interop.Excel.Range rng;
                    rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 16]];
                    int rowCount = WS.UsedRange.Rows.Count;
                    object[,] o = new object[1, 1];
                    o = (object[,])rng.Value2;
                    int wscount = analyWK.Worksheets.Count;

                    for (int i = 2; i <= rowCount; i++)
                    {
                        clsKEYinfo temp = new clsKEYinfo();

                        #region 基础信息

                        temp.diqu = "";
                        if (o[i, 1] != null)
                            temp.diqu = o[i, 1].ToString().Trim();

                        temp.fuzeren = "";
                        if (o[i, 2] != null)
                            temp.fuzeren = o[i, 2].ToString().Trim();

                        if (temp.fuzeren == null || temp.fuzeren == "")
                            continue;

                        temp.maichangname = "";
                        if (o[i, 3] != null)
                            temp.maichangname = o[i, 3].ToString().Trim();

                        //卖场代码

                        temp.Maichangdaima = "";
                        if (o[i, 4] != null)
                            temp.Maichangdaima = o[i, 4].ToString().Trim();

                        temp.naishuidanwei = "";
                        if (o[i, 5] != null)
                            temp.naishuidanwei = o[i, 5].ToString().Trim();


                        //6-12

                        temp.yinhangjiancheng = "";
                        if (o[i, 6] != null)
                            temp.yinhangjiancheng = o[i, 6].ToString().Trim();


                        temp.yinhangkemu = "";
                        if (o[i, 7] != null)
                            temp.yinhangkemu = o[i, 7].ToString().Trim();

                        temp.shoukuanfangshi = "";
                        if (o[i, 8] != null)
                            temp.shoukuanfangshi = o[i, 8].ToString().Trim();


                        temp.duiyinglie = "";
                        if (o[i, 9] != null)
                            temp.duiyinglie = o[i, 9].ToString().Trim();

                        temp.keytext = "";
                        if (o[i, 10] != null)
                            temp.keytext = o[i, 10].ToString().Trim();

                        temp.keytext11 = "";
                        if (o[i, 11] != null)
                            temp.keytext11 = o[i, 11].ToString().Trim();

                        temp.keytext12 = "";
                        if (o[i, 12] != null)
                            temp.keytext12 = o[i, 12].ToString().Trim();



                        temp.Message = "2.13";

                        #endregion
                        MAPPINGResult.Add(temp);

                        var ex = NAMEAlist.Find(v => v == temp.fuzeren);
                        if (ex == null || ex == "")
                        {
                            NAMEAlist.Add(temp.fuzeren);

                        }
                        else
                        {

                        }
                    }

                    #region 2
                    List<clsKEYinfo> MAPPINGResult2 = new List<clsKEYinfo>();
                    WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["2.14"];

                    rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 16]];
                    rowCount = WS.UsedRange.Rows.Count;
                    o = new object[1, 1];
                    o = (object[,])rng.Value2;
                    wscount = analyWK.Worksheets.Count;

                    for (int i = 2; i <= rowCount; i++)
                    {
                        clsKEYinfo temp = new clsKEYinfo();

                        #region 基础信息

                        temp.diqu = "";
                        if (o[i, 1] != null)
                            temp.diqu = o[i, 1].ToString().Trim();

                        temp.fuzeren = "";
                        if (o[i, 2] != null)
                            temp.fuzeren = o[i, 2].ToString().Trim();
                        if (temp.fuzeren == null || temp.fuzeren == "")
                            continue;

                        temp.maichangname = "";
                        if (o[i, 3] != null)
                            temp.maichangname = o[i, 3].ToString().Trim();

                        //卖场代码

                        temp.Maichangdaima = "";
                        if (o[i, 4] != null)
                            temp.Maichangdaima = o[i, 4].ToString().Trim();

                        temp.naishuidanwei = "";
                        if (o[i, 5] != null)
                            temp.naishuidanwei = o[i, 5].ToString().Trim();


                        //6-12

                        temp.yinhangjiancheng = "";
                        if (o[i, 6] != null)
                            temp.yinhangjiancheng = o[i, 6].ToString().Trim();


                        temp.yinhangkemu = "";
                        if (o[i, 7] != null)
                            temp.yinhangkemu = o[i, 7].ToString().Trim();

                        temp.shoukuanfangshi = "";
                        if (o[i, 8] != null)
                            temp.shoukuanfangshi = o[i, 8].ToString().Trim();


                        temp.duiyinglie = "";
                        if (o[i, 9] != null)
                            temp.duiyinglie = o[i, 9].ToString().Trim();

                        temp.keytext = "";
                        if (o[i, 10] != null)
                            temp.keytext = o[i, 10].ToString().Trim();

                        temp.keytext11 = "";
                        if (o[i, 11] != null)
                            temp.keytext11 = o[i, 11].ToString().Trim();

                        temp.keytext12 = "";
                        if (o[i, 12] != null)
                            temp.keytext12 = o[i, 12].ToString().Trim();

                        temp.Message = "2.14";


                        #endregion
                        MAPPINGResult2.Add(temp);

                        var ex = NAMEAlist.Find(v => v == temp.fuzeren);
                        if (ex == null || ex == "")
                        {
                            NAMEAlist.Add(temp.fuzeren);

                        }
                    }

                    #endregion
                    #region 3
                    List<clsKEYinfo> MAPPINGResult3 = new List<clsKEYinfo>();
                    WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["2.15"];

                    rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 16]];
                    rowCount = WS.UsedRange.Rows.Count;
                    o = new object[1, 1];
                    o = (object[,])rng.Value2;
                    wscount = analyWK.Worksheets.Count;

                    for (int i = 2; i <= rowCount; i++)
                    {
                        clsKEYinfo temp = new clsKEYinfo();

                        #region 基础信息

                        temp.diqu = "";
                        if (o[i, 1] != null)
                            temp.diqu = o[i, 1].ToString().Trim();

                        temp.fuzeren = "";
                        if (o[i, 2] != null)
                            temp.fuzeren = o[i, 2].ToString().Trim();

                        if (temp.fuzeren == null || temp.fuzeren == "")
                            continue;
                        temp.maichangname = "";
                        if (o[i, 3] != null)
                            temp.maichangname = o[i, 3].ToString().Trim();

                        //卖场代码

                        temp.Maichangdaima = "";
                        if (o[i, 4] != null)
                            temp.Maichangdaima = o[i, 4].ToString().Trim();

                        temp.naishuidanwei = "";
                        if (o[i, 5] != null)
                            temp.naishuidanwei = o[i, 5].ToString().Trim();


                        //6-12

                        temp.yinhangjiancheng = "";
                        if (o[i, 6] != null)
                            temp.yinhangjiancheng = o[i, 6].ToString().Trim();


                        temp.yinhangkemu = "";
                        if (o[i, 7] != null)
                            temp.yinhangkemu = o[i, 7].ToString().Trim();

                        temp.shoukuanfangshi = "";
                        if (o[i, 8] != null)
                            temp.shoukuanfangshi = o[i, 8].ToString().Trim();


                        temp.duiyinglie = "";
                        if (o[i, 9] != null)
                            temp.duiyinglie = o[i, 9].ToString().Trim();

                        temp.keytext = "";
                        if (o[i, 10] != null)
                            temp.keytext = o[i, 10].ToString().Trim();

                        temp.keytext11 = "";
                        if (o[i, 11] != null)
                            temp.keytext11 = o[i, 11].ToString().Trim();

                        temp.keytext12 = "";
                        if (o[i, 12] != null)
                            temp.keytext12 = o[i, 12].ToString().Trim();


                        temp.Message = "2.15";

                        #endregion
                        MAPPINGResult3.Add(temp);

                        var ex = NAMEAlist.Find(v => v == temp.fuzeren);
                        if (ex == null || ex == "")
                        {
                            NAMEAlist.Add(temp.fuzeren);

                        }
                    }

                    #endregion
                    #region 4
                    List<clsKEYinfo> MAPPINGResult4 = new List<clsKEYinfo>();
                    WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["2.18"];

                    rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 16]];
                    rowCount = WS.UsedRange.Rows.Count;
                    o = new object[1, 1];
                    o = (object[,])rng.Value2;
                    wscount = analyWK.Worksheets.Count;

                    for (int i = 2; i <= rowCount; i++)
                    {
                        clsKEYinfo temp = new clsKEYinfo();

                        #region 基础信息

                        temp.diqu = "";
                        if (o[i, 1] != null)
                            temp.diqu = o[i, 1].ToString().Trim();

                        temp.fuzeren = "";
                        if (o[i, 2] != null)
                            temp.fuzeren = o[i, 2].ToString().Trim();

                        if (temp.fuzeren == null || temp.fuzeren == "")
                            continue;
                        temp.maichangname = "";
                        if (o[i, 3] != null)
                            temp.maichangname = o[i, 3].ToString().Trim();

                        //卖场代码

                        temp.Maichangdaima = "";
                        if (o[i, 4] != null)
                            temp.Maichangdaima = o[i, 4].ToString().Trim();

                        temp.naishuidanwei = "";
                        if (o[i, 5] != null)
                            temp.naishuidanwei = o[i, 5].ToString().Trim();


                        //6-12

                        temp.yinhangjiancheng = "";
                        if (o[i, 6] != null)
                            temp.yinhangjiancheng = o[i, 6].ToString().Trim();


                        temp.yinhangkemu = "";
                        if (o[i, 7] != null)
                            temp.yinhangkemu = o[i, 7].ToString().Trim();

                        temp.shoukuanfangshi = "";
                        if (o[i, 8] != null)
                            temp.shoukuanfangshi = o[i, 8].ToString().Trim();


                        temp.duiyinglie = "";
                        if (o[i, 9] != null)
                            temp.duiyinglie = o[i, 9].ToString().Trim();

                        temp.keytext = "";
                        if (o[i, 10] != null)
                            temp.keytext = o[i, 10].ToString().Trim();

                        temp.keytext11 = "";
                        if (o[i, 11] != null)
                            temp.keytext11 = o[i, 11].ToString().Trim();

                        temp.keytext12 = "";
                        if (o[i, 12] != null)
                            temp.keytext12 = o[i, 12].ToString().Trim();


                        temp.Message = "2.18";

                        #endregion
                        MAPPINGResult4.Add(temp);

                        var ex = NAMEAlist.Find(v => v == temp.fuzeren);
                        if (ex == null || ex == "")
                        {
                            NAMEAlist.Add(temp.fuzeren);

                        }
                    }

                    #endregion
                    #region 5
                    List<clsKEYinfo> MAPPINGResult5 = new List<clsKEYinfo>();
                    WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["2.19"];

                    rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 16]];
                    rowCount = WS.UsedRange.Rows.Count;
                    o = new object[1, 1];
                    o = (object[,])rng.Value2;
                    wscount = analyWK.Worksheets.Count;

                    for (int i = 2; i <= rowCount; i++)
                    {
                        clsKEYinfo temp = new clsKEYinfo();

                        #region 基础信息

                        temp.diqu = "";
                        if (o[i, 1] != null)
                            temp.diqu = o[i, 1].ToString().Trim();

                        temp.fuzeren = "";
                        if (o[i, 2] != null)
                            temp.fuzeren = o[i, 2].ToString().Trim();

                        if (temp.fuzeren == null || temp.fuzeren == "")
                            continue;
                        temp.maichangname = "";
                        if (o[i, 3] != null)
                            temp.maichangname = o[i, 3].ToString().Trim();

                        //卖场代码

                        temp.Maichangdaima = "";
                        if (o[i, 4] != null)
                            temp.Maichangdaima = o[i, 4].ToString().Trim();

                        temp.naishuidanwei = "";
                        if (o[i, 5] != null)
                            temp.naishuidanwei = o[i, 5].ToString().Trim();


                        //6-12

                        temp.yinhangjiancheng = "";
                        if (o[i, 6] != null)
                            temp.yinhangjiancheng = o[i, 6].ToString().Trim();


                        temp.yinhangkemu = "";
                        if (o[i, 7] != null)
                            temp.yinhangkemu = o[i, 7].ToString().Trim();

                        temp.shoukuanfangshi = "";
                        if (o[i, 8] != null)
                            temp.shoukuanfangshi = o[i, 8].ToString().Trim();


                        temp.duiyinglie = "";
                        if (o[i, 9] != null)
                            temp.duiyinglie = o[i, 9].ToString().Trim();

                        temp.keytext = "";
                        if (o[i, 10] != null)
                            temp.keytext = o[i, 10].ToString().Trim();

                        temp.keytext11 = "";
                        if (o[i, 11] != null)
                            temp.keytext11 = o[i, 11].ToString().Trim();

                        temp.keytext12 = "";
                        if (o[i, 12] != null)
                            temp.keytext12 = o[i, 12].ToString().Trim();


                        temp.Message = "2.19";

                        #endregion
                        MAPPINGResult5.Add(temp);
                        var ex = NAMEAlist.Find(v => v == temp.fuzeren);
                        if (ex == null || ex == "")
                        {
                            NAMEAlist.Add(temp.fuzeren);

                        }
                    }

                    #endregion
                    #region 6
                    List<clsKEYinfo> MAPPINGResult6 = new List<clsKEYinfo>();
                    WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["2.20"];

                    rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 16]];
                    rowCount = WS.UsedRange.Rows.Count;
                    o = new object[1, 1];
                    o = (object[,])rng.Value2;
                    wscount = analyWK.Worksheets.Count;

                    for (int i = 2; i <= rowCount; i++)
                    {
                        clsKEYinfo temp = new clsKEYinfo();

                        #region 基础信息

                        temp.diqu = "";
                        if (o[i, 1] != null)
                            temp.diqu = o[i, 1].ToString().Trim();

                        temp.fuzeren = "";
                        if (o[i, 2] != null)
                            temp.fuzeren = o[i, 2].ToString().Trim();
                        if (temp.fuzeren == null || temp.fuzeren == "")
                            continue;

                        temp.maichangname = "";
                        if (o[i, 3] != null)
                            temp.maichangname = o[i, 3].ToString().Trim();

                        //卖场代码

                        temp.Maichangdaima = "";
                        if (o[i, 4] != null)
                            temp.Maichangdaima = o[i, 4].ToString().Trim();

                        temp.naishuidanwei = "";
                        if (o[i, 5] != null)
                            temp.naishuidanwei = o[i, 5].ToString().Trim();


                        //6-12

                        temp.yinhangjiancheng = "";
                        if (o[i, 6] != null)
                            temp.yinhangjiancheng = o[i, 6].ToString().Trim();


                        temp.yinhangkemu = "";
                        if (o[i, 7] != null)
                            temp.yinhangkemu = o[i, 7].ToString().Trim();

                        temp.shoukuanfangshi = "";
                        if (o[i, 8] != null)
                            temp.shoukuanfangshi = o[i, 8].ToString().Trim();


                        temp.duiyinglie = "";
                        if (o[i, 9] != null)
                            temp.duiyinglie = o[i, 9].ToString().Trim();

                        temp.keytext = "";
                        if (o[i, 10] != null)
                            temp.keytext = o[i, 10].ToString().Trim();

                        temp.keytext11 = "";
                        if (o[i, 11] != null)
                            temp.keytext11 = o[i, 11].ToString().Trim();

                        temp.keytext12 = "";
                        if (o[i, 12] != null)
                            temp.keytext12 = o[i, 12].ToString().Trim();

                        temp.Message = "2.20";


                        #endregion
                        MAPPINGResult6.Add(temp);
                        var ex = NAMEAlist.Find(v => v == temp.fuzeren);
                        if (ex == null || ex == "")
                        {
                            NAMEAlist.Add(temp.fuzeren);

                        }
                    }

                    #endregion
                    #region 7
                    List<clsKEYinfo> MAPPINGResult7 = new List<clsKEYinfo>();
                    WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["2.21"];

                    rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 16]];
                    rowCount = WS.UsedRange.Rows.Count;
                    o = new object[1, 1];
                    o = (object[,])rng.Value2;
                    wscount = analyWK.Worksheets.Count;

                    for (int i = 2; i <= rowCount; i++)
                    {
                        clsKEYinfo temp = new clsKEYinfo();

                        #region 基础信息

                        temp.diqu = "";
                        if (o[i, 1] != null)
                            temp.diqu = o[i, 1].ToString().Trim();

                        temp.fuzeren = "";
                        if (o[i, 2] != null)
                            temp.fuzeren = o[i, 2].ToString().Trim();

                        if (temp.fuzeren == null || temp.fuzeren == "")
                            continue;
                        temp.maichangname = "";
                        if (o[i, 3] != null)
                            temp.maichangname = o[i, 3].ToString().Trim();

                        //卖场代码

                        temp.Maichangdaima = "";
                        if (o[i, 4] != null)
                            temp.Maichangdaima = o[i, 4].ToString().Trim();

                        temp.naishuidanwei = "";
                        if (o[i, 5] != null)
                            temp.naishuidanwei = o[i, 5].ToString().Trim();


                        //6-12

                        temp.yinhangjiancheng = "";
                        if (o[i, 6] != null)
                            temp.yinhangjiancheng = o[i, 6].ToString().Trim();


                        temp.yinhangkemu = "";
                        if (o[i, 7] != null)
                            temp.yinhangkemu = o[i, 7].ToString().Trim();

                        temp.shoukuanfangshi = "";
                        if (o[i, 8] != null)
                            temp.shoukuanfangshi = o[i, 8].ToString().Trim();


                        temp.duiyinglie = "";
                        if (o[i, 9] != null)
                            temp.duiyinglie = o[i, 9].ToString().Trim();

                        temp.keytext = "";
                        if (o[i, 10] != null)
                            temp.keytext = o[i, 10].ToString().Trim();

                        temp.keytext11 = "";
                        if (o[i, 11] != null)
                            temp.keytext11 = o[i, 11].ToString().Trim();

                        temp.keytext12 = "";
                        if (o[i, 12] != null)
                            temp.keytext12 = o[i, 12].ToString().Trim();


                        temp.Message = "2.21";

                        #endregion
                        MAPPINGResult7.Add(temp);
                        var ex = NAMEAlist.Find(v => v == temp.fuzeren);
                        if (ex == null || ex == "")
                        {
                            NAMEAlist.Add(temp.fuzeren);

                        }
                    }

                    #endregion
                    #region 8
                    List<clsKEYinfo> MAPPINGResult8 = new List<clsKEYinfo>();
                    WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["2.22"];

                    rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 16]];
                    rowCount = WS.UsedRange.Rows.Count;
                    o = new object[1, 1];
                    o = (object[,])rng.Value2;
                    wscount = analyWK.Worksheets.Count;

                    for (int i = 2; i <= rowCount; i++)
                    {
                        clsKEYinfo temp = new clsKEYinfo();

                        #region 基础信息

                        temp.diqu = "";
                        if (o[i, 1] != null)
                            temp.diqu = o[i, 1].ToString().Trim();

                        temp.fuzeren = "";
                        if (o[i, 2] != null)
                            temp.fuzeren = o[i, 2].ToString().Trim();
                        if (temp.fuzeren == null || temp.fuzeren == "")
                            continue;

                        temp.maichangname = "";
                        if (o[i, 3] != null)
                            temp.maichangname = o[i, 3].ToString().Trim();

                        //卖场代码

                        temp.Maichangdaima = "";
                        if (o[i, 4] != null)
                            temp.Maichangdaima = o[i, 4].ToString().Trim();

                        temp.naishuidanwei = "";
                        if (o[i, 5] != null)
                            temp.naishuidanwei = o[i, 5].ToString().Trim();


                        //6-12

                        temp.yinhangjiancheng = "";
                        if (o[i, 6] != null)
                            temp.yinhangjiancheng = o[i, 6].ToString().Trim();


                        temp.yinhangkemu = "";
                        if (o[i, 7] != null)
                            temp.yinhangkemu = o[i, 7].ToString().Trim();

                        temp.shoukuanfangshi = "";
                        if (o[i, 8] != null)
                            temp.shoukuanfangshi = o[i, 8].ToString().Trim();


                        temp.duiyinglie = "";
                        if (o[i, 9] != null)
                            temp.duiyinglie = o[i, 9].ToString().Trim();

                        temp.keytext = "";
                        if (o[i, 10] != null)
                            temp.keytext = o[i, 10].ToString().Trim();

                        temp.keytext11 = "";
                        if (o[i, 11] != null)
                            temp.keytext11 = o[i, 11].ToString().Trim();

                        temp.keytext12 = "";
                        if (o[i, 12] != null)
                            temp.keytext12 = o[i, 12].ToString().Trim();


                        temp.Message = "2.22";

                        #endregion
                        MAPPINGResult8.Add(temp);
                        var ex = NAMEAlist.Find(v => v == temp.fuzeren);
                        if (ex == null || ex == "")
                        {
                            NAMEAlist.Add(temp.fuzeren);

                        }
                        else
                        {

                        }
                    }

                    #endregion

                    #region 读取模板列对应名称
                    List<clsR2RBankMappinginfo> R2RBankMapping = new List<clsR2RBankMappinginfo>();

                    WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["all"];

                    rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 16]];
                    rowCount = WS.UsedRange.Rows.Count;
                    o = new object[1, 1];
                    o = (object[,])rng.Value2;
                    wscount = analyWK.Worksheets.Count;

                    MAPPINGResult = MAPPINGResult.Concat(MAPPINGResult2).Concat(MAPPINGResult3).Concat(MAPPINGResult4).Concat(MAPPINGResult5).Concat(MAPPINGResult6).Concat(MAPPINGResult7).Concat(MAPPINGResult8).ToList();

                    int index = 2;

                    for (int i = 0; i < NAMEAlist.Count; i++)
                    {
                        clsKEYinfo temp = new clsKEYinfo();

                        List<clsKEYinfo> mlist = MAPPINGResult.FindAll(oe => oe.fuzeren != null && oe.fuzeren == NAMEAlist[i]).ToList();
                        if (mlist.Count > 0)
                        {
                            temp = mlist[0];

                            string tx1 = "";
                            string tx2 = "";
                            foreach (clsKEYinfo iten in mlist)
                            {
                                if (iten.maichangname != null && iten.maichangname != "")
                                    tx1 += " " + iten.Message + " " + iten.maichangname;
                                if (iten.keytext12 != null && iten.keytext12 != "")
                                    tx2 += " " + iten.Message + " " + iten.keytext12;


                            }
                            temp.maichangname = tx1;
                            temp.keytext12 = tx2;
                        }
                        WS.Cells[index, 2] = temp.fuzeren;
                        WS.Cells[index, 3] = temp.maichangname;
                        WS.Cells[index, 4] = temp.Maichangdaima;

                        WS.Cells[index, 5] = temp.naishuidanwei;
                        WS.Cells[index, 6] = temp.yinhangjiancheng;
                        WS.Cells[index, 7] = temp.yinhangkemu;
                        WS.Cells[index, 8] = temp.shoukuanfangshi;
                        WS.Cells[index, 9] = temp.duiyinglie;
                        WS.Cells[index, 10] = temp.keytext;
                        WS.Cells[index, 11] = temp.keytext11;
                        WS.Cells[index, 12] = temp.keytext12;

                        WS.Cells[index, 13] = mlist.Count;
                        index++;


                    }

                    #endregion
                    excelApp.Visible = true;
                    excelApp.ScreenUpdating = true;



                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: 01032" + ex);
                return null;

                throw;
            }
            return MAPPINGResult;

        }


        public List<clsKEYinfo> GetKEYnfo_列出客户所在的区或县()
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "Resources\\【名单】需列出客户所在的区或县-0614完成.xlsx";



            List<clsKEYinfo> MAPPINGResult = new List<clsKEYinfo>();
            try
            {
                List<clsKEYinfo> WANGYINResult = new List<clsKEYinfo>();
                System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                Microsoft.Office.Interop.Excel.Application excelApp;
                {

                    excelApp = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook analyWK = excelApp.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing,
                        "htc", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    MAPPINGResult = new List<clsKEYinfo>();
                    Microsoft.Office.Interop.Excel.Worksheet WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["区县参考表B"];
                    Microsoft.Office.Interop.Excel.Range rng;
                    rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 16]];
                    int rowCount = WS.UsedRange.Rows.Count;
                    object[,] o = new object[1, 1];
                    o = (object[,])rng.Value2;
                    int wscount = analyWK.Worksheets.Count;

                    for (int i = 2; i <= rowCount; i++)
                    {
                        clsKEYinfo temp = new clsKEYinfo();

                        #region 基础信息

                        temp.diqu = "";
                        if (o[i, 1] != null)
                            temp.diqu = o[i, 1].ToString().Trim();

                        temp.fuzeren = "";
                        if (o[i, 2] != null)
                            temp.fuzeren = o[i, 2].ToString().Trim();


                

                        #endregion
                        MAPPINGResult.Add(temp);
                    }

                    #region 读取模板列对应名称
                    List<clsR2RBankMappinginfo> R2RBankMapping = new List<clsR2RBankMappinginfo>();

                    WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["需完善的表A"];

                    rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 16]];
                    rowCount = WS.UsedRange.Rows.Count;
                    o = new object[1, 1];
                    o = (object[,])rng.Value2;
                    wscount = analyWK.Worksheets.Count;

                    for (int i = 2; i <= rowCount; i++)
                    {
                        clsR2RBankMappinginfo temp = new clsR2RBankMappinginfo();

                        #region 基础信息

                        temp.jiaoyiriqi = "";
                        if (o[i, 1] != null)
                            temp.jiaoyiriqi = o[i, 1].ToString().Trim();
                        if (temp.jiaoyiriqi == null || temp.jiaoyiriqi == "")
                            continue;

                        temp.jiefang = "";
                        if (o[i, 2] != null)
                            temp.jiefang = o[i, 2].ToString().Trim();


                        temp.yongtu = "";
                        if (o[i, 3] != null)
                            temp.yongtu = o[i, 3].ToString().Trim();

                        #endregion
                        if (temp.jiefang.Contains("重庆市云阳县电教站"))
                        { 
                        
                        }
                        List<clsKEYinfo> mlist = MAPPINGResult.FindAll(oe => oe.diqu != null && oe.diqu == temp.jiaoyiriqi && temp.jiefang.Contains(oe.fuzeren)).ToList();
                        if (mlist.Count > 0 && temp.yongtu.Length==0)
                        {

                            WS.Cells[i, 3] = mlist[0].fuzeren;

                        }
                        R2RBankMapping.Add(temp);
                    }

                    #endregion
                    excelApp.Visible = true;
                    excelApp.ScreenUpdating = true;
 
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: 01032" + ex);
                return null;

                throw;
            }
            return MAPPINGResult;

        }
      
    }
}
