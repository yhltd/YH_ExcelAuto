using clsBuiness;
using clsCommon;
using dblist;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelAutoMAIL
{
    public partial class frm同一单元格加逗号 : Form
    {
        private string Copyfile = "";

        public frm同一单元格加逗号()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog tbox = new OpenFileDialog();
            tbox.Multiselect = false;
            tbox.Filter = "Excel Files(*.xls,*.xlsx,*.xlsm,*.xlsb)|*.xls;*.xlsx;*.xlsm;*.xlsb";
            if (tbox.ShowDialog() == DialogResult.OK)
            {
                Copyfile = tbox.FileName;
                textBox1.Text = Copyfile;
            }




        }

        private void importButton_Click(object sender, EventArgs e)
        {
            if (Copyfile != null && Copyfile != "" && Copyfile.Length > 0)
            {
            }
            else
            {
                return;
            }
            clsAllnew BusinessHelp = new clsAllnew();
            GetKEYnfo(Copyfile);

            MessageBox.Show("处理结束！");

            return;
        }
        public List<clstongyidanyuangehebing> GetKEYnfo(string path)
        {
            System.Reflection.Missing missingValue = System.Reflection.Missing.Value;
            List<clstongyidanyuangehebing> MAPPINGResult = new List<clstongyidanyuangehebing>();
            try
            {
                List<clstongyidanyuangehebing> WANGYINResult = new List<clstongyidanyuangehebing>();
                System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                Microsoft.Office.Interop.Excel.Application excelApp;
                {

                    excelApp = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook analyWK = excelApp.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing,
                        "htc", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    MAPPINGResult = new List<clstongyidanyuangehebing>();
                    Microsoft.Office.Interop.Excel.Worksheet WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["模板"];
                    Microsoft.Office.Interop.Excel.Range rng;
                    rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 16]];
                    int rowCount = WS.UsedRange.Rows.Count;
                    object[,] o = new object[1, 1];
                    o = (object[,])rng.Value2;
                    int wscount = analyWK.Worksheets.Count;


                    for (int i = 2; i <= rowCount; i++)
                    {
                        clstongyidanyuangehebing temp = new clstongyidanyuangehebing();
                        if (i == 15)
                        {



                        }

                        #region 基础信息

                        temp.A_lie = "";
                        if (o[i, 1] != null)
                            temp.A_lie = o[i, 1].ToString().Trim();
                        if (temp.A_lie == null || temp.A_lie == "")
                            continue;
                        //if (temp.A_lie.Contains(" ") || temp.A_lie.Contains("/") || temp.A_lie.Contains("-") || temp.A_lie.Contains("-") || temp.A_lie.Contains("、"))
                        {
                            //string[] fileText = System.Text.RegularExpressions.Regex.Split(temp.A_lie, ",");
                            temp.B_lie = temp.A_lie.Replace(" ", ",").Replace("/", ",").Replace("-", ",").Replace("、", ",").Replace("\r\n", ",").Replace("/", ",").Replace("，", ",").Replace("；", ",");

                            temp.B_lie = removeblank(temp.B_lie);                            //new 
                            temp.B_lie = removeblank_txt(temp.B_lie);
                            WS.Cells[i, 2] = temp.B_lie;

                        }
                        //else
                        //{
                        //    temp.A_lie = removeblank(temp.A_lie);
                        //    //new 
                        //    temp.A_lie = removeblank_txt(temp.A_lie);

                        //    WS.Cells[i, 2] = temp.A_lie.Trim();

                        //}
                        //temp.B_lie = "";
                        //if (o[i, 2] != null)
                        //    temp.B_lie = o[i, 2].ToString().Trim();

                        #endregion

                        MAPPINGResult.Add(temp);
                    }
                    #region 写入文件
                    string strDesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                    string filename = strDesktopPath + "\\Export  " + DateTime.Now.ToString("yyyyMMdd-ss") + ".xlsx";

                    excelApp.ScreenUpdating = true;
                    analyWK.SaveAs(filename, missingValue, missingValue, missingValue, missingValue, missingValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, missingValue, missingValue, missingValue, missingValue, missingValue);
                    excelApp.DisplayAlerts = false;
                    #endregion
                    //excelApp.Visible = true;
                    //excelApp.ScreenUpdating = true;
                    clsCommHelp.CloseExcel(excelApp, analyWK);
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
        private static string removeblank_txt(string sp_txt)
        {
            sp_txt = sp_txt.Trim().Replace(" ", "\t").Replace("\t\t", "\t");

            while (true)
            {
                if (sp_txt.Contains("  "))
                {
                    sp_txt = sp_txt.Replace("  ", " ");

                }
                else
                    break;

            }
            while (true)
            {
                if (sp_txt.Contains("\t\t"))
                {
                    sp_txt = sp_txt.Replace("\t\t", "\t");

                }
                else
                    break;

            }
            return sp_txt;
        }
        private static string removeblank(string sp_txt)
        {


            while (true)
            {
                if (sp_txt.Contains("  "))
                {
                    sp_txt = sp_txt.Replace("  ", " ");

                }
                else
                    break;

            }
            while (true)
            {
                if (sp_txt.Contains("\t\t"))
                {
                    sp_txt = sp_txt.Replace("\t\t", "\t");

                }
                else
                    break;

            }
            return sp_txt;
        }
    }
}
