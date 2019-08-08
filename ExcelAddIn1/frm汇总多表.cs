using clsBuiness;
using clsCommon;
using dblist;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelAddIn1
{
    public partial class frm汇总多表 : Form
    {
        private List<string> Alist = new List<string>();
        private string Copyfile = "";
        private string username;
        private string Password;
        // 后台执行控件
        private BackgroundWorker bgWorker;
        // 消息显示窗体
        private bakfrmMessageShow frmMessageShow;
        // 后台操作是否正常完成
        private bool blnBackGroundWorkIsOK = false;
        //后加的后台属性显
        private bool backGroundRunResult;
        public string path;

        List<clsFenbiaoInfo> FenbiaoResult;
        List<clsmoban_biaoInfo> mobanResult;


        public frm汇总多表()
        {
            InitializeComponent();


        }
        private void InitialBackGroundWorker()
        {
            bgWorker = new BackgroundWorker();
            bgWorker.WorkerReportsProgress = true;
            bgWorker.WorkerSupportsCancellation = true;
            bgWorker.RunWorkerCompleted +=
                new RunWorkerCompletedEventHandler(bgWorker_RunWorkerCompleted);
            bgWorker.ProgressChanged +=
                new ProgressChangedEventHandler(bgWorker_ProgressChanged);
        }

        private void bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                blnBackGroundWorkIsOK = false;
            }
            else if (e.Cancelled)
            {
                blnBackGroundWorkIsOK = true;
            }
            else
            {
                blnBackGroundWorkIsOK = true;
            }
        }

        private void bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (frmMessageShow != null && frmMessageShow.Visible == true)
            {
                //设置显示的消息
                frmMessageShow.setMessage(e.UserState.ToString());
                //设置显示的按钮文字
                if (e.ProgressPercentage == clsConstant.Thread_Progress_OK)
                {
                    frmMessageShow.setStatus(clsConstant.Dialog_Status_Enable);
                }
            }
        }

        private void openFileBtton_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog();
            dialog.Description = "请选择所在文件夹";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (string.IsNullOrEmpty(dialog.SelectedPath))
                {
                    MessageBox.Show(this, "文件夹路径不能为空", "提示");
                    return;
                }
                path = dialog.SelectedPath;
                pathTextBox.Text = dialog.SelectedPath;


            }
            else
                return;

            Alist = new List<string>();

            Alist = GetBy_CategoryReportFileName(path);
        }
        public List<string> GetBy_CategoryReportFileName(string dirPath)
        {

            List<string> FileNameList = new List<string>();
            ArrayList list = new ArrayList();

            if (Directory.Exists(dirPath))
            {
                list.AddRange(Directory.GetFiles(dirPath));
            }
            if (list.Count > 0)
            {
                foreach (object item in list)
                {
                    if (!item.ToString().Contains("~$"))
                        //FileNameList.Add(item.ToString().Replace(dirPath + "\\", ""));
                        FileNameList.Add(item.ToString());
                }
            }

            return FileNameList;
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
            if (Alist == null && Alist.Count < 1)
                return;

            if (Copyfile == null || Copyfile.Length < 1)
                return;

            clsAllnew BusinessHelp = new clsAllnew();
            FenbiaoResult = BusinessHelp.Buiness_Bankcharge(ref this.bgWorker, "A", Password, username, Alist, Copyfile);

            MessageBox.Show("处理结束！");

            return;
            try
            {
                InitialBackGroundWorker();
                bgWorker.DoWork += new DoWorkEventHandler(KEYFile);

                bgWorker.RunWorkerAsync();
                // 启动消息显示画面
                frmMessageShow = new bakfrmMessageShow(clsShowMessage.MSG_001,
                                                    clsShowMessage.MSG_007,
                                                    clsConstant.Dialog_Status_Disable);
                frmMessageShow.ShowDialog();
                // 数据读取成功后在画面显示
                if (blnBackGroundWorkIsOK)
                {
                    //string ZFCEPath = Path.Combine(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Results"), "");
                    System.Diagnostics.Process.Start("explorer.exe", Copyfile);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }


        
         

        }
        private void KEYFile(object sender, DoWorkEventArgs e)
        {
            FenbiaoResult = new List<clsFenbiaoInfo>();

            //初始化信息
            clsAllnew BusinessHelp = new clsAllnew();
            //BusinessHelp.pbStatus = pbStatus;
            //BusinessHelp.tsStatusLabel1 = toolStripLabel1;
            DateTime oldDate = DateTime.Now;
            FenbiaoResult = BusinessHelp.Buiness_Bankcharge(ref this.bgWorker, "A", Password, username, Alist, Copyfile);
            DateTime FinishTime = DateTime.Now;  //   
            TimeSpan s = DateTime.Now - oldDate;
            string timei = s.Minutes.ToString() + ":" + s.Seconds.ToString();
            string Showtime = clsShowMessage.MSG_029 + timei.ToString();
            bgWorker.ReportProgress(clsConstant.Thread_Progress_OK, clsShowMessage.MSG_009 + "\r\n" + Showtime);
        }


        public List<clsTitleinfo> GetKEYnfo(string path)
        {
            //string path = AppDomain.CurrentDomain.BaseDirectory + "Resources\\航宇星111 - 副本.xls";

            List<clsTitleinfo> MAPPINGResult = new List<clsTitleinfo>();
            try
            {
                List<clsTitleinfo> WANGYINResult = new List<clsTitleinfo>();
                System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                Microsoft.Office.Interop.Excel.Application excelApp;
                {

                    excelApp = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook analyWK = excelApp.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing,
                        "htc", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    MAPPINGResult = new List<clsTitleinfo>();
                    Microsoft.Office.Interop.Excel.Worksheet WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets[1];
                    Microsoft.Office.Interop.Excel.Range rng;
                    rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 16]];
                    int rowCount = WS.UsedRange.Rows.Count;
                    object[,] o = new object[1, 1];
                    o = (object[,])rng.Value2;
                    int wscount = analyWK.Worksheets.Count;
                    clsCommHelp.CloseExcel(excelApp, analyWK);

                    for (int i = 3; i <= rowCount; i++)
                    {
                        clsTitleinfo temp = new clsTitleinfo();

                        #region 基础信息

                        temp.huozhubiaohaoA = "";
                        if (o[i, 1] != null)
                            temp.huozhubiaohaoA = o[i, 1].ToString().Trim();

                        temp.huozhumingchengB = "";
                        if (o[i, 2] != null)
                            temp.huozhumingchengB = o[i, 2].ToString().Trim();
                        if (temp.huozhumingchengB == null || temp.huozhumingchengB == "")
                            continue;


                        temp.xiangmubianmaC = "";
                        if (o[i, 3] != null)
                            temp.xiangmubianmaC = o[i, 3].ToString().Trim();

                        //卖场代码

                        temp.xiangmumingchengD = "";
                        if (o[i, 4] != null)
                            temp.xiangmumingchengD = o[i, 4].ToString().Trim();

                        temp.danjiaE = "";
                        if (o[i, 5] != null)
                            temp.danjiaE = o[i, 5].ToString().Trim();


                        temp.chanshengriqiF = "";
                        if (o[i, 6] != null)
                            temp.chanshengriqiF = clsCommHelp.objToDateTime(o[i, 6]);// o[i, 6].ToString().Trim();


                        temp.jifeiliangG = "";
                        if (o[i, 7] != null)
                            temp.jifeiliangG = o[i, 7].ToString().Trim();


                        temp.jineH = "";
                        if (o[i, 8] != null)
                            temp.jineH = o[i, 8].ToString().Trim();


                        temp.youhuijineI = "";
                        if (o[i, 9] != null)
                            temp.youhuijineI = o[i, 9].ToString().Trim();


                        temp.kaishiriqiJ = "";
                        if (o[i, 10] != null)
                            temp.kaishiriqiJ = clsCommHelp.objToDateTime(o[i, 10]); //o[i, 10].ToString().Trim();

                        temp.jieshuriqiK = "";
                        if (o[i, 11] != null)
                            temp.jieshuriqiK = clsCommHelp.objToDateTime(o[i, 11]);// o[i, 11].ToString().Trim();



                        #endregion
                        MAPPINGResult.Add(temp);
                    }


                    //excelApp.Visible = true;
                    //excelApp.ScreenUpdating = true;

                    DownbankExcel(MAPPINGResult);


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

        private void DownbankExcel(List<clsTitleinfo> Results)
        {
            #region 获取模板路径
            System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            string fullPath = Path.Combine(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "System\\"), "yzyg.xls");
            SaveFileDialog sfdDownFile = new SaveFileDialog();
            sfdDownFile.OverwritePrompt = false;
            string DesktopPath = Convert.ToString(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            sfdDownFile.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*.xlsx";
            string file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Results\\");
            if (Results.Count > 1)
                sfdDownFile.FileName = Path.Combine(file, Results[0].huozhumingchengB + DateTime.Now.ToString("yyyyMMdd"));

            string strExcelFileName = string.Empty;

            #endregion

            #region 导出前校验模板信息
            if (string.IsNullOrEmpty(sfdDownFile.FileName))
            {
                MessageBox.Show("File name can't be empty, please confirm, thanks!");
                return;
            }
            if (!File.Exists(fullPath))
            {
                MessageBox.Show("Template file does not exist, please confirm, thanks!");
                return;
            }
            else
            {
                strExcelFileName = sfdDownFile.FileName + ".xlsx";
            }
            #endregion
            #region Excel 初始化

            Microsoft.Office.Interop.Excel.Application ExcelApp;
            ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            System.Reflection.Missing missingValue = System.Reflection.Missing.Value;
            //
            Microsoft.Office.Interop.Excel._Workbook ExcelBook =
            ExcelApp.Workbooks.Open(fullPath, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue);
            #endregion
            #region Sheet 初始化
            try
            {

                {
                    Microsoft.Office.Interop.Excel._Worksheet ExcelSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelBook.Worksheets[1];
                    //打开时是否显示Excel
                    ExcelApp.Visible = false;
                    ExcelApp.ScreenUpdating = false;

                    int dou = ExcelSheet.UsedRange.Rows.Count + 1;
                    string las = "BT" + dou.ToString();
                    bool issave = false;
                    Microsoft.Office.Interop.Excel.Range rng = ExcelSheet.get_Range("A4", las);
                    // rng.Delete();

            #endregion
                    //ExcelApp.Visible = true;
                    //ExcelApp.ScreenUpdating = true;
                    #region 填充数据
                    int RowIndex = 4;
                    int xuhao = 1;

                    List<string> nashui = new List<string>();
                    List<int> xuhaoindex = new List<int>();

                    List<string> farenvalue = (from v in Results select v.kaishiriqiJ).Distinct().ToList();
                    farenvalue.Sort();

                    for (int i = 0; i < farenvalue.Count; i++)
                    {

                        double total = 0;

                        List<clsTitleinfo> cloumnlist = Results.FindAll(s => s.kaishiriqiJ != null && s.kaishiriqiJ == farenvalue[i]);
                        {
                            foreach (clsTitleinfo item in cloumnlist)
                            {

                                //if (item.daifang == null || item.daifang == 0)
                                //    continue;
                                if (item.xiangmumingchengD != null && item.xiangmumingchengD.Contains("入库"))
                                {



                                    ExcelSheet.Cells[RowIndex, 3] = item.jifeiliangG;

                                    ExcelSheet.Cells[RowIndex, 4] = "33";
                                    ExcelSheet.Cells[RowIndex, 5] = item.jineH;
                                    total = Convert.ToDouble(item.jineH);

                                }
                                else if (item.xiangmumingchengD != null && item.xiangmumingchengD.Contains("仓储费"))
                                {
                                    ExcelSheet.Cells[RowIndex, 6] = item.jifeiliangG;
                                    ExcelSheet.Cells[RowIndex, 7] = "3.8";
                                    ExcelSheet.Cells[RowIndex, 8] = item.jineH;

                                    total += Convert.ToDouble(item.jineH);

                                }
                                else if (item.xiangmumingchengD != null && item.xiangmumingchengD.Contains("拆零"))
                                {
                                    ExcelSheet.Cells[RowIndex, 9] = item.jifeiliangG;
                                    ExcelSheet.Cells[RowIndex, 10] = "0.38";
                                    ExcelSheet.Cells[RowIndex, 11] = item.jineH;
                                    total += Convert.ToDouble(item.jineH);
                                }
                                else if (item.xiangmumingchengD != null && item.xiangmumingchengD.Contains("出库"))
                                {
                                    ExcelSheet.Cells[RowIndex, 12] = item.jifeiliangG;
                                    ExcelSheet.Cells[RowIndex, 13] = "33";
                                    ExcelSheet.Cells[RowIndex, 14] = item.jineH;
                                    total += Convert.ToDouble(item.jineH);
                                }
                                ExcelSheet.Cells[RowIndex, 2] = item.chanshengriqiF;

                                ExcelSheet.Cells[RowIndex, 1] = RowIndex - 3;
                                ExcelSheet.Cells[RowIndex, 15] = total;



                            }
                            RowIndex++;
                        }



                    }
                    ExcelSheet.Cells[RowIndex, 1] = RowIndex - 3;
                    ExcelSheet.Cells[RowIndex, 2] = "合计";
                    int l = RowIndex - 1;

                    ExcelSheet.Cells[RowIndex, 3] = "=SUM(C4:C" + l.ToString() + ")";

                    ExcelSheet.Cells[RowIndex, 5] = "=SUM(E4:E" + l.ToString() + ")";
                    ExcelSheet.Cells[RowIndex, 6] = "=SUM(F4:F" + l.ToString() + ")";
                    ExcelSheet.Cells[RowIndex, 8] = "=SUM(H4:H" + l.ToString() + ")";
                    ExcelSheet.Cells[RowIndex, 9] = "=SUM(I4:I" + l.ToString() + ")";
                    ExcelSheet.Cells[RowIndex, 11] = "=SUM(K4:K" + l.ToString() + ")";
                    ExcelSheet.Cells[RowIndex, 12] = "=SUM(L4:L" + l.ToString() + ")";
                    ExcelSheet.Cells[RowIndex, 14] = "=SUM(N4:N" + l.ToString() + ")";
                    ExcelSheet.Cells[RowIndex, 15] = "=SUM(O4:O" + l.ToString() + ")";


                    ExcelSheet.Cells[RowIndex + 3, 1] = "亚洲渔港冷链（大连）有限公司";
                    ExcelSheet.Cells[RowIndex + 3, 11] = "大连铁越集团有限公司沈阳城市物流配送分公司";


                    ExcelSheet.Cells[RowIndex + 6, 1] = "负责人：";
                    ExcelSheet.Cells[RowIndex + 6, 11] = "负责人：";


                    ExcelSheet.Cells[1, 2] = "亚 洲 渔 港" + DateTime.Now.ToString("yyyy") + " 年 " + DateTime.Now.ToString("MM") + " 月" + " 计 费 明 细";
                    ExcelBook.RefreshAll();
                    #region 写入文件
                    sfdDownFile.FileName = Path.Combine(file, Results[0].huozhumingchengB + " " + DateTime.Now.ToString("yyyyMMdd-ss"));
                    strExcelFileName = sfdDownFile.FileName + ".xls";


                    ExcelApp.ScreenUpdating = true;
                    ExcelBook.SaveAs(strExcelFileName, missingValue, missingValue, missingValue, missingValue, missingValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, missingValue, missingValue, missingValue, missingValue, missingValue);
                    ExcelApp.DisplayAlerts = false;

                    #endregion
                }
            }

            #region 异常处理
            catch (Exception ex)
            {
                ExcelApp.DisplayAlerts = false;
                ExcelApp.Quit();
                ExcelBook = null;
                ExcelApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                throw ex;
            }
            #endregion

            #region Finally垃圾回收
            finally
            {
                ExcelBook.Close(false, missingValue, missingValue);
                ExcelBook = null;
                ExcelApp.DisplayAlerts = true;
                ExcelApp.Quit();
                clsKeyMyExcelProcess.Kill(ExcelApp);

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            #endregion

                    #endregion
        }


    }
}
