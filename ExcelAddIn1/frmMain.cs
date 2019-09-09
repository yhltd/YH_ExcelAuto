using clsBuiness;
using clsCommon;
using dblist;
using newclscommon;
using Spire.Pdf;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

using System.Drawing.Imaging;
using System.IO;

namespace ExcelAddIn1
{
    public partial class frmMain : Form
    {
        // 后台执行控件
        private BackgroundWorker bgWorker;
        // 消息显示窗体
        private frmMessageShow frmMessageShow;
        // 后台操作是否正常完成
        private bool blnBackGroundWorkIsOK = false;
        //后加的后台属性显
        private bool backGroundRunResult;
        string Copyfile;
        private Thread GetDataforRawDataThread;
        private System.Timers.Timer timerAlter_new;
        List<clsSendmailinfo> MAPPINGResult;
        private bool IsRun = false;

        public frmMain()
        {
            InitializeComponent();
            NewMethod();


            string testvalue = "警告：由于客户未付清费用当前系统为测试系统，禁止转包模仿 破解等商业用途，如违反将追究相关法律责任";

            var form = new Login(testvalue);

            if (form.ShowDialog() == DialogResult.OK)
            {


            }
            else
                System.Environment.Exit(0);
        }
        private void NewMethod()
        {
            timerAlter_new = new System.Timers.Timer(6666);
            timerAlter_new.Elapsed += new System.Timers.ElapsedEventHandler(TimeControl);
            timerAlter_new.AutoReset = true;
            timerAlter_new.Start();
        }
        private void TimeControl(object sender, EventArgs e)
        {
            if (!IsRun)
            {
                IsRun = true;
                GetDataforRawDataThread = new Thread(TimeMethod);
                GetDataforRawDataThread.Start();
            }
        }
        private void TimeMethod()
        {
            bool istrue = true;
            clsmytest buiness = new clsmytest();

            bool istue = buiness.checkname("YH_ExcelAuto", "yhltd");
            if (istue == false)
            {
                Control.CheckForIllegalCrossThreadCalls = false;
                this.Visible = false;
                //MessageBox.Show("缺失系统文件，或电脑系统更新导致，请联系开发人员 !", "系统错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var form = new frmAlterinfo("缺失系统文件，或电脑系统更新导致，请联系开发人员 !");

                if (form.ShowDialog() == DialogResult.OK)
                {

                }


                System.Environment.Exit(0);
            }

            IsRun = false;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("找不到SAP主窗口！");
        }

        private void 汇总多表ToolStripMenuItem_Click(object sender, EventArgs e)
        {

            var form = new frm汇总多表();
            if (form.ShowDialog() == DialogResult.OK)
            {

            }
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            var form = new frmSendpage();
            if (form.ShowDialog() == DialogResult.OK)
            {

            }

        }

        private void 多邮箱ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void 导入数据模板ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog tbox = new OpenFileDialog();
            tbox.Multiselect = false;
            tbox.Filter = "Excel Files(*.xls,*.xlsx,*.xlsm,*.xlsb)|*.xls;*.xlsx;*.xlsm;*.xlsb";
            if (tbox.ShowDialog() == DialogResult.OK)
            {
                Copyfile = tbox.FileName;


            }
            MAPPINGResult = new List<clsSendmailinfo>();

            MAPPINGResult = GetKEYnfo(Copyfile);

            dataGridView2.AutoGenerateColumns = false;

            dataGridView2.DataSource = MAPPINGResult;




        }
        public List<clsSendmailinfo> GetKEYnfo(string path)
        {
            //string path = AppDomain.CurrentDomain.BaseDirectory + "Resources\\航宇星111 - 副本.xls";

            List<clsSendmailinfo> MAPPINGResult = new List<clsSendmailinfo>();
            try
            {
                List<clsSendmailinfo> WANGYINResult = new List<clsSendmailinfo>();
                System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                Microsoft.Office.Interop.Excel.Application excelApp;
                {

                    excelApp = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook analyWK = excelApp.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing,
                        "htc", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    MAPPINGResult = new List<clsSendmailinfo>();
                    Microsoft.Office.Interop.Excel.Worksheet WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["群发多邮箱"];
                    Microsoft.Office.Interop.Excel.Range rng;
                    rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 16]];
                    int rowCount = WS.UsedRange.Rows.Count;
                    object[,] o = new object[1, 1];
                    o = (object[,])rng.Value2;
                    int wscount = analyWK.Worksheets.Count;
                    clsCommHelp.CloseExcel(excelApp, analyWK);

                    for (int i = 2; i <= rowCount; i++)
                    {
                        clsSendmailinfo temp = new clsSendmailinfo();

                        #region 基础信息
                        //
                        temp.sendfrom = "";
                        if (o[i, 1] != null)
                            temp.sendfrom = o[i, 1].ToString().Trim();

                        temp.sendto = "";
                        if (o[i, 2] != null)
                            temp.sendto = o[i, 2].ToString().Trim();
                        if (temp.sendto == null || temp.sendto == "")
                            continue;


                        temp.subject = "";
                        if (o[i, 3] != null)
                            temp.subject = o[i, 3].ToString().Trim();

                        //卖场代码

                        temp.bodyinfo = "";
                        if (o[i, 4] != null)
                            temp.bodyinfo = o[i, 4].ToString().Trim();

                        temp.acc = "";
                        if (o[i, 5] != null)
                            temp.acc = o[i, 5].ToString().Trim();


                        temp.msg_tel = "";
                        if (o[i, 6] != null)
                            temp.msg_tel = o[i, 6].ToString().Trim();


                        temp.host = "";
                        if (o[i, 7] != null)
                            temp.host = o[i, 7].ToString().Trim();


                        temp.password = "";
                        if (o[i, 8] != null)
                            temp.password = o[i, 8].ToString().Trim();


                        temp.bushiyong = "";
                        if (o[i, 9] != null)
                            temp.bushiyong = o[i, 9].ToString().Trim();


                        temp.jian_ge = "";
                        if (o[i, 11] != null)
                            temp.jian_ge = o[i, 11].ToString().Trim();


                        #endregion

                        MAPPINGResult.Add(temp);
                    }


                    //excelApp.Visible = true;
                    //excelApp.ScreenUpdating = true;

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

        private void 开始批量发信ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MAPPINGResult == null || MAPPINGResult.Count < 1)
            {

                MessageBox.Show("没有找到发信的信息，请先导入发信模板内容！");

                return;

            }

            //初始化信息
            clsAllnew BusinessHelp = new clsAllnew();
            int index = 1;

            toolStripLabel2.Text = "正在发送  :   " + index.ToString() + "/" + MAPPINGResult.Count.ToString();

            foreach (clsSendmailinfo item in MAPPINGResult)
            {
                if (item.bushiyong != null && item.bushiyong.Length >= 1)
                    continue;

                //bgWorker.ReportProgress(0, "已发送  :  " + index.ToString() + "/" + MAPPINGResult.ToString());

                toolStripLabel2.Text = "正在发送  :   " + index.ToString() + "/" + MAPPINGResult.Count.ToString();


                //添加模板
                if (item.msg_tel != null)
                {
                    string A_Path = Path.GetDirectoryName(Copyfile) + "\\" + item.subject + ".txt";
                    if (File.Exists(A_Path))
                    {
                        string[] fileTextnew = File.ReadAllLines(A_Path);
                        string htmlbody = "";

                        for (int i = 0; i < fileTextnew.Length; i++)
                        {
                            htmlbody = htmlbody + fileTextnew[i];
                        }

                        item.bodyinfo = htmlbody;
                    }
                }
                if (item.jian_ge != null && item.jian_ge.Length >= 0)
                    Thread.Sleep(Convert.ToInt32(item.jian_ge)*1000);


                string[] fileText = System.Text.RegularExpressions.Regex.Split(item.acc, ",");
                if (item.msg_tel == null || item.msg_tel.Length == 0)

                    BusinessHelp.SendMail_Allport(item.host, item.sendfrom, item.password, item.sendto, item.subject, item.bodyinfo, fileText, item.msg_tel);
                else

                    BusinessHelp.outllook_moban_Send(item.host, item.sendfrom, item.password, item.sendto, item.subject, item.bodyinfo, fileText, item.msg_tel);

                index++;
            }
            MessageBox.Show("运行结束，已发送邮件：  " + (index - 1).ToString());


            return;

            //add in  不支持多线程


            try
            {

                InitialBackGroundWorker();
                bgWorker.DoWork += new DoWorkEventHandler(BSendMail);

                bgWorker.RunWorkerAsync();

                // 启动消息显示画面
                frmMessageShow = new frmMessageShow(clsShowMessage.MSG_001,
                                                    clsShowMessage.MSG_007,
                                                    clsConstant.Dialog_Status_Disable);
                frmMessageShow.ShowDialog();

                // 数据读取成功后在画面显示
                if (blnBackGroundWorkIsOK)
                {

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);

                return;
                throw ex;



            }




        }
        private void BSendMail(object sender, DoWorkEventArgs e)
        {


            DateTime oldDate = DateTime.Now;

            //初始化信息
            clsAllnew BusinessHelp = new clsAllnew();
            int index = 0;

            foreach (clsSendmailinfo item in MAPPINGResult)
            {
                bgWorker.ReportProgress(0, "已发送  :  " + index.ToString() + "/" + MAPPINGResult.ToString());
                string[] fileText = System.Text.RegularExpressions.Regex.Split(item.acc, ",");
                BusinessHelp.SendMail_Allport(item.host, item.sendfrom, item.password, item.sendto, item.subject, item.bodyinfo, fileText, item.msg_tel);
                index++;
            }

            DateTime FinishTime = DateTime.Now;
            TimeSpan s = DateTime.Now - oldDate;
            string timei = s.Minutes.ToString() + ":" + s.Seconds.ToString();
            string Showtime = clsShowMessage.MSG_029 + timei.ToString();
            bgWorker.ReportProgress(clsConstant.Thread_Progress_OK, clsShowMessage.MSG_009 + "\r\n" + Showtime);

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

        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            selectPDF();


        }

        private void selectPDF()
        {
            OpenFileDialog tbox = new OpenFileDialog();
            tbox.Multiselect = false;
            tbox.Filter = "PDF Files(*.PDF,*.PDF,*.PDF,*.PDF)|*.PDF;*.PDF;*.PDF;*.PDF";
            if (tbox.ShowDialog() == DialogResult.OK)
            {
                Copyfile = tbox.FileName;
            }

            MessageBox.Show("导入成功，请点击开始转换  ");
        }

        private void toolStripMenuItem5_Click(object sender, EventArgs e)
        {
            toolStripLabel1.Text = "转换期间比较漫长请耐心等待.....";

            Thread t = new Thread(converpdf);
            t.Start();//线程开始执行


            //Thread.Sleep(TimeSpan.FromSeconds(2));
            //t.Abort();

            return;

            try
            {

                InitialBackGroundWorker();
                bgWorker.DoWork += new DoWorkEventHandler(cover_pdftoword);

                bgWorker.RunWorkerAsync();

                // 启动消息显示画面
                frmMessageShow = new frmMessageShow(clsShowMessage.MSG_001,
                                                    clsShowMessage.MSG_007,
                                                    clsConstant.Dialog_Status_Disable);
                frmMessageShow.ShowDialog();

                // 数据读取成功后在画面显示
                if (blnBackGroundWorkIsOK)
                {

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);

                return;
                throw ex;



            }


        }
        private void cover_pdftoword(object sender, DoWorkEventArgs e)
        {


            DateTime oldDate = DateTime.Now;


            converpdf();



            DateTime FinishTime = DateTime.Now;
            TimeSpan s = DateTime.Now - oldDate;
            string timei = s.Minutes.ToString() + ":" + s.Seconds.ToString();
            string Showtime = clsShowMessage.MSG_029 + timei.ToString();
            bgWorker.ReportProgress(clsConstant.Thread_Progress_OK, clsShowMessage.MSG_009 + "\r\n" + Showtime);

        }

        private void converpdf()
        {
            string strDesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string filename = strDesktopPath + "\\Export PDF " + DateTime.Now.ToString("yyyyMMdd-ss") + ".doc";

            #region Pdf转word
            PdfDocument doc = new PdfDocument();
            doc.LoadFromFile(Copyfile);
            doc.SaveToFile(filename, Spire.Pdf.FileFormat.DOC);

            toolStripLabel1.Text = "生成完成-" + filename;

            #endregion
        }


        private void toolStripMenuItem7_Click(object sender, EventArgs e)
        {
            OpenFileDialog tbox = new OpenFileDialog();
            tbox.Multiselect = false;
            tbox.Filter = "PDF Files(*.doc,*.docx,*.DOC,*.DOCX)|*.doc;*.docx;*.DOC;*.DOCX";
            if (tbox.ShowDialog() == DialogResult.OK)
            {
                Copyfile = tbox.FileName;
            }

            MessageBox.Show("导入成功，请点击开始转换  ");

        }

        private void toolStripMenuItem8_Click(object sender, EventArgs e)
        {

            string strDesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string filename = strDesktopPath + "\\Export image " + DateTime.Now.ToString("yyyyMMdd-ss") + ".jpg";
            clsAllnew bu = new clsAllnew();


            bu.WordToJPGBySpire(Copyfile, filename);


        }

        private void 导入数据模板ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            selectPDF();

        }

        private void 开始转换ToolStripMenuItem_Click(object sender, EventArgs e)
        {

            try
            {

                labshow();

                Thread t = new Thread(pdf_jpg);
                t.Start();//线程开始执行


                return;

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);


                throw;
            }
        }

        private void pdf_jpg()
        {
            string strDesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string filename = strDesktopPath + "\\pic\\Export image " + DateTime.Now.ToString("yyyyMMdd-ss") + "";


            string subPath = strDesktopPath + "/pic/";
            if (false == System.IO.Directory.Exists(subPath))
            {
                //创建pic文件夹
                System.IO.Directory.CreateDirectory(subPath);
            }

            //初始化PdfDocument实例
            PdfDocument doc = new PdfDocument();

            //加载PDF文档
            doc.LoadFromFile(Copyfile);

            //遍历PDF每一页
            for (int i = 0; i < doc.Pages.Count; i++)
            {
                //将PDF页转换成bitmap图形
                System.Drawing.Image bmp = doc.SaveAsImage(i);

                //将bitmap图形保存为png图片
                string fileName = string.Format("Page-{0}.png", i + 1);
                toolStripLabel1.Text = "生成:" + (i + 1) + "/" + doc.Pages.Count;

                bmp.Save(filename + fileName, System.Drawing.Imaging.ImageFormat.Png);
            }

            toolStripLabel1.Text = "生成完成-" + subPath;
        }

        private void labshow()
        {
            toolStripLabel1.Text = "转换期间比较漫长请耐心等待.....";
        }

    }
}
