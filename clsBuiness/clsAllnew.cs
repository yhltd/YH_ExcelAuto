using clsCommon;
using dblist;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace clsBuiness
{
    public class clsAllnew
    {
        public BackgroundWorker bgWorker1;
        //private object missing = System.Reflection.Missing.Value;
        public ToolStripProgressBar pbStatus { get; set; }
        public ToolStripStatusLabel tsStatusLabel1 { get; set; }
        public BackgroundWorker backgroundWorker1;
        public List<clsFenbiaoInfo> FenbiaoInfo;
        private string fullPath;

        public List<clsFenbiaoInfo> Buiness_Bankcharge(ref BackgroundWorker bgWorker, string casetype, string Password, string USER, List<string> Alist, string fullPath1)
        {

            fullPath = fullPath1;

            bgWorker1 = bgWorker;
            try
            {

                #region 读取 本地日报所有信息表

                FenbiaoInfo = new List<clsFenbiaoInfo>();
                DownbankExcel(Alist);


                #endregion
                return FenbiaoInfo;

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: 0032" + ex);


                throw;
            }
        }

        private void DownbankExcel(List<string> Alist)
        {
            #region 获取模板路径
            System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            //string fullPath = Path.Combine(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "System\\"), "yzyg.xls");
            SaveFileDialog sfdDownFile = new SaveFileDialog();
            sfdDownFile.OverwritePrompt = false;
            string DesktopPath = Convert.ToString(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            sfdDownFile.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*.xlsx";
            string file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Results\\");
            if (Alist.Count > 1)
                sfdDownFile.FileName = Path.Combine(DesktopPath, "合并-" + DateTime.Now.ToString("yyyyMMdd"));

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
                strExcelFileName = sfdDownFile.FileName + ".xls";
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
                    //ExcelApp.Visible = true;
                    //ExcelApp.ScreenUpdating = true;
                    int lineinex = ExcelSheet.UsedRange.Rows.Count + 1;
                    int dou = ExcelSheet.UsedRange.Rows.Count + 1;
                    string las = "BT" + dou.ToString();
                    bool issave = false;
                    Microsoft.Office.Interop.Excel.Range rng = ExcelSheet.get_Range("A4", las);
                    // rng.Delete();

            #endregion

                    #region 填充数据
                    int RowIndex = 4;
                    int xuhao = 1;
                    ExcelApp.DisplayAlerts = false;
                    for (int i = 0; i < Alist.Count; i++)
                    {
                   
                        Microsoft.Office.Interop.Excel._Workbook ExcelBook2 =
                                 ExcelApp.Workbooks.Open(Alist[i], missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue);

                        Microsoft.Office.Interop.Excel._Worksheet ExcelSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ExcelBook2.Worksheets[1];

                        Excel.Range r = ExcelSheet2.Range[ExcelSheet2.Cells[1, 1], ExcelSheet2.Cells[ExcelSheet.UsedRange.Rows.Count, 135]];
                        r.Copy(missingValue);

                        int lineinex2 = ExcelSheet2.UsedRange.Rows.Count;
                    

                        Excel.Range r2 = ExcelSheet.Range[ExcelSheet.Cells[lineinex, 1], ExcelSheet.Cells[lineinex, 135]];
                        r2.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues);

                        lineinex = lineinex2 + lineinex;

                        ExcelBook2.Close();
                        ExcelBook2 = null;
                    }
                    ExcelBook.RefreshAll();
                    #region 写入文件
                    sfdDownFile.FileName = Path.Combine(DesktopPath,"合并-" + " " + DateTime.Now.ToString("yyyyMMdd-ss"));
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

        public void SendMail_Allport(string Hosti, string fromi, string passkey, string toi, string Subjecti, string Bodyi, string[] Attachmentlist)
        {

            {
                System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient();
                client.Host = Hosti;// "smtp.126.com";
                client.UseDefaultCredentials = false;
                //
                //启用功能修改处
                //
                client.Credentials = new System.Net.NetworkCredential(fromi, passkey);
                client.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                client.Port = 25;
                client.EnableSsl = true;//经过ssl加密    
                //
                //启用功能修改处
                //
                System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage(fromi, toi);
                message.Subject = Subjecti;
                message.Body = Bodyi;
                message.BodyEncoding = System.Text.Encoding.UTF8;
                message.IsBodyHtml = true;
                //  message.Headers.Add("X-Mailer", "Microsoft Outlook");

                //添加附件需将(附件先上传到服务器)
                for (int i = 0; i < Attachmentlist.Length; i++)
                {
                    System.Net.Mail.Attachment data = new System.Net.Mail.Attachment(Attachmentlist[i], System.Net.Mime.MediaTypeNames.Application.Octet);
                    message.Attachments.Add(data);
                }
                try
                {
                    client.Send(message);
                    //  this.lbMessage.Text = "登录名和密码已经发送到您的" + "512250428@qq.com" + "邮箱!";
                }
                catch (Exception ex)
                {
                    // this.lbMessage.Text = "Send Email Failed." + ex.ToString();
                }
            }
        }

    }
}
