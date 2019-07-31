using dblist;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
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

namespace 测试功能专用
{
    public partial class frmPDF : Form
    {

        private int year;


        public frmPDF()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string A_Path = AppDomain.CurrentDomain.BaseDirectory + "Resources\\PDFfolder";

            List<string> Alist = GetFileName(A_Path);




            List<clsPDFInfo> PDFInfoList = ReadingFile(Alist);
        }
        private List<string> GetFileName(string dirPath)
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
                    FileNameList.Add(item.ToString().Replace(dirPath + "\\", ""));
                }
            }
            return FileNameList;
        }
        private List<clsPDFInfo> ReadingFile(List<string> PathList)
        {
            year = 2017;

            try
            {
                List<clsPDFInfo> PDFInfoList = new List<clsPDFInfo>();

                //string B_Path = AppDomain.CurrentDomain.BaseDirectory + "DataResources\\Readed";
                string B_Path = AppDomain.CurrentDomain.BaseDirectory + "Resources\\PDFfolder";

                #region 循环新文件路径List 读取本地PDF
                foreach (string Path in PathList)
                {
                    string rdfile = B_Path + "\\" + Path;
                    //  rdfile = @"C:\Test\测试功能专用\测试功能专用\bin\Debug\Resources\PDFfolder\bak\万向信托 2017 公司年报.pdf";

                    PdfReader reader = new PdfReader(rdfile);
                    PdfReaderContentParser parser = new PdfReaderContentParser(reader);
                    int rowCount = reader.NumberOfPages;
                    clsPDFInfo temp = new clsPDFInfo();


                    int nextisme = 0;

                    for (int i = 1; i <= rowCount; i++)
                    {
                        ITextExtractionStrategy strategy = parser.ProcessContent<SimpleTextExtractionStrategy>(i, new SimpleTextExtractionStrategy());
                        //将文本内容赋值给一个富文本框
                        string Text = strategy.GetResultantText();
                        string[] texttemp = System.Text.RegularExpressions.Regex.Split(Text, "\n");
                        //取得PDF内详细信息
                        //temp.PDFinfo = Text;
                        //texttemp[]
                        temp.PDF_File_Name = Path;
                        temp.PDF_File_pag = i.ToString();

                        //信托公司简称
                        if (texttemp.Length > 3 && i == 1)
                        {
                            int index = Array.FindIndex(texttemp, a => a.Contains("公司"));

                            temp.xintuogongsijiancheng = texttemp[index].ToUpper();
                            //信托公司全称
                            temp.xintuogongsiquancheng = texttemp[index].ToUpper();
                        }
                        //成立日期

                        //注册地省份
                        if (Text.Contains("注册地址"))
                        {
                            int index = Array.FindIndex(texttemp, a => a.Contains("注册地址"));
                            if (index > 0)
                            {
                                temp.zhucedizhi = texttemp[index].Replace("注册地址", "").Replace(":", "");
                            }
                        }
                        //注册地省份


                        //所属银监局


                        //信托资产运用与分布表

                        #region 信托资产运用与分布表
                        if (Text.Replace(" ", "").Contains("信托资产运用与") || nextisme == 1)
                        {
                            if (nextisme == 0)
                            {
                                nextisme++;

                                continue;
                            }
                            int index = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("货币资产"));
                            if (index > 0)
                            {
                                string[] texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //货币资金
                                temp.huobizijinA9 = texttemp1[2];

                                //基础产业
                                temp.jichuchanyeA18 = texttemp1[6];


                                index = Array.FindIndex(texttemp, a => a.Contains("贷款"));
                                if (index > 0)
                                {
                                    texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                    //贷款
                                    temp.huokuanA10 = texttemp1[3];
                                    //房地产
                                    temp.fangdichanA19 = texttemp1[7];
                                }
                                //
                                index = Array.FindIndex(texttemp, a => a.Contains("交易性"));
                                if (index > 0)
                                {
                                    texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                    //交易性金融资产
                                    temp.jiaoyixingjinrongzichanA11 = texttemp1[3];
                                    //证券市场
                                    temp.zhengjuanshichangA20 = texttemp1[7];
                                }
                                index = Array.FindIndex(texttemp, a => a.Contains("工商企"));
                                if (index > 0)
                                {
                                    texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                    //工商企业
                                    temp.gongshangqiyeA21 = texttemp1[8];
                                }

                                index = Array.FindIndex(texttemp, a => a.Contains("可供出"));
                                if (index > 0)
                                {
                                    texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                    //可供出售金融资产
                                    temp.kegongchushoujinrongzichanA12 = texttemp1[4];
                                    //金融机构
                                    temp.jinrongjigouA22 = texttemp1[8];
                                }

                                index = Array.FindIndex(texttemp, a => a.Contains("持有至"));
                                if (index > 0)
                                {
                                    texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                    //持有至到期投资
                                    temp.chiyouzhidaoqitouziA13 = texttemp1[3];
                                }


                                index = Array.FindIndex(texttemp, a => a.Contains("长期股"));
                                if (index > 0)
                                {
                                    texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                    //长期股权投资
                                    temp.chiyouzhidaoqitouziA13 = texttemp1[3];

                                }
                                index = Array.FindIndex(texttemp, a => a.Contains("资产总"));
                                if (index > 0)
                                {
                                    texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                    //合计
                                    temp.hejiA16 = texttemp1[2];
                                    //合计
                                    //    temp.hejiA24 = texttemp1[6];
                                }
                            }
                            nextisme = 0;

                        }
                        #endregion

                        //
                        #region  信托资产 	 2018年末 	 占比 	 2017年末 	 占比 	 2016年末 	 占比

                        if (Text.Replace(" ", "").Contains("信托资产管理情"))
                        {

                            int index = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("集合"));
                            if (index > 0)
                            {
                                string[] texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //集合
                                if (texttemp1.Length > 1)
                                    temp.jiheA27 = texttemp1[1];
                                //集合2016
                                if (texttemp1.Length > 2)
                                    temp.jiheA27_F27 = texttemp1[2];


                                index = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("单一"));
                                if (index > 0)
                                {
                                    texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                    //单一
                                    if (texttemp1.Length > 1)
                                        temp.danyiA28 = texttemp1[1];
                                    //单一2016
                                    if (texttemp1.Length > 2)
                                        temp.danyiA28_F27 = texttemp1[2];
                                }
                            }
                        }

                        #endregion

                        #region  // 主动管理型信托 	 2018年末 	 占比 	 2017年末 	 占比 	 2016年末 	 占比

                        if (Text.Replace(" ", "").Contains("主动管理型信托") && temp.zhengjuanleiA33 == null)
                        {
                            string[] texttemp1 = null;
                            int type2 = 1;

                            int index = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("证券"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //集合
                                if (texttemp1.Length > 2)
                                {
                                    string result = System.Text.RegularExpressions.Regex.Replace(texttemp1[2], @"[^0-9]+", "");
                                    if (result != "")
                                        temp.zhengjuanleiA33 = texttemp1[2];
                                    else
                                    {
                                        texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index + 1], " ");
                                        result = System.Text.RegularExpressions.Regex.Replace(texttemp1[1], @"[^0-9]+", "");
                                        if (result != "")
                                        {
                                            temp.zhengjuanleiA33 = texttemp1[1];
                                            if (texttemp1.Length > 2)
                                            {
                                                temp.zhengjuanleiA33_F41 = texttemp1[2];
                                                type2 = 2;
                                            }
                                        }
                                    }
                                }
                                //集合2016
                                if (texttemp1.Length > 3 && type2 == 1)
                                {
                                    temp.zhengjuanleiA33_F41 = texttemp1[3];


                                }
                            }

                            index = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("股权"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //单一
                                if (texttemp1.Length > 2)
                                    temp.guquanleiA34 = texttemp1[2];
                                //单一2016
                                if (texttemp1.Length > 3)
                                    temp.guquanleiA34_F34 = texttemp1[3];
                            }

                            index = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("融资"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //单一
                                if (texttemp1.Length > 1)
                                    temp.rongzileiA35 = texttemp1[1];
                                //单一2016
                                if (texttemp1.Length > 2)
                                    temp.guquanleiA34_F34 = texttemp1[2];
                            }

                            index = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("事务"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //单一
                                if (texttemp1.Length > 2)
                                    temp.shiwuguanlileiA37 = texttemp1[2];
                                //单一2016
                                if (texttemp1.Length > 3)
                                    temp.shiwuguanlileiA37_F37 = texttemp1[3];
                            }
                        }
                        #endregion
                        #region  //  被动管理型信托 	 2018年末 	 占比 	 2017年末 	 占比 	 2016年末 	 占比

                        if (Text.Replace(" ", "").Contains("被动管理型信托") && temp.zhengjuanleiA41 == null)
                        {
                            string[] texttemp1 = null;
                            bool isdone = false;

                            int index = Array.FindLastIndex(texttemp, a => a.Replace(" ", "").Contains("证券"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //集合
                                if (texttemp1.Length > 2)
                                {
                                    string result = System.Text.RegularExpressions.Regex.Replace(texttemp1[2], @"[^0-9]+", "");
                                    if (result != "")
                                        temp.zhengjuanleiA41 = texttemp1[2];
                                    else
                                    {
                                        int newindex = index;

                                        while (true)
                                        {
                                            checkisshuzi(texttemp, ref isdone, ref newindex);

                                            if (isdone == true)
                                            {
                                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[newindex], " ");

                                                if (texttemp1.Length > 2)
                                                {
                                                    temp.zhengjuanleiA41 = texttemp1[2];

                                                }
                                                if (texttemp1.Length > 3)
                                                {
                                                    temp.zhengjuanleiA33_F41 = texttemp1[3];

                                                }
                                                break;

                                            }
                                        }
                                    }
                                }
                                //集合2016
                                if (texttemp1.Length > 3 == isdone == false)
                                    temp.zhengjuanleiA33_F41 = texttemp1[3];
                            }

                            index = Array.FindLastIndex(texttemp, a => a.Replace(" ", "").Contains("股权"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //单一
                                if (texttemp1.Length > 2)
                                    temp.guquanleiA42 = texttemp1[2];
                                //单一2016
                                if (texttemp1.Length > 3)
                                    temp.guquanleiA42_F41 = texttemp1[3];
                            }

                            index = Array.FindLastIndex(texttemp, a => a.Replace(" ", "").Contains("融资"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //单一
                                if (texttemp1.Length > 1)
                                    temp.rongzileiA43 = texttemp1[1];
                                //单一2016
                                if (texttemp1.Length > 2)
                                    temp.rongzileiA43_F41 = texttemp1[2];
                            }

                        }
                        #endregion

                        #region  本年清算信托 	 2018年个数 	 2018年金额 	 2018年加权年化收益率 	 2018加权信托报酬 	 2017年个数 	 2017年金额 	 2017年加权年化收益率 	 2017加权信托报酬
                        if (Text.Replace(" ", "").Contains("已清算结束信托项目") && temp.jiheA49 == null)
                        {
                            string[] texttemp1 = null;

                            int index = Array.FindLastIndex(texttemp, a => a.Replace(" ", "").Contains("集合"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //2017年个数
                                if (texttemp1.Length > 1)
                                    temp.jiheA49 = texttemp1[1];
                                //2017年金额
                                if (texttemp1.Length > 2)
                                    temp.jiheA49_G49 = texttemp1[2];
                                //2017年加权年化收益率
                                if (texttemp1.Length > 3)
                                    temp.jiheA49_H49 = texttemp1[3];

                            }
                            index = Array.FindLastIndex(texttemp, a => a.Replace(" ", "").Contains("单一"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //2017年个数
                                if (texttemp1.Length > 1)
                                    temp.danyiA50 = texttemp1[1];
                                //2017年金额
                                if (texttemp1.Length > 2)
                                    temp.danyiA50_G50 = texttemp1[2];
                                //2017年加权年化收益率
                                if (texttemp1.Length > 3)
                                    temp.danyiA50_H50 = texttemp1[3];
                            }


                            index = Array.FindLastIndex(texttemp, a => a.Replace(" ", "").Contains("财产"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //2017年个数
                                if (texttemp1.Length > 1)
                                    temp.caichanquanA51 = texttemp1[1];
                                //2017年金额
                                if (texttemp1.Length > 2)
                                    temp.caichanquanA51_F51 = texttemp1[2];
                                //2017年加权年化收益率
                                if (texttemp1.Length > 3)
                                    temp.caichanquanA51_G51 = texttemp1[3];
                            }
                        }
                        //已清算的主动管理型信托项目情况
                        int indexS = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("已清算的主动管理型信托项目情况"));

                        if (Text.Replace(" ", "").Contains("已清算结束信托项目") && indexS > 0)
                        {
                            string[] texttemp1 = null;

                            // 证券类-主动
                            int index = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("证券投"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //2017年个数
                                if (texttemp1.Length > 1)
                                    temp.zhengjuanleizhudongA53 = texttemp1[1];
                                //2017年金额
                                if (texttemp1.Length > 2)
                                    temp.zhengjuanleizhudongA53_G53 = texttemp1[2];
                                //2017年加权年化收益率
                                if (texttemp1.Length > 3)
                                    temp.zhengjuanleizhudongA53_H53 = texttemp1[3];
                            }
                            //股权类-主动
                            index = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("股权投"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //2017年个数
                                if (texttemp1.Length > 1)
                                    temp.guquanleizhudongA54 = texttemp1[1];
                                //2017年金额
                                if (texttemp1.Length > 2)
                                    temp.guquanleizhudongA54_G54 = texttemp1[2];
                                //2017年加权年化收益率
                                if (texttemp1.Length > 3)
                                    temp.guquanleizhudongA54_H54 = texttemp1[3];
                            }

                            //融资类-主动
                            index = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("融资类"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //2017年个数
                                if (texttemp1.Length > 1)
                                    temp.rongzileizhudongA55 = texttemp1[1];
                                //2017年金额
                                if (texttemp1.Length > 2)
                                    temp.rongzileizhudongA55_G55 = texttemp1[2];
                                //2017年加权年化收益率
                                if (texttemp1.Length > 3)
                                    temp.rongzileizhudongA55_H55 = texttemp1[3];
                            }
                            //其他类-主动
                            //index = Array.FindLastIndex(texttemp, a => a.Replace(" ", "").Contains("财产"));
                            //texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                            ////2017年个数
                            //temp.caichanquanA51 = texttemp1[1];
                            ////2017年金额
                            //temp.caichanquanA51_F51 = texttemp1[2];
                            ////2017年加权年化收益率
                            //temp.caichanquanA51_G51 = texttemp1[3];

                            //事务管理类-主动
                            index = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("事务"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //2017年个数
                                if (texttemp1.Length > 2)
                                    temp.shiwuguanlileizhudongA57 = texttemp1[2];
                                //2017年金额
                                if (texttemp1.Length > 3)
                                    temp.shiwuguanlileizhudongA57_G57 = texttemp1[3];
                                //2017年加权年化收益率
                                if (texttemp1.Length > 4)
                                    temp.shiwuguanlileizhudongA57_H57 = texttemp1[4];

                            }
                        }
                        //

                        #region 已清算的被动管理型信托项目情况
                        indexS = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("已清算的被动管理型信托项目情况"));

                        if (Text.Replace(" ", "").Contains("已清算结束信托项目") && indexS > 0)
                        {
                            string[] texttemp1 = null;

                            //证券类-被动
                            int index = Array.FindLastIndex(texttemp, a => a.Replace(" ", "").Contains("证券投"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //2017年个数
                                if (texttemp1.Length > 1)
                                    temp.zhengjuanleibeidongA59 = texttemp1[1];
                                //2017年金额
                                if (texttemp1.Length > 2)
                                    temp.zhengjuanleibeidongA59_F59 = texttemp1[2];
                                //2017年加权年化收益率
                                if (texttemp1.Length > 3)
                                    temp.zhengjuanleibeidongA59_H59 = texttemp1[3];
                            }

                            //股权类-被动
                            index = Array.FindLastIndex(texttemp, a => a.Replace(" ", "").Contains("股权"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //2017年个数
                                if (texttemp1.Length > 2)
                                    temp.guquanleibeidongA60 = texttemp1[2];
                                //2017年金额
                                if (texttemp1.Length > 3)
                                    temp.guquanleibeidongA60_F60 = texttemp1[3];
                                //2017年加权年化收益率
                                if (texttemp1.Length > 4)
                                    temp.guquanleibeidongA60_H60 = texttemp1[4];
                            }

                            //融资类-被动
                            index = Array.FindLastIndex(texttemp, a => a.Replace(" ", "").Contains("融资类"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //2017年个数
                                if (texttemp1.Length > 1)
                                    temp.rongzileibeidongA61 = texttemp1[1];
                                //2017年金额
                                if (texttemp1.Length > 2)
                                    temp.rongzileibeidongA61_F61 = texttemp1[2];
                                //2017年加权年化收益率
                                if (texttemp1.Length > 3)
                                    temp.rongzileibeidongA61_H61 = texttemp1[3];
                            }

                            ////其他类-被动
                            //index = Array.FindLastIndex(texttemp, a => a.Replace(" ", "").Contains("财产"));
                            //texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                            ////2017年个数
                            //temp.caichanquanA51 = texttemp1[1];
                            ////2017年金额
                            //temp.caichanquanA51_F51 = texttemp1[2];
                            ////2017年加权年化收益率
                            //temp.caichanquanA51_G51 = texttemp1[3];

                            //事务管理类-被动
                            index = Array.FindLastIndex(texttemp, a => a.Replace(" ", "").Contains("事务管"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //2017年个数
                                if (texttemp1.Length > 2)
                                    temp.shiwuguanlileibeidongA63 = texttemp1[2];
                                //2017年金额
                                if (texttemp1.Length > 3)
                                    temp.shiwuguanlileibeidongA63_F63 = texttemp1[3];
                                //2017年加权年化收益率
                                if (texttemp1.Length > 4)
                                    temp.shiwuguanlileibeidongA63_H63 = texttemp1[4];
                            }


                        }
                        #endregion

                        #endregion


                        #region  本年新增信托 	 2018年个数 	 2018年金额 	 占比 	 2017年个数 	 2017年金额 	 占比

                        //本年度新增的集合类
                        indexS = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("本年度新增的集合类"));

                        if (Text.Replace(" ", "").Contains("新增信托项目") && indexS > 0)
                        {
                            string[] texttemp1 = null;

                            // 集合
                            int index = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("集合") && !a.Replace(" ", "").Contains("本年度新增"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //2017年个数
                                if (texttemp1.Length > 1)
                                    temp.jieheA67 = texttemp1[1];
                                //2017年金额
                                if (texttemp1.Length > 2)
                                    temp.jieheA67_F67 = texttemp1[2];
                            }
                            //单一
                            index = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("单一") && !a.Replace(" ", "").Contains("本年度新增"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //2017年个数
                                if (texttemp1.Length > 1)
                                    temp.danyiA68 = texttemp1[1];
                                //2017年金额
                                if (texttemp1.Length > 2)
                                    temp.danyiA68_F68 = texttemp1[2];
                            }

                            //财产权
                            index = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("财产") && !a.Replace(" ", "").Contains("本年度新增"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //2017年个数
                                if (texttemp1.Length > 2)
                                    temp.danyiA68 = texttemp1[2];
                                //2017年金额
                                if (texttemp1.Length > 3)
                                    temp.danyiA68_F68 = texttemp1[3];

                            }
                            //主动管理
                            index = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("主动管") && a.Replace(" ", "").Contains("其中："));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //2017年个数
                                if (texttemp1.Length > 4)
                                    temp.zhudongguanliA71 = texttemp1[4];
                                //2017年金额
                                if (texttemp1.Length > 5)
                                    temp.zhudongguanliA71_F71 = texttemp1[5];
                            }
                        }

                        #endregion


                        #region  公司收入结构 	 2018年万元 	 占比 	 2017年万元 	 占比 	 2016年万元 	 占比

                        //本年度新增的集合类
                        indexS = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("公司当年的收入结构"));

                        if (Text.Replace(" ", "").Contains("公司当年的收入结构") && indexS > 0)
                        {
                            string[] texttemp1 = null;

                            // 集合
                            int index = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("手续费") && !a.Replace(" ", "").Contains("本年度新增"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //2017年个数
                                if (texttemp1.Length > 4)
                                    temp.shouxufeijiyongjinA75 = texttemp1[4];
                            }

                            //利息收入
                            index = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("利息收入") && !a.Replace(" ", "").Contains("本年度新增"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //2017年个数
                                if (texttemp1.Length > 2)
                                    temp.lixishouruA76 = texttemp1[2];
                            }


                            //投资收益 
                            index = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("投资收") && !a.Replace(" ", "").Contains("本年度新增"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //2017年个数
                                if (texttemp1.Length > 2)
                                    temp.touzishouyiA77 = texttemp1[2];

                            }

                            // 其中：股权 

                            index = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("股权") && a.Replace(" ", "").Contains("其中："));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //2017年个数
                                if (texttemp1.Length > 4)
                                    temp.qizhongguquanA78 = texttemp1[4];
                            }
                            // 其中：证券 

                            index = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("证券投"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //2017年个数
                                if (texttemp1.Length > 3)
                                    temp.qizhongzhengjuanA79 = texttemp1[3];
                            }

                            // 其中：其他

                            index = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("其他投资收益"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //2017年个数
                                if (texttemp1.Length > 3)
                                    temp.qizhongqitaA80 = texttemp1[3];
                            }

                            //公允价值变动损益

                            index = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("公允价值"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //2017年个数
                                if (texttemp1.Length > 5)
                                    temp.gongyunjiazhibiandongshunyiA81 = texttemp1[5];
                            }

                            //收入合计

                            index = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("收入合"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //2017年个数
                                if (texttemp1.Length > 2)
                                    temp.shouruhejiA83 = texttemp1[2];
                            }

                        }



                        #endregion

                        #region  单位：亿元 	 2018年 	 2017年 	 2016年 	 2015年 	 2014年


                        indexS = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("母公司名称"));

                        if (Text.Replace(" ", "").Contains("母公司名称") && indexS > 0)
                        {

                            indexS = Text.Replace(" ", "").IndexOf("关联交易方与本公司的关系情况");
                            int indexE = Text.Replace(" ", "").IndexOf("其他关联交易方情况");

                            string newtxt = Text.Substring(indexS, indexS + 500 - indexS);

                            string[] texttemp2 = System.Text.RegularExpressions.Regex.Split(newtxt, "\n");
                            // 注册资本
                            int index = Array.FindIndex(texttemp2, a => a.Replace(" ", "").Contains("本公司的最终控制方"));
                            string[] texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp2[index - 1], " ");
                            //2017年个数
                            if (texttemp1.Length > 4)
                                temp.zhucezibenJ2 = texttemp1[4];


                            //固有资产
                            index = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("固有业务") && !a.Replace(" ", "").Contains("本年度新增"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //2017年个数
                                if (texttemp1.Length > 2)
                                    temp.lixishouruA76 = texttemp1[2];
                            }
                        }
                        indexS = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("固有业务"));

                        if (Text.Replace(" ", "").Contains("固有业务") && indexS > 0)
                        {
                            //固有资产
                            int index = Array.FindLastIndex(texttemp, a => a.Replace(" ", "").Contains("资产总计") && !a.Replace(" ", "").Contains("本年度新增"));
                            if (index > 0)
                            {
                                string[] texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                                //2017年个数
                                if (texttemp1.Length > 2)
                                    temp.guyouzichanJ3 = texttemp1[2];
                            }
                        }

                        #endregion

                        #region  职工人数 	 2018年 	 占比 	 2017年 	 占比 	 2016年
                        //本年度新增的集合类
                        indexS = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("员工分布表"));

                        if (Text.Replace(" ", "").Contains("员工分布表") && indexS > 0)
                        {

                            indexS = Text.Replace(" ", "").IndexOf("员工分布表");
                            string newtxt = Text.Substring(indexS, indexS + 500 - indexS);
                            string[] texttemp2 = System.Text.RegularExpressions.Regex.Split(newtxt, "% \n");

                            int index = Array.FindIndex(texttemp2, a => a.Replace(" ", "").Contains("董事") && !a.Replace(" ", "").Contains("注"));
                            string[] texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp2[index], "\n");
                            string[] texttemp3 = System.Text.RegularExpressions.Regex.Split(texttemp1[texttemp1.Length - 1], " ");

                            //管理人员
                            if (texttemp1.Length > 0)
                                temp.J20guanlirenyuan = texttemp1[0];


                            //固有业务
                            //index = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("利息收入") && !a.Replace(" ", "").Contains("本年度新增"));
                            //texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                            ////2017年个数
                            //temp.lixishouruA76 = texttemp1[2];

                            //信托业务

                            index = Array.FindIndex(texttemp2, a => a.Replace(" ", "").Contains("信托"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp2[index], " ");
                                //信托业务
                                if (texttemp1.Length > 3)
                                    temp.J22xintuoyewu = texttemp1[3];
                            }
                            //其他
                            index = Array.FindIndex(texttemp2, a => a.Replace(" ", "").Contains("其他人"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp2[index], " ");
                                //其他
                                if (texttemp1.Length > 2)
                                    temp.J23qita = texttemp1[2];
                            }
                            //30以下
                            index = Array.FindIndex(texttemp2, a => a.Replace(" ", "").Contains("30以下"));
                            if (index > 0)
                            {
                                string[] texttemp11 = System.Text.RegularExpressions.Regex.Split(texttemp2[index], "\n");
                                index = Array.FindIndex(texttemp11, a => a.Replace(" ", "").Contains("30以下"));

                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp11[index], " ");
                                //30以下
                                if (texttemp1.Length > 2)
                                    temp.J25_30yixia = texttemp1[2];
                            }
                            //30-39
                            index = Array.FindIndex(texttemp2, a => a.Replace(" ", "").Contains("30-39"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp2[index], " ");
                                // 
                                if (texttemp1.Length > 1)
                                    temp.J24heji = texttemp1[1];
                            }

                            //40以上
                            index = Array.FindIndex(texttemp2, a => a.Replace(" ", "").Contains("40以上"));
                            if (index > 0)
                            {
                                texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp2[index], " ");
                                // 
                                if (texttemp1.Length > 2)
                                    temp.J27_40yishang = texttemp1[2];
                            }


                        }



                        #endregion


                        #region  股东
                        //indexS = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("股东名称"));

                        //if (Text.Replace(" ", "").Contains("股东") && indexS > 0)
                        //{
                        //    indexS = Text.Replace(" ", "").IndexOf("股东名称");
                        //    string tx2 = Text.Substring(indexS, Text.Length - indexS);
                        //    int indexE = tx2.IndexOf("注：");
                        //    string newtxt = tx2.Substring(1, indexE - 1);

                        //    string[] texttemp2 = System.Text.RegularExpressions.Regex.Split(newtxt, "\n");
                        //    // 股东名称
                        //    int index = Array.FindIndex(texttemp2, a => a.Replace(" ", "").Contains("经营业务"));
                        //    string[] texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp2[index], " ");
                        //    //股东名称
                        //    //   temp.J32 = texttemp1[4];










                        //    ////利息收入
                        //    //index = Array.FindIndex(texttemp, a => a.Replace(" ", "").Contains("利息收入") && !a.Replace(" ", "").Contains("本年度新增"));
                        //    //texttemp1 = System.Text.RegularExpressions.Regex.Split(texttemp[index], " ");
                        //    ////2017年个数
                        //    //temp.lixishouruA76 = texttemp1[2];

                        //}





                        #endregion
                    }
                    PDFInfoList.Add(temp);
                    DownExcel(PDFInfoList, Path);

                    //清空
                    PDFInfoList = new List<clsPDFInfo>();

                }
                return PDFInfoList;
                #endregion
            }
            catch (Exception ex)
            {
                MessageBox.Show("Read error information in PDF：" + ex.Message);
                return null;
            }
            return null;
        }

        private static void checkisshuzi(string[] texttemp, ref bool isdone, ref int newindex)
        {
            if (System.Text.RegularExpressions.Regex.Replace(texttemp[newindex], @"[^0-9]+", "") != "")
            {
                isdone = true;

            }
            else
            {
                newindex++;

            }
        }

        private void DownExcel(List<clsPDFInfo> Results, string filename)
        {
            #region 获取模板路径
            System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            string fullPath = Path.Combine(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "System\\"), "信托数据分析-案例 2.xlsx");

            //fullPath = @"C:\Test\测试功能专用\测试功能专用\bin\Debug\Resources\PDFfolder\results\万向信托 2016-2017公司年报.xlsx";
            SaveFileDialog sfdDownFile = new SaveFileDialog();
            sfdDownFile.OverwritePrompt = false;
            string DesktopPath = Convert.ToString(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            sfdDownFile.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*.xlsx";
            string file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Results\\");

            sfdDownFile.FileName = Path.Combine(file, filename + DateTime.Now.ToString("yyyyMMdd"));
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
                    Microsoft.Office.Interop.Excel._Worksheet ExcelSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelBook.Worksheets["模版"];
                    //打开时是否显示Excel




            #endregion

                    #region 填充数据
                    int RowIndex = 1;
                    int xuhao = 1;

                    List<string> nashui = new List<string>();
                    List<int> xuhaoindex = new List<int>();

                    foreach (clsPDFInfo item in Results)
                    {

                        RowIndex++;
                        #region 2016
                        if (year == 2016)
                        {
                            //ExcelSheet.Cells[RowIndex, 1] = item.xuhao;
                            //ExcelSheet.Cells[RowIndex, 2] = item.gongsidaima;////B列公司代码= line 预提上传纳税单位
                            //ExcelSheet.Cells[RowIndex, 3] = DateTime.Now.ToString("yyyyMMdd");
                            ////ExcelSheet.Cells[RowIndex, 4] = item.jiaoyiriqi;
                            //ExcelSheet.Cells[RowIndex, 4] = clsCommHelp.objToDateTime1(item.jiaoyiriqi).Replace("/", "");//  
                            //ExcelSheet.Cells[RowIndex, 5] = "EV";
                            //ExcelSheet.Cells[RowIndex, 6] = "CNY";
                            //ExcelSheet.Cells[RowIndex, 9] = DateTime.Now.ToString("yyyyMM") + "利息收入";
                            ExcelSheet.Cells[1, 2] = item.xintuogongsijiancheng;
                            ExcelSheet.Cells[2, 2] = item.xintuogongsiquancheng;
                            ExcelSheet.Cells[3, 2] = item.chengligongsi;
                            ExcelSheet.Cells[4, 2] = item.zhucedishengfen;
                            ExcelSheet.Cells[5, 2] = item.zhucedizhi;


                            ExcelSheet.Cells[1, 7] = item.dongshizhang;
                            ExcelSheet.Cells[2, 7] = item.zongjinli;
                            ExcelSheet.Cells[3, 5] = item.xinyeyinhanggongfenyouxiangongsi;
                            ExcelSheet.Cells[4, 5] = item.aodaliyaguomingyinhang;
                            ExcelSheet.Cells[5, 5] = item.fujianshengnengyuanjituanyouxiangongsi;


                            // 信托资产运用 	 2018年 	 占比 	 2017年 	 占比 	 2016年 	 占比 

                            ExcelSheet.Cells[9, 6] = item.huobizijinA9;
                            ExcelSheet.Cells[10, 6] = item.huokuanA10;
                            ExcelSheet.Cells[11, 6] = item.jiaoyixingjinrongzichanA11;
                            ExcelSheet.Cells[12, 6] = item.kegongchushoujinrongzichanA12;
                            ExcelSheet.Cells[13, 6] = item.chiyouzhidaoqitouziA13;
                            ExcelSheet.Cells[14, 6] = item.changqiguquantouziA14;
                            ExcelSheet.Cells[16, 6] = item.hejiA16;

                            ExcelSheet.Cells[18, 6] = item.jichuchanyeA18;
                            ExcelSheet.Cells[19, 6] = item.fangdichanA19;
                            ExcelSheet.Cells[20, 6] = item.zhengjuanshichangA20;
                            ExcelSheet.Cells[21, 6] = item.gongshangqiyeA21;
                            ExcelSheet.Cells[22, 6] = item.jinrongjigouA22;

                            // 信托资产 	 2018年末 	 占比 	 2017年末 	 占比 	 2016年末 	 占比 

                            ExcelSheet.Cells[27, 6] = item.jiheA27_F27;
                            ExcelSheet.Cells[28, 6] = item.danyiA28_F27;


                            // 主动管理型信托 	 2018年末 	 占比 	 2017年末 	 占比 	 2016年末 	 占比 

                            ExcelSheet.Cells[33, 6] = item.zhengjuanleiA33_F41;
                            ExcelSheet.Cells[34, 6] = item.guquanleiA34_F34;
                            ExcelSheet.Cells[35, 6] = item.rongzileiA35_F43;
                            ExcelSheet.Cells[36, 6] = item.qitaleiA36;
                            ExcelSheet.Cells[37, 6] = item.shiwuguanlileiA37_F37;
                            // 被动管理型信托 	 2018年末 	 占比 	 2017年末 	 占比 	 2016年末 	 占比 

                            ExcelSheet.Cells[41, 6] = item.zhengjuanleiA41_F41;
                            ExcelSheet.Cells[42, 6] = item.guquanleiA42_F41;
                            ExcelSheet.Cells[43, 6] = item.rongzileiA43_F41;
                            ExcelSheet.Cells[44, 6] = item.qitaleiA44;




                            // 公司收入结构 	 2018年万元 	 占比 	 2017年万元 	 占比 	 2016年万元 	 占比 

                            ExcelSheet.Cells[75, 6] = item.shouxufeijiyongjinA75;
                            ExcelSheet.Cells[76, 6] = item.lixishouruA76;
                            ExcelSheet.Cells[77, 6] = item.touzishouyiA77;
                            ExcelSheet.Cells[78, 6] = item.qizhongguquanA78;
                            ExcelSheet.Cells[79, 6] = item.qizhongzhengjuanA79;
                            ExcelSheet.Cells[80, 6] = item.qizhongqitaA80;
                            ExcelSheet.Cells[81, 6] = item.gongyunjiazhibiandongshunyiA81;
                            ExcelSheet.Cells[83, 6] = item.shouruhejiA83;


                            // 单位：亿元 	 2018年 	 2017年 	 2016年 	 2015年 	 2014年 

                            ExcelSheet.Cells[2, 13] = item.zhucezibenJ2;
                            ExcelSheet.Cells[3, 13] = item.guyouzichanJ3;

                            //  职工人数 	 2018年 	 占比 	 2017年 	 占比 	 2016年 
                            ExcelSheet.Cells[20, 15] = item.J20guanlirenyuan;
                            ExcelSheet.Cells[21, 15] = item.J21guyouyewu;

                            ExcelSheet.Cells[22, 15] = item.J22xintuoyewu;
                            ExcelSheet.Cells[23, 15] = item.J23qita;

                            ExcelSheet.Cells[24, 15] = item.J24heji;
                            ExcelSheet.Cells[25, 15] = item.J25_30yixia;

                            ExcelSheet.Cells[26, 15] = item.J26_30_39yixia;
                            ExcelSheet.Cells[27, 15] = item.J27_40yishang;


                        }
                        #endregion

                        #region 2017
                        if (year == 2017)
                        {
                            // 信托资产运用 	 2018年 	 占比 	 2017年 	 占比 	 2016年 	 占比 

                            ExcelSheet.Cells[9, 4] = item.huobizijinA9;
                            ExcelSheet.Cells[10, 4] = item.huokuanA10;
                            ExcelSheet.Cells[11, 4] = item.jiaoyixingjinrongzichanA11;
                            ExcelSheet.Cells[12, 4] = item.kegongchushoujinrongzichanA12;
                            ExcelSheet.Cells[13, 4] = item.chiyouzhidaoqitouziA13;
                            ExcelSheet.Cells[14, 4] = item.changqiguquantouziA14;
                            ExcelSheet.Cells[16, 4] = item.hejiA16;

                            ExcelSheet.Cells[18, 4] = item.jichuchanyeA18;
                            ExcelSheet.Cells[19, 4] = item.fangdichanA19;
                            ExcelSheet.Cells[20, 4] = item.zhengjuanshichangA20;
                            ExcelSheet.Cells[21, 4] = item.gongshangqiyeA21;
                            ExcelSheet.Cells[22, 4] = item.jinrongjigouA22;


                            // 信托资产 	 2018年末 	 占比 	 2017年末 	 占比 	 2016年末 	 占比 

                            ExcelSheet.Cells[27, 4] = item.jiheA27;
                            ExcelSheet.Cells[28, 4] = item.danyiA28;


                            // 主动管理型信托 	 2018年末 	 占比 	 2017年末 	 占比 	 2016年末 	 占比 

                            ExcelSheet.Cells[33, 4] = item.zhengjuanleiA33_F41;
                            ExcelSheet.Cells[34, 4] = item.guquanleiA34_F34;
                            ExcelSheet.Cells[35, 4] = item.rongzileiA35_F43;
                            ExcelSheet.Cells[36, 4] = item.qitaleiA36;
                            ExcelSheet.Cells[37, 4] = item.shiwuguanlileiA37_F37;


                            // 被动管理型信托 	 2018年末 	 占比 	 2017年末 	 占比 	 2016年末 	 占比 

                            ExcelSheet.Cells[41, 4] = item.zhengjuanleiA41;
                            ExcelSheet.Cells[42, 4] = item.guquanleiA42;
                            ExcelSheet.Cells[43, 4] = item.rongzileiA43;
                            ExcelSheet.Cells[44, 4] = item.qitaleiA44;

                            // 本年清算信托 	 2018年个数 	 2018年金额 	 2018年加权年化收益率 	 2018加权信托报酬 	 2017年个数 	 2017年金额 	 2017年加权年化收益率 	 2017加权信托报酬 

                            ExcelSheet.Cells[49, 6] = item.jiheA49_G49;
                            ExcelSheet.Cells[49, 6] = item.jiheA49_H49;


                            ExcelSheet.Cells[50, 6] = item.danyiA50_G50;
                            ExcelSheet.Cells[50, 7] = item.danyiA50_H50;


                            ExcelSheet.Cells[51, 6] = item.caichanquanA51_F51;
                            ExcelSheet.Cells[51, 7] = item.caichanquanA51_G51;




                            ExcelSheet.Cells[53, 6] = item.zhengjuanleizhudongA53_G53;

                            ExcelSheet.Cells[53, 7] = item.zhengjuanleizhudongA53_H53;


                            ExcelSheet.Cells[54, 6] = item.guquanleizhudongA54_G54;
                            ExcelSheet.Cells[54, 7] = item.guquanleizhudongA54_H54;
                            ExcelSheet.Cells[55, 6] = item.rongzileizhudongA55_G55;
                            ExcelSheet.Cells[55, 6] = item.rongzileizhudongA55_H55;



                            ExcelSheet.Cells[56, 6] = item.qitaleizhudongA56;

                            ExcelSheet.Cells[57, 6] = item.shiwuguanlileizhudongA57_G57;
                            ExcelSheet.Cells[57, 7] = item.shiwuguanlileizhudongA57_H57;

                            ExcelSheet.Cells[59, 6] = item.zhengjuanleibeidongA59_F59;
                            ExcelSheet.Cells[59, 7] = item.zhengjuanleibeidongA59_H59;
                            ExcelSheet.Cells[60, 6] = item.guquanleibeidongA60_F60;
                            ExcelSheet.Cells[60, 7] = item.guquanleibeidongA60_H60;
                            ExcelSheet.Cells[61, 6] = item.rongzileibeidongA61_F61;
                            ExcelSheet.Cells[61, 7] = item.rongzileibeidongA61_H61;

                            ExcelSheet.Cells[62, 6] = item.qitaleibeidongA62;


                            ExcelSheet.Cells[41, 6] = item.zhengjuanleiA41;
                            ExcelSheet.Cells[42, 6] = item.guquanleiA42;
                            ExcelSheet.Cells[43, 4] = item.rongzileiA43;
                            ExcelSheet.Cells[44, 4] = item.qitaleiA44;

                            // 本年新增信托 	 2018年个数 	 2018年金额 	 占比 	 2017年个数 	 2017年金额 	 占比 
                            ExcelSheet.Cells[67, 5] = item.jieheA67;
                            ExcelSheet.Cells[68, 5] = item.jieheA67_F67;
                            ExcelSheet.Cells[69, 5] = item.danyiA68;
                            ExcelSheet.Cells[67, 6] = item.danyiA68_F68;
                            ExcelSheet.Cells[68, 6] = item.caichanquanA69;
                            //ExcelSheet.Cells[69, 6] = item.69;

                            ExcelSheet.Cells[71, 5] = item.zhudongguanliA71;
                            ExcelSheet.Cells[72, 6] = item.zhudongguanliA71_F71;

                            // 公司收入结构 	 2018年万元 	 占比 	 2017年万元 	 占比 	 2016年万元 	 占比 

                            ExcelSheet.Cells[75, 4] = item.shouxufeijiyongjinA75;
                            ExcelSheet.Cells[76, 4] = item.lixishouruA76;
                            ExcelSheet.Cells[77, 4] = item.touzishouyiA77;
                            ExcelSheet.Cells[78, 4] = item.qizhongguquanA78;
                            ExcelSheet.Cells[79, 4] = item.qizhongzhengjuanA79;
                            ExcelSheet.Cells[80, 4] = item.qizhongqitaA80;
                            ExcelSheet.Cells[81, 4] = item.gongyunjiazhibiandongshunyiA81;
                            ExcelSheet.Cells[83, 4] = item.shouruhejiA83;

                        }
                        #endregion
                    }
                    ExcelApp.Visible = true;
                    ExcelApp.ScreenUpdating = true;
                    ExcelBook.RefreshAll();
                    #region 写入文件
                    //sfdDownFile.FileName = Path.Combine(file, "Header rate " + farenvalue[i] + " " + DateTime.Now.ToString("yyyyMMdd")) + "利息";
                    //strExcelFileName = sfdDownFile.FileName + ".xlsx";


                    ExcelApp.ScreenUpdating = true;
                    //  ExcelBook.SaveAs(strExcelFileName, missingValue, missingValue, missingValue, missingValue, missingValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, missingValue, missingValue, missingValue, missingValue, missingValue);
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
                //clsKeyMyExcelProcess.Kill(ExcelApp);

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            #endregion

                    #endregion
        }

    }
}
