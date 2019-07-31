using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.Odbc;

namespace clsCommon
{
    public class clsServerID
    {
        public string dbid;

        public string ConStr = "Provider=Microsoft.Jet.OleDb.4.0;Data Source=\\\\9.112.114.167\\db\\ElandSalseDailyReport.mdb;";
        //private string mdbpath = "\\\\IBM536-PC065SDE\\db\\ElandSalseDailyReport.mdb;";
        public string mdbpath = "\\\\9.112.114.167\\db\\ElandSalseDailyReport.mdb;";

        public string ConStr2 = "Provider=Microsoft.Jet.OleDb.4.0;Data Source=\\\\9.112.114.167\\db\\ElandSalseDailyReport_Backup.mdb;";//记录 Status  click 和选择哪个服务器

        public string mdbpath2 = "\\\\9.112.114.167\\db\\ElandSalseDailyReport_Backup.mdb;";//记录 Status  click 和选择哪个服务器

        //切分后的服务器地址
        public string ConStr_S = "Provider=Microsoft.Jet.OleDb.4.0;Data Source=\\\\9.112.114.167\\db\\copy\\Split\\ElandSalseDailyReport1.mdb;";//ElandSalseDailyReport1_be_DailyReport
        public string mdbpath_S = "\\\\9.112.114.167\\db\\copy\\Split\\ElandSalseDailyReport1.mdb;";//ElandSalseDailyReport1_be_DailyReport

        //ctrix server path
        //public string ConStr_Ctirx = "Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "data\\ElandSalseDailyReport1_be_DailyReport.mdb;";
        //public string mdbpath_Ctirx = AppDomain.CurrentDomain.BaseDirectory + "data\\ElandSalseDailyReport1_be_DailyReport.mdb;";
        //2017 10 20 实验拿到本地 （拆分后的带箭头表）看看速度 //ElandSalseDailyReport1 为测试系统，ElandSalseDailyReport_qianduan 为正式数据库

        public string ConStr_Ctirx = "Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "data\\System\\ElandSalseDailyReport_qianduan.mdb;";//ElandSalseDailyReport_qianduan  ElandSalseDailyReport1
        public string mdbpath_Ctirx = AppDomain.CurrentDomain.BaseDirectory + "data\\System\\ElandSalseDailyReport_qianduan.mdb;"; //ElandSalseDailyReport_qianduan   ElandSalseDailyReport1
        //Internetbank 
        public string ConStr_Ctirx_Internetbank = "Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "data\\System\\ElandSalseInternetbank_qianduan.mdb;";//ElandSalseDailyReport_qianduan  ElandSalseDailyReport1
        public string mdbpath_Ctirx_Internetbank = AppDomain.CurrentDomain.BaseDirectory + "data\\System\\ElandSalseInternetbank_qianduan.mdb;"; //ElandSalseDailyReport_qianduan   ElandSalseDailyReport1
        //Internetbank all
        public string ConStr_Ctirx_InternetbankAll = "Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "data\\System\\ElandSalseInternetbankAll_qianduan.mdb;";//ElandSalseDailyReport_qianduan  ElandSalseDailyReport1
        public string mdbpath_Ctirx_InternetbankAll = AppDomain.CurrentDomain.BaseDirectory + "data\\System\\ElandSalseInternetbankAll_qianduan.mdb;"; //ElandSalseDailyReport_qianduan   ElandSalseDailyReport1

        //银行的总表
        public string Internetbank_allpath = "[;database=" + AppDomain.CurrentDomain.BaseDirectory + "data\\System\\ElandSalseInternetbankAll_qianduan.mdb].";

        //new 方法
        //public string ConStr_Ctirx = "Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "data\\System\\ElandSalseDailyReport1.mdb;Persist Security Info=False";
        // 黄网IP 地址
        //  \\172.18.6.26\test\db

        //ctrix server path记录点击 的事件
        public string ConStr2_Ctirx = "Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "data\\ElandSalseDailyReport_Backup.mdb;";//记录 Status  click 和选择哪个服务器
        public string mdbpath2_Ctirx = AppDomain.CurrentDomain.BaseDirectory + "data\\ElandSalseDailyReport_Backup.mdb;";//记录 Status  click 和选择哪个服务器 ElandSalseDailyReport_Backup
        //

        //string conStr = "DSN=TestEland;UID=test;PWD=!QAZ2wsx";
        //OdbcConnection odbcCon = new OdbcConnection()
        //sql link 
        // public string ConStr_sql = @"Provider=SQLOLEDB;server=127.0.0.1;uid=sa;pwd=Lyh07910;database=TestEland"; //本地自己的数据库
        public string ConStr_sql = @"Provider=SQLOLEDB;server=172.18.6.26,1433;uid=Test;pwd=!QAZ2wsx;database=TestEland"; //本地自己的数据库



        public void ServerID()
        {

            dbid = "";
            // dbid = Read_DBportsAll();

            Local_IP();
            ////临时替换
         //   dbid = "5";//XXXXXX
            ////sql 数据库
            //  dbid = "NewDB";//sql 数据库
            ////
            ////dbid = "2";//access 
            //    dbid = "4";//AP ctrix server 


            ServerIP();
        }

        private void Local_IP()
        {
            string A_Path = AppDomain.CurrentDomain.BaseDirectory + "data\\IP.txt";
            string[] fileText = File.ReadAllLines(A_Path);
            if (fileText.Length > 0 && fileText[0] != null && fileText[0] != "")
            {
                if (fileText[0] != null && fileText[0] != "")
                    dbid = fileText[0];
                if (dbid == "NewDB")
                    ConStr_sql = @fileText[1]; //本地自己的数据库
                //Provider=SQLOLEDB;server=172.18.6.26,1433;uid=Test;pwd=!QAZ2wsx;database=TestEland
                //Provider=SQLOLEDB;server=172.18.6.26,1433;uid=Test;pwd=!QAZ2wsx;database=TestEland
            }
        }

        private void ServerIP()
        {
            if (dbid == "2")//切分后的 数据  此方法 调整地方为   isrun = GetTables(aConnection, "DailyReport"); 默认都返回 True
            {
                //ConStr = "Provider=Microsoft.Jet.OleDb.4.0;Data Source=\\\\9.112.114.167\\db\\copy\\Split\\ElandSalseDailyReport1.mdb;";
                //mdbpath = "\\\\9.112.114.167\\db\\copy\\Split\\ElandSalseDailyReport.mdb;";
                ConStr = ConStr_S;
                mdbpath = mdbpath_S;
                ConStr2 = ConStr2;
                mdbpath2 = mdbpath2;
                //避免读去拆分后的表
                ConStr_Ctirx_Internetbank = "";
                mdbpath_Ctirx_Internetbank = "";
                ConStr_Ctirx_InternetbankAll = "";
                mdbpath_Ctirx_InternetbankAll = "";
            }
            else if (dbid == "1")//切分前的 数据
            {
                ConStr = ConStr;// "Provider=Microsoft.Jet.OleDb.4.0;Data Source=\\\\9.112.114.167\\db\\ElandSalseDailyReport.mdb;";
                mdbpath = mdbpath;// "\\\\9.112.114.167\\db\\ElandSalseDailyReport.mdb;";

                ConStr2 = ConStr2;
                mdbpath2 = mdbpath2;
                //避免读去拆分后的表
                ConStr_Ctirx_Internetbank = "";
                mdbpath_Ctirx_Internetbank = "";
                ConStr_Ctirx_InternetbankAll = "";
                mdbpath_Ctirx_InternetbankAll = "";
            }
            else if (dbid == "3")//读取本地配置文件的信息
            {
                string A_Path = AppDomain.CurrentDomain.BaseDirectory + "data\\IP2.txt";
                string[] fileText = File.ReadAllLines(A_Path);
                if (fileText.Length > 2 && fileText[0] != null && fileText[0] != "")
                {
                    ConStr = fileText[1];
                    mdbpath = fileText[2];
                    ConStr2 = fileText[3];
                    mdbpath2 = fileText[4];
                    ConStr_Ctirx_Internetbank = fileText[5];
                    mdbpath_Ctirx_Internetbank = fileText[6];
                    ConStr_Ctirx_InternetbankAll = fileText[7];
                    mdbpath_Ctirx_InternetbankAll = fileText[8];
                }
            }
            else if (dbid == "4")//ctrix server path
            {
                ConStr = ConStr_Ctirx;
                mdbpath = mdbpath_Ctirx;
                ConStr2 = ConStr2_Ctirx;
                mdbpath2 = mdbpath2_Ctirx;

                //避免读去拆分后的表
                ConStr_Ctirx_Internetbank = "";
                mdbpath_Ctirx_Internetbank = "";

                ConStr_Ctirx_InternetbankAll = "";
                mdbpath_Ctirx_InternetbankAll = "";
            }
            else if (dbid == "5")//ctrix server path &拆分网银信息
            {
                ConStr = ConStr_Ctirx;
                mdbpath = mdbpath_Ctirx;
                ConStr2 = ConStr2_Ctirx;
                mdbpath2 = mdbpath2_Ctirx;
                ConStr_Ctirx_Internetbank = ConStr_Ctirx_Internetbank;
                mdbpath_Ctirx_Internetbank = mdbpath_Ctirx_Internetbank;
                Internetbank_allpath = Internetbank_allpath;
                ConStr_Ctirx_InternetbankAll = ConStr_Ctirx_InternetbankAll;
                mdbpath_Ctirx_InternetbankAll = mdbpath_Ctirx_InternetbankAll;
            }
            else if (dbid == "NewDB")//sql automation IT server
            {
                ConStr = "NewDB";
                ConStr = ConStr_sql;
                mdbpath = ConStr_sql;
                ConStr2 = ConStr_sql;
                mdbpath2 = ConStr_sql;
                ConStr_Ctirx_Internetbank = ConStr_sql;
                mdbpath_Ctirx_Internetbank = ConStr_sql;
                Internetbank_allpath = ConStr_sql;
                ConStr_Ctirx_InternetbankAll = ConStr_sql;
                mdbpath_Ctirx_InternetbankAll = ConStr_sql;
            }

        }
        //
    }
}
