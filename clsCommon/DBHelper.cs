using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;

namespace clsCommon
{
   public class DBHelper
    {
        private OleDbConnection conn;
        private OleDbDataAdapter oda = new OleDbDataAdapter();
        private OleDbCommand cmd;
        private DataSet myds = new DataSet();
        public DBHelper()
        {
            //
            // TODO: 在此处添加构造函数逻辑
            //
            //conn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath + @"/db/dbtest.mdb");
        }
        //得到config中链接数据库字符串
        public OleDbConnection get_conn()
        {
            OleDbConnection conn;
            conn = new OleDbConnection(System.Configuration.ConfigurationManager.ConnectionStrings["conn"].ToString());
            return conn;
        }

        public DataSet getDS(string strSQL)
        {
            conn = new OleDbConnection(System.Configuration.ConfigurationManager.ConnectionStrings["conn"].ToString());
            myds = new DataSet();
            oda = new OleDbDataAdapter(strSQL, conn);
            oda.Fill(myds);
            return myds;
        }
        //查询
        public DataSet getDS2(string strSQL, int si, int mi)
        {
            conn = new OleDbConnection(System.Configuration.ConfigurationManager.ConnectionStrings["conn"].ToString());
            conn.Open();
            myds = new DataSet();
            oda = new OleDbDataAdapter(strSQL, conn);
            oda.Fill(myds, si, mi, "tab1");
            conn.Close();
            return myds;

        }

        public bool setDS(string strSQL)
        {
            conn = new OleDbConnection(System.Configuration.ConfigurationManager.ConnectionStrings["conn"].ToString());
            conn.Open();
            cmd = new OleDbCommand(strSQL, conn);
            cmd.ExecuteNonQuery();
            conn.Close();
            return true;
        }
        //添加 删除 修改
        public int add(string strSQL)
        {
            conn = new OleDbConnection(System.Configuration.ConfigurationManager.ConnectionStrings["conn"].ToString());
            conn.Open();
            cmd = new OleDbCommand(strSQL, conn);
            int k = cmd.ExecuteNonQuery();
            conn.Close();
            return k;
        }

    }
}
