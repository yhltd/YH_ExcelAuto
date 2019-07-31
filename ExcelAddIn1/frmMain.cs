using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelAddIn1
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
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
    }
}
