using System;
using System.Collections.Generic;
using System.Linq;
//using System.Threading.Tasks;
using System.Windows.Forms;

namespace 测试功能专用
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
           //Application.Run(new frmExcel());
            Application.Run(new frm汇总多表());
        
          // Application.Run(new frmPDF());

           //Application.Run(new frmTB小青龙1028());
           //Application.Run(new frm横向2());
        }
    }
}
