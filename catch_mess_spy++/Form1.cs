using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace catch_mess_spy__
{
    public partial class Form1 : Form
    {     
        #region Import API
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("User32.dll", EntryPoint = "SendMessage")]
        private static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        [DllImport("user32.dll")]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll", EntryPoint = "GetParent")]
        public static extern IntPtr GetParent(IntPtr hwndChild);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern bool IsWindowVisible(IntPtr hwnd);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
        #endregion

        #region Property
        private TextBox txtPW;
        private RegistryKey key;
        private RegistryKey User;
        private RegistryKey PassWord;
        private bool LoginStatus = false;

        private int ScreenStatus, intCnt;
        private IntPtr hwnd_main, hwnd_ReportTree, hwnd_ReportTree1, hwnd_Control;
        private System.Timers.Timer t = new System.Timers.Timer(5000);//实例化Timer类，设置间隔时间为10000毫秒； 
        private const int WM_KEYDOWN = 0x100;
        private const int WM_KEYUP = 0x101;
        private const int VK_TAB = 0x9;
        private const int VK_CONTROL = 0x11;
        private const int VK_PRIOR = 0x21;
        private const int VK_UP = 0x26;
        private const int VK_HOME = 0x24;
        private const int BM_CLICK = 0xF5;
        private const int WM_LBUTTONDOWN = 0x0201;
        private const int WM_LBUTTONUP = 0x0202;
        private const int SYSKEYDOWN = 0x104;
        private const int WM_SETTEXT = 0x000C;

  
        
        private bool bolExist01 = false, bolExist0205 = false;
        private string Wbs_ID = "";
        private string FileName = "";
        private bool SAPBoolStatus = false;
 
        private string Plan_Cost_90_300 = string.Empty;
        private string Plan_Cost_90_0 = string.Empty;
        #endregion


        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }
        #region Windows API Function
        private List<IntPtr> getSAPWindow()
        {
            IntPtr hwnd_sap = IntPtr.Zero;
            List<IntPtr> arrHwnd = new List<IntPtr>();

            while (true)
            {
                hwnd_sap = FindWindowEx(IntPtr.Zero, hwnd_sap, "SAP_FRONTEND_SESSION", null);
                if (hwnd_sap != IntPtr.Zero)
                {
                    arrHwnd.Add(hwnd_sap);
                }
                else
                {
                    break;
                }
            }

            return arrHwnd;
        }

        private void monitorSAP()
        {
            if (hwnd_main == IntPtr.Zero)
            {
                //MessageBox.Show("找不到SAP主窗口！");
                return;
            }

            if (IsWindowVisible(hwnd_main))
            {
                if (ScreenStatus == 0)
                {
                    hwnd_ReportTree1 = findReportTree1();
                    //MessageBox.Show(hwnd_ReportTree1.ToInt32().ToString());
                    selectReport(intCnt);

                    IntPtr btnExecute = findExecuteButton();
                    SendMessage(btnExecute, BM_CLICK, 0, 0);
                    ScreenStatus = 1;
                }
                //clickYes(btnExecute);
            }
        }

        private IntPtr findExecuteButton()
        {
            IntPtr children = FindWindowEx(hwnd_main, IntPtr.Zero, null, "");
            while (children != IntPtr.Zero)
            {
                children = FindWindowEx(hwnd_main, children, null, "");
                int nRet;
                StringBuilder ClassName = new StringBuilder(100);
                //Get the window class name
                nRet = GetClassName(children, ClassName, ClassName.Capacity);
                Regex r = new Regex("Afx:[(a-z)|(A-Z)|(0-9)]{8}:8:[0-9]{8}:00000000:00000000");
                if (nRet != 0 && r.Match(ClassName.ToString()).Success)
                {
                    IntPtr hwnd_level2 = FindWindowEx(children, IntPtr.Zero, "Button", null);
                    if (hwnd_level2 != IntPtr.Zero)
                    {
                        return hwnd_level2;
                    }
                }
            }
            return IntPtr.Zero;
        }

        private void clickYes(IntPtr hwnd_Control)
        {
            IntPtr hwnd_Button = FindWindowEx(hwnd_Control, new IntPtr(0), "Button", null);
            SendMessage(hwnd_Button, BM_CLICK, 0, 0);
        }

        private void selectReport(int intCount)
        {
            sendPageUp();
            sendTab(intCount);
        }

        private void sendPageUp()
        {
            for (int i = 0; i < 5; i++)
            {
                SendMessage(hwnd_ReportTree1, WM_KEYDOWN, VK_CONTROL, 0);
                SendMessage(hwnd_ReportTree1, WM_KEYDOWN, VK_PRIOR, 0);
                SendMessage(hwnd_ReportTree1, WM_KEYUP, VK_PRIOR, 0);
                SendMessage(hwnd_ReportTree1, WM_KEYUP, VK_CONTROL, 0);
            }
        }

        private void sendTab(int intCount)
        {
            //Tab
            for (int i = 0; i < intCount; i++)
            {
                SendMessage(hwnd_ReportTree1, WM_KEYDOWN, VK_TAB, 0);
                SendMessage(hwnd_ReportTree1, WM_KEYUP, VK_TAB, 0);
            }
        }

        private IntPtr findReportTree1()
        {
            IntPtr hwnd_level1, hwnd_level2, hwnd_level3;
            hwnd_level1 = FindWindowEx(hwnd_main, IntPtr.Zero, "Docking Container Class", null);
            //MessageBox.Show("Docking Container Class:" + hwnd_level1.ToInt32().ToString());
            if (hwnd_level1 != IntPtr.Zero)
            {
                hwnd_level2 = FindWindowEx(hwnd_level1, IntPtr.Zero, "Shell Window Class", "Control  Container");
                //MessageBox.Show("Control  Container:" + hwnd_level2.ToInt32().ToString());
                if (hwnd_level2 != IntPtr.Zero)
                {
                    hwnd_level3 = FindWindowEx(hwnd_level2, IntPtr.Zero, "AfxOleControl80", null);
                    if (hwnd_level3 == IntPtr.Zero)
                        hwnd_level3 = FindWindowEx(hwnd_level2, IntPtr.Zero, "AfxOleControl90", null);
                    //MessageBox.Show("AfxOleControl80:" + hwnd_level3.ToInt32().ToString());
                    if (hwnd_level3 != IntPtr.Zero)
                    {
                        //MessageBox.Show("SAPTreeList:" + FindWindowEx(hwnd_level3, IntPtr.Zero, "SAPTreeList", "SAP's Advanced Treelist").ToInt32().ToString());
                        return FindWindowEx(hwnd_level3, IntPtr.Zero, "SAPTreeList", "SAP's Advanced Treelist");
                        //return FindWindowEx(hwnd_level2, IntPtr.Zero, "SAPTreeList", "SAP's Advanced Treelist");
                    }
                }
            }
            return IntPtr.Zero;
        }

        private IntPtr findControlWindow()
        {
            string strCaption = "Execute Project Report: Initial Screen";

            IntPtr hwnd_Child = IntPtr.Zero;
            while (true)
            {
                hwnd_Child = FindWindowEx(IntPtr.Zero, hwnd_Child, "#32770", strCaption);
                if (GetParent(hwnd_Child) == hwnd_main || hwnd_Child == IntPtr.Zero)
                {
                    break;
                }
            }
            return hwnd_Child;
        }
        #endregion

    }
}
