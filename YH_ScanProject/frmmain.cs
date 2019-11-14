using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace YH_ScanProject
{
    public partial class frmmain : Form
    {
        private string Copyfile = "";
        public string path;
        public string exit_path;
        private List<string> Alist = new List<string>();
        string fullname; //文件路径+文件名，用于保存
        public frmmain()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
                OpenFileDialog1.Filter = "bmp,jpg,gif,png,tiff,icon|*.bmp;*.jpg;*.gif;*.png;*.tiff;*.icon";
                OpenFileDialog1.Title = "选择图片";
                if (OpenFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    fullname = OpenFileDialog1.FileName.ToString();
                }


                string ax = scanstep(fullname);

                this.textBox2.Text = ax.ToString();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: 01w22" + ex);
                throw;
            }

        }

        private static string scanstep(string fullname)
        {
            //string path = AppDomain.CurrentDomain.BaseDirectory + "1.jpg";
            //path = AppDomain.CurrentDomain.BaseDirectory + "Capture321.JPG";

            var ApiKey = "e375aac2fd624863b631ec5e45c81bdb";
            var SecretKey = "ecee0983771f451ab86ef6fea63b4847";
            var tuPian = fullname;

            var client = new Baidu.Aip.Ocr.Ocr(ApiKey, SecretKey);
            var image = File.ReadAllBytes(tuPian);

            // 通用文字识别
            var result = client.GeneralBasic(image, null);

            string ax = "";

            JsonTextReader reader = new JsonTextReader(new StringReader(result.ToString()));
            while (reader.Read())
            {
                if (reader.Value != null && reader.Value.ToString() != "words" && reader.Value.ToString() != "words_result" && reader.Value.ToString() != "words_result_num")
                {
                    //Console.WriteLine("Token: {0}, Value: {1}", reader.TokenType, reader.Value);
                    ax += "\r\n" + reader.Value;
                }
                //else
                //    Console.WriteLine("Token: {0}", reader.TokenType);
            }
            return ax;
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

            label7.Text = "已选中文件数：" + Alist.Count();

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

        private void button2_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog();
            dialog.Description = "请选择导出文件夹";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (string.IsNullOrEmpty(dialog.SelectedPath))
                {
                    MessageBox.Show(this, "文件夹路径不能为空", "提示");
                    return;
                }
                exit_path = dialog.SelectedPath;
                textBox3.Text = dialog.SelectedPath;


            }
            else
                return;
        }

        private void importButton_Click(object sender, EventArgs e)
        {
            try
            {
                for (int i = 0; i < Alist.Count; i++)
                {
                    //GetKEYnfo(path + "\\" + Alist[i]);
                    label7.Text = "正在执行：" + i + 1 + "/" + Alist.Count;

                    string ax = scanstep(Alist[i]);
                    string filename = System.IO.Path.GetFileName(Alist[i]);
                    downcsv(ax, filename);
                }

                MessageBox.Show("下载完成 ！", "System", MessageBoxButtons.OK, MessageBoxIcon.Information);
                label7.Text = "已完成";

                //downcsv(dataGridView);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: 0122" + ex);
                return;

                throw;
            }
        }

        public void downcsv(string ax, string filename)
        {
            string[] dataGridView = System.Text.RegularExpressions.Regex.Split(ax, "\r\n");

            if (dataGridView.Length == 0)
            {
                MessageBox.Show("Sorry , No Data Output !", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string strFileName = "";

            if (exit_path != "" && exit_path.Length > 0)
            {
                if (false == System.IO.Directory.Exists(exit_path))
                {
                    MessageBox.Show("路径不存在,请重新确认, 谢谢!");
                    return;
                }
               
                strFileName = exit_path + "\\Export  " + filename + " " + DateTime.Now.ToString("yyyyMMdd-ss") + ".csv";

            }
            else
            {
                string strDesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                strFileName = strDesktopPath + "\\SystemResults\\Export  " + filename + " " + DateTime.Now.ToString("yyyyMMdd-ss") + ".csv";

                string subPath = strDesktopPath + "/SystemResults/";
                if (false == System.IO.Directory.Exists(subPath))
                {
                    //创建pic文件夹
                    System.IO.Directory.CreateDirectory(subPath);
                }

            }

            FileStream fa = new FileStream(strFileName, FileMode.Create);
            StreamWriter sw = new StreamWriter(fa, Encoding.Unicode);
            string delimiter = "\r\n";
            string strHeader = "";

            {
                string strRowValue = "";

                for (int k = 0; k < dataGridView.Length; k++)
                {
                    if (dataGridView[k] != null)
                    {
                        strRowValue += dataGridView[k].ToString().Replace("\r\n", " ").Replace("\n", "") + delimiter;


                    }
                    else
                    {
                        strRowValue += dataGridView[k] + delimiter;
                    }
                }
                sw.WriteLine(strRowValue);
            }
            sw.Close();
            fa.Close();
        }

        private void closeButton_Click(object sender, EventArgs e)
        {
            this.Close();

        }


    }
}
