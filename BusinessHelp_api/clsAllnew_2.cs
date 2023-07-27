using Newtonsoft.Json;
using RestSharp.Contrib;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BusinessHelp_api
{
    public class clsAllnew_2
    {
        private int pu_type_name;


        #region 百度图片API
        public void diaoyongfangfa(int type_name)
        {
            try
            {
                pu_type_name = type_name;

                string token = "24.55d3e57d41ced3ca940944da7d0463df.2592000.1692931980.282335-36658435";


                List<string> filename = new List<string>();
                List<string> filename_mingcheng = new List<string>();
                OpenFileDialog tbox = new OpenFileDialog();
                tbox.Multiselect = false;
                tbox.Filter = "所有文件|*.*";
                tbox.Multiselect = true;
                tbox.SupportMultiDottedExtensions = true;
                if (tbox.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    int i = 0;

                    foreach (string s in tbox.SafeFileNames)
                    {
                        filename.Add(tbox.FileNames[i]);
                        filename_mingcheng.Add(s);
                        i++;
                    }
                }
                //循环文件一 一处理
                for (int j = 0; j < filename.Count; j++)
                {
                    string result = "";

                    //  apibaidu();
                    //DehazeDemo();

                    if (type_name == 1)
                        result = image_definition_enhance(filename[j], token);     // 图像清晰度增强

                    if (type_name == 2)
                        result = image_quality_enhance(filename[j], token);      // 图像无损放大
                    if (type_name == 3)
                        result = contrast_enhance(filename[j], token);   // 图像对比度增强
                    if (type_name == 5)
                        result = inpainting(filename[j], token);   // 图像修复
                    if (type_name == 6)
                        result = colorEnhance(filename[j], token);   // 图像色彩增强
                    if (type_name == 7)
                        result = docRepair(filename[j], token);   // 文档图片去底纹
                    if (type_name == 8)
                        result = quzao_demo(filename[j], token);   // 文档图片去底纹
                    if (type_name == 9)
                        result = quwu_dehaze(filename[j], token);   // 图像去雾


                    //如果没有子文件夹创建子文件夹
                    string newpaht = filename[j].Replace(filename_mingcheng[j], "") + "output";

                    if (false == System.IO.Directory.Exists(newpaht))
                    {
                        //创建pic文件夹
                        System.IO.Directory.CreateDirectory(newpaht);
                    }


                    Base64StringToImage(result, newpaht + "\\" + filename_mingcheng[j]);
                }
                MessageBox.Show("转换完成" + filename[0].Replace(filename_mingcheng[0], "") + "output");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "ex");
                return;

                throw;
            }
        }

        private static void apibaidu()
        {
            // 设置APPID/AK/SK
            var APP_ID = "36658435";
            var API_KEY = "CRZpbTYeK5EAufaAS70ZCO8d";
            var SECRET_KEY = "IDOigyq3SXqsG30zboHAeDDnqAP9zhzS";

            var client = new Baidu.Aip.ImageProcess.ImageProcess(API_KEY, SECRET_KEY);
            client.Timeout = 60000;  // 修改超时时间

            string url1 = AppDomain.CurrentDomain.BaseDirectory + "test\\9e6b1b7b-50fe-4369-b226-129e311ffe6c.jpg";


            var image = File.ReadAllBytes(url1);
            // 图像清晰增强，可能会抛出网络等异常，请使用try/catch捕获
            var result = client.ImageDefinitionEnhance(image);
            Console.WriteLine(result);



            // 文件url
            var url = "http://host/test.jpeg";
            // url = AppDomain.CurrentDomain.BaseDirectory + "test\\9e6b1b7b-50fe-4369-b226-129e311ffe6c.jpg";

            result = client.ImageDefinitionEnhanceUrl(url);

        }
        //public void ImageQualityEnhanceDemo()
        //{
        //    string path1 = AppDomain.CurrentDomain.BaseDirectory + "test\\9e6b1b7b-50fe-4369-b226-129e311ffe6c.jpg";

        //    //var client = NewMethod1();

        //    var client = NewMethod();


        //    var image = File.ReadAllBytes(path1);
        //    // 调用图像无损放大，可能会抛出网络等异常，请使用try/catch捕获
        //    var result = client.ImageQualityEnhance(image);
        //    Console.WriteLine(result);
        //}

        //private static Baidu.Aip.ImageProcess.ImageProcess NewMethod()
        //{
        //    // 设置APPID/AK/SK
        //    var APP_ID = "36658435";
        //    var API_KEY = "CRZpbTYeK5EAufaAS70ZCO8d";
        //    var SECRET_KEY = "IDOigyq3SXqsG30zboHAeDDnqAP9zhzS";

        //    var client = new Baidu.Aip.ImageProcess.ImageProcess(API_KEY, SECRET_KEY);
        //    client.Timeout = 60000;  // 修改超时时间

        //    return client;
        //}



        public class Person
        {
            public string log_id { get; set; }
            public string image { get; set; }
            public string result { get; set; }


        }

        //base64编码的字符串转为图片 
        private Bitmap Base64StringToImage(string strbase64, string nemos)
        {
            try
            {

                var person = JsonConvert.DeserializeObject<Person>(strbase64);
                if (pu_type_name == 7 || pu_type_name == 8)
                    strbase64 = person.result;
                else
                    strbase64 = person.image;

                strbase64 = strbase64.Replace("data:image/png;base64,", "").Replace("data:image/jgp;base64,", "").Replace("data:image/jpg;base64,", "").Replace("data:image/jpeg;base64,", "");//将base64头部信息替换
                strbase64 = strbase64.Replace("data:image/png;base64,", "");
                byte[] arr = Convert.FromBase64String(strbase64);
                MemoryStream ms = new MemoryStream(arr);
                Bitmap bmp = new Bitmap(ms);
                //string nemos = AppDomain.CurrentDomain.BaseDirectory + "test\\1.jpg";
                bmp.Save(nemos, System.Drawing.Imaging.ImageFormat.Jpeg);
                //bmp.Save(@"d:\"test.bmp", ImageFormat.Bmp);

                //bmp.Save(@"d:\"test.gif", ImageFormat.Gif);

                //bmp.Save(@"d:\"test.png", ImageFormat.Png);
                ms.Close();
                return bmp;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "ex");
                return null;

            }
        }
        public void DehazeDemo()
        {
            // 设置APPID/AK/SK
            var APP_ID = "36658435";
            var API_KEY = "CRZpbTYeK5EAufaAS70ZCO8d";
            var SECRET_KEY = "IDOigyq3SXqsG30zboHAeDDnqAP9zhzS";

            var client = new Baidu.Aip.ImageProcess.ImageProcess(API_KEY, SECRET_KEY);
            client.Timeout = 60000;  // 修改超时时间

            string path1 = AppDomain.CurrentDomain.BaseDirectory + "test\\9e6b1b7b-50fe-4369-b226-129e311ffe6c.jpg";

            var image = File.ReadAllBytes(path1);
            // 调用图像去雾，可能会抛出网络等异常，请使用try/catch捕获
            var result = client.Dehaze(image);
            Console.WriteLine(result);
        }
        // 图像清晰度增强
        public static string image_definition_enhance(string path1, string token)
        {
            //string token = "[调用鉴权接口获取的token]";//24.55d3e57d41ced3ca940944da7d0463df.2592000.1692931980.282335-36658435
            //string token = "24.55d3e57d41ced3ca940944da7d0463df.2592000.1692931980.282335-36658435";
            string host = "https://aip.baidubce.com/rest/2.0/image-process/v1/image_definition_enhance?access_token=" + token;
            Encoding encoding = Encoding.Default;
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(host);
            request.Method = "post";
            request.KeepAlive = true;
            // 图片的base64编码
            //string base64 = getFileBase64("[本地图片文件]");
            //string path1 = AppDomain.CurrentDomain.BaseDirectory + "test\\9e6b1b7b-50fe-4369-b226-129e311ffe6c.jpg";

            string base64 = getFileBase64(path1);
            String str = "image=" + HttpUtility.UrlEncode(base64);
            byte[] buffer = encoding.GetBytes(str);
            request.ContentLength = buffer.Length;
            request.GetRequestStream().Write(buffer, 0, buffer.Length);
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.Default);
            string result = reader.ReadToEnd();
            Console.WriteLine("图像清晰度增强:");
            //Console.WriteLine(result);


            return result;
        }
        // 图像无损放大
        public static string image_quality_enhance(string path1, string token)
        {
            //string token = "[调用鉴权接口获取的token]";
            //string token = "24.55d3e57d41ced3ca940944da7d0463df.2592000.1692931980.282335-36658435";

            string host = "https://aip.baidubce.com/rest/2.0/image-process/v1/image_quality_enhance?access_token=" + token;
            Encoding encoding = Encoding.Default;
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(host);
            request.Method = "post";
            request.KeepAlive = true;
            // 图片的base64编码
            //string path1 = AppDomain.CurrentDomain.BaseDirectory + "test\\9e6b1b7b-50fe-4369-b226-129e311ffe6c.jpg";

            string base64 = getFileBase64(path1);
            String str = "image=" + HttpUtility.UrlEncode(base64);
            byte[] buffer = encoding.GetBytes(str);
            request.ContentLength = buffer.Length;
            request.GetRequestStream().Write(buffer, 0, buffer.Length);
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.Default);
            string result = reader.ReadToEnd();
            Console.WriteLine("图像无损放大:");
            //Console.WriteLine(result);
            return result;
        }
        // 图像对比度增强
        public static string contrast_enhance(string path1, string token)
        {
            //string token = "24.55d3e57d41ced3ca940944da7d0463df.2592000.1692931980.282335-36658435";

            string host = "https://aip.baidubce.com/rest/2.0/image-process/v1/contrast_enhance?access_token=" + token;
            Encoding encoding = Encoding.Default;
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(host);
            request.Method = "post";
            request.KeepAlive = true;
            // 图片的base64编码
            //string path1 = AppDomain.CurrentDomain.BaseDirectory + "test\\9e6b1b7b-50fe-4369-b226-129e311ffe6c.jpg";

            string base64 = getFileBase64(path1);
            String str = "image=" + HttpUtility.UrlEncode(base64);
            byte[] buffer = encoding.GetBytes(str);
            request.ContentLength = buffer.Length;
            request.GetRequestStream().Write(buffer, 0, buffer.Length);
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.Default);
            string result = reader.ReadToEnd();
            Console.WriteLine("图像对比度增强:");
            //Console.WriteLine(result);
            return result;
        }
        // 图像修复
        public static string inpainting(string path1, string token)
        {
            //string token = "[调用鉴权接口获取的token]";
            string host = "https://aip.baidubce.com/rest/2.0/image-process/v1/inpainting?access_token=" + token;
            Encoding encoding = Encoding.Default;
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(host);
            request.Method = "post";
            request.KeepAlive = true;
            //String str = "{\"rectangle\":[{\"width\":92,\"top\":95,\"height\":36,\"left\":543}],\"image\":\"图片base64编码\"}";
            string base64 = getFileBase64(path1);

            String str = "{\"rectangle\":[{\"width\":92,\"top\":95,\"height\":36,\"left\":543}],\"image\":\"" + base64 + "\"}";
            byte[] buffer = encoding.GetBytes(str);
            request.ContentLength = buffer.Length;
            request.GetRequestStream().Write(buffer, 0, buffer.Length);
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.Default);
            string result = reader.ReadToEnd();
            //Console.WriteLine("图像修复:");
            //  Console.WriteLine(result);
            return result;
        }
        // 图像色彩增强
        public static string colorEnhance(string path1, string token)
        {
            //string token = "[调用鉴权接口获取的token]";
            string host = "https://aip.baidubce.com/rest/2.0/image-process/v1/color_enhance?access_token=" + token;
            Encoding encoding = Encoding.Default;
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(host);
            request.Method = "post";
            request.KeepAlive = true;
            // 图片的base64编码
            string base64 = getFileBase64(path1);
            String str = "image=" + HttpUtility.UrlEncode(base64);
            byte[] buffer = encoding.GetBytes(str);
            request.ContentLength = buffer.Length;
            request.GetRequestStream().Write(buffer, 0, buffer.Length);
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.Default);
            string result = reader.ReadToEnd();
            //Console.WriteLine("图像色彩增强:");
            //Console.WriteLine(result);
            return result;
        }
        // 文档图片去底纹
        public static string docRepair(string path1, string token)
        {
            //string token = "[调用鉴权接口获取的token]";
            string host = "https://aip.baidubce.com/rest/2.0/image-process/v1/doc_repair?access_token=" + token;
            Encoding encoding = Encoding.Default;
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(host);
            request.Method = "post";
            request.KeepAlive = true;
            // 图片的base64编码
            string base64 = getFileBase64(path1);
            String str = "image=" + HttpUtility.UrlEncode(base64);
            byte[] buffer = encoding.GetBytes(str);
            request.ContentLength = buffer.Length;
            request.GetRequestStream().Write(buffer, 0, buffer.Length);
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.Default);
            string result = reader.ReadToEnd();
            //Console.WriteLine("文档图片去底纹:");
            //Console.WriteLine(result);
            return result;
        }

        // 图像去噪
        public static string quzao_demo(string path1, string token)
        {
            //string token = "[调用鉴权接口获取的token]";
            string host = "https://aip.baidubce.com/rest/2.0/image-process/v1/denoise?access_token=" + token;
            Encoding encoding = Encoding.Default;
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(host);
            request.Method = "post";
            request.KeepAlive = true;
            // 图片的base64编码
            string base64 = getFileBase64(path1);
            String str = "option=" + "0" + "&image=" + HttpUtility.UrlEncode(base64);
            byte[] buffer = encoding.GetBytes(str);
            request.ContentLength = buffer.Length;
            request.GetRequestStream().Write(buffer, 0, buffer.Length);
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.Default);
            string result = reader.ReadToEnd();
            //Console.WriteLine("图像去噪:");
            //Console.WriteLine(result);
            return result;
        }
        public static String getFileBase64(String fileName)
        {
            FileStream filestream = new FileStream(fileName, FileMode.Open);
            byte[] arr = new byte[filestream.Length];
            filestream.Read(arr, 0, (int)filestream.Length);
            string baser64 = Convert.ToBase64String(arr);
            filestream.Close();
            return baser64;
        }

        // 图像去雾
        public static string quwu_dehaze(string path1, string token)
        {
            //string token = "[调用鉴权接口获取的token]";
            string host = "https://aip.baidubce.com/rest/2.0/image-process/v1/dehaze?access_token=" + token;
            Encoding encoding = Encoding.Default;
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(host);
            request.Method = "post";
            request.KeepAlive = true;
            // 图片的base64编码
            string base64 = getFileBase64(path1);
            String str = "image=" + HttpUtility.UrlEncode(base64);
            byte[] buffer = encoding.GetBytes(str);
            request.ContentLength = buffer.Length;
            request.GetRequestStream().Write(buffer, 0, buffer.Length);
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.Default);
            string result = reader.ReadToEnd();
            //Console.WriteLine("图像去雾:");
            //Console.WriteLine(result);
            return result;
        }
        #endregion


    }
}
