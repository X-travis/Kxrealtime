using RestSharp;
using System;
using System.Drawing;
using System.IO;
using System.Net;
using System.Windows.Forms;

namespace kxrealtime.utils
{
    public delegate void ProgressTip(double value);
    public static class Utils
    {
        public static Form downloadForm;

        // 创建唯一id
        public static string createGUID()
        {
            Random R = new Random();
            string strDateTimeNumber = DateTime.Now.ToString("yyyyMMddHHmmssms");
            string strRandomResult = R.Next(1, 1000).ToString();
            return strDateTimeNumber + strRandomResult;
        }

        // 截屏
        public static byte[] getScreenImg()
        {
            //创建图象，保存将来截取的图象
            foreach (var curScreen in Screen.AllScreens)
            {

            }
            var primaryScreenArea = Screen.PrimaryScreen.Bounds;
            Bitmap image = new Bitmap(primaryScreenArea.Width, primaryScreenArea.Height);
            Graphics imgGraphics = Graphics.FromImage(image);
            //设置截屏区域
            int xTmp = primaryScreenArea.Left;
            int yTmp = primaryScreenArea.Top;
            int wTmp = primaryScreenArea.Width;
            int hTmp = primaryScreenArea.Height;
            imgGraphics.CopyFromScreen(xTmp, yTmp, xTmp, yTmp, new System.Drawing.Size(wTmp, hTmp));
            ImageConverter converter = new ImageConverter();
            return (byte[])converter.ConvertTo(image, typeof(byte[]));
        }

        //获取屏幕信息
        public static System.Drawing.Point getScreenPosition(bool isPrimary = false)
        {
            System.Drawing.Point curPoint = new System.Drawing.Point(Screen.PrimaryScreen.Bounds.Left, Screen.PrimaryScreen.Bounds.Top);
            if (isPrimary)
            {
                return curPoint;
            }
            foreach (var curScreen in Screen.AllScreens)
            {
                if (!curScreen.Primary)
                {
                    curPoint = new System.Drawing.Point(curScreen.Bounds.Left, curScreen.Bounds.Top);
                }
            }
            return curPoint;
        }

        // log函数
        public static void LOG(object e)
        {
            System.Windows.Forms.MessageBox.Show(e.ToString());
            System.Diagnostics.Debug.WriteLine(e);
        }

        // 获取对应的时间戳
        public static long getTimeStamp()
        {
            System.DateTime startTime = TimeZone.CurrentTimeZone.ToLocalTime(new System.DateTime(1970, 1, 1)); // 当地时区
            long timeStamp = (long)(DateTime.Now - startTime).TotalMilliseconds;
            if(KXINFO.srvTimeDif != 0)
            {
                timeStamp += KXINFO.srvTimeDif;
            }
            return timeStamp;
        }

        // 下载文件 （弃用）
        public static void dlFile(string fileUrl, string savePath)
        {
            try
            {
                var writer = File.OpenWrite(savePath);
                var client = new RestClient();
                var request = new RestRequest(fileUrl);
                request.ResponseWriter = responseStream =>
                {
                    using (responseStream)
                    {
                        responseStream.CopyTo(writer);
                    }
                };
                client.DownloadData(request);
                writer.Close();
                writer.Dispose();
            }catch(Exception e)
            {
                Utils.LOG(e.Message);
            }
        }

        // 获取文件路径
        public static string getFilePath()
        {
            string curDir = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);
            var fileDict = curDir + @"\kxrealtime\files";
            if (!Directory.Exists(fileDict))
            {
                try
                {
                    Directory.CreateDirectory(fileDict);
                }
                catch (Exception e)
                {
                    MessageBox.Show("创建文件失败" + e.Message);
                }

            }
            return fileDict;
        }

        // 下载文件（以字节的方式）
        public static bool dlFileOrigin(string url, string path, string contentType, ProgressTip cb)
        {
            try
            {
                HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
                req.ServicePoint.Expect100Continue = false;
                req.Method = "GET";
                req.KeepAlive = true;
                //req.ContentType = contentType;// "image/png";
                using (HttpWebResponse rsp = (HttpWebResponse)req.GetResponse())
                {
                    long totalBytes = rsp.ContentLength;
                    using (Stream reader = rsp.GetResponseStream())
                    {
                        using (FileStream writer = new FileStream(path, FileMode.OpenOrCreate, FileAccess.Write))
                        {
                            cb(0);
                            byte[] buff = new byte[512];
                            int c = 0; //实际读取的字节数
                            double readBuff = 0;
                            while ((c = reader.Read(buff, 0, buff.Length)) > 0)
                            {
                                readBuff += buff.Length;
                                double curPG = (double)readBuff / (double)totalBytes;
                                cb(curPG);
                                writer.Write(buff, 0, c);
                            }
                        }
                    }
                }
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }
    }
}
