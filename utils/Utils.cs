using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace kxrealtime.utils
{
    public static class Utils
    {
        public static string createGUID()
        {
            Random R = new Random();
            string strDateTimeNumber = DateTime.Now.ToString("yyyyMMddHHmmssms");
            string strRandomResult = R.Next(1, 1000).ToString();
           return strDateTimeNumber + strRandomResult;
        }

        public static byte[] getScreenImg()
        {
            //创建图象，保存将来截取的图象
            foreach (var curScreen in Screen.AllScreens) {
                
            }
            var primaryScreenArea = Screen.PrimaryScreen.Bounds;
            Bitmap image = new Bitmap(primaryScreenArea.Width, primaryScreenArea.Height);
            Graphics imgGraphics = Graphics.FromImage(image);
            //设置截屏区域
            int xTmp = primaryScreenArea.Left;
            int yTmp = primaryScreenArea.Top;
            int wTmp = primaryScreenArea.Width;
            int hTmp = primaryScreenArea.Height;
            imgGraphics.CopyFromScreen(xTmp, yTmp, xTmp, yTmp, new Size(wTmp, hTmp));
            ImageConverter converter = new ImageConverter();
            return (byte[])converter.ConvertTo(image, typeof(byte[]));
        }

        public static Point getScreenPosition(bool isPrimary = false)
        {
            Point curPoint = new Point(Screen.PrimaryScreen.Bounds.Left, Screen.PrimaryScreen.Bounds.Top);
            if(isPrimary)
            {
                return curPoint;
            }
            foreach (var curScreen in Screen.AllScreens)
            {
                if (!curScreen.Primary)
                {
                    curPoint = new Point(curScreen.Bounds.Left, curScreen.Bounds.Top);
                }
            }
            return curPoint;
        }

        public static void LOG(object e)
        {
            MessageBox.Show(e.ToString());
            System.Diagnostics.Debug.WriteLine(e);
        }

        public static long getTimeStamp()
        {
            System.DateTime startTime = TimeZone.CurrentTimeZone.ToLocalTime(new System.DateTime(1970, 1, 1)); // 当地时区
            long timeStamp = (long)(DateTime.Now - startTime).TotalMilliseconds;
            return timeStamp;
        }
    }
}
