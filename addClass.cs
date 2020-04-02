using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace kxrealtime
{
    
    [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    //[System.Runtime.InteropServices.ComVisible(true)]
    public partial class addClass : Form
    {
        private WebBrowser curWebBrowser;

        public addClass(string classId, string className)
        {
            InitializeComponent();
            //initAddClassQR(classId, className);
            initWebPage();
        }

        private void initWebPage()
        {
            var timeStamp = utils.Utils.getTimeStamp();
            var url = $"{utils.KXINFO.KXADMINURL}/?timestamp={timeStamp}#/pptComponents/startTeach";
            this.curWebBrowser = new WebBrowser();
            this.curWebBrowser.Navigate(new Uri(url));
            this.curWebBrowser.Visible = true;
            this.curWebBrowser.Dock = DockStyle.Fill;
            this.curWebBrowser.Refresh();
            this.curWebBrowser.DocumentCompleted += WebBrowser1_DocumentCompleted;
            
            this.Controls.Add(this.curWebBrowser);
            this.curWebBrowser.ObjectForScripting = this;
        }

        private void WebBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            this.sendData();
        }

        private void sendData()
        {
            HtmlDocument curDoc = this.curWebBrowser.Document;
            object sendData = (new
            {
                user_id = utils.KXINFO.KXOUTUID,
                class_id = utils.KXINFO.KXCHOSECLASSID,
                course_id = utils.KXINFO.KXCHOSECOURSEID,
                chapter_id = utils.KXINFO.KXCHOSECHAPTERID,
                session_id = utils.KXINFO.KXSID,
                nickName = utils.KXINFO.KXUNAME,
                course_title = utils.KXINFO.KXCHOSECOURSETITLE,
                class_name = utils.KXINFO.KXCHOSECLASSNAME,
                avatar = utils.KXINFO.KXUAVATAR
            });
            JObject o = JObject.FromObject(sendData);
            string tmp = o.ToString();
            curDoc.InvokeScript("setPageData", new[]
            {
                tmp
            });
        }

        private void initAddClassQR(string classId, string className)
        {
            /*foreach (var curScreen in Screen.AllScreens)
            {
                if(!curScreen.Primary)
                {
                    this.Location = new Point(curScreen.Bounds.Left, curScreen.Bounds.Top);
                }
            }*/
            this.Location = utils.Utils.getScreenPosition();
            string connectUrl = $"{utils.KXINFO.KXURL}/mp/#/user/join-class?class_id={classId}&class_name={className}";
            // create qrcode
            string qrStr = QRcode.CreateQRCodeToBase64(connectUrl, false);
            byte[] imgBytes = Convert.FromBase64String(qrStr);
            System.Drawing.Image image;
            using (MemoryStream ms = new MemoryStream(imgBytes))
            {
                image = System.Drawing.Image.FromStream(ms);
            }
            var imageSize = image.Size;
            //pictureBox2.Image = image;
            //pictureBox2.Size = imageSize;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.CloseWin();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            this.Dispose();
        }

        public void CloseWin()
        {
            this.curWebBrowser.Dispose();
            this.Visible = false;
            this.Close();
        }

        public void getData()
        {
            this.sendData();
        }
    }
}
