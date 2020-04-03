using kxrealtime.kxdata;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace kxrealtime
{
    [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public partial class utilDialog : Form
    {

        private string paperId;
        private string testId;
        private Timer myTimer;

        private WebBrowser infoWebPage;
        private Form infoForm;

        public utilDialog()
        {
            InitializeComponent();
            //this.utilsBtn.Top = (this.Height - this.utilsBtn.Height)/ 2;
            var tmp = utils.Utils.getScreenPosition();
            this.Location = tmp;
            //var otherLocation = utils.Utils.getScreenPosition(true);
            this.utilsBtn.Top = this.Height/2;
            
        }

        private void label1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void sendBtn_Click(object sender, EventArgs e)
        {
            var result = this.getExamInfo();
            if(result != null)
            {
                Int64 difTime = result.startTimeStamp + result.duringTime - utils.Utils.getTimeStamp();
                // clear old the info of exam
                if (result.noTime || difTime < 1000)
                {
                    Globals.ThisAddIn.removeExamItem(this.paperId);
                    this.paperId = null;
                    this.testId = null;
                }
            }
            
            showTimeChoice();
        }

        public void showSendBtn()
        {
            var tmp = utils.Utils.getScreenPosition();
            this.Location = tmp;
            checkMod();
            //isSendMod(false);
            this.utilsBtn.Top = this.Height / 2;
            this.Show();
        }

        public void onlyUtils()
        {
            examUtils.Visible = false;
            sendBtn.Visible = false;
            checkAns.Visible = false;
            this.utilsBtn.Top = this.Height / 2;
            this.Show();
        }

        private slideExamInfo getExamInfo()
        {
            var curIdx = Globals.ThisAddIn.PlaySlideIdx;
            var curSld = Globals.ThisAddIn.Application.ActivePresentation.Slides[curIdx];
            var curSldName = curSld.Name;
            return Globals.ThisAddIn.findExamInfo(curSldName);
        }

        public void checkMod()
        {
            this.paperId = null;
            var result = this.getExamInfo();
            isSendMod(result == null);
            if(result != null)
            {
                examUtils.Visible = true;
                this.paperId = result.paperId;
                this.testId = result.testId;
                examInfoHandle(result);
            }
            else
            {
                examUtils.Visible = false;
            }
        }

        public void examInfoHandle(slideExamInfo eData)
        {
            examUtils.Visible = false;
            if(eData.noTime)
            {
                timeLeft.Visible = false;
                delayBtn.Visible = false;
                stopBtn.Visible = false;
                examUtils.Visible = false;
                //sendBtn = true;
            } else
            {
                Int64 difTime = eData.startTimeStamp + eData.duringTime - utils.Utils.getTimeStamp();
                if (difTime < 0)
                {
                    this.stopMod();
                }
                else
                {
                    timeLeft.Visible = true;
                    startTiming(difTime);
                    delayBtn.Visible = true;
                    // todo
                    stopBtn.Visible = true;
                    examUtils.Visible = true;
                }
            }
        }

        private void stopMod()
        {
            timeLeft.Visible = true;
            delayBtn.Visible = false;
            stopBtn.Visible = false;
            examUtils.Visible = true;
            timeLeft.Text = "练习结束";
        }

        private void startTiming(Int64 leftTime)
        {
            if(this.myTimer != null)
            {
                this.myTimer.Stop();
                this.myTimer.Dispose();
                this.myTimer = null;
            }
            this.myTimer = new System.Windows.Forms.Timer();//实例化　
            myTimer.Tick += new EventHandler((s, e) =>
            {
                timeLeft.Text = s2Format(leftTime);
                leftTime -= 1000;
                if (leftTime < 0)
                {
                    timeLeft.Text = "练习结束";
                    myTimer.Stop();
                    this.myTimer.Dispose();
                    this.myTimer = null;
                    this.delayBtn.Visible = false;
                }
            }); //给timer挂起事件
            myTimer.Enabled = true;
            myTimer.Interval = 1000;
        }

        private string s2Format(Int64 time)
        {
            //Int64 h = time / 3600000;
            //time %= 3600000;
            Int64 m = time / 60000;
            time %= 60000;
            Int64 s = time / 1000;
            time %= 1000;
            return  m.ToString() + ":" + s.ToString();
        }

        public void isSendMod(bool flag)
        {
            var canShowCheck = curType();
            checkAns.Visible = !flag && canShowCheck;
            sendBtn.Visible = true;// flag;
        }

        public bool curType()
        {
            var curIdx = Globals.ThisAddIn.PlaySlideIdx;
            var curSld = Globals.ThisAddIn.Application.ActivePresentation.Slides[curIdx];
            var singleTitle = "kx-title-" + singleSelCtl.TypeSelEnum.singleSel;
            var mulitTitle = "kx-title-" + singleSelCtl.TypeSelEnum.multiSel;
            var singleVoteTitle = "kx-title-" + singleSelCtl.TypeSelEnum.voteSingleSel;
            var mulitVoteTitle = "kx-title-" + singleSelCtl.TypeSelEnum.voteMultiSel;
            foreach (PowerPoint.Shape shapeTmp in curSld.Shapes)
            {
                if(shapeTmp.Name.Contains("kx-title-"))
                {
                    return shapeTmp.Name == singleTitle || shapeTmp.Name == mulitTitle || shapeTmp.Name == singleVoteTitle || shapeTmp.Name == mulitVoteTitle;
                }
            }
            return false;
        }

        public void showTimeChoice()
        {
            var tmp = new choseTime(this);
            tmp.showFn(this.paperId, this.testId);
        }

        private void checkAns_Click(object sender, EventArgs e)
        {
            var webBrowser1 = new WebBrowser();
            //webBrowser1.Width = 800;
            //webBrowser1.Height = 500;
            //webBrowser1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            webBrowser1.Dock = DockStyle.Fill;
            var uriTmp = new Uri($"{utils.KXINFO.KXADMINURL}/?token={utils.KXINFO.KXTOKEN}#/pptComponents/countAnswerChart?aid={paperId}&token={utils.KXINFO.KXTOKEN}&testId={testId}");
            webBrowser1.Navigate(uriTmp);
            webBrowser1.Visible = true;

            var formTmp = new Form();
            formTmp.Width = 800;
            formTmp.Height = 500;
            formTmp.Controls.Add(webBrowser1);
            formTmp.Owner = this;
            formTmp.FormBorderStyle = FormBorderStyle.FixedSingle;
            formTmp.StartPosition = FormStartPosition.Manual;
            formTmp.Location = utils.Utils.getScreenPosition();
            formTmp.ShowIcon = false;
            formTmp.ShowInTaskbar = false;
            formTmp.Left += checkAns.Left - 700;
            formTmp.Top += checkAns.Top - 500;
            formTmp.TopMost = true;
            formTmp.Visible = true;
            //this.Controls.Add(webBrowser1);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //this.Close();
            this.Visible = false;
        }

        private void delayBtn_Click(object sender, EventArgs e)
        {
            this.showTimeChoice();
        }

        private void stopBtn_Click(object sender, EventArgs e)
        {
            var examInfo = this.getExamInfo();
            createExam(examInfo);
            sendStop(examInfo.paperId, examInfo.testId);
        }

        private void sendStop(string paperId, string testId)
        {
            object sendData = (new
            {
                key = "exam",
                type = "TEST_STOP",
                data = new
                {
                    paperId = paperId,
                    testId = testId
                },
                timestamp = utils.Utils.getTimeStamp()
            });
            JObject o = JObject.FromObject(sendData);
            string tmp = o.ToString();
            Globals.ThisAddIn.SendTchInfo(tmp);
        }

        private string createExam(slideExamInfo eInfo)
        {
            Uri reqUrl = new Uri($"{utils.KXINFO.KXCOURSEURL}/usr/api/upsertTest");
            Dictionary<string, string> args = new Dictionary<string, string> { };
            args.Add("token", utils.KXINFO.KXTOKEN);
            var curTime = utils.Utils.getTimeStamp();
            var postData = new createExamInfo()
            {
                aids = new List<string>() { paperId },
                owner = "30",
                multi = 100,
                start_time = eInfo.startTimeStamp,
                end_time = eInfo.startTimeStamp + 1000,
                title = eInfo.paperTitle,
                cost_time = eInfo.duringTime,
                id = testId
            };
            RestSharp.IRestResponse response = utils.request.SendRequest(Globals.ThisAddIn.CurHttpReq, reqUrl, RestSharp.Method.POST, args, postData);
            if (response.ErrorException != null)
            {
                utils.Utils.LOG("usr/api/upsertTest api error: " + response.ErrorException.Message);
                return null;
            }
            else
            {
                JObject data = JObject.Parse(response.Content);
                string code = (string)data["code"];
                if (code != "0")
                {
                    if (code == "401")
                    {
                        this.Visible = false;
                        Globals.ThisAddIn.loginOut();
                        MessageBox.Show("登录失效请重新登录");
                    }
                    utils.Utils.LOG("usr/api/upsertTest api error: code" + code);
                }
                else
                {
                    eInfo.duringTime = 0;
                    if (this.myTimer != null)
                    {
                        this.myTimer.Stop();
                        this.myTimer.Dispose();
                        this.myTimer = null;
                    }
                    this.stopMod();
                    return (string)data["data"]["id"];
                }
            }
            return null;
        }

        private void utilsBtn_Click(object sender, EventArgs e)
        {
            this.utilsBtn.Visible = false;
            this.utilsPanel.Left = this.utilsBtn.Left + this.utilsBtn.Width - this.utilsPanel.Width ;
            this.utilsPanel.Top = this.utilsBtn.Top - 100;
            this.utilsPanel.Visible = true;
        }

        private void utilDialog_Load(object sender, EventArgs e)
        {
            this.utilsBtn.Visible = true;
            this.utilsPanel.Visible = false;
        }

        private void createWebForm(string url)
        {
            if(this.infoForm != null)
            {
                this.infoForm.Dispose();
            }
            this.infoForm = new Form();
            this.infoForm.FormBorderStyle = FormBorderStyle.None;
            this.infoForm.StartPosition = FormStartPosition.Manual;
            this.infoForm.ShowIcon = false;
            this.infoForm.ShowInTaskbar = false;
            this.infoForm.WindowState = FormWindowState.Maximized;
            this.infoForm.TopMost = true;
            this.infoForm.Opacity = 0.9;
            this.infoForm.BackColor = System.Drawing.Color.AliceBlue;
            this.infoForm.Owner = this;
            this.infoForm.TransparencyKey = System.Drawing.Color.AliceBlue;
            this.infoForm.Location = utils.Utils.getScreenPosition();

            this.infoForm.KeyUp += InfoForm_KeyUp;

            if (this.infoWebPage != null)
            {
                this.infoWebPage.Dispose();
            }
            this.infoWebPage = new WebBrowser();
            this.infoWebPage.Url = new Uri(url);
            this.infoWebPage.Navigate(new Uri(url));
            this.infoWebPage.Visible = true;
            this.infoWebPage.Dock = DockStyle.Fill;
            this.infoWebPage.Refresh();
            this.infoWebPage.ObjectForScripting = this;


            this.infoForm.Controls.Add(this.infoWebPage);
            this.infoForm.Show();
        }

        private void InfoForm_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.CloseWin();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string divideGroup = $"{utils.KXINFO.KXADMINURL}/?timestamp={utils.Utils.getTimeStamp()}/#/pptComponents/group?teach_record_id={utils.KXINFO.KXTCHRECORDID}&session_id={utils.KXINFO.KXSID}";
            createWebForm(divideGroup);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string checkStudent = $"{utils.KXINFO.KXADMINURL}/?session_id={utils.KXINFO.KXSID}&timestamp={utils.Utils.getTimeStamp()}/#/pptComponents/rollcall?teach_record_id={utils.KXINFO.KXTCHRECORDID}&session_id={utils.KXINFO.KXSID}";
            createWebForm(checkStudent);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string checkQRCode = $"{utils.KXINFO.KXADMINURL}/?session_id={utils.KXINFO.KXSID}&timestamp={utils.Utils.getTimeStamp()}/#/pptComponents/signInQrcode?teach_record_id={utils.KXINFO.KXTCHRECORDID}&class_id={utils.KXINFO.KXCHOSECLASSID}&chapter_id={utils.KXINFO.KXCHOSECHAPTERID}&course_id={utils.KXINFO.KXCHOSECOURSEID}&title={utils.KXINFO.KXCHOSECOURSETITLE}&session_id={utils.KXINFO.KXSID}";
            createWebForm(checkQRCode);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string question = $"{utils.KXINFO.KXADMINURL}/?session_id={utils.KXINFO.KXSID}&timestamp={utils.Utils.getTimeStamp()}/#/pptComponents/nounderstand?teach_record_id={utils.KXINFO.KXTCHRECORDID}&session_id={utils.KXINFO.KXSID}";
            createWebForm(question);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string question = $"{utils.KXINFO.KXADMINURL}/?session_id={utils.KXINFO.KXSID}&timestamp={utils.Utils.getTimeStamp()}/#/pptComponents/studentContribute?teach_record_id={utils.KXINFO.KXTCHRECORDID}&session_id={utils.KXINFO.KXSID}";
            createWebForm(question);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            var courseTitle = utils.KXINFO.KXCHOSECOURSETITLE;
            string courseQRCode = $"{utils.KXINFO.KXADMINURL}/?session_id={utils.KXINFO.KXSID}&timestamp={utils.Utils.getTimeStamp()}/#/pptComponents/courseQrcode?teach_record_id={utils.KXINFO.KXTCHRECORDID}&class_id={utils.KXINFO.KXCHOSECLASSID}&chapter_id={utils.KXINFO.KXCHOSECHAPTERID}&course_id={utils.KXINFO.KXCHOSECOURSEID}&title={courseTitle}&session_id={utils.KXINFO.KXSID}  ";
            createWebForm(courseQRCode);
        }

        private void label2_Click(object sender, EventArgs e)
        {
            this.utilsBtn.Visible = true;
            this.utilsPanel.Visible = false;
        }

        public void CloseWin()
        {
            this.infoWebPage.Dispose();
            this.infoForm.Close();
            this.infoWebPage = null;
            this.infoForm = null;
        }

    }

    public class SendKxOutContent
    {
        public string paperId { get; set; }
        public string testId { get; set; }
    }

    public class SendKxOut
    {
        public Int64 teach_record_id {get;set;}
        public Int16 type { get; set; }
        public SendKxOutContent content { get; set; }
    }

    public class PaperC
    {
        public string text { get; set; }
    }

    public class PaperContent
    {
        public string t { get; set; }
        public PaperC c { get; set; }
    }

    public class PaperSel
    {
        public List<PaperContent> contents { get; set; }
    }
}
