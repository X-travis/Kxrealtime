using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Websocket.Client;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using CustomTaskPane =  Microsoft.Office.Tools.CustomTaskPane;

namespace kxrealtime
{
    public partial class Ribbon1
    {

        PowerPoint.Application app;

        private IWebsocketClient loginWebSocket = null;
        private PictureBox loginPictureBox;
        private loginDialog curLoginDialog;

        private choseClass curChoseForm;

        private Hashtable customTaskHash = new Hashtable();

        private singleSelCtl singleSelCtlInstance
        {
            get
            {
                var curWn = Globals.ThisAddIn.getCurActiveWn;
                var curKey = curWn + "-singleSelCtl";
                if (customTaskHash.ContainsKey(curKey))
                {
                    return customTaskHash[curKey] as singleSelCtl;
                }
                else
                {
                    return null;
                }
            }
            set
            {
                var curWn = Globals.ThisAddIn.getCurActiveWn;
                var curKey = curWn + "-singleSelCtl";
                if (!customTaskHash.ContainsKey(curKey))
                {
                    customTaskHash.Add(curKey, value);
                }
            }
        }

        private kxResource ksResourceCtl
        {
            get
            {
                var curWn = Globals.ThisAddIn.getCurActiveWn;
                var curKey = curWn + "-resourceCtl";
                if (customTaskHash.ContainsKey(curKey))
                {
                    return customTaskHash[curKey] as kxResource;
                }
                else
                {
                    return null;
                }
            }
            set
            {
                var curWn = Globals.ThisAddIn.getCurActiveWn;
                var curKey = curWn + "-resourceCtl";
                if (!customTaskHash.ContainsKey(curKey))
                {
                    customTaskHash.Add(curKey, value);
                }
            }
        }

        public CustomTaskPane myCustomTaskPane
        {
            get
            {
                var curWn = Globals.ThisAddIn.getCurActiveWn;
                var curKey = curWn + "-question";
                if(customTaskHash.ContainsKey(curKey))
                {
                    return customTaskHash[curKey] as CustomTaskPane;
                } else
                {
                    return null;
                }
            }
            set
            {
                var curWn = Globals.ThisAddIn.getCurActiveWn;
                var curKey = curWn + "-question";
                if (!customTaskHash.ContainsKey(curKey))
                {
                    customTaskHash.Add(curKey, value);
                }
            }
        }

        public CustomTaskPane kxResourceTaskPane
        {
            get
            {
                var curWn = Globals.ThisAddIn.getCurActiveWn;
                var curKey = curWn + "-resource";
                if (customTaskHash.ContainsKey(curKey))
                {
                    return customTaskHash[curKey] as CustomTaskPane;
                }
                else
                {
                    return null;
                }
            }
            set
            {
                var curWn = Globals.ThisAddIn.getCurActiveWn;
                var curKey = curWn + "-resource";
                if (!customTaskHash.ContainsKey(curKey))
                {
                    customTaskHash.Add(curKey, value);
                }
            }
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            app = Globals.ThisAddIn.Application;
        }

        private void initPanes(string title)
        {
            singleSelCtlInstance = new singleSelCtl();
            myCustomTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(singleSelCtlInstance, "编辑习题", app.ActiveWindow);
            myCustomTaskPane.Visible = true;
            myCustomTaskPane.Width = 380;
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            if (this.myCustomTaskPane == null)
            {
                this.initPanes("单选题");
            }
            else
            {
                this.myCustomTaskPane.Visible = true;
            }
            //this.createSingleCtx("单选题", singleSelCtl.TypeSelEnum.singleSel);
            utils.pptContent.createPaperItem("单选题", singleSelCtl.TypeSelEnum.singleSel);
            this.singleSelCtlInstance.setCurSelType = singleSelCtl.TypeSelEnum.singleSel;
            //char[] ans = new char[] { };
            //this.singleSelCtlInstance.resetData(0,ans,4);
        }

        public void resetSingleSel(PowerPoint.SlideRange slide)
        {
            System.Diagnostics.Debug.WriteLine("resetSingleSel");
            float curScore = 0;

            ArrayList ans = new ArrayList();
            ArrayList labelArr = new ArrayList();
            Hashtable shapeMap = new Hashtable();
            PowerPoint.Shapes curShapes = slide.Shapes;
            foreach (PowerPoint.Shape shapeTmp in curShapes)
            {
                /*if (shapeTmp.Name.Contains("kx-choice"))
                {
                    string tmp = shapeTmp.Name.Substring(10);
                    char curC = tmp[0];
                    labelArr.Add(curC);
                    if (shapeTmp.Fill.ForeColor.RGB == utils.KXINFO.ChoseColor)
                    {
                        ans.Add(curC);
                    }
                }*/
                if (shapeTmp.Name == "kx-score")
                {
                    string scoreTmp = shapeTmp.TextFrame.TextRange.Text;
                    curScore = float.Parse(scoreTmp.Replace('分', ' '));
                }
                shapeMap.Add(shapeTmp.Name, shapeTmp);
            }
            char sChar = 'A';
            char lastChar = 'A';
            for (int i = 0; i < 26; i++)
            {
                char curChar = (char)(sChar + i);
                var curKey = "kx-choice-" + curChar;
                var textKey = "kx-text-" + curChar;
                if (shapeMap.Contains(curKey))
                {
                    var curShape = shapeMap[curKey] as Shape;
                    if (lastChar != curChar)
                    {
                        curShape.Name = "kx-choice-" + lastChar;
                        curShape.TextFrame.TextRange.Text = lastChar.ToString();
                        if (shapeMap.Contains(textKey))
                        {
                            var textShap = shapeMap[textKey] as Shape;
                            textShap.Name = "kx-text-" + lastChar;
                        }
                    }
                    labelArr.Add(lastChar);
                    if (curShape.Fill.ForeColor.RGB == utils.KXINFO.ChoseColor)
                    {
                        ans.Add(lastChar);
                    }
                    lastChar = (char)(lastChar + 1);
                }
                else
                {
                    if (shapeMap.Contains(textKey))
                    {
                        Shape textTmp = shapeMap[textKey] as Shape;
                        textTmp.Delete();
                    }
                }
            }
            shapeMap.Clear();
            this.singleSelCtlInstance.resetData(curScore, ans, labelArr);
        }

        private void resetFillQustion(PowerPoint.SlideRange slide)
        {
            PowerPoint.Shapes curShapes = slide.Shapes;
            string ansStr = "";
            string questionStr = "";
            List<fillOption> ansTmp = new List<fillOption>();
            foreach (PowerPoint.Shape shapeTmp in curShapes)
            {
                if (shapeTmp.Name == "kx-question")
                {
                    questionStr = shapeTmp.TextFrame.TextRange.Text;
                }
                else if (shapeTmp.Name == "kx-qInfo")
                {
                    ansStr = shapeTmp.TextFrame.TextRange.Text;
                }
            }
            if (questionStr.Length > 0)
            {
                string pattern = @"(\[填空\d\])";
                var fillOptionCollect = Regex.Matches(questionStr, pattern);
                if (ansStr.Length > 0)
                {
                    var tmp = JsonConvert.DeserializeObject<List<fillOption>>(ansStr);
                    for (int i = 0; i < fillOptionCollect.Count; i++)
                    {
                        if (tmp.Count > i)
                        {
                            ansTmp.Add(tmp[i]);
                        }
                        else
                        {
                            ansTmp.Add(new fillOption());
                        }
                    }
                }

            }
            this.singleSelCtlInstance.resetFill(ansTmp);
        }


        public void checkSelExist(PowerPoint.SlideRange slide)
        {
            if (this.singleSelCtlInstance.setCurSelType == singleSelCtl.TypeSelEnum.singleSel || this.singleSelCtlInstance.setCurSelType == singleSelCtl.TypeSelEnum.multiSel || this.singleSelCtlInstance.setCurSelType == singleSelCtl.TypeSelEnum.voteMultiSel || this.singleSelCtlInstance.setCurSelType == singleSelCtl.TypeSelEnum.voteSingleSel)
            {
                System.Diagnostics.Debug.WriteLine("checkSelExist");
                resetSingleSel(slide);
            }
            else if (this.singleSelCtlInstance.setCurSelType == singleSelCtl.TypeSelEnum.fillQuestion)
            {
                resetFillQustion(slide);
            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            if (this.myCustomTaskPane == null)
            {
                this.initPanes("多选题");
            }
            else
            {
                this.myCustomTaskPane.Visible = true;
            }
            utils.pptContent.createPaperItem("多选题", singleSelCtl.TypeSelEnum.multiSel);
            this.singleSelCtlInstance.setCurSelType = singleSelCtl.TypeSelEnum.multiSel;
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            if (this.myCustomTaskPane == null)
            {
                this.initPanes("主观题");
            }
            else
            {
                this.myCustomTaskPane.Visible = true;
            }
            utils.pptContent.createPaperItem("主观题", singleSelCtl.TypeSelEnum.textQuestion);
            this.singleSelCtlInstance.setCurSelType = singleSelCtl.TypeSelEnum.textQuestion;
            this.singleSelCtlInstance.initSubjectiveQ(0);
        }


        // 重新检测ppt内容
        public void resestContent(PowerPoint.SlideRange slide)
        {
            if (this.myCustomTaskPane == null)
            {
                this.initPanes("");
            }
            PowerPoint.Shapes curShapes = slide.Shapes;
            bool isSel = false;
            bool isFill = false;
            bool isText = false;
            bool isVote = false;
            float score = 0;
            foreach (PowerPoint.Shape shapeTmp in curShapes)
            {
                string targetName = "kx-title-";
                if (shapeTmp.Name.Contains("kx-title"))
                {
                    isSel = shapeTmp.Name == (targetName + singleSelCtl.TypeSelEnum.singleSel) || shapeTmp.Name == (targetName + singleSelCtl.TypeSelEnum.multiSel);
                    isVote = shapeTmp.Name == (targetName + singleSelCtl.TypeSelEnum.voteSingleSel) || shapeTmp.Name == (targetName + singleSelCtl.TypeSelEnum.voteMultiSel);
                    isFill = shapeTmp.Name == (targetName + singleSelCtl.TypeSelEnum.fillQuestion);
                    isText = shapeTmp.Name == (targetName + singleSelCtl.TypeSelEnum.textQuestion);
                }

                if (shapeTmp.Name == "kx-score")
                {
                    string scoreTmp = shapeTmp.TextFrame.TextRange.Text;
                    score = float.Parse(scoreTmp.Replace('分', ' '));
                }
            }
            if (isSel || isVote)
            {
                this.resetSingleSel(slide);
            }
            else if (isFill)
            {
                this.resetFillQustion(slide);
            }
            else if (isText)
            {
                this.singleSelCtlInstance.setCurSelType = singleSelCtl.TypeSelEnum.textQuestion;
                this.singleSelCtlInstance.initSubjectiveQ(score,true);
            }
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            if (this.curLoginDialog != null)
            {
                this.curLoginDialog.Focus();
                return;
            }
            //Int32 curW = (Int32)app.ActivePresentation.SlideMaster.Width;
            //Int32 curH = (Int32)app.ActivePresentation.SlideMaster.Height;
            curLoginDialog = new loginDialog();
            curLoginDialog.getClose.Visible = false;
            curLoginDialog.getTitle.Visible = false;
            curLoginDialog.getLogo.Visible = false;
            string connectID = utils.Utils.createGUID();
            var loginContentTmp = curLoginDialog.getContent;
            loginContentTmp.Controls.Clear();
            loginPictureBox = getLoginQR(connectID);
            curLoginDialog.Width = 500;
            curLoginDialog.Height = 700;
            curLoginDialog.BackColor = System.Drawing.Color.Black;
            loginDialog.Show(curLoginDialog, System.Drawing.Color.Empty, 0.01);
            curLoginDialog.getClose.Visible = true;
            curLoginDialog.getTitle.Visible = true;
            curLoginDialog.getLogo.Visible = true;
            Label tipText = new Label();
            tipText.Visible = true;
            tipText.Text = "使用微信扫码登录";
            tipText.Top = 400;
            tipText.Left = 100;
            tipText.Width = 300;
            tipText.Height = 50;
            tipText.ForeColor = System.Drawing.Color.White;
            tipText.Font = new System.Drawing.Font(tipText.Font.Name, 14F);
            tipText.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            loginContentTmp.Controls.Add(loginPictureBox);
            loginContentTmp.Controls.Add(tipText);
            initLoginListener(connectID);
        }

        private PictureBox getLoginQR(string connectID)
        {
            string connectUrl = $"{utils.KXINFO.KXURL}/mp/#/user?client_id={connectID}";// $"http://kx-v010-wap.dev.resfair.com/#/?client_id={connectID}";
            // create qrcode
            string qrStr = QRcode.CreateQRCodeToBase64(connectUrl, false);
            byte[] imgBytes = Convert.FromBase64String(qrStr);
            System.Drawing.Image image;
            using (MemoryStream ms = new MemoryStream(imgBytes))
            {
                image = System.Drawing.Image.FromStream(ms);
            }
            PictureBox pictureBox = new PictureBox();
            pictureBox.Image = image;
            pictureBox.Width = 300;
            pictureBox.Height = 300;
            pictureBox.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox.Left = 100;
            pictureBox.Top = 50;
            pictureBox.BackColor = System.Drawing.Color.Black;
            pictureBox.Visible = true;
            return pictureBox;
        }

        private void initLoginListener(string curID)
        {
            if (loginWebSocket != null)
            {
                //loginWebSocket.Stop(System.Net.WebSockets.WebSocketCloseStatus.NormalClosure, "close");
                this.closeLoginConnect();
            }
            loginWebSocket = utils.webSocketClient.StartWebSocket($"{utils.KXINFO.KXSOCKETURL}/mobileLogin?client_id={curID}");
            loginWebSocket.MessageReceived.Subscribe(msg =>
            {
                try
                {
                    if (loginWebSocket == null)
                    {
                        return;
                    }

                    string msgStr = msg.Text;
                    utils.KXINFO.initUsr(msgStr);
                    var startTch = initTchInfo(msgStr);
                    // close login pane
                    Action<bool> actionDelegate = this.LoginSuccess;
                    // 或者
                    // Action<string> actionDelegate = delegate(string txt) { this.label2.Text = txt; };
                    this.curLoginDialog.Invoke(actionDelegate, startTch);

                    this.button5.Visible = false;
                    this.menu1.Visible = true;
                    this.menu1.Label = utils.KXINFO.KXUNAME;
                    this.resourceGroup.Visible = true;
                    this.closeLoginConnect();
                    if(ksResourceCtl != null && !ksResourceCtl.IsDisposed)
                    {
                        ksResourceCtl.Dispose();
                        ksResourceCtl = null;
                        kxResourceTaskPane.Visible = false;
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show("登录失败" + e.Message);
                }
            });
        }

        private bool initTchInfo(string dataInfo)
        {
            try
            {
                JObject data = JObject.Parse(dataInfo);
                var tchList = data["teach_record_list"];
                var tchInfo = tchList[0];
                if (tchList == null || tchInfo == null)
                {
                    return false;
                }
                var classId = (string)tchInfo["class_id"];
                var courseId = (string)tchInfo["course_id"];
                var chapterId = (string)tchInfo["chapter_id"];
                var className = (string)data["class_info"][classId]["name"];
                var courseName = (string)data["course_info"][courseId]["title"];
                var chapterName = (string)data["chapter_info"][chapterId]["title"];
                utils.KXINFO.KXCHOSECLASSID = (Int64)tchInfo["class_id"];
                utils.KXINFO.KXCHOSECOURSEID = (Int64)tchInfo["course_id"];
                utils.KXINFO.KXCHOSECHAPTERID = (Int64)tchInfo["chapter_id"];
                utils.KXINFO.KXCHOSECOURSETITLE = courseName;
                utils.KXINFO.KXCHOSECLASSNAME = className;
                utils.KXINFO.KXTCHRECORDID = (string)tchInfo["tid"];
                utils.KXINFO.KXCHOSECHAPTERTITLE = chapterName;
                return true;
            }catch(Exception)
            {

            }
            return false;
        }


        public void closeLoginConnect()
        {
            if (this.loginWebSocket != null)
            {
                //this.loginWebSocket.Stop(System.Net.WebSockets.WebSocketCloseStatus.NormalClosure, "close");

                this.loginWebSocket.Dispose();
                this.loginWebSocket = null;
            }
            this.curLoginDialog = null;
        }

        private void CloseLoginDialog()
        {
            if (loginDialog.frmBack != null)
            {
                loginDialog.frmBack.Close();
            }
            if (this.curLoginDialog != null)
            {
                this.curLoginDialog.Close();
                this.curLoginDialog = null;
            }

            // need rechose the class info
            ChangeTchBtn(false);
            Globals.ThisAddIn.CloseTchSocket();
        }

        private void LoginSuccess(bool startTch)
        {
            var url = utils.KXINFO.KXUAVATAR;
            PictureBox pictureBox = new PictureBox();
            try
            {
                System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12 | System.Net.SecurityProtocolType.Tls11 | System.Net.SecurityProtocolType.Tls;
                pictureBox.Load(url);
                menu1.Image = pictureBox.Image;
            }
            catch (WebException)
            {
                //utils.Utils.LOG("loginsuccess load url error： " + e.Message);
            }

            pictureBox.Width = 300;
            pictureBox.Height = 300;
            pictureBox.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox.Left = 100;
            pictureBox.Top = 50;
            pictureBox.BackColor = System.Drawing.Color.Black;
            pictureBox.Visible = true;
            Label tipText = new Label();
            tipText.Visible = true;
            tipText.Text = "登录成功,3秒自动关闭";
            tipText.Top = 400;
            tipText.Left = 100;
            tipText.Width = 300;
            tipText.Height = 50;
            tipText.ForeColor = System.Drawing.Color.White;
            tipText.Font = new System.Drawing.Font(tipText.Font.Name, 14F);
            tipText.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;

            Panel loginContentTmp = curLoginDialog.getContent;
            loginContentTmp.Controls.Clear();

            loginContentTmp.Controls.Add(pictureBox);
            loginContentTmp.Controls.Add(tipText);

            int timeLeft = 2;
            System.Windows.Forms.Timer myTimer = new System.Windows.Forms.Timer();//实例化　
            myTimer.Tick += new EventHandler((s, e) =>
            {
                string textTmp = $"登录成功,{timeLeft}秒自动关闭";
                tipText.Text = textTmp;
                timeLeft--;
                if (timeLeft < 0)
                {
                    myTimer.Stop();
                    this.CloseLoginDialog();
                    if (startTch)
                    {
                        MessageBoxButtons messbutton = MessageBoxButtons.OKCancel;
                        DialogResult dr = MessageBox.Show($"是否需要继续上次授课（班级：{utils.KXINFO.KXCHOSECLASSNAME}    课程：{utils.KXINFO.KXCHOSECOURSETITLE}）", "温馨提示", messbutton);
                        if (dr == DialogResult.OK)
                        {
                            curChoseForm = new choseClass();
                            ChangeTchBtn(true);
                            Globals.ThisAddIn.InitTchSocket();
                        }
                    }
                }
            }); //给timer挂起事件
            myTimer.Enabled = true;
            myTimer.Interval = 1000;
        }

        public void settingChange(bool isSetting)
        {
            for (int i = 1; i <= app.ActivePresentation.Slides.Count; i++)
            {
                PowerPoint.Slide curSld = app.ActivePresentation.Slides[i];
                //bool isKxItem = curSld.Name.Contains("kx-slide");
                //if (!isKxItem)
                //{
                //    continue;
                //}
                ArrayList curAnswerArr = (ArrayList)AnswerStore.getAnswer(curSld.Name);
                if (curAnswerArr == null)
                {
                    curAnswerArr = new ArrayList();
                }
                if(!isSetting)
                {
                    curAnswerArr.Clear();
                }
                PowerPoint.Shapes curShapes = curSld.Shapes;
                foreach (PowerPoint.Shape shapeTmp in curShapes)
                {
                    string targetName = "kx-setting";
                    if (shapeTmp.Name == targetName)
                    {
                        bool isKxItem = curSld.Name.Contains("kx-slide");
                        if(!isKxItem)
                        {
                            curSld.Name = "kx-slide-" + curSld.Name;
                        }
                        shapeTmp.Visible = isSetting ? Office.MsoTriState.msoCTrue : Office.MsoTriState.msoFalse;
                    }
                    targetName = "kx-sending";
                    if (shapeTmp.Name == targetName)
                    {
                        shapeTmp.Visible = isSetting ? Office.MsoTriState.msoCTrue : Office.MsoTriState.msoFalse;
                    }
                    if (shapeTmp.Name.Contains("kx-choice"))
                    {
                        string tmp = shapeTmp.Name.Substring(10);
                        if (tmp.Length == 0)
                        {
                            continue;
                        }
                        char curC = tmp[0];
                        if (isSetting)
                        {
                            if (curAnswerArr.Contains(curC))
                            {
                                shapeTmp.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(1, 0, 255, 0).ToArgb();
                            }
                        }
                        else
                        {
                            if (shapeTmp.Fill.ForeColor.RGB == utils.KXINFO.ChoseColor)
                            {
                                curAnswerArr.Add(curC);
                            }
                            shapeTmp.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(1, 128, 128, 128).ToArgb();
                        }
                    }

                }
                if (!isSetting)
                {
                    AnswerStore.setAnswer(curSld.Name, curAnswerArr);
                }
            }
        }

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {
            if (utils.KXINFO.KXUID == null)
            {
                MessageBox.Show("请先登录");
                return;
            }
            curChoseForm = new choseClass();
            curChoseForm.Visible = true;
            //curChoseForm.Location = utils.Utils.getScreenPosition(true);
            curChoseForm.initClassList();
            curChoseForm.initCourseList();
        }

        // relogin
        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            button5_Click(sender, e);
        }

        // loginOut
        private void button8_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.loginOut();
            this.button5.Visible = true;
            this.menu1.Visible = false;
            this.menu1.Label = "";
            ChangeTchBtn(false);
            if(kxResourceTaskPane != null)
            {
                kxResourceTaskPane.Visible = false;
            }
        }

        // play slide
        private void button9_Click(object sender, RibbonControlEventArgs e)
        {
            //Globals.ThisAddIn.Application.ActivePresentation.SlideShowSettings.RangeType = PpSlideShowRangeType.ppShowAll;
            //Globals.ThisAddIn.Application.ActivePresentation.SlideShowSettings.StartingSlide = 1;
            //Globals.ThisAddIn.Application.ActivePresentation.SlideShowSettings.ShowWithNarration = Office.MsoTriState.msoTrue;
            Globals.ThisAddIn.Application.ActivePresentation.SlideShowSettings.Run();
        }

        // close couse
        private void button10_Click(object sender, RibbonControlEventArgs e)
        {
            this.stopTch();
        }

        public void stopTch()
        {
            object sendData = (new
            {
                key = "classroom",
                value = utils.KXINFO.KXCHOSECLASSID,
                type = "COURSE_END",
                data = new { },
                teach_record_id = utils.KXINFO.KXTCHRECORDID,
                timestamp = utils.Utils.getTimeStamp()
            });
            JObject o = JObject.FromObject(sendData);
            string tmp = o.ToString();
            Globals.ThisAddIn.SendTchInfo(tmp);
            try
            {
                curChoseForm.stopTching();
                Globals.ThisAddIn.CloseTchSocket();
                Globals.ThisAddIn.kxSlideExam.Clear();
                ChangeTchBtn(false);
                curChoseForm.Close();
                utils.KXINFO.tchClear();
            }
            catch (Exception e) {
                MessageBox.Show("结束授课失败");
            }
        }

        public void ChangeTchBtn(bool tching)
        {
            this.box1.Visible = tching;
            this.button6.Visible = !tching;
        }

        private void button11_Click(object sender, RibbonControlEventArgs e)
        {
            if (this.myCustomTaskPane == null)
            {
                this.initPanes("投票");
            }
            else
            {
                this.myCustomTaskPane.Visible = true;
            }
            utils.pptContent.createPaperItem("投票", singleSelCtl.TypeSelEnum.voteSingleSel);
            this.singleSelCtlInstance.setCurSelType = singleSelCtl.TypeSelEnum.voteSingleSel;
            //this.singleSelCtlInstance.initVoteQ(singleSelCtl.TypeSelEnum.voteSingleSel);
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            if (this.myCustomTaskPane == null)
            {
                this.initPanes("填空题");
            }
            else
            {
                this.myCustomTaskPane.Visible = true;
            }
            utils.pptContent.createPaperItem("填空题", singleSelCtl.TypeSelEnum.fillQuestion,"此处插入描述", 0);
            this.singleSelCtlInstance.setCurSelType = singleSelCtl.TypeSelEnum.fillQuestion;
            this.singleSelCtlInstance.initFillQ(0);
        }

        private void button12_Click(object sender, RibbonControlEventArgs e)
        {
            var curIdx = Globals.ThisAddIn.CurSlideIdx;
            //Globals.ThisAddIn.Application.ActivePresentation.SlideShowSettings.RangeType = PpSlideShowRangeType.ppShowSlideRange;
            //Globals.ThisAddIn.Application.ActivePresentation.SlideShowSettings.StartingSlide = curIdx;
            //Globals.ThisAddIn.Application.ActivePresentation.SlideShowSettings.ShowWithNarration = Office.MsoTriState.msoFalse;
            var showWin = Globals.ThisAddIn.Application.ActivePresentation.SlideShowSettings.Run();
            showWin.View.GotoSlide(curIdx);
        }

        private void resourceBtn_Click(object sender, RibbonControlEventArgs e)
        {
            if (utils.KXINFO.KXUID == null)
            {
                MessageBox.Show("请先登录");
                return;
            }
            if (ksResourceCtl == null)
            {
                ksResourceCtl = new kxResource();
                kxResourceTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(ksResourceCtl, "资源库");
                kxResourceTaskPane.Visible = true;
                kxResourceTaskPane.Width = 500;
            }
            kxResourceTaskPane.Visible = true;
        }
    }
}
