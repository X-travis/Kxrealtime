using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Collections;
using System.IO;
using RestSharp;
using Newtonsoft.Json.Linq;
using System.Windows.Forms;
using Websocket.Client;
using System.Net;
using Microsoft.Office.Interop.PowerPoint;
using Newtonsoft.Json;
using System.Text.RegularExpressions;

namespace kxrealtime
{
    public partial class Ribbon1
    {

        PowerPoint.Application app;
        private singleSelCtl singleSelCtlInstance;
        public Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
        public Microsoft.Office.Tools.CustomTaskPane loginPane;
        private loginCtl loginCtlInstance;
        private IWebsocketClient loginWebSocket;
        private PictureBox loginPictureBox;

        private loginDialog curLoginDialog;

        private choseClass curChoseForm;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            app = Globals.ThisAddIn.Application;
        }

        private void initPanes(string title)
        {
            singleSelCtlInstance = new singleSelCtl();
            myCustomTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(singleSelCtlInstance, "编辑习题");
            myCustomTaskPane.Visible = true;
            myCustomTaskPane.Width = 380;
        }

        private void createSingleCtx(string titleName, singleSelCtl.TypeSelEnum questionType)
        {
            PowerPoint.CustomLayout ppt_layout = app.ActivePresentation.SlideMaster.CustomLayouts[PowerPoint.PpSlideLayout.ppLayoutText];
            PowerPoint.Slide slide;
            int curSld = app.ActivePresentation.Slides.Count;
            slide = app.ActivePresentation.Slides.AddSlide(curSld + 1, ppt_layout);
            slide.Select();
            if(slide.Shapes.Count > 0)
            {
                slide.Shapes[1].Delete();
                slide.Shapes.Placeholders[1].Delete();
            }
            
            slide.Name = "kx-slide-" + slide.Name;

            Int32 curW = (Int32)app.ActivePresentation.SlideMaster.Width;
            Int32 curH = (Int32)app.ActivePresentation.SlideMaster.Height;

            PowerPoint.Shape sendBtn = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeActionButtonCustom, curW - 150, curH - 60, 100, 40);
            sendBtn.TextFrame.TextRange.InsertAfter("发送题目");
            sendBtn.Name = "kx-sending";
            sendBtn.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(1, 170, 170, 170).ToArgb();
            sendBtn.Line.ForeColor.RGB = System.Drawing.Color.FromArgb(1, 170, 170, 170).ToArgb();

            //sendBtn.TextFrame.TextRange.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick].Action = PowerPoint.PpActionType.ppActionRunMacro;
            //sendBtn.TextFrame.TextRange.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick].Run = "createText";


            // 题干
            PowerPoint.Shape textBoxTitle = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal, 100, 100, curW - 120, 100);
            textBoxTitle.TextFrame.TextRange.InsertAfter("此处插入描述");
            textBoxTitle.Name = "kx-question";
            textBoxTitle.Height = 80;
  

            // 题干额外信息
            PowerPoint.Shape qInfo = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal, -100, -100, curW - 120, 400);
            qInfo.Name = "kx-qInfo";
            qInfo.Visible = Office.MsoTriState.msoFalse;

            // 题目类型
            PowerPoint.Shape titleCom = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, curW, 30);
            titleCom.TextFrame.TextRange.InsertAfter(titleName);
            titleCom.Name = "kx-title-" + questionType;

            // 不是投票
            PowerPoint.Shape scoreCom = null;
            if (questionType != singleSelCtl.TypeSelEnum.voteSingleSel && questionType != singleSelCtl.TypeSelEnum.voteMultiSel)
            {
                // 分数
                scoreCom = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal, 100, 0, 100, 30);
                scoreCom.TextFrame.TextRange.InsertAfter("10分");
                scoreCom.Name = "kx-score";
            }


            //PowerPoint.Shape setBtn = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeActionButtonCustom, curW - 150, 0, 100, 40);
            //setBtn.TextFrame.TextRange.InsertBefore("设置");
            //setBtn.Name = "kx-setting";

            Globals.ThisAddIn.initSetting();
            string curDir = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);
            var settingImg = curDir + @"\kxrealtime\imgs\setting.png";
            PowerPoint.Shape setBtn = slide.Shapes.AddPicture(settingImg, Office.MsoTriState.msoTrue, Office.MsoTriState.msoTrue, curW - 150, 0, 100, 40);
            setBtn.Name = "kx-setting";


            if (questionType == singleSelCtl.TypeSelEnum.singleSel || questionType == singleSelCtl.TypeSelEnum.multiSel || questionType == singleSelCtl.TypeSelEnum.voteSingleSel || questionType == singleSelCtl.TypeSelEnum.voteMultiSel)
            {
                this.initOption(slide, 4, questionType == singleSelCtl.TypeSelEnum.multiSel);
            }
            else if(questionType == singleSelCtl.TypeSelEnum.textQuestion)
            {

            }
            else if(questionType == singleSelCtl.TypeSelEnum.fillQuestion && scoreCom != null)
            {
                scoreCom.TextFrame.TextRange.Text = ("0分");
            }
            //slide.Select();
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            if(this.myCustomTaskPane == null)
            {
                this.initPanes("单选题");
            }
            else
            {
                this.myCustomTaskPane.Visible = true;
            }
            this.createSingleCtx("单选题", singleSelCtl.TypeSelEnum.singleSel);
            this.singleSelCtlInstance.setCurSelType = singleSelCtl.TypeSelEnum.singleSel;
            //char[] ans = new char[] { };
            //this.singleSelCtlInstance.resetData(0,ans,4);
        }

        private void initOption(PowerPoint.Slide slide, int n, bool isMul)
        {
            char sChar = 'A';
            int posY = 200;
            float difY = (250 - n * 50) / (n - 1);
            Office.MsoAutoShapeType curShapeType = !isMul ? Office.MsoAutoShapeType.msoShapeOval : Office.MsoAutoShapeType.msoShapeRectangle;
            for (int i=0; i<n; i++)
            {
                char curChar = (char)(sChar + i);
                PowerPoint.Shape circleTmp = slide.Shapes.AddShape(curShapeType, 100, posY+ difY *i - 5, 40, 40);
                circleTmp.TextFrame.TextRange.InsertAfter(curChar.ToString());
                circleTmp.Name = "kx-choice-" + curChar.ToString();
                circleTmp.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(1,128,128,128).ToArgb();
                circleTmp.Line.ForeColor.RGB = System.Drawing.Color.FromArgb(1, 128, 128, 128).ToArgb();
                PowerPoint.Shape textBox = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal, 150, posY + difY * i, 500, 50);
                textBox.TextFrame.TextRange.InsertAfter("此处添加选项内容");
                textBox.Name = "kx-text-" + curChar.ToString();
                posY += 50;
            }
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
            for(int i=0; i<26; i++)
            {
                char curChar = (char)(sChar + i);
                var curKey = "kx-choice-" + curChar;
                var textKey = "kx-text-" + curChar;
                if (shapeMap.Contains(curKey))
                {
                    var curShape = shapeMap[curKey] as Shape;
                    if(lastChar != curChar)
                    {
                        curShape.Name = "kx-choice-" + lastChar;
                        curShape.TextFrame.TextRange.Text = lastChar.ToString();
                        if(shapeMap.Contains(textKey))
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
                    if(shapeMap.Contains(textKey))
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
                if(shapeTmp.Name == "kx-question")
                {
                    questionStr = shapeTmp.TextFrame.TextRange.Text;
                }
                else if(shapeTmp.Name == "kx-qInfo")
                {
                    ansStr = shapeTmp.TextFrame.TextRange.Text;
                }
            }
            if(questionStr.Length > 0)
            {
                string pattern = @"(\[填空\d\])";
                var fillOptionCollect = Regex.Matches(questionStr, pattern);
                if (ansStr.Length > 0)
                {
                    var tmp = JsonConvert.DeserializeObject<List<fillOption>>(ansStr);
                    for(int i=0; i<fillOptionCollect.Count; i++)
                    {
                        if(tmp.Count > i)
                        {
                            ansTmp.Add(tmp[i]);
                        } else
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
            if(this.singleSelCtlInstance.setCurSelType == singleSelCtl.TypeSelEnum.singleSel || this.singleSelCtlInstance.setCurSelType == singleSelCtl.TypeSelEnum.multiSel || this.singleSelCtlInstance.setCurSelType == singleSelCtl.TypeSelEnum.voteMultiSel || this.singleSelCtlInstance.setCurSelType == singleSelCtl.TypeSelEnum.voteSingleSel)
            {
                System.Diagnostics.Debug.WriteLine("checkSelExist");
                resetSingleSel(slide);
            }
            else if(this.singleSelCtlInstance.setCurSelType == singleSelCtl.TypeSelEnum.fillQuestion)
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
            this.createSingleCtx("多选题", singleSelCtl.TypeSelEnum.multiSel);
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
            this.createSingleCtx("主观题", singleSelCtl.TypeSelEnum.textQuestion);
            this.singleSelCtlInstance.setCurSelType = singleSelCtl.TypeSelEnum.textQuestion;
            this.singleSelCtlInstance.initSubjectiveQ(0);
        }

        
        // 重新检测ppt内容
        public void resestContent(PowerPoint.SlideRange slide)
        {
            if(this.myCustomTaskPane == null)
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
                this.singleSelCtlInstance.initSubjectiveQ(score);
            }
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            if(this.curLoginDialog != null)
            {
                return;
            }
            Int32 curW = (Int32)app.ActivePresentation.SlideMaster.Width;
            Int32 curH = (Int32)app.ActivePresentation.SlideMaster.Height;

            /*if (loginCtlInstance == null)
            {
                loginCtlInstance = new loginCtl();
            }
            
            if(loginPane == null)
            {
                loginPane = Globals.ThisAddIn.CustomTaskPanes.Add(loginCtlInstance, "扫码登录");
                loginPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionFloating;
                loginPane.Height = 800;
                loginPane.Width = 800;
                loginPane.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;

            }

            if(loginPane.Visible)
            {
                return;
            }*/

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
            this.initLoginListener(connectID);

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
            if(loginWebSocket != null)
            {
                loginWebSocket.Stop(System.Net.WebSockets.WebSocketCloseStatus.NormalClosure, "close");
            }
            loginWebSocket = utils.webSocketClient.StartWebSocket($"{utils.KXINFO.KXSOCKETURL}/mobileLogin?client_id={curID}");
            loginWebSocket.MessageReceived.Subscribe(msg => {
                try
                {
                    if(loginWebSocket == null)
                    {
                        return;
                    }
                    string msgStr = msg.Text;
                    utils.KXINFO.initUsr(msgStr);
                    // close login pane
                    Action actionDelegate = this.LoginSuccess;
                    // 或者
                    // Action<string> actionDelegate = delegate(string txt) { this.label2.Text = txt; };
                    this.curLoginDialog.Invoke(actionDelegate);
                    
                    this.button5.Visible = false;
                    this.menu1.Visible = true;
                    this.menu1.Label = utils.KXINFO.KXUNAME;
                    this.closeLoginConnect();
                }
                catch(Exception e)
                {
                    MessageBox.Show("登录失败");
                }
               
            });
        }

        public void closeLoginConnect()
        {
            if (this.loginWebSocket != null)
            {
                this.loginWebSocket.Stop(System.Net.WebSockets.WebSocketCloseStatus.NormalClosure, "close");
                
                this.loginWebSocket.Dispose();
                this.loginWebSocket = null;
            }
            this.curLoginDialog = null;
        }

        private void CloseLoginDialog()
        {
            if(loginDialog.frmBack != null)
            {
                loginDialog.frmBack.Close();
            }
            if(this.curLoginDialog != null)
            {
                this.curLoginDialog.Close();
                this.curLoginDialog = null;
            }
            
            // need rechose the class info
            ChangeTchBtn(false);
            Globals.ThisAddIn.CloseTchSocket();
        }

        private void LoginSuccess()
        {
            var url = utils.KXINFO.KXUAVATAR;
            PictureBox pictureBox = new PictureBox();
            try
            {
                System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12 | System.Net.SecurityProtocolType.Tls11 | System.Net.SecurityProtocolType.Tls;
                pictureBox.Load(url);
                menu1.Image = pictureBox.Image;
            }
            catch(WebException e)
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

            var loginContentTmp = curLoginDialog.getContent;
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
                if(timeLeft < 0)
                {
                    myTimer.Stop();
                    this.CloseLoginDialog();
                }
            }); //给timer挂起事件
            myTimer.Enabled = true;
            myTimer.Interval = 1000;  
        }

        public void settingChange(bool isSetting)
        {
            for(int i=1; i<= app.ActivePresentation.Slides.Count; i++) {
                PowerPoint.Slide curSld = app.ActivePresentation.Slides[i];
                bool isKxItem = curSld.Name.Contains("kx-slide");
                if(!isKxItem)
                {
                    continue;
                }
                ArrayList curAnswerArr = (ArrayList)AnswerStore.getAnswer(curSld.Name);
                if(curAnswerArr == null)
                {
                    curAnswerArr = new ArrayList();
                }
                PowerPoint.Shapes curShapes = curSld.Shapes;
                foreach (PowerPoint.Shape shapeTmp in curShapes)
                {
                    string targetName = "kx-setting";
                    if (shapeTmp.Name == targetName)
                    {
                        shapeTmp.Visible = isSetting ? Office.MsoTriState.msoCTrue : Office.MsoTriState.msoFalse;
                    }
                    targetName = "kx-sending";
                    if (shapeTmp.Name == targetName)
                    {
                        shapeTmp.Visible = isSetting ? Office.MsoTriState.msoCTrue : Office.MsoTriState.msoFalse;
                    }
                    if(shapeTmp.Name.Contains("kx-choice"))
                    {
                        string tmp = shapeTmp.Name.Substring(10);
                        if(tmp.Length == 0)
                        {
                            continue;
                        }
                        char curC =tmp[0];
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
                if(!isSetting)
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
        }

        // play slide
        private void button9_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ActivePresentation.SlideShowSettings.Run();
        }

        // close couse
        private void button10_Click(object sender, RibbonControlEventArgs e)
        {
            object sendData = (new
            {
                key = "classroom",
                value = utils.KXINFO.KXCHOSECLASSID,
                type = "COURSE_END",
                data = new {},
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
            }
            catch (Exception) { }
            
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
            this.createSingleCtx("投票", singleSelCtl.TypeSelEnum.voteSingleSel);
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
            this.createSingleCtx("填空题", singleSelCtl.TypeSelEnum.fillQuestion);
            this.singleSelCtlInstance.setCurSelType = singleSelCtl.TypeSelEnum.fillQuestion;
            this.singleSelCtlInstance.initFillQ(0);
        }
    }
}
