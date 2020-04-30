using kxrealtime.kxdata;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Websocket.Client;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace kxrealtime
{

    public delegate void WebSocketMsgHandle(string msg);

    public partial class ThisAddIn
    {
        private int curSlideIdx = 1;
        private int playSlideIdx = 1;
        private bool isKxPage;
        private PowerPoint.PpSelectionType lastType;
        private RestSharp.RestClient curHttpReq;
        private bool canShowAddClassFlag = false;
        private IWebsocketClient TchWebSocket;
        private utilDialog utilDialogInstance;

        public event WebSocketMsgHandle WebSocketMsg;

        public int CurSlideIdx
        {
            get
            {
                return this.curSlideIdx;
            }
            set
            {
                this.curSlideIdx = value;
            }
        }

        public int PlaySlideIdx
        {
            get
            {
                return this.playSlideIdx;
            }
        }

        public RestSharp.RestClient CurHttpReq
        {
            get
            {
                return this.curHttpReq;
            }
        }

        public bool canShowAddClass
        {
            get
            {
                return this.canShowAddClassFlag;
            }
            set
            {
                this.canShowAddClassFlag = value;
            }
        }


        // slide exame
        private List<slideExamInfo> kxSlideExamList = new List<slideExamInfo>();

        public List<slideExamInfo> kxSlideExam
        {
            get
            {
                return this.kxSlideExamList;
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.SlideSelectionChanged += Application_SlideSelectionChanged;
            this.Application.WindowSelectionChange += selectionChange;
            this.Application.SlideShowBegin += new PowerPoint.EApplication_SlideShowBeginEventHandler(SlideShowBegin);
            this.Application.SlideShowEnd += new PowerPoint.EApplication_SlideShowEndEventHandler(SlideShowEnd);
            this.Application.SlideShowNextSlide += Application_SlideShowNextSlide;
            this.Application.SlideShowOnNext += Application_SlideShowOnNext;
            this.Application.SlideShowOnPrevious += Application_SlideShowOnPrevious;
            this.Application.PresentationBeforeClose += Application_PresentationBeforeClose;

            this.curHttpReq = utils.request.GetClient();
        }

        private void Application_PresentationBeforeClose(PowerPoint.Presentation Pres, ref bool Cancel)
        {
            if (utils.KXINFO.KXTCHRECORDID != null && utils.KXINFO.KXTCHRECORDID.Length > 0)
            {
                MessageBoxButtons messbutton = MessageBoxButtons.OKCancel;
                DialogResult dr = MessageBox.Show("是否需要结束授课", "温馨提示", messbutton);
                if (dr == DialogResult.OK)
                {
                    Globals.Ribbons.Ribbon1.stopTch();
                }
            }
            Cancel = false;

        }

        private void Application_SlideShowOnPrevious(PowerPoint.SlideShowWindow Wn)
        {
            System.Diagnostics.Debug.WriteLine("this is on previous");
        }

        private void Application_SlideShowOnNext(PowerPoint.SlideShowWindow Wn)
        {
            System.Diagnostics.Debug.WriteLine("this is on next");
        }

        private void Application_SlideShowNextSlide(PowerPoint.SlideShowWindow Wn)
        {
            System.Diagnostics.Debug.WriteLine("this is new slide" + Wn.View.Slide.SlideIndex);
            if (utils.KXINFO.KXTCHRECORDID == null)
            {
                return;
            }
            this.playSlideIdx = Wn.View.Slide.SlideIndex;
            var winNum = Screen.AllScreens.Length;
            if (this.utilDialogInstance != null || winNum > 1)
            {
                this.checkUtils(Wn);
            }
            else
            {
                slideHandle(Wn);
            }
            sendScreen(Wn, this.playSlideIdx);
        }

        private void checkUtils(PowerPoint.SlideShowWindow Wn)
        {
            if (this.utilDialogInstance == null || this.utilDialogInstance.IsDisposed)
            {
                this.utilDialogInstance = new utilDialog();
            }
            this.utilDialogInstance.Location = utils.Utils.getScreenPosition();
            if (Wn.View.Slide.Name.Contains("kx-slide"))
            {
                this.utilDialogInstance.showSendBtn();
            }
            else
            {
                //this.utilDialogInstance.Close();
                this.utilDialogInstance.onlyUtils();
            }
            if (this.canShowAddClassFlag)
            {
                this.openAddClass();
                this.canShowAddClassFlag = false;
            }
        }

        private void slideHandle(PowerPoint.SlideShowWindow Wn)
        {
            System.Windows.Forms.Timer myTimer = new System.Windows.Forms.Timer();//实例化　
            myTimer.Tick += new EventHandler((s, e) =>
            {
                this.checkUtils(Wn);
                myTimer.Stop();
                myTimer.Dispose();
                myTimer = null;
            }); //给timer挂起事件
            myTimer.Enabled = true;
            myTimer.Interval = 300;
        }

        private void openAddClass()
        {
            // 邀请进入班级窗口
            var classId = (utils.KXINFO.KXCHOSECLASSID).ToString();
            var className = utils.KXINFO.KXCHOSECLASSNAME;
            var curWn = new addClass(classId, className);
            curWn.Location = utils.Utils.getScreenPosition();
            curWn.Show();
            Globals.Ribbons.Ribbon1.ChangeTchBtn(true);
        }

        public void initSetting()
        {
            string curDir = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);
            var fileDict = curDir + @"\kxrealtime\imgs";
            var imgPath = fileDict + @"\setting.png";
            if (File.Exists(imgPath))
            {
                return;
            }
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
            System.Drawing.Bitmap rs = (System.Drawing.Bitmap)(Properties.Resources.setting);
            rs.Save(fileDict + @"\setting.png");
        }

        private void sendScreen(PowerPoint.SlideShowWindow Wn, int curIdx)
        {
            // 开启了授课
            if (TchWebSocket != null && TchWebSocket.IsRunning)
            {
                //var imgTmp = utils.Utils.getScreenImg();
                //MessageBox.Show(curDirTmp);
                //Task.Run(() => {
                //string curDir = Directory.GetCurrentDirectory();
                try
                {
                    string curDir = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);
                    var curSld = Wn.View.Slide;
                    var fileDict = curDir + @"\kxrealtime\imgs";
                    var imgFile = fileDict + @"\" + curSld.Name + ".png";
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
                    curSld.Export(imgFile, "png");
                    utils.request.UploadImg($"{utils.KXINFO.KXURL}/usr/upload?session_id={utils.KXINFO.KXSID}", imgFile, curIdx);
                }
                catch (Exception e)
                {
                    utils.Utils.LOG("export失败" + e.Message);
                }

                // });

            }
            else
            {
                utils.Utils.LOG("授课连接中...");
                InitTchSocket();
            }
        }

        private void Application_SlideSelectionChanged(PowerPoint.SlideRange SldRange)
        {
            if (this.CurSlideIdx == SldRange.SlideIndex)
            {
                return;
            }
            this.curSlideIdx = SldRange.SlideIndex;
            bool isKxItem = SldRange.Name.Contains("kx-slide");
            if (!isKxItem)
            {
                if (Globals.Ribbons.Ribbon1 != null && Globals.Ribbons.Ribbon1.myCustomTaskPane != null)
                {
                    Globals.Ribbons.Ribbon1.myCustomTaskPane.Visible = false;
                }
                this.isKxPage = false;
            }
            else
            {
                if (Globals.Ribbons.Ribbon1 != null && Globals.Ribbons.Ribbon1.myCustomTaskPane != null && Globals.Ribbons.Ribbon1.myCustomTaskPane.Visible)
                {
                    Globals.Ribbons.Ribbon1.resestContent(SldRange);
                }
                this.isKxPage = true;
            }
        }

        private void selectionChange(PowerPoint.Selection Sel)
        {
            if(!this.isKxPage && Sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                var checkResult = checkIsKx();
                if(checkResult)
                {
                    this.isKxPage = true;
                }
            }
            if (this.isKxPage)
            {
                if (Sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    if (Sel.ShapeRange.Count == 1 && Sel.ShapeRange.Name == "kx-setting" && Globals.Ribbons.Ribbon1 != null)
                    {
                        Globals.Ribbons.Ribbon1.resestContent(Sel.SlideRange);
                        Globals.Ribbons.Ribbon1.myCustomTaskPane.Visible = true;
                    }
                }
                else if (this.lastType == PowerPoint.PpSelectionType.ppSelectionShapes && Globals.Ribbons.Ribbon1 != null && Globals.Ribbons.Ribbon1.myCustomTaskPane != null && Globals.Ribbons.Ribbon1.myCustomTaskPane.Visible)
                {
                    // 需要检测是否发送了删除选项等
                    Globals.Ribbons.Ribbon1.checkSelExist(Sel.SlideRange);
                }
                this.lastType = Sel.Type;
            }

        }

        private bool checkIsKx()
        {
            var curSlide = Application.ActivePresentation.Slides[curSlideIdx];
            var curShapes = curSlide.Shapes;
            foreach (PowerPoint.Shape shapeTmp in curShapes)
            {
                if (shapeTmp.Name == "kx-setting")
                {
                    curSlide.Name = "kx-slide-" + curSlide.Name;
                    return true;
                }
            }
            return false;
        }

        private void SlideShowEnd(PowerPoint.Presentation Pres)
        {
            //("结束放映");
            Globals.Ribbons.Ribbon1.settingChange(true);
            if (utilDialogInstance != null && !utilDialogInstance.IsDisposed)
            {
                utilDialogInstance.Close();
            }
            this.utilDialogInstance = null;

        }

        private void SlideShowBegin(PowerPoint.SlideShowWindow Wn)
        {
            //("开始放映");
            Globals.Ribbons.Ribbon1.settingChange(false);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

        }

        public void loginOut()
        {
            this.kxSlideExamList.Clear();
            utils.KXINFO.clear();
            CloseTchSocket();
            Globals.Ribbons.Ribbon1.ChangeTchBtn(false);
            if (this.utilDialogInstance != null)
            {
                this.utilDialogInstance.Close();
            }
        }

        public void InitTchSocket()
        {
            if (TchWebSocket != null)
            {
                //TchWebSocket.Stop(System.Net.WebSockets.WebSocketCloseStatus.NormalClosure, "stop old connect");
                CloseTchSocket();
            }
            string url = $"{utils.KXINFO.KXSOCKETURL}/im?user_id={utils.KXINFO.KXOUTUID}";
            TchWebSocket = utils.webSocketClient.StartWebSocket(url);
            TchWebSocket.MessageReceived.Subscribe(tchSocketHandle);
        }

        private void tchSocketHandle(ResponseMessage info)
        {
            try
            {
                if (info.Text == "HeartBeat")
                {
                    return;
                }
                JObject data = JObject.Parse(info.Text);
                string curType = (string)data["type"];
                if (curType == "barrage")
                {
                    string tchId = (string)data["teach_record_id"];
                    if (tchId != utils.KXINFO.KXTCHRECORDID)
                    {
                        return;
                    }
                    string contentStr = (data["data"]).ToString();
                    this.WebSocketMsg(contentStr);
                }
            }
            catch (Exception e)
            {
            }
        }

        public void SendTchInfo(string info)
        {
            if (TchWebSocket == null)
            {
                return;
            }
            if (TchWebSocket.IsRunning)
            {
                utils.webSocketClient.clientSend(TchWebSocket, info);
            }
        }

        public void CloseTchSocket()
        {
            if (TchWebSocket == null)
            {
                return;
            }
            //TchWebSocket.Stop(System.Net.WebSockets.WebSocketCloseStatus.NormalClosure, "USER");
            TchWebSocket.Dispose();
            TchWebSocket = null;
        }

        public slideExamInfo findExamInfo(string sldName)
        {
            foreach (slideExamInfo curExam in this.kxSlideExamList)
            {
                if (curExam.slideName == sldName)
                {
                    return curExam;
                }
            }
            return null;
        }

        public slideExamInfo findExamInfoByTestId(string testId)
        {
            foreach (slideExamInfo curExam in this.kxSlideExamList)
            {
                if (curExam.testId == testId)
                {
                    return curExam;
                }
            }
            return null;
        }

        public void removeExamItem(string paperId)
        {
            var itemToRemove = this.kxSlideExamList.SingleOrDefault(r => r.paperId == paperId);
            if (itemToRemove != null)
            {
                this.kxSlideExamList.Remove(itemToRemove);
            }
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
