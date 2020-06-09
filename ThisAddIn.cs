using kxrealtime.kxdata;
using kxrealtime.utils;
using Microsoft.Office.Interop.PowerPoint;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Websocket.Client;
using WebSocket4Net;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace kxrealtime
{
    // 接收websocket的函数
    public delegate void WebSocketMsgHandle(string msg);

    public partial class ThisAddIn
    {
        // 当前幻灯片序号
        private int curSlideIdx = 1;
        // 当前播放的幻灯片序号
        private int playSlideIdx = 1;
        // 是否是酷校页面
        private bool isKxPage;
        // 选择类型
        private PowerPoint.PpSelectionType lastType;
        // 请求对象
        private RestSharp.RestClient curHttpReq;
        // 是否显示加入班级页面
        private bool canShowAddClassFlag = false;
        // webscoket对象
        private webSocketClient TchWebSocket;
        // 工具栏对象
        private utilDialog utilDialogInstance;
        // 当前窗口句柄
        private int curActiveWn;
        // 是否在发送中
        private bool isSameSending = false;
        // webscoket 接收事件
        public event WebSocketMsgHandle WebSocketMsg;

        // 获取当前幻灯片index
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

        // 获取当前播放中幻灯片index
        public int PlaySlideIdx
        {
            get
            {
                return this.playSlideIdx;
            }
            set
            {
                this.playSlideIdx = value;
            }
        }

        // 获取请求对象
        public RestSharp.RestClient CurHttpReq
        {
            get
            {
                return this.curHttpReq;
            }
        }

        // 获取是否显示加入班级
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

        // 幻灯片考试信息列表
        public List<slideExamInfo> kxSlideExam
        {
            get
            {
                return this.kxSlideExamList;
            }
        }

        // 获取当前窗口句柄
        public int getCurActiveWn
        {
            get
            {
                return this.curActiveWn;
            }
        }

        // 开始函数
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
            this.Application.WindowActivate += Application_WindowActivate;

            this.curHttpReq = utils.request.GetClient();
        }

        // 窗口active事件回调
        private void Application_WindowActivate(PowerPoint.Presentation Pres, PowerPoint.DocumentWindow Wn)
        {
            this.curActiveWn = Wn.HWND;
        }

        // ppt关闭事件回调
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

        // 上一张事件回调
        private void Application_SlideShowOnPrevious(PowerPoint.SlideShowWindow Wn)
        {
            System.Diagnostics.Debug.WriteLine("this is on previous");
        }
        // 下一张事件回调
        private void Application_SlideShowOnNext(PowerPoint.SlideShowWindow Wn)
        {
            System.Diagnostics.Debug.WriteLine("this is on next");
        }

        // 切换幻灯片事件回调
        private void Application_SlideShowNextSlide(PowerPoint.SlideShowWindow Wn)
        {
            if(isSameSending && Wn.View.Slide.SlideIndex == this.playSlideIdx)
            {
                return;
            }
            isSameSending = true;
            System.Diagnostics.Debug.WriteLine("this is new slide" + Wn.View.Slide.SlideIndex + " old= " + this.playSlideIdx);
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
            isSameSending = false;
        }

        // 检查工具按钮
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

        // 处理当前slide
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

        // 打开加入班级
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

        // 初始化设置图片按钮
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

        // 发送ppt图片
        private void sendScreen(PowerPoint.SlideShowWindow Wn, int curIdx)
        {
            // 开启了授课
            if (TchWebSocket != null && TchWebSocket.State == WebSocketState.Open)
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
                //utils.Utils.LOG("授课连接中...");
                InitTchSocket();
            }
        }

        // 非播放状态切换ppt回调
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
                if (Globals.Ribbons.Ribbon1 != null && Globals.Ribbons.Ribbon1.myCustomTaskPane != null  && Globals.Ribbons.Ribbon1.myCustomTaskPane.Visible)
                {
                    Globals.Ribbons.Ribbon1.resestContent(SldRange);
                }
                this.isKxPage = true;
            }
        }

        // 选中回调
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

        // 检查是否是酷校ppt
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

        // 结束放映回调
        private void SlideShowEnd(PowerPoint.Presentation Pres)
        {
            //("结束放映");
            Globals.Ribbons.Ribbon1.settingChange(true);
            if (utilDialogInstance != null && !utilDialogInstance.IsDisposed)
            {
                utilDialogInstance.Close();
            }
            this.utilDialogInstance = null;
            this.playSlideIdx = 1;
        }

        // 开始放映回调
        private void SlideShowBegin(PowerPoint.SlideShowWindow Wn)
        {
            //("开始放映");
            if(!Globals.Ribbons.Ribbon1.isPlaying)
            {
                Globals.Ribbons.Ribbon1.settingChange(false);
            }
            isSameSending = false;
            if (this.playSlideIdx != 1)
            {
                Wn.View.GotoSlide(this.playSlideIdx);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

        }

        // 退出
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

        // 初始化授课webscoket
        public void InitTchSocket()
        {
            if (TchWebSocket != null)
            {
                //TchWebSocket.Stop(System.Net.WebSockets.WebSocketCloseStatus.NormalClosure, "stop old connect");
                CloseTchSocket();
            }
            string url = $"{utils.KXINFO.KXSOCKETURL}/im?user_id={utils.KXINFO.KXOUTUID}";
            TchWebSocket = new utils.webSocketClient();
            TchWebSocket.StartWebSocket(url);
            TchWebSocket.MessageReceived += tchSocketHandle; ;
        }

        // 接收授课webssocket内容
        private void tchSocketHandle(String msg)
        {
            try
            {
                if (msg == "HeartBeat")
                {
                    return;
                }
                JObject data = JObject.Parse(msg);
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

        // 发送授课信息
        public void SendTchInfo(string info)
        {
            if (TchWebSocket == null)
            {
                return;
            }
            if (TchWebSocket.State == WebSocketState.Open)
            {
                TchWebSocket.clientSend(info);
            }
        }

        // 关闭授课websocket
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

        // 查找ppt考试信息
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

        // 查找ppt考试信息 根据id
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

        // 删除ppt考试信息
        public void removeExamItem(string paperId)
        {
            var itemToRemove = this.kxSlideExamList.SingleOrDefault(r => r.paperId == paperId);
            if (itemToRemove != null)
            {
                this.kxSlideExamList.Remove(itemToRemove);
            }
        }

        // 关闭相关的弹窗
        public void closeWin()
        {
            Globals.Ribbons.Ribbon1.closeAllWin();
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
