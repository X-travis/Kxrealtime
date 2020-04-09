using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Websocket.Client;
using System.Threading.Tasks;
using System.IO;
using kxrealtime.kxdata;
using System.Windows.Forms;

namespace kxrealtime
{
    public partial class ThisAddIn
    {
        public int CurSlideIdx = 1;
        public int PlaySlideIdx = 1;
        private bool isKxPage;
        private PowerPoint.PpSelectionType lastType;
        public RestSharp.RestClient CurHttpReq;
        public bool canShowAddClass = false;
        private IWebsocketClient TchWebSocket;
        private utilDialog utilDialogInstance;


        // slide exame
        public List<slideExamInfo> kxSlideExam = new List<slideExamInfo>();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.SlideSelectionChanged += Application_SlideSelectionChanged;
            this.Application.WindowSelectionChange += selectionChange;
            this.Application.SlideShowBegin += new PowerPoint.EApplication_SlideShowBeginEventHandler(SlideShowBegin);
            this.Application.SlideShowEnd += new PowerPoint.EApplication_SlideShowEndEventHandler(SlideShowEnd);
            this.Application.SlideShowNextSlide += Application_SlideShowNextSlide;
            this.Application.SlideShowOnNext += Application_SlideShowOnNext;
            this.Application.SlideShowOnPrevious += Application_SlideShowOnPrevious;
            
            CurHttpReq = utils.request.GetClient();
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
            System.Diagnostics.Debug.WriteLine("this is new slide");
            if (utils.KXINFO.KXTCHRECORDID == null)
            {
                return;
            }
            this.PlaySlideIdx = Wn.View.Slide.SlideIndex;
            sendScreen(Wn);
            if(this.utilDialogInstance != null)
            {
                this.checkUtils(Wn);
            } else
            {
                slideHandle(Wn);
            }
        }

        private void checkUtils(PowerPoint.SlideShowWindow Wn)
        {
            if (this.utilDialogInstance == null)
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
            if (this.canShowAddClass)
            {
                this.openAddClass();
                this.canShowAddClass = false;
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
            myTimer.Interval = 500;
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
            if(File.Exists(imgPath))
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

        private void sendScreen(PowerPoint.SlideShowWindow Wn)
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
                    utils.request.UploadImg($"{utils.KXINFO.KXURL}/usr/upload?session_id={utils.KXINFO.KXSID}", imgFile);
                }catch(Exception e)
                {
                    utils.Utils.LOG("export失败" + e.Message);
                }
                    
               // });
                
            } else
            {
                utils.Utils.LOG("授课连接中断，请重新开启授课");
                TchWebSocket.Reconnect();
            }
        }

        private void Application_SlideSelectionChanged(PowerPoint.SlideRange SldRange)
        {
            if (this.CurSlideIdx == SldRange.SlideIndex)
            {
                return;
            }
            this.CurSlideIdx = SldRange.SlideIndex;
            bool isKxItem = SldRange.Name.Contains("kx-slide");
            if(!isKxItem)
            {
                if(Globals.Ribbons.Ribbon1 != null && Globals.Ribbons.Ribbon1.myCustomTaskPane != null)
                {
                    Globals.Ribbons.Ribbon1.myCustomTaskPane.Visible = false;
                }
                this.isKxPage = false;
            } else
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
            if (this.isKxPage)
            {
                if (Sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    if (Sel.ShapeRange.Count == 1 && Sel.ShapeRange.Name == "kx-setting" && Globals.Ribbons.Ribbon1 != null)
                    {
                        Globals.Ribbons.Ribbon1.resestContent(Sel.SlideRange);
                        Globals.Ribbons.Ribbon1.myCustomTaskPane.Visible = true;
                    }
                } else if(this.lastType == PowerPoint.PpSelectionType.ppSelectionShapes && Globals.Ribbons.Ribbon1 != null && Globals.Ribbons.Ribbon1.myCustomTaskPane != null && Globals.Ribbons.Ribbon1.myCustomTaskPane.Visible)
                {
                    // 需要检测是否发送了删除选项等
                    Globals.Ribbons.Ribbon1.checkSelExist(Sel.SlideRange);
                }
                this.lastType = Sel.Type;
            }
            
        }

        private void SlideShowEnd(PowerPoint.Presentation Pres)
        {
            //("结束放映");
            Globals.Ribbons.Ribbon1.settingChange(true);
            if(utilDialogInstance != null)
            {
                utilDialogInstance.Close();
                this.utilDialogInstance = null;
            }
            
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
            kxSlideExam.Clear();
            utils.KXINFO.clear();
            CloseTchSocket();
            Globals.Ribbons.Ribbon1.ChangeTchBtn(false);
            if(this.utilDialogInstance != null)
            {
                this.utilDialogInstance.Close();
            }
        }

        public void InitTchSocket()
        {
            if(TchWebSocket != null)
            {
                TchWebSocket.Stop(System.Net.WebSockets.WebSocketCloseStatus.NormalClosure, "stop old connect");
            }
            string url = $"{utils.KXINFO.KXSOCKETURL}/im?user_id={utils.KXINFO.KXOUTUID}";
            TchWebSocket = utils.webSocketClient.StartWebSocket(url);
        }

        public void SendTchInfo(string info)
        {
            if(TchWebSocket == null)
            {
                return;
            }
            if(TchWebSocket.IsRunning)
            {
                utils.webSocketClient.clientSend(TchWebSocket, info);
            }
        }

        public void CloseTchSocket()
        {
            if(TchWebSocket == null)
            {
                return;
            }
            TchWebSocket.Stop(System.Net.WebSockets.WebSocketCloseStatus.NormalClosure, "USER");
            TchWebSocket.Dispose();
            TchWebSocket = null;
        }

        public slideExamInfo findExamInfo(string sldName)
        {
            foreach(slideExamInfo curExam in kxSlideExam)
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
            foreach (slideExamInfo curExam in kxSlideExam)
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
            var itemToRemove = kxSlideExam.Single(r => r.paperId == paperId);
            if (itemToRemove != null)
            {
                kxSlideExam.Remove(itemToRemove);
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
