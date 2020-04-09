using kxrealtime.kxdata;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace kxrealtime
{
    public partial class choseTime : Form
    {
        private Form frmBack;
        private int btnWidthTmp = 400;
        private int difWidth = 80;
        private utilDialog parentForm;
        private string paperTitle = "";
        private float paperScore;
        private Int64 paperTime;
        private string paperId;
        private string testId;
        public choseTime(utilDialog ownerForm)
        {
            InitializeComponent();
            frmBack = new Form();
            frmBack.FormBorderStyle = FormBorderStyle.None;
            frmBack.StartPosition = FormStartPosition.Manual;
            frmBack.ShowIcon = false;
            frmBack.ShowInTaskbar = false;
            frmBack.WindowState = FormWindowState.Maximized;
            frmBack.TopMost = true;
            this.TopMost = true;
            frmBack.Opacity = 0.5;
            frmBack.BackColor = System.Drawing.Color.Black;
            this.Owner = frmBack;
            frmBack.Owner = ownerForm;
           
            parentForm = ownerForm;

            this.paperTitle = utils.KXINFO.KXCHOSECOURSETITLE + "练习";
        }

        public void showFn(string paperId, string testId)
        {
            var tmp = utils.Utils.getScreenPosition();
            this.frmBack.Location = tmp;
            this.Location = tmp;
            this.frmBack.Show();
            this.Show();
            this.createSel();
            this.paperId = paperId;
            this.testId = testId;
        }

        public void closeFn()
        {
            this.paperId = null;
            this.testId = null;
            this.Close();
            this.frmBack.Close();
        }

        private void createSel()
        {
            var txtArr = new string[] { "30秒", "1分钟", "2分钟", "3分钟", "4分钟", "5分钟" };
            var valueArr = new string[] { "30", "60", "120", "180", "240", "300" };
            int hTmp = panel1.Height;
            int wTmp = panel1.Width;
            button2.Top = hTmp / 2 + panel1.Top;
            button3.Top = hTmp / 2 + panel1.Top;
            int countTmp = txtArr.Count();
            btnWidthTmp = (wTmp - difWidth* countTmp) / countTmp;
            for (int i=0; i<txtArr.Length; i++)
            {
                var btnTmp = new Button();
                btnTmp.Text = txtArr[i];
                btnTmp.Name = valueArr[i];
                btnTmp.Top = hTmp / 6;
                btnTmp.Left = i*(btnWidthTmp + difWidth) + difWidth;
                btnTmp.Width = btnWidthTmp;
                btnTmp.Height = hTmp * 2 / 3;
                btnTmp.Visible = true;
                btnTmp.BackColor =  System.Drawing.Color.FromArgb(100,74,137,211);
                btnTmp.FlatStyle = FlatStyle.Flat;
                btnTmp.FlatAppearance.CheckedBackColor = System.Drawing.Color.Red;
                btnTmp.Click += BtnTmp_Click;
                btnTmp.MouseHover += BtnTmp_MouseHover;
                btnTmp.ForeColor = System.Drawing.Color.White;
                btnTmp.Font = new System.Drawing.Font(btnTmp.Font.Name, 30F);
                this.panel1.Controls.Add(btnTmp);
            }
            this.panel1.Visible = true;
        }

        private void BtnTmp_MouseHover(object sender, EventArgs e)
        {
            Button curBtn = sender as Button;
            //curBtn.Width += 40;
            
        }

        private void BtnTmp_Click(object sender, EventArgs e)
        {
            Button curBtn = sender as Button;
            if(curBtn != null)
            {
                this.paperTime = Int64.Parse(curBtn.Name) * 1000;
                if(this.paperId == null || this.paperId.Length == 0)
                {
                    try
                    {
                        sendPageFn();
                    }catch(Exception)
                    {
                        return;
                    }
                    
                } else
                {
                    var curTime = utils.Utils.getTimeStamp();
                    // 延时操作
                    if (this.testId != null)
                    {
                        var examTmp = Globals.ThisAddIn.findExamInfoByTestId(this.testId);
                        if (examTmp != null)
                        {
                            examTmp.duringTime += this.paperTime;
                            this.paperTime = examTmp.duringTime;
                            curTime = examTmp.startTimeStamp;
                        }
                    }
                    createExam(paperId, curTime);
                    sendPaperExt(paperId, false);
                    sendKXOUT(this.paperId, this.testId);

                    // send exam time
                    sendChangeTime(this.paperId, this.testId);
                }
                closeFn();
                this.parentForm.checkMod();
            }
        }

        private void sendChangeTime(string paperId, string testId)
        {
            object sendData = (new
            {
                key = "exam",
                type = "TEST_TIME",
                data = new
                {
                    time = this.paperTime,
                    paperId = paperId,
                    testId = testId
                },
                timestamp = utils.Utils.getTimeStamp()
            });
            JObject o = JObject.FromObject(sendData);
            string tmp = o.ToString();
            Globals.ThisAddIn.SendTchInfo(tmp);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.closeFn();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            moveBtn(1);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            moveBtn(-1);
        }

        private void moveBtn(int way)
        {
            foreach(Control col in panel1.Controls)
            {
                if(col is Button)
                {
                    col.Left += way * (btnWidthTmp + difWidth);
                }
            }
        }

        public void sendPageFn()
        {
            
            var curIdx = Globals.ThisAddIn.PlaySlideIdx;
            var curSld = Globals.ThisAddIn.Application.ActivePresentation.Slides[curIdx];

            string singleSelTitle = "kx-title-" + singleSelCtl.TypeSelEnum.singleSel;
            string multiSelTitle = "kx-title-" + singleSelCtl.TypeSelEnum.multiSel;
            string voteSingleSelTitle = "kx-title-" + singleSelCtl.TypeSelEnum.voteSingleSel;
            string voteMultiSelTitle = "kx-title-" + singleSelCtl.TypeSelEnum.voteMultiSel;
            string textTitle = "kx-title-" + singleSelCtl.TypeSelEnum.textQuestion;
            string fillTitle = "kx-title-" + singleSelCtl.TypeSelEnum.fillQuestion;

            // 答案
            ArrayList curAnswerArr = (ArrayList)AnswerStore.getAnswer(curSld.Name);
            // 选择题选项
            ArrayList labelArr = new ArrayList();
            // 选项内容
            Hashtable optionMap = new Hashtable();
            // 分数
            float curScore = 0;
            // 类型
            string curTitle = "";
            //题干
            string curQuestion = "";
            // 题目中答案信息 -- 目前只有填空题有
            string ansStr = "";
            foreach (Microsoft.Office.Interop.PowerPoint.Shape shapeTmp in curSld.Shapes)
            {
                if (shapeTmp.Name.Contains("kx-choice"))
                {
                    string tmp = shapeTmp.Name.Substring(10);
                    char curC = tmp[0];
                    labelArr.Add(curC);
                }
                if (shapeTmp.Name.Contains("kx-text"))
                {
                    optionMap.Add(shapeTmp.Name, shapeTmp.TextFrame.TextRange.Text);
                }
                if (shapeTmp.Name == "kx-score")
                {
                    string scoreTmp = shapeTmp.TextFrame.TextRange.Text;
                    curScore = float.Parse(scoreTmp.Replace('分', ' '));
                }
                if (shapeTmp.Name.Contains("kx-title"))
                {
                    curTitle = shapeTmp.Name;
                }
                if (shapeTmp.Name == "kx-question")
                {
                    curQuestion = shapeTmp.TextFrame.TextRange.Text;
                }
                if (shapeTmp.Name == "kx-qInfo")
                {
                    ansStr = shapeTmp.TextFrame.TextRange.Text;
                }
            }

            List<fillOption> fillAnswerArr = new List<fillOption>();
            if (curTitle == fillTitle && curQuestion.Length > 0)
            {
                string pattern = @"(\[填空\d\])";
                var fillOptionCollect = Regex.Matches(curQuestion, pattern);
                if (ansStr.Length > 0)
                {
                    var tmp = JsonConvert.DeserializeObject<List<fillOption>>(ansStr);
                    for (int i = 0; i < fillOptionCollect.Count; i++)
                    {
                        if (tmp.Count > i)
                        {
                            fillAnswerArr.Add(tmp[i]);
                        }
                        else
                        {
                            fillAnswerArr.Add(new fillOption());
                        }
                    }
                }

            }

            this.paperScore = curScore;

            this.paperTitle = utils.KXINFO.KXCHOSECOURSETITLE + "练习";
            bool isQ = curTitle == voteSingleSelTitle || curTitle == voteMultiSelTitle;
            string paperType = isQ ? "kkt_question" : "qset";
            var paperId = createPaper(paperType);
            string testId = "";
            if (paperId == null)
            {
                MessageBox.Show("发送失败");
                throw new Exception("发送失败");
            }//
            if(curTitle != voteSingleSelTitle && curTitle != voteMultiSelTitle)
            {
                var curTime = utils.Utils.getTimeStamp();
                testId = createExam(paperId, curTime);
                if (testId == null)
                {
                    MessageBox.Show("创建考试失败");
                    throw new Exception("创建考试失败");
                }
            }

            //this.paperTime = 6000;
            object questionData = null;
            
            if (curTitle == singleSelTitle || curTitle == voteSingleSelTitle)
            {
                questionData = createSingle(paperId, "single", curScore, curQuestion, labelArr, optionMap, curAnswerArr);
            }
            else if (curTitle == multiSelTitle || curTitle == voteMultiSelTitle)
            {
                questionData = createSingle(paperId, "multi", curScore, curQuestion, labelArr, optionMap, curAnswerArr);
            }
            else if (curTitle == textTitle)
            {
                questionData = createText(paperId, curScore, curQuestion, curAnswerArr);
            }
            else if (curTitle == fillTitle)
            {
                if(fillAnswerArr.Count < 1)
                {
                    MessageBox.Show("请先添加选项");
                    throw new Exception("填空题格式错误");
                }
                questionData = createFill(paperId, curScore, curQuestion, fillAnswerArr);
            }

            var slideExam = new slideExamInfo()
            {
                paperId = paperId,
                testId = testId,
                slideName = curSld.Name,
                startTimeStamp = utils.Utils.getTimeStamp(),
                duringTime = this.paperTime,
                noTime = isQ,
                paperTitle = this.paperTitle
            };
            Globals.ThisAddIn.kxSlideExam.Add(slideExam);

            sendPaperExt(paperId, isQ);
            if (questionData != null)
            {
                sendPaper(paperId, testId, questionData);
            }

        }

        private string createPaper(string paperType)
        {
            Uri reqUrl = new Uri($"{utils.KXINFO.KXCOURSEURL}/usr/api/create");
            Dictionary<string, string> args = new Dictionary<string, string> { };
            args.Add("oid", utils.KXINFO.KXUID);
            args.Add("token", utils.KXINFO.KXTOKEN);
            args.Add("type", paperType);
            args.Add("owner", "user");
            RestSharp.IRestResponse response = utils.request.SendRequest(Globals.ThisAddIn.CurHttpReq, reqUrl, RestSharp.Method.GET, args);
            if (response.ErrorException != null)
            {
                utils.Utils.LOG("usr/api/create api error: " + response.ErrorException.Message);
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
                    utils.Utils.LOG("usr/api/create api error: code" + code);
                }
                else
                {
                    return (string)data["data"]["id"];
                }
            }
            return null;
        }

        private string createExam(string paperId, long sTime)
        {
            Uri reqUrl = new Uri($"{utils.KXINFO.KXCOURSEURL}/usr/api/upsertTest");
            Dictionary<string, string> args = new Dictionary<string, string> { };
            args.Add("token", utils.KXINFO.KXTOKEN);
            //var curTime = utils.Utils.getTimeStamp();
            var postData = new createExamInfo()
            {
                aids = new List<string>() { paperId },
                owner = "30",
                multi = 100,
                start_time = sTime,
                end_time = sTime + this.paperTime,
                title = this.paperTitle,
                cost_time = this.paperTime
            };
            if(this.testId != null)
            {
                postData.id = this.testId;
            }
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
                    return (string)data["data"]["id"];
                }
            }
            return null;
        }

        private object createSingle(string aid, string selType, float score, string title, ArrayList labelArr, Hashtable optionMap, ArrayList curAnswerArr)
        {
            var titleArr = new List<PaperContent>
            {
                new PaperContent
                {
                    t = "text",
                    c = new PaperC()
                    {
                        text = title
                    }
                }
            };
            List<PaperSel> optionArr = new List<PaperSel>();
            labelArr.Sort();
            List<int> ans = new List<int>();
            int startIdx = 0;
            foreach (char obj in labelArr)
            {
                var curSel = new PaperContent
                {
                    t = "text",
                    c = new PaperC()
                    {
                        text = (string)optionMap["kx-text-" + obj]
                    }
                };
                optionArr.Add(new PaperSel()
                {
                    contents = new List<PaperContent> { curSel }
                });
                if (curAnswerArr.Contains(obj))
                {
                    ans.Add(startIdx);
                }
                startIdx++;
            }

            return (new
            {
                ver = 1,
                title = this.paperTitle,
                desc = this.paperTitle,
                id = aid,
                idx = new
                {
                    iids = new List<string>() { "l0", "l1" },
                },
                items = new
                {
                    l0 = new
                    {
                        t = "T1",
                        c = new
                        {
                            title = this.paperTitle,
                            allScore = score,
                            count = 1,
                        },
                        g = "_",
                        visiable = true
                    },
                    l1 = new
                    {
                        t = "e_sel",
                        c = new
                        {
                            t = selType,
                            titles = titleArr,
                            //analysis = null,
                            options = optionArr,
                            level = 1,
                            score = score,
                            answers = ans
                        },
                        g = "_",
                        visiable = true
                    }
                }
            });
        }

        private object createText(string aid, float score, string title, ArrayList curAnswerArr)
        {
            var titleArr = new List<PaperContent>
            {
                new PaperContent
                {
                    t = "text",
                    c = new PaperC()
                    {
                        text = title
                    }
                }
            };

            List<int> ans = new List<int>();

            return (new
            {
                ver = 1,
                title = this.paperTitle,
                desc = this.paperTitle,
                id = aid,
                idx = new
                {
                    iids = new List<string>() { "l0", "l1" },
                },
                items = new
                {
                    l0 = new
                    {
                        t = "T1",
                        c = new
                        {
                            title = this.paperTitle,
                            allScore = score,
                            count = 1,
                        },
                        g = "_",
                        visiable = true
                    },
                    l1 = new
                    {
                        t = "e_text",
                        c = new
                        {
                            titles = titleArr,
                            //analysis = null,
                            level = 1,
                            score = score,
                            answers = new List<object>()
                            {
                                new {
                                    t = "text",
                                    c = new
                                    {
                                        text = ""
                                    }
                                }
                            }
                        },
                        g = "_",
                        visiable = true
                    }
                }
            });
        }

        private object createFill(string aid, float score, string title, List<fillOption> curAnswerArr)
        {
            var titleArr = new List<PaperContent>
            {
                new PaperContent
                {
                    t = "text",
                    c = new PaperC()
                    {
                        text = title
                    }
                }
            };

            List<object> ans = new List<object>();
            foreach(var fillItem in curAnswerArr)
            {
                ans.Add(new
                {
                    t = "text",
                    c = new
                    {
                        text = fillItem.answer,
                        score = fillItem.score,
                        textArray = new List<string>() { fillItem.answer }
                    }
                });
            }

            return (new
            {
                ver = 1,
                title = this.paperTitle,
                desc = this.paperTitle,
                id = aid,
                idx = new
                {
                    iids = new List<string>() { "l0", "l1" },
                },
                items = new
                {
                    l0 = new
                    {
                        t = "T1",
                        c = new
                        {
                            title = this.paperTitle,
                            allScore = score,
                            count = 1,
                        },
                        g = "_",
                        visiable = true
                    },
                    l1 = new
                    {
                        t = "e_fill",
                        c = new
                        {
                            titles = titleArr,
                            //analysis = null,
                            level = 1,
                            score = score,
                            answers = ans
                        },
                        g = "_",
                        visiable = true
                    }
                }
            });
        }

        private void sendPaper(string paperId, string testId, object dataArgs)
        {
            Uri reqUrl = new Uri($"{utils.KXINFO.KXCOURSEURL}/art/api/update");
            Dictionary<string, string> args = new Dictionary<string, string> { };

            args.Add("aid", paperId);
            args.Add("token", utils.KXINFO.KXTOKEN);
            args.Add("uid", utils.KXINFO.KXUID);

            RestSharp.IRestResponse response = utils.request.SendRequest(Globals.ThisAddIn.CurHttpReq, reqUrl, RestSharp.Method.POST, args, dataArgs);
            if (response.ErrorException != null)
            {
                utils.Utils.LOG("art/api/update api error: " + response.ErrorException.Message);
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
                    utils.Utils.LOG("art/api/update api error: code " + code);
                }
                else
                {
                    //string testId = (string)data["data"]["id"];
                    sendKXOUT(paperId, testId);
                }
            }
        }

        private void sendPaperExt(string paperId, bool isQuestion)
        {
            Uri reqUrl = new Uri($"{utils.KXINFO.KXCOURSEURL}/usr/api/updateExt");
            Dictionary<string, string> args = new Dictionary<string, string> { };
            
            args.Add("aid", paperId);
            args.Add("token", utils.KXINFO.KXTOKEN);
            if(!isQuestion)
            {
                var extObj = JObject.FromObject(new
                {
                    advise_cost = this.paperTime,
                    es_auto_published = 1,
                    ease = 0
                });
                args.Add("ext", extObj.ToString());
            } else
            {
                args.Add("ext", "{\"folder\":\"\"}");
            }
            

            RestSharp.IRestResponse response = utils.request.SendRequest(Globals.ThisAddIn.CurHttpReq, reqUrl, RestSharp.Method.GET, args);
            if (response.ErrorException != null)
            {
                utils.Utils.LOG("usr/api/updateExt api error: " + response.ErrorException.Message);
                throw new Exception("updateExt api error");
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
                    utils.Utils.LOG("usr/api/updateExt api error: code " + code);
                    throw new Exception("updateExt api error");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("update ext success");
                }
            }
        }

        private void sendKXOUT(string paperId, string testId)
        {
            var recordArgs = new SendKxOut()
            {
                teach_record_id = Int64.Parse(utils.KXINFO.KXTCHRECORDID),
                content = new SendKxOutContent()
                {
                    paperId = paperId,
                    testId = testId
                },
                type = 200
            };
            Uri reqUrl = new Uri($"{utils.KXINFO.KXURL}/usr/upsertTeachResource?session_id={utils.KXINFO.KXSID}");
            Dictionary<string, string> args = new Dictionary<string, string> { };
            args.Add("session_id", utils.KXINFO.KXSID);
            List<SendKxOut> recordArgsArr = new List<SendKxOut> { recordArgs };
            RestSharp.IRestResponse response = utils.request.SendRequest(Globals.ThisAddIn.CurHttpReq, reqUrl, RestSharp.Method.POST, args, recordArgsArr);
            if (response.ErrorException != null)
            {
                utils.Utils.LOG("usr/upsertTeachResource api error: " + response.ErrorException.Message);
                throw new Exception("upsertTeachResource api error");
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
                    utils.Utils.LOG("usr/upsertTeachResource api error: code" + code);
                    throw new Exception("upsertTeachResource api error");
                }
                else
                {
                    string tidTmp = "";
                    try
                    {
                        tidTmp = (string)data["data"]["teach_resource_list"][0]["tid"];
                    }
                    catch (Exception e)
                    {
                        utils.Utils.LOG("read teach_resource_list.0.tid error");
                    }

                    object sendData = (new
                    {
                        key = "classroom",
                        value = utils.KXINFO.KXCHOSECLASSID,
                        type = "TEST",
                        data = new
                        {
                            course = new
                            {
                                title = utils.KXINFO.KXCHOSECOURSETITLE
                            },
                            exam = new
                            {
                                tid = tidTmp,
                                title = this.paperTitle,
                                content = new
                                {
                                    paperId = paperId,
                                    testId = testId
                                },
                                total = this.paperScore,
                                count = 1,
                                advise_cost = this.paperTime
                            }
                        },
                        timestamp = utils.Utils.getTimeStamp()
                    });
                    JObject o = JObject.FromObject(sendData);
                    string tmp = o.ToString();
                    Globals.ThisAddIn.SendTchInfo(tmp);
                    //recordTch(sendData);
                    utils.request.recordTch(sendData);
                }
            }
        }

        public void recordTch(object stepContent)
        {
            Uri reqUrl = new Uri($"{utils.KXINFO.KXURL}/usr/upsertTeachRecord");
            Dictionary<string, string> args = new Dictionary<string, string> { };
            args.Add("session_id", utils.KXINFO.KXSID);
            List<object> recordArgsArr = new List<object> {
                new
                {
                    class_id = utils.KXINFO.KXCHOSECLASSID,
                    course_id = utils.KXINFO.KXCHOSECOURSEID,
                    chapter_id = utils.KXINFO.KXCHOSECHAPTERID,
                    step_content = new List<object>(){stepContent }
                }
            };
            RestSharp.IRestResponse response = utils.request.SendRequest(Globals.ThisAddIn.CurHttpReq, reqUrl, RestSharp.Method.POST, args, recordArgsArr);
            if (response.ErrorException != null)
            {
                utils.Utils.LOG("usr/upsertTeachRecord api error: " + response.ErrorException.Message);
                throw new Exception("upsertTeachRecord api error");
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
                    utils.Utils.LOG("usr/upsertTeachRecord api error: code" + code);
                    throw new Exception("upsertTeachRecord api error");
                }
            }
        }
    }

    public class createExamInfo
    {
        public string title { get; set; }
        public string id { get; set; }
        public Int64 cost_time { get; set; }
        public int multi { get; set; }
        public string owner { get; set; }
        public List<string> aids { get; set; }
        public Int64 start_time { get; set; }
        public Int64 end_time { get; set; }
    }
}
