using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace kxrealtime
{
    public partial class choseClass : Form
    {
        private Int64 classID;
        private Int64 courseID;
        private Int64 chapterID;
        private string className;
        private string courseName;

        public choseClass()
        {
            InitializeComponent();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            var curbox = sender as ComboBox;
            if(curbox == null)
            {
                return;
            }
            var courseItem = (ClassItem)curbox.SelectedItem;
            if (courseItem == null)
            {
                return;
            }
            classID = courseItem.tid;
            className = courseItem.name;
            initChapter();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            var curbox = sender as ComboBox;
            if (curbox == null)
            {
                return;
            }
            var courseItem = (CourseItem)curbox.SelectedItem;
            if (courseItem == null)
            {
                return;
            }
            courseID = courseItem.tid;
            courseName = courseItem.title;
            initChapter();
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            var curbox = sender as ComboBox;
            if (curbox == null)
            {
                return;
            }
            var courseItem = (ChapterItem)curbox.SelectedItem;
            if(courseItem == null)
            {
                return;
            }
            chapterID = courseItem.tid;
        }


        public void initClassList()
        {
             Uri reqUrl = new Uri($"{utils.KXINFO.KXURL}/usr/listClass?skip=0&limit=1000&ret_teach_record=1&session_id={utils.KXINFO.KXSID}");
             RestSharp.IRestResponse response = utils.request.SendRequest(Globals.ThisAddIn.CurHttpReq, reqUrl, RestSharp.Method.GET);
             if (response.ErrorException != null)
             {
                utils.Utils.LOG("usr/listClass api error: " + response.ErrorException.Message);
            } else
             {
                JObject data = JObject.Parse(response.Content);
                string code = (string)data["code"];
                if(code != "0")
                {
                    System.Diagnostics.Debug.WriteLine("initclass" + code);
                    utils.Utils.LOG("usr/listClass api error: code" + code);
                } else
                {
                    try
                    {
                        comboBox1.Items.Clear();
                        JArray listStr = (JArray)data["data"]["list"];
                        List<ClassItem> listArr = listStr.Select(c => c.ToObject<ClassItem>()).ToList();
                        foreach (ClassItem tmp in listArr)
                        {
                            comboBox1.Items.Add(tmp);
                        }
                        comboBox1.ValueMember = "tid";
                        comboBox1.DisplayMember = "name";
                        if(listArr.Count == 0)
                        {
                            this.showTip("暂无课程信息，请前往酷课堂添加课程");
                        }
                    }catch(Exception)
                    {
                        this.showTip("暂无课程信息，请前往酷课堂添加课程");
                    }
                    
                }
             }
        }

        public void initCourseList()
        {
            Uri reqUrl = new Uri($"{utils.KXINFO.KXURL}/usr/listCourse?skip=0&limit=1000&session_id={utils.KXINFO.KXSID}");
            RestSharp.IRestResponse response = utils.request.SendRequest(Globals.ThisAddIn.CurHttpReq, reqUrl, RestSharp.Method.GET);
            if (response.ErrorException != null)
            {
                utils.Utils.LOG("usr/listCourse api error: " + response.ErrorException.Message);
            }
            else
            {
                JObject data = JObject.Parse(response.Content);
                string code = (string)data["code"];
                if (code != "0")
                {
                    if(code == "401")
                    {
                        this.Visible = false;
                        Globals.ThisAddIn.loginOut();
                        MessageBox.Show("登录失效请重新登录");
                    }
                    utils.Utils.LOG("usr/listCourse api error: code" + code);
                }
                else
                {
                    try
                    {
                        comboBox2.Items.Clear();
                        JArray listStr = (JArray)data["data"]["list"];
                        List<CourseItem> listArr = listStr.Select(c => c.ToObject<CourseItem>()).ToList();
                        foreach (CourseItem tmp in listArr)
                        {
                            comboBox2.Items.Add(tmp);
                        }
                        comboBox2.ValueMember = "tid";
                        comboBox2.DisplayMember = "title";
                        if (listArr.Count == 0)
                        {
                            this.showTip("暂无班级信息，请前往酷课堂添加班级");
                        }
                    }
                    catch(Exception)
                    {
                        this.showTip("暂无班级信息，请前往酷课堂添加班级");
                    }
                    
                }
            }
        }

        public void initChapter()
        {
            string curClassId = classID.ToString();
            string curCourseId = courseID.ToString();
            if(curCourseId.Length == 0 || curClassId.Length == 0 || curCourseId == "0" || curClassId == "0")
            {
                return;
            }
            Dictionary<string, string> args = new Dictionary<string, string> { };
            args.Add("skip", "0");
            args.Add("limit", "1000");
            args.Add("session_id", utils.KXINFO.KXSID);
            args.Add("class_id", curClassId);
            args.Add("course_id", curCourseId);
            Uri reqUrl = new Uri($"{utils.KXINFO.KXURL}/usr/listChapter");
            RestSharp.IRestResponse response = utils.request.SendRequest(Globals.ThisAddIn.CurHttpReq, reqUrl, RestSharp.Method.GET, args);
            if (response.ErrorException != null)
            {
                utils.Utils.LOG("usr/listChapter api error: " + response.ErrorException.Message);
            }
            else
            {
                JObject data = JObject.Parse(response.Content);
                string code = (string)data["code"];
                if (code != "0")
                {
                    System.Diagnostics.Debug.WriteLine("initchapter" + code);
                    utils.Utils.LOG("usr/listChapter api error: code" + code);
                }
                else
                {
                    try
                    {
                        comboBox3.Items.Clear();
                        JArray listStr = (JArray)data["data"]["list"];
                        List<ChapterItem> listArr = listStr.Select(c => c.ToObject<ChapterItem>()).ToList();
                        foreach (ChapterItem tmp in listArr)
                        {
                            comboBox3.Items.Add(tmp);
                        }
                        comboBox3.ValueMember = "tid";
                        comboBox3.DisplayMember = "title";
                        if (listArr.Count == 0)
                        {
                            this.showTip("暂无课时信息，请前往酷课堂添加课时");
                        }
                    }
                    catch(Exception) {
                        this.showTip("暂无课时信息，请前往酷课堂添加课时");
                    }
                    
                }
            }
        }


        private void button1_Click(object sender, EventArgs e)
        {
            this.Visible = false;
        }

        private void showTip(string text)
        {
            panel1.Controls.Clear();
            System.Windows.Forms.Label label = new System.Windows.Forms.Label();
            label.Text = text;
            label.ForeColor = System.Drawing.Color.Red;
            label.Visible = true;
            label.Width = 200;
            label.AutoSize = true;
            panel1.Controls.Add(label);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(chapterID == 0 || classID == 0 || courseID == 0)
            {
                this.showTip("请选择上述内容");
                return;
            }
            this.loadingBox.Visible = true;
            this.loadingBox.Dock = DockStyle.Fill;
            var curArgs = new SendArgs();
            curArgs.chapter_id = chapterID;
            curArgs.class_id = classID;
            curArgs.course_id = courseID;
            curArgs.status = 100;
            List<SendArgs> curArgsArr = new List<SendArgs> { curArgs };
           
            Dictionary<string, string> queryArgs = new Dictionary<string, string> { };
            queryArgs.Add("session_id", utils.KXINFO.KXSID);
            var reqUrl = new Uri($"{utils.KXINFO.KXURL}/usr/upsertTeachRecord?session_id={utils.KXINFO.KXSID}");
            RestSharp.IRestResponse response = utils.request.SendRequest(Globals.ThisAddIn.CurHttpReq, reqUrl, RestSharp.Method.POST, queryArgs, curArgsArr);
            this.loadingBox.Visible = false ;
            if (response.ErrorException != null)
            {
                utils.Utils.LOG("usr/upsertTeachRecord api error: " + response.ErrorException.Message);
            }
            else
            {

                JObject data = JObject.Parse(response.Content);
                string code = (string)data["code"];
                if (code != "0")
                {
                    System.Diagnostics.Debug.WriteLine("send tch record" + code);
                    utils.Utils.LOG("usr/upsertTeachRecord api error: code" + code);
                }
                else
                {
                    utils.KXINFO.KXCHOSECLASSID = classID;
                    utils.KXINFO.KXCHOSECOURSEID = courseID;
                    utils.KXINFO.KXCHOSECHAPTERID = chapterID;
                    utils.KXINFO.KXCHOSECOURSETITLE = courseName;
                    utils.KXINFO.KXCHOSECLASSNAME = className;
                    try
                    {
                        utils.KXINFO.KXTCHRECORDID = (string)data["data"]["teach_record_list"][0]["tid"];
                    }catch(Exception)
                    {
                        utils.Utils.LOG("read teach_record_list.0.tid error");
                    }
                    
                    this.Visible = false;
                    Globals.ThisAddIn.InitTchSocket();
                    createScreenWin(classID.ToString(), className);
                }
            }
        }

        private void createScreenWin(string classId, string className)
        {
            Globals.ThisAddIn.canShowAddClass = true;
            Globals.ThisAddIn.Application.ActivePresentation.SlideShowSettings.Run();
            return;
        }

        public void stopTching()
        {
            var curArgs = new SendArgs();
            curArgs.chapter_id = chapterID;
            curArgs.class_id = classID;
            curArgs.course_id = courseID;
            curArgs.status = 200;
            curArgs.tid = Int64.Parse(utils.KXINFO.KXTCHRECORDID);
            curArgs.user_id = Int64.Parse(utils.KXINFO.KXOUTUID);
            List<SendArgs> curArgsArr = new List<SendArgs> { curArgs };

            Dictionary<string, string> queryArgs = new Dictionary<string, string> { };
            queryArgs.Add("session_id", utils.KXINFO.KXSID);
            var reqUrl = new Uri($"{utils.KXINFO.KXURL}/usr/upsertTeachRecord?session_id={utils.KXINFO.KXSID}");
            RestSharp.IRestResponse response = utils.request.SendRequest(Globals.ThisAddIn.CurHttpReq, reqUrl, RestSharp.Method.POST, queryArgs, curArgsArr);
            if (response.ErrorException != null)
            {
                utils.Utils.LOG("usr/upsertTeachRecord api error: " + response.ErrorException.Message);
                throw new Exception("error");
            }
            else
            {
                try
                {
                    JObject data = JObject.Parse(response.Content);
                    string code = (string)data["code"];
                    if (code != "0")
                    {
                        System.Diagnostics.Debug.WriteLine("send stop tch record" + code);
                        utils.Utils.LOG("usr/upsertTeachRecord api error: code" + code);
                        throw new Exception("error");
                    }
                }catch(JsonReaderException)
                {
                    utils.Utils.LOG("close course error upsertTeachRecord, received data not json");
                    throw new Exception("error");
                }
                
            }
        }
    }


    public class ClassItem {
        public Int64 tid { get; set; }
        public Int64 user_id { get; set; }
        public string name { get; set; }
    };

    public class CourseItem
    {
        public Int64 tid { get; set; }
        public Int64 user_id { get; set; }
        public string title { get; set; }
    }

    public class ChapterItem
    {
        public Int64 tid { get; set; }
        public Int64 course_id { get; set; }
        public string title { get; set; }
    }

    public class SendArgs
    {
        public Int64 course_id { get; set; }
        public Int64 class_id { get; set; }
        public Int64 chapter_id { get; set; }

        public Int64 user_id { get; set; }

        public Int64 tid { get; set; }

        public int status { get; set; }
    }

}
