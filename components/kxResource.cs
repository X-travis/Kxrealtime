using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Security.Permissions;
using Newtonsoft.Json;
using System.Threading;
using System.Diagnostics;

namespace kxrealtime
{
    [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public partial class kxResource : UserControl
    {
        // 构造函数
        public kxResource()
        {
            InitializeComponent();
            initWeb();
        }

        // 初始化网页
        private void initWeb()
        {
            //string urlTmp = "http://192.168.19.168:8080" + $"/?token=BBE356F4817DE17FCAB90A1CC45CE960&timestamp={utils.Utils.getTimeStamp()}#/pptComponents/resourceLibrary?session_id=5ea0fbe157c9f5119e00002d&token=BBE356F4817DE17FCAB90A1CC45CE960";
            string urlTmp = $"{utils.KXINFO.KXADMINURL}/?token={utils.KXINFO.KXTOKEN}&timestamp={utils.Utils.getTimeStamp()}#/pptComponents/resourceLibrary?token={utils.KXINFO.KXTOKEN}&session_id={utils.KXINFO.KXSID}";
            resourceWebBrowser.Navigate(urlTmp);
            resourceWebBrowser.ObjectForScripting = this;
            resourceWebBrowser.Visible = true;
        }

        // 弃用
        public void showPaper(string data)
        {
            var paperInfo = JsonConvert.DeserializeObject<kxdata.simplePaper>(data);
            bool isQ = paperInfo.type == "question";
            var paperList = paperInfo.data;
            foreach(var pItem in paperList)
            {
                var itemType = pItem.type;
                string titleName;
                singleSelCtl.TypeSelEnum qType;
                float score;
                switch (itemType)
                {
                    case "single":
                        titleName = "单选题";
                        qType = isQ ? singleSelCtl.TypeSelEnum.voteSingleSel : singleSelCtl.TypeSelEnum.singleSel;
                        utils.pptContent.createPaperItem(titleName, qType, pItem.title, pItem.score, pItem.answers, pItem.options);
                        break;
                    case "multi":
                        titleName = "多选题";
                        qType = isQ ? singleSelCtl.TypeSelEnum.voteMultiSel : singleSelCtl.TypeSelEnum.multiSel;
                        utils.pptContent.createPaperItem(titleName, qType, pItem.title, pItem.score, pItem.answers, pItem.options);
                        break;
                    case "e_fill":
                        titleName = "填空题";
                        qType = singleSelCtl.TypeSelEnum.fillQuestion;
                        utils.pptContent.createPaperItem(titleName, qType, pItem.title, pItem.score, pItem.answers);
                        break;
                    case "e_text":
                        titleName = "简答题";
                        qType = singleSelCtl.TypeSelEnum.textQuestion;
                        utils.pptContent.createPaperItem(titleName, qType, pItem.title, pItem.score);
                        break;
                }
            }
        }

        // 弃用
        public void showFile(string fileLink, string fileName, string type)
        {
            showFilePathTip(fileName);
            utils.pptContent.openFile(fileLink, fileName, type, isShowProgress, changeProgress);
        }

        // 显示保存地址
        private void showFilePathTip(string fileName)
        {
            var savePath = utils.Utils.getFilePath();
            var filePath = savePath + @"\" + fileName;
          //  this.savePathLabel.Text = $"保存地址：{filePath}";
        }

        // 是否显示进度条
        public void isShowProgress(bool flag)
        {
            Action<bool> action = (bool isShow) =>
            {
                this.progresslabel.Text = "下载进度：0%";
                this.fileLoading.Visible = isShow;
            };
            this.Invoke(action, flag);
            
        }

        // 改变当前进度
        public void changeProgress(double value)
        {
            Action<double> action = (double curPer) =>
            {
                int curPerTmp = (int)(100 * curPer);
                if(curPerTmp > 100)
                {
                    curPerTmp = 100;
                }
                this.progresslabel.Text = "下载进度：" + (curPerTmp).ToString() + "%";
                int pg =  (Int32)(curPer * 100) % 10;
                // 优化
                if(pg%2 == 0)
                {
                    System.Windows.Forms.Application.DoEvents();
                }
            };
            this.Invoke(action, value);
        }

        // 提供给web的调用方法 插入图片
        public void showImage(string imgLink)
        {
            //utils.pptContent.InsertImage(imgLink);
            var nameArr = imgLink.Split('/');
            var curName = nameArr[nameArr.Length - 1];
            showFilePathTip(curName);
            utils.pptContent.openFile(imgLink, curName, "image", isShowProgress, changeProgress);
        }

        // 提供给web的调用方法 插入视频
        public void showVideo(string videoLink)
        {
           // Console.WriteLine(videoLink);
            //utils.pptContent.InserVideo(videoLink);
            var nameArr = videoLink.Split('/');
            var curName = nameArr[nameArr.Length - 1];
            showFilePathTip(curName);
            utils.pptContent.openFile(videoLink, curName, "video", isShowProgress, changeProgress);
           
        }

        // 提供给web的调用方法 插入链接
        public void showLink(string link, string name, string info)
        {
            utils.pptContent.InserLink(link, name, info);
        }

        public void showQuestion()
        {

        }

        private void savePathLabel_Click(object sender, EventArgs e)
        {

        }

        private void progresslabel_Click(object sender, EventArgs e)
        {

        }
    }
}
