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

namespace kxrealtime
{
    [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public partial class kxResource : UserControl
    {
        public kxResource()
        {
            InitializeComponent();
            initWeb();
        }

        private void initWeb()
        {
            //string urlTmp = $"{utils.KXINFO.KXADMINURL}/?session_id={utils.KXINFO.KXSID}&timestamp={utils.Utils.getTimeStamp()}&token={utils.KXINFO.KXTOKEN}#/pptComponents/resourceLibrary?teach_record_id={utils.KXINFO.KXTCHRECORDID}&session_id={utils.KXINFO.KXSID}";
            string urlTmp = $"{utils.KXINFO.KXADMINURL}/?token={utils.KXINFO.KXTOKEN}&timestamp={utils.Utils.getTimeStamp()}#/pptComponents/resourceLibrary?token={utils.KXINFO.KXTOKEN}&session_id={utils.KXINFO.KXSID}";
            resourceWebBrowser.Navigate(urlTmp);
            resourceWebBrowser.ObjectForScripting = this;
        }

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


        public void showFile(string fileLink, string fileName)
        {
            utils.pptContent.openFile(fileLink, fileName);
        }

        public void showImage(string imgLink)
        {
            utils.pptContent.InsertImage(imgLink);
        }

        public void showVideo(string videoLink)
        {
            utils.pptContent.InserVideo(videoLink);
        }

        public void showLink(string link, string name)
        {
            utils.pptContent.InserLink(link);
        }

        public void showQuestion()
        {

        }
    }
}
