﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Newtonsoft.Json;

namespace kxrealtime.utils
{
    public static class pptContent
    {
        public static PowerPoint.Slide NewSlide()
        {
            var app = Globals.ThisAddIn.Application;
            PowerPoint.CustomLayout ppt_layout = app.ActivePresentation.SlideMaster.CustomLayouts[PowerPoint.PpSlideLayout.ppLayoutText];
            PowerPoint.Slide slide;
            int curSld = Globals.ThisAddIn.CurSlideIdx;
            slide = app.ActivePresentation.Slides.AddSlide(curSld + 1, ppt_layout);
            slide.Select();
            if (slide.Shapes.Count > 0)
            {
                slide.Shapes[1].Delete();
                slide.Shapes.Placeholders[1].Delete();
            }

            slide.Name = "kx-slide-" + slide.Name;
            return slide;
        }

        public static void InsertImage(string picUrl)
        {
            var slide = NewSlide();
            var shapeTmp = slide.Shapes.AddPicture(picUrl, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0);
            shapeTmp.Left = 100;
            shapeTmp.Top = 100;
        }

        public static void InserVideo(string videlUrl)
        {
            var slide = NewSlide();
            slide.Shapes.AddMediaObject2(videlUrl, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue);
        }

        public static void InserLink(string linkUrl)
        {
            var slide = NewSlide();
            var shapeTmp = slide.Shapes.AddTitle();
            var objText = shapeTmp.TextFrame.TextRange;
            objText.Text = linkUrl;
            objText.ActionSettings[Microsoft.Office.Interop.PowerPoint.PpMouseActivation.ppMouseClick].Hyperlink.Address = linkUrl;
        }

        public static void openPPT(string filePath)
        {
            //Utils.dlFile(filePath);
            //var objApp = new PowerPoint.Application();
            //Globals.ThisAddIn.Application.Presentations.Open(filePath);
            //new PowerPoint.Application().Presentations.Open(filePath);
        }

        public static void openWrold()
        {
            
        }

        public static void openExcel()
        {

        }

        public static void openFile(string pathTmp, string fileName)
        {
            var savePath = Utils.getFilePath();
            var filePath = savePath + @"\" + fileName;
            var task = Task.Run(() =>
            {
                Utils.dlFile(pathTmp, filePath);
                try
                {
                    System.Diagnostics.Process.Start(filePath);
                }
                catch (Exception e)
                {

                }
            });
            
        }

        public static void createPaperItem(string titleName, singleSelCtl.TypeSelEnum questionType, string stem = "此处插入描述", float score = 10, List<kxdata.simpleAnswerItem> answers =null, List<string> options = null)
        {
            var app = Globals.ThisAddIn.Application;
            var slide = NewSlide();
            Int32 curW = (Int32)app.ActivePresentation.SlideMaster.Width;
            Int32 curH = (Int32)app.ActivePresentation.SlideMaster.Height;

            PowerPoint.Shape sendBtn = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeActionButtonCustom, curW - 150, curH - 60, 100, 40);
            sendBtn.TextFrame.TextRange.InsertAfter("发送题目");
            sendBtn.Name = "kx-sending";
            sendBtn.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(1, 170, 170, 170).ToArgb();
            sendBtn.Line.ForeColor.RGB = System.Drawing.Color.FromArgb(1, 170, 170, 170).ToArgb();

            // 题干
            PowerPoint.Shape textBoxTitle = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal, 100, 100, curW - 120, 100);
            textBoxTitle.TextFrame.TextRange.InsertAfter(stem);
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
                scoreCom.TextFrame.TextRange.InsertAfter(score.ToString() + "分");
                scoreCom.Name = "kx-score";
            }

            Globals.ThisAddIn.initSetting();
            string curDir = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);
            var settingImg = curDir + @"\kxrealtime\imgs\setting.png";
            PowerPoint.Shape setBtn = slide.Shapes.AddPicture(settingImg, Office.MsoTriState.msoTrue, Office.MsoTriState.msoTrue, curW - 150, 0, 100, 40);
            setBtn.Name = "kx-setting";


            if (questionType == singleSelCtl.TypeSelEnum.singleSel || questionType == singleSelCtl.TypeSelEnum.multiSel || questionType == singleSelCtl.TypeSelEnum.voteSingleSel || questionType == singleSelCtl.TypeSelEnum.voteMultiSel)
            {
                var ans = new List<string>();
                foreach (var item in answers)
                {
                    ans.Add(item.text);
                }
                initOption(slide, options, questionType == singleSelCtl.TypeSelEnum.multiSel, ans);
            }
            else if (questionType == singleSelCtl.TypeSelEnum.textQuestion)
            {

            }
            else if (questionType == singleSelCtl.TypeSelEnum.fillQuestion && scoreCom != null)
            {
                scoreCom.TextFrame.TextRange.Text = score.ToString() + "分";
                var fillAns = new List<kxdata.simpleFillAnswer>();
                foreach (var item in answers)
                {
                    var fans = new kxdata.simpleFillAnswer
                    {
                        score = item.score,
                        answer = item.text
                    };
                    fillAns.Add(fans);
                }
                string output = JsonConvert.SerializeObject(fillAns);
                qInfo.TextFrame.TextRange.Text = output;
            }

            //slide.Select();
        }

        public static void initOption(PowerPoint.Slide slide,List<string> options, bool isMul, List<string> ans)
        {
            char sChar = 'A';
            int posY = 200;
            int n = options.Count;
            float difY = (250 - n * 50) / (n - 1);
            Office.MsoAutoShapeType curShapeType = !isMul ? Office.MsoAutoShapeType.msoShapeOval : Office.MsoAutoShapeType.msoShapeRectangle;
            for (int i = 0; i < n; i++)
            {
                char curChar = (char)(sChar + i);
                PowerPoint.Shape circleTmp = slide.Shapes.AddShape(curShapeType, 100, posY + difY * i - 5, 40, 40);
                circleTmp.TextFrame.TextRange.InsertAfter(curChar.ToString());
                circleTmp.Name = "kx-choice-" + curChar.ToString();
                var colorTmp = System.Drawing.Color.FromArgb(1, 128, 128, 128).ToArgb();
                if (ans.Contains(curChar.ToString()))
                {
                    colorTmp = KXINFO.ChoseColor;
                }
                circleTmp.Fill.ForeColor.RGB = colorTmp;
                circleTmp.Line.ForeColor.RGB = colorTmp;
                PowerPoint.Shape textBox = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal, 150, posY + difY * i, 500, 50);
                var optionText = options[i];
                textBox.TextFrame.TextRange.InsertAfter(optionText);
                textBox.Name = "kx-text-" + curChar.ToString();
                posY += 50;
            }
        }
    }
}
