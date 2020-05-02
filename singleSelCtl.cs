using Newtonsoft.Json;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace kxrealtime
{
    public partial class singleSelCtl : UserControl
    {

        public enum TypeSelEnum
        {
            singleSel,
            multiSel,
            voteSingleSel,
            voteMultiSel,
            fillQuestion,
            textQuestion
        }

        public singleSelCtl()
        {
            InitializeComponent();
        }

        private Panel selectPanel;
        private Button selectBtn;
        private TypeSelEnum curType;

        private ArrayList curSelLabelArr;

        private List<fillOption> fillOptionArr;

        public TypeSelEnum setCurSelType
        {
            get
            {
                return this.curType;
            }
            set
            {
                this.curType = value;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox comboBox = (ComboBox)sender;
            //Globals.ThisAddIn.Application.ActiveWindow.Panes[2].Activate();

            //int currentSlideIndex = Globals.ThisAddIn.CurSlideIdx;
            //PowerPoint.Shapes curShapes = Globals.ThisAddIn.Application.ActivePresentation.Slides[currentSlideIndex].Shapes;
            //foreach (PowerPoint.Shape shapeTmp in curShapes)
            //{
            //    System.Diagnostics.Debug.WriteLine(shapeTmp.TextFrame.TextRange.Text);
            //}
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            NumericUpDown numericUpDown = (NumericUpDown)sender;
            System.Diagnostics.Debug.WriteLine(numericUpDown.Value);
            int currentSlideIndex = Globals.ThisAddIn.CurSlideIdx;
            PowerPoint.Shapes curShapes = Globals.ThisAddIn.Application.ActivePresentation.Slides[currentSlideIndex].Shapes;
            foreach (PowerPoint.Shape shapeTmp in curShapes)
            {
                string targetName = "kx-score";
                if (shapeTmp.Name == targetName)
                {
                    shapeTmp.TextFrame.TextRange.Text = (numericUpDown.Value).ToString() + '分';
                    //shapeTmp.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(1, 0, 255, 0).ToArgb();
                    //shapeTmp.Line.ForeColor.RGB = System.Drawing.Color.FromArgb(1, 128, 128, 128).ToArgb();
                }

            }
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            ArrayList ansTmp = new ArrayList();
            char nextValue = FindNextSel();
            ArrayList addList = new ArrayList();
            addList.Add(nextValue);
            this.createRadio(addList, ansTmp);
            this.curSelLabelArr.Add(nextValue);
            this.resetOption(nextValue);
        }

        private char FindNextSel()
        {
            char result = (char)this.curSelLabelArr[0];
            for (int i = 1; i < this.curSelLabelArr.Count; i++)
            {
                char tmp = (char)this.curSelLabelArr[i];
                if (tmp > result)
                {
                    result = tmp;
                }
            }
            return (char)(result + 1);
        }

        public void createRadio(ArrayList labelArr, ArrayList ans)
        {
            int difY = 35;
            int difX = 50;
            int oldNum = this.selectPanel.Controls.Count - 1;
            int len = labelArr.Count;
            int x = (oldNum % 4) * difX;
            int y = (oldNum / 4) * difY;
            bool isMul = this.curType == TypeSelEnum.multiSel;

            for (int i = 0; i < len; i++)
            {
                char tmp = (char)(labelArr[i]);
                if (isMul)
                {
                    System.Windows.Forms.CheckBox cTmp = this.createCheckBtn(x, y, tmp.ToString(), ans.Contains(tmp));
                    this.selectPanel.Controls.Add(cTmp);
                }
                else
                {
                    System.Windows.Forms.RadioButton radioTmp = this.createRadioBtn(x, y, tmp.ToString(), ans.Contains(tmp));
                    this.selectPanel.Controls.Add(radioTmp);
                }

                x += difX;
                if (i != 0 && ((i + 1) % 4) == 0)
                {
                    x = 0;
                    y += difY;
                }
            }
            this.selectBtn.Location = new System.Drawing.Point(x, y);
        }

        private System.Windows.Forms.RadioButton createRadioBtn(int posX, int posY, string text, bool isChecked)
        {
            System.Windows.Forms.RadioButton radioTmp = new System.Windows.Forms.RadioButton();
            radioTmp.Location = new System.Drawing.Point(posX, posY);
            radioTmp.Size = new System.Drawing.Size(50, 15);
            radioTmp.Text = text;
            radioTmp.CheckedChanged += new EventHandler(this.radioButton_CheckedChanged);
            radioTmp.Checked = isChecked;
            radioTmp.Height = 30;
            return radioTmp;
        }

        private System.Windows.Forms.CheckBox createCheckBtn(int posX, int posY, string text, bool isChecked)
        {
            System.Windows.Forms.CheckBox checkTmp = new System.Windows.Forms.CheckBox();
            checkTmp.Location = new System.Drawing.Point(posX, posY);
            checkTmp.Size = new System.Drawing.Size(50, 15);
            checkTmp.Text = text;
            checkTmp.CheckedChanged += new EventHandler(this.checkButton_CheckedChanged);
            checkTmp.Checked = isChecked;
            checkTmp.Height = 30;
            return checkTmp;
        }

        private void resetCurType()
        {
            int currentSlideIndex = Globals.ThisAddIn.CurSlideIdx;
            PowerPoint.Shapes curShapes = Globals.ThisAddIn.Application.ActivePresentation.Slides[currentSlideIndex].Shapes;
            foreach (PowerPoint.Shape shapeTmp in curShapes)
            {
                string targetName = "kx-title-";
                if (shapeTmp.Name == (targetName + TypeSelEnum.singleSel))
                {
                    this.curType = TypeSelEnum.singleSel;
                }
                else if (shapeTmp.Name == (targetName + TypeSelEnum.multiSel))
                {
                    this.curType = TypeSelEnum.multiSel;
                }
                else if (shapeTmp.Name == (targetName + TypeSelEnum.voteSingleSel))
                {
                    this.curType = TypeSelEnum.voteSingleSel;
                }
                else if (shapeTmp.Name == (targetName + TypeSelEnum.voteMultiSel))
                {
                    this.curType = TypeSelEnum.voteMultiSel;
                }
                else if (shapeTmp.Name == (targetName + TypeSelEnum.fillQuestion))
                {
                    this.curType = TypeSelEnum.fillQuestion;
                }
                else if (shapeTmp.Name == (targetName + TypeSelEnum.textQuestion))
                {
                    this.curType = TypeSelEnum.textQuestion;
                }

            }
        }

        public void resetData(float score, ArrayList ans, ArrayList labelArr)
        {
            this.resetCurType();
            this.changePannelShow(this.curType);
            this.curSelLabelArr = labelArr;

            if (this.curType == TypeSelEnum.singleSel || this.curType == TypeSelEnum.multiSel)
            {
                resetSelData(score, ans, labelArr);
            }
            else if (this.curType == TypeSelEnum.voteMultiSel || this.curType == TypeSelEnum.voteSingleSel)
            {
                resetVoteData();
            }


        }

        private void resetVoteData()
        {
            this.voteMBtn.Checked = this.curType == TypeSelEnum.voteMultiSel;
            this.voteSBtn.Checked = this.curType == TypeSelEnum.voteSingleSel;
        }

        private void resetSelData(float score, ArrayList ans, ArrayList labelArr)
        {
            numericUpDown1.Value = (decimal)score;
            if (this.selectPanel != null)
            {
                panel2.Controls.Remove(this.selectPanel);
            }
            this.selectPanel = new Panel();
            var upLocation = label3.Location;
            selectPanel.Location = new Point(upLocation.X, upLocation.Y + label3.Height * 3 / 2);
            selectPanel.Size = new Size(400, 100);
            this.selectPanel.AutoScroll = true;
            this.selectBtn = new Button();
            this.selectBtn.Click += this.button1_Click;
            this.selectBtn.Text = "+";
            this.selectBtn.Width = 30;
            this.selectBtn.FlatStyle = FlatStyle.Flat;

            this.selectPanel.Controls.Add(this.selectBtn);
            panel2.Controls.Add(selectPanel);
            this.createRadio(labelArr, ans);

            if (this.curType == TypeSelEnum.singleSel)
            {
                this.comboBox1.SelectedItem = "单选题";
            }
            else
            {
                this.comboBox1.SelectedItem = "多选题";
            }
        }


        private void radioButton_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton rb = sender as RadioButton;

            if (rb == null)
            {
                MessageBox.Show("Sender is not a RadioButton");
                return;
            }

            // Ensure that the RadioButton.Checked property
            // changed to true.
            if (rb.Checked)
            {
                //System.Diagnostics.Debug.WriteLine(rb.Text);
                int currentSlideIndex = Globals.ThisAddIn.CurSlideIdx;
                PowerPoint.Slide Sld = Globals.ThisAddIn.Application.ActivePresentation.Slides[currentSlideIndex];
                PowerPoint.Shapes curShapes = Sld.Shapes;
                foreach (PowerPoint.Shape shapeTmp in curShapes)
                {
                    string targetName = "kx-choice-" + rb.Text;
                    if (shapeTmp.Name == targetName)
                    {
                        shapeTmp.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(1, 0, 255, 0).ToArgb();
                        //shapeTmp.Line.ForeColor.RGB = System.Drawing.Color.FromArgb(1, 128, 128, 128).ToArgb();
                    }
                    else if (shapeTmp.Name.Contains("kx-choice"))
                    {
                        shapeTmp.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(1, 128, 128, 128).ToArgb();
                    }
                }
            }
        }

        private void checkButton_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox rb = sender as CheckBox;

            if (rb == null)
            {
                MessageBox.Show("Sender is not a checkButton");
                return;
            }

            int currentSlideIndex = Globals.ThisAddIn.CurSlideIdx;
            PowerPoint.Shapes curShapes = Globals.ThisAddIn.Application.ActivePresentation.Slides[currentSlideIndex].Shapes;
            foreach (PowerPoint.Shape shapeTmp in curShapes)
            {
                string targetName = "kx-choice-" + rb.Text;
                if (shapeTmp.Name == targetName)
                {
                    if (rb.Checked)
                    {
                        shapeTmp.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(1, 0, 255, 0).ToArgb();
                    }
                    else
                    {
                        shapeTmp.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(1, 128, 128, 128).ToArgb();
                    }

                    //shapeTmp.Line.ForeColor.RGB = System.Drawing.Color.FromArgb(1, 128, 128, 128).ToArgb();
                }

            }
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }


        // 用于增加选项
        private void resetOption(char nextChar)
        {
            char startChar = 'A';
            int i = this.curSelLabelArr.Count - 1; // this.selectPanel.Controls.Count - 1;
            int currentSlideIndex = Globals.ThisAddIn.CurSlideIdx;
            PowerPoint.Shapes curShapes = Globals.ThisAddIn.Application.ActivePresentation.Slides[currentSlideIndex].Shapes;

            int posIdx = 1;
            Hashtable shapeMap = new Hashtable();
            bool isMul = this.curType == TypeSelEnum.multiSel || this.curType == TypeSelEnum.voteMultiSel;
            foreach (PowerPoint.Shape shapeTmp in curShapes)
            {
                if (shapeMap.ContainsKey(shapeTmp.Name))
                {

                }
                else
                {
                    shapeMap.Add(shapeTmp.Name, posIdx);
                    posIdx += 1;
                }
            }
            var app = Globals.ThisAddIn.Application;
            Int32 curH = (Int32)app.ActivePresentation.SlideMaster.Height;
            float posY = 200;
            
            char curChar = (char)(startChar);
            int optionH = curH > 450 ? 50 : 40;
            int selectCtxHeight = curH - 240;
            // 计算新增一个，每个选项之间的间隔
            float difY = (selectCtxHeight - (i+1) * optionH) / (i);
            int difNum = 0;
            // 从A往后检查100个， 或者满足当前个数
            for (int j = 0; j < 100 && difNum < i; j++)
            {
                string choiceKeyTmp = "kx-choice-" + curChar.ToString();
                string choiceText = "kx-text-" + curChar.ToString();
                bool hasFound = false;

                if (shapeMap.ContainsKey(choiceKeyTmp))
                {
                    int keyIdx = (int)shapeMap[choiceKeyTmp];
                    curShapes[keyIdx].Top = difY * difNum + posY - 5;
                    hasFound = true;

                }
                if (shapeMap.ContainsKey(choiceText))
                {
                    int textIdx = (int)shapeMap[choiceText];
                    curShapes[textIdx].Top = difY * difNum + posY;
                    hasFound = true;

                }
                if (difNum != i - 1)
                {
                    curChar = (char)(startChar + difNum + 1);
                }
                if (hasFound)
                {
                    posY += optionH;
                    difNum += 1;
                }
            }
            posY += difY * i;
            Office.MsoAutoShapeType curShapeType = !isMul ? Office.MsoAutoShapeType.msoShapeOval : Office.MsoAutoShapeType.msoShapeRectangle;
            PowerPoint.Shape circleTmp = curShapes.AddShape(curShapeType, 100, posY - 5, optionH - 10, optionH-10);
            circleTmp.TextFrame.TextRange.InsertAfter(nextChar.ToString());
            circleTmp.Name = "kx-choice-" + nextChar.ToString();
            circleTmp.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(1, 128, 128, 128).ToArgb();
            circleTmp.Line.ForeColor.RGB = System.Drawing.Color.FromArgb(1, 128, 128, 128).ToArgb();
            PowerPoint.Shape textBox = curShapes.AddTextbox(
            Office.MsoTextOrientation.msoTextOrientationHorizontal, 150, posY, 500, optionH);
            textBox.TextFrame.TextRange.InsertAfter("此处插入描述");
            textBox.Name = "kx-text-" + nextChar.ToString();
        }

        public void initSubjectiveQ(float score, bool fillScore = false)
        {
            this.changePannelShow(TypeSelEnum.textQuestion);
            if(fillScore)
            {
                numericUpDown2.Value = (decimal)score;
            }
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            numericUpDown1_ValueChanged(sender, e);
        }

        public void initVoteQ(TypeSelEnum curType)
        {
            this.changePannelShow(curType);
        }

        private void changePannelShow(TypeSelEnum curType)
        {
            this.panel2.Visible = curType == TypeSelEnum.singleSel || curType == TypeSelEnum.multiSel;
            this.panel3.Visible = curType == TypeSelEnum.textQuestion;
            this.panel4.Visible = curType == TypeSelEnum.voteSingleSel || curType == TypeSelEnum.voteMultiSel;
            this.panel5.Visible = curType == TypeSelEnum.fillQuestion;
            this.panel3.Top = 30;
            this.panel4.Top = 30;
            this.panel5.Top = 30;

        }

        private void addVoteOption_Click(object sender, EventArgs e)
        {
            ArrayList ansTmp = new ArrayList();
            char nextValue = FindNextSel();
            ArrayList addList = new ArrayList();
            addList.Add(nextValue);
            this.curSelLabelArr.Add(nextValue);
            this.resetOption(nextValue);
        }

        private void voteSBtn_CheckedChanged(object sender, EventArgs e)
        {
            var curObj = sender as RadioButton;
            if (curObj.Checked && this.curType != TypeSelEnum.voteSingleSel)
            {
                this.curType = TypeSelEnum.voteSingleSel;
                changeVoteType();
            }
        }

        private void voteMBtn_CheckedChanged(object sender, EventArgs e)
        {
            var curObj = sender as RadioButton;
            if(curObj.Checked && this.curType != TypeSelEnum.voteMultiSel)
            {
                this.curType = TypeSelEnum.voteMultiSel;
                changeVoteType();
            }
        }

        private void changeVoteType()
        {
            int currentSlideIndex = Globals.ThisAddIn.CurSlideIdx;
            PowerPoint.Shapes curShapes = Globals.ThisAddIn.Application.ActivePresentation.Slides[currentSlideIndex].Shapes;
            bool isMul = this.curType == TypeSelEnum.voteMultiSel;
            foreach (PowerPoint.Shape shapeTmp in curShapes)
            {
                if (shapeTmp.Name.Contains("kx-title"))
                {
                    shapeTmp.Name = "kx-title-" + this.curType;
                }
                if (shapeTmp.Name.Contains("kx-choice"))
                {

                    Office.MsoAutoShapeType curShapeType = !isMul ? Office.MsoAutoShapeType.msoShapeOval : Office.MsoAutoShapeType.msoShapeRectangle;
                    shapeTmp.AutoShapeType = curShapeType;

                }
            }
        }

        public void initFillQ(float score)
        {
            this.changePannelShow(TypeSelEnum.fillQuestion);
            //numericUpDown2.Value = (decimal)score;
        }

        private void fillAddBtn_Click(object sender, EventArgs e)
        {
            var tmp = new List<fillOption>()
            {
                new fillOption{
                    score = 0,
                    answer = ""
                }
            };
            addFillOption(tmp);
            this.fillOptionArr.AddRange(tmp);
            changeFillContent(this.fillOptionArr.Count);
            getFillContent();
        }

        private void addFillOption(List<fillOption> options, bool isInit = false)
        {
            if (this.fillOptionArr == null)
            {
                this.fillOptionArr = new List<fillOption>();
            }
            var sX = 20;
            var sY = 30;
            var labelW = 50;
            var contentW = 150;
            var diff = 200;
            var inDiff = 50;
            var count = isInit ? 0 : this.fillOptionArr.Count;
            sY += diff * count;
            for (int i = 0; i < options.Count; i++)
            {
                var labelTmp = new Label();
                labelTmp.Text = $"[填空{i + 1 + count}]";
                labelTmp.ForeColor = System.Drawing.Color.FromArgb(100, 99, 158, 244);
                labelTmp.Visible = true;
                labelTmp.Location = new Point(0, 0);
                var scoreText = new Label();
                scoreText.Text = "分值";
                scoreText.Visible = true;
                scoreText.Width = labelW;
                scoreText.Location = new Point(0, inDiff);
                var scoreInput = new NumericUpDown();
                scoreInput.Visible = true;
                scoreInput.Location = new Point(labelW, inDiff);
                scoreInput.Width = contentW;
                scoreInput.ValueChanged += ScoreInput_ValueChanged;
                var ansText = new Label();
                ansText.Text = "答案";
                ansText.Visible = true;
                ansText.Location = new Point(0, inDiff * 2);
                ansText.Width = labelW;
                var ansInput = new TextBox();
                ansInput.Visible = true;
                ansInput.Location = new Point(labelW, inDiff * 2);
                ansInput.Width = contentW;
                ansInput.TextChanged += AnsInput_TextChanged;
                if (isInit)
                {
                    var curValue = options[i];
                    ansInput.Text = curValue.answer;
                    scoreInput.Value = (decimal)curValue.score;
                }
                var panelTmp = new Panel();
                panelTmp.Left = sX;
                panelTmp.Top = sY + i * diff;
                panelTmp.Visible = true;
                panelTmp.Height = 150;
                panelTmp.Controls.Add(labelTmp);
                panelTmp.Controls.Add(scoreText);
                panelTmp.Controls.Add(scoreInput);
                panelTmp.Controls.Add(ansText);
                panelTmp.Controls.Add(ansInput);
                this.fillOptionPanel.Controls.Add(panelTmp);
            }
            this.panel5.Height = 800;
            this.fillOptionPanel.Height = 700;
            //this.fillOptionPanel.AutoScroll = true;
        }

        private void ScoreInput_ValueChanged(object sender, EventArgs e)
        {
            getFillContent();
        }

        private void AnsInput_TextChanged(object sender, EventArgs e)
        {
            getFillContent();
        }

        private void getFillContent()
        {
            if (this.fillOptionPanel.Controls.Count < 1)
            {
                return;
            }
            var ansTmp = new List<fillOption>();
            float score = 0;
            foreach (var curPanel in this.fillOptionPanel.Controls)
            {
                var inPanel = curPanel as Panel;
                var inAns = new fillOption();
                foreach (var curItem in inPanel.Controls)
                {

                    if (curItem is NumericUpDown)
                    {
                        var cItem = curItem as NumericUpDown;
                        inAns.score = (float)cItem.Value;
                        score += inAns.score;
                    }
                    else if (curItem is TextBox)
                    {
                        var cItem = curItem as TextBox;
                        inAns.answer = cItem.Text;
                    }

                }
                ansTmp.Add(inAns);
            }
            int currentSlideIndex = Globals.ThisAddIn.CurSlideIdx;
            PowerPoint.Shapes curShapes = Globals.ThisAddIn.Application.ActivePresentation.Slides[currentSlideIndex].Shapes;
            foreach (PowerPoint.Shape shapeTmp in curShapes)
            {
                if (shapeTmp.Name == "kx-qInfo")
                {
                    string output = JsonConvert.SerializeObject(ansTmp);
                    shapeTmp.TextFrame.TextRange.Text = (output);
                }
                if (shapeTmp.Name == "kx-score")
                {
                    shapeTmp.TextFrame.TextRange.Text = score.ToString() + "分";
                }
            }
        }

        private void changeFillContent(int idx)
        {
            int currentSlideIndex = Globals.ThisAddIn.CurSlideIdx;
            PowerPoint.Shapes curShapes = Globals.ThisAddIn.Application.ActivePresentation.Slides[currentSlideIndex].Shapes;
            foreach (PowerPoint.Shape shapeTmp in curShapes)
            {
                if (shapeTmp.Name == "kx-question")
                {
                    //shapeTmp.TextFrame.TextRange.Font.Color = System.Drawing.Color.FromArgb(100, 99, 158, 244);
                    shapeTmp.TextFrame.TextRange.InsertAfter($"[填空{idx}]");
                }
            }
        }

        public void resetFill(List<fillOption> options)
        {
            changePannelShow(TypeSelEnum.fillQuestion);
            this.fillOptionPanel.Controls.Clear();
            if (this.fillOptionArr != null)
            {
                this.fillOptionArr.Clear();
            }
            this.addFillOption(options, true);
            this.fillOptionArr.AddRange(options);

        }
    }

    public class fillOption
    {
        public string answer;
        public float score;
    }
}
