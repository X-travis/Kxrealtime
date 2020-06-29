using Microsoft.Office.Interop.PowerPoint;

namespace kxrealtime
{
    partial class utilDialog
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.sendBtn = new System.Windows.Forms.Button();
            this.checkAns = new System.Windows.Forms.Button();
            this.examUtils = new System.Windows.Forms.Panel();
            this.stopBtn = new System.Windows.Forms.Button();
            this.timeLeft = new System.Windows.Forms.Label();
            this.delayBtn = new System.Windows.Forms.Button();
            this.utilsPanel = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.utilsBtn = new System.Windows.Forms.Button();
            this.button8 = new System.Windows.Forms.Button();
            this.button6 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.examUtils.SuspendLayout();
            this.utilsPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // sendBtn
            // 
            this.sendBtn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.sendBtn.BackColor = System.Drawing.SystemColors.MenuHighlight;
            this.sendBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.sendBtn.FlatAppearance.BorderSize = 0;
            this.sendBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.sendBtn.Font = new System.Drawing.Font("宋体", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.sendBtn.ForeColor = System.Drawing.Color.White;
            this.sendBtn.Location = new System.Drawing.Point(552, 498);
            this.sendBtn.Margin = new System.Windows.Forms.Padding(4);
            this.sendBtn.Name = "sendBtn";
            this.sendBtn.Size = new System.Drawing.Size(200, 62);
            this.sendBtn.TabIndex = 1;
            this.sendBtn.Text = "发送题目";
            this.sendBtn.UseVisualStyleBackColor = false;
            this.sendBtn.Click += new System.EventHandler(this.sendBtn_Click);
            // 
            // checkAns
            // 
            this.checkAns.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.checkAns.BackColor = System.Drawing.SystemColors.MenuHighlight;
            this.checkAns.Cursor = System.Windows.Forms.Cursors.Hand;
            this.checkAns.FlatAppearance.BorderSize = 0;
            this.checkAns.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkAns.Font = new System.Drawing.Font("宋体", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.checkAns.ForeColor = System.Drawing.Color.White;
            this.checkAns.Location = new System.Drawing.Point(796, 498);
            this.checkAns.Margin = new System.Windows.Forms.Padding(4);
            this.checkAns.Name = "checkAns";
            this.checkAns.Size = new System.Drawing.Size(200, 62);
            this.checkAns.TabIndex = 2;
            this.checkAns.Text = "作答情况";
            this.checkAns.UseVisualStyleBackColor = false;
            this.checkAns.Click += new System.EventHandler(this.checkAns_Click);
            // 
            // examUtils
            // 
            this.examUtils.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.examUtils.Controls.Add(this.stopBtn);
            this.examUtils.Controls.Add(this.timeLeft);
            this.examUtils.Controls.Add(this.delayBtn);
            this.examUtils.Location = new System.Drawing.Point(245, 1);
            this.examUtils.Margin = new System.Windows.Forms.Padding(4);
            this.examUtils.Name = "examUtils";
            this.examUtils.Size = new System.Drawing.Size(676, 75);
            this.examUtils.TabIndex = 4;
            // 
            // stopBtn
            // 
            this.stopBtn.BackColor = System.Drawing.SystemColors.MenuHighlight;
            this.stopBtn.FlatAppearance.BorderSize = 0;
            this.stopBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.stopBtn.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.stopBtn.ForeColor = System.Drawing.Color.White;
            this.stopBtn.Location = new System.Drawing.Point(488, 2);
            this.stopBtn.Margin = new System.Windows.Forms.Padding(4);
            this.stopBtn.Name = "stopBtn";
            this.stopBtn.Size = new System.Drawing.Size(160, 68);
            this.stopBtn.TabIndex = 2;
            this.stopBtn.Text = "结束练习";
            this.stopBtn.UseVisualStyleBackColor = false;
            this.stopBtn.Click += new System.EventHandler(this.stopBtn_Click);
            // 
            // timeLeft
            // 
            this.timeLeft.AutoSize = true;
            this.timeLeft.Font = new System.Drawing.Font("宋体", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.timeLeft.Location = new System.Drawing.Point(4, 14);
            this.timeLeft.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.timeLeft.Name = "timeLeft";
            this.timeLeft.Size = new System.Drawing.Size(0, 37);
            this.timeLeft.TabIndex = 0;
            // 
            // delayBtn
            // 
            this.delayBtn.BackColor = System.Drawing.SystemColors.MenuHighlight;
            this.delayBtn.FlatAppearance.BorderSize = 0;
            this.delayBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.delayBtn.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.delayBtn.ForeColor = System.Drawing.Color.White;
            this.delayBtn.Location = new System.Drawing.Point(307, 2);
            this.delayBtn.Margin = new System.Windows.Forms.Padding(4);
            this.delayBtn.Name = "delayBtn";
            this.delayBtn.Size = new System.Drawing.Size(156, 68);
            this.delayBtn.TabIndex = 1;
            this.delayBtn.Text = "延时";
            this.delayBtn.UseVisualStyleBackColor = false;
            this.delayBtn.Click += new System.EventHandler(this.delayBtn_Click);
            // 
            // utilsPanel
            // 
            this.utilsPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.utilsPanel.BackColor = System.Drawing.Color.White;
            this.utilsPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.utilsPanel.Controls.Add(this.button8);
            this.utilsPanel.Controls.Add(this.label2);
            this.utilsPanel.Controls.Add(this.button6);
            this.utilsPanel.Controls.Add(this.button5);
            this.utilsPanel.Controls.Add(this.button4);
            this.utilsPanel.Controls.Add(this.button3);
            this.utilsPanel.Controls.Add(this.button2);
            this.utilsPanel.Controls.Add(this.label1);
            this.utilsPanel.Location = new System.Drawing.Point(585, 139);
            this.utilsPanel.Margin = new System.Windows.Forms.Padding(4);
            this.utilsPanel.Name = "utilsPanel";
            this.utilsPanel.Size = new System.Drawing.Size(389, 274);
            this.utilsPanel.TabIndex = 5;
            this.utilsPanel.Visible = false;
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.ForeColor = System.Drawing.Color.Black;
            this.label2.Location = new System.Drawing.Point(336, 14);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(34, 24);
            this.label2.TabIndex = 7;
            this.label2.Text = "×";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.Color.Black;
            this.label1.Location = new System.Drawing.Point(17, 21);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(37, 15);
            this.label1.TabIndex = 0;
            this.label1.Text = "工具";
            // 
            // utilsBtn
            // 
            this.utilsBtn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.utilsBtn.BackColor = System.Drawing.SystemColors.Control;
            this.utilsBtn.FlatAppearance.BorderSize = 0;
            this.utilsBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.utilsBtn.Image = global::kxrealtime.Properties.Resources.设置备份3;
            this.utilsBtn.Location = new System.Drawing.Point(982, 238);
            this.utilsBtn.Margin = new System.Windows.Forms.Padding(4);
            this.utilsBtn.Name = "utilsBtn";
            this.utilsBtn.Size = new System.Drawing.Size(84, 73);
            this.utilsBtn.TabIndex = 6;
            this.utilsBtn.UseVisualStyleBackColor = false;
            this.utilsBtn.Click += new System.EventHandler(this.utilsBtn_Click);
            // 
            // button8
            // 
            this.button8.FlatAppearance.BorderSize = 0;
            this.button8.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button8.ForeColor = System.Drawing.Color.Black;
            this.button8.Image = global::kxrealtime.Properties.Resources.course_qrcode1;
            this.button8.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.button8.Location = new System.Drawing.Point(261, 156);
            this.button8.Margin = new System.Windows.Forms.Padding(4);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(100, 96);
            this.button8.TabIndex = 8;
            this.button8.Text = "弹幕";
            this.button8.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Visible = false;
            this.button8.Click += new System.EventHandler(this.button8_Click);
            // 
            // button6
            // 
            this.button6.FlatAppearance.BorderSize = 0;
            this.button6.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button6.ForeColor = System.Drawing.Color.Black;
            this.button6.Image = global::kxrealtime.Properties.Resources.contribute1;
            this.button6.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.button6.Location = new System.Drawing.Point(147, 156);
            this.button6.Margin = new System.Windows.Forms.Padding(4);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(89, 96);
            this.button6.TabIndex = 5;
            this.button6.Text = "查看投稿";
            this.button6.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // button5
            // 
            this.button5.FlatAppearance.BorderSize = 0;
            this.button5.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button5.ForeColor = System.Drawing.Color.Black;
            this.button5.Image = global::kxrealtime.Properties.Resources.unstandand1;
            this.button5.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.button5.Location = new System.Drawing.Point(20, 159);
            this.button5.Margin = new System.Windows.Forms.Padding(4);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(89, 96);
            this.button5.TabIndex = 4;
            this.button5.Text = "学生不懂";
            this.button5.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // button4
            // 
            this.button4.FlatAppearance.BorderSize = 0;
            this.button4.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button4.ForeColor = System.Drawing.Color.Black;
            this.button4.Image = global::kxrealtime.Properties.Resources.sign_qrcode_big;
            this.button4.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.button4.Location = new System.Drawing.Point(261, 52);
            this.button4.Margin = new System.Windows.Forms.Padding(4);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(100, 96);
            this.button4.TabIndex = 3;
            this.button4.Text = "签到二维码";
            this.button4.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button3
            // 
            this.button3.FlatAppearance.BorderSize = 0;
            this.button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button3.ForeColor = System.Drawing.Color.Black;
            this.button3.Image = global::kxrealtime.Properties.Resources.check_stu_big;
            this.button3.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.button3.Location = new System.Drawing.Point(147, 52);
            this.button3.Margin = new System.Windows.Forms.Padding(4);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(89, 96);
            this.button3.TabIndex = 2;
            this.button3.Text = "点名";
            this.button3.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button2
            // 
            this.button2.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.button2.FlatAppearance.BorderSize = 0;
            this.button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button2.ForeColor = System.Drawing.Color.Black;
            this.button2.Image = global::kxrealtime.Properties.Resources.divide_group1;
            this.button2.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.button2.Location = new System.Drawing.Point(20, 55);
            this.button2.Margin = new System.Windows.Forms.Padding(4);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(89, 92);
            this.button2.TabIndex = 1;
            this.button2.Text = "分组";
            this.button2.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.BackColor = System.Drawing.SystemColors.Control;
            this.button1.BackgroundImage = global::kxrealtime.Properties.Resources.close_btn;
            this.button1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button1.FlatAppearance.BorderColor = System.Drawing.SystemColors.Control;
            this.button1.FlatAppearance.BorderSize = 0;
            this.button1.FlatAppearance.MouseDownBackColor = System.Drawing.SystemColors.Control;
            this.button1.FlatAppearance.MouseOverBackColor = System.Drawing.SystemColors.Control;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.ForeColor = System.Drawing.Color.White;
            this.button1.Location = new System.Drawing.Point(961, 12);
            this.button1.Margin = new System.Windows.Forms.Padding(4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(129, 50);
            this.button1.TabIndex = 3;
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            this.button1.KeyDown += Button1_KeyDown;
            this.button1.KeyUp += Button1_KeyUp;
            //新增可以鼠标上下键关闭PPT顺利翻页

            // 
            // utilDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(1107, 681);
            this.Controls.Add(this.utilsBtn);
            this.Controls.Add(this.utilsPanel);
            this.Controls.Add(this.examUtils);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.checkAns);
            this.Controls.Add(this.sendBtn);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "utilDialog";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "utilDialog";
            this.TopMost = true;
            this.TransparencyKey = System.Drawing.SystemColors.Control;
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.utilDialog_Load);
            this.examUtils.ResumeLayout(false);
            this.examUtils.PerformLayout();
            this.utilsPanel.ResumeLayout(false);
            this.utilsPanel.PerformLayout();
            this.ResumeLayout(false);

        }

        private void Button1_KeyUp(object sender, System.Windows.Forms.KeyEventArgs e)
        {
           this.Close() ;
            
        }

        private void Button1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            this.Close(); ;
        }

        #endregion
        private System.Windows.Forms.Button sendBtn;
        private System.Windows.Forms.Button checkAns;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Panel examUtils;
        private System.Windows.Forms.Button stopBtn;
        private System.Windows.Forms.Button delayBtn;
        private System.Windows.Forms.Label timeLeft;
        private System.Windows.Forms.Panel utilsPanel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button8;
        protected System.Windows.Forms.Button utilsBtn;
    }
}