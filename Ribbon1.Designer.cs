namespace kxrealtime
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();   
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button5 = this.Factory.CreateRibbonButton();
            this.menu1 = this.Factory.CreateRibbonMenu();
            this.button7 = this.Factory.CreateRibbonButton();
            this.button8 = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.button6 = this.Factory.CreateRibbonButton();
            this.box1 = this.Factory.CreateRibbonBox();
            this.button12 = this.Factory.CreateRibbonButton();
            this.button10 = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.button11 = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.resourceGroup = this.Factory.CreateRibbonGroup();
            this.resourceBtn = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.box1.SuspendLayout();
            this.group1.SuspendLayout();
            this.resourceGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.resourceGroup);
            this.tab1.Label = "酷课堂";
            this.tab1.Name = "tab1";
            // 
            // group2
            // 
            this.group2.Items.Add(this.button5);
            this.group2.Items.Add(this.menu1);
            this.group2.Label = "教师登录";
            this.group2.Name = "group2";
            // 
            // button5
            // 
            this.button5.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button5.Image = global::kxrealtime.Properties.Resources.wechat;
            this.button5.Label = "微信扫一扫";
            this.button5.Name = "button5";
            this.button5.ShowImage = true;
            this.button5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button5_Click);
            // 
            // menu1
            // 
            this.menu1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menu1.Items.Add(this.button7);
            this.menu1.Items.Add(this.button8);
            this.menu1.Label = "menu1";
            this.menu1.Name = "menu1";
            this.menu1.ShowImage = true;
            this.menu1.Visible = false;
            // 
            // button7
            // 
            this.button7.Label = "重新登录";
            this.button7.Name = "button7";
            this.button7.ShowImage = true;
            this.button7.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button7_Click);
            // 
            // button8
            // 
            this.button8.Label = "退出登录";
            this.button8.Name = "button8";
            this.button8.ShowImage = true;
            this.button8.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button8_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.button6);
            this.group3.Items.Add(this.box1);
            this.group3.Label = "课堂教学";
            this.group3.Name = "group3";
            // 
            // button6
            // 
            this.button6.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button6.Image = global::kxrealtime.Properties.Resources.start_new;
            this.button6.Label = " 开启授课";
            this.button6.Name = "button6";
            this.button6.ShowImage = true;
            this.button6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button6_Click);
            // 
            // box1
            // 
            this.box1.Items.Add(this.button12);
            this.box1.Items.Add(this.button10);
            this.box1.Name = "box1";
            this.box1.Visible = false;
            // 
            // button12
            // 
            this.button12.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button12.Image = global::kxrealtime.Properties.Resources.start_new;
            this.button12.Label = "从当前页放映";
            this.button12.Name = "button12";
            this.button12.ShowImage = true;
            this.button12.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button12_Click);
            // 
            // button10
            // 
            this.button10.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button10.Image = global::kxrealtime.Properties.Resources.stop1;
            this.button10.Label = "结束本次授课";
            this.button10.Name = "button10";
            this.button10.ShowImage = true;
            this.button10.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button10_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1);
            this.group1.Items.Add(this.button2);
            this.group1.Items.Add(this.button3);
            this.group1.Items.Add(this.button11);
            this.group1.Items.Add(this.button4);
            this.group1.Label = "插入题目";
            this.group1.Name = "group1";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = global::kxrealtime.Properties.Resources.single_new1;
            this.button1.Label = "                单选题";
            this.button1.Name = "button1";
            this.button1.ScreenTip = "插入单项选择题";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Description = "多选题";
            this.button2.Image = global::kxrealtime.Properties.Resources.multi_new1;
            this.button2.Label = "                     多选题";
            this.button2.Name = "button2";
            this.button2.ScreenTip = "插入多项选择题";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button3.Image = global::kxrealtime.Properties.Resources.fill_new1;
            this.button3.Label = "                填空题 ";
            this.button3.Name = "button3";
            this.button3.ScreenTip = "插入填空题";
            this.button3.ShowImage = true;
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // button11
            // 
            this.button11.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button11.Image = global::kxrealtime.Properties.Resources.vote_new1;
            this.button11.Label = "             投票";
            this.button11.Name = "button11";
            this.button11.ShowImage = true;
            this.button11.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button11_Click);
            // 
            // button4
            // 
            this.button4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button4.Image = global::kxrealtime.Properties.Resources.subject_new1;
            this.button4.Label = "               主观题 ";
            this.button4.Name = "button4";
            this.button4.ScreenTip = "插入主观题";
            this.button4.ShowImage = true;
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button4_Click);
            // 
            // resourceGroup
            // 
            this.resourceGroup.Items.Add(this.resourceBtn);
            this.resourceGroup.Label = "资源库";
            this.resourceGroup.Name = "resourceGroup";
            // 
            // resourceBtn
            // 
            this.resourceBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.resourceBtn.Image = global::kxrealtime.Properties.Resources.resource;
            this.resourceBtn.Label = "               酷课堂";
            this.resourceBtn.Name = "resourceBtn";
            this.resourceBtn.ShowImage = true;
            this.resourceBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.resourceBtn_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.resourceGroup.ResumeLayout(false);
            this.resourceGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button8;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button10;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button11;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button12;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup resourceGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton resourceBtn;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
