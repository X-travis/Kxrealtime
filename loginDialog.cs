using System;
using System.Drawing;
using System.Windows.Forms;

namespace kxrealtime
{
    public partial class loginDialog : Form
    {
        public static Form frmBack;
        public loginDialog()
        {
            InitializeComponent();
        }

        public Panel getContent
        {
            get
            {
                return this.panel1;
            }
        }

        public Label getClose
        {
            get
            {
                return this.label1;
            }
        }

        public Label getTitle
        {
            get
            {
                return this.label2;
            }
        }

        public PictureBox getLogo
        {
            get
            {
                return this.pictureBox1;
            }
        }

        public static void Show(Form frmTop, Color frmBackColor, double frmBackOpacity)
        {
            // 背景窗体设置
            //if (frmBack == null)
            {
                frmBack = new Form();
                frmBack.FormBorderStyle = FormBorderStyle.None;
                frmBack.StartPosition = FormStartPosition.Manual;
                frmBack.ShowIcon = false;
                frmBack.ShowInTaskbar = false;
                //frmBack.Size = frmTop.Size;
                frmBack.WindowState = FormWindowState.Maximized;
                frmBack.TopMost = false;

            }

            frmBack.Opacity = frmBackOpacity;
            frmBack.BackColor = frmBackColor;
            frmBack.Location = frmTop.Location;
            // 顶部窗体设置
            frmTop.Owner = frmBack;
            frmTop.StartPosition = FormStartPosition.CenterScreen;

            //frmTop.LocationChanged += (o, args) => { frmBack.Location = frmTop.Location; };
            // 显示窗体
            frmTop.Show();
            frmBack.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            frmBack.Close();
            this.Close();
            frmBack = null;
        }

        private void label1_Click(object sender, EventArgs e)
        {
            if (frmBack != null)
            {
                frmBack.Close();
                frmBack.Dispose();
            }
            this.Close();
            frmBack = null;
            Globals.Ribbons.Ribbon1.closeLoginConnect();
        }
    }
}
