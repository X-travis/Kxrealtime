using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace kxrealtime.components
{
    public partial class cancelForm : Form
    {
        public cancelForm()
        {
            InitializeComponent();
        }

        public void onClosing()
        {
            this.Show();
            
            this.pictureBox1.Visible = true;
            this.undone.Visible = true;
            this.cancel1.Visible = true;
            this.uncancel1.Visible = true;
        }

        private void UserControl1_Load(object sender, EventArgs e)
        {

            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
        }

        private void uncancel1_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {

        }

        private void undone_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
}
