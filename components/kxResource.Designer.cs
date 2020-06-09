namespace kxrealtime
{
    partial class kxResource
    {
        /// <summary> 
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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
            this.resourceWebBrowser = new System.Windows.Forms.WebBrowser();
            this.fileLoadingPic = new System.Windows.Forms.PictureBox();
            this.progresslabel = new System.Windows.Forms.Label();
            this.fileLoading = new System.Windows.Forms.Panel();
            this.savePathLabel = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.fileLoadingPic)).BeginInit();
            this.fileLoading.SuspendLayout();
            this.SuspendLayout();
            // 
            // resourceWebBrowser
            // 
            this.resourceWebBrowser.Dock = System.Windows.Forms.DockStyle.Fill;
            this.resourceWebBrowser.Location = new System.Drawing.Point(0, 0);
            this.resourceWebBrowser.Margin = new System.Windows.Forms.Padding(4);
            this.resourceWebBrowser.MinimumSize = new System.Drawing.Size(27, 25);
            this.resourceWebBrowser.Name = "resourceWebBrowser";
            this.resourceWebBrowser.Size = new System.Drawing.Size(443, 652);
            this.resourceWebBrowser.TabIndex = 0;
            // 
            // fileLoadingPic
            // 
            this.fileLoadingPic.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.fileLoadingPic.BackColor = System.Drawing.Color.White;
            this.fileLoadingPic.Image = global::kxrealtime.Properties.Resources.page_loading;
            this.fileLoadingPic.Location = new System.Drawing.Point(141, 195);
            this.fileLoadingPic.Margin = new System.Windows.Forms.Padding(4);
            this.fileLoadingPic.Name = "fileLoadingPic";
            this.fileLoadingPic.Size = new System.Drawing.Size(172, 181);
            this.fileLoadingPic.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.fileLoadingPic.TabIndex = 1;
            this.fileLoadingPic.TabStop = false;
            // 
            // progresslabel
            // 
            this.progresslabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progresslabel.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.progresslabel.Location = new System.Drawing.Point(77, 486);
            this.progresslabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.progresslabel.Name = "progresslabel";
            this.progresslabel.Size = new System.Drawing.Size(299, 52);
            this.progresslabel.TabIndex = 2;
            this.progresslabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.progresslabel.Click += new System.EventHandler(this.progresslabel_Click);
            // 
            // fileLoading
            // 
            this.fileLoading.BackColor = System.Drawing.Color.White;
            this.fileLoading.Controls.Add(this.savePathLabel);
            this.fileLoading.Controls.Add(this.fileLoadingPic);
            this.fileLoading.Controls.Add(this.progresslabel);
            this.fileLoading.Dock = System.Windows.Forms.DockStyle.Fill;
            this.fileLoading.Location = new System.Drawing.Point(0, 0);
            this.fileLoading.Margin = new System.Windows.Forms.Padding(4);
            this.fileLoading.Name = "fileLoading";
            this.fileLoading.Size = new System.Drawing.Size(443, 652);
            this.fileLoading.TabIndex = 3;
            this.fileLoading.Visible = false;
            // 
            // savePathLabel
            // 
            this.savePathLabel.AutoSize = true;
            this.savePathLabel.Font = new System.Drawing.Font("宋体", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.savePathLabel.Location = new System.Drawing.Point(25, 570);
            this.savePathLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.savePathLabel.Name = "savePathLabel";
            this.savePathLabel.Size = new System.Drawing.Size(403, 30);
            this.savePathLabel.TabIndex = 3;
            this.savePathLabel.Text = "正在缓存到酷课堂文件目录中";
            this.savePathLabel.Click += new System.EventHandler(this.savePathLabel_Click);
            // 
            // kxResource
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.fileLoading);
            this.Controls.Add(this.resourceWebBrowser);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "kxResource";
            this.Size = new System.Drawing.Size(443, 652);
            ((System.ComponentModel.ISupportInitialize)(this.fileLoadingPic)).EndInit();
            this.fileLoading.ResumeLayout(false);
            this.fileLoading.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.WebBrowser resourceWebBrowser;
        private System.Windows.Forms.PictureBox fileLoadingPic;
        private System.Windows.Forms.Label progresslabel;
        private System.Windows.Forms.Panel fileLoading;
        private System.Windows.Forms.Label savePathLabel;
    }
}
