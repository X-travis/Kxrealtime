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
            this.SuspendLayout();
            // 
            // resourceWebBrowser
            // 
            this.resourceWebBrowser.Dock = System.Windows.Forms.DockStyle.Fill;
            this.resourceWebBrowser.Location = new System.Drawing.Point(0, 0);
            this.resourceWebBrowser.MinimumSize = new System.Drawing.Size(20, 20);
            this.resourceWebBrowser.Name = "resourceWebBrowser";
            this.resourceWebBrowser.Size = new System.Drawing.Size(332, 522);
            this.resourceWebBrowser.TabIndex = 0;
            // 
            // kxResource
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.resourceWebBrowser);
            this.Name = "kxResource";
            this.Size = new System.Drawing.Size(332, 522);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.WebBrowser resourceWebBrowser;
    }
}
