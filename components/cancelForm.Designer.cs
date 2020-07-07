namespace kxrealtime.components
{
    partial class cancelForm
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
            this.undone = new System.Windows.Forms.Button();
            this.uncancel1 = new System.Windows.Forms.Button();
            this.cancel1 = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // undone
            // 
            this.undone.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(10)))), ((int)(((byte)(165)))), ((int)(((byte)(238)))));
            this.undone.Image = global::kxrealtime.Properties.Resources.取消;
            this.undone.Location = new System.Drawing.Point(770, 268);
            this.undone.Name = "undone";
            this.undone.Size = new System.Drawing.Size(65, 36);
            this.undone.TabIndex = 2;
            this.undone.UseVisualStyleBackColor = true;
            this.undone.Click += new System.EventHandler(this.undone_Click);
            // 
            // uncancel1
            // 
            this.uncancel1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(10)))), ((int)(((byte)(165)))), ((int)(((byte)(238)))));
            this.uncancel1.Image = global::kxrealtime.Properties.Resources.不结束;
            this.uncancel1.Location = new System.Drawing.Point(699, 268);
            this.uncancel1.Name = "uncancel1";
            this.uncancel1.Size = new System.Drawing.Size(65, 36);
            this.uncancel1.TabIndex = 1;
            this.uncancel1.UseVisualStyleBackColor = true;
            this.uncancel1.Click += new System.EventHandler(this.uncancel1_Click);
            // 
            // cancel1
            // 
            this.cancel1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(236)))), ((int)(((byte)(73)))), ((int)(((byte)(73)))));
            this.cancel1.Image = global::kxrealtime.Properties.Resources.结束;
            this.cancel1.Location = new System.Drawing.Point(628, 268);
            this.cancel1.Name = "cancel1";
            this.cancel1.Size = new System.Drawing.Size(65, 36);
            this.cancel1.TabIndex = 0;
            this.cancel1.UseVisualStyleBackColor = true;
            this.cancel1.Click += new System.EventHandler(this.button1_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::kxrealtime.Properties.Resources.结束授课_1_;
            this.pictureBox1.Location = new System.Drawing.Point(178, 35);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(704, 313);
            this.pictureBox1.TabIndex = 3;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // cancelForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.undone);
            this.Controls.Add(this.uncancel1);
            this.Controls.Add(this.cancel1);
            this.Controls.Add(this.pictureBox1);
            this.Name = "cancelForm";
            this.Size = new System.Drawing.Size(1018, 398);
            this.Load += new System.EventHandler(this.UserControl1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button cancel1;
        private System.Windows.Forms.Button uncancel1;
        private System.Windows.Forms.Button undone;
        private System.Windows.Forms.PictureBox pictureBox1;
    }
}
