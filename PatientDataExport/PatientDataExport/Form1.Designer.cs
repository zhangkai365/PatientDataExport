namespace PatientDataExport
{
    partial class Form1
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

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btn_beginProgress = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.progressNum = new System.Windows.Forms.Label();
            this.totalNum = new System.Windows.Forms.Label();
            this.iffinished = new System.Windows.Forms.Label();
            this.btn_selectSavePath = new System.Windows.Forms.Button();
            this.txtbox_FilePath = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(237, 54);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(200, 21);
            this.textBox1.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(92, 54);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(124, 25);
            this.label1.TabIndex = 1;
            this.label1.Text = "待连接数据库IP：";
            // 
            // btn_beginProgress
            // 
            this.btn_beginProgress.Location = new System.Drawing.Point(315, 304);
            this.btn_beginProgress.Name = "btn_beginProgress";
            this.btn_beginProgress.Size = new System.Drawing.Size(122, 32);
            this.btn_beginProgress.TabIndex = 2;
            this.btn_beginProgress.Text = "开始导出数据";
            this.btn_beginProgress.UseVisualStyleBackColor = true;
            this.btn_beginProgress.Click += new System.EventHandler(this.button1_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(92, 126);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 12);
            this.label2.TabIndex = 3;
            this.label2.Text = "进度：";
            // 
            // progressNum
            // 
            this.progressNum.AutoSize = true;
            this.progressNum.Location = new System.Drawing.Point(235, 126);
            this.progressNum.Name = "progressNum";
            this.progressNum.Size = new System.Drawing.Size(11, 12);
            this.progressNum.TabIndex = 4;
            this.progressNum.Text = "0";
            // 
            // totalNum
            // 
            this.totalNum.AutoSize = true;
            this.totalNum.Location = new System.Drawing.Point(310, 126);
            this.totalNum.Name = "totalNum";
            this.totalNum.Size = new System.Drawing.Size(17, 12);
            this.totalNum.TabIndex = 5;
            this.totalNum.Text = "/0";
            // 
            // iffinished
            // 
            this.iffinished.AutoSize = true;
            this.iffinished.Location = new System.Drawing.Point(218, 186);
            this.iffinished.Name = "iffinished";
            this.iffinished.Size = new System.Drawing.Size(41, 12);
            this.iffinished.TabIndex = 6;
            this.iffinished.Text = "未完成";
            // 
            // btn_selectSavePath
            // 
            this.btn_selectSavePath.Location = new System.Drawing.Point(102, 305);
            this.btn_selectSavePath.Name = "btn_selectSavePath";
            this.btn_selectSavePath.Size = new System.Drawing.Size(144, 31);
            this.btn_selectSavePath.TabIndex = 7;
            this.btn_selectSavePath.Text = "选择Excel文件存储路径";
            this.btn_selectSavePath.UseVisualStyleBackColor = true;
            this.btn_selectSavePath.Click += new System.EventHandler(this.btn_selectSavePath_Click);
            // 
            // txtbox_FilePath
            // 
            this.txtbox_FilePath.Location = new System.Drawing.Point(177, 240);
            this.txtbox_FilePath.Name = "txtbox_FilePath";
            this.txtbox_FilePath.Size = new System.Drawing.Size(310, 21);
            this.txtbox_FilePath.TabIndex = 8;
            this.txtbox_FilePath.Text = "C:\\Users\\Win7x64_20140606\\Documents\\zk2.xls";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(44, 243);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(107, 12);
            this.label3.TabIndex = 9;
            this.label3.Text = "Excel文件存储路径";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(549, 363);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtbox_FilePath);
            this.Controls.Add(this.btn_selectSavePath);
            this.Controls.Add(this.iffinished);
            this.Controls.Add(this.totalNum);
            this.Controls.Add(this.progressNum);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btn_beginProgress);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox1);
            this.Name = "Form1";
            this.Text = "体检数据导出为Excel格式";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btn_beginProgress;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label progressNum;
        private System.Windows.Forms.Label totalNum;
        private System.Windows.Forms.Label iffinished;
        private System.Windows.Forms.Button btn_selectSavePath;
        private System.Windows.Forms.TextBox txtbox_FilePath;
        private System.Windows.Forms.Label label3;
    }
}

