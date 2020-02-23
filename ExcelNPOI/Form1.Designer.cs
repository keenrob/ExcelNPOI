namespace ExcelNPOI
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
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.groupBoxChoose = new System.Windows.Forms.GroupBox();
            this.checkBox8 = new System.Windows.Forms.CheckBox();
            this.checkBox7 = new System.Windows.Forms.CheckBox();
            this.checkBox6 = new System.Windows.Forms.CheckBox();
            this.checkBox5 = new System.Windows.Forms.CheckBox();
            this.checkBox4 = new System.Windows.Forms.CheckBox();
            this.checkBox3 = new System.Windows.Forms.CheckBox();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.dtBegin = new System.Windows.Forms.DateTimePicker();
            this.dtEnd = new System.Windows.Forms.DateTimePicker();
            this.btnImport = new System.Windows.Forms.Button();
            this.lblVer = new System.Windows.Forms.Label();
            this.BtnFunc = new System.Windows.Forms.Button();
            this.BtnCountEmp = new System.Windows.Forms.Button();
            this.groupBoxChoose.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBoxChoose
            // 
            this.groupBoxChoose.Controls.Add(this.checkBox8);
            this.groupBoxChoose.Controls.Add(this.checkBox7);
            this.groupBoxChoose.Controls.Add(this.checkBox6);
            this.groupBoxChoose.Controls.Add(this.checkBox5);
            this.groupBoxChoose.Controls.Add(this.checkBox4);
            this.groupBoxChoose.Controls.Add(this.checkBox3);
            this.groupBoxChoose.Controls.Add(this.checkBox2);
            this.groupBoxChoose.Controls.Add(this.checkBox1);
            this.groupBoxChoose.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBoxChoose.Location = new System.Drawing.Point(9, 66);
            this.groupBoxChoose.Margin = new System.Windows.Forms.Padding(2);
            this.groupBoxChoose.Name = "groupBoxChoose";
            this.groupBoxChoose.Padding = new System.Windows.Forms.Padding(2);
            this.groupBoxChoose.Size = new System.Drawing.Size(545, 207);
            this.groupBoxChoose.TabIndex = 1;
            this.groupBoxChoose.TabStop = false;
            this.groupBoxChoose.Text = "选择部门/车间";
            // 
            // checkBox8
            // 
            this.checkBox8.AutoSize = true;
            this.checkBox8.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.checkBox8.ForeColor = System.Drawing.Color.Green;
            this.checkBox8.Location = new System.Drawing.Point(326, 163);
            this.checkBox8.Margin = new System.Windows.Forms.Padding(2);
            this.checkBox8.Name = "checkBox8";
            this.checkBox8.Size = new System.Drawing.Size(75, 21);
            this.checkBox8.TabIndex = 7;
            this.checkBox8.Text = "行政部门";
            this.checkBox8.UseVisualStyleBackColor = true;
            this.checkBox8.CheckedChanged += new System.EventHandler(this.CheckBox8_CheckedChanged);
            // 
            // checkBox7
            // 
            this.checkBox7.AutoSize = true;
            this.checkBox7.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.checkBox7.ForeColor = System.Drawing.Color.Green;
            this.checkBox7.Location = new System.Drawing.Point(326, 127);
            this.checkBox7.Margin = new System.Windows.Forms.Padding(2);
            this.checkBox7.Name = "checkBox7";
            this.checkBox7.Size = new System.Drawing.Size(75, 21);
            this.checkBox7.TabIndex = 6;
            this.checkBox7.Text = "电工车间";
            this.checkBox7.UseVisualStyleBackColor = true;
            // 
            // checkBox6
            // 
            this.checkBox6.AutoSize = true;
            this.checkBox6.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.checkBox6.ForeColor = System.Drawing.Color.Green;
            this.checkBox6.Location = new System.Drawing.Point(326, 90);
            this.checkBox6.Margin = new System.Windows.Forms.Padding(2);
            this.checkBox6.Name = "checkBox6";
            this.checkBox6.Size = new System.Drawing.Size(75, 21);
            this.checkBox6.TabIndex = 5;
            this.checkBox6.Text = "动力车间";
            this.checkBox6.UseVisualStyleBackColor = true;
            // 
            // checkBox5
            // 
            this.checkBox5.AutoSize = true;
            this.checkBox5.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.checkBox5.ForeColor = System.Drawing.Color.Green;
            this.checkBox5.Location = new System.Drawing.Point(326, 58);
            this.checkBox5.Margin = new System.Windows.Forms.Padding(2);
            this.checkBox5.Name = "checkBox5";
            this.checkBox5.Size = new System.Drawing.Size(87, 21);
            this.checkBox5.TabIndex = 4;
            this.checkBox5.Text = "前处理车间";
            this.checkBox5.UseVisualStyleBackColor = true;
            // 
            // checkBox4
            // 
            this.checkBox4.AutoSize = true;
            this.checkBox4.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.checkBox4.ForeColor = System.Drawing.Color.Green;
            this.checkBox4.Location = new System.Drawing.Point(118, 163);
            this.checkBox4.Margin = new System.Windows.Forms.Padding(2);
            this.checkBox4.Name = "checkBox4";
            this.checkBox4.Size = new System.Drawing.Size(75, 21);
            this.checkBox4.TabIndex = 3;
            this.checkBox4.Text = "提取车间";
            this.checkBox4.UseVisualStyleBackColor = true;
            // 
            // checkBox3
            // 
            this.checkBox3.AutoSize = true;
            this.checkBox3.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.checkBox3.ForeColor = System.Drawing.Color.Green;
            this.checkBox3.Location = new System.Drawing.Point(118, 127);
            this.checkBox3.Margin = new System.Windows.Forms.Padding(2);
            this.checkBox3.Name = "checkBox3";
            this.checkBox3.Size = new System.Drawing.Size(87, 21);
            this.checkBox3.TabIndex = 2;
            this.checkBox3.Text = "制剂三车间";
            this.checkBox3.UseVisualStyleBackColor = true;
            // 
            // checkBox2
            // 
            this.checkBox2.AutoSize = true;
            this.checkBox2.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.checkBox2.ForeColor = System.Drawing.Color.Green;
            this.checkBox2.Location = new System.Drawing.Point(118, 90);
            this.checkBox2.Margin = new System.Windows.Forms.Padding(2);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(87, 21);
            this.checkBox2.TabIndex = 1;
            this.checkBox2.Text = "制剂二车间";
            this.checkBox2.UseVisualStyleBackColor = true;
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.checkBox1.ForeColor = System.Drawing.Color.Green;
            this.checkBox1.Location = new System.Drawing.Point(118, 58);
            this.checkBox1.Margin = new System.Windows.Forms.Padding(2);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(87, 21);
            this.checkBox1.TabIndex = 0;
            this.checkBox1.Text = "制剂一车间";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.ForeColor = System.Drawing.Color.Green;
            this.label1.Location = new System.Drawing.Point(13, 20);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(57, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "开始日期";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.ForeColor = System.Drawing.Color.Green;
            this.label2.Location = new System.Drawing.Point(222, 20);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(57, 12);
            this.label2.TabIndex = 3;
            this.label2.Text = "结束日期";
            // 
            // dtBegin
            // 
            this.dtBegin.Location = new System.Drawing.Point(70, 14);
            this.dtBegin.Margin = new System.Windows.Forms.Padding(2);
            this.dtBegin.Name = "dtBegin";
            this.dtBegin.Size = new System.Drawing.Size(128, 21);
            this.dtBegin.TabIndex = 4;
            this.dtBegin.ValueChanged += new System.EventHandler(this.DtBegin_ValueChanged);
            // 
            // dtEnd
            // 
            this.dtEnd.Location = new System.Drawing.Point(280, 14);
            this.dtEnd.Margin = new System.Windows.Forms.Padding(2);
            this.dtEnd.Name = "dtEnd";
            this.dtEnd.Size = new System.Drawing.Size(128, 21);
            this.dtEnd.TabIndex = 5;
            // 
            // btnImport
            // 
            this.btnImport.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnImport.ForeColor = System.Drawing.Color.Green;
            this.btnImport.Location = new System.Drawing.Point(411, 13);
            this.btnImport.Margin = new System.Windows.Forms.Padding(2);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(58, 26);
            this.btnImport.TabIndex = 7;
            this.btnImport.Text = "导出";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.BtnImport_Click);
            // 
            // lblVer
            // 
            this.lblVer.AutoSize = true;
            this.lblVer.Location = new System.Drawing.Point(440, 276);
            this.lblVer.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblVer.Name = "lblVer";
            this.lblVer.Size = new System.Drawing.Size(53, 12);
            this.lblVer.TabIndex = 9;
            this.lblVer.Text = "version ";
            // 
            // BtnFunc
            // 
            this.BtnFunc.Font = new System.Drawing.Font("Franklin Gothic Medium", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnFunc.ForeColor = System.Drawing.Color.Green;
            this.BtnFunc.Location = new System.Drawing.Point(516, 13);
            this.BtnFunc.Margin = new System.Windows.Forms.Padding(2);
            this.BtnFunc.Name = "BtnFunc";
            this.BtnFunc.Size = new System.Drawing.Size(38, 26);
            this.BtnFunc.TabIndex = 10;
            this.BtnFunc.Text = ">>>";
            this.BtnFunc.UseVisualStyleBackColor = true;
            this.BtnFunc.Click += new System.EventHandler(this.BtnFunc_Click);
            // 
            // BtnCountEmp
            // 
            this.BtnCountEmp.Font = new System.Drawing.Font("微软雅黑", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.BtnCountEmp.ForeColor = System.Drawing.Color.Green;
            this.BtnCountEmp.Location = new System.Drawing.Point(473, 13);
            this.BtnCountEmp.Margin = new System.Windows.Forms.Padding(2);
            this.BtnCountEmp.Name = "BtnCountEmp";
            this.BtnCountEmp.Size = new System.Drawing.Size(38, 26);
            this.BtnCountEmp.TabIndex = 11;
            this.BtnCountEmp.Text = "统计";
            this.BtnCountEmp.UseVisualStyleBackColor = true;
            this.BtnCountEmp.Click += new System.EventHandler(this.BtnCountEmp_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(566, 295);
            this.Controls.Add(this.BtnCountEmp);
            this.Controls.Add(this.BtnFunc);
            this.Controls.Add(this.lblVer);
            this.Controls.Add(this.btnImport);
            this.Controls.Add(this.dtEnd);
            this.Controls.Add(this.dtBegin);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.groupBoxChoose);
            this.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2);
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "考勤导出";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBoxChoose.ResumeLayout(false);
            this.groupBoxChoose.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.GroupBox groupBoxChoose;
        private System.Windows.Forms.CheckBox checkBox8;
        private System.Windows.Forms.CheckBox checkBox7;
        private System.Windows.Forms.CheckBox checkBox6;
        private System.Windows.Forms.CheckBox checkBox5;
        private System.Windows.Forms.CheckBox checkBox4;
        private System.Windows.Forms.CheckBox checkBox3;
        private System.Windows.Forms.CheckBox checkBox2;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker dtBegin;
        private System.Windows.Forms.DateTimePicker dtEnd;
        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.Label lblVer;
        private System.Windows.Forms.Button BtnFunc;
        private System.Windows.Forms.Button BtnCountEmp;
    }
}

