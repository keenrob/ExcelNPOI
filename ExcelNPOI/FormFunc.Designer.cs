namespace ExcelNPOI
{
    partial class FormFunc
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormFunc));
            this.BtnImporAllDep = new System.Windows.Forms.Button();
            this.BtnMoveLastYear = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // BtnImporAllDep
            // 
            this.BtnImporAllDep.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.BtnImporAllDep.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.BtnImporAllDep.Location = new System.Drawing.Point(124, 23);
            this.BtnImporAllDep.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.BtnImporAllDep.Name = "BtnImporAllDep";
            this.BtnImporAllDep.Size = new System.Drawing.Size(102, 29);
            this.BtnImporAllDep.TabIndex = 0;
            this.BtnImporAllDep.Text = "导出所有部门";
            this.BtnImporAllDep.UseVisualStyleBackColor = true;
            this.BtnImporAllDep.Click += new System.EventHandler(this.BtnImporAllDep_Click);
            // 
            // BtnMoveLastYear
            // 
            this.BtnMoveLastYear.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.BtnMoveLastYear.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.BtnMoveLastYear.Location = new System.Drawing.Point(124, 109);
            this.BtnMoveLastYear.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.BtnMoveLastYear.Name = "BtnMoveLastYear";
            this.BtnMoveLastYear.Size = new System.Drawing.Size(102, 29);
            this.BtnMoveLastYear.TabIndex = 1;
            this.BtnMoveLastYear.Text = "归档上一年度";
            this.BtnMoveLastYear.UseVisualStyleBackColor = true;
            this.BtnMoveLastYear.Click += new System.EventHandler(this.BtnMoveLastYear_Click);
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button1.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.button1.Location = new System.Drawing.Point(124, 66);
            this.button1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(102, 29);
            this.button1.TabIndex = 2;
            this.button1.Text = "发薪日请假人员";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // FormFunc
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(370, 172);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.BtnMoveLastYear);
            this.Controls.Add(this.BtnImporAllDep);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormFunc";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "功能窗体";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button BtnImporAllDep;
        private System.Windows.Forms.Button BtnMoveLastYear;
        private System.Windows.Forms.Button button1;
    }
}