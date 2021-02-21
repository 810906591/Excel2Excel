namespace Excel2Excel
{
    partial class Form2
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form2));
            this.labelFilePath = new System.Windows.Forms.Label();
            this.textBoxViceSheet = new System.Windows.Forms.TextBox();
            this.textBoxLordSheet = new System.Windows.Forms.TextBox();
            this.lblStatus = new System.Windows.Forms.Label();
            this.buttonConfirm = new System.Windows.Forms.Button();
            this.buttonNew = new System.Windows.Forms.Button();
            this.textBoxNew = new System.Windows.Forms.TextBox();
            this.labelNew = new System.Windows.Forms.Label();
            this.buttonVice = new System.Windows.Forms.Button();
            this.textBoxVice = new System.Windows.Forms.TextBox();
            this.labelVice = new System.Windows.Forms.Label();
            this.buttonLord = new System.Windows.Forms.Button();
            this.textBoxLord = new System.Windows.Forms.TextBox();
            this.labelLord = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // labelFilePath
            // 
            this.labelFilePath.Location = new System.Drawing.Point(21, 142);
            this.labelFilePath.Name = "labelFilePath";
            this.labelFilePath.Size = new System.Drawing.Size(253, 36);
            this.labelFilePath.TabIndex = 28;
            this.labelFilePath.Text = "导出文件路径";
            // 
            // textBoxViceSheet
            // 
            this.textBoxViceSheet.Location = new System.Drawing.Point(89, 68);
            this.textBoxViceSheet.Name = "textBoxViceSheet";
            this.textBoxViceSheet.Size = new System.Drawing.Size(76, 21);
            this.textBoxViceSheet.TabIndex = 27;
            this.textBoxViceSheet.Text = "Sheet1";
            // 
            // textBoxLordSheet
            // 
            this.textBoxLordSheet.Location = new System.Drawing.Point(89, 26);
            this.textBoxLordSheet.Name = "textBoxLordSheet";
            this.textBoxLordSheet.Size = new System.Drawing.Size(76, 21);
            this.textBoxLordSheet.TabIndex = 26;
            this.textBoxLordSheet.Text = "新产品明细表";
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(21, 188);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(0, 12);
            this.lblStatus.TabIndex = 25;
            // 
            // buttonConfirm
            // 
            this.buttonConfirm.Location = new System.Drawing.Point(280, 159);
            this.buttonConfirm.Name = "buttonConfirm";
            this.buttonConfirm.Size = new System.Drawing.Size(127, 38);
            this.buttonConfirm.TabIndex = 24;
            this.buttonConfirm.Text = "确   定";
            this.buttonConfirm.UseVisualStyleBackColor = true;
            this.buttonConfirm.Click += new System.EventHandler(this.buttonConfirm_Click);
            // 
            // buttonNew
            // 
            this.buttonNew.Location = new System.Drawing.Point(373, 108);
            this.buttonNew.Name = "buttonNew";
            this.buttonNew.Size = new System.Drawing.Size(34, 23);
            this.buttonNew.TabIndex = 23;
            this.buttonNew.Text = "...";
            this.buttonNew.UseVisualStyleBackColor = true;
            this.buttonNew.Click += new System.EventHandler(this.buttonNew_Click);
            // 
            // textBoxNew
            // 
            this.textBoxNew.Location = new System.Drawing.Point(88, 109);
            this.textBoxNew.Name = "textBoxNew";
            this.textBoxNew.ReadOnly = true;
            this.textBoxNew.Size = new System.Drawing.Size(279, 21);
            this.textBoxNew.TabIndex = 22;
            // 
            // labelNew
            // 
            this.labelNew.AutoSize = true;
            this.labelNew.Location = new System.Drawing.Point(29, 113);
            this.labelNew.Name = "labelNew";
            this.labelNew.Size = new System.Drawing.Size(53, 12);
            this.labelNew.TabIndex = 21;
            this.labelNew.Text = "导出目录";
            // 
            // buttonVice
            // 
            this.buttonVice.Location = new System.Drawing.Point(373, 67);
            this.buttonVice.Name = "buttonVice";
            this.buttonVice.Size = new System.Drawing.Size(34, 23);
            this.buttonVice.TabIndex = 20;
            this.buttonVice.Text = "...";
            this.buttonVice.UseVisualStyleBackColor = true;
            this.buttonVice.Click += new System.EventHandler(this.buttonVice_Click);
            // 
            // textBoxVice
            // 
            this.textBoxVice.Location = new System.Drawing.Point(171, 68);
            this.textBoxVice.Name = "textBoxVice";
            this.textBoxVice.ReadOnly = true;
            this.textBoxVice.Size = new System.Drawing.Size(196, 21);
            this.textBoxVice.TabIndex = 19;
            // 
            // labelVice
            // 
            this.labelVice.AutoSize = true;
            this.labelVice.Location = new System.Drawing.Point(29, 72);
            this.labelVice.Name = "labelVice";
            this.labelVice.Size = new System.Drawing.Size(53, 12);
            this.labelVice.TabIndex = 18;
            this.labelVice.Text = "副-Excel";
            // 
            // buttonLord
            // 
            this.buttonLord.Location = new System.Drawing.Point(373, 25);
            this.buttonLord.Name = "buttonLord";
            this.buttonLord.Size = new System.Drawing.Size(34, 23);
            this.buttonLord.TabIndex = 17;
            this.buttonLord.Text = "...";
            this.buttonLord.UseVisualStyleBackColor = true;
            this.buttonLord.Click += new System.EventHandler(this.buttonLord_Click);
            // 
            // textBoxLord
            // 
            this.textBoxLord.Location = new System.Drawing.Point(171, 26);
            this.textBoxLord.Name = "textBoxLord";
            this.textBoxLord.ReadOnly = true;
            this.textBoxLord.Size = new System.Drawing.Size(196, 21);
            this.textBoxLord.TabIndex = 16;
            // 
            // labelLord
            // 
            this.labelLord.AutoSize = true;
            this.labelLord.Location = new System.Drawing.Point(29, 29);
            this.labelLord.Name = "labelLord";
            this.labelLord.Size = new System.Drawing.Size(53, 12);
            this.labelLord.TabIndex = 15;
            this.labelLord.Text = "主-Excel";
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(433, 210);
            this.Controls.Add(this.labelFilePath);
            this.Controls.Add(this.textBoxViceSheet);
            this.Controls.Add(this.textBoxLordSheet);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.buttonConfirm);
            this.Controls.Add(this.buttonNew);
            this.Controls.Add(this.textBoxNew);
            this.Controls.Add(this.labelNew);
            this.Controls.Add(this.buttonVice);
            this.Controls.Add(this.textBoxVice);
            this.Controls.Add(this.labelVice);
            this.Controls.Add(this.buttonLord);
            this.Controls.Add(this.textBoxLord);
            this.Controls.Add(this.labelLord);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximumSize = new System.Drawing.Size(449, 249);
            this.MinimumSize = new System.Drawing.Size(449, 249);
            this.Name = "Form2";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Excel2Excel";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelFilePath;
        private System.Windows.Forms.TextBox textBoxViceSheet;
        private System.Windows.Forms.TextBox textBoxLordSheet;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Button buttonConfirm;
        private System.Windows.Forms.Button buttonNew;
        private System.Windows.Forms.TextBox textBoxNew;
        private System.Windows.Forms.Label labelNew;
        private System.Windows.Forms.Button buttonVice;
        private System.Windows.Forms.TextBox textBoxVice;
        private System.Windows.Forms.Label labelVice;
        private System.Windows.Forms.Button buttonLord;
        private System.Windows.Forms.TextBox textBoxLord;
        private System.Windows.Forms.Label labelLord;
    }
}