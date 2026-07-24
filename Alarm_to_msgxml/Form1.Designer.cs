namespace Alarm_to_msgxml
{
    partial class Form1
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.titleLabel = new System.Windows.Forms.Label();
            this.descriptionLabel = new System.Windows.Forms.Label();
            this.step1Label = new System.Windows.Forms.Label();
            this.step2Label = new System.Windows.Forms.Label();
            this.step3Label = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.file_name = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.noteLabel = new System.Windows.Forms.Label();
            this.separator1 = new System.Windows.Forms.Panel();
            this.separator2 = new System.Windows.Forms.Panel();
            this.SuspendLayout();
            // 
            // titleLabel
            // 
            this.titleLabel.AutoSize = true;
            this.titleLabel.Font = new System.Drawing.Font("Microsoft JhengHei UI", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.titleLabel.Location = new System.Drawing.Point(31, 24);
            this.titleLabel.Name = "titleLabel";
            this.titleLabel.Size = new System.Drawing.Size(171, 30);
            this.titleLabel.TabIndex = 0;
            this.titleLabel.Text = "Alarm Builder";
            // 
            // descriptionLabel
            // 
            this.descriptionLabel.AutoSize = true;
            this.descriptionLabel.ForeColor = System.Drawing.SystemColors.GrayText;
            this.descriptionLabel.Location = new System.Drawing.Point(34, 65);
            this.descriptionLabel.Name = "descriptionLabel";
            this.descriptionLabel.Size = new System.Drawing.Size(290, 15);
            this.descriptionLabel.TabIndex = 1;
            this.descriptionLabel.Text = "選擇 PICTURE XML 與 Excel，產生 Alarm msgxml。";
            // 
            // step1Label
            // 
            this.step1Label.AutoSize = true;
            this.step1Label.Font = new System.Drawing.Font("Microsoft JhengHei UI", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.step1Label.Location = new System.Drawing.Point(34, 111);
            this.step1Label.Name = "step1Label";
            this.step1Label.Size = new System.Drawing.Size(116, 18);
            this.step1Label.TabIndex = 2;
            this.step1Label.Text = "1. PICTURE XML";
            // 
            // step2Label
            // 
            this.step2Label.AutoSize = true;
            this.step2Label.Font = new System.Drawing.Font("Microsoft JhengHei UI", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.step2Label.Location = new System.Drawing.Point(34, 172);
            this.step2Label.Name = "step2Label";
            this.step2Label.Size = new System.Drawing.Size(103, 18);
            this.step2Label.TabIndex = 4;
            this.step2Label.Text = "2. Alarm Excel";
            // 
            // step3Label
            // 
            this.step3Label.AutoSize = true;
            this.step3Label.Font = new System.Drawing.Font("Microsoft JhengHei UI", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.step3Label.Location = new System.Drawing.Point(34, 233);
            this.step3Label.Name = "step3Label";
            this.step3Label.Size = new System.Drawing.Size(80, 18);
            this.step3Label.TabIndex = 6;
            this.step3Label.Text = "3. 輸出設定";
            // 
            // button3
            // 
            this.button3.Cursor = System.Windows.Forms.Cursors.Hand;
            this.button3.Location = new System.Drawing.Point(399, 101);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(154, 38);
            this.button3.TabIndex = 3;
            this.button3.Text = "選擇 XML 檔案";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.open_xml);
            // 
            // button1
            // 
            this.button1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.button1.Location = new System.Drawing.Point(399, 162);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(154, 38);
            this.button1.TabIndex = 5;
            this.button1.Text = "選擇 Excel 檔案";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.open_file);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(34, 268);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(67, 15);
            this.label1.TabIndex = 7;
            this.label1.Text = "輸出檔名：";
            // 
            // file_name
            // 
            this.file_name.Location = new System.Drawing.Point(107, 264);
            this.file_name.Name = "file_name";
            this.file_name.Size = new System.Drawing.Size(286, 23);
            this.file_name.TabIndex = 8;
            // 
            // button2
            // 
            this.button2.Cursor = System.Windows.Forms.Cursors.Hand;
            this.button2.Font = new System.Drawing.Font("Microsoft JhengHei UI", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button2.Location = new System.Drawing.Point(399, 254);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(154, 44);
            this.button2.TabIndex = 9;
            this.button2.Text = "開始建置";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.execute);
            // 
            // noteLabel
            // 
            this.noteLabel.AutoSize = true;
            this.noteLabel.ForeColor = System.Drawing.SystemColors.GrayText;
            this.noteLabel.Location = new System.Drawing.Point(104, 295);
            this.noteLabel.Name = "noteLabel";
            this.noteLabel.Size = new System.Drawing.Size(165, 15);
            this.noteLabel.TabIndex = 10;
            this.noteLabel.Text = "不需輸入副檔名，例如：RVN";
            // 
            // separator1
            // 
            this.separator1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.separator1.Location = new System.Drawing.Point(36, 148);
            this.separator1.Name = "separator1";
            this.separator1.Size = new System.Drawing.Size(517, 1);
            this.separator1.TabIndex = 11;
            // 
            // separator2
            // 
            this.separator2.BackColor = System.Drawing.SystemColors.ControlLight;
            this.separator2.Location = new System.Drawing.Point(36, 209);
            this.separator2.Name = "separator2";
            this.separator2.Size = new System.Drawing.Size(517, 1);
            this.separator2.TabIndex = 12;
            // 
            // Form1
            // 
            this.AcceptButton = this.button2;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(590, 339);
            this.Controls.Add(this.separator2);
            this.Controls.Add(this.separator1);
            this.Controls.Add(this.noteLabel);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.file_name);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.step3Label);
            this.Controls.Add(this.step2Label);
            this.Controls.Add(this.step1Label);
            this.Controls.Add(this.descriptionLabel);
            this.Controls.Add(this.titleLabel);
            this.Font = new System.Drawing.Font("Microsoft JhengHei UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Alarm Builder";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label titleLabel;
        private System.Windows.Forms.Label descriptionLabel;
        private System.Windows.Forms.Label step1Label;
        private System.Windows.Forms.Label step2Label;
        private System.Windows.Forms.Label step3Label;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox file_name;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label noteLabel;
        private System.Windows.Forms.Panel separator1;
        private System.Windows.Forms.Panel separator2;
    }
}
