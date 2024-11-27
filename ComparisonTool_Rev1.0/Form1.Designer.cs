namespace ComparisonTool_Rev1._0
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            UploadFirstTemplateBtn = new Button();
            UploadLastTemplateBtn = new Button();
            FirstUse = new RadioButton();
            LastUse = new RadioButton();
            BlankTempText = new Label();
            MutiFilesText = new Label();
            SymbolText1 = new Label();
            UploadMultipleFileBtn = new Button();
            SummitButton = new Button();
            NonFileFirst = new Label();
            NonFileLast = new Label();
            NonFIleMuti = new Label();
            BySRD3IT = new Label();
            Fii = new Label();
            panel1 = new Panel();
            Version = new Label();
            panel1.SuspendLayout();
            SuspendLayout();
            // 
            // UploadFirstTemplateBtn
            // 
            UploadFirstTemplateBtn.Font = new Font("Verdana", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            UploadFirstTemplateBtn.Location = new Point(337, 24);
            UploadFirstTemplateBtn.Margin = new Padding(4);
            UploadFirstTemplateBtn.Name = "UploadFirstTemplateBtn";
            UploadFirstTemplateBtn.Size = new Size(163, 61);
            UploadFirstTemplateBtn.TabIndex = 0;
            UploadFirstTemplateBtn.Text = "上傳檔案";
            UploadFirstTemplateBtn.UseVisualStyleBackColor = true;
            UploadFirstTemplateBtn.Click += UploadFirstTemplate;
            // 
            // UploadLastTemplateBtn
            // 
            UploadLastTemplateBtn.Font = new Font("Verdana", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            UploadLastTemplateBtn.Location = new Point(337, 125);
            UploadLastTemplateBtn.Margin = new Padding(4);
            UploadLastTemplateBtn.Name = "UploadLastTemplateBtn";
            UploadLastTemplateBtn.Size = new Size(163, 61);
            UploadLastTemplateBtn.TabIndex = 4;
            UploadLastTemplateBtn.Text = "上傳檔案";
            UploadLastTemplateBtn.UseVisualStyleBackColor = true;
            UploadLastTemplateBtn.Click += UploadLastTemplate;
            // 
            // FirstUse
            // 
            FirstUse.AutoSize = true;
            FirstUse.Font = new Font("新細明體", 14F, FontStyle.Regular, GraphicsUnit.Point, 136);
            FirstUse.Location = new Point(34, 36);
            FirstUse.Margin = new Padding(4);
            FirstUse.Name = "FirstUse";
            FirstUse.Size = new Size(233, 32);
            FirstUse.TabIndex = 5;
            FirstUse.TabStop = true;
            FirstUse.Text = "首次上傳使用：";
            FirstUse.UseVisualStyleBackColor = true;
            FirstUse.CheckedChanged += FirstUseChecked;
            // 
            // LastUse
            // 
            LastUse.AutoSize = true;
            LastUse.Font = new Font("新細明體", 14F, FontStyle.Regular, GraphicsUnit.Point, 136);
            LastUse.Location = new Point(34, 138);
            LastUse.Margin = new Padding(4);
            LastUse.Name = "LastUse";
            LastUse.Size = new Size(233, 32);
            LastUse.TabIndex = 6;
            LastUse.TabStop = true;
            LastUse.Text = "前次上傳使用：";
            LastUse.UseVisualStyleBackColor = true;
            LastUse.CheckedChanged += LastUseChecked;
            // 
            // BlankTempText
            // 
            BlankTempText.AutoSize = true;
            BlankTempText.Font = new Font("新細明體", 10F, FontStyle.Regular, GraphicsUnit.Point, 136);
            BlankTempText.ForeColor = Color.FromArgb(192, 0, 192);
            BlankTempText.Location = new Point(62, 85);
            BlankTempText.Margin = new Padding(4, 0, 4, 0);
            BlankTempText.Name = "BlankTempText";
            BlankTempText.Size = new Size(195, 20);
            BlankTempText.TabIndex = 7;
            BlankTempText.Text = "( *請上傳空白範本！)";
            // 
            // MutiFilesText
            // 
            MutiFilesText.AutoSize = true;
            MutiFilesText.Font = new Font("新細明體", 14F, FontStyle.Regular, GraphicsUnit.Point, 136);
            MutiFilesText.Location = new Point(59, 239);
            MutiFilesText.Margin = new Padding(4, 0, 4, 0);
            MutiFilesText.Name = "MutiFilesText";
            MutiFilesText.Size = new Size(208, 28);
            MutiFilesText.TabIndex = 8;
            MutiFilesText.Text = "多筆檔案上傳：";
            // 
            // SymbolText1
            // 
            SymbolText1.AutoSize = true;
            SymbolText1.ForeColor = Color.FromArgb(192, 0, 0);
            SymbolText1.Location = new Point(31, 227);
            SymbolText1.Margin = new Padding(4, 0, 4, 0);
            SymbolText1.Name = "SymbolText1";
            SymbolText1.Size = new Size(28, 23);
            SymbolText1.TabIndex = 9;
            SymbolText1.Text = "＊";
            // 
            // UploadMultipleFileBtn
            // 
            UploadMultipleFileBtn.Font = new Font("Verdana", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            UploadMultipleFileBtn.Location = new Point(335, 227);
            UploadMultipleFileBtn.Margin = new Padding(4);
            UploadMultipleFileBtn.Name = "UploadMultipleFileBtn";
            UploadMultipleFileBtn.Size = new Size(163, 61);
            UploadMultipleFileBtn.TabIndex = 10;
            UploadMultipleFileBtn.Text = "上傳檔案";
            UploadMultipleFileBtn.UseVisualStyleBackColor = true;
            UploadMultipleFileBtn.Click += UploadMultipleFile;
            // 
            // SummitButton
            // 
            SummitButton.Font = new Font("Verdana", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            SummitButton.Location = new Point(59, 328);
            SummitButton.Margin = new Padding(4);
            SummitButton.Name = "SummitButton";
            SummitButton.Size = new Size(152, 52);
            SummitButton.TabIndex = 11;
            SummitButton.Text = "送出";
            SummitButton.UseVisualStyleBackColor = true;
            SummitButton.Click += SummitBtn;
            // 
            // NonFileFirst
            // 
            NonFileFirst.AutoSize = true;
            NonFileFirst.Font = new Font("新細明體", 10F, FontStyle.Regular, GraphicsUnit.Point, 136);
            NonFileFirst.Location = new Point(535, 43);
            NonFileFirst.Margin = new Padding(4, 0, 4, 0);
            NonFileFirst.Name = "NonFileFirst";
            NonFileFirst.Size = new Size(161, 20);
            NonFileFirst.TabIndex = 12;
            NonFileFirst.Text = "(尚未上傳檔案！)";
            // 
            // NonFileLast
            // 
            NonFileLast.AutoSize = true;
            NonFileLast.Font = new Font("新細明體", 10F, FontStyle.Regular, GraphicsUnit.Point, 136);
            NonFileLast.Location = new Point(535, 148);
            NonFileLast.Margin = new Padding(4, 0, 4, 0);
            NonFileLast.Name = "NonFileLast";
            NonFileLast.Size = new Size(161, 20);
            NonFileLast.TabIndex = 13;
            NonFileLast.Text = "(尚未上傳檔案！)";
            // 
            // NonFIleMuti
            // 
            NonFIleMuti.AutoSize = true;
            NonFIleMuti.Font = new Font("新細明體", 10F, FontStyle.Regular, GraphicsUnit.Point, 136);
            NonFIleMuti.Location = new Point(533, 246);
            NonFIleMuti.Margin = new Padding(4, 0, 4, 0);
            NonFIleMuti.Name = "NonFIleMuti";
            NonFIleMuti.Size = new Size(161, 20);
            NonFIleMuti.TabIndex = 14;
            NonFIleMuti.Text = "(尚未上傳檔案！)";
            // 
            // BySRD3IT
            // 
            BySRD3IT.AutoSize = true;
            BySRD3IT.Font = new Font("Verdana", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            BySRD3IT.Location = new Point(706, 523);
            BySRD3IT.Margin = new Padding(4, 0, 4, 0);
            BySRD3IT.Name = "BySRD3IT";
            BySRD3IT.Size = new Size(114, 22);
            BySRD3IT.TabIndex = 15;
            BySRD3IT.Text = "By SRD3 IT";
            BySRD3IT.UseMnemonic = false;
            // 
            // Fii
            // 
            Fii.AutoSize = true;
            Fii.Font = new Font("Verdana", 18F, FontStyle.Bold | FontStyle.Italic, GraphicsUnit.Point, 0);
            Fii.ForeColor = Color.FromArgb(0, 64, 0);
            Fii.Location = new Point(13, 9);
            Fii.Margin = new Padding(4, 0, 4, 0);
            Fii.Name = "Fii";
            Fii.Size = new Size(66, 44);
            Fii.TabIndex = 16;
            Fii.Text = "Fii";
            // 
            // panel1
            // 
            panel1.BorderStyle = BorderStyle.FixedSingle;
            panel1.Controls.Add(NonFIleMuti);
            panel1.Controls.Add(NonFileLast);
            panel1.Controls.Add(NonFileFirst);
            panel1.Controls.Add(SummitButton);
            panel1.Controls.Add(UploadMultipleFileBtn);
            panel1.Controls.Add(SymbolText1);
            panel1.Controls.Add(MutiFilesText);
            panel1.Controls.Add(BlankTempText);
            panel1.Controls.Add(LastUse);
            panel1.Controls.Add(FirstUse);
            panel1.Controls.Add(UploadLastTemplateBtn);
            panel1.Controls.Add(UploadFirstTemplateBtn);
            panel1.Location = new Point(44, 77);
            panel1.Margin = new Padding(4);
            panel1.Name = "panel1";
            panel1.Size = new Size(742, 412);
            panel1.TabIndex = 17;
            // 
            // Version
            // 
            Version.AutoSize = true;
            Version.Font = new Font("Verdana", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            Version.Location = new Point(13, 523);
            Version.Margin = new Padding(4, 0, 4, 0);
            Version.Name = "Version";
            Version.Size = new Size(125, 22);
            Version.TabIndex = 18;
            Version.Text = "Version : 1.0";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(11F, 23F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(833, 563);
            Controls.Add(Version);
            Controls.Add(panel1);
            Controls.Add(Fii);
            Controls.Add(BySRD3IT);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Margin = new Padding(4);
            Name = "Form1";
            Text = "資料比對工具";
            Load += Form1_Load;
            panel1.ResumeLayout(false);
            panel1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Button UploadFirstTemplateBtn;
        private System.Windows.Forms.Button UploadLastTemplateBtn;
        private System.Windows.Forms.RadioButton FirstUse;
        private System.Windows.Forms.RadioButton LastUse;
        private System.Windows.Forms.Label BlankTempText;
        private System.Windows.Forms.Label MutiFilesText;
        private System.Windows.Forms.Label SymbolText1;
        private System.Windows.Forms.Button UploadMultipleFileBtn;
        private System.Windows.Forms.Button SummitButton;
        private System.Windows.Forms.Label NonFileFirst;
        private System.Windows.Forms.Label NonFileLast;
        private System.Windows.Forms.Label NonFIleMuti;
        private System.Windows.Forms.Label BySRD3IT;
        private System.Windows.Forms.Label Fii;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label Version;
    }
}
