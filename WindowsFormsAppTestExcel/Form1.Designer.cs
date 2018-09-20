namespace WindowsFormsAppTestExcel
{
    partial class Form1
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
            this.btnRun = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.textBoxAddinPath = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.textBoxExcelSheetPath = new System.Windows.Forms.TextBox();
            this.btnSelectExcelFile = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.textBoxResult = new System.Windows.Forms.TextBox();
            this.btnParserTest = new System.Windows.Forms.Button();
            this.openExcelFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.label5 = new System.Windows.Forms.Label();
            this.textBoxActiveParserName = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnRun
            // 
            this.btnRun.Location = new System.Drawing.Point(12, 173);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(197, 37);
            this.btnRun.TabIndex = 5;
            this.btnRun.Text = "Запуск надстройки";
            this.btnRun.UseVisualStyleBackColor = true;
            this.btnRun.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(197, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Пример запуска надстройки Парсер.";
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Location = new System.Drawing.Point(15, 25);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(552, 1);
            this.panel1.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 43);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(102, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Путь к надстройке";
            // 
            // textBoxAddinPath
            // 
            this.textBoxAddinPath.Location = new System.Drawing.Point(12, 59);
            this.textBoxAddinPath.Name = "textBoxAddinPath";
            this.textBoxAddinPath.Size = new System.Drawing.Size(491, 20);
            this.textBoxAddinPath.TabIndex = 1;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(8, 121);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(142, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Файл Excel для обработки";
            // 
            // textBoxExcelSheetPath
            // 
            this.textBoxExcelSheetPath.Location = new System.Drawing.Point(11, 137);
            this.textBoxExcelSheetPath.Name = "textBoxExcelSheetPath";
            this.textBoxExcelSheetPath.Size = new System.Drawing.Size(514, 20);
            this.textBoxExcelSheetPath.TabIndex = 3;
            // 
            // btnSelectExcelFile
            // 
            this.btnSelectExcelFile.Location = new System.Drawing.Point(531, 135);
            this.btnSelectExcelFile.Name = "btnSelectExcelFile";
            this.btnSelectExcelFile.Size = new System.Drawing.Size(30, 23);
            this.btnSelectExcelFile.TabIndex = 4;
            this.btnSelectExcelFile.Text = "...";
            this.btnSelectExcelFile.UseVisualStyleBackColor = true;
            this.btnSelectExcelFile.Click += new System.EventHandler(this.btnSelectExcelFile_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(9, 227);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(59, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "Результат";
            // 
            // textBoxResult
            // 
            this.textBoxResult.Location = new System.Drawing.Point(11, 243);
            this.textBoxResult.Multiline = true;
            this.textBoxResult.Name = "textBoxResult";
            this.textBoxResult.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBoxResult.Size = new System.Drawing.Size(514, 120);
            this.textBoxResult.TabIndex = 6;
            // 
            // btnParserTest
            // 
            this.btnParserTest.Location = new System.Drawing.Point(512, 56);
            this.btnParserTest.Name = "btnParserTest";
            this.btnParserTest.Size = new System.Drawing.Size(55, 23);
            this.btnParserTest.TabIndex = 2;
            this.btnParserTest.Text = "Тест";
            this.btnParserTest.UseVisualStyleBackColor = true;
            this.btnParserTest.Click += new System.EventHandler(this.btnParserTest_Click);
            // 
            // openExcelFileDialog
            // 
            this.openExcelFileDialog.Filter = "Excel files|*.xls;*.xlsx|All files|*.*";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(9, 82);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(136, 13);
            this.label5.TabIndex = 9;
            this.label5.Text = "Настройка по умолчанию";
            // 
            // textBoxActiveParserName
            // 
            this.textBoxActiveParserName.Location = new System.Drawing.Point(11, 98);
            this.textBoxActiveParserName.Name = "textBoxActiveParserName";
            this.textBoxActiveParserName.Size = new System.Drawing.Size(492, 20);
            this.textBoxActiveParserName.TabIndex = 10;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(576, 373);
            this.Controls.Add(this.textBoxActiveParserName);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.textBoxAddinPath);
            this.Controls.Add(this.btnParserTest);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.textBoxExcelSheetPath);
            this.Controls.Add(this.btnSelectExcelFile);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.textBoxResult);
            this.Name = "Form1";
            this.Text = "FormMain";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBoxAddinPath;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBoxExcelSheetPath;
        private System.Windows.Forms.Button btnSelectExcelFile;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBoxResult;
        private System.Windows.Forms.Button btnParserTest;
        private System.Windows.Forms.OpenFileDialog openExcelFileDialog;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox textBoxActiveParserName;
    }
}

