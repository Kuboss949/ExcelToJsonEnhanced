namespace ExcelToJSON_eng
{
    partial class ExcelToJSONeng
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            label1 = new System.Windows.Forms.Label();
            wierszBox = new System.Windows.Forms.TextBox();
            Cena = new System.Windows.Forms.Label();
            Ilość = new System.Windows.Forms.Label();
            Index = new System.Windows.Forms.Label();
            cenaBox0 = new System.Windows.Forms.TextBox();
            iloscBox0 = new System.Windows.Forms.TextBox();
            indexBox = new System.Windows.Forms.TextBox();
            nazwaJSON = new System.Windows.Forms.Label();
            textJSON = new System.Windows.Forms.TextBox();
            BtnGenerateJSON = new System.Windows.Forms.Button();
            textFile = new System.Windows.Forms.TextBox();
            BtnSelectFile = new System.Windows.Forms.Button();
            File = new System.Windows.Forms.Label();
            textWALUTA = new System.Windows.Forms.TextBox();
            label4 = new System.Windows.Forms.Label();
            dataOd = new System.Windows.Forms.DateTimePicker();
            label7 = new System.Windows.Forms.Label();
            boxPanel = new System.Windows.Forms.Panel();
            addBox = new System.Windows.Forms.Button();
            deleteBox = new System.Windows.Forms.Button();
            templateBox = new System.Windows.Forms.TextBox();
            boxPanel.SuspendLayout();
            SuspendLayout();
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new System.Drawing.Point(113, 87);
            label1.Name = "label1";
            label1.Size = new System.Drawing.Size(41, 15);
            label1.TabIndex = 26;
            label1.Text = "Wiersz";
            // 
            // wierszBox
            // 
            wierszBox.Location = new System.Drawing.Point(167, 83);
            wierszBox.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            wierszBox.Name = "wierszBox";
            wierszBox.Size = new System.Drawing.Size(114, 23);
            wierszBox.TabIndex = 25;
            // 
            // Cena
            // 
            Cena.AutoSize = true;
            Cena.Location = new System.Drawing.Point(712, 106);
            Cena.Name = "Cena";
            Cena.Size = new System.Drawing.Size(34, 15);
            Cena.TabIndex = 24;
            Cena.Text = "Cena";
            // 
            // Ilość
            // 
            Ilość.AutoSize = true;
            Ilość.Location = new System.Drawing.Point(565, 106);
            Ilość.Name = "Ilość";
            Ilość.Size = new System.Drawing.Size(53, 15);
            Ilość.TabIndex = 23;
            Ilość.Text = "Od Ilosci";
            // 
            // Index
            // 
            Index.AutoSize = true;
            Index.Location = new System.Drawing.Point(304, 87);
            Index.Name = "Index";
            Index.Size = new System.Drawing.Size(41, 15);
            Index.TabIndex = 22;
            Index.Text = "Indeks";
            // 
            // cenaBox
            // 
            cenaBox0.Location = new System.Drawing.Point(141, 11);
            cenaBox0.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            cenaBox0.Name = "cenaBox0";
            cenaBox0.Size = new System.Drawing.Size(114, 23);
            cenaBox0.TabIndex = 21;
            cenaBox0.TextChanged += cenaBox_TextChanged;
            // 
            // iloscBox
            // 
            iloscBox0.Location = new System.Drawing.Point(3, 11);
            iloscBox0.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            iloscBox0.Name = "iloscBox0";
            iloscBox0.Size = new System.Drawing.Size(114, 23);
            iloscBox0.TabIndex = 20;
            // 
            // indexBox
            // 
            indexBox.Location = new System.Drawing.Point(358, 83);
            indexBox.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            indexBox.Name = "indexBox";
            indexBox.Size = new System.Drawing.Size(114, 23);
            indexBox.TabIndex = 19;
            // 
            // nazwaJSON
            // 
            nazwaJSON.AutoSize = true;
            nazwaJSON.Location = new System.Drawing.Point(14, 133);
            nazwaJSON.Name = "nazwaJSON";
            nazwaJSON.Size = new System.Drawing.Size(73, 15);
            nazwaJSON.TabIndex = 18;
            nazwaJSON.Text = "Nazwa JSON";
            // 
            // textJSON
            // 
            textJSON.Location = new System.Drawing.Point(14, 157);
            textJSON.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            textJSON.Name = "textJSON";
            textJSON.Size = new System.Drawing.Size(114, 23);
            textJSON.TabIndex = 17;
            // 
            // BtnGenerateJSON
            // 
            BtnGenerateJSON.Location = new System.Drawing.Point(14, 265);
            BtnGenerateJSON.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            BtnGenerateJSON.Name = "BtnGenerateJSON";
            BtnGenerateJSON.Size = new System.Drawing.Size(109, 31);
            BtnGenerateJSON.TabIndex = 16;
            BtnGenerateJSON.Text = "Generate JSON";
            BtnGenerateJSON.UseVisualStyleBackColor = true;
            BtnGenerateJSON.Click += BtnGenerateJSON_Click;
            // 
            // textFile
            // 
            textFile.Location = new System.Drawing.Point(14, 43);
            textFile.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            textFile.Name = "textFile";
            textFile.Size = new System.Drawing.Size(871, 23);
            textFile.TabIndex = 15;
            // 
            // BtnSelectFile
            // 
            BtnSelectFile.Location = new System.Drawing.Point(14, 81);
            BtnSelectFile.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            BtnSelectFile.Name = "BtnSelectFile";
            BtnSelectFile.Size = new System.Drawing.Size(86, 31);
            BtnSelectFile.TabIndex = 14;
            BtnSelectFile.Text = "Select File";
            BtnSelectFile.UseVisualStyleBackColor = true;
            BtnSelectFile.Click += BtnSelectFile_Click;
            // 
            // File
            // 
            File.AutoSize = true;
            File.Location = new System.Drawing.Point(14, 19);
            File.Name = "File";
            File.Size = new System.Drawing.Size(25, 15);
            File.TabIndex = 27;
            File.Text = "File";
            // 
            // textWALUTA
            // 
            textWALUTA.Location = new System.Drawing.Point(157, 157);
            textWALUTA.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            textWALUTA.Name = "textWALUTA";
            textWALUTA.Size = new System.Drawing.Size(114, 23);
            textWALUTA.TabIndex = 41;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new System.Drawing.Point(157, 133);
            label4.Name = "label4";
            label4.Size = new System.Drawing.Size(44, 15);
            label4.TabIndex = 42;
            label4.Text = "Waluta";
            // 
            // dataOd
            // 
            dataOd.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            dataOd.Location = new System.Drawing.Point(14, 216);
            dataOd.Name = "dataOd";
            dataOd.Size = new System.Drawing.Size(114, 23);
            dataOd.TabIndex = 44;
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Location = new System.Drawing.Point(14, 193);
            label7.Name = "label7";
            label7.Size = new System.Drawing.Size(48, 15);
            label7.TabIndex = 46;
            label7.Text = "Data od";
            // 
            // boxPanel
            // 
            boxPanel.Controls.Add(iloscBox0);
            boxPanel.Controls.Add(cenaBox0);
            boxPanel.Location = new System.Drawing.Point(533, 133);
            boxPanel.Name = "boxPanel";
            boxPanel.Size = new System.Drawing.Size(258, 455);
            boxPanel.TabIndex = 47;
            // 
            // addBox
            // 
            addBox.BackColor = System.Drawing.Color.LightGreen;
            addBox.Font = new System.Drawing.Font("Segoe UI", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            addBox.Location = new System.Drawing.Point(808, 133);
            addBox.Name = "addBox";
            addBox.Size = new System.Drawing.Size(38, 34);
            addBox.TabIndex = 48;
            addBox.Text = "+";
            addBox.UseVisualStyleBackColor = false;
            addBox.Click += addBox_Click;
            // 
            // deleteBox
            // 
            deleteBox.BackColor = System.Drawing.Color.LightCoral;
            deleteBox.Font = new System.Drawing.Font("Segoe UI", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            deleteBox.Location = new System.Drawing.Point(808, 173);
            deleteBox.Name = "deleteBox";
            deleteBox.Size = new System.Drawing.Size(38, 34);
            deleteBox.TabIndex = 49;
            deleteBox.Text = "-";
            deleteBox.UseVisualStyleBackColor = false;
            deleteBox.Click += deleteBox_Click;
            // 
            // templateBox
            // 
            templateBox.Location = new System.Drawing.Point(31, 542);
            templateBox.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            templateBox.Name = "templateBox";
            templateBox.Size = new System.Drawing.Size(114, 23);
            templateBox.TabIndex = 50;
            templateBox.Visible = false;
            // 
            // ExcelToJSONeng
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(914, 600);
            Controls.Add(templateBox);
            Controls.Add(deleteBox);
            Controls.Add(addBox);
            Controls.Add(boxPanel);
            Controls.Add(label7);
            Controls.Add(dataOd);
            Controls.Add(label4);
            Controls.Add(textWALUTA);
            Controls.Add(File);
            Controls.Add(label1);
            Controls.Add(wierszBox);
            Controls.Add(Cena);
            Controls.Add(Ilość);
            Controls.Add(Index);
            Controls.Add(indexBox);
            Controls.Add(nazwaJSON);
            Controls.Add(textJSON);
            Controls.Add(BtnGenerateJSON);
            Controls.Add(textFile);
            Controls.Add(BtnSelectFile);
            Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            Text = "Form1";
            boxPanel.ResumeLayout(false);
            boxPanel.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        private System.Windows.Forms.TextBox templateBox;

        private System.Windows.Forms.Button deleteBox;

        private System.Windows.Forms.Panel boxPanel;
        private System.Windows.Forms.Button addBox;

        private System.Windows.Forms.DateTimePicker dataOd;
        private System.Windows.Forms.Label label7;

        #endregion

        private Label label1;
        private TextBox wierszBox;
        private System.Windows.Forms.Label Cena;
        private System.Windows.Forms.Label Ilość;
        private Label Index;
        private System.Windows.Forms.TextBox cenaBox0;
        private System.Windows.Forms.TextBox iloscBox0;
        private TextBox indexBox;
        private Label nazwaJSON;
        private TextBox textJSON;
        private System.Windows.Forms.Button BtnGenerateJSON;
        private TextBox textFile;
        private Button BtnSelectFile;
        private Label File;
        private TextBox textWALUTA;
        private Label label4;
    }
}