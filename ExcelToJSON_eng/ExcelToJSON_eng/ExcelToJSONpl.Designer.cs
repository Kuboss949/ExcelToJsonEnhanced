namespace KsaweryAPP
{
    partial class ExcelToJSONpl
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
            label1 = new System.Windows.Forms.Label();
            wierszBox = new System.Windows.Forms.TextBox();
            Cena = new System.Windows.Forms.Label();
            Ilość = new System.Windows.Forms.Label();
            Index = new System.Windows.Forms.Label();
            cenaBox = new System.Windows.Forms.TextBox();
            iloscBox = new System.Windows.Forms.TextBox();
            indexBox = new System.Windows.Forms.TextBox();
            nazwaJSON = new System.Windows.Forms.Label();
            textJSON = new System.Windows.Forms.TextBox();
            BtnGenerateJSON = new System.Windows.Forms.Button();
            File = new System.Windows.Forms.Label();
            textFile = new System.Windows.Forms.TextBox();
            BtnSelectFile = new System.Windows.Forms.Button();
            dataDo = new System.Windows.Forms.DateTimePicker();
            label2 = new System.Windows.Forms.Label();
            label3 = new System.Windows.Forms.Label();
            dataOd = new System.Windows.Forms.DateTimePicker();
            detal = new System.Windows.Forms.CheckBox();
            SuspendLayout();
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new System.Drawing.Point(113, 88);
            label1.Name = "label1";
            label1.Size = new System.Drawing.Size(53, 20);
            label1.TabIndex = 27;
            label1.Text = "Wiersz";
            // 
            // wierszBox
            // 
            wierszBox.Location = new System.Drawing.Point(167, 84);
            wierszBox.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            wierszBox.Name = "wierszBox";
            wierszBox.Size = new System.Drawing.Size(114, 27);
            wierszBox.TabIndex = 26;
            // 
            // Cena
            // 
            Cena.AutoSize = true;
            Cena.Location = new System.Drawing.Point(726, 88);
            Cena.Name = "Cena";
            Cena.Size = new System.Drawing.Size(42, 20);
            Cena.TabIndex = 25;
            Cena.Text = "Cena";
            // 
            // Ilość
            // 
            Ilość.AutoSize = true;
            Ilość.Location = new System.Drawing.Point(504, 88);
            Ilość.Name = "Ilość";
            Ilość.Size = new System.Drawing.Size(67, 20);
            Ilość.TabIndex = 24;
            Ilość.Text = "Od Ilosci";
            // 
            // Index
            // 
            Index.AutoSize = true;
            Index.Location = new System.Drawing.Point(304, 88);
            Index.Name = "Index";
            Index.Size = new System.Drawing.Size(51, 20);
            Index.TabIndex = 23;
            Index.Text = "Indeks";
            // 
            // cenaBox
            // 
            cenaBox.Location = new System.Drawing.Point(771, 84);
            cenaBox.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            cenaBox.Name = "cenaBox";
            cenaBox.Size = new System.Drawing.Size(114, 27);
            cenaBox.TabIndex = 22;
            // 
            // iloscBox
            // 
            iloscBox.Location = new System.Drawing.Point(571, 84);
            iloscBox.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            iloscBox.Name = "iloscBox";
            iloscBox.Size = new System.Drawing.Size(114, 27);
            iloscBox.TabIndex = 21;
            // 
            // indexBox
            // 
            indexBox.Location = new System.Drawing.Point(358, 84);
            indexBox.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            indexBox.Name = "indexBox";
            indexBox.Size = new System.Drawing.Size(114, 27);
            indexBox.TabIndex = 20;
            // 
            // nazwaJSON
            // 
            nazwaJSON.AutoSize = true;
            nazwaJSON.Location = new System.Drawing.Point(14, 135);
            nazwaJSON.Name = "nazwaJSON";
            nazwaJSON.Size = new System.Drawing.Size(93, 20);
            nazwaJSON.TabIndex = 19;
            nazwaJSON.Text = "Nazwa JSON";
            // 
            // textJSON
            // 
            textJSON.Location = new System.Drawing.Point(14, 159);
            textJSON.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            textJSON.Name = "textJSON";
            textJSON.Size = new System.Drawing.Size(114, 27);
            textJSON.TabIndex = 18;
            // 
            // BtnGenerateJSON
            // 
            BtnGenerateJSON.Location = new System.Drawing.Point(14, 197);
            BtnGenerateJSON.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            BtnGenerateJSON.Name = "BtnGenerateJSON";
            BtnGenerateJSON.Size = new System.Drawing.Size(109, 31);
            BtnGenerateJSON.TabIndex = 17;
            BtnGenerateJSON.Text = "Generate JSON";
            BtnGenerateJSON.UseVisualStyleBackColor = true;
            BtnGenerateJSON.Click += BtnGenerateJSON_Click;
            // 
            // File
            // 
            File.AutoSize = true;
            File.Location = new System.Drawing.Point(14, 20);
            File.Name = "File";
            File.Size = new System.Drawing.Size(32, 20);
            File.TabIndex = 16;
            File.Text = "File";
            // 
            // textFile
            // 
            textFile.Location = new System.Drawing.Point(14, 44);
            textFile.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            textFile.Name = "textFile";
            textFile.Size = new System.Drawing.Size(871, 27);
            textFile.TabIndex = 15;
            // 
            // BtnSelectFile
            // 
            BtnSelectFile.Location = new System.Drawing.Point(14, 83);
            BtnSelectFile.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            BtnSelectFile.Name = "BtnSelectFile";
            BtnSelectFile.Size = new System.Drawing.Size(86, 31);
            BtnSelectFile.TabIndex = 14;
            BtnSelectFile.Text = "Select File";
            BtnSelectFile.UseVisualStyleBackColor = true;
            BtnSelectFile.Click += BtnSelectFile_Click;
            // 
            // dataDo
            // 
            dataDo.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            dataDo.Location = new System.Drawing.Point(388, 159);
            dataDo.Name = "dataDo";
            dataDo.Size = new System.Drawing.Size(130, 27);
            dataDo.TabIndex = 29;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new System.Drawing.Point(208, 135);
            label2.Name = "label2";
            label2.Size = new System.Drawing.Size(63, 20);
            label2.TabIndex = 30;
            label2.Text = "Data od";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new System.Drawing.Point(388, 135);
            label3.Name = "label3";
            label3.Size = new System.Drawing.Size(63, 20);
            label3.TabIndex = 31;
            label3.Text = "Data do";
            // 
            // dataOd
            // 
            dataOd.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            dataOd.Location = new System.Drawing.Point(208, 159);
            dataOd.Name = "dataOd";
            dataOd.Size = new System.Drawing.Size(130, 27);
            dataOd.TabIndex = 32;
            // 
            // detal
            // 
            detal.Location = new System.Drawing.Point(555, 161);
            detal.Name = "detal";
            detal.Size = new System.Drawing.Size(104, 24);
            detal.TabIndex = 33;
            detal.Text = "Detal";
            detal.UseVisualStyleBackColor = true;
            // 
            // ExcelToJSONpl
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(914, 600);
            Controls.Add(detal);
            Controls.Add(dataOd);
            Controls.Add(label3);
            Controls.Add(label2);
            Controls.Add(dataDo);
            Controls.Add(label1);
            Controls.Add(wierszBox);
            Controls.Add(Cena);
            Controls.Add(Ilość);
            Controls.Add(Index);
            Controls.Add(cenaBox);
            Controls.Add(iloscBox);
            Controls.Add(indexBox);
            Controls.Add(nazwaJSON);
            Controls.Add(textJSON);
            Controls.Add(BtnGenerateJSON);
            Controls.Add(File);
            Controls.Add(textFile);
            Controls.Add(BtnSelectFile);
            Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            Text = "ExcelToJSONpl";
            ResumeLayout(false);
            PerformLayout();
        }

        private System.Windows.Forms.CheckBox detal;

        private System.Windows.Forms.DateTimePicker dataOd;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;

        private System.Windows.Forms.DateTimePicker dataDo;

        #endregion

        private Label label1;
        private TextBox wierszBox;
        private Label Cena;
        private Label Ilość;
        private Label Index;
        private TextBox cenaBox;
        private TextBox iloscBox;
        private TextBox indexBox;
        private Label nazwaJSON;
        private TextBox textJSON;
        private Button BtnGenerateJSON;
        private Label File;
        private TextBox textFile;
        private Button BtnSelectFile;
    }
}