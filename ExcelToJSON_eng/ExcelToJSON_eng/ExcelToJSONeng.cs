using System;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace ExcelToJSON_eng
{
    public partial class ExcelToJSONeng : Form
    {
        private string selectedFile;
        private const string ConfigFilePath = "config.txt";
        private const string ConfigFilePath2 = "config2.txt";
        private string start;
        private string srodekTemplate;
        private string koniec;
        public ExcelToJSONeng()
        {
            InitializeComponent();
            var (textFilePath, index, ilosc, cena, wiersz, ilosc2, ilosc3, cena2, cena3, waluta) = ReadFromConfig();
            textFile.Text = textFilePath;
            indexBox.Text = index;
            iloscBox.Text = ilosc;
            cenaBox.Text = cena;
            wierszBox.Text = wiersz;
            iloscBox2.Text = ilosc2;
            iloscBox3.Text = ilosc3;
            cenaBox2.Text = cena2;
            cenaBox3.Text = cena3;
            textWALUTA.Text = waluta;
            (start, srodekTemplate, koniec) = ReadFromConfig2();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private void BtnSelectFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                textFile.Text = openFileDialog.FileName;
                selectedFile = openFileDialog.FileName;
            }
        }

        private void BtnGenerateJSON_Click(object sender, EventArgs e)
        {
            if (!validationCreateJsonFromExcel()) return;
            string excelFilePath = textFile.Text;
            string directory = Path.GetDirectoryName(excelFilePath);
            string jsonOutputPath = Path.Combine(directory, $"{textJSON.Text}.json");

            CreateJsonFromExcel(excelFilePath, jsonOutputPath);
        }

        private void CreateJsonFromExcel(string excelFilePath, string jsonOutputPath)
        {
            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                string opis = textJSON.Text;
                string waluta = textWALUTA.Text;
                string tomorrowDate = DateTime.Now.AddDays(1).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                string stringKoniec = string.Format(koniec, tomorrowDate, opis);
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;
                List<RowData> rows = new List<RowData>();

                SaveToConfig(textFile.Text, indexBox.Text, iloscBox.Text, cenaBox.Text, wierszBox.Text, iloscBox2.Text, iloscBox3.Text, cenaBox2.Text, cenaBox3.Text, textWALUTA.Text);

                //StringBuilder srodekBuilder = new StringBuilder();
                System.Windows.Forms.TextBox[] iloscBoxes = { iloscBox, iloscBox2, iloscBox3 };
                System.Windows.Forms.TextBox[] cenaBoxes = { cenaBox, cenaBox2, cenaBox3 };
                for(int i = 0; i < 3; i++)
                {
                    for (int row = int.Parse(wierszBox.Text); row <= rowCount; row++)
                    {
                        string odIlosci = worksheet.Cells[row, ConvertColumnLetterToNumber(iloscBoxes[i].Text)].Text;
                        odIlosci = odIlosci.Replace("\r\n", "").Replace("\n", "");
                        string cena = worksheet.Cells[row, ConvertColumnLetterToNumber(cenaBoxes[i].Text)].Text.Replace(",", ".");
                        string indeks = worksheet.Cells[row, ConvertColumnLetterToNumber(indexBox.Text)].Text;

                        if (string.IsNullOrEmpty(odIlosci) || string.IsNullOrEmpty(cena) || string.IsNullOrEmpty(indeks)) continue;

                        string srodekRow = string.Format(srodekTemplate, odIlosci, cena, indeks, waluta);
                        rows.Add(new RowData { Indeks = indeks, OdIlosci = odIlosci, Cena = cena, FormattedRow = srodekRow });
                    }
                }
                rows = rows.OrderBy(r => r.OdIlosci).ToList();

                StringBuilder srodekBuilder = new StringBuilder();
                foreach (var rowData in rows)
                {
                    srodekBuilder.AppendLine(rowData.FormattedRow);
                }
                string srodek = srodekBuilder.ToString().TrimEnd(',');
                srodek = srodek.Substring(0, srodek.Length - 3);
                string finalJson = start + Environment.NewLine + srodek + Environment.NewLine + stringKoniec;
                System.IO.File.WriteAllText(jsonOutputPath, finalJson);
                MessageBox.Show("Wyeksportowano pomy�lnie!", "Eksport", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        public class RowData
        {
            public string Indeks { get; set; }
            public string OdIlosci { get; set; }
            public string Cena { get; set; }
            public string FormattedRow { get; set; }
        }
        private void SaveToConfig(string textFilePath, string index, string ilosc, string cena, string wiersz, string ilosc2, string ilosc3, string cena2, string cena3, string waluta)
        {
            string configData = $"{textFilePath}|{index}|{ilosc}|{cena}|{wiersz}|{ilosc2}|{ilosc3}|{cena2}|{cena3}|{waluta}";
            SaveToLine(2, configData);
        }

        private (string textFilePath, string index, string ilosc, string cena, string wiersz, string ilosc2, string ilosc3, string cena2, string cena3, string waluta) ReadFromConfig()
        {
            string readedLine = ReadFromLine(ConfigFilePath,2);
            var parts = readedLine.Split('|');
            if (parts.Length == 10)
            {
                return (parts[0], parts[1], parts[2], parts[3], parts[4], parts[5], parts[6], parts[7], parts[8], parts[9]);
            }
            return (string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
        }
        private (string start, string srodekTemplate, string koniec) ReadFromConfig2()
        {
            string start = ReadFromLine(ConfigFilePath2, 1);
            string srodekTemplate = ReadFromLine(ConfigFilePath2, 2);
            string koniec = ReadFromLine(ConfigFilePath2,  3);
            return (start, srodekTemplate, koniec);
        }
        private void SaveToLine(int lineNumber, string data)
        {
            List<string> lines = new List<string>();
            if (System.IO.File.Exists(ConfigFilePath))
            {
                lines.AddRange(System.IO.File.ReadAllLines(ConfigFilePath));
            }

            while (lines.Count < lineNumber)
            {
                lines.Add(string.Empty);
            }

            lines[lineNumber - 1] = data;
            System.IO.File.WriteAllLines(ConfigFilePath, lines);
        }

        private string ReadFromLine(string filePath, int lineNumber)
        {
            if (System.IO.File.Exists(filePath))
            {
                var lines = System.IO.File.ReadAllLines(filePath);
                if (lineNumber <= lines.Length)
                {
                    return lines[lineNumber - 1];
                }
            }
            return string.Empty;
        }

        public int ConvertColumnLetterToNumber(string columnLetter)
        {
            int columnNumber = 0;
            foreach (char c in columnLetter.ToUpper())
            {
                columnNumber = columnNumber * 26 + (c - 'A' + 1);
            }
            return columnNumber;
        }
        private bool validationCreateJsonFromExcel()
        {
            if (string.IsNullOrEmpty(textFile.Text))
            {
                MessageBox.Show("Prosz� wybra� plik.", "B��d", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (string.IsNullOrEmpty(textJSON.Text))
            {
                MessageBox.Show("Prosz� wpisa� nazw�.", "B��d", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            // Sprawdzanie, czy wierszBox, indexBox, cenaBox i iloscBox s� liczbami ca�kowitymi
            if (!int.TryParse(wierszBox.Text, out _))
            {
                MessageBox.Show("Nieprawid�owa warto�� w polu Wiersz. Prosz� wpisa� liczb� ca�kowit�.", "B��d", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }
    }
}