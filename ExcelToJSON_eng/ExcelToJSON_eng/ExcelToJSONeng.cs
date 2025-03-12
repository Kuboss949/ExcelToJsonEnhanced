using System;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using TextBox = System.Windows.Forms.TextBox;

namespace ExcelToJSON_eng
{
    public partial class ExcelToJSONeng : Form
    {
        private const string ConfigFilePath = "config.txt";
        private const string ConfigFilePath2 = "config2.txt";
        private string start;
        private string srodekTemplate;
        private string koniec;
        private Point[] firstBoxLocation;
        private int boxCounter;
        private const int boxMargin = 40;
        
        public ExcelToJSONeng()
        {
            InitializeComponent();
            boxPanel.Controls.Add(iloscBox0);
            boxPanel.Controls.Add(cenaBox0);
            firstBoxLocation = new [] {iloscBox0.Location, cenaBox0.Location};
            boxCounter = 1;
            LoadFromConfig();
            (start, srodekTemplate, koniec) = ReadFromConfig2();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private void BtnSelectFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                textFile.Text = openFileDialog.FileName;
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
                string dataOdStr = "";
                DateTime parsedDate;
                if (DateTime.TryParse(dataOd.Text, out parsedDate))
                {
                    dataOdStr = parsedDate.ToString("yyyy-MM-dd");
                }
                SaveToConfig(textFile.Text, indexBox.Text, wierszBox.Text, waluta);
                koniec = "], \"ListaGrupKart\":[], \"DataOd\":\"{0}\",\"DataDo\":\"\",\"OdIlosci\":0,\"ZakresMag\":0,\"ZakresDok\":0," +
                         "\"ListaMag\":[], \"ListaDok\":[], \"SposLaczPromZUmonNaBonif\":0,\"ZakresWylaczenKontrah\":0," +
                         "\"ListaCech\":[], \"Procent\":0,\"ZakresGrupKontrah\":1,\"Uwagi\":\"\",\"Parametr\":1," +
                         "\"Opis\":\"{1}\",\"ZakresGrupKart\":0,\"UmowaDla\":\"35\"}}";
                string stringKoniec = string.Format(koniec, dataOdStr, opis);
                var worksheet = package.Workbook.Worksheets[0];
                List<RowData> rows = CreateRowList(worksheet, waluta);
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
        
        private void SaveToConfig(string textFilePath, string index, string wiersz, string waluta)
        {
            string configData = $"{textFilePath}|{index}|{wiersz}|{waluta}|{boxCounter}";
            boxPanel.Controls.OfType<TextBox>().OrderBy(c => c.Location.Y).ToList().ForEach(control =>
            {
                configData += $"|{control.Name}:{control.Text}";
            });
            SaveToLine(2, configData);
        }

        private bool LoadFromConfig()
        {
            string readedLine = ReadFromLine(ConfigFilePath,2);
            var parts = readedLine.Split('|');
            if (parts.Length < 5)
            {
                return false;
            }
            textFile.Text = parts[0];
            indexBox.Text = parts[1];
            wierszBox.Text = parts[2];
            textWALUTA.Text = parts[3];
            
            int counter = int.Parse(parts[4]);
            for (int i = 0; i < counter - 1; i++)
                AddInputBoxToPanel();
            
            boxPanel.Controls.OfType<TextBox>().ToList().ForEach(control =>
            {
                var boxName = control.Name;
                foreach (var part in parts)
                {
                    var substrings = part.Split(':');
                    if (substrings[0] == control.Name)
                    {
                        control.Text = substrings[1];
                        break;
                    }
                }
            });
            return true;
        }
        
        private (string start, string srodekTemplate, string koniec) ReadFromConfig2()
        {
            string start = ReadFromLine(ConfigFilePath2, 1);
            string srodekTemplate = ReadFromLine(ConfigFilePath2, 2);
            srodekTemplate = "{{\"CenaBrutto\":0,\"Waluta\":\"{3}\",\"OdIlosci\":{0},\"Procent\":0,\"Cena\":{1},\"Indeks\":\"{2}\"}},";
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

        private int ConvertColumnLetterToNumber(string columnLetter)
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
                MessageBox.Show("Proszę wybrać plik.", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (string.IsNullOrEmpty(textJSON.Text))
            {
                MessageBox.Show("Proszę wpisać nazwę.", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            
            if (!int.TryParse(wierszBox.Text, out _))
            {
                MessageBox.Show("Nieprawidłowa wartość w polu Wiersz. Proszę wpisać liczbę całkowitą.", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            
            return true;
        }

        private void cenaBox_TextChanged(object sender, EventArgs e)
        {
            //throw new System.NotImplementedException();
        }

        private void addBox_Click(object sender, EventArgs e)
        {
            AddInputBoxToPanel();
        }

        private void deleteBox_Click(object sender, EventArgs e)
        {
            DeleteInputBoxFromPanel();
        }

        private void AddInputBoxToPanel()
        {
            int margin = boxMargin * boxCounter;
            TextBox newIloscBox = new TextBox();
            newIloscBox.Name = "iloscBox" + boxCounter;
            newIloscBox.Size = iloscBox0.Size;
            newIloscBox.Location = new Point(firstBoxLocation[0].X, firstBoxLocation[0].Y + margin);
            
            TextBox newCenaBox = new TextBox();
            newCenaBox.Name = "cenaBox" + boxCounter;
            newCenaBox.Size = cenaBox0.Size;
            newCenaBox.Location = new Point(firstBoxLocation[1].X, firstBoxLocation[1].Y + margin);
            
            boxPanel.Controls.Add(newIloscBox);
            boxPanel.Controls.Add(newCenaBox);

            boxCounter++;
        }

        private void DeleteInputBoxFromPanel()
        {
            if (boxCounter <= 1) return;
            for (int i = 0; i < 2; i++)
            {
                var maxYControl = boxPanel.Controls
                    .OfType<Control>()
                    .OrderByDescending(c => c.Location.Y)
                    .FirstOrDefault();
                if (maxYControl != null)
                {
                    boxPanel.Controls.Remove(maxYControl);
                }
            }
            boxCounter--;
        }

        private List<RowData> CreateRowList(ExcelWorksheet worksheet, string waluta)
        {
            int rowCount = worksheet.Dimension.Rows;
            List<RowData> rows = new List<RowData>();
            var iloscBoxes = boxPanel.Controls
                .OfType<TextBox>()
                .Where(c => c.Name.StartsWith("iloscBox"))
                .ToArray();
            var cenaBoxes = boxPanel.Controls
                .OfType<TextBox>()
                .Where(c => c.Name.StartsWith("cenaBox"))
                .ToArray();
            for (int i = 0; i < boxCounter; i++)
            {
                for (int row = int.Parse(wierszBox.Text); row <= rowCount; row++)
                {
                    string odIlosci = worksheet.Cells[row, ConvertColumnLetterToNumber(iloscBoxes[i].Text)].Text;
                    odIlosci = odIlosci.Replace("\r\n", "").Replace("\n", "");
                    string cena = worksheet.Cells[row, ConvertColumnLetterToNumber(cenaBoxes[i].Text)].Text
                        .Replace(",", ".");
                    string indeks = worksheet.Cells[row, ConvertColumnLetterToNumber(indexBox.Text)].Text;

                    if (string.IsNullOrEmpty(odIlosci) || string.IsNullOrEmpty(cena) ||
                        string.IsNullOrEmpty(indeks)) continue;

                    string srodekRow = string.Format(srodekTemplate, odIlosci, cena, indeks, waluta);
                    rows.Add(new RowData
                        { Indeks = indeks, OdIlosci = odIlosci, Cena = cena, FormattedRow = srodekRow });
                }
            }

            return rows;
        }
    }
}