using System;
using System.Data;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;


namespace KsaweryAPP
{
    public partial class ExcelToJSONpl : Form
    {
        private string selectedFile;
        private const string ConfigFilePath = "config.txt";
        public ExcelToJSONpl()
        {
            InitializeComponent();
            var (textFilePath, index, ilosc, cena, wiersz, dataOdStr, dataDoStr, isDetal) = ReadFromConfig();
            textFile.Text = textFilePath;
            indexBox.Text = index;
            iloscBox.Text = ilosc;
            cenaBox.Text = cena;
            wierszBox.Text = wiersz;
            SetDates(dataOdStr, dataDoStr);
            detal.Checked = isDetal.Equals("True");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private void SetDates(string dataOdStr, string dataDoStr)
        {
            DateTime date;
            if (DateTime.TryParse(dataOdStr, out date))
                dataOd.Value = date;
            if (DateTime.TryParse(dataDoStr, out date))
                dataDo.Value = date;
        }
        
        private void BtnSelectFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;*.xlsb|All Files|*.*";

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
                const string start = "{\"ZakresWylaczenGrupKontrah\":0,\"SposLaczUmowyZRabat\":0,\"ZakresKontrah\":0,\"SposLaczPromZUmonNaCene\":2,\"Typ\":0,\"ListaWylaczenKontrah\":[{\"Indeks\": \"SAGA SP. Z O. O.\"}],\"ListaGrupKontrah\":[],\"ZakresKarotek\":1,\"ZaleznaOd\":0,\"ListaKartotek\":[";
                const string srodekTemplate = "{{\"CenaBrutto\":1,\"Waluta\":\"\",\"OdIlosci\":{0},\"Procent\":0,\"Cena\":{1},\"Indeks\":\"{2}\"}},";
                const string koniec = "],\"ListaGrupKart\":[],\"DataOd\":\"{0}\",\"DataDo\":\"{1}\",\"OdIlosci\":0,\"ZakresMag\":0,\"ZakresDok\":0,\"ListaMag\":[],\"ListaDok\":[],\"SposLaczPromZUmonNaBonif\":23,\"ZakresWylaczenKontrah\":1,\"ListaKontrah\":[],\"ListaCech\":[],\"ListaWylaczenGrupKontrah\":[],\"Procent\":0,\"ZakresGrupKontrah\":0,\"Uwagi\":\"\",\"Parametr\":1,\"Opis\":\"{2}\",\"ZakresGrupKart\":0,\"UmowaDla\":\"\"}}";
                //błąd pretiża - nie widzi dokumentów o grupach 150 oraz 220
                //const string koniecDetal = "],\"ListaGrupKart\":[],\"DataOd\":\"{0}\",\"DataDo\":\"{1}\",\"OdIlosci\":0,\"ZakresMag\":0,\"ZakresDok\":1,\"ListaMag\":[],\"ListaDok\":[{{\"GrupaDok\": 10,\"Skrot\": \"PAR\"}},{{\"GrupaDok\": 150,\"Skrot\":\"OFEODB\"}},{{\"GrupaDok\": 150,\"Skrot\": \"OF_ZAMB\"}}, {{\"GrupaDok\": 150, \"Skrot\": \"OF_ZAMD\"}}, {{\"GrupaDok\": 150, \"Skrot\": \"OF_ZAMIN\"}}, {{\"GrupaDok\": 150, \"Skrot\": \"OF_ZAMK\"}}, {{\"GrupaDok\": 220, \"Skrot\": \"ZAOFEO\"}}, {{\"GrupaDok\": 80, \"Skrot\": \"ZAMIN\"}}, {{\"GrupaDok\": 80, \"Skrot\": \"ZAMINC\"}}, {{\"GrupaDok\": 10, \"Skrot\": \"PARA\"}}, {{\"GrupaDok\": 10, \"Skrot\": \"FVAT\"}}, {{\"GrupaDok\": 10, \"Skrot\": \"FDETAL\"}}, {{\"GrupaDok\": 10, \"Skrot\": \"FRA BON\"}}, {{\"GrupaDok\": 80, \"Skrot\": \"ZAMB\"}}, {{\"GrupaDok\": 80, \"Skrot\": \"ZAMK\"}}, {{\"GrupaDok\": 80, \"Skrot\": \"ZAMD\"}}, {{\"GrupaDok\": 10, \"Skrot\": \"FVATD\"}}],\"SposLaczPromZUmonNaBonif\":23,\"ZakresWylaczenKontrah\":1,\"ListaKontrah\":[],\"ListaCech\":[],\"ListaWylaczenGrupKontrah\":[],\"Procent\":0,\"ZakresGrupKontrah\":0,\"Uwagi\":\"\",\"Parametr\":1,\"Opis\":\"{2}\",\"ZakresGrupKart\":0,\"UmowaDla\":\"\"}}";
                const string koniecDetal = "],\"ListaGrupKart\":[],\"DataOd\":\"{0}\",\"DataDo\":\"{1}\",\"OdIlosci\":0,\"ZakresMag\":0,\"ZakresDok\":1,\"ListaMag\":[],\"ListaDok\":[{{\"GrupaDok\": 10,\"Skrot\": \"PAR\"}}, {{\"GrupaDok\": 80, \"Skrot\": \"ZAMIN\"}}, {{\"GrupaDok\": 80, \"Skrot\": \"ZAMINC\"}}, {{\"GrupaDok\": 10, \"Skrot\": \"PARA\"}}, {{\"GrupaDok\": 10, \"Skrot\": \"FVAT\"}}, {{\"GrupaDok\": 10, \"Skrot\": \"FDETAL\"}}, {{\"GrupaDok\": 10, \"Skrot\": \"FRA BON\"}}, {{\"GrupaDok\": 80, \"Skrot\": \"ZAMB\"}}, {{\"GrupaDok\": 80, \"Skrot\": \"ZAMK\"}}, {{\"GrupaDok\": 80, \"Skrot\": \"ZAMD\"}}, {{\"GrupaDok\": 10, \"Skrot\": \"FVATD\"}}],\"SposLaczPromZUmonNaBonif\":23,\"ZakresWylaczenKontrah\":1,\"ListaKontrah\":[],\"ListaCech\":[],\"ListaWylaczenGrupKontrah\":[],\"Procent\":0,\"ZakresGrupKontrah\":0,\"Uwagi\":\"\",\"Parametr\":1,\"Opis\":\"{2}\",\"ZakresGrupKart\":0,\"UmowaDla\":\"\"}}";
                bool detal = this.detal.Checked;
                string dataOdStr = "", dataDoStr = "";
                DateTime parsedDate;
                if (DateTime.TryParse(dataOd.Text, out parsedDate))
                {
                    dataOdStr = parsedDate.ToString("yyyy-MM-dd");
                }
                if (DateTime.TryParse(dataDo.Text, out parsedDate))
                {
                    dataDoStr = parsedDate.ToString("yyyy-MM-dd");
                }
                //var dataDoStr = dataDo.Text.Replace('.', '-');
                string opis = textJSON.Text.ToString();
                string stringKoniec = !detal ? string.Format(koniec, dataOdStr, dataDoStr, opis) : string.Format(koniecDetal, dataOdStr, dataDoStr, opis);
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                SaveToConfig(textFile.Text, indexBox.Text, iloscBox.Text, cenaBox.Text, wierszBox.Text, dataOd.Text, dataDo.Text, detal.ToString());

                StringBuilder srodekBuilder = new StringBuilder();

                for (int row = int.Parse(wierszBox.Text); row <= rowCount; row++)
                {
                    string odIlosci = worksheet.Cells[row, ConvertColumnLetterToNumber(iloscBox.Text)].Text;
                    odIlosci = odIlosci.Replace("\r\n", "").Replace("\n", "");
                    string cena = worksheet.Cells[row, ConvertColumnLetterToNumber(cenaBox.Text)].Text.Replace(",", ".");
                    string indeks = worksheet.Cells[row, ConvertColumnLetterToNumber(indexBox.Text)].Text;

                    if (string.IsNullOrEmpty(odIlosci) || string.IsNullOrEmpty(cena) || string.IsNullOrEmpty(indeks)) continue;

                    string srodekRow = string.Format(srodekTemplate, odIlosci, cena, indeks);
                    srodekBuilder.AppendLine(srodekRow);
                }
                string srodek = srodekBuilder.ToString().TrimEnd(',');
                srodek = srodek.Substring(0, srodek.Length - 3);
                string finalJson = start + Environment.NewLine + srodek + Environment.NewLine + stringKoniec;
                System.IO.File.WriteAllText(jsonOutputPath, finalJson);
                MessageBox.Show("Wyeksportowano pomyślnie!", "Eksport", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void SaveToConfig(string textFilePath, string index, string ilosc, string cena, string wiersz, string dataOdStr, string dataDoStr, string isDetal)
        {
            string configData = $"{textFilePath}|{index}|{ilosc}|{cena}|{wiersz}|{dataOdStr}|{dataDoStr}|{isDetal}";
            SaveToLine(1, configData);
        }

        private (string textFilePath, string index, string ilosc, string cena, string wiersz, string dataOd, string dataDo, string isDetal)
            ReadFromConfig()
        {
            string readedLine = ReadFromLine(1);
            var parts = readedLine.Split('|');
            if (parts.Length == 8)
            {
                return (parts[0], parts[1], parts[2], parts[3],
                    parts[4], parts[5], parts[6], parts[7]);
            }
            return (string.Empty, string.Empty, string.Empty, string.Empty,
                string.Empty, string.Empty,  string.Empty , string.Empty);
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

        private string ReadFromLine(int lineNumber)
        {
            if (System.IO.File.Exists(ConfigFilePath))
            {
                var lines = System.IO.File.ReadAllLines(ConfigFilePath);
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
                MessageBox.Show("Proszę wybrać plik.", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (string.IsNullOrEmpty(textJSON.Text))
            {
                MessageBox.Show("Proszę wpisać nazwę.", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            // Sprawdzanie, czy wierszBox, indexBox, cenaBox i iloscBox są liczbami całkowitymi
            if (!int.TryParse(wierszBox.Text, out _))
            {
                MessageBox.Show("Nieprawidłowa wartość w polu Wiersz. Proszę wpisać liczbę całkowitą.", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }
    }
}
