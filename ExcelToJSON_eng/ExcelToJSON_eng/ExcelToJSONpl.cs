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
        private FirebirdConnector _db;
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
            _db = new FirebirdConnector("User=SYSDBA;" +
                                        "Password=masterkey;" +
                                        "Database=/data01/db/LYSON.ib;" +
                                        "DataSource=192.168.1.16;" +
                                        "Port=3050;" +
                                        "Dialect=3;" +
                                        "Charset=NONE;");
            SetGroups();
        }

        private void SetGroups()
        {
            var list = _db.GetGroupsKontrah();
            foreach (var group in list)
            {
                GroupsListBox.Items.Add(group);
            }
        }

        private string SelectedGroupsToString()
        {
            string result = "";
            string positionPattern = "{{\"KodZl\": \"{0}\"}},";  // Uwaga: podwójne nawiasy klamrowe

            foreach (var item in GroupsListBox.CheckedItems)
            {
                if (item is GrupaKontrah group)
                {
                    string kodZl = group.KodZl.ToString() ?? string.Empty;
                    result += string.Format(positionPattern, EscapeJsonString(kodZl));
                }
            }

            return result.Length > 0 ? result.TrimEnd(',') : "";
        }
        
        string EscapeJsonString(string input)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;
            return input.Replace("\\", "\\\\").Replace("\"", "\\\"");
        }
        
        private JArray GetSelectedGroupsAsJArray()
        {
            var array = new JArray();

            foreach (var item in GroupsListBox.CheckedItems)
            {
                if (item is GrupaKontrah group)
                {
                    // group.KodZl może być intem albo stringiem, 
                    // dostosuj typ poniżej jeśli trzeba
                    string kodZl = group.KodZl.ToString();
                    // jeśli chcesz escapować np. cudzysłowy wewnątrz kodu:
                    kodZl = EscapeJsonString(kodZl);

                    array.Add(new JObject(
                        new JProperty("KodZl", kodZl)
                    ));
                }
            }

            return array;
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
    // 1) Załaduj Excela
    using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
    {
        var worksheet = package.Workbook.Worksheets[0];
        int rowCount = worksheet.Dimension.Rows;

        // 2) Pobierz parametry z UI
        bool detal = this.detal.Checked;
        string waluta = this.walutaInput.Text;
        string opis = this.textJSON.Text;
        
        // Parsowanie dat
        DateTime tmp;
        string dataOdStr = DateTime.TryParse(this.dataOd.Text, out tmp) 
                           ? tmp.ToString("yyyy-MM-dd") 
                           : "";
        string dataDoStr = DateTime.TryParse(this.dataDo.Text, out tmp) 
                           ? tmp.ToString("yyyy-MM-dd") 
                           : "";

        // 3) Lista grup kontrahentów jako List<int> (zamiast stringa "1,2,3")
        JArray listaGrupKontrahArray = GetSelectedGroupsAsJArray();

        // 5) Tworzymy JArray z wierszami „kartotek”
        var listaKartotekArray = new JArray();
        int colIlosc  = ConvertColumnLetterToNumber(this.iloscBox.Text);
        int colCena   = ConvertColumnLetterToNumber(this.cenaBox.Text);
        int colIndeks = ConvertColumnLetterToNumber(this.indexBox.Text);

        for (int row = int.Parse(this.wierszBox.Text); row <= rowCount; row++)
        {
            string odIlosci = worksheet.Cells[row, colIlosc].Text?.Trim();
            string cenaTxt  = worksheet.Cells[row, colCena].Text?.Replace(",", ".").Trim();
            string indeks   = worksheet.Cells[row, colIndeks].Text?.Trim();

            if (string.IsNullOrEmpty(odIlosci) ||
                string.IsNullOrEmpty(cenaTxt)  ||
                string.IsNullOrEmpty(indeks))
                continue;

            // Konwersja liczbowa
            int    odIl = int.Parse(odIlosci);
            decimal cena = decimal.Parse(cenaTxt, System.Globalization.CultureInfo.InvariantCulture);

            // Jeden wpis kartoteki
            var entry = new JObject(
                new JProperty("CenaBrutto", detal ? 1 : 0),
                new JProperty("Waluta",     waluta),
                new JProperty("OdIlosci",   odIl),
                new JProperty("Procent",    0),
                new JProperty("Cena",       cena),
                new JProperty("Indeks",     indeks)
            );

            listaKartotekArray.Add(entry);
        }

        // 6) ListaDok – pusta lub z predefiniowanymi obiektami, gdy detal==true
        JArray listaDokArray;
        if (!detal)
        {
            listaDokArray = new JArray();
        }
        else
        {
            // tylko zestaw zdefiniowanych „GrupaDok”/„Skrot”
            var doki = new[]
            {
                (10, "PAR"),
                ( 80, "ZAMIN"),
                ( 80, "ZAMINC"),
                ( 10, "PARA"),
                ( 10, "FVAT"),
                ( 10, "FDETAL"),
                ( 10, "FRA BON"),
                ( 80, "ZAMB"),
                ( 80, "ZAMK"),
                ( 80, "ZAMD"),
                ( 10, "FVATD")
            };
            listaDokArray = new JArray(
                doki.Select(d => 
                    new JObject(
                        new JProperty("GrupaDok", d.Item1),
                        new JProperty("Skrot",    d.Item2)
                    )
                )
            );
        }

        // 7) Składamy „root” JObject
        var root = new JObject(
            new JProperty("ZakresWylaczenGrupKontrah",     0),
            new JProperty("SposLaczUmowyZRabat",          0),
            new JProperty("ZakresKontrah",                0),
            new JProperty("SposLaczPromZUmonNaCene",      2),
            new JProperty("Typ",                          0),
            new JProperty("ListaWylaczenKontrah", 
                new JArray(
                    new JObject(new JProperty("Indeks", "SAGA SP. Z O. O."))
                )
            ),
            new JProperty("ListaGrupKontrah",             listaGrupKontrahArray),
            new JProperty("ZakresKarotek",                1),
            new JProperty("ZaleznaOd",                    0),
            new JProperty("ListaKartotek",                listaKartotekArray),
            new JProperty("ListaGrupKart",                new JArray()),
            new JProperty("DataOd",                       dataOdStr),
            new JProperty("DataDo",                       dataDoStr),
            new JProperty("OdIlosci",                     0),
            new JProperty("ZakresMag",                    0),
            new JProperty("ZakresDok",                    detal ? 1 : 0),
            new JProperty("ListaMag",                     new JArray()),
            new JProperty("ListaDok",                     listaDokArray),
            new JProperty("SposLaczPromZUmonNaBonif",     23),
            new JProperty("ZakresWylaczenKontrah",        1),
            new JProperty("ListaKontrah",                 new JArray()),
            new JProperty("ListaCech",                    new JArray()),
            new JProperty("ListaWylaczenGrupKontrah",     new JArray()),
            new JProperty("Procent",                      0),
            new JProperty("ZakresGrupKontrah",            listaGrupKontrahArray.Count > 0 ? 1 : 0),
            new JProperty("Uwagi",                        ""),
            new JProperty("Parametr",                     1),
            new JProperty("Opis",                         opis),
            new JProperty("ZakresGrupKart",               0),
            new JProperty("UmowaDla",                     "")
        );

        // 8) Serializacja do pliku
        System.IO.File.WriteAllText(jsonOutputPath, root.ToString(Formatting.None));

        MessageBox.Show("Wyeksportowano pomyślnie!", "Eksport", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
}
        
        private void CreateJsonFromExcel1(string excelFilePath, string jsonOutputPath)
        {
            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                bool detal = this.detal.Checked;
                string waluta = this.walutaInput.Text;
                string grupyKontrah = SelectedGroupsToString();
                string start = "{\"ZakresWylaczenGrupKontrah\":0,\"SposLaczUmowyZRabat\":0,\"ZakresKontrah\":0,\"SposLaczPromZUmonNaCene\":2,\"Typ\":0,\"ListaWylaczenKontrah\":[{\"Indeks\": \"SAGA SP. Z O. O.\"}],\"ListaGrupKontrah\":[{0}],\"ZakresKarotek\":1,\"ZaleznaOd\":0,\"ListaKartotek\":[";
                
                string srodekTemplate = !detal ? "{{\"CenaBrutto\":0,\"Waluta\":\"{0}\",\"OdIlosci\":{1},\"Procent\":0,\"Cena\":{2},\"Indeks\":\"{3}\"}}," : "{{\"CenaBrutto\":1,\"Waluta\":\"{0}\",\"OdIlosci\":{1},\"Procent\":0,\"Cena\":{2},\"Indeks\":\"{3}\"}},";
                
                const string koniec = "],\"ListaGrupKart\":[],\"DataOd\":\"{0}\",\"DataDo\":\"{1}\",\"OdIlosci\":0,\"ZakresMag\":0,\"ZakresDok\":0,\"ListaMag\":[],\"ListaDok\":[],\"SposLaczPromZUmonNaBonif\":23,\"ZakresWylaczenKontrah\":1,\"ListaKontrah\":[],\"ListaCech\":[],\"ListaWylaczenGrupKontrah\":[],\"Procent\":0,\"ZakresGrupKontrah\":0,\"Uwagi\":\"\",\"Parametr\":1,\"Opis\":\"{2}\",\"ZakresGrupKart\":0,\"UmowaDla\":\"\"}}";
                //błąd pretiża - nie widzi dokumentów o grupach 150 oraz 220
                //const string koniecDetal = "],\"ListaGrupKart\":[],\"DataOd\":\"{0}\",\"DataDo\":\"{1}\",\"OdIlosci\":0,\"ZakresMag\":0,\"ZakresDok\":1,\"ListaMag\":[],\"ListaDok\":[{{\"GrupaDok\": 10,\"Skrot\": \"PAR\"}},{{\"GrupaDok\": 150,\"Skrot\":\"OFEODB\"}},{{\"GrupaDok\": 150,\"Skrot\": \"OF_ZAMB\"}}, {{\"GrupaDok\": 150, \"Skrot\": \"OF_ZAMD\"}}, {{\"GrupaDok\": 150, \"Skrot\": \"OF_ZAMIN\"}}, {{\"GrupaDok\": 150, \"Skrot\": \"OF_ZAMK\"}}, {{\"GrupaDok\": 220, \"Skrot\": \"ZAOFEO\"}}, {{\"GrupaDok\": 80, \"Skrot\": \"ZAMIN\"}}, {{\"GrupaDok\": 80, \"Skrot\": \"ZAMINC\"}}, {{\"GrupaDok\": 10, \"Skrot\": \"PARA\"}}, {{\"GrupaDok\": 10, \"Skrot\": \"FVAT\"}}, {{\"GrupaDok\": 10, \"Skrot\": \"FDETAL\"}}, {{\"GrupaDok\": 10, \"Skrot\": \"FRA BON\"}}, {{\"GrupaDok\": 80, \"Skrot\": \"ZAMB\"}}, {{\"GrupaDok\": 80, \"Skrot\": \"ZAMK\"}}, {{\"GrupaDok\": 80, \"Skrot\": \"ZAMD\"}}, {{\"GrupaDok\": 10, \"Skrot\": \"FVATD\"}}],\"SposLaczPromZUmonNaBonif\":23,\"ZakresWylaczenKontrah\":1,\"ListaKontrah\":[],\"ListaCech\":[],\"ListaWylaczenGrupKontrah\":[],\"Procent\":0,\"ZakresGrupKontrah\":0,\"Uwagi\":\"\",\"Parametr\":1,\"Opis\":\"{2}\",\"ZakresGrupKart\":0,\"UmowaDla\":\"\"}}";
                const string koniecDetal = "],\"ListaGrupKart\":[],\"DataOd\":\"{0}\",\"DataDo\":\"{1}\",\"OdIlosci\":0,\"ZakresMag\":0,\"ZakresDok\":1,\"ListaMag\":[],\"ListaDok\":[{{\"GrupaDok\": 10,\"Skrot\": \"PAR\"}}, {{\"GrupaDok\": 80, \"Skrot\": \"ZAMIN\"}}, {{\"GrupaDok\": 80, \"Skrot\": \"ZAMINC\"}}, {{\"GrupaDok\": 10, \"Skrot\": \"PARA\"}}, {{\"GrupaDok\": 10, \"Skrot\": \"FVAT\"}}, {{\"GrupaDok\": 10, \"Skrot\": \"FDETAL\"}}, {{\"GrupaDok\": 10, \"Skrot\": \"FRA BON\"}}, {{\"GrupaDok\": 80, \"Skrot\": \"ZAMB\"}}, {{\"GrupaDok\": 80, \"Skrot\": \"ZAMK\"}}, {{\"GrupaDok\": 80, \"Skrot\": \"ZAMD\"}}, {{\"GrupaDok\": 10, \"Skrot\": \"FVATD\"}}],\"SposLaczPromZUmonNaBonif\":23,\"ZakresWylaczenKontrah\":1,\"ListaKontrah\":[],\"ListaCech\":[],\"ListaWylaczenGrupKontrah\":[],\"Procent\":0,\"ZakresGrupKontrah\":0,\"Uwagi\":\"\",\"Parametr\":1,\"Opis\":\"{2}\",\"ZakresGrupKart\":0,\"UmowaDla\":\"\"}}";
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

                    string srodekRow = string.Format(srodekTemplate, waluta, odIlosci, cena, indeks);
                    srodekBuilder.AppendLine(srodekRow);
                }
                string srodek = srodekBuilder.ToString().TrimEnd(',');
                srodek = srodek.Substring(0, srodek.Length - 3);
                string insertedStart = GenerateStartString(grupyKontrah);
                string finalJson = start + Environment.NewLine + srodek + Environment.NewLine + stringKoniec;
                System.IO.File.WriteAllText(jsonOutputPath, finalJson);
                MessageBox.Show("Wyeksportowano pomyślnie!", "Eksport", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private string GenerateStartString(string grupyKontrah)
        {
            var o = new JObject {
                ["ZakresWylaczenGrupKontrah"] = 0,
                ["SposLaczUmowyZRabat"]     = 0,
                ["ZakresKontrah"]           = 0,
                ["SposLaczPromZUmonNaCene"] = 2,
                ["Typ"]                     = 0,
                ["ListaWylaczenKontrah"]    = new JArray(
                    new JObject(
                        new JProperty("Indeks", "SAGA SP. Z O. O.")
                    )
                ),
                ["ListaGrupKontrah"]        = JArray.FromObject(grupyKontrah),
                ["ZakresKarotek"]           = 1,
                ["ZaleznaOd"]               = 0,
                ["ListaKartotek"]           = new JArray(/* … */)
            };

            string json = JsonConvert.SerializeObject(o, Formatting.None);
            return json;
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

        private void label6_Click(object sender, EventArgs e)
        {
            throw new System.NotImplementedException();
        }
    }
}
