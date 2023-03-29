using System.Data.SQLite;
using System.IO;
using System;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;


namespace webuhelp
{
    class Import
    {

        public string DBFile = "WebUntisDB.db";
        public bool IsNewDataBase = false;
        public string sql;

        SQLiteCommand command;
        SQLiteConnection m_dbConnection;


        public void CreateDB()
        {
            if (!File.Exists(DBFile))
            {
                SQLiteConnection.CreateFile(DBFile);
                IsNewDataBase = true;
            }
            else
            {
                IsNewDataBase = false;
            }

            if (m_dbConnection == null || m_dbConnection.ConnectionString != $"Data Source={DBFile};Version=3;")
            {
                m_dbConnection = new SQLiteConnection($"Data Source={DBFile};Version=3;");
            }

            m_dbConnection.Open();

            if (IsNewDataBase)
            {
                sql = @"CREATE TABLE Pupil
                        (
                        name VARCHAR(15), 
                        vorname VARCHAR(20), 
                        schuelerID INTEGER PRIMARY KEY, 
                        klasse VARCHAR(10)
                        );

                        CREATE TABLE Absence
                        (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        status VARCHAR(30), 
                        datum DATETIME,
                        stundennr INTEGER,
                        lehrkraft VARCHAR(4),
                        fach VARCHAR(4),
                        fehlminute INTEGER,
                        grund VARCHAR(50),
                        entschuldigungstext VARCHAR(20),
                        text VARCHAR(50),
                        schuelerID INTEGER,
                        FOREIGN KEY (schuelerID) 
                        REFERENCES Pupil (schuelerID) 
                        );";

                command = new SQLiteCommand(sql, m_dbConnection);
                command.ExecuteNonQuery();
            }
            else
            {           //Hier Drop verwendet, da der ID Count bei Absences sonst nicht richtig zählt
                sql = @"DROP TABLE Pupil;
                        DROP TABLE Absence;
                        CREATE TABLE Pupil
                        (
                        name VARCHAR(15), 
                        vorname VARCHAR(20), 
                        schuelerID INTEGER PRIMARY KEY, 
                        klasse VARCHAR(10)
                        );

                        CREATE TABLE Absence
                        (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        status VARCHAR(30), 
                        datum DATETIME,
                        stundennr INTEGER,
                        lehrkraft VARCHAR(4),
                        fach VARCHAR(4),
                        fehlminute INTEGER,
                        grund VARCHAR(50),
                        entschuldigungstext VARCHAR(20),
                        text VARCHAR(50),
                        schuelerID INTEGER,
                        FOREIGN KEY (schuelerID) 
                        REFERENCES Pupil (schuelerID) 
                        );
                        ";

                command = new SQLiteCommand(sql, m_dbConnection);
                command.ExecuteNonQuery();
            }

            m_dbConnection.Close();
        }


        public void ExcelImport()
        {

            String Programpath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string[] ExcelFiles = Directory.GetFiles(Programpath, "*.xls*");            

            if (ExcelFiles.Length <= 0)
            {
                throw new Exception($"Es konnte keine Excel-Datei im Pogrammpfad ({Programpath}) gefunden werden.");
            }
            else
            {
                foreach (string file in ExcelFiles)
                {
                    Excel.Application xlApp;
                    Excel.Workbook xlWorkbook;
                    Excel.Worksheet xlWorkSheet;
                    Excel.Range range;

                    int count = 0;
                    xlApp = new Excel.Application();
                    xlWorkbook = xlApp.Workbooks.Open(file);
                    xlWorkSheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);
                    range = xlWorkSheet.UsedRange;
                    m_dbConnection.Open();

                    for (count = 2; count <= range.Rows.Count; count++)
                    {
                        string FullName = (string)Convert.ToString((range.Cells[count, 1] as Excel.Range).Value2);
                        string[] NameKomplett;
                        NameKomplett = FullName.Split(' ');
                        string name = NameKomplett[0];
                        string vorname = NameKomplett[1];

                        string SchuelerID = (string)Convert.ToString((range.Cells[count, 2] as Excel.Range).Value2);                        
                        string Klasse = (string)Convert.ToString((range.Cells[count, 4] as Excel.Range).Value2);
                        string status = (string)Convert.ToString((range.Cells[count, 18] as Excel.Range).Value2);

                        string date = (string)Convert.ToString((range.Cells[count, 5] as Excel.Range).Value);
                        string[] datumkomplett;
                        datumkomplett = date.Split('.');                        
                        string Tag = datumkomplett[0];
                        string Monat = datumkomplett[1];
                        string JahrUF = datumkomplett[2];
                        string[] JahrAr = JahrUF.Split(' ');
                        string Jahr = JahrAr[0];
                        string datum = $"{Jahr}-{Monat}-{Tag}";

                        string stundennr = (string)Convert.ToString((range.Cells[count, 17] as Excel.Range).Value2);
                        string lehrkraft = (string)Convert.ToString((range.Cells[count, 9] as Excel.Range).Value2);
                        string fach = (string)Convert.ToString((range.Cells[count, 10] as Excel.Range).Value2);
                        string fehlminute = (string)Convert.ToString((range.Cells[count, 8] as Excel.Range).Value2);
                        string grund = (string)Convert.ToString((range.Cells[count, 11] as Excel.Range).Value2);
                        string entschtext = (string)Convert.ToString((range.Cells[count, 16] as Excel.Range).Value2);
                        string text = (string)Convert.ToString((range.Cells[count, 12] as Excel.Range).Value2);

                        
                        if (range.Cells[count, 1].value2 != range.Cells[count + 1, 1].value2)
                        {
                            sql = "INSERT OR IGNORE INTO Pupil(name, vorname, schuelerID, klasse) VALUES('" + name + "','" + vorname + "'," + SchuelerID + ",'" + Klasse + "')";

                            command = new SQLiteCommand(sql, m_dbConnection);
                            command.ExecuteNonQuery();

                        }
                        
                        

                        sql = @"INSERT OR IGNORE INTO Absence(status, datum, stundennr, lehrkraft, fach, fehlminute, grund, entschuldigungstext, text, schuelerID)
                                VALUES('" + status + "','" + datum + "'," + stundennr + ",'" + lehrkraft + "','" + fach + "'," + fehlminute + ",'" + grund + "','" + entschtext + "','" + text + "'," + SchuelerID + ")";

                        command = new SQLiteCommand(sql, m_dbConnection);
                        command.ExecuteNonQuery();

                    }
                    xlWorkbook.Close(true, null, null);
                    xlApp.Quit();
                    Console.WriteLine($"Die Datei {Path.GetFileName(file)} wurde erfolgreich importiert.");
                    m_dbConnection.Close();
                }

            }

            
        }
    }
}

