#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace webuhelp
{
    internal class Program
    {
        private static Dictionary <string, string> _legalCommands = new Dictionary<string, string>()
        {
            ["-n"] = "Gibt die Namen aller Projektbeteiligten aus.",
            ["-i"] = "Automatisierter Import aller Gesamtdateien im aktuellen Verzeichnis",
            ["-e"] = "Automatisierter Export aller Einzeldateien",
            ["-s"] = "Automatisierter Export einer Zusammenfassung"
        };
        private static bool existingData = false;

        private static string currentDirectory = Directory.GetCurrentDirectory();

        static void Main(string[] args)
        {
            bool runApp = true;
            string command = "";

            while (runApp)
            { 
                if (command.Length == 0 && args.Length == 0 || args.Length > 1)
                {
                    WriteLegalCommands();
                }
                else
                {
                    command = args.Length != 0 ? args[0] : command;
                    switch (command)
                    {
                        case "-n":
                            {
                                Console.WriteLine($"{Environment.NewLine}Annika Schäfer - Kaya Kopp - Marika Lübbers {Environment.NewLine}");
                            }
                            break;

                        case "-i":
                            {
                                Import Import = new Import();
                                Import.CreateDB();
                                //TODO: Import aus Vezeichnis von Excel Dateien in die Datenbank
                                existingData = true;
                            }
                            break;

                        case "-e":
                            {
                                if (!existingData) NoData();
                                else ExportStudentSummary();
                            }
                            break;

                        case "-s":
                            {
                                if (!existingData) NoData();
                                else ExportClassSummary();
                            }
                            break;

                        default:
                            {
                                WriteLegalCommands();
                            }
                            break;
                    }
                }
                Console.WriteLine("Drücken sie ESC zum Schließen oder geben Sie einen Befehl ein.");
                ConsoleKeyInfo key = Console.ReadKey();
                if (key.Key == ConsoleKey.Escape)
                {
                    runApp = false;
                    break;
                }
                else
                {
                    command = "-";
                }
                command += Console.ReadLine();
            }
            
        }

        private static void ExportStudentSummary()
        {
            var rowList = DataAccess.GetPersonData();

            for (int i = 0; i < rowList.Count; i++)
            {
                try
                {
                    var excelApp = new Application();
                    if (excelApp == null)
                    {
                        Console.WriteLine("Excel ist nicht korrekt installiert, es kann kein Excel Export durchgeführt werden.");
                        return;
                    }

                    excelApp.Workbooks.Add();
                    _Worksheet activeSheet = (Worksheet)excelApp.ActiveSheet;
                    
                    activeSheet.Cells[1, "A"] = "name";
                    activeSheet.Cells[1, "B"] = "id";
                    activeSheet.Cells[1, "C"] = "klasse";
                    activeSheet.Cells[1, "D"] = "status";
                    activeSheet.Cells[1, "E"] = "datum";
                    activeSheet.Cells[1, "F"] = "wochentag";
                    activeSheet.Cells[1, "G"] = "stundennr";
                    activeSheet.Cells[1, "H"] = "lehrkraft";
                    activeSheet.Cells[1, "I"] = "fach";
                    activeSheet.Cells[1, "J"] = "fehlstunden";
                    activeSheet.Cells[1, "K"] = "fehlminuten";
                    activeSheet.Cells[1, "L"] = "grund";
                    activeSheet.Cells[1, "M"] = "entschuldigungstext";
                    activeSheet.Cells[1, "N"] = "text";

                    string personName = "";
                    var row = 1;
                    foreach (PupilData data in rowList[i].Data)
                    {
                        if (row <= rowList[i].Data.Count)
                        {
                            row++;

                            activeSheet.Cells[row, "A"] = rowList[i].GetFullNameGermanBurocratic();
                            activeSheet.Cells[row, "B"] = rowList[i].ID;
                            activeSheet.Cells[row, "C"] = rowList[i].Class;

                            if (data.IsExcused) activeSheet.Cells[row, "D"] = "entsch.";
                            else
                            {
                                activeSheet.Cells[row, "D"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red); 
                                activeSheet.Cells[row, "D"] = "nicht entsch.";
                            }

                            activeSheet.Cells[row, "E"] = data.Date;
                            activeSheet.Cells[row, "F"] = data.Weekday;
                            activeSheet.Cells[row, "G"] = data.LessonNr;
                            activeSheet.Cells[row, "H"] = data.Teacher;
                            activeSheet.Cells[row, "I"] = data.Lesson;
                            activeSheet.Cells[row, "J"] = data.MissingHour;
                            activeSheet.Cells[row, "K"] = data.MissingMinute;
                            activeSheet.Cells[row, "L"] = data.Reason;
                            activeSheet.Cells[row, "M"] = data.MissingText;
                            activeSheet.Cells[row, "N"] = data.Text;
                        }

                        if (row >= rowList[i].Data.Count)
                        {
                            personName = rowList[i].GetFileName();

                            activeSheet.Cells[row + 1, "A"] = "A";
                            activeSheet.Cells[row + 2, "A"] = "N";
                            activeSheet.Cells[row + 3, "A"] = "B";
                            activeSheet.Cells[row + 4, "A"] = "V";

                            activeSheet.Cells[row + 1, "B"] = rowList[i].A;
                            activeSheet.Cells[row + 2, "B"] = rowList[i].N;
                            activeSheet.Cells[row + 3, "B"] = rowList[i].B;
                            activeSheet.Cells[row + 4, "B"] = rowList[i].V;
                        }
                    }

                    activeSheet.Columns[1].AutoFit();
                    activeSheet.Columns[2].AutoFit();
                    activeSheet.Columns[3].AutoFit();
                    activeSheet.Columns[4].AutoFit();
                    activeSheet.Columns[5].AutoFit();
                    activeSheet.Columns[6].AutoFit();
                    activeSheet.Columns[7].AutoFit();
                    activeSheet.Columns[8].AutoFit();
                    activeSheet.Columns[9].AutoFit();
                    activeSheet.Columns[10].AutoFit();
                    activeSheet.Columns[11].AutoFit();
                    activeSheet.Columns[12].AutoFit();
                    activeSheet.Columns[13].AutoFit();
                    activeSheet.Columns[14].AutoFit();

                    var fileName = personName + ".xlsx";
                    var fullPath = currentDirectory + "\\" + fileName;

                    activeSheet.SaveAs(fullPath);
                    excelApp.Workbooks.Close();
                    excelApp.Quit();
                    Console.WriteLine($"Die Datei {fileName} wurde erfolgreich exportiert.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Es gab einen Fehler beim Exportieren der Einzeldateien, bitte versuchen Sie es nochmal.");
                } 
            }

        }
        private static void ExportClassSummary()
        {
            try
            {
                var excelApp = new Application();
                if (excelApp == null)
                {
                    Console.WriteLine("Excel ist nicht korrekt installiert, es kann kein Excel Export durchgeführt werden.");
                    return;
                }
                excelApp.Workbooks.Add();
                _Worksheet activeSheet = (Worksheet)excelApp.ActiveSheet;
               
                var rowList = DataAccess.GetSummaryData();

                activeSheet.Cells[1, "A"] = "Name";
                activeSheet.Cells[1, "B"] = "Vorname";
                activeSheet.Cells[1, "C"] = "A";
                activeSheet.Cells[1, "D"] = "N";
                activeSheet.Cells[1, "E"] = "B";
                activeSheet.Cells[1, "F"] = "V";

                var row = 1;
                foreach (PupilRow pupilRow in rowList)
                {
                    row++;
                    activeSheet.Cells[row, "A"] = pupilRow.Name;
                    activeSheet.Cells[row, "B"] = pupilRow.FirstName;
                    activeSheet.Cells[row, "C"] = pupilRow.A;
                    activeSheet.Cells[row, "D"] = pupilRow.N;
                    activeSheet.Cells[row, "E"] = pupilRow.B;
                    activeSheet.Cells[row, "F"] = pupilRow.V;
                }

                activeSheet.Columns[1].AutoFit();
                activeSheet.Columns[2].AutoFit();
                activeSheet.Columns[3].AutoFit();
                activeSheet.Columns[4].AutoFit();
                activeSheet.Columns[5].AutoFit();
                activeSheet.Columns[6].AutoFit();

                var className = DataAccess.GetClassName();
                var fileName = className + "_Zusammenfassung.xlsx";
                var fullPath = currentDirectory + "\\" + fileName;

                activeSheet.SaveAs(fullPath);
                excelApp.Workbooks.Close();
                excelApp.Quit();
                Console.WriteLine($"Die Datei {fileName} wurde erfolgreich erstellt.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Es gab einen Fehler beim Exportieren der Zusammenfassung, bitte versuchen Sie es nochmal.");
            }

        }
        

        private static void WriteLegalCommands()
        {
            Console.WriteLine("Sie haben keinen oder einen ungültigen Befehl eingegeben, bitte geben Sie einen der folgenden Befehle ein: ");

            foreach (KeyValuePair<string, string> entry in _legalCommands)
            {
                Console.WriteLine(entry.Key + ": " + entry.Value);
            }
        }

        private static void NoData()
        {
            Console.WriteLine("Die Datenbank ist leer, es kann nichts exportiert werden, über \"-i\" können Sie Daten importieren.");
        }
    }
}
