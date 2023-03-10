using System;
using System.Collections.Generic;
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

                            Console.WriteLine($"{Environment.NewLine}Annika Schäfer - Kaya Koop - Marika Lübbers {Environment.NewLine}");                            

                            break;

                        case "-i":

                            Import Import = new Import();
                            Import.CreateDB();
                            //TODO: Import aus Vezeichnis von Excel Dateien in die Datenbank

                            break;

                        case "-e":
                            if (!existingData)
                            {
                                NoData();
                            }
                            break;

                        case "-s":
                            if (!existingData)
                            {
                                NoData();
                            }
                            break;

                        default:
                            WriteLegalCommands();
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

        private static void ExportClassSummary()
        {
            try
            {
                var activeSheet = GetActiveWorksheet();
                if (activeSheet == null)
                {
                    Console.WriteLine("Excel ist nicht korrekt installiert, es kann kein Excel Export durchgeführt werden.");
                    return;
                }
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

                activeSheet.SaveAs(fileName);
                Console.WriteLine($"Die Datei {fileName} wurde erfolgreich erstellt.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Es gab einen Fehler beim Exportieren der Zusammenfassung, bitte versuchen Sie es nochmal.");
            }

        }

        private static _Worksheet? GetActiveWorksheet()
        {
            var excelApp = new Application();
            if (excelApp == null)
            {
                return null;
            }
            excelApp.Visible = true;
            excelApp.Workbooks.Add();
            _Worksheet workSheet = (Worksheet)excelApp.ActiveSheet;
            return workSheet;
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
