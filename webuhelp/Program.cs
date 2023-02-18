using System;
using System.Collections.Generic;

namespace webuhelp
{
    internal class Program
    {
        private static List<string> _developers = new List<string>() { "Marika Lübbers", "Annika Schäfer", "Kaya Kopp"};
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
                            break;
                        case "-i":
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
