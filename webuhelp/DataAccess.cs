using System.Collections.Generic;
using System.Data.SQLite;

namespace webuhelp
{
    internal class DataAccess
    {
        private static SQLiteConnection GetOpenConnection()
        {
            var connection = new SQLiteConnection($"Data Source=WebUntisDB.db;Version=3;");
            connection.Open();
            return connection;
        }

        public List<PupilRow> GetSummaryData()
        {
            // Status = entsch., nicht entsch. Grund = Verspätet Entschuldigungstext = B
            var query = @"SELECT
                            P.schuelerID,
                            P.klasse,
                            P.name,
                            P.vorname,
                            (SELECT SUM(*) 
                                FROM Absence A 
                                WHERE A.schuelerID = P.schuelerID
                                AND status = 'entsch.';";
        }

    }
}
