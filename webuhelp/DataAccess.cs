using System.Collections.Generic;
using System.Data.SQLite;

namespace webuhelp
{
    public class DataAccess
    {
        private static SQLiteConnection GetOpenConnection()
        {
            var connection = new SQLiteConnection($"Data Source=WebUntisDB.db;Version=3;");
            connection.Open();
            return connection;
        }

        private static int? GetIntNullable(SQLiteDataReader reader, int index)
        {
            if (reader.IsDBNull(index))
            {
                return null;
            }
            return reader.GetInt32(index);
        }

        public static List<PupilRow> GetSummaryData()
        {
            var query = @"SELECT
                            P.name,
                            P.vorname,
                            (
                                SELECT
                                COUNT(*) 
                                WHERE A.status = 'entsch.'
                             ) AS excused,
                             (
                                SELECT
                                COUNT(*)
                                WHERE A.status = 'nicht entsch.'
                             ) AS not_excused,
                             (
                                SELECT
                                COUNT(*)
                                WHERE A.entschuldigungstext = 'B'
                             ) AS work_related,
                             (
                                SELECT
                                COUNT(*)
                                WHERE A.grund = 'Verspätet'
                             ) AS late
                            FROM Pupil P
                            JOIN Absence A ON A.schuelerID = P.schuelerID                                
                            ;";

            var rowList = new List<PupilRow>();

            using (var connection = GetOpenConnection())
            {
                using (var command = connection.CreateCommand())
                {
                    command.CommandText = query;
                    
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var row = new PupilRow();

                            row.Name = reader.GetString(0);
                            row.FirstName= reader.GetString(1);
                            row.A = GetIntNullable(reader, 2) ?? 0;
                            row.N = GetIntNullable(reader, 3) ?? 0;
                            row.B = GetIntNullable(reader, 4) ?? 0;
                            row.V = GetIntNullable(reader, 5) ?? 0;

                            rowList.Add(row);
                        }
                    }
                }
            }
            return rowList;
        }

        public static string GetClassName()
        {
            var query = @"SELECT
                            klasse
                          FROM Pupil LIMIT 1;";

            using (var connection = GetOpenConnection())
            {
                using (var command = connection.CreateCommand())
                {
                    command.CommandText = query;
                    return (string) command.ExecuteScalar();                   
                }
            }
        }
    }
}
