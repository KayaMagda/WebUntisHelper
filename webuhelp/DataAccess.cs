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

        private static string? GetStringNullable(SQLiteDataReader reader, int index)
        {
            if (reader.IsDBNull(index))
            {
                return null;
            }
            return reader.GetString(index);
        }

        public static List<Pupil> GetPersonData()
        {
            var query = @"SELECT
                            P.schuelerID,
                            P.name,
                            P.vorname,
                            P.klasse,
                            (
                                SELECT
                                COUNT(*)
                                FROM Absence A
                                WHERE A.status = 'entsch.' AND A.schuelerID = P.schuelerID
                             ) AS excused,
                             (
                                SELECT
                                COUNT(*)
                                FROM Absence A
                                WHERE A.status = 'nicht entsch.' AND A.schuelerID = P.schuelerID
                             ) AS not_excused,
                             (
                                SELECT
                                COUNT(*)
                                FROM Absence A
                                WHERE A.entschuldigungstext = 'B' AND A.schuelerID = P.schuelerID
                             ) AS work_related,
                             (
                                SELECT
                                COUNT(*)
                                FROM Absence A
                                WHERE A.grund = 'Verspätet' AND A.schuelerID = P.schuelerID
                             ) AS late
                            FROM Pupil P
                            ;";

            var rowList = new List<Pupil>();

            using var connection = GetOpenConnection();

            using var command = connection.CreateCommand();
            command.CommandText = query;

            using var reader = command.ExecuteReader();

            while (reader.Read())
            {
                var row = new Pupil();

                row.ID = reader.GetInt32(0);
                row.LastName = reader.GetString(1);
                row.FirstName = reader.GetString(2);
                row.Class = reader.GetString(3);
                row.Data = new List<PupilData> ();
                row.A = GetIntNullable(reader, 4) ?? 0;
                row.N = GetIntNullable(reader, 5) ?? 0;
                row.B = GetIntNullable(reader, 6) ?? 0;
                row.V = GetIntNullable(reader, 7) ?? 0;

                rowList.Add(row);
            }
            rowList = GetPupilData(rowList);

            return rowList;
        }

        private static List<Pupil> GetPupilData(List<Pupil> rowList)
        {
            for (int i = 0; i < rowList.Count; i++)
            {
                var query = $@"SELECT
                            A.status,
                            A.datum,
                            A.stundennr,
                            A.lehrkraft,
                            A.fach,
                            A.fehlminute,
                            A.grund,
                            A.entschuldigungstext,
                            A.text
                            FROM Pupil P
                            JOIN Absence A ON A.schuelerID = P.schuelerID
                            WHERE P.schuelerID = '{rowList[i].ID}'
                            ;";

                using var connection = GetOpenConnection();

                using var command = connection.CreateCommand();
                command.CommandText = query;

                using var reader = command.ExecuteReader();

                while (reader.Read())
                {
                    var row = new PupilData();

                    row.IsExcused = (reader.GetString(0).ToLower().Trim() == "entsch.") ? true : false;
                    row.Date = "";//reader.GetDateTime(1).ToString("dd.MM.yyyy");
                    row.Weekday = ""; //reader.GetDateTime(1).ToString("ddd").Substring(0, 2) + ".";
                    row.LessonNr = reader.GetInt32(2);
                    row.Teacher = reader.GetString(3);
                    row.Lesson = reader.GetString(4);
                    row.MissingHour = (45 - reader.GetInt32(5) == 0) ? 1 : 0;
                    row.MissingMinute = reader.GetInt32(5); 
                    row.Reason = reader.GetString(6);
                    row.MissingText = GetStringNullable(reader, 7);
                    row.Text = GetStringNullable(reader, 8);

                    rowList[i].Data.Add(row);
                }
            }

            return rowList;
        }

        public static List<PupilRow> GetSummaryData()
        {
            var query = @"SELECT
                            P.name,
                            P.vorname,
                            (
                                SELECT
                                COUNT(*) 
                                FROM Absence A
                                WHERE A.status = 'entsch.' AND A.schuelerID = P.schuelerID
                             ) AS excused,
                             (
                                SELECT
                                COUNT(*)
                                FROM Absence A
                                WHERE A.status = 'nicht entsch.' AND A.schuelerID = P.schuelerID
                             ) AS not_excused,
                             (
                                SELECT
                                COUNT(*)
                                FROM Absence A
                                WHERE A.entschuldigungstext = 'B' AND A.schuelerID = P.schuelerID
                             ) AS work_related,
                             (
                                SELECT
                                COUNT(*)
                                FROM Absence A
                                WHERE A.grund = 'Verspätet' AND A.schuelerID = P.schuelerID
                             ) AS late
                            FROM Pupil P
                            ;";

            var rowList = new List<PupilRow>();

            using var connection = GetOpenConnection();

            using var command = connection.CreateCommand();                
                    command.CommandText = query;

            using var reader = command.ExecuteReader();
                    
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
            
            return rowList;
        }

        public static string GetClassName()
        {
            var query = @"SELECT
                            klasse
                          FROM Pupil LIMIT 1;";

            using var connection = GetOpenConnection();

            using var command = connection.CreateCommand();
                
            command.CommandText = query;
            return (string) command.ExecuteScalar();              
        }
    }
}
