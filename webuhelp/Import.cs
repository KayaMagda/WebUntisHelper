using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.Data;
using System.IO;

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

            if (m_dbConnection == null || m_dbConnection.ConnectionString != $"Data Source={DBFile};Version=3;")
            {
                m_dbConnection = new SQLiteConnection($"Data Source={DBFile};Version=3;");
            }

            m_dbConnection.Open();

            if (IsNewDataBase)
            {
                sql = "CREATE TABLE Pupil (name VARCHAR(15), vorname VARCHAR(20), schuelerID INTEGER, klasse VARCHAR(10)); CREATE TABLE Absence (status VARCHAR(30), datum DATETIME, stundennr INTEGER, lehrkraft VARCHAR(4), fach VARCHAR(4), fehlminute INTEGER, grund VARCHAR(50), entschuldigungstext VARCHAR(20), text VARCHAR(50) )";

                command = new SQLiteCommand(sql, m_dbConnection);
                command.ExecuteNonQuery();
            }
        }

    }
}
