using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Data.OleDb;
using System.IO;
using System.Linq;

namespace DBConverter
{
    class AccessToSQLite
    {       
        static void Main()
        {
            // Prompt the user to install the driver
            if (!DriverInstalled())
            {
                Console.WriteLine("The Microsoft Access Driver is not installed.\n" +
                  "It can be downloaded from 'https://www.microsoft.com/en-au/download/confirmation.aspx?id=13255'.");
                return;
            }

            Console.WriteLine("");

            // Enumerate over files in the directory of the executable
            var files = Directory.EnumerateFiles(Environment.CurrentDirectory);
            foreach (var file in files)
            {
                if (Path.GetExtension(file) != ".accdb") continue;

                Console.WriteLine("Fetching data from " + Path.GetFileName(file));
                List<DataTable> tables = GetAccdbTables(file);

                Console.WriteLine("Writing data to SQLite . . .\n");
                CreateDB(Path.ChangeExtension(file, ".db"), tables);
            }                        
        }

        /// <summary>
        /// Reads table data from an Access database (.accdb) into a collection of
        /// System.Data DataTable objects
        /// </summary>
        /// <param name="path">Location of the Access database</param>
        public static List<DataTable> GetAccdbTables(string path)
        {
            List<DataTable> tables = new List<DataTable>();

            // Connect to the .accdb file
            string connector = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source = {path}";
            OleDbConnection connection = new OleDbConnection(connector);
            connection.Open();

            // Find all TABLE objects in the schema
            string[] restrictions = new string[4];
            restrictions[3] = "TABLE";
            DataTable schema = connection.GetSchema("Tables", restrictions);

            // Import each table from the schema
            foreach (DataRow row in schema.Rows)
            {
                // Create an empty DataTable
                string name = row.ItemArray[2].ToString();
                DataTable table = new DataTable(name);

                // Fill the table with data
                OleDbCommand command = new OleDbCommand($"SELECT * FROM {name}", connection);
                OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                adapter.Fill(table);

                tables.Add(table);
            }
            connection.Close();

            return tables;
        }

        /// <summary>
        /// Create an SQLite (.db) database from System.Data DataTables
        /// </summary>
        /// <param name="path">Full path for location of database</param>
        /// <param name="tables">Collection of tables to be included in database</param>
        public static void CreateDB(string path, List<DataTable> tables)
        {
            // Create DataBase and open connection to it
            SQLiteConnection.CreateFile(path);
            SQLiteConnection connection = new SQLiteConnection($"Data Source={path};Version=3;PRAGMA journal_mode=WAL");
            connection.Open();

            // Begin the transaction
            SQLiteCommand command = new SQLiteCommand("BEGIN", connection);
            command.ExecuteNonQuery();

            foreach (DataTable table in tables)
            {
                // Find the variable data for the table
                var x = from DataColumn col
                        in table.Columns
                        select $"{col.ColumnName} {col.DataType.Name.ToUpper().Replace("32", "")}";
                string variables = "(" + string.Join(", ", x) + ")";

                // Create the table
                command.CommandText = $"CREATE TABLE {table.TableName} {variables}";
                command.ExecuteNonQuery();

                // Find the variable names
                var cols = from DataColumn col
                           in table.Columns
                           select col.ColumnName;
                string names = "(" + string.Join(", ", cols) + ")";

                // Insert each row into the table
                foreach (DataRow row in table.Rows)
                {
                    string values = "(" + string.Join(", ", row.ItemArray.Select(item =>
                        string.IsNullOrEmpty(item.ToString()) ? "NULL" :
                            (item is string ? $"'{item.ToString().Replace("'", "''")}'" : item.ToString()))) + ")";
                    command.CommandText = $"insert into {table.TableName} {names} values {values};";
                    command.ExecuteNonQuery();
                }
            }

            // End the transaction and close the file
            command.CommandText = "END";
            command.ExecuteNonQuery();
            connection.Close();
        }

        /// <summary>
        /// Test if the required driver is installed
        /// </summary>
        public static bool DriverInstalled()
        {
            // Open the product registry
            RegistryKey Key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Classes\Installer\Products");
            if (Key == null) return false;
            
            // Look through each product in the registry
            foreach (string name in Key.GetSubKeyNames())
            {                   
                // Check if it is an instance of the Access database engine
                RegistryKey SubKey = Key.OpenSubKey(name);
                if ((SubKey.GetValue("ProductName") is string value) && value.Contains("Access database engine")) return true;                
            }            

            return false;
        }


    }
}
