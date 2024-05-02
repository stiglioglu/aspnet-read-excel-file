using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Web;
using System.Web.Mvc;
using OfficeOpenXml;

namespace Proje.Models
{
    public class DatabaseLogicLayer
    {
        private static DatabaseLogicLayer instance;
        private static readonly object lockObject = new object();

        private string connectionString;
        private string tableName;

        private DatabaseLogicLayer()
        {
            connectionString = 
                ConfigurationManager.ConnectionStrings["MyConnectionString"].ConnectionString;
            tableName = "";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public static DatabaseLogicLayer Instance
        {
            get
            {
                lock (lockObject)
                {
                    if (instance == null)
                    {
                        instance = new DatabaseLogicLayer();
                    }
                    return instance;
                }
            }
        }

        public List<string> GetTableNames()
        {
            List<string> tableNames = new List<string>();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                string query = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'";
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string tableName = reader["TABLE_NAME"].ToString();
                            tableNames.Add(tableName);
                        }
                    }
                }

                connection.Close();
            }

            return tableNames;
        }

        public int GetTableCount()
        {
            int tableCount = 0;

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                string query = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'";
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    tableCount = (int)command.ExecuteScalar();
                }

                connection.Close();
            }

            return tableCount;
        }

        public bool IsTableExists(string tableName)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                string query = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = @TableName";

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@TableName", tableName);

                    int count = (int)command.ExecuteScalar();

                    return count > 0;
                }
            }
        }

        public void CreateTableWithColumnCount(int columnCount)
        {
            tableName = "MyTable" + GetTableCount().ToString();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                StringBuilder createTableQuery = new StringBuilder();
                createTableQuery.Append("CREATE TABLE "+ tableName + "(");
                createTableQuery.Append("ID INT PRIMARY KEY IDENTITY,");

                for (int i = 0; i < columnCount; i++)
                {
                    createTableQuery.Append($"Column{i + 1} NVARCHAR(100),");
                }

                createTableQuery.Length--;

                createTableQuery.Append(")");

                using (SqlCommand command = new SqlCommand(createTableQuery.ToString(), connection))
                {
                    command.ExecuteNonQuery();
                    Console.WriteLine("Tablo oluşturuldu.");
                }

                connection.Close();
            }
        }

        public void InsertDataToDatabase(List<String> value)
        {
            int columnCount = value.Count();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                StringBuilder insertQuery = new StringBuilder($"INSERT INTO {tableName} (");

                for (int i = 1; i <= columnCount; i++)
                {
                    insertQuery.Append($"Column{i}");

                    if (i < columnCount)
                    {
                        insertQuery.Append(", ");
                    }
                }

                insertQuery.Append(") VALUES (");

                for (int i = 1; i <= columnCount; i++)
                {
                    insertQuery.Append($"@Value{i}");

                    if (i < columnCount)
                    {
                        insertQuery.Append(", ");
                    }
                }

                insertQuery.Append(")");

                using (SqlCommand command = new SqlCommand(insertQuery.ToString(), connection))
                {
                    for (int i = 1; i <= columnCount; i++)
                    {
                        command.Parameters.AddWithValue($"@Value{i}", value[i-1]);
                    }

                    command.ExecuteNonQuery();
                }

                connection.Close();
            }
        }

        public List<List<string>> GetTableData(string tableName)
        {
            if (string.IsNullOrEmpty(tableName))
            {
                return null;
            }
            if (!IsTableExists(tableName))
            {
                return null;
            }
            List<List<string>> tableData = new List<List<string>>();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                string query = $"SELECT * FROM {tableName}";

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            List<string> rowData = new List<string>();
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                if (!reader.IsDBNull(i))
                                {
                                    object value = reader.GetValue(i);
                                    if (value.GetType() == typeof(string))
                                    {
                                        string data = (string)value;
                                        rowData.Add(data);
                                    }
                                    else if (value.GetType() == typeof(int))
                                    {
                                        int intValue = (int)value;
                                        string data = intValue.ToString();
                                        rowData.Add(data);
                                    }
                                }
                            }
                            tableData.Add(rowData);
                        }
                    }
                }

                connection.Close();
            }

            return tableData;
        }

        public List<string> GetItemById(string tableName, int id)
        {
            List<string> itemData = new List<string>();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                string query = $"SELECT * FROM {tableName} WHERE ID = @Id";

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Id", id);

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                if (!reader.IsDBNull(i))
                                {
                                    object value = reader.GetValue(i);
                                    if (value.GetType() == typeof(string))
                                    {
                                        string data = (string)value;
                                        itemData.Add(data);
                                    }
                                    else if (value.GetType() == typeof(int))
                                    {
                                        int intValue = (int)value;
                                        string data = intValue.ToString();
                                        itemData.Add(data);
                                    }
                                }
                            }
                        }
                    }
                }

                connection.Close();
            }

            return itemData;
        }

        public void UpdateItemById(string tableName, int id, List<string> newData)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                StringBuilder updateQuery = new StringBuilder($"UPDATE {tableName} SET ");

                for (int i = 0; i < newData.Count; i++)
                {
                    updateQuery.Append($"Column{i + 1} = @Value{i + 1}");

                    if (i < newData.Count - 1)
                    {
                        updateQuery.Append(", ");
                    }
                }

                updateQuery.Append(" WHERE ID = @Id");

                using (SqlCommand command = new SqlCommand(updateQuery.ToString(), connection))
                {
                    for (int i = 0; i < newData.Count; i++)
                    {
                        command.Parameters.AddWithValue($"@Value{i + 1}", newData[i]);
                    }
                    command.Parameters.AddWithValue("@Id", id);

                    command.ExecuteNonQuery();
                }

                connection.Close();
            }
        }

        public void DeleteItemById(string tableName, int id)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                string deleteQuery = $"DELETE FROM {tableName} WHERE ID = @Id";

                using (SqlCommand command = new SqlCommand(deleteQuery, connection))
                {
                    command.Parameters.AddWithValue("@Id", id);

                    command.ExecuteNonQuery();
                }

                connection.Close();
            }
        }

        public FileContentResult DownloadTableDataAsExcel(string tableName)
        {
            List<List<string>> tableData = GetTableData(tableName);

            using (var package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(tableName);


                for (int i = 0; i < tableData.Count; i++)
                {
                    for (int j = 1; j < tableData[i].Count; j++)
                    {
                        worksheet.Cells[i + 1, j ].Value = tableData[i][j];
                    }
                }

                byte[] fileBytes = package.GetAsByteArray();

                string fileName = tableName + ".xlsx";

                return new FileContentResult(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                {
                    FileDownloadName = fileName
                };
            }
        }








    }

}