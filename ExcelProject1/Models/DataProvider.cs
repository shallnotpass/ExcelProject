using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelProject1.Models
{
    class DataProvider
    {
        private string ConnectionString = Properties.Settings.Default.BoilerValuesConnectionString.ToString();
        public void SaveDataToDatabase(List<ValueDBO> values, List<TimedValueDBO> timedValues)
        {
            string output = "";
            using (SqlConnection connection = new SqlConnection(this.ConnectionString))
            {
                using (var command = new SqlCommand(output, connection))
                {
                    connection.Open();
                    command.CommandText = @"IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='BoilerVls')
                        CREATE TABLE BoilerVls
                        (idvalues int IDENTITY(1,1) PRIMARY KEY,  
                        tagname char(20) NOT NULL,  
                        value_type char(20) NOT NULL,  
                        boiler_value float(12) NOT NULL,);";
                    command.ExecuteNonQuery();
                    command.CommandText = @"IF NOT EXISTS(SELECT * FROM sysobjects WHERE name = 'TimedBoilerVls')
                        CREATE TABLE TimedBoilerVls
                        (idtimedvalues int IDENTITY(1,1) PRIMARY KEY,
                        time_date char(30) NOT NULL,
                        tagname char(20) NOT NULL,
                        value_type char(20) NOT NULL,  
                        boiler_value float(12) NOT NULL,);";
                    command.ExecuteNonQuery();
                    for (int i = 1; i < values.Count; i++)
                    {
                        output = String.Format("INSERT INTO dbo.BoilerVls ( tagname, value_type, boiler_value) VALUES " +
                            "('{0}', " +
                            "'{1}', " +
                            "'{2}');",
                            values[i].TagName.ToString(), 
                            values[i].Type.ToString(), 
                            values[i].Value.ToString().Replace(',', '.'));
                        command.CommandText = output;
                        command.ExecuteNonQuery();
                    }
                    for (int i = 1; i < timedValues.Count; i++)
                    {
                        output = String.Format("INSERT INTO timedBoilerVls (time_date, tagname, value_type, boiler_value) VALUES " +
                            "('{0}', " +
                            "'{1}', " +
                            "'{2}'," +
                            "'{3}');",
                            timedValues[i].DateTime.ToString(),
                            timedValues[i].TagName.ToString(),
                            timedValues[i].Type.ToString(),
                            timedValues[i].Value.ToString().Replace(',', '.'));
                        command.CommandText = output;
                        int result = command.ExecuteNonQuery();
                    }
                }
            }
        }
        public async Task<List<ValueDBO>> ReadValuesFromDatabase()
        {
            List<ValueDBO> values = new List<ValueDBO>();
            string readCommand = @"SELECT * FROM dbo.BoilerVls;";
            using (SqlConnection connection = new SqlConnection(this.ConnectionString))
            {
                using (SqlCommand command = new SqlCommand(readCommand, connection))
                {
                    connection.Open();
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            object[] row = new object[4];
                            reader.GetValues(row);
                            values.Add(new ValueDBO { TagName = (string)row[1], Type = (string)row[2], Value = row[3].ToString() });
                        }
                        return values;
                    }
                }
            }
        }
        public async Task<List<TimedValueDBO>> ReadTimedValuesFromDatabase()
        {
            List<TimedValueDBO> values = new List<TimedValueDBO>();
            string readCommand = @"SELECT * FROM dbo.TimedBoilerVls;";
            using (SqlConnection connection = new SqlConnection(this.ConnectionString))
            {
                using (SqlCommand command = new SqlCommand(readCommand, connection))
                {
                    connection.Open();
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            object[] row = new object[5];
                            reader.GetValues(row);
                            values.Add(new TimedValueDBO {DateTime = (string)row[1], TagName = (string)row[2], Type = (string)row[3], Value = row[4].ToString() });
                        }
                        return values;
                    }
                }
            }
        }
        public List<ValueDBO> sdf()
        {
            List<ValueDBO> values = new List<ValueDBO>(); 
            string readCommand = @"SELECT * FROM dbo.BoilerVls;";
            using (SqlConnection connection = new SqlConnection(this.ConnectionString))
            {
                using (SqlCommand command = new SqlCommand(readCommand, connection))
                {
                    connection.Open();
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            object[] row = new object[4];
                            reader.GetValues(row);
                            values.Add(new ValueDBO { TagName = (string)row[1], Type = (string)row[2], Value = row[3].ToString() });
                        }
                        return values;
                    }
                }
            }
        }
    }
}
