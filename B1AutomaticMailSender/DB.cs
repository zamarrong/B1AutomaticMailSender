using Sap.Data.Hana;
using System;
using System.Data;
using System.Data.SqlClient;

namespace B1AutomaticMailSender
{
    public class DB
    {
        private string _connectionString { get; set; }

        public DB()
        {
            _connectionString = (Properties.Settings.Default.HANA) ? SAP.hana.ConnectionString : SAP.sql.ConnectionString;
        }

        public DataTable ExecuteQuery(string commandText)
        {
            var dataTable = new DataTable();
            var _sqlConnection = new SqlConnection(_connectionString);
            var cmd = new SqlCommand(commandText, _sqlConnection);
            var da = new SqlDataAdapter(cmd);

            try
            {
                _sqlConnection.Open();
                da.Fill(dataTable);
            }
            catch (Exception ex)
            {
                var errorText = string.Format("Error : Command={0} :: Error={1}", commandText, ex.Message);
                Program.WriteLog(errorText);
            }
            finally
            {
                da.Dispose();
                _sqlConnection.Dispose();
            }

            return dataTable;
        }
        public DataTable ExecuteQueryHana(string commandText)
        {
            var dataTable = new DataTable();
            var _sqlConnection = new HanaConnection(_connectionString);
            var cmd = new HanaCommand(commandText, _sqlConnection);
            var da = new HanaDataAdapter(cmd);

            try
            {
                _sqlConnection.Open();
                da.Fill(dataTable);
            }
            catch (Exception ex)
            {
                var errorText = string.Format("Error : Command={0} :: Error={1}", commandText, ex.Message);
                Program.WriteLog(errorText);
            }
            finally
            {
                da.Dispose();
                _sqlConnection.Dispose();
            }

            return dataTable;
        }
        public string ExecuteScalarString(string commandText)
        {
            var _sqlConnection = new SqlConnection(_connectionString);
            string row = "";

            try
            {
                var cmd = new SqlCommand(commandText, _sqlConnection);
                _sqlConnection.Open();
                row = cmd.ExecuteScalar().ToString();
            }
            catch (Exception ex)
            {
                var errorText = string.Format("Error : Command={0} :: Error={1}", commandText, ex.Message);
                Program.WriteLog(errorText);
            }
            finally
            {
                _sqlConnection.Dispose();
            }

            return row.ToString();
        }

        public string ExecuteScalarStringHana(string commandText)
        {
            var _sqlConnection = new HanaConnection(_connectionString);
            string row = "";

            try
            {
                var cmd = new HanaCommand(commandText, _sqlConnection);
                _sqlConnection.Open();
                row = cmd.ExecuteScalar().ToString();
            }
            catch (Exception ex)
            {
                var errorText = string.Format("Error : Command={0} :: Error={1}", commandText, ex.Message);
                Program.WriteLog(errorText);
            }
            finally
            {
                _sqlConnection.Dispose();
            }

            return row.ToString();
        }
    }
}
