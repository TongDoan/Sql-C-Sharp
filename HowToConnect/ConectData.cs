using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HowToConnect
{
    internal class ConectData
    {
        SqlConnection sqlConnection;
        public void Connect()
        {
            sqlConnection = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=D:\\MenuObject\\LTTQ\\HowToConnect\\HowToConnect\\Database1.mdf;Integrated Security=True");
            if (sqlConnection.State != ConnectionState.Open)
            {
                sqlConnection.Open();
            }

        }
        public void closeConnect()
        {

            if (sqlConnection.State != ConnectionState.Closed)
            {
                sqlConnection.Close();
            }
        }
        public DataTable table(string query)
        {
            DataTable table = new DataTable();
            Connect();
            SqlDataAdapter adapter = new SqlDataAdapter(query, sqlConnection);
            adapter.Fill(table);
            closeConnect();
            table.Dispose();
            return table;
        }
        public void excute(string query)
        {
            Connect();
            SqlCommand sqlCommand = new SqlCommand(query, sqlConnection);
            sqlCommand.ExecuteNonQuery();
            closeConnect();
        }
    }
}
