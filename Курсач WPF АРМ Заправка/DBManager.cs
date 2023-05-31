using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace Курсач_WPF_АРМ_Заправка
{
    public class DBManager
    {
        SqlConnection sqlConnection = new SqlConnection(@"Data Source=DESKTOP-9714ODM\SQLEXPRESS;Initial Catalog=Заправка;Integrated Security=True");


        public void openConnection()
        {
            if(sqlConnection.State==System.Data.ConnectionState.Closed)
            {
                sqlConnection.Open();
            }
        }

        public void closeConnection()
        {
            if (sqlConnection.State == System.Data.ConnectionState.Open)
            {
                sqlConnection.Close();
            }
        }

        public SqlConnection getConnection()
        {
            return sqlConnection;
        }
    }

}
