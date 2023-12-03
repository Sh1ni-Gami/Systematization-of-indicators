using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using MySql.Data.MySqlClient;




namespace Programm
{
    internal class DataBase
    {
       // MySqlConnection sqlConnection = new MySqlConnection(@"user id=ruslan;password=Password1!;server=104.248.42.111;DATABASE=test;");

        public void openConnection()
        {
           /* if (sqlConnection.State == System.Data.ConnectionState.Closed)
            {
                sqlConnection.Open();
            } */
        }
        public void closeConnection()
        {
           /* if (sqlConnection.State == System.Data.ConnectionState.Open)
            {
                sqlConnection.Close();
            } */
        }
        /* public MySqlConnection getConnection()
        { 
            return sqlConnection; 
        }
        */
    }
}
