using System;

using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace Grafics.NET
{
    internal class DB
    {
        // Підключення до БД

        MySqlConnection connection = new MySqlConnection("server=localhost;port=3306;username=root;password=Jeka-Nikita228;database=abonents");
        public void OpenConnection()
        {
            try
            {
                if (connection.State == System.Data.ConnectionState.Closed)
                    connection.Open();
            }
            catch (Exception ex)
            { 
                MessageBox.Show(ex.Message);
            }

        }

        // Закриття з'єднання

        public void CloseConnection()
        {
            if (connection.State == System.Data.ConnectionState.Open)
                connection.Close();
        }

        // Отримання з'єднання

        public MySqlConnection GetConnection()
        {
            return connection;
        }

    }
}
