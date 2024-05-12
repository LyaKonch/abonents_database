using System;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace Grafics.NET
{
    public partial class Login : Form
    {
        private DB db = new DB();

        public Login()
        {
            InitializeComponent();
        }

        private void btnSignIn_Click(object sender, EventArgs e)
        {
            // Get the username and password from the input fields
            string username = txtUsername.Text;
            string password = txtPassword.Text;

            try
            {
                // Open the database connection
                db.OpenConnection();

                // Prepare the SQL query to retrieve the user's credentials
                string query = "SELECT * FROM editor_info WHERE Login = @login AND Password = @password";
                MySqlCommand command = new MySqlCommand(query, db.GetConnection());
                command.Parameters.AddWithValue("@login", username);
                command.Parameters.AddWithValue("@password", password);

                // Execute the query
                MySqlDataReader reader = command.ExecuteReader();

                // Check if any rows were returned
                if (reader.HasRows)
                {
                    // Credentials are valid, close the login form and return DialogResult.OK
                    MessageBox.Show("Successful login!");
                    this.DialogResult = DialogResult.OK;
                }
                else
                {
                    // No matching user found, display error message
                    MessageBox.Show("Invalid username or password!");
                }
            }
            catch (Exception ex)
            {
                // Handle any exceptions
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                // Close the database connection
                db.CloseConnection();
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            // Close the login form and return DialogResult.Cancel
            this.DialogResult = DialogResult.Cancel;
        }
    }
}
