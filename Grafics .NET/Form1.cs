using Grafics.NET;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Coursework
{
    public partial class Form1 : Form
    {
        private DB db = new DB();
        private string placeholderText = "Type last name or year";
        private string userRole;
        private List<string[]> bufferList = new List<string[]>();
        private DataGridView bufferDataGridView;
        public Form1()
        {
            InitializeComponent();
            panel1.Visible = false;
            this.MouseDown += new MouseEventHandler(Form1_MouseDown);
            PopulateDataGridView();
        }

        private void InitializePanel1()
        {
            panel1.Visible = true;
            // Add labels
            Label labelLastName = new Label();
            labelLastName.Text = "Last Name:";
            labelLastName.Location = new System.Drawing.Point(20, 0);
            labelLastName.Size = new System.Drawing.Size(100, 20);

            Label labelPhoneNumber = new Label();
            labelPhoneNumber.Text = "Phone number:";
            labelPhoneNumber.Location = new System.Drawing.Point(20, 40);
            labelPhoneNumber.Size = new System.Drawing.Size(100, 20);

            Label labelAddress = new Label();
            labelAddress.Text = "Address:";
            labelAddress.Location = new System.Drawing.Point(20, 80);
            labelAddress.Size = new System.Drawing.Size(100, 20);

            Label labelYear = new Label();
            labelYear.Text = "Year:";
            labelYear.Location = new System.Drawing.Point(20, 120);
            labelYear.Size = new System.Drawing.Size(100, 20);

            // Add text boxes for Last Name, Phone number, Address, and Year
            TextBox textBoxLastName = new TextBox();
            textBoxLastName.ForeColor = System.Drawing.Color.Gray;
            textBoxLastName.Location = new System.Drawing.Point(20, 20);
            textBoxLastName.Size = new System.Drawing.Size(150, 20);

            TextBox textBoxPhoneNumber = new TextBox();
            textBoxPhoneNumber.ForeColor = System.Drawing.Color.Gray;
            textBoxPhoneNumber.Location = new System.Drawing.Point(20, 60);
            textBoxPhoneNumber.Size = new System.Drawing.Size(150, 20);

            TextBox textBoxAddress = new TextBox();
            textBoxAddress.ForeColor = System.Drawing.Color.Gray;
            textBoxAddress.Location = new System.Drawing.Point(20, 100);
            textBoxAddress.Size = new System.Drawing.Size(150, 20);

            TextBox textBoxYear = new TextBox();
            textBoxYear.ForeColor = System.Drawing.Color.Gray;
            textBoxYear.Location = new System.Drawing.Point(20, 140);
            textBoxYear.Size = new System.Drawing.Size(150, 20);

            // Add an "Add" button
            Button buttonAdd = new Button();
            buttonAdd.Text = "Add";
            buttonAdd.Location = new System.Drawing.Point(20, 170);
            buttonAdd.Size = new System.Drawing.Size(70, 30);
            buttonAdd.Click += ButtonAdd_Click;

            // Add a "Move to Main" button
            Button buttonMoveToMain = new Button();
            buttonMoveToMain.Text = "Move to DB";
            buttonMoveToMain.Location = new System.Drawing.Point(90, 170);
            buttonMoveToMain.Size = new System.Drawing.Size(80, 30);
            buttonMoveToMain.Click += ButtonMoveToMain_Click;

            // Add a DataGridView inside Panel1
            bufferDataGridView = new DataGridView();
            bufferDataGridView.Location = new System.Drawing.Point(200, 0);
            bufferDataGridView.Size = new System.Drawing.Size(450, 200);
            bufferDataGridView.AllowUserToAddRows = false;
            bufferDataGridView.Columns.Add("LastName", "Last Name");
            bufferDataGridView.Columns.Add("PhoneNumber", "Phone Number");
            bufferDataGridView.Columns.Add("Address", "Address");
            bufferDataGridView.Columns.Add("Year", "Year");
            bufferDataGridView.UserDeletingRow += bufferDataGridView_UserDeletingRow;
            bufferDataGridView.CellEndEdit += bufferDataGridView_CellEndEdit;


            // Add controls to Panel1
            panel1.Controls.Add(textBoxLastName);
            panel1.Controls.Add(textBoxPhoneNumber);
            panel1.Controls.Add(textBoxAddress);
            panel1.Controls.Add(textBoxYear);
            panel1.Controls.Add(labelLastName);
            panel1.Controls.Add(labelPhoneNumber);
            panel1.Controls.Add(labelAddress);
            panel1.Controls.Add(labelYear);
            panel1.Controls.Add(buttonAdd);
            panel1.Controls.Add(bufferDataGridView);
            panel1.Controls.Add(buttonMoveToMain);
        }

        private void ButtonMoveToMain_Click(object sender, EventArgs e)
        {
            if (bufferList.Count < 10)
            {
                MessageBox.Show("Кількість записів менше 10. Заповніть ще");
                return;
            }

            try
            {
                // Open the database connection
                db.OpenConnection();
                DataTable dataTable = (DataTable)dataGridView1.DataSource;
                // Add data from buffer list to the main DataGridView and to the database
                foreach (string[] abonent in bufferList)
                {
                    // Add each element of abonent array to the DataGridView
                    DataRow row = dataTable.NewRow();
                    row["name"] = abonent[0];
                    row["number"] = abonent[1];
                    row["address"] = abonent[2];
                    row["year"] = abonent[3];
                    dataTable.Rows.Add(row);

                    // Insert the data into the database
                    string query = "INSERT INTO abonent (name, number, address, year) VALUES (@LastName, @PhoneNumber, @Address, @Year)";
                    MySqlCommand command = new MySqlCommand(query, db.GetConnection());
                    command.Parameters.AddWithValue("@LastName", abonent[0]);
                    command.Parameters.AddWithValue("@PhoneNumber", abonent[1]);
                    command.Parameters.AddWithValue("@Address", abonent[2]);
                    command.Parameters.AddWithValue("@Year", abonent[3]);
                    command.ExecuteNonQuery();
                    dataGridView1.DataSource = dataTable;
                }

                // Clear the buffer list and update the buffer DataGridView
                bufferList.Clear();
                bufferDataGridView.Rows.Clear(); // Clear the bufferDataGridView
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка при додаванні записів до бази даних: " + ex.Message);
            }
            finally
            {
                // Close the database connection
                db.CloseConnection();
            }
        }

        private void ButtonAdd_Click(object sender, EventArgs e)
        {
            // Get data from text boxes
            string lastName = ((TextBox)panel1.Controls[0]).Text;
            string phoneNumber = ((TextBox)panel1.Controls[1]).Text;
            string address = ((TextBox)panel1.Controls[2]).Text;
            string year = ((TextBox)panel1.Controls[3]).Text;


            if (!int.TryParse(year, out _))
            {
                MessageBox.Show("Введіть рік як число");
                return;
            }

            // Check if all fields are filled
            if (string.IsNullOrWhiteSpace(lastName) || string.IsNullOrWhiteSpace(phoneNumber) ||
                string.IsNullOrWhiteSpace(address) || string.IsNullOrWhiteSpace(year))
            {
                MessageBox.Show("Please fill in all fields.");
                return;
            }

            // Add data to buffer list
            bufferList.Add(new string[] { lastName, phoneNumber, address, year });

            // Add data from buffer list to bufferDataGridView
            bufferDataGridView.Rows.Add(); // Add a new row
            int rowIndex = bufferDataGridView.Rows.Count - 1; // Get the index of the last row

            // Set the values of cells in the new row
            bufferDataGridView.Rows[rowIndex].Cells["LastName"].Value = lastName;
            bufferDataGridView.Rows[rowIndex].Cells["PhoneNumber"].Value = phoneNumber;
            bufferDataGridView.Rows[rowIndex].Cells["Address"].Value = address;
            bufferDataGridView.Rows[rowIndex].Cells["Year"].Value = year;
        }

        private void bufferDataGridView_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            // Отримуємо індекс видаляємого рядка
            int rowIndex = e.Row.Index;

            // Отримуємо дані про абонента, який видаляється
            string lastName = bufferDataGridView.Rows[rowIndex].Cells["LastName"].Value.ToString();
            string phoneNumber = bufferDataGridView.Rows[rowIndex].Cells["PhoneNumber"].Value.ToString();
            string address = bufferDataGridView.Rows[rowIndex].Cells["Address"].Value.ToString();
            string year = bufferDataGridView.Rows[rowIndex].Cells["Year"].Value.ToString();

            // Перевіряємо дані абонента
            DialogResult result = MessageBox.Show($"Видалення абонента:\nLast Name: {lastName}\nPhone Number: {phoneNumber}\nAddress: {address}\nYear: {year}", "Підтвердження видалення", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                // Видалення абонента
                bufferList.RemoveAll(abonent => abonent[0] == lastName && abonent[1] == phoneNumber && abonent[2] == address && abonent[3] == year);
            }
        }

        private void bufferDataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int rowIndex = e.RowIndex;
            int columnIndex = e.ColumnIndex;

            // Отримання нового значення зміненої комірки
            string newValue = bufferDataGridView.Rows[rowIndex].Cells[columnIndex].Value.ToString();

            // Оновлення відповідного запису у списку bufferList
            bufferList[rowIndex][columnIndex] = newValue;
        }

        private void DataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            // Отримання інформації про змінену комірку
            int rowIndex = e.RowIndex;
            int columnIndex = e.ColumnIndex;
            string newValue = dataGridView1.Rows[rowIndex].Cells[columnIndex].Value.ToString();

            // Отримання ідентифікатора або будь-яких інших даних, які можуть бути потрібні для оновлення в базі даних
            string id = dataGridView1.Rows[rowIndex].Cells["id"].Value.ToString();

            // Оновлення відповідного рядка в базі даних
            string columnName = dataGridView1.Columns[columnIndex].Name;
            string query = $"UPDATE abonent SET {columnName} = @NewValue WHERE id = @Id";

            using (MySqlCommand command = new MySqlCommand(query, db.GetConnection()))
            {
                command.Parameters.AddWithValue("@NewValue", newValue);
                command.Parameters.AddWithValue("@Id", id);
                try
                {
                    db.OpenConnection();
                    int rowsAffected = command.ExecuteNonQuery();
                    if (rowsAffected > 0)
                    {
                        MessageBox.Show("Дані успішно оновлено в базі даних.");
                    }
                    else
                    {
                        MessageBox.Show("Не вдалося оновити дані в базі даних.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Помилка при оновленні даних в базі даних: " + ex.Message);
                }
                finally
                {
                    db.CloseConnection();
                }
            }
        }

        private void DataGridView1_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            // Отримання індексу видаляємого рядка
            int rowIndex = e.Row.Index;

            // Отримання інформації про абонента, який видаляється
            string name = dataGridView1.Rows[rowIndex].Cells["name"].Value.ToString();
            string number = dataGridView1.Rows[rowIndex].Cells["number"].Value.ToString();
            string address = dataGridView1.Rows[rowIndex].Cells["address"].Value.ToString();
            string year = dataGridView1.Rows[rowIndex].Cells["year"].Value.ToString();
            // Виконання операції видалення в базі даних
            DialogResult result = MessageBox.Show($"Видалення абонента:\nLast Name: {name}\nPhone Number: {number}\nAddress: {address}\nYear: {year}", "Підтвердження видалення", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.No)
            {
                return;
            }
            string query = "DELETE FROM abonent WHERE name = @Name AND number = @Number AND address = @Address AND year = @Year";

            using (MySqlCommand command = new MySqlCommand(query, db.GetConnection()))
            {
                command.Parameters.AddWithValue("@Name", name);
                command.Parameters.AddWithValue("@Number", number);
                command.Parameters.AddWithValue("@Address", address);
                command.Parameters.AddWithValue("@Year", year);
                try
                {
                    db.OpenConnection();
                    int rowsAffected = command.ExecuteNonQuery();
                    if (rowsAffected > 0)
                    {
                        MessageBox.Show($"Абонента {name} успішно видалено з бази даних.");
                    }
                    else
                    {
                        MessageBox.Show("Не вдалося видалити абонента з бази даних.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Помилка при видаленні абонента з бази даних: " + ex.Message);
                }
                finally
                {
                    db.CloseConnection();
                }
            }
        }
        private void PopulateDataGridView(string searchText = "")
        {
            try
            {
                db.OpenConnection(); // Open the database connection
                MySqlCommand command;
                string query;
                if (!string.IsNullOrEmpty(searchText))
                {
                    if (int.TryParse(searchText, out _)) // Якщо searchText є числом
                    {
                        query = "SELECT * FROM abonent WHERE CONVERT(year, CHAR) LIKE @SearchYear";
                        command = new MySqlCommand(query, db.GetConnection());
                        command.Parameters.AddWithValue("@SearchYear", "%" + searchText + "%");
                    }
                    else // Якщо searchText не є числом, тоді шукаємо за текстом у полі name
                    {
                        query = "SELECT * FROM abonent WHERE name LIKE @SearchText";
                        command = new MySqlCommand(query, db.GetConnection());
                        command.Parameters.AddWithValue("@SearchText", "%" + searchText + "%");
                    }
                }
                else // Якщо searchText порожній, вибираємо всі записи
                {
                    query = "SELECT * FROM abonent";
                    command = new MySqlCommand(query, db.GetConnection());
                }

                MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);

                dataGridView1.DataSource = dataTable;
                dataGridView1.Columns["id"].Visible = false;
                dataGridView1.ReadOnly = true;

                if (userRole == "editor")
                {
                    dataGridView1.ReadOnly = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                db.CloseConnection(); // Close the database connection
            }
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // Приховати подію натискання Enter
            }
        }

        private void SetPlaceholderText(TextBox textBox, string text)
        {
            textBox.Text = text;
            textBox.ForeColor = System.Drawing.Color.Gray;
        }


        private void Form1_MouseDown(object sender, MouseEventArgs e)
        {
            if (textBox1.Focused)
            {
                this.ActiveControl = null; // Знімаємо фокус з textBox1
            }
        }


        private void pictureBox3_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_Enter_1(object sender, EventArgs e)
        {
            if (textBox1.Text == placeholderText)
            {
                textBox1.Text = "";
                textBox1.ForeColor = System.Drawing.Color.Black;
            }
        }

        private void textBox1_Leave_1(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox1.Text))
            {
                SetPlaceholderText(textBox1, placeholderText);
                PopulateDataGridView();
            }
        }

        private void pictureBox4_MouseEnter(object sender, EventArgs e)
        {
            Cursor = Cursors.Hand; // Зміна курсора на руку
        }

        private void pictureBox4_MouseLeave(object sender, EventArgs e)
        {
            Cursor = Cursors.Default; // Повернення стандартного курсора
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text == placeholderText || textBox1.Text == "")
            {

            }
            else
            {
                dataGridView1.Visible = true;
                label1.Visible = false;
                label2.Visible = false;
                pictureBox4.Visible = false;
                this.pictureBox5.Visible = true;

            }
            PopulateDataGridView(textBox1.Text);
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            // Створюємо новий екземпляр форми авторизації
            Login loginForm = new Login();

            // Показуємо форму модально, це блокуватиме взаємодію з головною формою,
            // поки не буде закрито вікно авторизації
            DialogResult result = loginForm.ShowDialog();

            // Після закриття вікна авторизації перевіряємо результат
            // Якщо результат Ok, то відображаємо список
            if (result == DialogResult.OK)
            {
                userRole = "editor";
                PopulateDataGridView();
                InitializePanel1();
                dataGridView1.Visible = true;
                dataGridView1.AllowUserToDeleteRows = true;
                label1.Visible = false;
                label2.Visible = false;
                pictureBox4.Visible = false;
                this.pictureBox5.Visible = true;
            }
        }

        // Останній критерій сортування
        private string lastSearchCriteria = "";

        // Остання кількість відсортованих абонентів
        private int lastSortedCount = 0;


        // Функція підрахунку кількості відсортованих абонентів
        private int CountSortedRows(DataGridView dataGridView, string searchText)
        {
            int count = 0;
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                if (!row.IsNewRow)
                {
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        if (cell.Value != null)
                        {
                            count++;
                            break; // Перериваємо внутрішній цикл, коли знайдено відповідне значення
                        }
                    }
                }
            }
            return count;
        }

        // Функція для створення та збереження документа Word
        private void SaveToWord(DataGridView dataGridView, string searchText, string filePath, int rowCount)
        {
            try
            {
                // Створення нового екземпляру додатку Word
                Word.Application wordApp = new Word.Application();
                wordApp.Visible = false;

                // Створення нового документа Word
                Word.Document wordDoc = wordApp.Documents.Add();

                // Додавання заголовка документа
                Word.Paragraph title = wordDoc.Content.Paragraphs.Add();
                title.Range.Text = "Список абонентів";
                title.Range.Font.Bold = 1;
                title.Range.Font.Size = 14;
                title.Range.Font.Name = "Arial"; // Встановлення шрифту
                title.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                title.Format.SpaceAfter = 12; // 12 pt spacing after the paragraph

                // Додавання таблиці з даними з DataGridView
                int rowsCount = dataGridView.Rows.Count;
                int colsCount = dataGridView.Columns.Count;
                object missing = System.Reflection.Missing.Value;
                Word.Range range = wordDoc.Range();
                Word.Table table = wordDoc.Tables.Add(range, rowsCount + 1, colsCount, ref missing, ref missing);

                // Додавання заголовків стопців
                for (int i = 0; i < colsCount; i++)
                {
                    table.Cell(1, i + 1).Range.Text = dataGridView.Columns[i].HeaderText;
                    table.Cell(1, i + 1).Range.Font.Bold = 1;
                    table.Cell(1, i + 1).Range.Font.Name = "Arial"; // Встановлення шрифту
                }

                // Додавання даних
                for (int i = 0; i < rowsCount; i++)
                {
                    for (int j = 0; j < colsCount; j++)
                    {
                        table.Cell(i + 2, j + 1).Range.Text = dataGridView.Rows[i].Cells[j].Value.ToString();
                        table.Cell(i + 2, j + 1).Range.Font.Name = "Arial"; // Встановлення шрифту
                    }
                }

                // Додавання роздільної лінії
                Word.Paragraph separator = wordDoc.Content.Paragraphs.Add();
                separator.Range.InsertParagraphAfter();

                // Додавання інформації про критерій пошуку

                Word.Paragraph searchInfo = wordDoc.Content.Paragraphs.Add();
                searchInfo.Range.Text = $"Кількість абонентів: {rowCount}";
                searchInfo.Range.Font.Size = 12;
                searchInfo.Range.Font.Name = "Arial"; // Встановлення шрифту
                searchInfo.Range.Font.Color = Word.WdColor.wdColorBlack; // Встановлення кольору
                searchInfo.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                searchInfo.Format.SpaceAfter = 6;

                // Додавання інформації про кількість знайдених абонентів
                Word.Paragraph rowCountInfo = wordDoc.Content.Paragraphs.Add();
                rowCountInfo.Range.Text = $"Критерій пошуку: {searchText}";
                rowCountInfo.Range.Font.Size = 12;
                rowCountInfo.Range.Font.Name = "Arial"; // Встановлення шрифту
                rowCountInfo.Range.Font.Color = Word.WdColor.wdColorBlack; // Встановлення кольору
                rowCountInfo.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                rowCountInfo.Format.SpaceAfter = 6;

                // Збереження документа Word
                wordDoc.SaveAs2(filePath); // Використання SaveAs2 для сумісності
                wordDoc.Close();
                wordApp.Quit();

                MessageBox.Show("Файл успішно збережено");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // Підрахунок відсортованих абонентів та збереження в Word
        private void SaveSortedToWord(DataGridView dataGridView, string searchText)
        {
            try
            {
                // Відкриття діалогового вікна для вибору шляху збереження
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Filter = "Word Document|*.docx";
                saveFileDialog1.Title = "Save as Word Document";
                DialogResult result = saveFileDialog1.ShowDialog();

                // Перевірка чи користувач обрав файл
                if (result == DialogResult.OK)
                {   
                    // Підрахунок відсортованих абонентів
                    int sortedCount = CountSortedRows(dataGridView, searchText);
                    if (sortedCount > 0)
                    {
                        // Збереження у файл Word
                        string filePath = saveFileDialog1.FileName;
                        SaveToWord(dataGridView, searchText, filePath, sortedCount);
                        lastSortedCount = sortedCount;
                        lastSearchCriteria = searchText;
                    }
                    else
                    {
                        MessageBox.Show("Немає відповідних записів для збереження");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            // Збереження даних у файл Word
            SaveSortedToWord(dataGridView1, textBox1.Text);
        }



    }
}
