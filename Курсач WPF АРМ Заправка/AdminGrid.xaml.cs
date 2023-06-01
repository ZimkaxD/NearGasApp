using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Курсач_WPF_АРМ_Заправка
{

    public partial class AdminGrid : UserControl
    {
        private DBManager databaseManager = new DBManager();
        public AdminGrid()
        {
            InitializeComponent();
            RefreshDataGrid();
        }

        private void AddUser(object sender, RoutedEventArgs e)
        {
            string loginValue = login.Text;
            string passwordValue = password.Text;
            string accessLevel = "user";

            if (login.Text.Length > 15 || password.Text.Length > 15 || login.Text.Length<6|| password.Text.Length<6)
            {
                MessageBox.Show("Не правильная длина Логин или Пароль. Введите другие значения.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else
            {
                AddUserToDatabase(loginValue, passwordValue, accessLevel);
                RefreshDataGrid();
                MessageBox.Show("Логин:"+login.Text+" \n"+"Пароль:"+password.Text, "Новый пользователь добавлен", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void DeleteUser(object sender, RoutedEventArgs e)
        {
            string selectedUser = LoginsComboBox.SelectedItem as string;
            if (selectedUser == null)
            {
                MessageBox.Show("Выберите пользователя из списка.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else
            {
                DeleteUserFromDatabase(selectedUser);
                RefreshDataGrid();
                MessageBox.Show("Пользователь с этими данными удалён:\nЛогин:" + selectedUser , "Пользователь удалён", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void AddUserToDatabase(string login, string password, string accessLevel)
        {
            try
            {
                databaseManager.openConnection();

                string query = "INSERT INTO employees (login, password, access_level) VALUES (@login, @password, @accessLevel)";
                SqlCommand command = new SqlCommand(query, databaseManager.getConnection());
                command.Parameters.AddWithValue("@login", login);
                command.Parameters.AddWithValue("@password", password);
                command.Parameters.AddWithValue("@accessLevel", accessLevel);

                command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                
            }
            finally
            {
                databaseManager.closeConnection();
            }
        }
        private List<string> GetLoginsFromDataGrid()
        {
            List<string> logins = new List<string>();

           
            var dataTable = ((DataView)data.ItemsSource).Table;

            
            foreach (DataRow row in dataTable.Rows)
            {
                if (row["access_level"].ToString() == "user")
                {       
                    logins.Add(row["login"].ToString());
                }
            }

            return logins;
        }


        private void DeleteUserFromDatabase(string login)
        {

            try
            {
                databaseManager.openConnection();

                string query = "DELETE FROM employees WHERE login = @login";
                SqlCommand command = new SqlCommand(query, databaseManager.getConnection());
                command.Parameters.AddWithValue("@login", login);

                command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                
            }
            finally
            {
                databaseManager.closeConnection();
            }
        }
        private void RefreshDataGrid()
        {


            try
            {

                databaseManager.openConnection();

                string query = "SELECT * FROM employees ";
                SqlCommand command = new SqlCommand(query, databaseManager.getConnection());
                SqlDataAdapter adapter = new SqlDataAdapter(command);
                DataSet dataset = new DataSet();
                adapter.Fill(dataset, "employees");
                data.ItemsSource = dataset.Tables["employees"].DefaultView;
                List<string> logins = GetLoginsFromDataGrid();
                LoginsComboBox.ItemsSource = logins;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Не удалось загрузить таблицу из базы данных.","Ошибка",MessageBoxButton.OK,MessageBoxImage.Error);
            }
            finally
            {
                databaseManager.closeConnection();
            }
        }

        private void data_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            DataRowView selectedRow = data.SelectedItem as DataRowView;

            if (selectedRow != null && selectedRow.Row.Table.Columns.Contains("login"))
            {
                
                string selectedUser = selectedRow.Row["login"].ToString();

                LoginsComboBox.SelectedItem = selectedUser;
            }
            else
            {
                
            }
        }
        private void DataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
           

            if (e.Column is DataGridBoundColumn column && e.EditingElement is TextBox textBox)
            {
             
                string newValue = textBox.Text.Trim();

               
                DataRowView selectedRow = (DataRowView)e.Row.Item;

                
                if (selectedRow != null && selectedRow.Row.Table.Columns.Contains("login") && selectedRow.Row.Table.Columns.Contains("password"))
                {

                    string loginValue = selectedRow.Row["login"].ToString();
                    string passwordValue = selectedRow.Row["password"].ToString();

                    UpdateDataInDatabase(loginValue, passwordValue, column.Header.ToString(), newValue);
                    RefreshDataGrid();
                    
                }
                
            }
        }
        private void UpdateDataInDatabase(string login, string password, string columnName, string newValue)
        {
            try
            {
                databaseManager.openConnection();

                string query;
                if (columnName == "password")
                {
                    query = $"UPDATE employees SET {columnName} = @newValue WHERE login = @login";
                    MessageBox.Show("Пароль изменён", "Пароль изменён", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    query = $"UPDATE employees SET {columnName} = @newValue WHERE login = @login";
                    MessageBox.Show("Логин изменён", "Логин изменён", MessageBoxButton.OK, MessageBoxImage.Information);
                }

                SqlCommand command = new SqlCommand(query, databaseManager.getConnection());
                command.Parameters.AddWithValue("@newValue", newValue);
                command.Parameters.AddWithValue("@login", login);

                command.ExecuteNonQuery();
            
            }
            catch (Exception ex)
            {
                // Обработка ошибок при выполнении запроса
            }
            finally
            {
                databaseManager.closeConnection();
            }
        }

        private void CloseFuelWindowButton_Click(object sender, RoutedEventArgs e)
        {
            if (Awtorization.accessLevel == "admin")
            {
                // Выполните действия для администратора
                MainMenuControlAdmin mainMenuControlAdmin = new MainMenuControlAdmin();
                AdminGridd.Children.Clear();
                AdminGridd.Children.Add(mainMenuControlAdmin);
            }
            else if (Awtorization.accessLevel == "user")
            {
                // Выполните действия для обычного пользователя
                MainMenuControl mainMenuControlUser = new MainMenuControl();
                AdminGridd.Children.Clear();
                AdminGridd.Children.Add(mainMenuControlUser);
            }
            else
            {
                MessageBox.Show("ТЫ кто?", "Ошибка доступа", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        private void DataGrid_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {

            DataRowView selectedRow = (DataRowView)e.Row.Item;
            DataGridColumn column = e.Column;

            if (column.Header.ToString() == "access_level")
            {
                e.Cancel = true;
                MessageBox.Show("Редактирование уровня доступа невозможно", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Stop);
            }

            if (selectedRow != null && selectedRow.Row.Table.Columns.Contains("access_level"))
            {

                string accessLevelValue = selectedRow.Row["access_level"].ToString();


                if (accessLevelValue == "admin")
                {
                    e.Cancel = true;
                    MessageBox.Show("Вам запрещено редактирование данного пользователя", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Stop);
                }
            }
        }

    }
}
