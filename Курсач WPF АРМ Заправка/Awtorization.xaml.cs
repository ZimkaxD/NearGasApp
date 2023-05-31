using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
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
using HtmlAgilityPack;
using System.Globalization;


namespace Курсач_WPF_АРМ_Заправка
{
    /// <summary>
    /// Логика взаимодействия для Awtorization.xaml
    /// </summary>
    public partial class Awtorization : UserControl
    {
        private DBManager databaseManager = new DBManager();
        private string Ai95Price;
        private string Ai96Price;
        private string Ai97Price;
        private string Ai98Price;
        private string Ai99Price;
        private string Ai100Price;
        public static string accessLevel = "";

        public Awtorization()
        {
            InitializeComponent();

    }



        private void Login_Click(object sender, RoutedEventArgs e)
        {
            string username = usernameTextBox.Text;
            string password = passwordBox.Text;

            SqlDataAdapter adapter = new SqlDataAdapter();
            DataTable table = new DataTable();

            string querystring = $"select id, login, password, access_level from employees where login='{username}' and password='{password}'";

            SqlCommand  command = new SqlCommand(querystring,databaseManager.getConnection());

            adapter.SelectCommand=command;
            //Парсинг данных с веб-страницы
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--headless"); 
            options.AddArgument("--disable-gpu"); 
            options.AddArgument("--silent");
            var driverService = ChromeDriverService.CreateDefaultService();
            driverService.HideCommandPromptWindow = true;

            using (IWebDriver driver = new ChromeDriver(driverService,options))
            {
                try
                {
                   
                    driver.Navigate().GoToUrl("https://azs.belorusneft.by/sitebeloil/ru/center/azs/center/fuelandService/price/");

                    
                    IWebElement tabliza = driver.FindElement(By.XPath("//*[@id='leftcontainer']/div[1]/div[1]/table"));

                    
                    IReadOnlyCollection<IWebElement> rows = tabliza.FindElements(By.XPath(".//tr"));

                    int index = 0;
                    List<string> prices = new List<string>();


                    foreach (IWebElement row in rows)
                    {
                        
                        IReadOnlyCollection<IWebElement> cells = row.FindElements(By.XPath(".//td"));

                        foreach (IWebElement cell in cells)
                        {
                           
                            string price = cell.Text;
                            if (price.Contains("."))
                            {
                                price = price.Replace(".", ",");
                            }
                            prices.Add(price);
                        }

                        index++;
                        if (index >= 6)
                            break;
                    }

                    Ai95Price = prices[0];
                    Ai96Price = prices[1];
                    Ai97Price = prices[2];
                    Ai98Price = prices[3];
                    Ai99Price = prices[4];
                    Ai100Price = prices[5];
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Возникла ошибка: {ex.Message}");
                }
            }
                    // Парсинг и запись данных в таблицу
            string[] fuelTypes = { "АВТОГАЗ (ПБА)*", "ДТ", "ДТ ECO", "АИ-92", "АИ-95", "АИ-98" };
            double[] priceNodes = { Double.Parse(Ai95Price), Double.Parse(Ai96Price), Double.Parse(Ai97Price), Double.Parse(Ai98Price), Double.Parse(Ai99Price), Double.Parse(Ai100Price) };
            double[] fuelQuantitys = { 2360.3453, 3243.234, 1233.342, 3423.323, 2343.324, 5432.243 };
            using (SqlCommand clearFuelCommand = new SqlCommand("TRUNCATE TABLE FuelPrices",databaseManager.getConnection()))
            {
                databaseManager.openConnection();
                clearFuelCommand.ExecuteNonQuery();
                databaseManager.closeConnection();
            }
                if (priceNodes != null && priceNodes.Length == fuelTypes.Length)
                {
                    for (int i = 0; i < fuelTypes.Length; i++)
                    {
                            string fuelType = fuelTypes[i];
                            decimal price = Convert.ToDecimal(priceNodes[i]);


                            // Вставка данных в таблицу
                            string insertQuery = "INSERT INTO FuelPrices (fuel_type, price) VALUES (@fuel_type, @price)";

                            using (SqlCommand insertCommand = new SqlCommand(insertQuery, databaseManager.getConnection()))
                            {

                                databaseManager.openConnection();
                                insertCommand.Parameters.AddWithValue("@fuel_type", fuelType);
                                insertCommand.Parameters.AddWithValue("@price", price);
                                insertCommand.ExecuteNonQuery();
                                databaseManager.closeConnection();
                            }                      
                    }
                }

                adapter.Fill(table);

                if(table.Rows.Count ==1 ) 
                {
                accessLevel = table.Rows[0]["access_level"].ToString();
                if (accessLevel == "admin")
                    {
                    // Выполните действия для администратора
                        MainMenuControlAdmin mainMenuControlAdminn = new MainMenuControlAdmin();

                        Awtorizationn.Children.Clear();
                        Awtorizationn.Children.Add(mainMenuControlAdminn);
                    }
                    else if (accessLevel == "user")
                    {
                        // Выполните действия для обычного пользователя
                        MainMenuControl mainMenuControlUser = new MainMenuControl();

                    Awtorizationn.Children.Clear();
                        Awtorizationn.Children.Add(mainMenuControlUser);
                    }
                }
                else
                {
                    MessageBox.Show("У вас нет доступа к системе", "Ошибка доступа", MessageBoxButton.OK, MessageBoxImage.Warning);
                }

        }
    }
}
