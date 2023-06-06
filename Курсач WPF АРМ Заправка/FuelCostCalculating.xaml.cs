using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
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
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using HtmlAgilityPack;


namespace Курсач_WPF_АРМ_Заправка
{
    public partial class FuelCostCalculating : UserControl
    {
        private double fuelAmount;
        private double fuelPrice;
        List<Label> labelTypeList = new List<Label>();
        List<Label> labelPriceList = new List<Label>();
        List<Label> labelFuelQuantityList = new List<Label>();
        int labelIndex = 0;
        List<int> SalesCount = new List<int> {0,0,0,0,0,0};
        List<double> SalesAmount = new List<double> { 0.0, 0.0, 0.0, 0.0, 0.0, 0.0 };
        List<double> FuelAmount = new List<double> { 0.0, 0.0, 0.0, 0.0, 0.0, 0.0 };
        private DBManager databaseManager = new DBManager();

        public FuelCostCalculating()
        {
            InitializeComponent();



            labelTypeList.Add(Ai95);
            labelTypeList.Add(Ai96);
            labelTypeList.Add(Ai97);
            labelTypeList.Add(Ai98);
            labelTypeList.Add(Ai99);
            labelTypeList.Add(Ai100);
            labelPriceList.Add(Ai95Price);
            labelPriceList.Add(Ai96Price);
            labelPriceList.Add(Ai97Price);
            labelPriceList.Add(Ai98Price);
            labelPriceList.Add(Ai99Price);
            labelPriceList.Add(Ai100Price);
            labelFuelQuantityList.Add(Ai95Availability);
            labelFuelQuantityList.Add(Ai96Availability);
            labelFuelQuantityList.Add(Ai97Availability);
            labelFuelQuantityList.Add(Ai98Availability);
            labelFuelQuantityList.Add(Ai99Availability);
            labelFuelQuantityList.Add(Ai100Availability);

            FillFuelPriceFromDatabase();

            FillFuelQuantityFromDatabase();
        }

        private void CloseFuelWindowButton_Click(object sender, RoutedEventArgs e)
        {

            if (Awtorization.accessLevel=="admin" )
            {
                // Выполните действия для администратора
                MainMenuControlAdmin mainMenuControlAdmin = new MainMenuControlAdmin();
                fuelCostCalculating.Children.Clear();
                fuelCostCalculating.Children.Add(mainMenuControlAdmin);
            }
            else if (Awtorization.accessLevel == "user")
            {
                // Выполните действия для обычного пользователя
                MainMenuControl mainMenuControlUser = new MainMenuControl();
                fuelCostCalculating.Children.Clear();
                fuelCostCalculating.Children.Add(mainMenuControlUser);
            }
            else
            {
                MessageBox.Show("ТЫ кто?", "Ошибка доступа", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }


        private void calculateButton_Click(object sender, RoutedEventArgs e)
        {
            Dictionary<string, double> fuelAmounts = new Dictionary<string, double>();

            StringBuilder messageBuilder = new StringBuilder();

            bool insufficientFuel = false;

            if (Ai95Check.IsChecked == true)
            {
                double fuelAmount = fuelSliderAi95.Value;
                double fuelQuantity = GetFuelQuantityByType(Ai95.Content.ToString());

                if (fuelAmount > fuelQuantity)
                {
                    insufficientFuel = true;
                    double fuelShortage = fuelAmount - fuelQuantity;
                    messageBuilder.AppendLine($"Недостаточное количество топлива \"{Ai95.Content}\": {fuelShortage} л. Пожалуйста, выберите другое количество или тип топлива.");
                }
                else
                {
                    fuelAmounts.Add(Ai95.Content.ToString(), fuelAmount);
                }
            }

            if (Ai96Check.IsChecked == true)
            {
                double fuelAmount = fuelSliderAi96.Value;
                double fuelQuantity = GetFuelQuantityByType(Ai96.Content.ToString());

                if (fuelAmount > fuelQuantity)
                {
                    insufficientFuel = true;
                    double fuelShortage = fuelAmount - fuelQuantity;
                    messageBuilder.AppendLine($"Недостаточное количество топлива \"{Ai96.Content}\": {fuelShortage} л. Пожалуйста, выберите другое количество или тип топлива.");
                }
                else
                {
                    fuelAmounts.Add(Ai96.Content.ToString(), fuelAmount);
                }
            }
            if (Ai97Check.IsChecked == true)
            {
                double fuelAmount = fuelSliderAi97.Value;
                double fuelQuantity = GetFuelQuantityByType(Ai97.Content.ToString());

                if (fuelAmount > fuelQuantity)
                {
                    insufficientFuel = true;
                    double fuelShortage = fuelAmount - fuelQuantity;
                    messageBuilder.AppendLine($"Недостаточное количество топлива \"{Ai97.Content}\": {fuelShortage} л. Пожалуйста, выберите другое количество или тип топлива.");
                }
                else
                {
                    fuelAmounts.Add(Ai97.Content.ToString(), fuelAmount);
                }
            }
            if (Ai98Check.IsChecked == true)
            {
                double fuelAmount = fuelSliderAi98.Value;
                double fuelQuantity = GetFuelQuantityByType(Ai98.Content.ToString());

                if (fuelAmount > fuelQuantity)
                {
                    insufficientFuel = true;
                    double fuelShortage = fuelAmount - fuelQuantity;
                    messageBuilder.AppendLine($"Недостаточное количество топлива \"{Ai98.Content}\": {fuelShortage} л. Пожалуйста, выберите другое количество или тип топлива.");
                }
                else
                {
                    fuelAmounts.Add(Ai98.Content.ToString(), fuelAmount);
                }
            }
            if (Ai99Check.IsChecked == true)
            {
                double fuelAmount = fuelSliderAi99.Value;
                double fuelQuantity = GetFuelQuantityByType(Ai99.Content.ToString());

                if (fuelAmount > fuelQuantity)
                {
                    insufficientFuel = true;
                    double fuelShortage = fuelAmount - fuelQuantity;
                    messageBuilder.AppendLine($"Недостаточное количество топлива \"{Ai99.Content}\": {fuelShortage} л. Пожалуйста, выберите другое количество или тип топлива.");
                }
                else
                {
                    fuelAmounts.Add(Ai99.Content.ToString(), fuelAmount);
                }
            }
            if (Ai100Check.IsChecked == true)
            {
                double fuelAmount = fuelSliderAi100.Value;
                double fuelQuantity = GetFuelQuantityByType(Ai100.Content.ToString());

                if (fuelAmount > fuelQuantity)
                {
                    insufficientFuel = true;
                    double fuelShortage = fuelAmount - fuelQuantity;
                    messageBuilder.AppendLine($"Недостаточное количество топлива \"{Ai100.Content}\": {fuelShortage} л. Пожалуйста, выберите другое количество или тип топлива.");
                }
                else
                {
                    fuelAmounts.Add(Ai100.Content.ToString(), fuelAmount);
                    Dictionary<string, double> fuelPrices = CalculateFuelPrice();

                    double totalFuelQuantity = 0.0;
                    double newFuelQuantity = 0.0;

                    foreach (var fuel in fuelPrices)
                    {
                        string fuelType = fuel.Key;
                        double totalPrice = fuel.Value;
                        messageBuilder.AppendLine($"Тип топлива: {fuelType}, Стоимость: {totalPrice} руб.");
                        fuelAmount = GetFuelAmountByType(fuelType);
                        totalFuelQuantity += fuelAmount;
                        newFuelQuantity -= fuelAmount;
                        messageBuilder.AppendLine($"Количество {fuelType}: {fuelAmount} л" + "\n");
                    }

                    double totalFuelPrice = fuelPrices.Sum(fuel => fuel.Value);
                    messageBuilder.AppendLine($"Общая стоимость: {totalFuelPrice} руб.");
                    messageBuilder.AppendLine($"Общее количество топлива: {totalFuelQuantity} л");

                    // Обновление количества топлива в базе данных
                    UpdateFuelQuantityInDatabase(fuelAmounts);

                    // Обновление отображения доступного количества топлива
                    FillFuelQuantityFromDatabase();
                }
            }

            else
            {
                // Выполнение расчетов и отображение информации о стоимости и количестве топлива...
                Dictionary<string, double> fuelPrices = CalculateFuelPrice();

                double totalFuelQuantity = 0.0;
                double newFuelQuantity = 0.0;

                foreach (var fuel in fuelPrices)
                {
                    string fuelType = fuel.Key;
                    double totalPrice = fuel.Value;
                    messageBuilder.AppendLine($"Тип топлива: {fuelType}, Стоимость: {totalPrice} руб.");
                    double fuelAmount = GetFuelAmountByType(fuelType);
                    totalFuelQuantity += fuelAmount;
                    newFuelQuantity -= fuelAmount;
                    messageBuilder.AppendLine($"Количество {fuelType}: {fuelAmount} л" + "\n");
                }

                double totalFuelPrice = fuelPrices.Sum(fuel => fuel.Value);
                messageBuilder.AppendLine($"Общая стоимость: {totalFuelPrice} руб.");
                messageBuilder.AppendLine($"Общее количество топлива: {totalFuelQuantity} л");

                // Обновление количества топлива в базе данных
                UpdateFuelQuantityInDatabase(fuelAmounts);

                // Обновление отображения доступного количества топлива
                FillFuelQuantityFromDatabase();
            }

            string message = messageBuilder.ToString();
            MessageBox.Show(message, "Расчет стоимости топлива", MessageBoxButton.OK, MessageBoxImage.Information);
        }



        public Dictionary<string, double> CalculateFuelPrice()
        {
            Dictionary<string, double> fuelPrices = new Dictionary<string, double>();
            Dictionary <string, double> fuelAmounts = new Dictionary<string, double>();

            StringBuilder messageBuilder = new StringBuilder();

            if (Ai95Check.IsChecked == true)
            {
                double fuelAmount = fuelSliderAi95.Value;
                double fuelPrice = Convert.ToDouble(Ai95Price.Content.ToString().Split(' ')[0]);
                double totalPrice = fuelAmount * fuelPrice;
                SalesAmount[0] += totalPrice;
                fuelPrices.Add(Ai95.Content.ToString(), totalPrice);
                fuelAmounts.Add(Ai95.Content.ToString(), fuelAmount);
            }

            if (Ai96Check.IsChecked == true)
            {
                double fuelAmount = fuelSliderAi96.Value;
                double fuelPrice = Convert.ToDouble(Ai96Price.Content.ToString().Split(' ')[0]);
                double totalPrice = fuelAmount * fuelPrice;
                SalesAmount[1] += totalPrice;
                fuelPrices.Add(Ai96.Content.ToString(), totalPrice);
                fuelAmounts.Add(Ai96.Content.ToString(), fuelAmount);
            }

            if (Ai97Check.IsChecked == true)
            {
                double fuelAmount = fuelSliderAi97.Value;
                double fuelPrice = Convert.ToDouble(Ai97Price.Content.ToString().Split(' ')[0]);
                double totalPrice = fuelAmount * fuelPrice;
                SalesAmount[2] += totalPrice;
                fuelPrices.Add(Ai97.Content.ToString(), totalPrice);
                fuelAmounts.Add(Ai97.Content.ToString(), fuelAmount);
            }

            if (Ai98Check.IsChecked == true)
            {
                double fuelAmount = fuelSliderAi98.Value;
                double fuelPrice = Convert.ToDouble(Ai98Price.Content.ToString().Split(' ')[0]);
                double totalPrice = fuelAmount * fuelPrice;
                SalesAmount[3] += totalPrice;
                fuelPrices.Add(Ai98.Content.ToString(), totalPrice);
                fuelAmounts.Add(Ai98.Content.ToString(), fuelAmount);
            }

            if (Ai99Check.IsChecked == true)
            {
                double fuelAmount = fuelSliderAi99.Value;
                double fuelPrice = Convert.ToDouble(Ai99Price.Content.ToString().Split(' ')[0]);
                double totalPrice = fuelAmount * fuelPrice;
                SalesAmount[4] += totalPrice;
                fuelPrices.Add(Ai99.Content.ToString(), totalPrice);
                fuelAmounts.Add(Ai99.Content.ToString(), fuelAmount);
            }

            if (Ai100Check.IsChecked == true)
            {
                double fuelAmount = fuelSliderAi100.Value;
                double fuelPrice = Convert.ToDouble(Ai100Price.Content.ToString().Split(' ')[0]);
                double totalPrice = fuelAmount * fuelPrice;
                SalesAmount[5] += totalPrice;
                fuelPrices.Add(Ai100.Content.ToString(), totalPrice);
                fuelAmounts.Add(Ai100.Content.ToString(), fuelAmount);
            }

            return fuelPrices;
        }

        private double GetFuelAmountByType(string fuelType)
        {
            double fuelAmount = 0.0;

            switch (fuelType)
            {
                case "АВТОГАЗ (ПБА)*":
                    fuelAmount = fuelSliderAi95.Value;
                    SalesCount[0] += 1;
                    FuelAmount[0] += fuelAmount;
                    break;
                case "ДТ":
                    fuelAmount = fuelSliderAi96.Value;
                    SalesCount[1] += 1;
                    FuelAmount[1] += fuelAmount;
                    break;
                case "ДТ ECO":
                    fuelAmount = fuelSliderAi97.Value;
                    SalesCount[2] += 1;
                    FuelAmount[2] += fuelAmount;
                    break;
                case "АИ-92":
                    fuelAmount = fuelSliderAi98.Value;
                    SalesCount[3] += 1;
                    FuelAmount[3] += fuelAmount;
                    break;
                case "АИ-95":
                    fuelAmount = fuelSliderAi99.Value;
                    SalesCount[4] += 1;
                    FuelAmount[4] += fuelAmount;
                    break;
                case "АИ-98":
                    fuelAmount = fuelSliderAi100.Value;
                    SalesCount[5] += 1;
                    FuelAmount[5] += fuelAmount;
                    break;
            }

            return fuelAmount;
        }


        private void FillFuelPriceFromDatabase()
        {
            string query = "SELECT fuel_type,price FROM FuelPrices";

            databaseManager.openConnection();
            using (SqlCommand command = new SqlCommand(query, databaseManager.getConnection()))
            {
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read() && labelIndex < 6)
                {
                    string value = reader.GetString(0);
                    string valuePrice = reader.GetDecimal(1).ToString();
                    Label labelPrice = labelPriceList[labelIndex];
                    Label label = labelTypeList[labelIndex];
                    labelPrice.Content = valuePrice + " руб.";
                    label.Content = value;
                    labelIndex++;
                }
                reader.Close();
                labelIndex = 0;
            }         
        }
        private void FillFuelQuantityFromDatabase()
        {
            string query = "SELECT fuel_type, quantity FROM FuelQuantity";

            databaseManager.openConnection();

            using (SqlCommand command = new SqlCommand(query, databaseManager.getConnection()))
            {
                SqlDataReader reader = command.ExecuteReader();

                while (reader.Read() && labelIndex < 6)
                {
                    string fuelType = reader.GetString(0);
                    string fuelQuantity = reader.GetDecimal(1).ToString();

                    Label labelQuantity = labelFuelQuantityList[labelIndex];
                    labelQuantity.Content = fuelQuantity+" л";

                    labelIndex++;
                }
                reader.Close();
                labelIndex = 0;
            }
        }

        private double GetFuelQuantityByType(string fuelType)
        {
            double fuelQuantity = 0.0;

            switch (fuelType)
            {
                case "АВТОГАЗ (ПБА)*":
                    fuelQuantity = Convert.ToDouble(Ai95Availability.Content.ToString().Split(' ')[0]);
                    break;
                case "ДТ":
                    fuelQuantity = Convert.ToDouble(Ai96Availability.Content.ToString().Split(' ')[0]);
                    break;
                case "ДТ ECO":
                    fuelQuantity = Convert.ToDouble(Ai97Availability.Content.ToString().Split(' ')[0]);
                    break;
                case "АИ-92":
                    fuelQuantity = Convert.ToDouble(Ai98Availability.Content.ToString().Split(' ')[0]);
                    break;
                case "АИ-95":
                    fuelQuantity = Convert.ToDouble(Ai99Availability.Content.ToString().Split(' ')[0]);
                    break;
                case "АИ-98":
                    fuelQuantity = Convert.ToDouble(Ai100Availability.Content.ToString().Split(' ')[0]);
                    break;
            }

            return fuelQuantity;
        }

        private void UpdateFuelQuantityInDatabase(Dictionary<string, double> fuelAmounts)
        {
            databaseManager.openConnection();
            foreach (var fuel in fuelAmounts)
            {
                string fuelType = fuel.Key;
                double fuelAmount = fuel.Value;

                string query = "UPDATE FuelQuantity SET quantity = quantity - @fuelAmount WHERE fuel_type = @fuelType";

                using (SqlCommand command = new SqlCommand(query, databaseManager.getConnection()))
                {
                    command.Parameters.AddWithValue("@fuelAmount", fuelAmount);
                    command.Parameters.AddWithValue("@fuelType", fuelType);
                    command.ExecuteNonQuery();
                }
            }
        }

        private void fuelSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {

        }

        private void Ai95Check_Checked(object sender, RoutedEventArgs e)
        {
            fuelSliderAi95.IsEnabled = true;
        }

        private void Ai96Check_Checked(object sender, RoutedEventArgs e)
        {
            fuelSliderAi96.IsEnabled = true;
        }

        private void Ai97Check_Checked(object sender, RoutedEventArgs e)
        {
            fuelSliderAi97.IsEnabled = true;
        }

        private void Ai98Check_Checked(object sender, RoutedEventArgs e)
        {
            fuelSliderAi98.IsEnabled = true;
        }

        private void Ai99Check_Checked(object sender, RoutedEventArgs e)
        {
            fuelSliderAi99.IsEnabled = true;
        }

        private void Ai100Check_Checked(object sender, RoutedEventArgs e)
        {
            fuelSliderAi100.IsEnabled = true;
        }

        private void Ai95Check_Unchecked(object sender, RoutedEventArgs e)
        {
            fuelSliderAi95.IsEnabled = false;
        }

        private void Ai96Check_Unchecked(object sender, RoutedEventArgs e)
        {
            fuelSliderAi96.IsEnabled = false;
        }

        private void Ai97Check_Unchecked(object sender, RoutedEventArgs e)
        {
            fuelSliderAi97.IsEnabled = false;
        }

        private void Ai98Check_Unchecked(object sender, RoutedEventArgs e)
        {
            fuelSliderAi98.IsEnabled = false;
        }

        private void Ai99Check_Unchecked(object sender, RoutedEventArgs e)
        {
            fuelSliderAi99.IsEnabled = false;
        }

        private void Ai100Check_Unchecked(object sender, RoutedEventArgs e)
        {
            fuelSliderAi100.IsEnabled = false;
        }

        private void DaySallesButton_Click(object sender, RoutedEventArgs e)
        {
            // Создаем приложение MS Word
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

            try
            {
                // Создаем новый документ
                Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Add();

                // Заголовок отчета
                Microsoft.Office.Interop.Word.Paragraph title = doc.Content.Paragraphs.Add();
                title.Range.Text = "Отчет о продажах за день";
                title.Range.Font.Bold = 1;
                title.Range.Font.Size = 14;
                title.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                title.Range.InsertParagraphAfter();

                // Дата
                Microsoft.Office.Interop.Word.Paragraph dateParagraph = doc.Content.Paragraphs.Add();
                dateParagraph.Range.Text = "Дата: " + DateTime.Now.ToShortDateString();
                dateParagraph.Range.Font.Bold = 0;
                dateParagraph.Range.Font.Size = 12;
                dateParagraph.Range.InsertParagraphAfter();

                // Данные о продажах
                List<SaleData> salesData = GetSalesData(); 

                Microsoft.Office.Interop.Word.Table salesTable = doc.Tables.Add(dateParagraph.Range, salesData.Count + 1, 4);
                salesTable.Borders.Enable = 1;

                // Заголовки столбцов
                salesTable.Cell(1, 1).Range.Text = "Топливо";
                salesTable.Cell(1, 2).Range.Text = "Количество продаж";
                salesTable.Cell(1, 3).Range.Text = "Количество топлива";
                salesTable.Cell(1, 4).Range.Text = "Сумма";


                // Заполнение таблицы данными
                for (int i = 0; i < salesData.Count; i++)
                {
                    salesTable.Cell(i + 2, 1).Range.Text = salesData[i].FuelType;
                    salesTable.Cell(i + 2, 2).Range.Text = salesData[i].SaleCount.ToString();
                    salesTable.Cell(i + 2, 3).Range.Text = salesData[i].FuelAmount.ToString();
                    salesTable.Cell(i + 2, 4).Range.Text = salesData[i].SaleAmount.ToString();

                }

                // Сохранение документа
                object fileName = System.IO.Path.Combine(Environment.CurrentDirectory, "SalesReport.docx"); 
                doc.SaveAs2(fileName);
                doc=wordApp.Documents.Open(fileName, ReadOnly:true);
                
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
            finally
            {

            }
        }
        

            private List<SaleData> GetSalesData()
            {

                List<SaleData> salesData = new List<SaleData>();


                salesData.Add(new SaleData { FuelType = Ai95.Content.ToString(), SaleCount = SalesCount[0], FuelAmount= Math.Round(FuelAmount[0],2).ToString()+" л." ,SaleAmount = Math.Round(SalesAmount[0],2).ToString() + " р." });
                salesData.Add(new SaleData { FuelType = Ai96.Content.ToString(), SaleCount = SalesCount[1], FuelAmount = Math.Round(FuelAmount[1],2).ToString() + " л.", SaleAmount = Math.Round(SalesAmount[1],2).ToString() + " р." });
                salesData.Add(new SaleData { FuelType = Ai97.Content.ToString(), SaleCount = SalesCount[2], FuelAmount = Math.Round(FuelAmount[2],2).ToString() + " л.", SaleAmount = Math.Round(SalesAmount[2],2).ToString() + " р." });
                salesData.Add(new SaleData { FuelType = Ai98.Content.ToString(), SaleCount = SalesCount[3], FuelAmount = Math.Round(FuelAmount[3],2).ToString() + " л.", SaleAmount = Math.Round(SalesAmount[3],2).ToString() + " р." });
                salesData.Add(new SaleData { FuelType = Ai99.Content.ToString(), SaleCount = SalesCount[4], FuelAmount = Math.Round(FuelAmount[4],2).ToString() + " л.", SaleAmount = Math.Round(SalesAmount[4],2).ToString() + " р." });
                salesData.Add(new SaleData { FuelType = Ai100.Content.ToString(), SaleCount = SalesCount[5], FuelAmount = Math.Round(FuelAmount[5],2).ToString() + " л.", SaleAmount = Math.Round(SalesAmount[5],2).ToString() + " р." });
                double TotalFuelAmount = FuelAmount[0] + FuelAmount[1] + FuelAmount[2] + FuelAmount[3] + FuelAmount[4] + FuelAmount[5];
                double TotalSaleAmount = SalesAmount[0] + SalesAmount[1] + SalesAmount[2] + SalesAmount[3] + SalesAmount[4] + SalesAmount[5];
                salesData.Add(new SaleData { FuelType = "Общее", SaleCount = SalesCount[0]+ SalesCount[1]+ SalesCount[2]+ SalesCount[3]+ SalesCount[4]+ SalesCount[5], FuelAmount = Math.Round(TotalFuelAmount,2).ToString()+" л." , SaleAmount = Math.Round(TotalSaleAmount,2).ToString()+" р."  });
            return salesData;
            }

        public class SaleData
        {
            public string FuelType { get; set; }
            public int SaleCount { get; set; }
            public string FuelAmount { get; set; }
            public string SaleAmount { get; set; }
        
        }


        private void EveryDaySallesPlanButton_Click(object sender, RoutedEventArgs e)
        {

        }

        private void FuelAmountOtchetButton_Click(object sender, RoutedEventArgs e)
        {
            // Создаем приложение MS Word
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

            try
            {
                // Создаем новый документ
                Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Add();

                // Заголовок отчета
                Microsoft.Office.Interop.Word.Paragraph title = doc.Content.Paragraphs.Add();
                title.Range.Text = "Отчет о количестве топлива";
                title.Range.Font.Bold = 1;
                title.Range.Font.Size = 14;
                title.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                title.Range.InsertParagraphAfter();

                // Дата
                Microsoft.Office.Interop.Word.Paragraph dateParagraph = doc.Content.Paragraphs.Add();
                dateParagraph.Range.Text = "Дата: " + DateTime.Now.ToShortDateString();
                dateParagraph.Range.Font.Bold = 0;
                dateParagraph.Range.Font.Size = 12;
                dateParagraph.Range.InsertParagraphAfter();

                // Данные о продажах
                List<AmountData> amountData = GetAmountData();

                Microsoft.Office.Interop.Word.Table salesTable = doc.Tables.Add(dateParagraph.Range, amountData.Count + 1, 2);
                salesTable.Borders.Enable = 1;

                // Заголовки столбцов
                salesTable.Cell(1, 1).Range.Text = "Топливо";
                salesTable.Cell(1, 2).Range.Text = "Количество топлива";



                // Заполнение таблицы данными
                for (int i = 0; i < amountData.Count; i++)
                {
                    salesTable.Cell(i + 2, 1).Range.Text = amountData[i].FuelType;
                    salesTable.Cell(i + 2, 2).Range.Text = amountData[i].FuelAmount.ToString();

                }

                // Сохранение документа
                object fileName = System.IO.Path.Combine(Environment.CurrentDirectory, "FuelAmountReport.docx");
                doc.SaveAs2(fileName);
                doc = wordApp.Documents.Open(fileName, ReadOnly: true);


            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
            finally
            {

            }
        }
        private List<AmountData> GetAmountData()
        {

            List<AmountData> amountData = new List<AmountData>();


            amountData.Add(new AmountData { FuelType = Ai95.Content.ToString(), FuelAmount = Ai95Availability.Content.ToString()});
            amountData.Add(new AmountData { FuelType = Ai96.Content.ToString(), FuelAmount = Ai96Availability.Content.ToString() });
            amountData.Add(new AmountData { FuelType = Ai97.Content.ToString(), FuelAmount = Ai97Availability.Content.ToString() });
            amountData.Add(new AmountData { FuelType = Ai98.Content.ToString(), FuelAmount = Ai98Availability.Content.ToString() });
            amountData.Add(new AmountData { FuelType = Ai99.Content.ToString(), FuelAmount = Ai99Availability.Content.ToString() });
            amountData.Add(new AmountData { FuelType = Ai100.Content.ToString(), FuelAmount = Ai100Availability.Content.ToString() });
            double Ai95Fuel = Convert.ToDouble(Ai95Availability.Content.ToString().Split(' ')[0]); double Ai96Fuel = Convert.ToDouble(Ai96Availability.Content.ToString().Split(' ')[0]); double Ai97Fuel= Convert.ToDouble(Ai97Availability.Content.ToString().Split(' ')[0]); double Ai98Fuel= Convert.ToDouble(Ai98Availability.Content.ToString().Split(' ')[0]); double Ai99Fuel= Convert.ToDouble(Ai99Availability.Content.ToString().Split(' ')[0]); double Ai100Fuel= Convert.ToDouble(Ai100Availability.Content.ToString().Split(' ')[0]);
            double TotalFuelAmount = Ai95Fuel + Ai96Fuel + Ai97Fuel + Ai98Fuel + Ai99Fuel + Ai100Fuel;
            amountData.Add(new AmountData { FuelType = "Общее", FuelAmount = TotalFuelAmount.ToString()+" л."});
            return amountData;
        }

        public class AmountData
        {
            public string FuelType { get; set; }
            public string FuelAmount { get; set; }

        }
    }
}

