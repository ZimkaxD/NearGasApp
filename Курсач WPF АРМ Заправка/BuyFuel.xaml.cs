using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace Курсач_WPF_АРМ_Заправка
{
    public partial class BuyFuel : UserControl
    {
        private List<string> FuelTypes = new List<string> { "АВТОГАЗ (ПБА)*", "ДТ", "ДТ ECO", "АИ-92", "АИ-95", "АИ-98" };
        private List<string> ChangeFuelTypes = new List<string> { "АВТОГАЗ (ПБА)*", "ДТ", "ДТ ECO", "АИ-92", "АИ-95", "АИ-98" };
        private List<string> OriginalFuelTypes;

        private DBManager databaseManager = new DBManager();

        private const int PortNumber = 12345; // Порт для подключения

        public BuyFuel()
        {
            InitializeComponent();
            InitializeFuelTypeComboBox();
           
        }

        private void InitializeFuelTypeComboBox()
        {
            OriginalFuelTypes = new List<string>(FuelTypes);
        }

        private void fuelTypeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int selectedFuelCount = fuelTypeComboBox.SelectedIndex + 1;
            CreateFuelLabels(selectedFuelCount);
        }

        private void CreateFuelLabels(int fuelCount)
        {
            fuelStackPanel.Children.Clear();
            ResetFuelTypes();

            for (int i = 0; i < fuelCount; i++)
            {
                StackPanel fuelPanel = new StackPanel();
                fuelPanel.Orientation = Orientation.Horizontal;

                ComboBox fuelTypeComboBox = new ComboBox();
                fuelTypeComboBox.Name = "fuelTypeComboBox" + i;
                fuelTypeComboBox.ItemsSource = FuelTypes;
                fuelTypeComboBox.Width = 200;
                fuelTypeComboBox.Height = 20;
                fuelTypeComboBox.SelectionChanged += FuelTypeComboBox_SelectionChanged;

                TextBox fuelQuantityTextBox = new TextBox();
                fuelQuantityTextBox.Name = "fuelTextBox" + i;

                fuelPanel.Children.Add(fuelTypeComboBox);

                var spacer = new TextBlock();
                spacer.Width = 20;
                fuelPanel.Children.Add(spacer);

                fuelPanel.Children.Add(fuelQuantityTextBox);

                fuelStackPanel.Children.Add(fuelPanel);
            }
            AddSendButton();
        }

        private void FuelTypeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox selectedComboBox = (ComboBox)sender;
            string selectedFuelType = selectedComboBox.SelectedItem.ToString();

            // Проверяем, есть ли выбранный тип топлива в списке FuelTypes
            if (FuelTypes.Contains(selectedFuelType))
            {
                // Удаляем выбранный тип топлива из списка
                FuelTypes.Remove(selectedFuelType);
            }

            // Добавляем предыдущий выбранный тип топлива обратно в список
            if (e.RemovedItems.Count > 0)
            {
                string previousFuelType = e.RemovedItems[0].ToString();
                FuelTypes.Add(previousFuelType);
            }

            selectedComboBox.IsEnabled = false;
        }




        private void AddSendButton()
        {

            StackPanel fuelPanel = new StackPanel();
            fuelPanel.Orientation = Orientation.Horizontal;
            Button sendButton = new Button();
            sendButton.Content = "Отправить";


            sendButton.Margin = new Thickness(0, 10, 0, 0);

            sendButton.Click += SendButton_Click;

            fuelPanel.Children.Add(sendButton);
            fuelStackPanel.Children.Add(fuelPanel);
        }

        private async void SendButton_Click(object sender, RoutedEventArgs e)
        {
            StringBuilder formData = new StringBuilder();
            foreach (StackPanel fuelPanel in fuelStackPanel.Children)
            {
                ComboBox fuelTypeComboBox = fuelPanel.Children.OfType<ComboBox>().FirstOrDefault();
                TextBox fuelQuantityTextBox = fuelPanel.Children.OfType<TextBox>().FirstOrDefault();

                if (fuelTypeComboBox != null && fuelQuantityTextBox != null)
                {
                    string fuelType = fuelTypeComboBox.SelectedItem?.ToString();
                    string fuelQuantity = fuelQuantityTextBox.Text;

                    formData.AppendLine($"Тип топлива: {fuelType}, Количество: {fuelQuantity + " л"}");
                }
            }

            try
            {
                using (TcpClient client = new TcpClient())
                {
                    await client.ConnectAsync(IPAddress.Parse("127.0.0.1"), PortNumber);

                    byte[] data = Encoding.UTF8.GetBytes(formData.ToString());
                    await client.GetStream().WriteAsync(data, 0, data.Length);
                }
                await SendConfirmationToFirstApp(formData.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при отправке данных: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task SendConfirmationToFirstApp(string formData)
        {
            using (NamedPipeServerStream pipeServer = new NamedPipeServerStream("MyNamedPipe", PipeDirection.In))
            {
                try
                {
                    await pipeServer.WaitForConnectionAsync();

                    using (StreamReader sr = new StreamReader(pipeServer))
                    {
                        string receivedMessage = await sr.ReadLineAsync();

                        Dispatcher.Invoke(() =>
                        {
                            MessageBox.Show(receivedMessage, "Информация о заказе", MessageBoxButton.OK, MessageBoxImage.Information);
                        });
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при принятии подключения: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }


        


        private void CloseFuelWindowButton_Click(object sender, RoutedEventArgs e)
        {

            if (Awtorization.accessLevel == "admin")
            {
                // Выполните действия для администратора
                MainMenuControlAdmin mainMenuControlAdmin = new MainMenuControlAdmin();
                BuyDuelControl.Children.Clear();
                mainMenuControlAdmin.InitializeComponent();
                BuyDuelControl.Children.Add(mainMenuControlAdmin);
            }
            else if (Awtorization.accessLevel == "user")
            {
                // Выполните действия для обычного пользователя
                MainMenuControl mainMenuControlUser = new MainMenuControl();
                BuyDuelControl.Children.Clear();
                mainMenuControlUser.InitializeComponent();
                BuyDuelControl.Children.Add(mainMenuControlUser);
            }
            else
            {
                MessageBox.Show("ТЫ кто?", "Ошибка доступа", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }


        public void ResetFuelTypes()
        {
            FuelTypes = new List<string>(OriginalFuelTypes);
        }
    }

}
