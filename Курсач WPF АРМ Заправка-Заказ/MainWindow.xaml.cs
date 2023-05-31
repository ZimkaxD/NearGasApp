using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.IO.Pipes;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace Курсач_WPF_АРМ_Заправка_Заказ
{
    public partial class MainWindow : Window
    {
        private const int PortNumber = 12345; // Порт для прослушивания
        private TcpListener listener;
        private Thread listenerThread;
        private int buttonIndex = 1;
        private DBManager databaseManager = new DBManager();

        public MainWindow()
        {
            InitializeComponent();
            StartListening();
        }

        private void StartListening()
        {
            try
            {
                // Запускаем прослушивание на указанном порту
                listener = new TcpListener(IPAddress.Parse("127.0.0.1"), PortNumber);
                listener.Start();

                // Запускаем новый поток для прослушивания подключений
                listenerThread = new Thread(ListenForClients);
                listenerThread.SetApartmentState(ApartmentState.STA); // Устанавливаем STA для потока
                listenerThread.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при прослушивании порта: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ListenForClients()
        {
            while (true)
            {
                try
                {
                    TcpClient client = listener.AcceptTcpClient();
                    StringBuilder formData = new StringBuilder();
                    byte[] buffer = new byte[4096];
                    int bytesRead;

                    while ((bytesRead = client.GetStream().Read(buffer, 0, buffer.Length)) > 0)
                    {
                        formData = new StringBuilder();
                        formData.Append(Encoding.UTF8.GetString(buffer, 0, bytesRead));
                    }

                    Dispatcher.InvokeAsync(() =>
                    {
                        Button appButton = new Button();
                        appButton.Content = "NearGas " + "Заказ";
                        appButton.Margin = new Thickness(0, 10, 0, 0);
                        appButton.Click += async (sender, e) =>
                        {

                            MessageBoxResult result = MessageBox.Show(formData.ToString(), (string)appButton.Content, MessageBoxButton.OKCancel, MessageBoxImage.Information);
                            // Извлечение типа топлива
                            int fuelTypeStartIndex = formData.ToString().IndexOf("Тип топлива: ") + "Тип топлива: ".Length;
                            int fuelTypeEndIndex = formData.ToString().IndexOf(",", fuelTypeStartIndex);
                            string fuelType = formData.ToString().Substring(fuelTypeStartIndex, fuelTypeEndIndex - fuelTypeStartIndex);

                            // Извлечение количества топлива
                            int fuelQuantityStartIndex = formData.ToString().IndexOf("Количество: ") + "Количество: ".Length;
                            int fuelQuantityEndIndex = formData.ToString().IndexOf(" л", fuelQuantityStartIndex);
                            string fuelQuantityString = formData.ToString().Substring(fuelQuantityStartIndex, fuelQuantityEndIndex - fuelQuantityStartIndex);
                            double fuelQuantity = double.Parse(fuelQuantityString);

                            if (result == MessageBoxResult.OK)
                            {
                                await SendConfirmationToFirstApp((string)appButton.Content+" "+buttonIndex+" принят"+" "+formData.ToString());
                                UpdateFuelQuantityInDatabase(fuelType,fuelQuantity);
                                await Dispatcher.InvokeAsync(() => appButtonsStackPanel.Children.Remove(appButton));

                            }
                            else if (result == MessageBoxResult.Cancel)
                            {
                                await SendConfirmationToFirstApp((string)appButton.Content +" "+ buttonIndex + " отклонён" +" "+ formData.ToString());
                                await Dispatcher.InvokeAsync(() => appButtonsStackPanel.Children.Remove(appButton));
                            }
                        };

                        appButtonsStackPanel.Children.Add(appButton);
                        buttonIndex++;


                    });

                    client.Close();
                }
                catch (SocketException)
                {
                    break;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при принятии подключения: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
        private void UpdateFuelQuantityInDatabase(string fuelType, double fuelQuantity)
        {
            databaseManager.openConnection();
            try
            {
                string updateQuery = $"UPDATE FuelQuantity SET quantity = quantity + {fuelQuantity} WHERE fuel_type = '{fuelType}'";


                using (SqlCommand command = new SqlCommand(updateQuery, databaseManager.getConnection()))
                {
                    command.ExecuteNonQuery();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при обновлении данных: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }

        private async Task SendConfirmationToFirstApp(string message)
        {
            using (NamedPipeClientStream pipeClient = new NamedPipeClientStream(".", "MyNamedPipe", PipeDirection.Out))
            {
                try
                {
                    await pipeClient.ConnectAsync(); // Асинхронное подключение к именованному каналу

                    using (StreamWriter sw = new StreamWriter(pipeClient))
                    {
                        await sw.WriteLineAsync(message); // Асинхронная отправка сообщения
                        await sw.FlushAsync();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при отправке данных: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }






        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);

            // Остановка прослушивания при закрытии приложения
            listener.Stop();
            listenerThread.Join();
        }
    }
}
