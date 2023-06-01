using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
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
using System.Runtime.InteropServices;

namespace Курсач_WPF_АРМ_Заправка
{
    /// <summary>
    /// Логика взаимодействия для MainMenuControl.xaml
    /// </summary>
    /// 
    public partial class MainMenuControlAdmin : UserControl
    {
        private Brush currentBackground = Brushes.White;

        public MainMenuControlAdmin()
        {
            InitializeComponent();
            
        }

        private void OpenFuelWindowButton_Click(object sender, RoutedEventArgs e)
        {
            
            FuelCostCalculating fuelUserControl = new FuelCostCalculating();
                MainMenuControlAdminn.Children.Clear();
                MainMenuControlAdminn.Children.Add(fuelUserControl);
        }

        private void ModeButton_Click(object sender, RoutedEventArgs e)
        {
            Color veryDarkColor = Color.FromRgb(20, 20, 20);
            if (currentBackground == Brushes.White)
            {
                currentBackground = new SolidColorBrush(veryDarkColor); ;
            }
            else
            {
                currentBackground = Brushes.White;
            }
            MainMenuControlAdminn.Background = currentBackground;
        }

        private void BuyFuelButton_Click(object sender, RoutedEventArgs e)
        {
            BuyFuel buyFuel = new BuyFuel();
            MainMenuControlAdminn.Children.Clear();
            MainMenuControlAdminn.Children.Add(buyFuel);
        }

        private void AdminPanelButton_Click(object sender, RoutedEventArgs e)
        {
            AdminGrid adminGrid = new AdminGrid();
            MainMenuControlAdminn.Children.Clear();
            MainMenuControlAdminn.Children.Add(adminGrid);
        }
    }
}
