using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using HtmlAgilityPack;
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
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System.Runtime.InteropServices;

namespace Курсач_WPF_АРМ_Заправка
{

    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            //this.WindowState = WindowState.Maximized;
            Awtorization awtorization = new Awtorization();
            MainMenu.Children.Add(awtorization);
            
        }

    }
}
