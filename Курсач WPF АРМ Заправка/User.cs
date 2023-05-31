using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace Курсач_WPF_АРМ_Заправка
{
    public class User : INotifyPropertyChanged
    {
        private int id;
        public int ID
        {
            get { return id; }
            set { id = value; OnPropertyChanged(); }
        }

        private string username;
        public string Username
        {
            get { return username; }
            set { username = value; OnPropertyChanged(); }
        }

        private string password;
        public string Password
        {
            get { return password; }
            set { password = value; OnPropertyChanged(); }
        }
        private string role;
        public string Role
        {
            get { return role; }
            set { if (role == "admin" || role == "employee") { role = value; } OnPropertyChanged(); }
        }
        private string fuel_price;
        public string Fuel_price
        {
            get { return fuel_price; }
            set { fuel_price = value; OnPropertyChanged(); }
        }
        private string remaining_fuel;
        public string Remaining_fuel
        {
            get { return remaining_fuel; }
            set { remaining_fuel = value; OnPropertyChanged(); }
        }

        // Реализация интерфейса INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

}
