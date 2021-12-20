using System;
using System.Windows;
using WorldSkillsYachts.Models;
using WorldSkillsYachts.Utils;

namespace WorldSkillsYachts.Views.Windows
{
    /// <summary>
    /// Interaction logic for WindowChangePassword.xaml
    /// </summary>
    public partial class WindowChangePassword : Window
    {
        private readonly User user;

        public WindowChangePassword()
        {
            InitializeComponent();
        }

        public WindowChangePassword(User user)
        {
            InitializeComponent();
            this.user = user;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (PassBox.Password == RepeatPassBox.Password)
            {
                if (PassBox.Password != user.Password)
                {
                    user.Password = PassBox.Password;
                    user.PasswordChanged = DateTime.Today;
                    AppData.db.SaveChanges();
                    MessageBox.Show("Пароль успешно изменен. Авторизируйтесь.", "", MessageBoxButton.OK, MessageBoxImage.Information);
                    Close();
                }
                else
                {
                    MessageBox.Show("Старый и новый пароли не должны совпадать!", "Внимание!", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
            }
            else
            {
                MessageBox.Show("Пароли не совпадают!", "Внимание!", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
        }
    }
}