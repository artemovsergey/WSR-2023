using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Navigation;
using WorldSkillsYachts.Models;
using WorldSkillsYachts.Utils;

namespace WorldSkillsYachts.Views.Pages
{
    /// <summary>
    /// Логика взаимодействия для PageAddUser.xaml
    /// </summary>
    public partial class PageAddUser : Page
    {
        private User NewUser { get; set; }

        public PageAddUser(User user)
        {
            InitializeComponent();
            NewUser = user;
            isNotBlocked.IsChecked = true;
            cbRole.SelectedIndex = 1;
        }

        private void Page_Loaded(object sender, RoutedEventArgs e)
        {
            cbRole.ItemsSource = AppData.db.Roles.ToList();
            if (NewUser == null)
            {
                NewUser = new User();
                NavigationService.Navigate(new PageAddUser(NewUser));
            }
            if (NewUser.User_ID != 0)
            {
                tbLogin.Text = NewUser.Login;
                tbPassword.Text = NewUser.Password;
                cbRole.SelectedIndex = (int)NewUser.Role_ID - 1;
                if (NewUser.IsBlocked)
                    isBlocked.IsChecked = true;
                else
                    isNotBlocked.IsChecked = true;
            }
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(tbLogin.Text) ||
                string.IsNullOrWhiteSpace(tbLogin.Text) ||
                string.IsNullOrEmpty(tbPassword.Text) ||
                string.IsNullOrWhiteSpace(tbPassword.Text))
            {
                MessageBox.Show("Заполните все значения!", "", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            NewUser.Login = tbLogin.Text;
            NewUser.Password = tbPassword.Text;
            NewUser.Role_ID = cbRole.SelectedIndex + 1;
            NewUser.IsBlocked = (bool)isBlocked.IsChecked;
            NewUser.RegisteredDate = DateTime.Today;
            if (NewUser.User_ID == 0)
            {
                AppData.db.Users.Add(NewUser);
            }
            AppData.db.SaveChanges();

            if (NavigationService.CanGoBack)
            {
                NavigationService.GoBack();
            }
        }
    }
}