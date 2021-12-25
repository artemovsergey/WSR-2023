using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Navigation;
using WorldSkillsYachts.Models;
using WorldSkillsYachts.Utils;

namespace WorldSkillsYachts.Views.Pages
{
    /// <summary>
    /// Логика взаимодействия для PageUsers.xaml
    /// </summary>
    public partial class PageUsers : Page
    {
        public PageUsers()
        {
            InitializeComponent();
        }

        private void EditButton_Click(object sender, RoutedEventArgs e)
        {
            User user = UserGrid.SelectedItem as User;
            NavigationService.Navigate(new PageAddUser(user));
        }

        private void RemoveButton_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Подтвердите удаление", "", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                User user = UserGrid.SelectedItem as User;
                AppData.db.Users.Remove(user);
                AppData.db.SaveChanges();
                GetUsers();
            }
        }

        private void GetUsers()
        {
            List<User> users = AppData.db.Users.ToList();
            UserGrid.ItemsSource = users;
        }

        private void Page_Loaded(object sender, RoutedEventArgs e)
        {
            GetUsers();
        }

        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new PageAddUser(new User()));
        }
    }
}