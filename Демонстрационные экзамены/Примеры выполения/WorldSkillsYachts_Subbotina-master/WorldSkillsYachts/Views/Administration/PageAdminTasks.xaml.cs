using System.Windows;
using System.Windows.Controls;
using System.Windows.Navigation;

namespace WorldSkillsYachts.Views.Pages
{
    /// <summary>
    /// Логика взаимодействия для PageAdminTasks.xaml
    /// </summary>
    public partial class PageAdminTasks : Page
    {
        public PageAdminTasks()
        {
            InitializeComponent();
        }

        private void ButtonBoats_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new PageBoats());
        }

        private void ButtonAccessory_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new PageAccessories());
        }

        private void ButtonUsers_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new PageUsers());
        }
    }
}