using System.Windows.Controls;

namespace WorldSkillsYachts.Views.Pages
{
    /// <summary>
    /// Interaction logic for PageAdmin.xaml
    /// </summary>
    public partial class PageAdmin : Page
    {
        public PageAdmin()
        {
            InitializeComponent();
        }

        private void ManagerButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            NavigationService.Navigate(new PageManager());
        }

        private void AdminButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            NavigationService.Navigate(new PageAdminTasks());
        }
    }
}