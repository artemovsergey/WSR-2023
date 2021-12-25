using System.Windows;

namespace WorldSkillsYachts
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void ButtonBack_Click(object sender, RoutedEventArgs e)
        {
            MainFrame.GoBack();
        }

        private void MainFrame_ContentRendered(object sender, System.EventArgs e)
        {
            if (MainFrame.CanGoBack)
                ButtonBack.Visibility = Visibility.Visible;
            else
                ButtonBack.Visibility = Visibility.Collapsed;
        }
    }
}