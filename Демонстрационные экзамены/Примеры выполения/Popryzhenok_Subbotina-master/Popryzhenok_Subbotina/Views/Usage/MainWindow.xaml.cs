using System.Windows;

namespace Popryzhenok_Subbotina
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

        private void MainFrame_ContentRendered(object sender, System.EventArgs e)
        {
            ButtonBack.Visibility = MainFrame.CanGoBack ?
                Visibility.Visible : Visibility.Collapsed;
        }

        private void ButtonBack_Click(object sender, RoutedEventArgs e)
        {
            MainFrame.GoBack();
        }
    }
}