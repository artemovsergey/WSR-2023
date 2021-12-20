using System.Windows;

namespace WSR2016_TP09__2020_
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            OpenPage(pages.HomePage);
        }

        // Список доступных страниц.
        public enum pages
        {
            HomePage
        }
        
        // Метод открытия страниц.
        public void OpenPage(pages pages)
        {
            // Стартовая страница.
            if (pages == pages.HomePage)
                frame.Navigate(new HomePage(this));
        }
    }
}
