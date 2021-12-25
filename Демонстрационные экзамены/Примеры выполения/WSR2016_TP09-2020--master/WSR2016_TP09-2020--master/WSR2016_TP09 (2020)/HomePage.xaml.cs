using System.Windows;
using System.Windows.Controls;

namespace WSR2016_TP09__2020_
{
    public partial class HomePage : Page
    {
        public MainWindow mainWindow;
        public HomePage(MainWindow _mainWindow)
        {
            InitializeComponent();
            mainWindow = _mainWindow;
            _MinimizexExit();
        }

        // Кнопки Свернуть и Закрыть
        void _MinimizexExit()
        {
            _Minimize.Click += (s, a) => { mainWindow.WindowState = WindowState.Minimized; };
            xExit.Click += (s, a) => { mainWindow.Close(); };
            mainWindow.MouseLeftButtonDown += (s, a) => { mainWindow.DragMove(); };
        }
    }
}
