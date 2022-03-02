using System.Windows;
using System.Windows.Controls;
using System.Windows.Navigation;
using WorldSkillsYachts.Models;
using WorldSkillsYachts.Utils;

namespace WorldSkillsYachts.Views.Pages
{
    /// <summary>
    /// Логика взаимодействия для PageAddBoat.xaml
    /// </summary>
    public partial class PageAddBoat : Page
    {
        public Boat NewBoat { get; set; }

        public PageAddBoat(Boat boat)
        {
            InitializeComponent();
            NewBoat = boat;
            isNotMast.IsChecked = true;
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(tbModel.Text) ||
                string.IsNullOrWhiteSpace(tbModel.Text) ||
                string.IsNullOrEmpty(tbBoatType.Text) ||
                string.IsNullOrWhiteSpace(tbBoatType.Text) ||
                string.IsNullOrEmpty(tbNumberOfRowers.Text) ||
                string.IsNullOrWhiteSpace(tbNumberOfRowers.Text) ||
                string.IsNullOrEmpty(tbColour.Text) ||
                string.IsNullOrWhiteSpace(tbColour.Text) ||
                string.IsNullOrEmpty(tbWood.Text) ||
                string.IsNullOrWhiteSpace(tbWood.Text) ||
                string.IsNullOrEmpty(tbBasePrice.Text) ||
                string.IsNullOrWhiteSpace(tbBasePrice.Text) ||
                string.IsNullOrEmpty(tbVAT.Text) ||
                string.IsNullOrWhiteSpace(tbVAT.Text))
            {
                MessageBox.Show("Заполните все данные!", "", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            if (!int.TryParse(tbNumberOfRowers.Text, out int numberOfRowers) ||
                !decimal.TryParse(tbBasePrice.Text, out decimal basePrice) ||
                !int.TryParse(tbVAT.Text, out int vat))
            {
                MessageBox.Show("Исправьте числовые данные!", "", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            if (vat > 100)
            {
                MessageBox.Show("Налог не может быть больше 100%!", "", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            NewBoat.Model = tbModel.Text;
            NewBoat.BoatType = tbBoatType.Text;
            NewBoat.NumberOfRowers = numberOfRowers;
            NewBoat.Mast = (bool)isMast.IsChecked;
            NewBoat.Colour = tbColour.Text;
            NewBoat.Wood = tbWood.Text;
            NewBoat.BasePrice = basePrice;
            NewBoat.VAT = vat.ToString() + "%";
            if (NewBoat.boat_ID == 0)
            {
                AppData.db.Boats.Add(NewBoat);
            }
            AppData.db.SaveChanges();

            if (NavigationService.CanGoBack)
            {
                NavigationService.GoBack();
            }
        }

        private void Page_Loaded(object sender, RoutedEventArgs e)
        {
            if (NewBoat == null)
            {
                NewBoat = new Boat();
                NavigationService.Navigate(new PageAddBoat(NewBoat));
            }
            if (NewBoat.boat_ID != 0)
            {
                tbModel.Text = NewBoat.Model;
                tbBoatType.Text = NewBoat.BoatType;
                tbNumberOfRowers.Text = NewBoat.NumberOfRowers.ToString();
                if (NewBoat.Mast) isMast.IsChecked = true;
                else isNotMast.IsChecked = true;
                tbColour.Text = NewBoat.Colour;
                tbWood.Text = NewBoat.Wood;
                tbBasePrice.Text = NewBoat.BasePrice.ToString();
                tbVAT.Text = NewBoat.VAT.TrimEnd('%');
            }
        }
    }
}