using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Navigation;
using WorldSkillsYachts.Models;
using WorldSkillsYachts.Utils;

namespace WorldSkillsYachts.Views.Pages
{
    /// <summary>
    /// Логика взаимодействия для PageAddAccessory.xaml
    /// </summary>
    public partial class PageAddAccessory : Page
    {
        public Accessory NewAccessory { get; set; }

        public PageAddAccessory(Accessory accessory)
        {
            InitializeComponent();
            NewAccessory = accessory;
            cbPartner.SelectedItem = 0;
        }

        private void Page_Loaded(object sender, RoutedEventArgs e)
        {
            cbPartner.ItemsSource = AppData.db.Partners.ToList();
            if (NewAccessory == null)
            {
                NewAccessory = new Accessory();
                NavigationService.Navigate(new PageAddAccessory(NewAccessory));
            }
            if (NewAccessory.Accessory_ID != 0)
            {
                tbAccName.Text = NewAccessory.AccName;
                tbDescriptionOfAccessory.Text = NewAccessory.DescriptionOfAccessory;
                tbPrice.Text = NewAccessory.Price.ToString();
                tbVAT.Text = NewAccessory.VAT.TrimEnd('%');
                tbInventory.Text = NewAccessory.Inventory.ToString();
                tbOrderLevel.Text = NewAccessory.OrderLevel.ToString();
                tbOrderBatch.Text = NewAccessory.OrderBatch.ToString();
                cbPartner.SelectedItem = NewAccessory.Partner;
            }
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(tbAccName.Text) ||
                string.IsNullOrWhiteSpace(tbAccName.Text) ||
                string.IsNullOrEmpty(tbDescriptionOfAccessory.Text) ||
                string.IsNullOrWhiteSpace(tbDescriptionOfAccessory.Text) ||
                string.IsNullOrEmpty(tbPrice.Text) ||
                string.IsNullOrWhiteSpace(tbPrice.Text) ||
                string.IsNullOrEmpty(tbVAT.Text) ||
                string.IsNullOrWhiteSpace(tbVAT.Text) ||
                string.IsNullOrEmpty(tbInventory.Text) ||
                string.IsNullOrWhiteSpace(tbInventory.Text) ||
                string.IsNullOrEmpty(tbOrderLevel.Text) ||
                string.IsNullOrWhiteSpace(tbOrderLevel.Text) ||
                string.IsNullOrEmpty(tbOrderBatch.Text) ||
                string.IsNullOrWhiteSpace(tbOrderBatch.Text) ||
                string.IsNullOrEmpty(cbPartner.Text) ||
                string.IsNullOrWhiteSpace(cbPartner.Text))
            {
                MessageBox.Show("Заполните все значения!", "", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            if (!decimal.TryParse(tbPrice.Text, out decimal price) ||
                !int.TryParse(tbVAT.Text, out int vat) ||
                !int.TryParse(tbInventory.Text, out int inventory) ||
                !int.TryParse(tbOrderLevel.Text, out int orderLevel) ||
                !int.TryParse(tbOrderBatch.Text, out int orderBatch))
            {
                MessageBox.Show("Исправьте числовые данные!", "", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            if (vat > 100)
            {
                MessageBox.Show("Налог не может быть больше 100%!", "", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            NewAccessory.AccName = tbAccName.Text;
            NewAccessory.DescriptionOfAccessory = tbDescriptionOfAccessory.Text;
            NewAccessory.Price = price;
            NewAccessory.VAT = vat.ToString() + "%";
            NewAccessory.Inventory = inventory;
            NewAccessory.OrderLevel = orderLevel;
            NewAccessory.OrderBatch = orderBatch;
            NewAccessory.Partner = cbPartner.SelectedItem as Partner;
            if (NewAccessory.Accessory_ID == 0)
            {
                AppData.db.Accessories.Add(NewAccessory);
            }
            AppData.db.SaveChanges();

            if (NavigationService.CanGoBack)
            {
                NavigationService.GoBack();
            }
        }
    }
}