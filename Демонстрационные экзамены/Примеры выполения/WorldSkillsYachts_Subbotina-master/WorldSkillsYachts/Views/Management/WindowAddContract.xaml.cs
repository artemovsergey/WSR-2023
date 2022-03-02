using System;
using System.Linq;
using System.Windows;
using WorldSkillsYachts.Models;
using WorldSkillsYachts.Utils;

namespace WorldSkillsYachts.Views.Management
{
    /// <summary>
    /// Логика взаимодействия для WindowAddContract.xaml
    /// </summary>
    public partial class WindowAddContract : Window
    {
        private int order_ID;

        public Contract NewContract { get; set; }

        public WindowAddContract()
        {
            InitializeComponent();
            spOrder.Visibility = Visibility.Visible;
            cbNumber.ItemsSource = AppData.db.Orders.ToList();
        }

        public WindowAddContract(int order_ID)
        {
            InitializeComponent();
            spOrder.Visibility = Visibility.Collapsed;
            this.order_ID = order_ID;
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            if (spOrder.Visibility == Visibility.Visible)
                order_ID = Convert.ToInt32(cbNumber.Text);

            if (string.IsNullOrEmpty(tbDeposit.Text) ||
                string.IsNullOrWhiteSpace(tbDeposit.Text) ||
                string.IsNullOrEmpty(tbSum.Text) ||
                string.IsNullOrWhiteSpace(tbSum.Text) ||
                string.IsNullOrEmpty(tbSumVAT.Text) ||
                string.IsNullOrWhiteSpace(tbSumVAT.Text))
            {
                MessageBox.Show("Заполните все значения!", "", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (!decimal.TryParse(tbDeposit.Text, out decimal deposit) ||
                !decimal.TryParse(tbSum.Text, out decimal sum) ||
                !decimal.TryParse(tbSumVAT.Text, out decimal sumVAT))
            {
                MessageBox.Show("Исправьте числовые данные!", "", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            AddContract(deposit, sum, sumVAT);
            if (MessageBox.Show("Контракт оформлен. Перейти к оформлению счета?", "",
                MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                WindowAddInvoice windowAddInvoice = new WindowAddInvoice(NewContract.Contract_ID);
                windowAddInvoice.ShowDialog();
                Close();
            }
        }

        private void AddContract(decimal deposit, decimal sum, decimal sumVAT)
        {
            NewContract = new Contract
            {
                Date = DateTime.Today,
                DepositPayed = deposit,
                Order_ID = order_ID,
                ContractTotalPrice = sum,
                ContracTotalPrice_inclVAT = sumVAT,
                ProductionProcess = "Работы не начаты"
            };
            AppData.db.Contracts.Add(NewContract);
            AppData.db.SaveChanges();
        }
    }
}
