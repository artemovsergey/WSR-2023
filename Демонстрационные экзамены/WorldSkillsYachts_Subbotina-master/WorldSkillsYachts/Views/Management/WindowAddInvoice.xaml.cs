using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using WorldSkillsYachts.Models;
using WorldSkillsYachts.Utils;

namespace WorldSkillsYachts.Views.Management
{
    /// <summary>
    /// Логика взаимодействия для WindowAddInvoice.xaml
    /// </summary>
    public partial class WindowAddInvoice : Window
    {
        private int contract_ID;
        public Invoice NewInvoice { get; set; }

        public WindowAddInvoice()
        {
            InitializeComponent();
            spContract.Visibility = Visibility.Visible;
            cbContractNumber.ItemsSource = AppData.db.Contracts.ToList();
        }

        public WindowAddInvoice(int contract_ID)
        {
            InitializeComponent();
            spContract.Visibility = Visibility.Collapsed;
            this.contract_ID = contract_ID;
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            if (spContract.Visibility == Visibility.Visible)
                contract_ID = Convert.ToInt32(cbContractNumber.Text);

            if (string.IsNullOrEmpty(tbSum.Text) ||
                string.IsNullOrWhiteSpace(tbSum.Text) ||
                string.IsNullOrEmpty(tbSumVAT.Text) ||
                string.IsNullOrWhiteSpace(tbSumVAT.Text))
            {
                MessageBox.Show("Заполните все значения!", "", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            if (!decimal.TryParse(tbSum.Text, out decimal sum) ||
                !decimal.TryParse(tbSumVAT.Text, out decimal sumVAT))
            {
                MessageBox.Show("Исправьте числовые данные!", "", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            AddInvoice(sum, (bool)isSettled.IsChecked, sumVAT);
            MessageBox.Show("Счет оформлен");
            Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void AddInvoice(decimal sum, bool settled, decimal sumVAT)
        {
            NewInvoice = new Invoice
            {
                Contract_ID = contract_ID,
                Settled = settled,
                Sum = sum,
                Sum_inclVAT = sumVAT,
                Date = DateTime.Today
            };
            AppData.db.Invoices.Add(NewInvoice);
            AppData.db.SaveChanges();
        }
    }
}
