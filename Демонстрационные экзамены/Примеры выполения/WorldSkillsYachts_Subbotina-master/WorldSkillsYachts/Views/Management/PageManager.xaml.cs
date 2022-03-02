using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using WorldSkillsYachts.Models;
using WorldSkillsYachts.Utils;
using WorldSkillsYachts.Views.Management;

namespace WorldSkillsYachts.Views.Pages
{
    /// <summary>
    /// Interaction logic for PageManager.xaml
    /// </summary>
    public partial class PageManager : Page
    {
        public Detail NewDetail { get; set; }
        public Order NewOrder { get; set; }

        public PageManager()
        {
            InitializeComponent();
        }

        private void Page_Loaded(object sender, RoutedEventArgs e)
        {
            cbCustomer.ItemsSource = AppData.db.Customers.ToList();
            cbSalesperson.ItemsSource = AppData.db.Salespersons.ToList();
            cbBoat.ItemsSource = AppData.db.Boats.ToList();
            lbAcc.ItemsSource = GetFit(cbBoat.SelectedItem.ToString());
        }

        private void cbBoat_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            lbAcc.ItemsSource = GetFit(cbBoat.SelectedItem.ToString());
        }

        private string[] GetFit(string boatModel)
        {
            int boatID = AppData.db.Boats.Where(x => x.Model == boatModel).FirstOrDefault().boat_ID;
            List<Fit> fits = AppData.db.Fits.Where(x => x.Boat_ID == boatID).ToList();
            string[] accessories = new string[fits.Count];
            for (int i = 0; i < fits.Count; i++)
            {
                accessories[i] = fits[i].Accessory.AccName;
            }
            return accessories;
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            AddOrder(cbCustomer.Text, cbSalesperson.Text, cbBoat.Text, tbAddress.Text, tbCity.Text);
            string[] accs = new string[lbAcc.Items.Count];
            int i = 0;
            foreach (var item in lbAcc.Items)
            {
                accs[i] = item.ToString();
                i++;
            }
            AddDetail(accs);
            if (MessageBox.Show("Заказ оформлен. Перейти к оформлению контракта?", "", 
                MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                WindowAddContract windowAddContract = new WindowAddContract(NewOrder.Order_ID);
                windowAddContract.ShowDialog();
            }
        }

        private void AddOrder(string customer, string salesperson, string boat, string address, string city)
        {
            NewOrder = new Order
            {
                Date = DateTime.Today,
                Salesperson_ID = AppData.db.Salespersons.Where(x => salesperson.Contains(x.FirstName) &&
                salesperson.Contains(x.FamilyName)).FirstOrDefault().SalesPerson_ID,
                Customer_ID = AppData.db.Customers.Where(x => customer.Contains(x.FistName) &&
                customer.Contains(x.FamilyName)).FirstOrDefault().Customer_ID,
                Boat_ID = AppData.db.Boats.Where(x => x.Model == boat).FirstOrDefault().boat_ID,
                DeliveryAddress = address,
                City = city
            };
            AppData.db.Orders.Add(NewOrder);
            AppData.db.SaveChanges();
        }
        
        private void AddDetail(string[] accessory)
        {
            if (string.IsNullOrEmpty(tbAddress.Text) ||
                string.IsNullOrWhiteSpace(tbAddress.Text) ||
                string.IsNullOrEmpty(tbCity.Text) ||
                string.IsNullOrWhiteSpace(tbCity.Text))
            {
                MessageBox.Show("Заполните все значения!", "", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            NewDetail = new Detail();
            int[] IDs = new int[lbAcc.SelectedItems.Count];
            int i = 0;
            foreach (var item in lbAcc.SelectedItems)
            {
                IDs[i] = AppData.db.Accessories.Where(x => x.AccName == item.ToString()).FirstOrDefault().Accessory_ID;
                i++;
            }

            foreach (var id in IDs)
            {
                NewDetail.Accessory_ID = id;
                NewDetail.Order_ID = NewOrder.Order_ID;
                AppData.db.Details.Add(NewDetail);
                AppData.db.SaveChanges();
            }
        }

        private void btnInvoice_Click(object sender, RoutedEventArgs e)
        {
            WindowAddInvoice windowAddInvoice = new WindowAddInvoice();
            windowAddInvoice.ShowDialog();
        }

        private void btnContract_Click(object sender, RoutedEventArgs e)
        {
            WindowAddContract windowAddContract = new WindowAddContract();
            windowAddContract.ShowDialog();
        }
    }
}