using System;
using System.Collections.Generic;
using System.Windows;
using WorldSkillsYachts.Models;
using WorldSkillsYachts.Utils;

namespace WorldSkillsYachts.Views.Windows
{
    /// <summary>
    /// Логика взаимодействия для WindowChangePrice.xaml
    /// </summary>
    public partial class WindowChangePrice : Window
    {
        private List<Boat> boats;
        private List<Accessory> accessories;
        private bool flag;

        public WindowChangePrice(List<Boat> boats)
        {
            InitializeComponent();
            this.boats = boats;
            flag = true;
        }

        public WindowChangePrice(List<Accessory> accessories)
        {
            InitializeComponent();
            this.accessories = accessories;
            flag = false;
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (flag)
            {
                foreach (var b in boats)
                {
                    if ((bool)Increase.IsChecked)
                        b.BasePrice = b.BasePrice / 100 * (100 + Convert.ToDecimal(tbIncrease.Text));
                    if ((bool)Decrease.IsChecked)
                        b.BasePrice = b.BasePrice / 100 * (100 - Convert.ToDecimal(tbDecrease.Text));
                }
            }
            if (!flag)
            {
                foreach (var a in accessories)
                {
                    if ((bool)Increase.IsChecked)
                        a.Price = a.Price / 100 * (100 + Convert.ToDecimal(tbIncrease.Text));
                    if ((bool)Decrease.IsChecked)
                        a.Price = a.Price / 100 * (100 - Convert.ToDecimal(tbDecrease.Text));
                }
            }
            AppData.db.SaveChanges();
            Close();
        }

        private void tbIncrease_GotFocus(object sender, RoutedEventArgs e)
        {
            Increase.IsChecked = true;
        }

        private void tbDecrease_GotFocus(object sender, RoutedEventArgs e)
        {
            Decrease.IsChecked = true;
        }
    }
}