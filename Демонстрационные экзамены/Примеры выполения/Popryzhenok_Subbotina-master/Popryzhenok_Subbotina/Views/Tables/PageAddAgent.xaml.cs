using ClassLibraryPopryzhenok.Models;
using ClassLibraryPopryzhenok.Utils;
using Microsoft.Win32;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;

namespace Popryzhenok_Subbotina
{
    /// <summary>
    /// Логика взаимодействия для PageAddAgent.xaml
    /// </summary>
    public partial class PageAddAgent : Page
    {
        public Agent NewAgent { get; set; }

        public static string workingDirectory = Environment.CurrentDirectory;
        public static string projectDirectory = Directory.GetParent(workingDirectory).Parent.Parent.FullName;

        public PageAddAgent(Agent agent)
        {
            InitializeComponent();
            if (agent != null)
            {
                NewAgent = agent;
            }
            DataContext = NewAgent;
            cbType.ItemsSource = AppData.db.AgentTypes.ToList();
            MessageBox.Show(NewAgent.SalesPerYear.ToString());
        }

        private void buttonLogo_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                DefaultExt = ".png",
                FileName = "picture.png",
                Filter = "Image Files(*.BMP;*.JPG;*.PNG)|*.BMP;*.JPG;*.PNG",
                InitialDirectory = projectDirectory + @"\Popryzhenok_Subbotina\agents\"
            };
            bool? result = openFileDialog.ShowDialog();

            if (result == true)
            {
                string filename = openFileDialog.FileName;
                photo.Source = new BitmapImage(new Uri(filename));
                tbLogo.Text = "/agents/" + Path.GetFileName(filename);
                try
                {
                    File.Copy(filename, $"{projectDirectory}/Popryzhenok_Subbotina{tbLogo.Text}");
                    // TODO: RESCAN AGENTS FOLDER
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                    return;
                }
            }
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            StringBuilder errors = new StringBuilder();

            if (string.IsNullOrEmpty(tbTitle.Text) || string.IsNullOrWhiteSpace(tbTitle.Text))
            {
                errors.AppendLine("Введите название");
            }
            if (cbType.SelectedIndex == -1)
            {
                errors.AppendLine("Выберите тип");
            }
            if (string.IsNullOrEmpty(tbINN.Text) || string.IsNullOrWhiteSpace(tbINN.Text))
            {
                errors.AppendLine("Введите ИНН");
            }
            if (string.IsNullOrEmpty(tbDirector.Text) || string.IsNullOrWhiteSpace(tbDirector.Text))
            {
                errors.AppendLine("Введите имя директора");
            }
            if (string.IsNullOrEmpty(tbPhone.Text) || string.IsNullOrWhiteSpace(tbPhone.Text))
            {
                errors.AppendLine("Введите номер телефона");
            }
            if (string.IsNullOrEmpty(tbPriority.Text) || string.IsNullOrWhiteSpace(tbPriority.Text))
            {
                errors.AppendLine("Введите приоритет");
            }
            else if (!int.TryParse(tbPriority.Text, out int temp))
            {
                errors.AppendLine("Приоритет должен быть числовым значением");
            }
            else if (temp < 0)
            {
                errors.AppendLine("Приоритет должен быть больше нуля");
            }

            if (errors.Length > 0)
            {
                MessageBox.Show(errors.ToString());
                return;
            }

            if (NewAgent.ID == 0)
            {
                AppData.db.Agents.Add(NewAgent);
            }
            try
            {
                NewAgent.Logo = tbLogo.Text;
                AppData.db.SaveChanges();
                MessageBox.Show("Информация сохранена");
                NavigationService.GoBack();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.GoBack();
        }

        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (NewAgent != null)
            {
                if (MessageBox.Show("Вы уверены?", "Подтверждение", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    try
                    {
                        AppData.db.Agents.Remove(NewAgent);
                        AppData.db.SaveChanges();
                        MessageBox.Show("Успешно удалено");
                        NavigationService.GoBack();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
            }
        }

        private void Page_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                photo.Source = new BitmapImage(new Uri($"{projectDirectory}/Popryzhenok_Subbotina/{NewAgent.Logo}"));
            }
            catch (Exception)
            {
                return;
            }
        }
    }
}