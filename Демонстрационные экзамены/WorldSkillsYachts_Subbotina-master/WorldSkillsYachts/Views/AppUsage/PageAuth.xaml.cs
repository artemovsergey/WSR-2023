using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Navigation;
using System.Windows.Threading;
using WorldSkillsYachts.Models;
using WorldSkillsYachts.Utils;
using WorldSkillsYachts.Views.Windows;

namespace WorldSkillsYachts.Views.Pages
{
    /// <summary>
    /// Interaction logic for PageAuth.xaml
    /// </summary>
    public partial class PageAuth : Page
    {
        public static User user = new User();

        private int attempts = 0;
        private int timeBlock;

        private DispatcherTimer timer;

        public PageAuth()
        {
            InitializeComponent();
            LoginBox.Text = "manager";
            PassBox.Password = "manager";
            User.CheckUsers();
        }

        public static string Authorize(string login, string password)
        {
            List<User> users = AppData.db.Users.ToList();
            user = users.SingleOrDefault(x => x.Login == login);
            if (user != null && user.Password == password && !user.IsBlocked)
            {
                Properties.Settings.Default.IsLoggedIn = true;
                user.LastLogin = DateTime.Today;
                AppData.db.SaveChanges();
                switch (user.Role_ID)
                {
                    case 1:
                        Properties.Settings.Default.IsAdmin = true;
                        break;

                    case 2:
                        Properties.Settings.Default.IsAdmin = false;
                        break;

                    default:
                        break;
                }
                return "OK";
            }
            else if (user == null)
                return "User not found";
            else if (user.IsBlocked && user.Password == password)
                return "This user is blocked";
            else if (user.Password != password)
                return "Incorrect password";
            else
                return null;
        }

        private void LoginButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(LoginBox.Text) ||
                string.IsNullOrEmpty(LoginBox.Text) ||
                string.IsNullOrEmpty(PassBox.Password) ||
                string.IsNullOrWhiteSpace(PassBox.Password))
            {
                MessageBox.Show("Пожалуйста, заполните все поля",
                    "Внимание!", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            var result = Authorize(LoginBox.Text, PassBox.Password);

            if (result == "OK")
            {
                attempts = 0;
                if (Properties.Settings.Default.IsLoggedIn)
                {
                    if (User.PasswordNotChanged(user) >= 14)
                    {
                        WindowChangePassword windowChangePassword = new WindowChangePassword(user);
                        windowChangePassword.ShowDialog();
                        Properties.Settings.Default.IsLoggedIn = false;
                        PassBox.Password = string.Empty;
                        return;
                    }
                    else
                    {
                        user.LastLogin = DateTime.Today;
                        if (Properties.Settings.Default.IsAdmin)
                            NavigationService.Navigate(new PageAdmin());
                        else
                            NavigationService.Navigate(new PageManager());
                    }
                }
            }
            else if (result == "User not found")
            {
                MessageBox.Show("Пользователь не найден",
                    "Внимание!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else if (result == "This user is blocked")
            {
                MessageBox.Show("Ваш аккаунт заблокирован спустя месяц неактивности. Обратитесь к системному администратору для разблокировки",
                    "Внимание!", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            else if (result == "Incorrect password")
            {
                if (attempts >= 3)
                {
                    MessageBox.Show("Неверный пароль введен более трех раз. Повторите позднее.",
                        "Внимание!", MessageBoxButton.OK, MessageBoxImage.Error);
                    LoginButton.IsEnabled = false;
                    PassBox.IsEnabled = false;
                    TimeSpan time = TimeSpan.FromSeconds(15 + timeBlock);
                    timer = new DispatcherTimer(new TimeSpan(0, 0, 1), DispatcherPriority.Normal, delegate
                    {
                        SecondsBlocked.Text = $"Попробуйте ввести пароль заново через {time:c} секунд";
                        if (time == TimeSpan.Zero)
                        {
                            timer.Stop();
                            LoginButton.IsEnabled = true;
                            PassBox.IsEnabled = true;
                            timeBlock += 20;
                        }
                        time = time.Add(TimeSpan.FromSeconds(-1));
                    },
                    Application.Current.Dispatcher);
                }
                else
                {
                    MessageBox.Show("Неверный пароль.", "Внимание!", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                attempts++;
            }
            else
            {
                MessageBox.Show("Неизвестная ошибка", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void RegisterButton_Click(object sender, RoutedEventArgs e)
        {
            WindowRegister windowRegister = new WindowRegister();
            windowRegister.ShowDialog();
        }
    }
}