# Репозиторий для подготовки к WSR 2022 
---

## План работы и что надо научиться делать
https://www.evernote.com/shard/s300/sh/64fac59b-edcc-2cc9-a56e-67ba66fd057e/2aab59d756e7d605154545b376b4fcf9
### Git
https://www.atlassian.com/ru/git/tutorials/learn-git-with-bitbucket-cloud


---

## 1 Подготовка по базам данных
### Прочитать книгу Осипова. Книга даст общее понимание про базы данных.
### Просмотреть курсы https://disk.yandex.ru/d/ObrMuo3lH1LBiQ

---
## 2 Подготовка по чистому С#
### Выбрать вариант любой и выполнить 10 заданий в файле <a href = "https://github.com/artemovsergey/WSR_2021/blob/master/%D0%9C%D0%B0%D1%82%D0%B5%D1%80%D0%B8%D0%B0%D0%BB%D1%8B%20%D0%A1%23/Zadachi.pdf">Задачи</a>
---


## 3 Windows Forms основы
https://metanit.com/sharp/windowsforms/

---

## 4 Windows Forms
### Пройти учебник Осипова по основам windows Forms. Особенно посмотреть практическую работу 10 по работе с базой данных. 



## Справка по командам
### Подключение MySql в проекта Windows Forms

Cначала надо прописать в директивах
using MySql.Data.MySqlClient. 
Для этого надо убедиться, что в ссылках установлен пакет NuGet. который называется Mysql Data
```
            try
            {

                // create connection
                string mysql = @"server=localhost;userid=root;password=root;database=test_data";
                MySqlConnection connect = new MySqlConnection(mysql);
                connect.Open();



                // create query
                string sql = "select * from billing";
                MySqlCommand query = new MySqlCommand(sql, connect);

                string result = query.ExecuteScalar().ToString();

                MySqlDataReader result1 = query.ExecuteReader();

                while (result1.Read())
                {
                    TextBox.Text = ($"{result1[0]}");
                    Console.WriteLine(result1[0].ToString());
                }
                result1.Close();


                //TextBox.Text = ($"{result}");
                connect.Close();
            }

            catch
            {
                TextBox.Text = ($"Не удалось подключиться к базе данных");
            }

            finally
            {
                Console.WriteLine("Сеанс работы с базой данных завершен!");
            }
```
## Применение DataReader при подключении к MySQL

string MyConString = "SERVER=localhost;" +"DATABASE=mydatabase;" 
        "UID=testuser;" +"PASSWORD=testpassword;";
    MySqlConnection connection = new MySqlConnection(MyConString);
    MySqlCommand command = connection.CreateCommand();
    MySqlDataReader Reader;
    command.CommandText = "select * from mycustomers";
    connection.Open();
    Reader = command.ExecuteReader();
    while (Reader.Read())
    {
        string thisrow = "";                
        for (int i = 0;i<Reader.FieldCount;i++)     
            thisrow += Reader.GetValue(i).ToString() + ","; 
        listBox1.Items.Add(thisrow);
    }
    connection.Close(); 
    
 
 ## Вспомогательный метод Load Data
 
 public void Load_Data()
        {

            try
            {

                string connection = @"server=localhost;userid=root;password=root;database=test_data";
                MySqlConnection connect = new MySqlConnection(connection);
                connect.Open();



                MySqlCommand cmd = connect.CreateCommand();
                cmd.CommandText = "select * from materials";

                //MySqlCommand query = new MySqlCommand("select * from materials", connect);

                //string result = query.ExecuteScalar().ToString();

                //MySqlDataReader result_array = query.ExecuteReader();



                //while (result_array.Read())
                //textBox1.Text = result_array[1].ToString();


                // Dataset settings



                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                DataSet dataset = new DataSet();
                adapter.Fill(dataset);

                dataGridView1.DataSource = dataset.Tables[0];

                adapter.Update(dataset);

                DataView materialDataView = new DataView(dataset.Tables[0]);
                dataGridView1.DataSource = materialDataView;

                // Присвоения исходного порядка сортировки
                materialDataView.Sort = Sort.Text;
                materialDataView.RowFilter = Filter.Text;

            }
            catch
            {
                MessageBox.Show("Не удалось подключиться к базе данных!");
                throw;

            }
        }
 
 
 
    
    

